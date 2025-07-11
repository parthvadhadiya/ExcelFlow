import os
import re
import pandas as pd
import streamlit as st
from dotenv import load_dotenv
from openai import OpenAI
from openai.types.chat import ChatCompletionMessageParam

load_dotenv()

# === Config ===
EXCEL_PATH = os.getenv("EXCEL_PATH", "./data.xlsx")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

def load_excel():
    return pd.read_excel(EXCEL_PATH, engine="openpyxl", header=None)

def save_excel(df):
    df.to_excel(EXCEL_PATH, index=False, header=False)

def excel_cell_to_index(cell: str):
    match = re.match(r"([A-Za-z]+)(\d+)", cell)
    if not match:
        raise ValueError(f"Invalid cell format: {cell}")
    col_letters, row = match.groups()
    col = sum([(ord(c.upper()) - ord('A') + 1) * (26 ** i) for i, c in enumerate(reversed(col_letters))]) - 1
    return int(row) - 1, col

# === Tools ===
def read_excel_range(range: str) -> str:
    df = load_excel()
    start_cell, end_cell = range.split(":")
    start_row, start_col = excel_cell_to_index(start_cell)
    end_row, end_col = excel_cell_to_index(end_cell)
    sub_df = df.iloc[start_row:end_row+1, start_col:end_col+1]
    return sub_df.to_csv(index=False, header=False)

def read_cell(cell: str = None, row: int = None, col: int = None) -> str:
    if cell:
        row_idx, col_idx = excel_cell_to_index(cell)
    elif row is not None and col is not None:
        row_idx, col_idx = row, col
    else:
        raise ValueError("Provide either cell or both row and col")
    df = load_excel()
    return str(df.iat[row_idx, col_idx])

def update_cell(value: str, cell: str = None, row: int = None, col: int = None) -> str:
    if cell:
        row_idx, col_idx = excel_cell_to_index(cell)
    elif row is not None and col is not None:
        row_idx, col_idx = row, col
    else:
        raise ValueError("Provide either cell or both row and col")
    df = load_excel()
    df.iat[row_idx, col_idx] = value
    save_excel(df)
    return f"✅ Updated cell ({row_idx}, {col_idx}) to '{value}'"

def update_range(range: str, values: list[list[str]]) -> str:
    df = load_excel()
    start_cell, _ = range.split(":")
    start_row, start_col = excel_cell_to_index(start_cell)
    for i, row_vals in enumerate(values):
        for j, val in enumerate(row_vals):
            df.iat[start_row + i, start_col + j] = val
    save_excel(df)
    return f"✅ Updated range {range}"

def find_and_replace(find: str, replace: str, range: str = None) -> str:
    df = load_excel()
    if range:
        start_cell, end_cell = range.split(":")
        start_row, start_col = excel_cell_to_index(start_cell)
        end_row, end_col = excel_cell_to_index(end_cell)
        sub_df = df.iloc[start_row:end_row+1, start_col:end_col+1]
        df.iloc[start_row:end_row+1, start_col:end_col+1] = sub_df.replace(find, replace, regex=False)
    else:
        df.replace(find, replace, regex=False, inplace=True)
    save_excel(df)
    return f"🔄 Replaced '{find}' with '{replace}'"

def summarize_range(range: str, operation: str) -> str:
    df = load_excel()
    start_cell, end_cell = range.split(":")
    start_row, start_col = excel_cell_to_index(start_cell)
    end_row, end_col = excel_cell_to_index(end_cell)
    sub_df = df.iloc[start_row:end_row+1, start_col:end_col+1]
    sub_df_numeric = pd.to_numeric(sub_df.stack(), errors='coerce')
    if operation == "sum":
        return f"🔢 Sum = {sub_df_numeric.sum()}"
    elif operation == "avg":
        return f"📊 Average = {sub_df_numeric.mean()}"
    elif operation == "min":
        return f"🔽 Min = {sub_df_numeric.min()}"
    elif operation == "max":
        return f"🔼 Max = {sub_df_numeric.max()}"
    else:
        return f"❌ Unsupported operation: {operation}"

def insert_row_or_col(type: str, index: int, count: int = 1) -> str:
    df = load_excel()
    if type == "row":
        for _ in range(count):
            df = pd.concat([df.iloc[:index], pd.DataFrame([[None]*df.shape[1]]), df.iloc[index:]], ignore_index=True)
    elif type == "column":
        for _ in range(count):
            df.insert(loc=index, column=df.shape[1], value=[None]*df.shape[0])
    else:
        return f"❌ Invalid type: {type}"
    save_excel(df)
    return f"➕ Inserted {count} {type}(s) at index {index}"

def delete_row_or_col(type: str, index: int, count: int = 1) -> str:
    df = load_excel()
    if type == "row":
        df.drop(index=list(range(index, index+count)), inplace=True)
        df.reset_index(drop=True, inplace=True)
    elif type == "column":
        df.drop(columns=list(range(index, index+count)), inplace=True)
    else:
        return f"❌ Invalid type: {type}"
    save_excel(df)
    return f"🗑️ Deleted {count} {type}(s) starting from index {index}"

def read_sheet_metadata() -> str:
    df = load_excel()
    rows, cols = df.shape
    return f"Sheet has {rows} rows and {cols} columns."

def get_column_values(index: int) -> str:
    df = load_excel()
    return df.iloc[:, index].to_csv(index=False, header=False)

def filter_rows(col_index: int, value: str) -> str:
    df = load_excel()
    filtered = df[df.iloc[:, col_index].astype(str) == value]
    return filtered.to_csv(index=False, header=False)




# === OpenAI Agent ===
client = OpenAI(api_key=OPENAI_API_KEY)

tools = tools = [
    {
        "type": "function",
        "function": {
            "name": "read_excel_range",
            "description": "Read a rectangular range from the Excel sheet (e.g., A1:C5).",
            "parameters": {
                "type": "object",
                "properties": {
                    "range": {
                        "type": "string",
                        "description": "Range in Excel notation (e.g., A1:C3)"
                    }
                },
                "required": ["range"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "read_cell",
            "description": "Read a single cell from the Excel sheet.",
            "parameters": {
                "type": "object",
                "properties": {
                    "cell": {"type": "string", "description": "Excel-style cell name (e.g., B2)"},
                    "row": {"type": "integer", "description": "Zero-based row index (alt to cell)"},
                    "col": {"type": "integer", "description": "Zero-based col index (alt to cell)"}
                },
                "required": []
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "update_cell",
            "description": "Update a single Excel cell with a new value.",
            "parameters": {
                "type": "object",
                "properties": {
                    "value": {"type": "string", "description": "New value to insert"},
                    "cell": {"type": "string", "description": "Excel-style cell name (e.g., B2)"},
                    "row": {"type": "integer", "description": "Zero-based row index (alt to cell)"},
                    "col": {"type": "integer", "description": "Zero-based col index (alt to cell)"}
                },
                "required": ["value"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "update_range",
            "description": "Update a range of cells with a 2D array of values.",
            "parameters": {
                "type": "object",
                "properties": {
                    "range": {"type": "string", "description": "Range to update (e.g., A1:C3)"},
                    "values": {
                        "type": "array",
                        "items": {
                            "type": "array",
                            "items": {"type": "string"}
                        },
                        "description": "2D list of new values"
                    }
                },
                "required": ["range", "values"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "find_and_replace",
            "description": "Find and replace values in the Excel sheet.",
            "parameters": {
                "type": "object",
                "properties": {
                    "find": {"type": "string", "description": "Text to find"},
                    "replace": {"type": "string", "description": "Replacement text"},
                    "range": {"type": "string", "description": "Optional range to restrict replacement"}
                },
                "required": ["find", "replace"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "summarize_range",
            "description": "Perform a summary operation over a numeric range.",
            "parameters": {
                "type": "object",
                "properties": {
                    "range": {"type": "string", "description": "Range to summarize (e.g., A2:C6)"},
                    "operation": {
                        "type": "string",
                        "enum": ["sum", "avg", "min", "max"],
                        "description": "Summary operation to perform"
                    }
                },
                "required": ["range", "operation"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "insert_row_or_col",
            "description": "Insert empty rows or columns into the Excel sheet.",
            "parameters": {
                "type": "object",
                "properties": {
                    "type": {
                        "type": "string",
                        "enum": ["row", "column"],
                        "description": "Whether to insert a row or column"
                    },
                    "index": {"type": "integer", "description": "Index at which to insert"},
                    "count": {"type": "integer", "description": "Number of rows/cols to insert (default 1)"}
                },
                "required": ["type", "index"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "delete_row_or_col",
            "description": "Delete rows or columns from the Excel sheet.",
            "parameters": {
                "type": "object",
                "properties": {
                    "type": {
                        "type": "string",
                        "enum": ["row", "column"],
                        "description": "Whether to delete a row or column"
                    },
                    "index": {"type": "integer", "description": "Starting index to delete"},
                    "count": {"type": "integer", "description": "Number of rows/cols to delete (default 1)"}
                },
                "required": ["type", "index"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "read_sheet_metadata",
            "description": "Get metadata like number of rows and columns from the Excel sheet.",
            "parameters": {
                "type": "object",
                "properties": {}
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "get_column_values",
            "description": "Return all values in a column.",
            "parameters": {
                "type": "object",
                "properties": {
                    "index": {"type": "integer", "description": "Zero-based column index"}
                },
                "required": ["index"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "filter_rows",
            "description": "Filter rows where column matches a value.",
            "parameters": {
                "type": "object",
                "properties": {
                    "col_index": {"type": "integer", "description": "Column index to filter on"},
                    "value": {"type": "string", "description": "Value to match"}
                },
                "required": ["col_index", "value"]
            }
        }
    }
]


def call_agent(message_history: list[ChatCompletionMessageParam]):
    response = client.chat.completions.create(
        model="gpt-4o",
        messages=message_history,
        tools=tools,
        tool_choice="auto"
    )
    msg = response.choices[0].message

    if msg.tool_calls:
        results = []
        for tool_call in msg.tool_calls:
            fn_name = tool_call.function.name
            args = eval(tool_call.function.arguments)
            print(f"Function: {fn_name}, Args: {args}")
            if fn_name == "read_excel_range":
                result = read_excel_range(**args)
            elif fn_name == "read_cell":
                result = read_cell(**args)
            elif fn_name == "update_cell":
                result = update_cell(**args)
            elif fn_name == "read_sheet_metadata":
                result = read_sheet_metadata()
            elif fn_name == "get_column_values":
                result = get_column_values(**args)
            elif fn_name == "filter_rows":
                result = filter_rows(**args)
            elif fn_name == "update_range":
                result = update_range(**args)
            elif fn_name == "find_and_replace":
                result = find_and_replace(**args)
            elif fn_name == "summarize_range":
                result = summarize_range(**args)
            elif fn_name == "insert_row_or_col":
                result = insert_row_or_col(**args)
            elif fn_name == "delete_row_or_col":
                result = delete_row_or_col(**args)
            else:
                result = f"❌ Unknown function {fn_name}"
            results.append({"tool_call_id": tool_call.id, "output": result})

        follow_up = client.chat.completions.create(
            model="gpt-4o",
            messages=message_history + [
                msg,
                *[{
                    "tool_call_id": r["tool_call_id"],
                    "role": "tool",
                    "name": msg.tool_calls[i].function.name,
                    "content": r["output"]
                } for i, r in enumerate(results)]
            ]
        )
        return follow_up.choices[0].message.content

    return msg.content

# === Streamlit UI ===
st.set_page_config(page_title="📊 Excel Agent", layout="wide")
st.title("📊 Excel Agent with OpenAI")

st.write(f"Using Excel: `{EXCEL_PATH}`")
df = load_excel()

# Fix datetime serialization issue for Streamlit UI
for col in df.columns:
    if pd.api.types.is_datetime64_any_dtype(df[col]):
        df[col] = df[col].astype(str)

st.dataframe(df)

user_input = st.chat_input("Ask me something about this sheet or make changes...")

if "history" not in st.session_state:
    st.session_state.history = [
        {"role": "system", "content": "You are a smart Excel assistant. You have access to functions. Always use them when asked to read/update Excel content."}
    ]

if user_input:
    st.session_state.history.append({"role": "user", "content": user_input})
    with st.spinner("Thinking..."):
        reply = call_agent(st.session_state.history)
        st.session_state.history.append({"role": "assistant", "content": reply})
        st.rerun()

for msg in st.session_state.history[1:]:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

st.divider()
st.caption("Built with ❤️ using OpenAI + Streamlit")