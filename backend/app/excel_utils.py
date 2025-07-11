import os
import re
import pandas as pd
from typing import List, Dict, Any, Optional, Tuple, Union

class ExcelUtils:
    def __init__(self, excel_path: str):
        self.excel_path = excel_path
    
    def load_excel(self) -> pd.DataFrame:
        """Load the Excel file into a pandas DataFrame"""
        return pd.read_excel(self.excel_path, engine="openpyxl", header=None)
    
    def save_excel(self, df: pd.DataFrame) -> None:
        """Save the DataFrame to the Excel file"""
        df.to_excel(self.excel_path, index=False, header=False)
    
    def excel_cell_to_index(self, cell: str) -> Tuple[int, int]:
        """Convert Excel cell reference (e.g., 'A1') to row and column indices"""
        match = re.match(r"([A-Za-z]+)(\d+)", cell)
        if not match:
            raise ValueError(f"Invalid cell format: {cell}")
        col_letters, row = match.groups()
        col = sum([(ord(c.upper()) - ord('A') + 1) * (26 ** i) for i, c in enumerate(reversed(col_letters))]) - 1
        return int(row) - 1, col
    
    def read_excel_range(self, range: str) -> str:
        """Read a range of cells from the Excel file"""
        df = self.load_excel()
        start_cell, end_cell = range.split(":")
        start_row, start_col = self.excel_cell_to_index(start_cell)
        end_row, end_col = self.excel_cell_to_index(end_cell)
        sub_df = df.iloc[start_row:end_row+1, start_col:end_col+1]
        return sub_df.to_csv(index=False, header=False)
    
    def read_cell(self, cell: Optional[str] = None, row: Optional[int] = None, col: Optional[int] = None) -> str:
        """Read a single cell from the Excel file"""
        if cell:
            row_idx, col_idx = self.excel_cell_to_index(cell)
        elif row is not None and col is not None:
            row_idx, col_idx = row, col
        else:
            raise ValueError("Provide either cell or both row and col")
        df = self.load_excel()
        return str(df.iat[row_idx, col_idx])
    
    def update_cell(self, value: str, cell: Optional[str] = None, row: Optional[int] = None, col: Optional[int] = None) -> str:
        """Update a single cell in the Excel file"""
        if cell:
            row_idx, col_idx = self.excel_cell_to_index(cell)
        elif row is not None and col is not None:
            row_idx, col_idx = row, col
        else:
            raise ValueError("Provide either cell or both row and col")
        df = self.load_excel()
        df.iat[row_idx, col_idx] = value
        self.save_excel(df)
        return f"✅ Updated cell ({row_idx}, {col_idx}) to '{value}'"
    
    def update_range(self, range: str, values: List[List[str]]) -> str:
        """Update a range of cells in the Excel file"""
        self.save_excel(self.load_excel())  # Optional: ensure latest file
        df = self.load_excel()  # ✅ reload to get any inserted rows
        
        start_cell, _ = range.split(":")
        start_row, start_col = self.excel_cell_to_index(start_cell)

        for i, row_vals in enumerate(values):
            for j, val in enumerate(row_vals):
                # Expand DataFrame if necessary
                target_row = start_row + i
                target_col = start_col + j

                if target_row >= len(df):
                    # Add new rows
                    for _ in range(target_row - len(df) + 1):
                        df = pd.concat([df, pd.DataFrame([[None] * df.shape[1]])], ignore_index=True)

                if target_col >= df.shape[1]:
                    # Add new columns
                    for _ in range(target_col - df.shape[1] + 1):
                        df[df.shape[1]] = None

                df.iat[target_row, target_col] = val

        self.save_excel(df)
        return f"✅ Updated range {range}"



    def find_and_replace(self, find: str, replace: str, range: Optional[str] = None) -> str:
        """Find and replace values in the Excel file"""
        df = self.load_excel()
        if range:
            start_cell, end_cell = range.split(":")
            start_row, start_col = self.excel_cell_to_index(start_cell)
            end_row, end_col = self.excel_cell_to_index(end_cell)
            sub_df = df.iloc[start_row:end_row+1, start_col:end_col+1]
            df.iloc[start_row:end_row+1, start_col:end_col+1] = sub_df.replace(find, replace, regex=False)
        else:
            df.replace(find, replace, regex=False, inplace=True)
        self.save_excel(df)
        return f"🔄 Replaced '{find}' with '{replace}'"
    
    def summarize_range(self, range: str, operation: str) -> str:
        """Perform a summary operation on a range of cells"""
        df = self.load_excel()
        start_cell, end_cell = range.split(":")
        start_row, start_col = self.excel_cell_to_index(start_cell)
        end_row, end_col = self.excel_cell_to_index(end_cell)
        sub_df = df.iloc[start_row:end_row+1, start_col:end_col+1]
        sub_df_numeric = pd.to_numeric(sub_df.stack(), errors='coerce')
        
        # Calculate the result based on the operation
        if operation == "sum":
            result = sub_df_numeric.sum()
        elif operation == "avg":
            result = sub_df_numeric.mean()
        elif operation == "min":
            result = sub_df_numeric.min()
        elif operation == "max":
            result = sub_df_numeric.max()
        else:
            return f"❌ Unsupported operation: {operation}"
        
        # Convert result to a JSON-serializable format
        if isinstance(result, pd.Timestamp) or hasattr(result, 'isoformat'):
            result_str = result.isoformat()
        else:
            result_str = str(result)
        
        # Return formatted result
        operation_emoji = {
            "sum": "🔢",
            "avg": "📊",
            "min": "🔽",
            "max": "🔼"
        }.get(operation, "")
        
        return f"{operation_emoji} {operation.capitalize()} = {result_str}"
    
    def insert_row_or_col(self, type: str, index: int, count: int = 1) -> str:
        """Insert rows or columns in the Excel file"""
        df = self.load_excel()
        if type == "row":
            for _ in range(count):
                df = pd.concat([df.iloc[:index], pd.DataFrame([[None]*df.shape[1]]), df.iloc[index:]], ignore_index=True)
        elif type == "column":
            for _ in range(count):
                df.insert(index, None, None)
        else:
            return f"❌ Invalid type: {type}. Use 'row' or 'column'."
        self.save_excel(df)
        return f"✅ Inserted {count} {type}(s) at index {index}"
    
    def delete_row_or_col(self, type: str, index: int, count: int = 1) -> str:
        """Delete rows or columns from the Excel file"""
        df = self.load_excel()
        if type == "row":
            df = df.drop(range(index, index + count)).reset_index(drop=True)
        elif type == "column":
            df = df.drop(columns=range(index, index + count))
        else:
            return f"❌ Invalid type: {type}. Use 'row' or 'column'."
        self.save_excel(df)
        return f"✅ Deleted {count} {type}(s) at index {index}"
    
    def read_sheet_metadata(self) -> str:
        """Read metadata about the Excel sheet"""
        df = self.load_excel()
        return f"Sheet dimensions: {df.shape[0]} rows × {df.shape[1]} columns"
    
    def get_column_values(self, index: int) -> str:
        """Get all values in a column"""
        df = self.load_excel()
        return df.iloc[:, index].to_csv(index=False, header=False)
    
    def filter_rows(self, col_index: int, value: str) -> str:
        """Filter rows where a column matches a value"""
        df = self.load_excel()
        filtered = df[df.iloc[:, col_index].astype(str) == value]
        return filtered.to_csv(index=False, header=False)
    
    def get_dataframe(self) -> pd.DataFrame:
        """Return the current DataFrame for display purposes"""
        return self.load_excel()

    def get_last_filled_row_index(self) -> int:
        """Return the index of the last non-empty row"""
        df = self.load_excel()
        # Drop fully empty rows from bottom up
        non_empty_df = df.dropna(how='all')
        return len(non_empty_df)
