import os
import re
import pandas as pd
from typing import List, Dict, Any, Optional, Tuple, Union

class ExcelUtils:
    def __init__(self, file_path: str):
        """Initialize with a spreadsheet file path"""
        self.file_path = file_path
        self.file_extension = os.path.splitext(file_path)[1].lower()
    
    def load_excel(self) -> pd.DataFrame:
        """Load the spreadsheet file into a pandas DataFrame"""
        # Determine the appropriate engine and options based on file extension
        if self.file_extension in ['.xlsx', '.xlsm', '.xltx', '.xltm']:
            return pd.read_excel(self.file_path, engine="openpyxl", header=None)
        elif self.file_extension == '.xls':
            return pd.read_excel(self.file_path, engine="xlrd", header=None)
        elif self.file_extension == '.csv':
            return pd.read_csv(self.file_path, header=None)
        elif self.file_extension == '.tsv':
            return pd.read_csv(self.file_path, sep='\t', header=None)
        elif self.file_extension in ['.ods', '.fods']:
            return pd.read_excel(self.file_path, engine="odf", header=None)
        else:
            # Default to openpyxl for unknown extensions
            return pd.read_excel(self.file_path, engine="openpyxl", header=None)
    
    def save_excel(self, df: pd.DataFrame) -> None:
        """Save the DataFrame to the spreadsheet file"""
        # Determine the appropriate save method based on file extension
        try:
            if self.file_extension == '.xlsx':
                df.to_excel(self.file_path, engine="openpyxl", index=False, header=False)
            elif self.file_extension == '.xls':
                # For .xls files, we need to convert to .xlsx as pandas doesn't support writing to .xls directly
                # with newer versions
                xlsx_path = self.file_path.replace('.xls', '.xlsx')
                df.to_excel(xlsx_path, engine="openpyxl", index=False, header=False)
                # If the original file was .xls, we'll keep it as the main file
                # but we'll work with the .xlsx version internally
                self.file_path = xlsx_path
                self.file_extension = '.xlsx'
            elif self.file_extension in ['.xlsm', '.xltx', '.xltm']:
                df.to_excel(self.file_path, engine="openpyxl", index=False, header=False)
            elif self.file_extension == '.csv':
                df.to_csv(self.file_path, index=False, header=False)
            elif self.file_extension == '.tsv':
                df.to_csv(self.file_path, sep='\t', index=False, header=False)
            elif self.file_extension in ['.ods', '.fods']:
                df.to_excel(self.file_path, engine="odf", index=False, header=False)
            else:
                # Default to Excel format for unknown extensions
                df.to_excel(self.file_path, engine="openpyxl", index=False, header=False)
        except Exception as e:
            print(f"Error saving file: {str(e)}")
            raise
    
    def excel_cell_to_index(self, cell: str) -> Tuple[int, int]:
        """Convert Excel cell reference (e.g., 'A1') to row and column indices"""
        match = re.match(r"([A-Za-z]+)(\d+)", cell)
        if not match:
            raise ValueError(f"Invalid cell format: {cell}")
        col_letters, row = match.groups()
        col = sum([(ord(c.upper()) - ord('A') + 1) * (26 ** i) for i, c in enumerate(reversed(col_letters))]) - 1
        return int(row) - 1, col
    
    def read_excel_range(self, range: str) -> str:
        """Read a range of cells from the spreadsheet file"""
        df = self.load_excel()
        start_cell, end_cell = range.split(":")
        start_row, start_col = self.excel_cell_to_index(start_cell)
        end_row, end_col = self.excel_cell_to_index(end_cell)
        sub_df = df.iloc[start_row:end_row+1, start_col:end_col+1]
        return sub_df.to_csv(index=False, header=False)
    
    def read_cell(self, cell: Optional[str] = None, row: Optional[int] = None, col: Optional[int] = None) -> str:
        """Read a single cell from the spreadsheet file"""
        if cell:
            row_idx, col_idx = self.excel_cell_to_index(cell)
        elif row is not None and col is not None:
            row_idx, col_idx = row, col
        else:
            raise ValueError("Provide either cell or both row and col")
        df = self.load_excel()
        return str(df.iat[row_idx, col_idx])
    
    def update_cell(self, value: str, cell: Optional[str] = None, row: Optional[int] = None, col: Optional[int] = None) -> str:
        """Update a single cell in the spreadsheet file"""
        try:
            if cell:
                row_idx, col_idx = self.excel_cell_to_index(cell)
            elif row is not None and col is not None:
                row_idx, col_idx = row, col
            else:
                raise ValueError("Provide either cell or both row and col")
                
            df = self.load_excel()
            
            # Ensure DataFrame has enough rows
            while row_idx >= len(df):
                # Create a new row with the same column structure
                new_row = [[None] * df.shape[1]]
                new_df = pd.DataFrame(new_row, columns=df.columns)
                df = pd.concat([df, new_df], ignore_index=True)
            
            # Ensure DataFrame has enough columns
            while col_idx >= df.shape[1]:
                # Add a new column
                df.insert(df.shape[1], str(df.shape[1]), None)
            
            # Update the cell value
            df.iat[row_idx, col_idx] = value
            self.save_excel(df)
            return f"✅ Updated cell ({row_idx}, {col_idx}) to '{value}'"
        except Exception as e:
            error_msg = str(e)
            print(f"Error updating cell: {error_msg}")
            return f"❌ Error updating cell: {error_msg}"
    
    def update_range(self, range: str, values: List[List[str]]) -> str:
        """Update a range of cells in the spreadsheet file"""
        try:
            # Ensure we have the latest file
            df = self.load_excel()
            
            # Parse the range
            start_cell, end_cell = range.split(":")
            start_row, start_col = self.excel_cell_to_index(start_cell)
            
            # Process each value in the provided values list
            for i, row_vals in enumerate(values):
                for j, val in enumerate(row_vals):
                    # Calculate target position
                    target_row = start_row + i
                    target_col = start_col + j
                    
                    # Ensure DataFrame has enough rows
                    while target_row >= len(df):
                        # Create a new row with the same column structure
                        new_row = [[None] * df.shape[1]]
                        new_df = pd.DataFrame(new_row, columns=df.columns)
                        df = pd.concat([df, new_df], ignore_index=True)
                    
                    # Ensure DataFrame has enough columns
                    while target_col >= df.shape[1]:
                        # Add a new column
                        df.insert(df.shape[1], str(df.shape[1]), None)
                    
                    # Update the cell value
                    df.iat[target_row, target_col] = val
            
            # Save the updated DataFrame
            self.save_excel(df)
            return f"✅ Updated range {range}"
        except Exception as e:
            error_msg = str(e)
            print(f"Error updating range: {error_msg}")
            return f"❌ Error updating range: {error_msg}"

    def find_and_replace(self, find: str, replace: str, range: Optional[str] = None) -> str:
        """Find and replace values in the spreadsheet file"""
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
        """Insert rows or columns in the spreadsheet file"""
        df = self.load_excel()
        if type == "row":
            # Create a new row with the same data types as the existing DataFrame
            if len(df) > 0:
                # Create a new DataFrame with the same structure but filled with None values
                num_cols = df.shape[1]
                # Use a different approach to avoid the FutureWarning
                # Create a new row as a list of None values
                new_row = [[None] * num_cols]
                # Create a new DataFrame with the same column structure
                empty_df = pd.DataFrame(new_row, columns=df.columns)
                
                # Insert the row at the specified index
                if index == 0:
                    df = pd.concat([empty_df, df], ignore_index=True)
                elif index >= len(df):
                    df = pd.concat([df, empty_df], ignore_index=True)
                else:
                    df = pd.concat([df.iloc[:index], empty_df, df.iloc[index:]], ignore_index=True)
            else:
                # If DataFrame is empty, just add a row with NaN values
                df = pd.DataFrame([[None] * (df.shape[1] or 1)])
        elif type == "column":
            for _ in range(count):
                # Insert a new column with None values
                col_name = f"new_col_{_}"
                df.insert(index, col_name, None)
                # Rename the column to match numeric indexing if needed
                df = df.rename(columns={col_name: df.columns.size - 1})
        else:
            return f"❌ Invalid type: {type}. Use 'row' or 'column'."
        self.save_excel(df)
        return f"✅ Inserted {count} {type}(s) at index {index}"
    
    def delete_row_or_col(self, type: str, index: int, count: int = 1) -> str:
        """Delete rows or columns from the spreadsheet file"""
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
        """Read metadata about the spreadsheet"""
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
