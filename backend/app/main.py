import os
import json
import uuid
import pandas as pd
from typing import Dict
from fastapi import FastAPI, WebSocket, WebSocketDisconnect, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
import uvicorn
from dotenv import load_dotenv

from .excel_utils import ExcelUtils
from .agent import ExcelAgent

load_dotenv()

app = FastAPI(title="Excel Agent API")

# Add CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # In production, replace with specific origins
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Global variables
UPLOAD_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

# Store active connections and their associated Excel files
active_connections: Dict[str, Dict] = {}

# Initialize the Excel agent
excel_agent = ExcelAgent()

class ConnectionManager:
    def __init__(self):
        self.active_connections: Dict[str, WebSocket] = {}
    
    async def connect(self, websocket: WebSocket, client_id: str):
        await websocket.accept()
        self.active_connections[client_id] = websocket
    
    def disconnect(self, client_id: str):
        if client_id in self.active_connections:
            del self.active_connections[client_id]
    
    async def send_message(self, client_id: str, message: str):
        if client_id in self.active_connections:
            await self.active_connections[client_id].send_text(message)
    
    async def broadcast(self, message: str):
        for connection in self.active_connections.values():
            await connection.send_text(message)

manager = ConnectionManager()

@app.get("/")
async def root():
    return {"message": "Excel Agent API is running"}

@app.post("/upload")
async def upload_excel(file: UploadFile = File(...)):
    """Upload an Excel file and return a client ID for WebSocket connection"""
    print(f"Received file upload: {file.filename}")
    
    try:
        # Generate a unique client ID
        client_id = str(uuid.uuid4())
        print(f"Generated client ID: {client_id}")
        
        # Save the uploaded file
        file_path = os.path.join(UPLOAD_DIR, f"{client_id}.xlsx")
        print(f"Saving file to: {file_path}")
        
        with open(file_path, "wb") as buffer:
            content = await file.read()
            print(f"Read {len(content)} bytes from file")
            buffer.write(content)
        
        # Initialize Excel utils for this client
        print("Initializing Excel utils")
        excel_utils = ExcelUtils(file_path)
        
        # Get DataFrame head and data range information
        df = excel_utils.get_dataframe()
        df_head_str = df.head().to_string()
        
        # Determine the range where data exists
        num_rows, num_cols = df.shape
        num_rows += 1
        from openpyxl.utils import get_column_letter
        data_range = f"A1:{get_column_letter(num_cols)}{num_rows}"

        
        system_message = f"""You are a smart Excel assistant. You have access to functions. Always use them when asked to read/update Excel content.
        After inserting a row or column, always re-evaluate the sheet before updating it.
        Make sure the inserted index exists before trying to write to it.
        When inserting a total row, always find the last filled row index using tools. Never hardcode the index.

        
        Here's a preview of top 5 rows of the Excel data structure it might contains header if not you can figure it out based on data: 
        {df_head_str}

        The Excel file contains data in the range {data_range} ({num_rows} rows × {num_cols} columns)."""
        active_connections[client_id] = {
            "file_path": file_path,
            "excel_utils": excel_utils,
            "message_history": [
                {"role": "system", "content": system_message}
            ]
        }
        
        print("File upload successful")
        return {"client_id": client_id}
    except Exception as e:
        print(f"Error in upload_excel: {str(e)}")
        raise

@app.get("/excel/{client_id}")
async def get_excel_data(client_id: str):
    """Get the current Excel data for a client"""
    if client_id not in active_connections:
        return JSONResponse(status_code=404, content={"error": "Client not found"})
    
    excel_utils = active_connections[client_id]["excel_utils"]
    df = excel_utils.get_dataframe()
    
    # Convert DataFrame to JSON-serializable format
    data = []
    for i in range(len(df)):
        row = {}
        for j in range(len(df.columns)):
            cell_value = df.iloc[i, j]
            # Handle non-serializable types
            if pd.isna(cell_value):
                cell_value = None
            elif isinstance(cell_value, pd.Timestamp):
                cell_value = cell_value.isoformat()
            row[str(j)] = cell_value
        data.append(row)
    
    # Get metadata
    rows, cols = df.shape
    
    return {
        "data": data,
        "metadata": {
            "rows": rows,
            "columns": cols
        }
    }

@app.websocket("/ws/{client_id}")
async def websocket_endpoint(websocket: WebSocket, client_id: str):
    """WebSocket endpoint for real-time communication with the Excel agent"""
    if client_id not in active_connections:
        await websocket.close(code=1000, reason="Client not found")
        return
    
    await manager.connect(websocket, client_id)
    
    try:
        while True:
            # Receive message from client
            data = await websocket.receive_text()
            message_data = json.loads(data)
            user_message = message_data.get("message", "")
            
            # Add user message to history
            active_connections[client_id]["message_history"].append({"role": "user", "content": user_message})
            
            # Process with the agent
            excel_utils = active_connections[client_id]["excel_utils"]
            message_history = active_connections[client_id]["message_history"]
            
            # Call the agent
            try:
                response, excel_modified = excel_agent.call_agent(message_history, excel_utils)
                
                # Add assistant response to history
                active_connections[client_id]["message_history"].append({"role": "assistant", "content": response})
                
                # Send response back to client
                try:
                    await manager.send_message(
                        client_id, 
                        json.dumps({
                            "response": response,
                            "excel_modified": excel_modified
                        })
                    )
                except TypeError as e:
                    print(f"JSON serialization error in response: {str(e)}")
                    # Try to serialize with a custom encoder
                    await manager.send_message(
                        client_id, 
                        json.dumps({
                            "response": str(response),
                            "excel_modified": excel_modified
                        })
                    )
                
                # If Excel was modified, notify client to refresh the data
                if excel_modified:
                    # Get updated Excel data
                    df = excel_utils.get_dataframe()
                    
                    # Convert DataFrame to JSON-serializable format
                    data = []
                    for i in range(len(df)):
                        row = {}
                        for j in range(len(df.columns)):
                            cell_value = df.iloc[i, j]
                            # Handle non-serializable types
                            if pd.isna(cell_value):
                                cell_value = None
                            elif isinstance(cell_value, pd.Timestamp) or hasattr(cell_value, 'isoformat'):
                                cell_value = cell_value.isoformat()
                            elif not isinstance(cell_value, (str, int, float, bool, type(None))):
                                cell_value = str(cell_value)
                            row[str(j)] = cell_value
                        data.append(row)
                    
                    # Get metadata
                    rows, cols = df.shape
                    
                    # Send updated data
                    try:
                        await manager.send_message(
                            client_id,
                            json.dumps({
                                "type": "excel_update",
                                "data": data,
                                "metadata": {
                                    "rows": rows,
                                    "columns": cols
                                }
                            })
                        )
                    except TypeError as e:
                        print(f"JSON serialization error in excel_update: {str(e)}")
            except Exception as e:
                print(f"Error in agent processing: {str(e)}")
                await manager.send_message(
                    client_id,
                    json.dumps({
                        "response": f"Error processing request: {str(e)}",
                        "excel_modified": False
                    })
                )
    
    except WebSocketDisconnect:
        manager.disconnect(client_id)
    except Exception as e:
        print(f"Error: {str(e)}")
        manager.disconnect(client_id)

@app.on_event("startup")
async def startup_event():
    # Create uploads directory if it doesn't exist
    os.makedirs(UPLOAD_DIR, exist_ok=True)

if __name__ == "__main__":
    uvicorn.run("app.main:app", host="0.0.0.0", port=8000, reload=True)
