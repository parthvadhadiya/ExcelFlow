import React, { useState, useEffect } from 'react';
import FileUpload from './components/FileUpload';
import ExcelViewer from './components/ExcelViewer';
import ChatInterface from './components/ChatInterface';
import './App.css';

function App() {
  const [clientId, setClientId] = useState(null);
  const [excelData, setExcelData] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [socket, setSocket] = useState(null);
  const [messages, setMessages] = useState([]);

  // Initialize WebSocket connection when clientId is set
  useEffect(() => {
    if (!clientId) return;

    // For WebSocket connections, we still need to use the full URL
    const ws = new WebSocket(`ws://localhost:8000/ws/${clientId}`);
    
    ws.onopen = () => {
      console.log('WebSocket connected');
    };
    
    ws.onmessage = (event) => {
      const data = JSON.parse(event.data);
      
      if (data.type === 'excel_update') {
        // Update Excel data when changes are made
        setExcelData({
          data: data.data,
          metadata: data.metadata
        });
      } else {
        // Handle chat messages
        setMessages(prev => [...prev, {
          role: 'assistant',
          content: data.response
        }]);
        
        // If Excel was modified, fetch the latest data
        if (data.excel_modified) {
          fetchExcelData();
        }
      }
    };
    
    ws.onclose = () => {
      console.log('WebSocket disconnected');
    };
    
    ws.onerror = (error) => {
      console.error('WebSocket error:', error);
      setError('Connection error. Please try again.');
    };
    
    setSocket(ws);
    
    // Fetch initial Excel data
    fetchExcelData();
    
    // Clean up WebSocket connection on component unmount
    return () => {
      if (ws) {
        ws.close();
      }
    };
  }, [clientId]);

  const fetchExcelData = async () => {
    try {
      const response = await fetch(`http://localhost:8000/excel/${clientId}`);
      if (!response.ok) {
        throw new Error('Failed to fetch Excel data');
      }
      const data = await response.json();
      setExcelData(data);
    } catch (error) {
      console.error('Error fetching Excel data:', error);
      setError('Failed to load Excel data. Please try again.');
    }
  };

  const handleFileUpload = async (file) => {
    console.log('Uploading file:', file.name);
    setLoading(true);
    setError(null);
    
    try {
      const formData = new FormData();
      formData.append('file', file);
      
      console.log('Sending file to backend...');
      const response = await fetch('http://localhost:8000/upload', {
        method: 'POST',
        body: formData,
      });
      
      console.log('Response status:', response.status);
      
      if (!response.ok) {
        const errorText = await response.text();
        console.error('Error response:', errorText);
        throw new Error(`Failed to upload file: ${response.status} ${errorText}`);
      }
      
      const data = await response.json();
      console.log('Upload successful, client ID:', data.client_id);
      setClientId(data.client_id);
    } catch (error) {
      console.error('Error uploading file:', error);
      setError('Failed to upload file. Please try again.');
    } finally {
      setLoading(false);
    }
  };

  const sendMessage = (message) => {
    if (!socket || socket.readyState !== WebSocket.OPEN) {
      setError('Connection lost. Please refresh the page.');
      return;
    }
    
    // Add user message to the chat
    setMessages(prev => [...prev, {
      role: 'user',
      content: message
    }]);
    
    // Send message to the server
    socket.send(JSON.stringify({ message }));
  };

  return (
    <div className="app-container">
      {!clientId ? (
        <FileUpload onFileUpload={handleFileUpload} loading={loading} error={error} />
      ) : (
        <div className="main-content">
          <div className="excel-container">
            {excelData ? (
              <ExcelViewer data={excelData.data} metadata={excelData.metadata} />
            ) : (
              <div className="loading">Loading Excel data...</div>
            )}
          </div>
          <div className="chat-sidebar">
            <ChatInterface 
              messages={messages} 
              onSendMessage={sendMessage} 
            />
          </div>
        </div>
      )}
    </div>
  );
}

export default App;
