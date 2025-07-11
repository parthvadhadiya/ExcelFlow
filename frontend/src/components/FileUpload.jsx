import React, { useState, useRef } from 'react';
import { FiUpload } from 'react-icons/fi';

const FileUpload = ({ onFileUpload, loading, error }) => {
  const [selectedFile, setSelectedFile] = useState(null);
  const fileInputRef = useRef(null);

  const handleFileChange = (event) => {
    const file = event.target.files[0];
    if (file && file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
      setSelectedFile(file);
    } else {
      setSelectedFile(null);
      alert('Please select a valid Excel file (.xlsx)');
    }
  };

  const handleDragOver = (event) => {
    event.preventDefault();
  };

  const handleDrop = (event) => {
    event.preventDefault();
    const file = event.dataTransfer.files[0];
    if (file && file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
      setSelectedFile(file);
    } else {
      alert('Please drop a valid Excel file (.xlsx)');
    }
  };

  const handleUpload = () => {
    if (selectedFile) {
      onFileUpload(selectedFile);
    }
  };

  return (
    <div 
      className="file-upload"
      onDragOver={handleDragOver}
      onDrop={handleDrop}
    >
      <input
        type="file"
        ref={fileInputRef}
        onChange={handleFileChange}
        accept=".xlsx"
        id="file-input"
      />
      <label htmlFor="file-input">
        <FiUpload />
        <h2>Upload Excel File</h2>
        <p>Drag & drop your Excel file here or click to browse</p>
        {selectedFile && (
          <p className="selected-file">Selected: {selectedFile.name}</p>
        )}
      </label>
      {error && <p className="error">{error}</p>}
      <button 
        onClick={handleUpload} 
        disabled={!selectedFile || loading}
      >
        {loading ? 'Uploading...' : 'Upload'}
      </button>
    </div>
  );
};

export default FileUpload;
