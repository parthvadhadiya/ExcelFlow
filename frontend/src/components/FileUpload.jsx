import React, { useState, useRef } from 'react';
import { FiUpload } from 'react-icons/fi';

const FileUpload = ({ onFileUpload, loading, error }) => {
  const [selectedFile, setSelectedFile] = useState(null);
  const fileInputRef = useRef(null);

  // Define supported file types
  const supportedFormats = {
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': '.xlsx',
    'application/vnd.ms-excel': '.xls',
    'application/vnd.oasis.opendocument.spreadsheet': '.ods',
    'text/csv': '.csv',
    'text/tab-separated-values': '.tsv',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.template': '.xltx',
    'application/vnd.ms-excel.template.macroEnabled.12': '.xltm',
    'application/vnd.ms-excel.sheet.macroEnabled.12': '.xlsm'
  };

  // Get file extensions for accept attribute
  const acceptedExtensions = Object.values(supportedFormats).join(',');
  
  // Check if file type is supported
  const isFileSupported = (file) => {
    // Check by MIME type
    if (supportedFormats[file.type]) {
      return true;
    }
    
    // Fallback to extension check if MIME type is not recognized
    const extension = file.name.toLowerCase().split('.').pop();
    return ['xlsx', 'xls', 'csv', 'tsv', 'ods', 'fods', 'xlsm', 'xltx', 'xltm'].includes(extension);
  };

  const handleFileChange = (event) => {
    const file = event.target.files[0];
    if (file && isFileSupported(file)) {
      setSelectedFile(file);
    } else {
      setSelectedFile(null);
      alert(`Please select a valid spreadsheet file (${acceptedExtensions})`);
    }
  };

  const handleDragOver = (event) => {
    event.preventDefault();
  };

  const handleDrop = (event) => {
    event.preventDefault();
    const file = event.dataTransfer.files[0];
    if (file && isFileSupported(file)) {
      setSelectedFile(file);
    } else {
      alert(`Please drop a valid spreadsheet file (${acceptedExtensions})`);
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
        accept={acceptedExtensions}
        id="file-input"
      />
      <label htmlFor="file-input">
        <FiUpload />
        <h2>Upload Spreadsheet</h2>
        <p>Drag & drop your spreadsheet file here or click to browse</p>
        <p className="supported-formats">
          Supported formats: XLSX, XLS, CSV, TSV, ODS and more
        </p>
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
