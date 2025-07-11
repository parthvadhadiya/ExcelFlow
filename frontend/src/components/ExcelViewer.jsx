import React, { useState, useEffect } from 'react';
import { DataGrid } from 'react-data-grid';
import 'react-data-grid/lib/styles.css';

const ExcelViewer = ({ data, metadata }) => {
  const [columns, setColumns] = useState([]);
  const [rows, setRows] = useState([]);

  useEffect(() => {
    if (!data || !metadata) return;

    // Create column definitions
    const cols = Array.from({ length: metadata.columns }, (_, i) => ({
      key: i.toString(),
      name: getExcelColumnName(i),
      resizable: true,
      sortable: true,
      width: 120
    }));
    
    // Transform data into rows format expected by react-data-grid
    const formattedRows = data.map((row, rowIndex) => {
      return {
        id: rowIndex,
        ...row
      };
    });

    setColumns(cols);
    setRows(formattedRows);
  }, [data, metadata]);

  // Convert column index to Excel column name (A, B, C, ..., Z, AA, AB, etc.)
  const getExcelColumnName = (index) => {
    let columnName = '';
    let tempIndex = index;
    
    while (tempIndex >= 0) {
      const remainder = tempIndex % 26;
      columnName = String.fromCharCode(65 + remainder) + columnName;
      tempIndex = Math.floor(tempIndex / 26) - 1;
    }
    
    return columnName;
  };

  if (!data || !metadata) {
    return <div>Loading Excel data...</div>;
  }

  return (
    <div className="excel-viewer">
      <h2 className="text-xl font-bold mb-4">Excel Spreadsheet</h2>
      <div style={{ height: 'calc(100vh - 100px)', width: '100%' }}>
        <DataGrid
          columns={columns}
          rows={rows}
          rowHeight={35}
          className="rdg-light"
        />
      </div>
    </div>
  );
};

export default ExcelViewer;
