'use client';

import { useState, useRef } from 'react';
import * as XLSX from 'xlsx';

interface ExcelData {
  headers: string[];
  rows: any[][];
  sheetNames: string[];
  stats: {
    totalRows: number;
    totalColumns: number;
    totalSheets: number;
  };
}

export default function Home() {
  const [file, setFile] = useState<File | null>(null);
  const [data, setData] = useState<ExcelData | null>(null);
  const [currentSheet, setCurrentSheet] = useState<string>('');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string>('');
  const [dragOver, setDragOver] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFileSelect = (selectedFile: File) => {
    if (!selectedFile) return;

    const validTypes = [
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'application/vnd.ms-excel',
      'text/csv'
    ];

    if (!validTypes.includes(selectedFile.type) &&
        !selectedFile.name.endsWith('.xlsx') &&
        !selectedFile.name.endsWith('.xls') &&
        !selectedFile.name.endsWith('.csv')) {
      setError('कृपया सही Excel फाइल चुनें (.xlsx, .xls, .csv)');
      return;
    }

    setFile(selectedFile);
    setError('');
    extractData(selectedFile);
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFile = e.target.files?.[0];
    if (selectedFile) {
      handleFileSelect(selectedFile);
    }
  };

  const handleDrop = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setDragOver(false);
    const droppedFile = e.dataTransfer.files[0];
    if (droppedFile) {
      handleFileSelect(droppedFile);
    }
  };

  const handleDragOver = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setDragOver(true);
  };

  const handleDragLeave = () => {
    setDragOver(false);
  };

  const extractData = async (file: File) => {
    setLoading(true);
    setError('');

    try {
      const arrayBuffer = await file.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, { type: 'array' });

      const sheetNames = workbook.SheetNames;
      const firstSheetName = sheetNames[0];
      setCurrentSheet(firstSheetName);

      const worksheet = workbook.Sheets[firstSheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];

      if (jsonData.length === 0) {
        setError('फाइल खाली है');
        setLoading(false);
        return;
      }

      const headers = jsonData[0] as string[];
      const rows = jsonData.slice(1);

      setData({
        headers,
        rows,
        sheetNames,
        stats: {
          totalRows: rows.length,
          totalColumns: headers.length,
          totalSheets: sheetNames.length
        }
      });

    } catch (err) {
      setError('फाइल पढ़ने में त्रुटि हुई');
      console.error(err);
    } finally {
      setLoading(false);
    }
  };

  const handleSheetChange = async (sheetName: string) => {
    if (!file) return;

    setCurrentSheet(sheetName);
    setLoading(true);

    try {
      const arrayBuffer = await file.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, { type: 'array' });
      const worksheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];

      const headers = jsonData[0] as string[];
      const rows = jsonData.slice(1);

      setData(prev => prev ? {
        ...prev,
        headers,
        rows,
        stats: {
          ...prev.stats,
          totalRows: rows.length,
          totalColumns: headers.length
        }
      } : null);

    } catch (err) {
      setError('शीट लोड करने में त्रुटि हुई');
      console.error(err);
    } finally {
      setLoading(false);
    }
  };

  const downloadAsCSV = () => {
    if (!data) return;

    const csvContent = [
      data.headers.join(','),
      ...data.rows.map(row => row.join(','))
    ].join('\n');

    const blob = new Blob([csvContent], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${file?.name.replace(/\.[^/.]+$/, '')}_${currentSheet}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const downloadAsJSON = () => {
    if (!data) return;

    const jsonData = data.rows.map(row => {
      const obj: any = {};
      data.headers.forEach((header, index) => {
        obj[header] = row[index];
      });
      return obj;
    });

    const blob = new Blob([JSON.stringify(jsonData, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${file?.name.replace(/\.[^/.]+$/, '')}_${currentSheet}.json`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const removeFile = () => {
    setFile(null);
    setData(null);
    setCurrentSheet('');
    setError('');
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  return (
    <div className="container">
      <div className="header">
        <h1>📊 Excel Data Extractor</h1>
        <p>अपनी Excel फाइल से आसानी से डेटा निकालें</p>
      </div>

      <div className="card">
        {!file ? (
          <div
            className={`upload-area ${dragOver ? 'dragover' : ''}`}
            onClick={() => fileInputRef.current?.click()}
            onDrop={handleDrop}
            onDragOver={handleDragOver}
            onDragLeave={handleDragLeave}
          >
            <div className="upload-icon">📁</div>
            <div className="upload-text">फाइल चुनने के लिए क्लिक करें या यहाँ खींचें</div>
            <div className="upload-hint">Excel (.xlsx, .xls) या CSV फाइलें समर्थित हैं</div>
            <input
              ref={fileInputRef}
              type="file"
              className="file-input"
              accept=".xlsx,.xls,.csv"
              onChange={handleFileChange}
            />
          </div>
        ) : (
          <div className="file-info">
            <span className="file-name">📄 {file.name}</span>
            <button className="remove-btn" onClick={removeFile}>हटाएं</button>
          </div>
        )}

        {error && <div className="error">{error}</div>}

        {loading && <div className="loading">⏳ डेटा लोड हो रहा है...</div>}

        {data && !loading && (
          <>
            <div className="stats">
              <div className="stat-card">
                <div className="stat-value">{data.stats.totalRows}</div>
                <div className="stat-label">कुल पंक्तियाँ</div>
              </div>
              <div className="stat-card">
                <div className="stat-value">{data.stats.totalColumns}</div>
                <div className="stat-label">कुल कॉलम</div>
              </div>
              <div className="stat-card">
                <div className="stat-value">{data.stats.totalSheets}</div>
                <div className="stat-label">कुल शीट्स</div>
              </div>
            </div>

            {data.sheetNames.length > 1 && (
              <div className="sheet-selector">
                <label>शीट चुनें:</label>
                <select value={currentSheet} onChange={(e) => handleSheetChange(e.target.value)}>
                  {data.sheetNames.map((name) => (
                    <option key={name} value={name}>{name}</option>
                  ))}
                </select>
              </div>
            )}

            <div className="actions">
              <button className="btn" onClick={downloadAsCSV}>
                💾 CSV के रूप में डाउनलोड करें
              </button>
              <button className="btn" onClick={downloadAsJSON}>
                💾 JSON के रूप में डाउनलोड करें
              </button>
            </div>

            <div className="table-container">
              <table>
                <thead>
                  <tr>
                    {data.headers.map((header, index) => (
                      <th key={index}>{header || `Column ${index + 1}`}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {data.rows.length > 0 ? (
                    data.rows.map((row, rowIndex) => (
                      <tr key={rowIndex}>
                        {row.map((cell, cellIndex) => (
                          <td key={cellIndex}>{cell !== undefined && cell !== null ? String(cell) : ''}</td>
                        ))}
                      </tr>
                    ))
                  ) : (
                    <tr>
                      <td colSpan={data.headers.length} className="empty-state">
                        कोई डेटा नहीं मिला
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </>
        )}
      </div>
    </div>
  );
}
