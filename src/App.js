import React, { useEffect, useState } from 'react';
import './App.css';

function App() {
  const [files, setFiles] = useState([]);
  const [selectedFile, setSelectedFile] = useState('');
  const [fieldName, setFieldName] = useState('');
  const [fieldType, setFieldType] = useState('');
  const [visibilityRules, setVisibilityRules] = useState('');
  const [visibilityAttributes, setVisibilityAttributes] = useState('');
  const [results, setResults] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);

  useEffect(() => {
    fetch('/api/excel-files')
      .then((res) => res.json())
      .then((data) => setFiles(data.files || []))
      .catch((err) => setError('Failed to load files'));
  }, []);

  const handleSearch = async (e) => {
    e.preventDefault();
    setLoading(true);
    setError(null);
    setResults([]);
    try {
      const payload = {
        fieldName,
        fieldType,
        visibilityRules,
        visibilityAttributes,
      };
      if (selectedFile) payload.fileName = selectedFile;

      const res = await fetch('/api/search-excel', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });
      const data = await res.json();
      setResults(data.results || []);
    } catch (err) {
      setError('Search failed');
    } finally {
      setLoading(false);
    }
  };

  const allColumns = () => {
    const cols = new Set();
    results.forEach((r) => Object.keys(r).forEach((k) => cols.add(k)));
    return Array.from(cols);
  };

  return (
    <div className="App">
      <header className="App-header">
        <h1>Excel Search</h1>
        <form className="search-form" onSubmit={handleSearch}>
          <label>
            File:
            <select value={selectedFile} onChange={(e) => setSelectedFile(e.target.value)}>
              <option value="">-- All files --</option>
              {files.map((f) => (
                <option key={f.id} value={f.id}>
                  {f.name}
                </option>
              ))}
            </select>
          </label>

          <label>
            Field Name:
            <input value={fieldName} onChange={(e) => setFieldName(e.target.value)} placeholder="Field Name" />
          </label>

          <label>
            Field Type:
            <input value={fieldType} onChange={(e) => setFieldType(e.target.value)} placeholder="Field Type" />
          </label>

          <label>
            Visibility Rules:
            <input value={visibilityRules} onChange={(e) => setVisibilityRules(e.target.value)} placeholder="Visibility Rules" />
          </label>

          <label>
            Visibility Attributes:
            <input value={visibilityAttributes} onChange={(e) => setVisibilityAttributes(e.target.value)} placeholder="Visibility Attributes" />
          </label>

          <div className="actions">
            <button type="submit" disabled={loading}>
              {loading ? 'Searching…' : 'Search'}
            </button>
            <button
              type="button"
              onClick={() => {
                setFieldName('');
                setFieldType('');
                setVisibilityRules('');
                setVisibilityAttributes('');
                setSelectedFile('');
                setResults([]);
                setError(null);
              }}
            >
              Reset
            </button>
          </div>
        </form>

        {error && <div className="error">{error}</div>}

        <div className="results">
          {results.length === 0 && !loading && <div className="empty">No results to show.</div>}

          {results.length > 0 && (
            <table>
              <thead>
                <tr>
                  {allColumns().map((col) => (
                    <th key={col}>{col}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {results.map((row, idx) => (
                  <tr key={idx}>
                    {allColumns().map((col) => (
                      <td key={col}>{row[col] != null ? String(row[col]) : ''}</td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>
      </header>
    </div>
  );
}

export default App;
