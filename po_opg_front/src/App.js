import React, { useState } from 'react';
import './App.css';

function App() {
  const [file, setFile] = useState(null);
  const [uploading, setUploading] = useState(false);
  const [importing, setImporting] = useState(false);
  const [parsedPOs, setParsedPOs] = useState(null);
  const [selectedPos, setSelectedPos] = useState([]);
  const [importResults, setImportResults] = useState(null);
  const [error, setError] = useState(null);
  const [expandedPo, setExpandedPo] = useState({});
  const [searchTerm, setSearchTerm] = useState('');

  const handleFileChange = (e) => {
    if (e.target.files && e.target.files[0]) {
      setFile(e.target.files[0]);
      setError(null);
      setParsedPOs(null);
      setImportResults(null);
    }
  };

  const handleDragOver = (e) => {
    e.preventDefault();
  };

  const handleDrop = (e) => {
    e.preventDefault();
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      setFile(e.dataTransfer.files[0]);
      setError(null);
      setParsedPOs(null);
      setImportResults(null);
    }
  };

  const handleUpload = async () => {
    if (!file) {
      setError('Please select an Excel file first.');
      return;
    }

    setUploading(true);
    setError(null);
    setParsedPOs(null);
    setImportResults(null);

    const formData = new FormData();
    formData.append('file', file);

    try {
      const response = await fetch('http://localhost:5000/api/upload', {
        method: 'POST',
        body: formData,
      });

      const data = await response.json();

      if (!response.ok) {
        throw new Error(data.error || 'Failed to process file');
      }

      setParsedPOs(data.poList);
      // Select all by default
      setSelectedPos(data.poList.map(po => po.poNumber));
    } catch (err) {
      setError(err.message || 'An error occurred while uploading the file.');
    } finally {
      setUploading(false);
    }
  };

  const handleImport = async () => {
    if (selectedPos.length === 0) {
      setError('Please select at least one Purchase Order to import.');
      return;
    }

    setImporting(true);
    setError(null);

    const posToImport = parsedPOs.filter(po => selectedPos.includes(po.poNumber));
    
    // Initialize results state for lively table
    const initialResults = posToImport.map(po => ({
      poNumber: po.poNumber,
      docType: po.docType,
      vendorName: po.vendorName,
      itemCount: po.itemCount,
      status: 'pending',
      xmlGenerated: '',
      tallyResponse: null,
      error: null,
      items: po.items
    }));
    
    setImportResults(initialResults);

    // Process each PO sequentially to show progress in real-time
    for (let i = 0; i < posToImport.length; i++) {
      const po = posToImport[i];
      
      // Mark current PO as processing
      setImportResults(prev => prev.map((item, idx) => 
        idx === i ? { ...item, status: 'processing' } : item
      ));

      try {
        const response = await fetch('http://localhost:5000/api/import', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({ selectedPOs: [po] }),
        });

        const data = await response.json();

        if (!response.ok) {
          throw new Error(data.error || 'Import failed');
        }

        const poResult = data.results[0];

        setImportResults(prev => prev.map((item, idx) => 
          idx === i ? { 
            ...item, 
            status: poResult.status,
            xmlGenerated: poResult.xmlGenerated,
            tallyResponse: poResult.tallyResponse,
            error: poResult.error
          } : item
        ));
      } catch (err) {
        setImportResults(prev => prev.map((item, idx) => 
          idx === i ? { 
            ...item, 
            status: 'failed',
            error: err.message || 'Connection Error'
          } : item
        ));
      }
    }

    setImporting(false);
  };

  const togglePoSelection = (poNumber) => {
    setSelectedPos(prev => 
      prev.includes(poNumber) 
        ? prev.filter(num => num !== poNumber) 
        : [...prev, poNumber]
    );
  };

  const toggleDocTypeSelection = (docType) => {
    const posInDocType = parsedPOs.filter(po => po.docType === docType).map(po => po.poNumber);
    const allSelected = posInDocType.every(num => selectedPos.includes(num));

    if (allSelected) {
      // Deselect all POs in this doc type
      setSelectedPos(prev => prev.filter(num => !posInDocType.includes(num)));
    } else {
      // Select all POs in this doc type
      setSelectedPos(prev => {
        const next = [...prev];
        posInDocType.forEach(num => {
          if (!next.includes(num)) {
            next.push(num);
          }
        });
        return next;
      });
    }
  };

  const toggleAllSelection = () => {
    const visiblePoNumbers = parsedPOs ? parsedPOs.filter(po => 
      po.poNumber.toLowerCase().includes(searchTerm.toLowerCase())
    ).map(po => po.poNumber) : [];
    const allVisibleSelected = visiblePoNumbers.every(num => selectedPos.includes(num));

    if (allVisibleSelected) {
      setSelectedPos(prev => prev.filter(num => !visiblePoNumbers.includes(num)));
    } else {
      setSelectedPos(prev => {
        const next = [...prev];
        visiblePoNumbers.forEach(num => {
          if (!next.includes(num)) {
            next.push(num);
          }
        });
        return next;
      });
    }
  };

  const handleReset = () => {
    setFile(null);
    setParsedPOs(null);
    setSelectedPos([]);
    setImportResults(null);
    setError(null);
    setSearchTerm('');
  };

  const togglePoExpand = (poNumber) => {
    setExpandedPo(prev => ({
      ...prev,
      [poNumber]: !prev[poNumber]
    }));
  };

  const copyToClipboard = (text) => {
    navigator.clipboard.writeText(text);
    alert('XML copied to clipboard!');
  };

  const downloadXml = (poNumber, xml) => {
    const blob = new Blob([xml], { type: 'text/xml' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `Purchase_Order_${poNumber}.xml`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  // Extract unique doc types and counts
  const getDocTypeCounts = () => {
    if (!parsedPOs) return {};
    const counts = {};
    parsedPOs.forEach(po => {
      counts[po.docType] = (counts[po.docType] || 0) + 1;
    });
    return counts;
  };

  const filteredPOs = parsedPOs ? parsedPOs.filter(po => 
    po.poNumber.toLowerCase().includes(searchTerm.toLowerCase())
  ) : [];

  const docTypeCounts = getDocTypeCounts();
  const uniqueDocTypes = Object.keys(docTypeCounts);

  return (
    <div className="App">
      <header className="App-header">
        <div className="logo-container">
          <div className="glow-circle"></div>
          <h1>OPG Tally Gateway</h1>
        </div>
        <p className="subtitle">Automated Excel-to-Tally Purchase Order Importer</p>
      </header>

      <main className="container">
        {/* Step Indicators */}
        <div className="steps-indicator">
          <div className={`step-item ${!parsedPOs && !importResults ? 'active' : ''} ${parsedPOs || importResults ? 'completed' : ''}`}>
            <span className="step-num">1</span>
            <span className="step-text">Upload Spreadsheet</span>
          </div>
          <div className="step-line"></div>
          <div className={`step-item ${parsedPOs && !importResults ? 'active' : ''} ${importResults ? 'completed' : ''}`}>
            <span className="step-num">2</span>
            <span className="step-text">Select Sections</span>
          </div>
          <div className="step-line"></div>
          <div className={`step-item ${importResults ? 'active' : ''}`}>
            <span className="step-num">3</span>
            <span className="step-text">Import Results</span>
          </div>
        </div>

        {/* ERROR ALERT */}
        {error && <div className="alert alert-danger">{error}</div>}

        {/* STEP 1: UPLOAD CARD */}
        {!parsedPOs && !importResults && (
          <div className="card upload-card">
            <h2>Upload PO Spreadsheet</h2>
            <p className="card-desc">Select or drag & drop your purchase order excel sheet containing the Details sheet.</p>
            
            <div 
              className={`dropzone ${file ? 'has-file' : ''}`}
              onDragOver={handleDragOver}
              onDrop={handleDrop}
            >
              <input 
                type="file" 
                id="fileInput" 
                accept=".xlsx, .xls" 
                onChange={handleFileChange} 
                className="hidden-input"
              />
              <label htmlFor="fileInput" className="dropzone-label">
                <svg className="upload-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
                  <polyline points="17 8 12 3 7 8" />
                  <line x1="12" y1="3" x2="12" y2="15" />
                </svg>
                {file ? (
                  <div className="file-info">
                    <span className="file-name">{file.name}</span>
                    <span className="file-size">{(file.size / 1024).toFixed(2)} KB</span>
                  </div>
                ) : (
                  <span>Drag & drop file here or <strong>browse files</strong></span>
                )}
              </label>
            </div>

            <button 
              className={`btn btn-primary ${uploading ? 'loading' : ''}`} 
              onClick={handleUpload}
              disabled={uploading || !file}
            >
              {uploading ? (
                <span className="spinner-container">
                  <span className="spinner"></span> Processing Spreadsheet...
                </span>
              ) : 'Parse Excel Document'}
            </button>
          </div>
        )}

        {/* STEP 2: SELECTION CARD */}
        {parsedPOs && !importResults && (
          <div className="selection-container">
            <div className="card selection-controls-card">
              <div className="card-header-flex">
                <div>
                  <h2>Select Purchase Orders to Import</h2>
                  <p className="card-desc">Select which PO doc types or individual orders to import into Tally.</p>
                </div>
                <button className="btn-sm btn-outline-danger" onClick={handleReset}>
                  Cancel / Upload New
                </button>
              </div>

              {/* Search Box */}
              <div className="search-bar-container">
                <svg className="search-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <circle cx="11" cy="11" r="8"></circle>
                  <line x1="21" y1="21" x2="16.65" y2="16.65"></line>
                </svg>
                <input
                  type="text"
                  className="search-input"
                  placeholder="Search by Purchasing Document (PO Number)..."
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                />
                {searchTerm && (
                  <button className="btn-search-clear" onClick={() => setSearchTerm('')}>
                    &times;
                  </button>
                )}
              </div>

              {/* Selection Sections by Doc Type */}
              <div className="doc-type-filters">
                <label className="checkbox-container select-all-btn">
                  <input 
                    type="checkbox"
                    checked={filteredPOs.length > 0 && filteredPOs.every(po => selectedPos.includes(po.poNumber))}
                    onChange={toggleAllSelection}
                  />
                  <span className="checkmark"></span>
                  <span className="filter-label">Select All Checked/Filtered ({filteredPOs.length} POs)</span>
                </label>

                <div className="filter-divider"></div>

                <div className="doc-type-pills">
                  {uniqueDocTypes.map(docType => {
                    const count = docTypeCounts[docType];
                    const posInDoc = parsedPOs.filter(po => po.docType === docType).map(po => po.poNumber);
                    const isAllDocSelected = posInDoc.every(num => selectedPos.includes(num));
                    const isSomeDocSelected = posInDoc.some(num => selectedPos.includes(num)) && !isAllDocSelected;

                    return (
                      <label 
                        key={docType} 
                        className={`doc-type-pill-label ${isAllDocSelected ? 'checked' : ''} ${isSomeDocSelected ? 'partial' : ''}`}
                      >
                        <input 
                          type="checkbox"
                          checked={isAllDocSelected}
                          ref={el => {
                            if (el) el.indeterminate = isSomeDocSelected;
                          }}
                          onChange={() => toggleDocTypeSelection(docType)}
                        />
                        <span className="doc-type-badge">{docType}</span>
                        <span className="doc-type-count">{count} POs</span>
                      </label>
                    );
                  })}
                </div>
              </div>

              <button 
                className={`btn btn-primary ${importing ? 'loading' : ''}`}
                onClick={handleImport}
                disabled={importing || selectedPos.length === 0}
              >
                {importing ? (
                  <span className="spinner-container">
                    <span className="spinner"></span> Importing to Tally...
                  </span>
                ) : `Import Selected Vouchers (${selectedPos.length} of ${parsedPOs.length})`}
              </button>
            </div>

            {/* List of individual POs */}
            <div className="po-grid">
              {filteredPOs.map(po => {
                const isSelected = selectedPos.includes(po.poNumber);
                return (
                  <div key={po.poNumber} className={`po-select-card ${isSelected ? 'selected' : ''}`}>
                    <div className="po-select-header">
                      <label className="checkbox-container">
                        <input 
                          type="checkbox"
                          checked={isSelected}
                          onChange={() => togglePoSelection(po.poNumber)}
                        />
                        <span className="checkmark"></span>
                      </label>

                      <div className="po-select-info">
                        <div className="po-title-row">
                          <span className="po-num-text">PO #{po.poNumber}</span>
                          <span className="doc-type-pill">{po.docType}</span>
                        </div>
                        <span className="po-vendor-text">{po.vendorName}</span>
                      </div>

                      <div className="po-select-meta">
                        <span className="po-items-count">{po.itemCount} items</span>
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {/* STEP 3: LIVELY IMPORT PROGRESS TABLE */}
        {importResults && (
          <div className="results-container">
            <div className="card results-summary-card">
              <div className="results-header">
                <h2>Import Live Dashboard</h2>
                <div className="results-summary-badges">
                  <span className="badge badge-success">
                    {importResults.filter(r => r.status === 'success').length} Imported
                  </span>
                  {importResults.filter(r => r.status === 'failed').length > 0 && (
                    <span className="badge badge-failed">
                      {importResults.filter(r => r.status === 'failed').length} Failed
                    </span>
                  )}
                  {importResults.filter(r => r.status === 'processing').length > 0 && (
                    <span className="badge badge-processing">
                      Processing...
                    </span>
                  )}
                </div>
              </div>
              <p className="card-desc">
                {importing 
                  ? `Importing selected vouchers to Tally... (${importResults.filter(r => ['success', 'failed'].includes(r.status)).length} of ${importResults.length} complete)`
                  : `Tally import completed. ${importResults.filter(r => r.status === 'success').length} successfully imported.`}
              </p>
              
              {!importing && (
                <div className="action-buttons-container">
                  <button className="btn btn-primary" onClick={handleReset}>
                    Clear & Start New Import
                  </button>
                </div>
              )}
            </div>

            {/* Lively Status Table */}
            <div className="card live-table-card">
              <div className="table-responsive">
                <table className="live-status-table">
                  <thead>
                    <tr>
                      <th>PO Number</th>
                      <th>Doc Type</th>
                      <th>Vendor</th>
                      <th>Items</th>
                      <th>Status</th>
                      <th>Actions</th>
                    </tr>
                  </thead>
                  <tbody>
                    {importResults.map((po) => {
                      const isExpanded = !!expandedPo[po.poNumber];
                      return (
                        <React.Fragment key={po.poNumber}>
                          <tr className={`table-row-${po.status} ${isExpanded ? 'row-expanded-header' : ''}`}>
                            <td className="po-num-cell">
                              <strong>#{po.poNumber}</strong>
                            </td>
                            <td>
                              <span className="doc-type-pill">{po.docType}</span>
                            </td>
                            <td className="vendor-cell">{po.vendorName}</td>
                            <td>{po.itemCount}</td>
                            <td>
                              <div className={`status-indicator ${po.status}`}>
                                {po.status === 'pending' && <span className="status-dot dot-pending"></span>}
                                {po.status === 'processing' && <span className="status-spinner"></span>}
                                {po.status === 'success' && <span className="status-dot dot-success"></span>}
                                {po.status === 'failed' && <span className="status-dot dot-failed"></span>}
                                <span className="status-text-cap">
                                  {po.status === 'success' ? 'Imported' : po.status}
                                </span>
                              </div>
                            </td>
                            <td>
                              <div className="table-actions">
                                {['success', 'failed'].includes(po.status) && (
                                  <>
                                    <button className="btn-table-icon" onClick={() => togglePoExpand(po.poNumber)} title="Toggle XML View">
                                      <svg className={`chevron ${isExpanded ? 'open' : ''}`} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                                        <polyline points="6 9 12 15 18 9" />
                                      </svg>
                                    </button>
                                    {po.xmlGenerated && (
                                      <button className="btn-table-icon" onClick={() => downloadXml(po.poNumber, po.xmlGenerated)} title="Download XML">
                                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="16" height="16">
                                          <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
                                          <polyline points="7 10 12 15 17 10" />
                                          <line x1="12" y1="15" x2="12" y2="3" />
                                        </svg>
                                      </button>
                                    )}
                                  </>
                                )}
                                {po.status === 'pending' && <span className="text-muted">-</span>}
                              </div>
                            </td>
                          </tr>

                          {/* Expandable XML view row */}
                          {isExpanded && (
                            <tr className="expanded-details-row">
                              <td colSpan="6">
                                <div className="po-card-details table-expanded-details">
                                  {po.status === 'failed' && (
                                    <div className="tally-error">
                                      <strong>Tally Response / Connection Error:</strong> {po.error || 'Tally Client is Offline. Review generated XML below.'}
                                    </div>
                                  )}
                                  <div className="xml-actions">
                                    <h4>Tally XML Payload</h4>
                                    <button className="btn-sm" onClick={() => copyToClipboard(po.xmlGenerated)}>
                                      Copy XML
                                    </button>
                                  </div>
                                  <pre className="xml-preview">
                                    <code>{po.xmlGenerated}</code>
                                  </pre>
                                </div>
                              </td>
                            </tr>
                          )}
                        </React.Fragment>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}

export default App;
