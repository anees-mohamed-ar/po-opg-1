import React, { useState, useRef, useCallback, useMemo } from 'react';
import './App.css';
import XmlVisualizer from './XmlVisualizer';
import DayBookVisualizer from './DayBookVisualizer';

function App() {
  const [isSidebarCollapsed, setIsSidebarCollapsed] = useState(false);
  const [activeMode, setActiveMode] = useState('importer'); // 'importer' | 'visualizer' | 'daybook'
  const [importType, setImportType] = useState('po'); // 'po' | 'grn'
  const [file, setFile] = useState(null);
  const [uploading, setUploading] = useState(false);
  const [parsingProgress, setParsingProgress] = useState(null); // { message, rowsParsed, totalRows }
  const [importing, setImporting] = useState(false);
  const [parsedPOs, setParsedPOs] = useState(null);
  const [selectedPos, setSelectedPos] = useState(new Set());
  const [importResults, setImportResults] = useState(null);
  const [error, setError] = useState(null);
  const [expandedPo, setExpandedPo] = useState({});
  const [searchTerm, setSearchTerm] = useState('');
  const [resultSearchTerm, setResultSearchTerm] = useState('');
  const [currentPage, setCurrentPage] = useState(1);
  const [rowsPerPage, setRowsPerPage] = useState(10);
  const [importTimeTaken, setImportTimeTaken] = useState(null);
  const [isDragging, setIsDragging] = useState(false);
  const [copySuccess, setCopySuccess] = useState(null);
  const [skipBlankMaterial, setSkipBlankMaterial] = useState(false);
  const [rangeFrom, setRangeFrom] = useState(1);
  const [rangeTo, setRangeTo] = useState(100);
  const [rangeCategory, setRangeCategory] = useState('ALL');
  const [showSummaryModal, setShowSummaryModal] = useState(false);
  const [onlyMissingMaterial, setOnlyMissingMaterial] = useState(false);
  const [ignoredInvoices, setIgnoredInvoices] = useState([]);
  const [showIgnoredModal, setShowIgnoredModal] = useState(false);

  const fileInputRef = useRef(null);

  const handleDropzoneClick = () => {
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
      fileInputRef.current.click();
    }
  };

  const handleFileChange = (e) => {
    if (e.target.files && e.target.files[0]) {
      setFile(e.target.files[0]);
      setError(null);
      setParsedPOs(null);
      setImportResults(null);
    }
  };

  const handleDragOver = useCallback((e) => {
    e.preventDefault();
    setIsDragging(true);
  }, []);

  const handleDragLeave = useCallback((e) => {
    e.preventDefault();
    setIsDragging(false);
  }, []);

  const handleDrop = useCallback((e) => {
    e.preventDefault();
    setIsDragging(false);
    const dropped = e.dataTransfer.files?.[0];
    if (dropped && (dropped.name.endsWith('.xlsx') || dropped.name.endsWith('.xls'))) {
      setFile(dropped);
      setError(null);
      setParsedPOs(null);
      setImportResults(null);
    } else if (dropped) {
      setError('Please drop a valid Excel file (.xlsx or .xls).');
    }
  }, []);

  const handleUpload = async () => {
    if (!file) {
      setError('Please select an Excel file first.');
      return;
    }

    setUploading(true);
    setError(null);
    setParsedPOs(null);
    setImportResults(null);
    setParsingProgress({ message: 'Uploading file and reading sheets...', rowsParsed: 0, totalRows: 0 });

    const formData = new FormData();
    formData.append('file', file);

    try {
      const response = await fetch(`http://192.168.1.166:5001/api/upload?importType=${importType}`, {
        method: 'POST',
        body: formData,
      });

      const contentType = response.headers.get('content-type') || '';
      if (contentType.includes('text/html') || contentType.includes('application/xhtml+xml')) {
        throw new Error('Received HTML response instead of JSON. The backend server might be offline or misconfigured.');
      }

      if (!response.ok) {
        const text = await response.text();
        let errMsg = 'Failed to process file';
        try {
          const parsed = JSON.parse(text);
          errMsg = parsed.error || errMsg;
        } catch (_) {
          errMsg = text || errMsg;
        }
        throw new Error(errMsg);
      }

      const reader = response.body.getReader();
      const decoder = new TextDecoder();
      let buffer = '';
      let finalData = null;

      while (true) {
        const { done, value } = await reader.read();
        if (done) break;

        buffer += decoder.decode(value, { stream: true });
        const lines = buffer.split('\n');
        buffer = lines.pop(); // Keep the last incomplete chunk

        for (const line of lines) {
          if (!line.trim()) continue;
          try {
            const msg = JSON.parse(line);
            if (msg.type === 'progress') {
              setParsingProgress({
                message: msg.message,
                rowsParsed: msg.rowsParsed,
                totalRows: msg.totalRows
              });
            } else if (msg.type === 'error') {
              throw new Error(msg.error);
            } else if (msg.type === 'result') {
              finalData = msg;
            }
          } catch (e) {
            console.error('Error parsing line:', line, e);
            if (e.message) throw e;
          }
        }
      }

      if (!finalData || !finalData.poList) {
        throw new Error('Upload completed without returning parsed data.');
      }

      setParsedPOs(finalData.poList);
      setIgnoredInvoices(finalData.ignoredPurchaseInvoices || []);
      setRangeFrom(1);
      const initialTo = Math.min(100, finalData.poList.length);
      setRangeTo(initialTo);
      
      const newSelected = new Set();
      for (let i = 0; i < initialTo; i++) {
        if (finalData.poList[i]) {
          const po = finalData.poList[i];
          const activeCount = po.items ? po.items.filter(item => {
            const d = String(item['Deletion Indicator'] || item['deletion indicator'] || '').trim().toUpperCase();
            return d !== 'L';
          }).length : 0;
          if (activeCount > 0) {
            newSelected.add(po.poNumber);
          }
        }
      }
      setSelectedPos(newSelected);
    } catch (err) {
      if (err.message && (err.message.toLowerCase().includes('fetch') || err.name === 'TypeError')) {
        setError('Cannot connect to the backend server. Please verify that the backend server is running on port 5001.');
      } else {
        setError(err.message || 'An error occurred while uploading the file.');
      }
    } finally {
      setUploading(false);
      setParsingProgress(null);
    }
  };

  const [etaSeconds, setEtaSeconds] = useState(null);

  const handleImport = async () => {
    if (selectedPos.size === 0) {
      setError('Please select at least one Purchase Order to import.');
      return;
    }

    setImporting(true);
    setError(null);
    setImportTimeTaken(null);
    setEtaSeconds(null);
    const startTime = Date.now();

    const posToImport = parsedPOs.filter(po => selectedPos.has(po.poNumber));

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
    setCurrentPage(1);

    for (let i = 0; i < posToImport.length; i++) {
      const po = posToImport[i];

      setImportResults(prev => prev.map((item, idx) =>
        idx === i ? { ...item, status: 'processing' } : item
      ));

      try {
        const response = await fetch('http://192.168.1.166:5001/api/import', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ selectedPOs: [po], importType, skipBlankMaterial }),
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
            tallyParsed: poResult.tallyParsed,
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

      // Calculate ETA
      const elapsedMs = Date.now() - startTime;
      const completedCount = i + 1;
      const remainingCount = posToImport.length - completedCount;
      if (completedCount > 0 && remainingCount > 0) {
        const avgTimePerItemMs = elapsedMs / completedCount;
        const estRemainingSec = Math.ceil((avgTimePerItemMs * remainingCount) / 1000);
        setEtaSeconds(estRemainingSec);
      } else {
        setEtaSeconds(0);
      }
    }

    const endTime = Date.now();
    setImportTimeTaken(((endTime - startTime) / 1000).toFixed(2));
    setEtaSeconds(null);
    setImporting(false);
    setShowSummaryModal(true);
  };

  const togglePoSelection = (poNumber) => {
    setSelectedPos(prev => {
      const next = new Set(prev);
      if (next.has(poNumber)) {
        next.delete(poNumber);
      } else {
        next.add(poNumber);
      }
      return next;
    });
  };

  const toggleDocTypeSelection = (docType) => {
    // Only target POs that have active items
    const posInDocType = parsedPOs.filter(po => {
      if (po.docType !== docType) return false;
      const activeCount = po.items ? po.items.filter(item => {
        const d = String(item['Deletion Indicator'] || item['deletion indicator'] || '').trim().toUpperCase();
        return d !== 'L';
      }).length : 0;
      return activeCount > 0;
    }).map(po => po.poNumber);

    const allSelected = posInDocType.every(num => selectedPos.has(num));

    setSelectedPos(prev => {
      const next = new Set(prev);
      if (allSelected) {
        posInDocType.forEach(num => next.delete(num));
      } else {
        posInDocType.forEach(num => next.add(num));
      }
      return next;
    });
  };

  const toggleAllSelection = () => {
    // Only target visible POs that have active items
    const visiblePoNumbers = filteredPOs.filter(po => {
      const activeCount = po.items ? po.items.filter(item => {
        const d = String(item['Deletion Indicator'] || item['deletion indicator'] || '').trim().toUpperCase();
        return d !== 'L';
      }).length : 0;
      return activeCount > 0;
    }).map(po => po.poNumber);

    const allVisibleSelected = visiblePoNumbers.every(num => selectedPos.has(num));

    setSelectedPos(prev => {
      const next = new Set(prev);
      if (allVisibleSelected) {
        visiblePoNumbers.forEach(num => next.delete(num));
      } else {
        visiblePoNumbers.forEach(num => next.add(num));
      }
      return next;
    });
  };

  const handleReset = () => {
    setFile(null);
    setParsedPOs(null);
    setSelectedPos(new Set());
    setImportResults(null);
    setError(null);
    setSearchTerm('');
    setResultSearchTerm('');
    setCurrentPage(1);
    setRowsPerPage(10);
    setImportTimeTaken(null);
    setExpandedPo({});
    setShowSummaryModal(false);
    setOnlyMissingMaterial(false);
  };

  const handleBackToReselect = () => {
    setImportResults(null);
    setCurrentPage(1);
    setSearchTerm('');
    setResultSearchTerm('');
    setExpandedPo({});
    setShowSummaryModal(false);
  };

  const togglePoExpand = (poNumber) => {
    setExpandedPo(prev => ({ ...prev, [poNumber]: !prev[poNumber] }));
  };

  const handleSelectRange = () => {
    if (!parsedPOs) return;
    
    const filteredByCategory = parsedPOs.filter(po => {
      if (rangeCategory === 'ALL') return true;
      return po.docType === rangeCategory;
    });

    const fromIndex = Math.max(1, parseInt(rangeFrom) || 1) - 1;
    const toIndex = Math.min(filteredByCategory.length, parseInt(rangeTo) || filteredByCategory.length) - 1;
    
    const newSelected = new Set(selectedPos);
    
    // First, clear current selection for items within the target category
    filteredByCategory.forEach(po => {
      newSelected.delete(po.poNumber);
    });

    // Add items within target range
    for (let i = fromIndex; i <= toIndex; i++) {
      if (filteredByCategory[i]) {
        const po = filteredByCategory[i];
        const activeCount = po.items ? po.items.filter(item => {
          const d = String(item['Deletion Indicator'] || item['deletion indicator'] || '').trim().toUpperCase();
          return d !== 'L';
        }).length : 0;
        if (activeCount > 0) {
          newSelected.add(po.poNumber);
        }
      }
    }
    setSelectedPos(newSelected);
  };

  const copyToClipboard = async (text, poNumber) => {
    try {
      await navigator.clipboard.writeText(text);
      setCopySuccess(poNumber);
      setTimeout(() => setCopySuccess(null), 2000);
    } catch {
      setCopySuccess(null);
    }
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

  const docTypeCounts = useMemo(() => {
    if (!parsedPOs) return {};
    return parsedPOs.reduce((acc, po) => {
      acc[po.docType] = (acc[po.docType] || 0) + 1;
      return acc;
    }, {});
  }, [parsedPOs]);

  const hasMissingMaterialItem = (po) => {
    if (!po || !po.items) return false;
    return po.items.some(item => {
      const d = String(item['Deletion Indicator'] || item['deletion indicator'] || '').trim().toUpperCase();
      if (d === 'L') return false;
      const getV = (row, key) => {
        if (!row) return '';
        const clean = key.toLowerCase().replace(/\s/g, '');
        for (const k of Object.keys(row)) {
          if (k.toLowerCase().replace(/\s/g, '') === clean) return row[k];
        }
        return '';
      };
      const mat = String(getV(item, 'Material') || '').split('.')[0].trim();
      return !mat || mat === '0' || /^0+$/.test(mat);
    });
  };

  const missingMaterialRecords = useMemo(() => {
    if (!parsedPOs) return [];
    return parsedPOs.filter(po => hasMissingMaterialItem(po));
  }, [parsedPOs]);

  const handleSelectOnlyMissingMaterial = () => {
    setOnlyMissingMaterial(true);
    const targetNumbers = missingMaterialRecords.map(po => po.poNumber);
    setSelectedPos(new Set(targetNumbers));
    setCurrentPage(1);
  };

  const filteredPOs = useMemo(() => {
    if (!parsedPOs) return [];
    let list = parsedPOs;
    if (onlyMissingMaterial) {
      list = list.filter(po => hasMissingMaterialItem(po));
    }
    const term = searchTerm.toLowerCase();
    if (!term) return list;
    return list.filter(po => 
      po.poNumber.toLowerCase().includes(term) ||
      (po.vendorName || '').toLowerCase().includes(term)
    );
  }, [parsedPOs, searchTerm, onlyMissingMaterial]);

  const uniqueDocTypes = useMemo(() => Object.keys(docTypeCounts), [docTypeCounts]);

  const currentStep = !parsedPOs && !importResults ? 1 : (parsedPOs && !importResults ? 2 : 3);

  const filteredResults = useMemo(() => {
    if (!importResults) return [];
    const term = resultSearchTerm.toLowerCase();
    if (!term) return importResults;
    return importResults.filter(po => 
      po.poNumber.toLowerCase().includes(term) ||
      (po.vendorName || '').toLowerCase().includes(term)
    );
  }, [importResults, resultSearchTerm]);

  const totalItemsCount = currentStep === 2 ? filteredPOs.length : filteredResults.length;
  const indexOfLastResult = currentPage * rowsPerPage;
  const indexOfFirstResult = indexOfLastResult - rowsPerPage;
  
  const currentPOsPage = filteredPOs.slice(indexOfFirstResult, indexOfLastResult);
  const currentResultsPage = filteredResults.slice(indexOfFirstResult, indexOfLastResult);
  const totalPages = Math.ceil(totalItemsCount / rowsPerPage);

  // Derived stats
  const successCount = importResults ? importResults.filter(r => r.status === 'success').length : 0;
  const failedCount = importResults ? importResults.filter(r => r.status === 'failed').length : 0;
  const processingCount = importResults ? importResults.filter(r => r.status === 'processing').length : 0;
  const pendingCount = importResults ? importResults.filter(r => r.status === 'pending').length : 0;
  const doneCount = successCount + failedCount;
  const totalCount = importResults ? importResults.length : 0;
  const progressPct = totalCount > 0 ? Math.round((doneCount / totalCount) * 100) : 0;
  const filteredTotalCount = filteredResults.length;

  const totalCreated = importResults ? importResults.reduce((acc, r) => acc + (r.tallyParsed?.created || 0), 0) : 0;
  const totalAltered = importResults ? importResults.reduce((acc, r) => acc + (r.tallyParsed?.altered || 0), 0) : 0;
  const totalExceptions = importResults ? importResults.reduce((acc, r) => acc + (r.tallyParsed?.exceptions || 0), 0) : 0;
  const totalErrors = importResults ? importResults.reduce((acc, r) => acc + (r.tallyParsed?.errors || 0), 0) : 0;

  const renderItemsTable = (items) => {
    const getV = (row, key) => {
      if (!row) return '';
      const clean = key.toLowerCase().replace(/\s/g, '');
      for (const k of Object.keys(row)) {
        if (k.toLowerCase().replace(/\s/g, '') === clean) return row[k];
      }
      return '';
    };

    return (
      <div className="items-table-wrap">
        <table className="items-table">
          <thead>
            <tr>
              <th>Material</th>
              <th>Description</th>
              <th>Quantity</th>
              <th>Unit Price</th>
              <th>Value</th>
              <th>Status</th>
            </tr>
          </thead>
          <tbody>
            {items.map((item, idx) => {
              const material = String(getV(item, 'Material') || '').split('.')[0].trim();
              const shortText = String(getV(item, 'Material Description') || getV(item, 'Purchase Order - Short Text') || getV(item, 'Short Text') || getV(item, 'Text') || '').trim();
              const qty = importType === 'po' ? getV(item, 'Order Quantity') :
                          (importType === 'purchase' ? getV(item, 'Quantity') || getV(item, 'Qty in OPUn') : getV(item, 'Qty in Un. of Entry'));
              const unit = importType === 'po' ? String(getV(item, 'Order Unit') || '').trim() :
                           (importType === 'purchase' ? String(getV(item, 'Order Unit') || getV(item, 'Order Price Unit') || '').trim() :
                            String(getV(item, 'Unit of Entry') || '').trim());
              let price = importType === 'po' ? getV(item, 'Net Order Price') : '';
              let val = importType === 'po' ? getV(item, 'Net Order Value') :
                        (importType === 'purchase' ? getV(item, 'Amount') : getV(item, 'Amount in LC'));
              if (importType === 'grn') {
                const parsedQty = parseFloat(qty) || 0;
                const parsedVal = parseFloat(val) || 0;
                const netVal = parsedVal / 1.18;
                val = netVal;
                price = parsedQty > 0 ? (netVal / parsedQty) : 0;
              } else if (importType === 'purchase') {
                const parsedQty = parseFloat(qty) || 0;
                const parsedVal = parseFloat(val) || 0;
                price = parsedQty > 0 ? (parsedVal / parsedQty) : 0;
              }
              const delInd = String(getV(item, 'Deletion Indicator') || '').trim().toUpperCase();
              const isIgnored = delInd === 'L';

              return (
                <tr key={idx} className={isIgnored ? 'row-ignored' : 'row-active'}>
                  <td><code className="mat-code">{material || '-'}</code></td>
                  <td className="desc-cell">{shortText || '-'}</td>
                  <td>{qty ? `${qty} ${unit || 'Nos'}` : '-'}</td>
                  <td className="num-cell">{price ? Number(price).toLocaleString('en-IN', { minimumFractionDigits: 2 }) : '-'}</td>
                  <td className="num-cell">{val ? Number(val).toLocaleString('en-IN', { minimumFractionDigits: 2 }) : '-'}</td>
                  <td>
                    {isIgnored
                      ? <span className="badge-ignored">Deleted</span>
                      : <span className="badge-active">Active</span>
                    }
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    );
  };

  return (
    <div className="app-shell">
      {/* Sidebar */}
      <aside className={`sidebar ${isSidebarCollapsed ? 'collapsed' : ''}`}>
        <div className="sidebar-brand" style={{ justifyContent: isSidebarCollapsed ? 'center' : 'flex-start' }}>
          <div className="brand-icon">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="22" height="22">
              <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
              <polyline points="14 2 14 8 20 8" />
              <line x1="16" y1="13" x2="8" y2="13" />
              <line x1="16" y1="17" x2="8" y2="17" />
            </svg>
          </div>
          {!isSidebarCollapsed && (
            <div className="brand-text">
              <span className="brand-name">OPG Gateway</span>
              <span className="brand-sub">Tally Importer</span>
            </div>
          )}
          <button 
            className="sidebar-toggle-btn" 
            onClick={() => setIsSidebarCollapsed(!isSidebarCollapsed)}
            title={isSidebarCollapsed ? "Expand Sidebar" : "Collapse Sidebar"}
            style={{
              background: 'transparent',
              border: 'none',
              color: 'var(--text-2)',
              marginLeft: isSidebarCollapsed ? '0' : 'auto',
              cursor: 'pointer',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              padding: '4px',
              borderRadius: '4px'
            }}
          >
            {isSidebarCollapsed ? (
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="16" height="16">
                <polyline points="9 18 15 12 9 6" />
              </svg>
            ) : (
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="16" height="16">
                <polyline points="15 18 9 12 15 6" />
              </svg>
            )}
          </button>
        </div>

        <div className="sidebar-mode-switcher">
          <button 
            className={`mode-btn ${activeMode === 'importer' ? 'active' : ''}`}
            onClick={() => setActiveMode('importer')}
            title="Excel Importer"
          >
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="14" height="14">
              <rect x="3" y="3" width="18" height="18" rx="2" ry="2"/>
              <line x1="9" y1="3" x2="9" y2="21"/>
            </svg>
            {!isSidebarCollapsed && <span>Excel Importer</span>}
          </button>
          <button 
            className={`mode-btn ${activeMode === 'visualizer' ? 'active' : ''}`}
            onClick={() => setActiveMode('visualizer')}
            title="XML Visualizer"
          >
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="14" height="14">
              <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
              <polyline points="14 2 14 8 20 8"/>
              <line x1="16" y1="13" x2="8" y2="13"/>
              <line x1="16" y1="17" x2="8" y2="17"/>
            </svg>
            {!isSidebarCollapsed && <span>DayBook Report</span>}
          </button>
          {/* <button 
            className={`mode-btn ${activeMode === 'daybook' ? 'active' : ''}`}
            onClick={() => setActiveMode('daybook')}
            title="DayBook Register"
          >
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="14" height="14">
              <rect x="3" y="4" width="18" height="18" rx="2" ry="2" />
              <line x1="16" y1="2" x2="16" y2="6" />
              <line x1="8" y1="2" x2="8" y2="6" />
              <line x1="3" y1="10" x2="21" y2="10" />
            </svg>
            {!isSidebarCollapsed && <span>DayBook Register</span>}
          </button> */}
        </div>

        {activeMode === 'importer' ? (
          <>
            <nav className="sidebar-nav">
              <div className={`nav-step ${currentStep >= 1 ? 'completed' : ''} ${currentStep === 1 ? 'active' : ''}`}>
                <div className="step-icon-wrap">
                  {currentStep > 1
                    ? <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" width="14" height="14"><polyline points="20 6 9 17 4 12" /></svg>
                    : <span>1</span>
                  }
                </div>
                <div className="step-label">
                  <span className="step-title">Upload</span>
                  <span className="step-desc">Select Excel file</span>
                </div>
              </div>

              <div className="step-connector" />

              <div className={`nav-step ${currentStep >= 2 ? 'completed' : ''} ${currentStep === 2 ? 'active' : ''}`}>
                <div className="step-icon-wrap">
                  {currentStep > 2
                    ? <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" width="14" height="14"><polyline points="20 6 9 17 4 12" /></svg>
                    : <span>2</span>
                  }
                </div>
                <div className="step-label">
                  <span className="step-title">Select POs</span>
                  <span className="step-desc">Choose orders</span>
                </div>
              </div>

              <div className="step-connector" />

              <div className={`nav-step ${currentStep === 3 ? 'active' : ''}`}>
                <div className="step-icon-wrap"><span>3</span></div>
                <div className="step-label">
                  <span className="step-title">Import</span>
                  <span className="step-desc">Push to Tally</span>
                </div>
              </div>
            </nav>

            {importResults && (
              <div className="sidebar-stats">
                <div className="stat-item stat-success">
                  <span className="stat-val">{successCount}</span>
                  <span className="stat-lbl">Imported</span>
                </div>
                <div className="stat-item stat-failed">
                  <span className="stat-val">{failedCount}</span>
                  <span className="stat-lbl">Failed</span>
                </div>
                <div className="stat-item stat-pending">
                  <span className="stat-val">{pendingCount + processingCount}</span>
                  <span className="stat-lbl">Pending</span>
                </div>
              </div>
            )}

            <div className="sidebar-footer">
              <button className="btn-reset" onClick={handleReset} title="Start over">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="14" height="14">
                  <polyline points="1 4 1 10 7 10" />
                  <path d="M3.51 15a9 9 0 1 0 .49-3.58" />
                </svg>
                Start Over
              </button>
            </div>
          </>
        ) : (
          <div className="sidebar-info-box">
            <p>Visualizing custom Tally XML files directly from database and file stores.</p>
          </div>
        )}
      </aside>

      {/* Main Content */}
      <main className="main-content">
        {error && (
          <div className="error-banner">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="16" height="16">
              <circle cx="12" cy="12" r="10" />
              <line x1="12" y1="8" x2="12" y2="12" />
              <line x1="12" y1="16" x2="12.01" y2="16" />
            </svg>
            <span>{error}</span>
            <button onClick={() => setError(null)} className="error-close">&times;</button>
          </div>
        )}

        {activeMode === 'importer' ? (
          <>
            {/* ── STEP 1: UPLOAD ── */}
            {currentStep === 1 && (
          <div className="page-section">
            <div className="import-type-selector" style={{ display: 'flex', gap: '8px', marginBottom: '24px', background: 'var(--border)', padding: '4px', borderRadius: '8px', width: 'fit-content' }}>
              <button 
                className={`mode-btn ${importType === 'po' ? 'active' : ''}`}
                onClick={() => { setImportType('po'); setFile(null); setError(null); }}
                style={{ padding: '8px 16px', borderRadius: '6px', border: 'none', background: importType === 'po' ? 'var(--accent)' : 'transparent', color: importType === 'po' ? '#fff' : 'var(--text-2)', fontWeight: '600', cursor: 'pointer', transition: 'all 0.2s' }}
              >
                Purchase Order (PO)
              </button>
              <button 
                className={`mode-btn ${importType === 'grn' ? 'active' : ''}`}
                onClick={() => { setImportType('grn'); setFile(null); setError(null); }}
                style={{ padding: '8px 16px', borderRadius: '6px', border: 'none', background: importType === 'grn' ? 'var(--accent)' : 'transparent', color: importType === 'grn' ? '#fff' : 'var(--text-2)', fontWeight: '600', cursor: 'pointer', transition: 'all 0.2s' }}
              >
                Goods Receipt Note (GRN)
              </button>
              <button 
                className={`mode-btn ${importType === 'purchase' ? 'active' : ''}`}
                onClick={() => { setImportType('purchase'); setFile(null); setError(null); }}
                style={{ padding: '8px 16px', borderRadius: '6px', border: 'none', background: importType === 'purchase' ? 'var(--accent)' : 'transparent', color: importType === 'purchase' ? '#fff' : 'var(--text-2)', fontWeight: '600', cursor: 'pointer', transition: 'all 0.2s' }}
              >
                Purchase Invoice
              </button>
              <button 
                className={`mode-btn ${importType === 'stock_journal' ? 'active' : ''}`}
                onClick={() => { setImportType('stock_journal'); setFile(null); setError(null); }}
                style={{ padding: '8px 16px', borderRadius: '6px', border: 'none', background: importType === 'stock_journal' ? 'var(--accent)' : 'transparent', color: importType === 'stock_journal' ? '#fff' : 'var(--text-2)', fontWeight: '600', cursor: 'pointer', transition: 'all 0.2s' }}
              >
                Stock Journal
              </button>
            </div>
            <div className="page-header">
              <h1 className="page-title">
                {importType === 'po' ? 'Upload PO Spreadsheet' : 
                 (importType === 'purchase' ? 'Upload Purchase Invoice Spreadsheet' : 
                  (importType === 'stock_journal' ? 'Upload Stock Journal Spreadsheet' : 'Upload GRN Spreadsheet'))}
              </h1>
              <p className="page-desc">
                {importType === 'po' ? 'Drop your SAP Excel export to start importing purchase orders into Tally.' : 
                 (importType === 'purchase' ? 'Drop your Purchase Invoice Excel export to start importing Purchase Invoices into Tally.' : 
                  (importType === 'stock_journal' ? 'Drop your SAP Excel export to start importing WA Stock Journal entries into Tally.' : 'Drop your SAP GRN Excel export to start importing Receipt Notes into Tally.'))}
              </p>
            </div>

            <div className="upload-area">
              <div
                className={`dropzone ${isDragging ? 'dragging' : ''} ${file ? 'has-file' : ''}`}
                onClick={handleDropzoneClick}
                onDragOver={handleDragOver}
                onDragLeave={handleDragLeave}
                onDrop={handleDrop}
              >
                <input
                  type="file"
                  id="fileInput"
                  ref={fileInputRef}
                  accept=".xlsx,.xls"
                  onChange={handleFileChange}
                  className="hidden-input"
                />

                {file ? (
                  <div className="file-ready">
                    <div className="file-icon-wrap success">
                      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="28" height="28">
                        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
                        <polyline points="14 2 14 8 20 8" />
                        <polyline points="9 15 11 17 15 13" />
                      </svg>
                    </div>
                    <div className="file-info-block">
                      <span className="file-name-text">{file.name}</span>
                      <span className="file-size-text">{(file.size / 1024).toFixed(1)} KB · Excel Workbook</span>
                    </div>
                    <button
                      className="file-remove-btn"
                      onClick={(e) => { e.stopPropagation(); setFile(null); }}
                      title="Remove file"
                    >
                      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="16" height="16">
                        <line x1="18" y1="6" x2="6" y2="18" />
                        <line x1="6" y1="6" x2="18" y2="18" />
                      </svg>
                    </button>
                  </div>
                ) : (
                  <div className="dropzone-prompt">
                    <div className={`upload-icon-wrap ${isDragging ? 'bounce' : ''}`}>
                      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" width="36" height="36">
                        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
                        <polyline points="17 8 12 3 7 8" />
                        <line x1="12" y1="3" x2="12" y2="15" />
                      </svg>
                    </div>
                    <p className="drop-title">{isDragging ? 'Release to upload' : 'Drag & drop your Excel file'}</p>
                    <p className="drop-sub">or <span className="drop-link">browse to choose a file</span></p>
                    <p className="drop-hint">.xlsx or .xls · {importType === 'po' ? 'SAP PO export format' : (importType === 'purchase' ? 'Purchase Invoice format' : 'SAP GRN export format')}</p>
                  </div>
                )}
              </div>

              <button
                className="btn-primary btn-lg"
                onClick={handleUpload}
                disabled={!file || uploading}
              >
                {uploading ? (
                  <><span className="btn-spinner" /> {parsingProgress ? parsingProgress.message : 'Parsing Spreadsheet…'}</>
                ) : (
                  <>
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="16" height="16">
                      <polyline points="9 11 12 14 22 4" />
                      <path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11" />
                    </svg>
                    Parse &amp; Continue
                  </>
                )}
              </button>

              {parsingProgress && (
                <div className="parsing-progress-container" style={{ marginTop: '20px', width: '100%', maxWidth: '500px', marginLeft: 'auto', marginRight: 'auto' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '8px', fontSize: '13px', color: 'var(--text-2)', fontWeight: '500' }}>
                    <span>{parsingProgress.message}</span>
                    {parsingProgress.totalRows > 0 && (
                      <span>{Math.round((parsingProgress.rowsParsed / parsingProgress.totalRows) * 100)}%</span>
                    )}
                  </div>
                  {parsingProgress.totalRows > 0 && (
                    <div className="progress-bar-wrap" style={{ height: '6px', background: 'var(--surface-3)', borderRadius: '3px', overflow: 'hidden', border: '1px solid var(--border)' }}>
                      <div className="progress-bar" style={{ height: '100%', background: 'var(--primary)', width: `${(parsingProgress.rowsParsed / parsingProgress.totalRows) * 100}%`, transition: 'width 0.1s ease-out' }} />
                    </div>
                  )}
                </div>
              )}
            </div>
          </div>
        )}

        {/* ── STEP 2: SELECT ── */}
        {currentStep === 2 && (
          <div className="page-section">
            <div className="page-header">
              <div>
                <h1 className="page-title">{importType === 'po' ? 'Select Purchase Orders' : (importType === 'purchase' ? 'Select Purchase Invoices' : (importType === 'stock_journal' ? 'Select Stock Journal (WA) Entries' : 'Select Goods Receipt Notes (GRN)'))}</h1>
                <p className="page-desc">
                  {parsedPOs.length} {importType === 'po' ? 'POs' : (importType === 'purchase' ? 'Purchase Invoices' : (importType === 'stock_journal' ? 'Stock Journals' : 'GRNs'))} found · {selectedPos.size} selected
                  {importType === 'purchase' && ignoredInvoices.length > 0 && ` · ${ignoredInvoices.length} ignored (Missing GRN Tracking)`}
                </p>
              </div>
              <div style={{ display: 'flex', gap: '10px' }}>
                {importType === 'purchase' && ignoredInvoices.length > 0 && (
                  <button 
                    className="btn-secondary" 
                    onClick={() => setShowIgnoredModal(true)}
                    style={{ background: '#fef2f2', border: '1px solid #fca5a5', color: '#dc2626', fontWeight: '600' }}
                  >
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="14" height="14">
                      <circle cx="12" cy="12" r="10" />
                      <line x1="12" y1="8" x2="12" y2="12" />
                      <line x1="12" y1="16" x2="12.01" y2="16" />
                    </svg>
                    View Ignored Invoices ({ignoredInvoices.length})
                  </button>
                )}
                <button className="btn-secondary" onClick={handleReset}>
                  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="14" height="14">
                    <line x1="19" y1="12" x2="5" y2="12" />
                    <polyline points="12 19 5 12 12 5" />
                  </svg>
                  Back
                </button>
              </div>
            </div>

            {/* Filter bar */}
            <div className="filter-bar">
              <div className="search-field">
                <svg className="search-ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <circle cx="11" cy="11" r="8" />
                  <line x1="21" y1="21" x2="16.65" y2="16.65" />
                </svg>
                <input
                  className="search-inp"
                  type="text"
                  placeholder={importType === 'po' ? 'Search by PO number or vendor…' : (importType === 'purchase' ? 'Search by Invoice number or supplier…' : 'Search by GRN number or vendor…')}
                  value={searchTerm}
                  onChange={(e) => {
                    setSearchTerm(e.target.value);
                    setCurrentPage(1);
                  }}
                />
                {searchTerm && (
                  <button className="search-clear" onClick={() => { setSearchTerm(''); setCurrentPage(1); }}>&times;</button>
                )}
              </div>

              <div className="dt-pills">
                <label
                  className={`dt-pill select-all-pill ${filteredPOs.length > 0 && filteredPOs.every(po => selectedPos.has(po.poNumber)) ? 'active' : ''}`}
                >
                  <input
                    type="checkbox"
                    checked={filteredPOs.length > 0 && filteredPOs.every(po => selectedPos.has(po.poNumber))}
                    onChange={toggleAllSelection}
                  />
                  All
                </label>
                {uniqueDocTypes.map(dt => {
                  const posInDoc = parsedPOs.filter(po => po.docType === dt).map(po => po.poNumber);
                  const allSel = posInDoc.every(n => selectedPos.has(n));
                  const someSel = posInDoc.some(n => selectedPos.has(n)) && !allSel;
                  return (
                    <label
                      key={dt}
                      className={`dt-pill ${allSel ? 'active' : ''} ${someSel ? 'partial' : ''}`}
                    >
                      <input
                        type="checkbox"
                        checked={allSel}
                        ref={el => { if (el) el.indeterminate = someSel; }}
                        onChange={() => toggleDocTypeSelection(dt)}
                      />
                      {dt}
                      <span className="dt-count">{docTypeCounts[dt]}</span>
                    </label>
                  );
                })}
              </div>
            </div>

            {/* Range Selector Bar */}
            <div className="range-selector-bar" style={{ display: 'flex', alignItems: 'center', gap: '12px', background: 'var(--surface-2)', padding: '12px 16px', borderRadius: '8px', marginBottom: '16px', border: '1px solid var(--border)', flexWrap: 'wrap' }}>
              <span style={{ fontSize: '14px', fontWeight: '600', color: 'var(--text-1)' }}>Select Range:</span>
              
              {importType === 'grn' && (
                <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                  <span style={{ fontSize: '13px', color: 'var(--text-2)' }}>Category</span>
                  <select 
                    value={rangeCategory} 
                    onChange={(e) => {
                      setRangeCategory(e.target.value);
                      const count = parsedPOs.filter(po => e.target.value === 'ALL' ? true : po.docType === e.target.value).length;
                      setRangeFrom(1);
                      setRangeTo(Math.min(100, count));
                    }} 
                    style={{ padding: '6px 8px', borderRadius: '4px', border: '1px solid var(--border)', background: 'var(--surface)', color: 'var(--text-1)' }}
                  >
                    <option value="ALL">All</option>
                    <option value="WE">WE (GRN)</option>
                    <option value="WA">WA (Internal Transfer)</option>
                  </select>
                </div>
              )}

              <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                <span style={{ fontSize: '13px', color: 'var(--text-2)' }}>From</span>
                <input 
                  type="number" 
                  value={rangeFrom} 
                  onChange={(e) => setRangeFrom(e.target.value)} 
                  style={{ width: '70px', padding: '6px 8px', borderRadius: '4px', border: '1px solid var(--border)', background: 'var(--surface)', color: 'var(--text-1)' }}
                  min="1"
                />
                <span style={{ fontSize: '13px', color: 'var(--text-2)' }}>To</span>
                <input 
                  type="number" 
                  value={rangeTo} 
                  onChange={(e) => setRangeTo(e.target.value)} 
                  style={{ width: '70px', padding: '6px 8px', borderRadius: '4px', border: '1px solid var(--border)', background: 'var(--surface)', color: 'var(--text-1)' }}
                  min="1"
                />
              </div>
              <button 
                className="btn-primary" 
                onClick={handleSelectRange}
                style={{ padding: '6px 12px', fontSize: '13px', height: 'auto' }}
              >
                Apply Range
              </button>
              <span style={{ fontSize: '12px', color: 'var(--text-3)' }}>
                (Total Category: {parsedPOs ? parsedPOs.filter(po => rangeCategory === 'ALL' ? true : po.docType === rangeCategory).length : 0})
              </span>

              {missingMaterialRecords.length > 0 && (
                <button
                  onClick={handleSelectOnlyMissingMaterial}
                  style={{
                    background: '#eab308',
                    color: '#000000',
                    fontWeight: '700',
                    padding: '6px 14px',
                    borderRadius: '6px',
                    fontSize: '13px',
                    border: 'none',
                    cursor: 'pointer',
                    display: 'flex',
                    alignItems: 'center',
                    gap: '6px',
                    boxShadow: '0 2px 4px rgba(0,0,0,0.1)'
                  }}
                  title="Select all GRNs/Vouchers that have missing Material codes"
                >
                  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" width="14" height="14">
                    <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14" />
                    <polyline points="22 4 12 14.01 9 11.01" />
                  </svg>
                  Select Only Missing Material ({missingMaterialRecords.length})
                </button>
              )}

              <label style={{ display: 'flex', alignItems: 'center', gap: '8px', fontSize: '13px', color: 'var(--text-1)', cursor: 'pointer', background: onlyMissingMaterial ? 'rgba(234, 179, 8, 0.15)' : 'var(--surface-3)', padding: '6px 12px', borderRadius: '6px', border: `1px solid ${onlyMissingMaterial ? '#eab308' : 'var(--border)'}`, transition: 'all 0.2s', marginLeft: missingMaterialRecords.length > 0 ? '0' : 'auto' }}>
                <input 
                  type="checkbox" 
                  checked={onlyMissingMaterial} 
                  onChange={(e) => {
                    setOnlyMissingMaterial(e.target.checked);
                    setCurrentPage(1);
                  }} 
                  style={{ width: '15px', height: '15px', cursor: 'pointer' }}
                />
                Filter only records with missing Material
              </label>

              <label style={{ display: 'flex', alignItems: 'center', gap: '8px', fontSize: '13px', color: 'var(--text-1)', cursor: 'pointer', background: 'var(--surface-3)', padding: '6px 12px', borderRadius: '6px', border: '1px solid var(--border)' }}>
                <input 
                  type="checkbox" 
                  checked={skipBlankMaterial} 
                  onChange={(e) => setSkipBlankMaterial(e.target.checked)} 
                  style={{ width: '15px', height: '15px', cursor: 'pointer' }}
                />
                Skip items with blank Material
              </label>
            </div>

            {/* PO List */}
            <div className="po-list">
              {filteredPOs.length === 0 && (
                <div className="empty-state">
                  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" width="40" height="40" style={{ opacity: 0.3 }}>
                    <circle cx="11" cy="11" r="8" />
                    <line x1="21" y1="21" x2="16.65" y2="16.65" />
                  </svg>
                  <p>No POs match your search.</p>
                </div>
              )}
              {currentPOsPage.map(po => {
                const isSelected = selectedPos.has(po.poNumber);
                const isExpanded = !!expandedPo[po.poNumber];
                const activeCount = po.items.filter(item => {
                  const d = String(item['Deletion Indicator'] || item['deletion indicator'] || '').trim().toUpperCase();
                  return d !== 'L';
                }).length;
                const ignoredCount = po.items.length - activeCount;
                const isFullyIgnored = activeCount === 0 && ignoredCount > 0;

                return (
                  <div key={po.poNumber} className={`po-row ${isSelected ? 'selected' : ''} ${isFullyIgnored ? 'po-row-ignored' : ''}`}>
                    <div className="po-row-main">
                      <label className={`po-check ${isFullyIgnored ? 'disabled' : ''}`} onClick={(e) => e.stopPropagation()}>
                        <input
                          type="checkbox"
                          disabled={isFullyIgnored}
                          checked={isSelected}
                          onChange={() => !isFullyIgnored && togglePoSelection(po.poNumber)}
                        />
                        <span className="checkmark" />
                      </label>

                      <div className="po-info" onClick={() => togglePoExpand(po.poNumber)}>
                        <div className="po-info-top">
                          <span className="po-num">#{po.poNumber}</span>
                          <span className="dt-tag">{po.docType}</span>
                          {hasMissingMaterialItem(po) && (
                            <span style={{ color: '#eab308', background: 'rgba(234, 179, 8, 0.15)', border: '1px solid rgba(234, 179, 8, 0.3)', padding: '1px 6px', borderRadius: '4px', fontSize: '11px', fontWeight: '600' }} title="Has item(s) without Material Code">
                              Missing Material
                            </span>
                          )}
                        </div>
                        <span className="po-vendor">{po.vendorName}</span>
                      </div>

                      <div className="po-meta" onClick={() => togglePoExpand(po.poNumber)}>
                        <span className="po-items-ct">
                          {activeCount} item{activeCount !== 1 ? 's' : ''}
                          {ignoredCount > 0 && <span className="ignored-ct"> · {ignoredCount} deleted</span>}
                        </span>
                        <svg
                          className={`chevron-icon ${isExpanded ? 'open' : ''}`}
                          viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"
                          width="14" height="14"
                        >
                          <polyline points="6 9 12 15 18 9" />
                        </svg>
                      </div>
                    </div>

                    {isExpanded && (
                      <div className="po-expand">
                        {renderItemsTable(po.items)}
                      </div>
                    )}
                  </div>
                );
              })}
            </div>

            {/* Pagination for Selection List */}
            {filteredPOs.length > 0 && (
              <div className="pagination" style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 'var(--radius)', padding: '0.85rem 1rem', marginBottom: '0.5rem', boxShadow: 'var(--shadow-sm)' }}>
                <span className="pg-info">
                  Showing {indexOfFirstResult + 1}–{Math.min(indexOfLastResult, filteredPOs.length)} of {filteredPOs.length}
                </span>
                <div className="pg-controls">
                  <div className="pg-rpp">
                    <span>Rows</span>
                    <select
                      value={rowsPerPage}
                      onChange={(e) => { setRowsPerPage(Number(e.target.value)); setCurrentPage(1); }}
                      className="rpp-select"
                    >
                      {[5, 10, 25, 50, 100].map(s => <option key={s} value={s}>{s}</option>)}
                    </select>
                  </div>
                  <div className="pg-btns">
                    <button className="pg-btn" onClick={() => setCurrentPage(1)} disabled={currentPage === 1}>&laquo;</button>
                    <button className="pg-btn" onClick={() => setCurrentPage(p => Math.max(p - 1, 1))} disabled={currentPage === 1}>&lsaquo;</button>
                    <span className="pg-indicator">{currentPage} / {totalPages}</span>
                    <button className="pg-btn" onClick={() => setCurrentPage(p => Math.min(p + 1, totalPages))} disabled={currentPage === totalPages}>&rsaquo;</button>
                    <button className="pg-btn" onClick={() => setCurrentPage(totalPages)} disabled={currentPage === totalPages}>&raquo;</button>
                  </div>
                </div>
              </div>
            )}

            <div className="import-footer">
              <div className="import-footer-info">
                <span className="sel-count-label">{selectedPos.size} of {parsedPOs.length} {importType === 'po' ? 'POs' : (importType === 'purchase' ? 'Purchase Invoices' : 'GRNs')} selected</span>
              </div>
              <button
                className="btn-primary btn-lg"
                onClick={handleImport}
                disabled={importing || selectedPos.size === 0}
              >
                {importing ? (
                  <><span className="btn-spinner" /> Importing…</>
                ) : (
                  <>
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="16" height="16">
                      <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14" />
                      <polyline points="22 4 12 14.01 9 11.01" />
                    </svg>
                    Import {selectedPos.size} {importType === 'po' ? 'PO' : (importType === 'purchase' ? 'Purchase Invoice' : 'GRN')}{selectedPos.size !== 1 ? 's' : ''} to Tally
                  </>
                )}
              </button>
            </div>
          </div>
        )}

        {/* ── STEP 3: RESULTS ── */}
        {currentStep === 3 && (
          <div className="page-section">
            <div className="page-header">
              <div>
                <h1 className="page-title">
                  {importing ? 'Importing to Tally…' : 'Import Complete'}
                </h1>
                <p className="page-desc">
                  {importing
                    ? `${doneCount} of ${totalCount} completed`
                    : `${successCount} imported · ${failedCount} failed · ${importTimeTaken ? `${importTimeTaken}s` : ''}`}
                </p>
              </div>
              {!importing && (
                <div style={{ display: 'flex', gap: '8px' }}>
                  <button className="btn-secondary" onClick={handleBackToReselect}>
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="14" height="14">
                      <line x1="19" y1="12" x2="5" y2="12" />
                      <polyline points="12 19 5 12 12 5" />
                    </svg>
                    Reselect / Adjust Range
                  </button>
                  <button className="btn-secondary" onClick={handleReset}>
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="14" height="14">
                      <polyline points="1 4 1 10 7 10" />
                      <path d="M3.51 15a9 9 0 1 0 .49-3.58" />
                    </svg>
                    New Import
                  </button>
                </div>
              )}
            </div>

            {/* Progress bar */}
            {importing && (
              <div className="progress-bar-wrap" style={{ flexDirection: 'column', alignItems: 'stretch', gap: '8px' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', fontSize: '13px', fontWeight: '600' }}>
                  <span style={{ color: 'var(--text-1)' }}>
                    Importing entries into Tally ({doneCount} of {totalCount})
                  </span>
                  <div style={{ display: 'flex', gap: '16px', alignItems: 'center' }}>
                    {etaSeconds !== null && etaSeconds > 0 && (
                      <span style={{ color: '#2563eb', fontWeight: '600' }}>
                        Estimated time remaining: {etaSeconds > 60 ? `${Math.floor(etaSeconds / 60)}m ${etaSeconds % 60}s` : `${etaSeconds}s`}
                      </span>
                    )}
                    <span className="progress-pct" style={{ fontSize: '14px' }}>{progressPct}%</span>
                  </div>
                </div>
                <div className="progress-track" style={{ width: '100%', height: '10px', background: 'var(--surface-3, #e2e8f0)', borderRadius: '5px', overflow: 'hidden', position: 'relative' }}>
                  <div 
                    className="progress-bar" 
                    style={{ 
                      width: `${progressPct}%`, 
                      height: '100%', 
                      background: 'linear-gradient(90deg, #2563eb, #3b82f6)', 
                      borderRadius: '5px',
                      transition: 'width 0.3s ease-in-out'
                    }} 
                  />
                </div>
              </div>
            )}

            {/* Summary cards */}
            <div className="summary-cards">
              <div className="sum-card sum-success">
                <span className="sum-num">{successCount}</span>
                <span className="sum-lbl">Imported</span>
              </div>
              <div className="sum-card sum-failed">
                <span className="sum-num">{failedCount}</span>
                <span className="sum-lbl">Failed</span>
              </div>
              <div className="sum-card sum-neutral">
                <span className="sum-num">{pendingCount + processingCount}</span>
                <span className="sum-lbl">Pending</span>
              </div>
              <div className="sum-card sum-total">
                <span className="sum-num">{totalCount}</span>
                <span className="sum-lbl">Total</span>
              </div>
            </div>

            {/* Search Box */}
            <div className="filter-bar" style={{ marginTop: '0.5rem', marginBottom: '0.5rem' }}>
              <div className="search-field">
                <svg className="search-ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <circle cx="11" cy="11" r="8" />
                  <line x1="21" y1="21" x2="16.65" y2="16.65" />
                </svg>
                <input
                  className="search-inp"
                  type="text"
                  placeholder={importType === 'po' ? 'Search by PO number or vendor in results…' : (importType === 'purchase' ? 'Search by Invoice number or supplier in results…' : (importType === 'stock_journal' ? 'Search by Material Doc number in results…' : 'Search by GRN number or vendor in results…'))}
                  value={resultSearchTerm}
                  onChange={(e) => {
                    setResultSearchTerm(e.target.value);
                    setCurrentPage(1);
                  }}
                />
                {resultSearchTerm && (
                  <button className="search-clear" onClick={() => setResultSearchTerm('')}>&times;</button>
                )}
              </div>
            </div>

            {/* Results table */}
            <div className="results-card">
              <div className="table-scroll">
                <table className="results-table">
                  <thead>
                    <tr>
                      <th>{importType === 'po' ? 'PO Number' : (importType === 'purchase' ? 'Invoice Number' : (importType === 'stock_journal' ? 'Material Doc' : 'GRN Number'))}</th>
                      <th>Type</th>
                      <th>Vendor</th>
                      <th style={{ textAlign: 'center' }}>Items</th>
                      <th>Status</th>
                      <th style={{ textAlign: 'right' }}>Actions</th>
                    </tr>
                  </thead>
                  <tbody>
                    {currentResultsPage.map((po) => {
                      const isExpanded = !!expandedPo[po.poNumber];
                      const activeCount = po.items ? po.items.filter(item => {
                        const d = String(item['Deletion Indicator'] || item['deletion indicator'] || '').trim().toUpperCase();
                        return d !== 'L';
                      }).length : 0;
                      const ignoredCount = po.items ? po.items.length - activeCount : 0;
                      const isFullyIgnored = activeCount === 0 && ignoredCount > 0;
                      return (
                        <React.Fragment key={po.poNumber}>
                          <tr className={`result-row status-${po.status} ${isExpanded ? 'row-open' : ''} ${isFullyIgnored ? 'result-row-ignored' : ''}`}>
                            <td><strong className="po-num-strong">#{po.poNumber}</strong></td>
                            <td><span className="dt-tag">{po.docType}</span></td>
                            <td className="vendor-td">{po.vendorName}</td>
                            <td style={{ textAlign: 'center' }}>
                              <span>{activeCount}</span>
                              {ignoredCount > 0 && (
                                <span style={{ color: 'var(--red)', fontWeight: '600', marginLeft: '0.25rem' }} title={`${ignoredCount} deleted items ignored`}>
                                  +{ignoredCount}L
                                </span>
                              )}
                            </td>
                            <td>
                              <span className={`status-badge badge-${po.status}`}>
                                {po.status === 'processing' && <span className="spin-sm" />}
                                {po.status === 'success' ? 'Imported' : po.status === 'failed' ? 'Failed' : po.status === 'processing' ? 'Processing' : 'Pending'}
                              </span>
                            </td>
                            <td>
                              <div className="row-actions">
                                {['success', 'failed'].includes(po.status) && (
                                  <>
                                    <button
                                      className={`icon-btn ${isExpanded ? 'active' : ''}`}
                                      onClick={() => togglePoExpand(po.poNumber)}
                                      title={isExpanded ? 'Collapse' : 'View Details'}
                                    >
                                      <svg className={`chevron-icon ${isExpanded ? 'open' : ''}`} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="14" height="14">
                                        <polyline points="6 9 12 15 18 9" />
                                      </svg>
                                    </button>
                                    {po.xmlGenerated && (
                                      <>
                                        <button
                                          className={`icon-btn ${copySuccess === po.poNumber ? 'copied' : ''}`}
                                          onClick={() => copyToClipboard(po.xmlGenerated, po.poNumber)}
                                          title="Copy XML"
                                        >
                                          {copySuccess === po.poNumber
                                            ? <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="14" height="14"><polyline points="20 6 9 17 4 12" /></svg>
                                            : <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="14" height="14"><rect x="9" y="9" width="13" height="13" rx="2" ry="2" /><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1" /></svg>
                                          }
                                        </button>
                                        <button
                                          className="icon-btn"
                                          onClick={() => downloadXml(po.poNumber, po.xmlGenerated)}
                                          title="Download XML"
                                        >
                                          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="14" height="14">
                                            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
                                            <polyline points="7 10 12 15 17 10" />
                                            <line x1="12" y1="15" x2="12" y2="3" />
                                          </svg>
                                        </button>
                                      </>
                                    )}
                                  </>
                                )}
                                {po.status === 'pending' && <span className="pending-dash">—</span>}
                              </div>
                            </td>
                          </tr>

                          {isExpanded && (
                            <tr className="expand-row">
                              <td colSpan="6">
                                <div className="expand-panel">
                                  <div className="expand-section">
                                    <h4 className="expand-title">Line Items</h4>
                                    {renderItemsTable(po.items)}
                                  </div>

                                  {po.status === 'failed' && po.error && (
                                    <div className="tally-err">
                                      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="14" height="14">
                                        <circle cx="12" cy="12" r="10" />
                                        <line x1="12" y1="8" x2="12" y2="12" />
                                        <line x1="12" y1="16" x2="12.01" y2="16" />
                                      </svg>
                                      <strong>Error:</strong> {po.error}
                                    </div>
                                  )}

                                  {po.tallyParsed && (
                                    <div className="expand-section" style={{ marginTop: '12px', padding: '12px', borderRadius: '8px', background: (po.tallyParsed.exceptions > 0 || po.tallyParsed.errors > 0) ? 'rgba(239, 68, 68, 0.08)' : 'rgba(34, 197, 94, 0.08)', border: `1px solid ${(po.tallyParsed.exceptions > 0 || po.tallyParsed.errors > 0) ? '#ef4444' : '#22c55e'}` }}>
                                      <h4 className="expand-title" style={{ margin: '0 0 8px 0', color: 'var(--text-1)' }}>Tally Response Breakdown</h4>
                                      <div style={{ display: 'flex', gap: '16px', flexWrap: 'wrap', fontSize: '13px', fontWeight: '600' }}>
                                        <span style={{ color: po.tallyParsed.created > 0 ? '#22c55e' : 'var(--text-2)' }}>Created: {po.tallyParsed.created}</span>
                                        <span style={{ color: po.tallyParsed.altered > 0 ? '#3b82f6' : 'var(--text-2)' }}>Altered: {po.tallyParsed.altered}</span>
                                        <span style={{ color: po.tallyParsed.exceptions > 0 ? '#ef4444' : 'var(--text-2)', background: po.tallyParsed.exceptions > 0 ? 'rgba(239,68,68,0.2)' : 'transparent', padding: '2px 6px', borderRadius: '4px' }}>Exceptions: {po.tallyParsed.exceptions}</span>
                                        <span style={{ color: po.tallyParsed.errors > 0 ? '#ef4444' : 'var(--text-2)', background: po.tallyParsed.errors > 0 ? 'rgba(239,68,68,0.2)' : 'transparent', padding: '2px 6px', borderRadius: '4px' }}>Errors: {po.tallyParsed.errors}</span>
                                      </div>
                                      {po.tallyParsed.lineError && (
                                        <div style={{ marginTop: '8px', color: '#ef4444', fontSize: '12px', fontWeight: '600' }}>
                                          Line Error: {po.tallyParsed.lineError}
                                        </div>
                                      )}
                                    </div>
                                  )}

                                  {po.tallyResponse && (
                                    <div className="expand-section">
                                      <h4 className="expand-title">Raw Tally Response</h4>
                                      <pre className="xml-code"><code>{po.tallyResponse}</code></pre>
                                    </div>
                                  )}

                                  {po.xmlGenerated && (
                                    <div className="expand-section">
                                      <div className="xml-header">
                                        <h4 className="expand-title">Tally XML Payload</h4>
                                        <div className="xml-btns">
                                          <button
                                            className={`btn-sm-outline ${copySuccess === po.poNumber ? 'copied' : ''}`}
                                            onClick={() => copyToClipboard(po.xmlGenerated, po.poNumber)}
                                          >
                                            {copySuccess === po.poNumber ? '✓ Copied!' : 'Copy XML'}
                                          </button>
                                          <button
                                            className="btn-sm-outline"
                                            onClick={() => downloadXml(po.poNumber, po.xmlGenerated)}
                                          >
                                            Download
                                          </button>
                                        </div>
                                      </div>
                                      <pre className="xml-code"><code>{po.xmlGenerated}</code></pre>
                                    </div>
                                  )}
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

              {/* Pagination */}
              {filteredTotalCount > 0 && (
                <div className="pagination">
                  <span className="pg-info">
                    Showing {indexOfFirstResult + 1}–{Math.min(indexOfLastResult, filteredTotalCount)} of {filteredTotalCount}
                  </span>
                  <div className="pg-controls">
                    <div className="pg-rpp">
                      <span>Rows</span>
                      <select
                        value={rowsPerPage}
                        onChange={(e) => { setRowsPerPage(Number(e.target.value)); setCurrentPage(1); }}
                        className="rpp-select"
                      >
                        {[5, 10, 25, 50, 100].map(s => <option key={s} value={s}>{s}</option>)}
                      </select>
                    </div>
                    <div className="pg-btns">
                      <button className="pg-btn" onClick={() => setCurrentPage(1)} disabled={currentPage === 1}>&laquo;</button>
                      <button className="pg-btn" onClick={() => setCurrentPage(p => Math.max(p - 1, 1))} disabled={currentPage === 1}>&lsaquo;</button>
                      <span className="pg-indicator">{currentPage} / {totalPages}</span>
                      <button className="pg-btn" onClick={() => setCurrentPage(p => Math.min(p + 1, totalPages))} disabled={currentPage === totalPages}>&rsaquo;</button>
                      <button className="pg-btn" onClick={() => setCurrentPage(totalPages)} disabled={currentPage === totalPages}>&raquo;</button>
                    </div>
                  </div>
                </div>
              )}
            </div>
          </div>
        )}
          </>
        ) : activeMode === 'visualizer' ? (
          <XmlVisualizer />
        ) : (
          <DayBookVisualizer />
        )}
      </main>

      {/* Completion Summary Popup Modal */}
      {showSummaryModal && importResults && (
        <div className="modal-overlay" style={{
          position: 'fixed',
          top: 0, left: 0, right: 0, bottom: 0,
          backgroundColor: 'rgba(15, 23, 42, 0.45)',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          zIndex: 9999,
          backdropFilter: 'blur(4px)'
        }}>
          <div className="modal-content" style={{
            background: '#ffffff',
            color: '#0f172a',
            border: '1px solid #e2e8f0',
            borderRadius: '16px',
            padding: '28px',
            maxWidth: '480px',
            width: '90%',
            boxShadow: '0 20px 25px -5px rgba(0, 0, 0, 0.1), 0 8px 10px -6px rgba(0, 0, 0, 0.1)'
          }}>
            <div style={{ textAlign: 'center', marginBottom: '20px' }}>
              <div style={{
                width: '56px', height: '56px', borderRadius: '50%',
                background: totalExceptions > 0 || failedCount > 0 ? '#fef2f2' : '#f0fdf4',
                color: totalExceptions > 0 || failedCount > 0 ? '#dc2626' : '#16a34a',
                display: 'flex', alignItems: 'center', justifyContent: 'center',
                margin: '0 auto 12px auto',
                border: `1px solid ${totalExceptions > 0 || failedCount > 0 ? '#fecaca' : '#bbf7d0'}`
              }}>
                {totalExceptions > 0 || failedCount > 0 ? (
                  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" width="30" height="30">
                    <circle cx="12" cy="12" r="10" />
                    <line x1="12" y1="8" x2="12" y2="12" />
                    <line x1="12" y1="16" x2="12.01" y2="16" />
                  </svg>
                ) : (
                  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" width="30" height="30">
                    <polyline points="20 6 9 17 4 12" />
                  </svg>
                )}
              </div>
              <h2 style={{ fontSize: '20px', fontWeight: '700', margin: '0 0 4px 0', color: '#0f172a' }}>
                {totalExceptions > 0 || failedCount > 0 ? 'Import Complete with Exceptions' : 'Import Completed Successfully'}
              </h2>
              <p style={{ fontSize: '13px', color: '#64748b', margin: 0 }}>
                Processed {totalCount} records in {importTimeTaken}s
              </p>
            </div>

            <div style={{
              display: 'grid',
              gridTemplateColumns: '1fr 1fr',
              gap: '12px',
              marginBottom: '20px'
            }}>
              <div style={{ background: '#f8fafc', padding: '12px 16px', borderRadius: '10px', textAlign: 'center', border: '1px solid #e2e8f0' }}>
                <div style={{ fontSize: '24px', fontWeight: '800', color: '#16a34a' }}>{successCount}</div>
                <div style={{ fontSize: '12px', color: '#64748b', fontWeight: '600', marginTop: '2px' }}>Imported</div>
              </div>
              <div style={{ background: '#f8fafc', padding: '12px 16px', borderRadius: '10px', textAlign: 'center', border: '1px solid #e2e8f0' }}>
                <div style={{ fontSize: '24px', fontWeight: '800', color: failedCount > 0 ? '#dc2626' : '#64748b' }}>{failedCount}</div>
                <div style={{ fontSize: '12px', color: '#64748b', fontWeight: '600', marginTop: '2px' }}>Failed / Exception</div>
              </div>
            </div>

            <div style={{
              background: '#f8fafc',
              borderRadius: '10px',
              padding: '14px 16px',
              marginBottom: '24px',
              border: '1px solid #e2e8f0'
            }}>
              <div style={{ fontSize: '12px', fontWeight: '700', textTransform: 'uppercase', letterSpacing: '0.5px', color: '#64748b', marginBottom: '10px' }}>
                Tally Response Counts
              </div>
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '10px', fontSize: '13px' }}>
                <div><span style={{ color: '#64748b' }}>Created:</span> <strong style={{ color: '#16a34a' }}>{totalCreated}</strong></div>
                <div><span style={{ color: '#64748b' }}>Altered:</span> <strong style={{ color: '#2563eb' }}>{totalAltered}</strong></div>
                <div><span style={{ color: '#64748b' }}>Exceptions:</span> <strong style={{ color: totalExceptions > 0 ? '#dc2626' : '#0f172a' }}>{totalExceptions}</strong></div>
                <div><span style={{ color: '#64748b' }}>Errors:</span> <strong style={{ color: totalErrors > 0 ? '#dc2626' : '#0f172a' }}>{totalErrors}</strong></div>
              </div>
            </div>

            <button
              className="btn-primary"
              onClick={() => setShowSummaryModal(false)}
              style={{
                width: '100%',
                padding: '12px',
                borderRadius: '8px',
                fontWeight: '600',
                cursor: 'pointer',
                border: 'none',
                background: '#2563eb',
                color: '#ffffff',
                fontSize: '14px'
              }}
            >
              View Import Results
            </button>
          </div>
        </div>
      )}

      {/* Ignored Invoices Modal */}
      {showIgnoredModal && (
        <div style={{
          position: 'fixed',
          top: 0,
          left: 0,
          right: 0,
          bottom: 0,
          background: 'rgba(15, 23, 42, 0.65)',
          backdropFilter: 'blur(4px)',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          zIndex: 9999,
          padding: '20px'
        }}>
          <div style={{
            background: 'var(--surface, #ffffff)',
            borderRadius: '16px',
            width: '100%',
            maxWidth: '900px',
            maxHeight: '85vh',
            display: 'flex',
            flexDirection: 'column',
            boxShadow: '0 20px 25px -5px rgba(0, 0, 0, 0.3)',
            border: '1px solid var(--border, #e2e8f0)',
            overflow: 'hidden'
          }}>
            <div style={{
              padding: '20px 24px',
              borderBottom: '1px solid var(--border, #e2e8f0)',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'space-between',
              background: '#fef2f2'
            }}>
              <div>
                <h3 style={{ margin: 0, fontSize: '18px', fontWeight: '700', color: '#991b1b', display: 'flex', alignItems: 'center', gap: '8px' }}>
                  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="20" height="20">
                    <circle cx="12" cy="12" r="10" />
                    <line x1="12" y1="8" x2="12" y2="12" />
                    <line x1="12" y1="16" x2="12.01" y2="16" />
                  </svg>
                  Ignored Purchase Invoices ({ignoredInvoices.length})
                </h3>
                <p style={{ margin: '4px 0 0 0', fontSize: '13px', color: '#b91c1c' }}>
                  These purchase invoices were ignored and not pushed because one or more line items miss a Reference Document (GRN tracking number).
                </p>
              </div>
              <button
                onClick={() => setShowIgnoredModal(false)}
                style={{
                  background: 'none',
                  border: 'none',
                  fontSize: '24px',
                  cursor: 'pointer',
                  color: '#991b1b',
                  lineHeight: '1'
                }}
              >
                &times;
              </button>
            </div>

            <div style={{ padding: '20px 24px', overflowY: 'auto', flex: 1 }}>
              {ignoredInvoices.map((inv, i) => (
                <div key={i} style={{
                  border: '1px solid var(--border, #e2e8f0)',
                  borderRadius: '8px',
                  marginBottom: '16px',
                  background: 'var(--surface-2, #f8fafc)',
                  overflow: 'hidden'
                }}>
                  <div style={{
                    padding: '12px 16px',
                    background: 'var(--surface-3, #f1f5f9)',
                    display: 'flex',
                    justifyContent: 'space-between',
                    alignItems: 'center',
                    borderBottom: '1px solid var(--border, #e2e8f0)'
                  }}>
                    <div>
                      <strong style={{ fontSize: '15px', color: 'var(--text-1, #0f172a)' }}>Invoice #{inv.poNumber}</strong>
                      <span style={{ marginLeft: '12px', fontSize: '13px', color: 'var(--text-2, #64748b)' }}>{inv.vendorName}</span>
                    </div>
                    <span style={{
                      fontSize: '12px',
                      background: '#fee2e2',
                      color: '#991b1b',
                      padding: '4px 8px',
                      borderRadius: '4px',
                      fontWeight: '600'
                    }}>
                      Missing Reference Document / GRN
                    </span>
                  </div>
                  <div style={{ padding: '12px 16px' }}>
                    {renderItemsTable(inv.items)}
                  </div>
                </div>
              ))}
            </div>

            <div style={{
              padding: '16px 24px',
              borderTop: '1px solid var(--border, #e2e8f0)',
              display: 'flex',
              justifyContent: 'flex-end',
              background: 'var(--surface-2, #f8fafc)'
            }}>
              <button
                className="btn-secondary"
                onClick={() => setShowIgnoredModal(false)}
                style={{ padding: '8px 20px', fontWeight: '600' }}
              >
                Close
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

export default App;
