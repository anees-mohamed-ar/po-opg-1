import React, { useState, useEffect } from 'react';

const API_BASE = 'http://192.168.1.166:5001';

export default function TemplatesViewer({ initialModule = 'po' }) {
  const [selectedModule, setSelectedModule] = useState(initialModule);
  const [metadata, setMetadata] = useState(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [searchQuery, setSearchQuery] = useState('');
  const [filterRequired, setFilterRequired] = useState('ALL'); // 'ALL' | 'REQUIRED' | 'OPTIONAL'

  useEffect(() => {
    fetch(`${API_BASE}/api/templates/meta`)
      .then(res => {
        if (!res.ok) throw new Error('Failed to load template metadata from server');
        return res.json();
      })
      .then(data => {
        setMetadata(data);
        setLoading(false);
      })
      .catch(err => {
        console.error('Error fetching template metadata:', err);
        setError(err.message);
        setLoading(false);
      });
  }, []);

  const currentConfig = metadata ? metadata[selectedModule] : null;

  const filteredColumns = (currentConfig?.columns || []).filter(col => {
    const matchesSearch = col.name.toLowerCase().includes(searchQuery.toLowerCase()) ||
      (col.description && col.description.toLowerCase().includes(searchQuery.toLowerCase()));
    if (!matchesSearch) return false;
    if (filterRequired === 'REQUIRED') return col.required;
    if (filterRequired === 'OPTIONAL') return !col.required;
    return true;
  });

  const handleDownloadXlsx = (modKey) => {
    const key = modKey || selectedModule;
    window.location.href = `${API_BASE}/api/templates/download/${key}`;
  };

  const handleDownloadCsv = (modKey) => {
    const key = modKey || selectedModule;
    window.location.href = `${API_BASE}/api/templates/csv/${key}`;
  };

  const modulesList = metadata ? Object.keys(metadata).map(k => ({
    key: k,
    title: metadata[k].title,
    count: metadata[k].columns.length,
    original: metadata[k].originalTotalCols
  })) : [
    { key: 'po', title: 'Purchase Order (PO)' },
    { key: 'grn', title: 'Goods Receipt Note (GRN)' },
    { key: 'purchase', title: 'Purchase Invoice' },
    { key: 'sales_order', title: 'Sales Order' },
    { key: 'fi', title: 'Financial Entry (FI)' },
    { key: 'stock_journal', title: 'Stock Journal' }
  ];

  return (
    <div className="templates-viewer-container" style={{ padding: '32px 40px', maxWidth: '1280px', margin: '0 auto', color: 'var(--text-1)' }}>
      {/* Top Header Banner */}
      <div style={{ marginBottom: '28px', display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', flexWrap: 'wrap', gap: '16px' }}>
        <div>
          <div style={{ display: 'inline-flex', alignItems: 'center', gap: '8px', background: 'rgba(59, 130, 246, 0.1)', color: '#3b82f6', padding: '4px 12px', borderRadius: '20px', fontSize: '12px', fontWeight: '600', marginBottom: '10px' }}>
            <svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" strokeWidth="2.5">
              <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
              <polyline points="14 2 14 8 20 8" />
              <line x1="16" y1="13" x2="8" y2="13" />
              <line x1="16" y1="17" x2="8" y2="17" />
            </svg>
            Clean Client Templates
          </div>
          <h1 style={{ fontSize: '26px', fontWeight: '700', margin: '0 0 6px 0', letterSpacing: '-0.02em' }}>
            Input Excel Header Templates
          </h1>
          <p style={{ margin: 0, color: 'var(--text-2)', fontSize: '14px', maxWidth: '700px', lineHeight: '1.5' }}>
            These templates contain <strong>only the exact columns required by the program</strong> to process into Tally XML. 
            Download and share these minimal sheets with clients instead of requiring hundreds of unneeded SAP export columns.
          </p>
        </div>

        {/* Global Action Buttons */}
        {currentConfig && (
          <div style={{ display: 'flex', gap: '10px', alignItems: 'center' }}>
            <button
              onClick={() => handleDownloadCsv()}
              style={{
                display: 'inline-flex', alignItems: 'center', gap: '8px',
                padding: '10px 18px', borderRadius: '8px', border: '1px solid var(--border)',
                background: 'var(--bg-card)', color: 'var(--text-1)', fontWeight: '600',
                fontSize: '13px', cursor: 'pointer', transition: 'all 0.2s',
                boxShadow: '0 1px 2px rgba(0,0,0,0.05)'
              }}
              title="Download clean CSV format"
            >
              <svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" strokeWidth="2">
                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
                <polyline points="7 10 12 15 17 10" />
                <line x1="12" y1="15" x2="12" y2="3" />
              </svg>
              Export CSV (.csv)
            </button>

            <button
              onClick={() => handleDownloadXlsx()}
              style={{
                display: 'inline-flex', alignItems: 'center', gap: '8px',
                padding: '10px 20px', borderRadius: '8px', border: 'none',
                background: '#10b981', color: '#ffffff', fontWeight: '600',
                fontSize: '13px', cursor: 'pointer', transition: 'all 0.2s',
                boxShadow: '0 4px 12px rgba(16, 185, 129, 0.25)'
              }}
              title="Download ready-to-use Excel template with sample row"
            >
              <svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" strokeWidth="2.5">
                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
                <polyline points="7 10 12 15 17 10" />
                <line x1="12" y1="15" x2="12" y2="3" />
              </svg>
              Download Excel Template (.xlsx)
            </button>
          </div>
        )}
      </div>

      {/* Module Selector Navigation Tabs */}
      <div style={{
        display: 'flex', gap: '8px', overflowX: 'auto', paddingBottom: '6px',
        borderBottom: '1px solid var(--border)', marginBottom: '24px'
      }}>
        {modulesList.map(mod => {
          const isActive = selectedModule === mod.key;
          return (
            <button
              key={mod.key}
              onClick={() => { setSelectedModule(mod.key); setSearchQuery(''); }}
              style={{
                padding: '10px 18px',
                borderRadius: '8px 8px 0 0',
                border: 'none',
                borderBottom: isActive ? '3px solid var(--accent)' : '3px solid transparent',
                background: isActive ? 'rgba(59, 130, 246, 0.08)' : 'transparent',
                color: isActive ? 'var(--accent)' : 'var(--text-2)',
                fontWeight: isActive ? '700' : '500',
                fontSize: '14px',
                cursor: 'pointer',
                whiteSpace: 'nowrap',
                transition: 'all 0.15s'
              }}
            >
              {mod.title}
              {mod.count ? (
                <span style={{
                  marginLeft: '8px',
                  fontSize: '11px',
                  padding: '2px 7px',
                  borderRadius: '10px',
                  background: isActive ? 'var(--accent)' : 'var(--border)',
                  color: isActive ? '#fff' : 'var(--text-2)'
                }}>
                  {mod.count} cols
                </span>
              ) : null}
            </button>
          );
        })}
      </div>

      {loading && (
        <div style={{ padding: '60px', textAlign: 'center', color: 'var(--text-2)' }}>
          <div className="spinner" style={{ margin: '0 auto 16px' }} />
          Loading column templates...
        </div>
      )}

      {error && (
        <div style={{ padding: '20px', background: '#fef2f2', color: '#dc2626', borderRadius: '8px', marginBottom: '24px' }}>
          <strong>Error loading templates:</strong> {error}
        </div>
      )}

      {!loading && currentConfig && (
        <>
          {/* Summary / Comparison Card */}
          <div style={{
            display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(240px, 1fr))',
            gap: '16px', marginBottom: '24px'
          }}>
            <div style={{
              background: 'var(--bg-card)', padding: '18px 20px', borderRadius: '12px',
              border: '1px solid var(--border)', display: 'flex', alignItems: 'center', gap: '16px'
            }}>
              <div style={{
                width: '46px', height: '46px', borderRadius: '10px',
                background: 'rgba(16, 185, 129, 0.12)', color: '#10b981',
                display: 'flex', alignItems: 'center', justifyContent: 'center'
              }}>
                <svg viewBox="0 0 24 24" width="24" height="24" fill="none" stroke="currentColor" strokeWidth="2">
                  <polyline points="20 6 9 17 4 12" />
                </svg>
              </div>
              <div>
                <div style={{ fontSize: '12px', color: 'var(--text-2)', fontWeight: '500' }}>Program Required Columns</div>
                <div style={{ fontSize: '24px', fontWeight: '700', color: '#10b981' }}>
                  {currentConfig.columns.length} <span style={{ fontSize: '14px', fontWeight: '400', color: 'var(--text-2)' }}>columns</span>
                </div>
              </div>
            </div>

            <div style={{
              background: 'var(--bg-card)', padding: '18px 20px', borderRadius: '12px',
              border: '1px solid var(--border)', display: 'flex', alignItems: 'center', gap: '16px'
            }}>
              <div style={{
                width: '46px', height: '46px', borderRadius: '10px',
                background: 'rgba(239, 68, 68, 0.1)', color: '#ef4444',
                display: 'flex', alignItems: 'center', justifyContent: 'center'
              }}>
                <svg viewBox="0 0 24 24" width="22" height="22" fill="none" stroke="currentColor" strokeWidth="2">
                  <line x1="18" y1="6" x2="6" y2="18" />
                  <line x1="6" y1="6" x2="18" y2="18" />
                </svg>
              </div>
              <div>
                <div style={{ fontSize: '12px', color: 'var(--text-2)', fontWeight: '500' }}>Unused SAP Columns Eliminated</div>
                <div style={{ fontSize: '24px', fontWeight: '700', color: 'var(--text-1)' }}>
                  {Math.max(0, currentConfig.originalTotalCols - currentConfig.columns.length)}{' '}
                  <span style={{ fontSize: '13px', fontWeight: '500', color: '#10b981' }}>
                    ({Math.round(((currentConfig.originalTotalCols - currentConfig.columns.length) / currentConfig.originalTotalCols) * 100)}% cleaner)
                  </span>
                </div>
              </div>
            </div>

            <div style={{
              background: 'var(--bg-card)', padding: '18px 20px', borderRadius: '12px',
              border: '1px solid var(--border)', display: 'flex', alignItems: 'center', gap: '16px'
            }}>
              <div style={{
                width: '46px', height: '46px', borderRadius: '10px',
                background: 'rgba(239, 68, 68, 0.12)', color: '#ef4444',
                display: 'flex', alignItems: 'center', justifyContent: 'center'
              }}>
                <svg viewBox="0 0 24 24" width="22" height="22" fill="none" stroke="currentColor" strokeWidth="2.5">
                  <path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z" />
                </svg>
              </div>
              <div>
                <div style={{ fontSize: '12px', color: 'var(--text-2)', fontWeight: '500' }}>Requirement Status</div>
                <div style={{ fontSize: '16px', fontWeight: '700', color: '#ef4444', marginTop: '4px' }}>
                  All {currentConfig.columns.length} Columns Mandatory
                </div>
              </div>
            </div>
          </div>

          {/* Description & Search Bar */}
          <div style={{
            display: 'flex', justifyContent: 'space-between', alignItems: 'center',
            flexWrap: 'wrap', gap: '14px', marginBottom: '18px'
          }}>
            <p style={{ margin: 0, fontSize: '14px', color: 'var(--text-2)' }}>
              {currentConfig.description}
            </p>

            <div style={{ display: 'flex', gap: '10px', alignItems: 'center' }}>
              {/* Search input */}
              <div style={{ position: 'relative' }}>
                <input
                  type="text"
                  placeholder="Search columns..."
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  style={{
                    padding: '7px 12px 7px 32px', borderRadius: '6px',
                    border: '1px solid var(--border)', background: 'var(--bg-card)',
                    color: 'var(--text-1)', fontSize: '13px', width: '240px'
                  }}
                />
                <svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" strokeWidth="2"
                  style={{ position: 'absolute', left: '10px', top: '10px', color: 'var(--text-2)' }}>
                  <circle cx="11" cy="11" r="8" />
                  <line x1="21" y1="21" x2="16.65" y2="16.65" />
                </svg>
              </div>
            </div>
          </div>

          {/* Table of Columns */}
          <div style={{
            background: 'var(--bg-card)', borderRadius: '12px', border: '1px solid var(--border)',
            overflow: 'hidden', boxShadow: '0 1px 3px rgba(0,0,0,0.05)'
          }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left', fontSize: '13px' }}>
              <thead>
                <tr style={{ background: 'rgba(0,0,0,0.02)', borderBottom: '1px solid var(--border)', color: 'var(--text-2)' }}>
                  <th style={{ padding: '12px 16px', width: '45px', textAlign: 'center' }}>#</th>
                  <th style={{ padding: '12px 16px', width: '280px' }}>Exact Excel Header Name</th>
                  <th style={{ padding: '12px 16px', width: '130px' }}>Requirement</th>
                  <th style={{ padding: '12px 16px', width: '200px' }}>Sample Value</th>
                  <th style={{ padding: '12px 16px' }}>Description / Field Role</th>
                </tr>
              </thead>
              <tbody>
                {filteredColumns.map((col, idx) => (
                  <tr
                    key={col.name + idx}
                    style={{
                      borderBottom: '1px solid var(--border)',
                      transition: 'background 0.15s'
                    }}
                    onMouseEnter={(e) => e.currentTarget.style.background = 'rgba(59, 130, 246, 0.03)'}
                    onMouseLeave={(e) => e.currentTarget.style.background = 'transparent'}
                  >
                    <td style={{ padding: '12px 16px', textAlign: 'center', color: 'var(--text-2)', fontWeight: '500' }}>
                      {idx + 1}
                    </td>
                    <td style={{ padding: '12px 16px', fontWeight: '600', color: 'var(--text-1)' }}>
                      <code style={{
                        background: 'rgba(0,0,0,0.05)', padding: '3px 7px', borderRadius: '4px',
                        fontSize: '12.5px', fontFamily: 'monospace', color: 'var(--accent)'
                      }}>
                        {col.name}
                      </code>
                    </td>
                    <td style={{ padding: '12px 16px' }}>
                      {col.required ? (
                        <span style={{
                          display: 'inline-flex', alignItems: 'center', gap: '4px',
                          background: 'rgba(239, 68, 68, 0.1)', color: '#dc2626',
                          padding: '3px 8px', borderRadius: '12px', fontSize: '11.5px', fontWeight: '600'
                        }}>
                          ● Mandatory
                        </span>
                      ) : (
                        <span style={{
                          display: 'inline-flex', alignItems: 'center', gap: '4px',
                          background: 'rgba(100, 116, 139, 0.1)', color: 'var(--text-2)',
                          padding: '3px 8px', borderRadius: '12px', fontSize: '11.5px', fontWeight: '500'
                        }}>
                          Optional
                        </span>
                      )}
                    </td>
                    <td style={{ padding: '12px 16px', color: 'var(--text-1)' }}>
                      <span style={{ fontFamily: 'monospace', fontSize: '12px', opacity: 0.9 }}>
                        {col.sample !== '' && col.sample !== undefined ? String(col.sample) : <em style={{ color: 'var(--text-2)' }}>(blank)</em>}
                      </span>
                    </td>
                    <td style={{ padding: '12px 16px', color: 'var(--text-2)', lineHeight: '1.4' }}>
                      {col.description}
                    </td>
                  </tr>
                ))}
                {filteredColumns.length === 0 && (
                  <tr>
                    <td colSpan="5" style={{ padding: '40px', textAlign: 'center', color: 'var(--text-2)' }}>
                      No column headers match the search or filter criteria.
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>

          {/* Quick Client Instructions Box */}
          <div style={{
            marginTop: '28px', background: 'rgba(59, 130, 246, 0.05)', borderRadius: '12px',
            border: '1px solid rgba(59, 130, 246, 0.2)', padding: '20px 24px', display: 'flex', gap: '16px'
          }}>
            <div style={{ color: 'var(--accent)', marginTop: '2px' }}>
              <svg viewBox="0 0 24 24" width="22" height="22" fill="none" stroke="currentColor" strokeWidth="2">
                <circle cx="12" cy="12" r="10" />
                <line x1="12" y1="16" x2="12" y2="12" />
                <line x1="12" y1="8" x2="12.01" y2="8" />
              </svg>
            </div>
            <div>
              <h4 style={{ margin: '0 0 6px 0', fontSize: '14px', fontWeight: '700', color: 'var(--accent)' }}>
                Instructions for Clients / Data Providers:
              </h4>
              <ul style={{ margin: 0, paddingLeft: '18px', color: 'var(--text-1)', fontSize: '13px', lineHeight: '1.6' }}>
                <li>Do <strong>not</strong> rename or alter the column headers in Row 1; our program maps these exact column names.</li>
                <li><strong>All listed columns are mandatory</strong>: they are directly read by the program to generate vouchers and ledgers in Tally.</li>
                <li>For date fields, both <code>DD.MM.YYYY</code> (e.g. <code>01.04.2023</code>) and <code>YYYYMMDD</code> formats are supported.</li>
                <li>Once filled, save as <code>.xlsx</code> and drag & drop into the <strong>Excel Importer</strong> tab.</li>
              </ul>
            </div>
          </div>
        </>
      )}
    </div>
  );
}
