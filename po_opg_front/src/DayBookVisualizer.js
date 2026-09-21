import React, { useState, useEffect } from 'react';

function DayBookVisualizer() {
  const [xmlText, setXmlText] = useState('');
  const [vouchers, setVouchers] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [copySuccess, setCopySuccess] = useState(false);

  useEffect(() => {
    fetchDayBook();
  }, []);

  const fetchDayBook = async () => {
    setLoading(true);
    setError(null);
    try {
      const response = await fetch('http://192.168.1.166:5001/api/daybook-xml');
      if (!response.ok) {
        throw new Error(`Failed to fetch DayBook XML (Status: ${response.status})`);
      }
      const data = await response.text();
      setXmlText(data);
      parseDayBook(data);
    } catch (err) {
      setError(err.message || 'Error fetching DayBook XML.');
    } finally {
      setLoading(false);
    }
  };

  const parseDayBook = (text) => {
    try {
      const parser = new DOMParser();
      const sanitizedText = text.replace(/&#\d+;/g, '');
      const xmlDoc = parser.parseFromString(sanitizedText, 'text/xml');

      const parserError = xmlDoc.querySelector('parsererror');
      if (parserError) {
        throw new Error('XML parsing error: ' + parserError.textContent);
      }

      const list = [];
      const fixedNodes = xmlDoc.querySelectorAll('DBCFIXED');

      fixedNodes.forEach((fixedNode) => {
        const date = fixedNode.querySelector('DBCDATE')?.textContent || '';
        const party = fixedNode.querySelector('DBCPARTY')?.textContent || '';

        // Walk siblings to extract other attributes before the next DBCFIXED
        let sibling = fixedNode.nextElementSibling;
        let buyerName = '';
        let vchNo = '';
        let amount = '';
        let grossAmt = '';
        const ledAmts = [];

        while (sibling && sibling.tagName !== 'DBCFIXED') {
          if (sibling.tagName === 'DBCBUYERNAME') {
            buyerName = sibling.textContent;
          } else if (sibling.tagName === 'DBCVCHNO') {
            vchNo = sibling.textContent;
          } else if (sibling.tagName === 'DBCAMOUNT') {
            amount = sibling.textContent;
          } else if (sibling.tagName === 'DBCGROSSAMT') {
            grossAmt = sibling.textContent;
          } else if (sibling.tagName === 'DBCLEDAMT') {
            ledAmts.push(sibling.textContent);
          }
          sibling = sibling.nextElementSibling;
        }

        list.push({
          date,
          party,
          buyerName,
          vchNo,
          amount,
          grossAmt,
          ledAmts
        });
      });

      setVouchers(list);
    } catch (err) {
      console.error(err);
      setError('Failed to parse DayBook XML: ' + err.message);
    }
  };

  const formatValue = (valStr, addDrCr = false) => {
    if (!valStr || valStr.trim() === '') return '';
    const num = parseFloat(valStr);
    if (isNaN(num)) return valStr;
    const absNum = Math.abs(num);
    const formatted = absNum.toLocaleString('en-IN', {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2
    });
    if (addDrCr) {
      return `${formatted}${num < 0 ? ' Dr' : ' Cr'}`;
    }
    return formatted;
  };

  const copyToClipboard = () => {
    navigator.clipboard.writeText(xmlText);
    setCopySuccess(true);
    setTimeout(() => setCopySuccess(false), 2000);
  };

  const exportTableToCsv = (tableId, filename) => {
    // Manually create flat single-row headers to avoid rowspan/colspan shifts in Excel
    const headers = [
      "Date", "Particulars", "Supplier", "Voucher No.", "Value", "Gross Total",
      "Purchase ZCOL", "IGST Tax", "800012-Clearing Cess", "CGST Tax", "SGST Tax", "Clean Engry Cess"
    ];

    let csvContent = headers.map(h => `"${h.replace(/"/g, '""')}"`).join(",") + "\n";

    vouchers.forEach(v => {
      const rowData = [
        v.date || "",
        v.party || "",
        v.buyerName || "",
        v.vchNo || "",
        v.amount ? formatValue(v.amount) : "0.00",
        v.grossAmt ? formatValue(v.grossAmt, true) : "",
        
        // Ledger values
        v.ledAmts[0] ? formatValue(v.ledAmts[0], true) : "",
        v.ledAmts[1] ? formatValue(v.ledAmts[1], true) : "",
        v.ledAmts[2] ? formatValue(v.ledAmts[2], true) : "",
        v.ledAmts[3] ? formatValue(v.ledAmts[3], true) : "",
        v.ledAmts[4] ? formatValue(v.ledAmts[4], true) : "",
        v.ledAmts[5] ? formatValue(v.ledAmts[5], true) : ""
      ];
      csvContent += rowData.map(val => `"${String(val).replace(/"/g, '""')}"`).join(",") + "\n";
    });

    const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${filename}.csv`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  if (loading) {
    return (
      <div className="xml-loading">
        <span className="spin" />
        <p>Loading Voucher Register DayBook...</p>
      </div>
    );
  }

  if (error) {
    return (
      <div className="xml-error-card">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="40" height="40" className="err-icon">
          <circle cx="12" cy="12" r="10" />
          <line x1="12" y1="8" x2="12" y2="12" />
          <line x1="12" y1="16" x2="12.01" y2="16" />
        </svg>
        <h3>Register Loading Failed</h3>
        <p>{error}</p>
        <button className="btn-primary" onClick={fetchDayBook} style={{ marginTop: '1rem' }}>Retry</button>
      </div>
    );
  }

  return (
    <div className="tally-register-container">
      {/* Tally Window Header Bar */}
      <div className="tally-register-window-header">
        <span className="window-title">Voucher Register</span>
        <span className="window-company">Opg Legacy Data 22-23</span>
        <span className="window-close-btn">&times;</span>
      </div>

      {/* Sub-Header */}
      <div className="tally-register-subheader">
        <span className="subheader-title">List of Purcahse Order ZCOL Vouchers</span>
        <span className="subheader-date-range">1-Apr-23 to 30-Apr-23</span>
      </div>

      {/* Grid Table */}
      <div className="tally-register-grid-wrap">
        <table className="tally-register-table" id="daybook-register-table">
          <thead>
            <tr>
              <th rowSpan="2" className="col-date">Date</th>
              <th rowSpan="2" className="col-particulars">Particulars</th>
              <th rowSpan="2" className="col-supplier">Supplier</th>
              <th rowSpan="2" className="col-vchno">Voucher No.</th>
              <th rowSpan="2" className="col-value align-right">Value</th>
              <th rowSpan="2" className="col-gross align-right">Gross Total</th>
              <th colSpan="6" className="col-ledgers-header align-center">Ledgers &amp; Taxes Breakdown</th>
            </tr>
            <tr>
              <th className="col-subled align-right">Purchase ZCOL</th>
              <th className="col-subled align-right">IGST Tax</th>
              <th className="col-subled align-right">800012-Clearing Cess</th>
              <th className="col-subled align-right">CGST Tax</th>
              <th className="col-subled align-right">SGST Tax</th>
              <th className="col-subled align-right">Clean Engry Cess</th>
            </tr>
          </thead>
          <tbody>
            {vouchers.map((v, idx) => (
              <tr key={idx} className={idx === 0 ? 'selected-row' : ''}>
                <td className="cell-date">{v.date}</td>
                <td className="cell-particulars font-highlight">{v.party}</td>
                <td className="cell-supplier">{v.buyerName}</td>
                <td className="cell-vchno">{v.vchNo}</td>
                <td className="cell-value align-right">{formatValue(v.amount)}</td>
                <td className="cell-gross align-right font-bold">{formatValue(v.grossAmt, true)}</td>
                
                {/* Ledger amounts in order */}
                <td className="cell-led align-right">{formatValue(v.ledAmts[0], true)}</td>
                <td className="cell-led align-right">{formatValue(v.ledAmts[1], true)}</td>
                <td className="cell-led align-right">{formatValue(v.ledAmts[2], true)}</td>
                <td className="cell-led align-right">{formatValue(v.ledAmts[3], true)}</td>
                <td className="cell-led align-right">{formatValue(v.ledAmts[4], true)}</td>
                <td className="cell-led align-right">{formatValue(v.ledAmts[5], true)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <div style={{ marginTop: '1rem', display: 'flex', gap: '0.5rem', justifyContent: 'flex-end' }}>
        <button className="btn-secondary" onClick={copyToClipboard}>
          {copySuccess ? 'Copied XML!' : 'Copy DayBook XML'}
        </button>
        <button className="btn-primary" onClick={() => exportTableToCsv('daybook-register-table', 'DayBook_Voucher_Register')} style={{ background: '#16a34a', borderColor: '#16a34a' }}>
          Export to Excel
        </button>
      </div>
    </div>
  );
}

export default DayBookVisualizer;
