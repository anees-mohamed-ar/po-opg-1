import React, { useState, useEffect } from 'react';

function XmlVisualizer() {
  const [xmlText, setXmlText] = useState('');
  const [vouchers, setVouchers] = useState([]);
  const [selectedVoucherIndex, setSelectedVoucherIndex] = useState(0);
  const [xmlFiles, setXmlFiles] = useState([]);
  const [selectedFile, setSelectedFile] = useState('Purchase Order_4600001048.xml');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [copySuccess, setCopySuccess] = useState(false);
  const [showXml, setShowXml] = useState(false);

  // Date range picker modal for live Tally integration
  const [fromDate, setFromDate] = useState('2022-04-01');
  const [toDate, setToDate] = useState('2023-03-31');
  const [showTallyModal, setShowTallyModal] = useState(true);

  // Pagination & Filtering state
  const [searchTerm, setSearchTerm] = useState('');
  const [currentPage, setCurrentPage] = useState(1);
  const [rowsPerPage, setRowsPerPage] = useState(25);
  const [colFilters, setColFilters] = useState({});

  useEffect(() => {
    fetchXmlFiles();
  }, []);

  useEffect(() => {
    if (selectedFile && !selectedFile.startsWith('Tally_Live_Export')) {
      fetchXml(selectedFile);
    }
  }, [selectedFile]);

  const fetchFromTally = async (from, to) => {
    setLoading(true);
    setError(null);
    setShowTallyModal(false);
    try {
      const formattedFrom = from.replace(/-/g, '');
      const formattedTo = to.replace(/-/g, '');

      const response = await fetch('http://192.168.1.166:5001/api/fetch-from-tally', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          fromDate: formattedFrom,
          toDate: formattedTo
        })
      });

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error || errorData.details || `Failed to fetch from Tally (Status: ${response.status})`);
      }

      const data = await response.text();
      setXmlText(data);
      setSelectedFile(`Tally_Live_Export_${formattedFrom}_${formattedTo}`);
      parseXmlData(data);
    } catch (err) {
      setError(err.message || 'Error fetching data from Tally.');
    } finally {
      setLoading(false);
    }
  };

  const fetchXmlFiles = async () => {
    try {
      const response = await fetch('http://192.168.1.166:5001/api/xml-files');
      if (response.ok) {
        const data = await response.json();
        setXmlFiles(data.xmlFiles || []);
        // If "DayBook copy.xml" exists in the files list, set it as default
        if (data.xmlFiles && data.xmlFiles.includes('DayBook copy.xml')) {
          setSelectedFile('DayBook copy.xml');
        } else if (data.xmlFiles && data.xmlFiles.length > 0) {
          setSelectedFile(data.xmlFiles[0]);
        }
      }
    } catch (err) {
      console.error('Error fetching XML files list:', err);
    }
  };

  const fetchXml = async (fileName) => {
    setLoading(true);
    setError(null);
    try {
      const response = await fetch(`http://192.168.1.166:5001/api/po-xml?file=${encodeURIComponent(fileName)}`);
      if (!response.ok) {
        throw new Error(`Failed to fetch XML (Status: ${response.status})`);
      }
      const data = await response.text();
      setXmlText(data);
      parseXmlData(data);
    } catch (err) {
      setError(err.message || 'Error fetching XML data.');
    } finally {
      setLoading(false);
    }
  };

  const parseXmlData = (text) => {
    try {
      const sanitizedText = text.replace(/&#\d+;/g, '');
      const parser = new DOMParser();
      const xmlDoc = parser.parseFromString(sanitizedText, 'text/xml');

      // Check parsing error
      const parserError = xmlDoc.querySelector('parsererror');
      if (parserError) {
        throw new Error('XML parsing error: ' + parserError.textContent);
      }

      const voucherNodes = xmlDoc.getElementsByTagName('VOUCHER');
      if (voucherNodes.length === 0) {
        throw new Error('Could not find any VOUCHER element in XML.');
      }

      const list = [];
      const udfKeys = [
        { inr: 'TRANSOCFREINR', label: 'TRANSOCFRE', sub: 'TRANSOCFRESUB', defaultName: 'Ocean Freight' },
        { inr: 'TRANSDDWGTINR', label: 'TRANSDDWGT', sub: 'TRANSDDWGTSUB', defaultName: 'Dead Weight' },
        { inr: 'TRANSDEMINR', label: 'TRANSDEM', sub: 'TRANSDEMSUB', defaultName: 'Demurrage' },
        { inr: 'TRANSINSURINR', label: 'TRANSINSUR', sub: 'TRANSINSURSUB', defaultName: 'Insurance' },
        { inr: 'TRANSSAMPLEINR', label: 'TRANSSAMPLE', sub: 'TRANSSAMPLESUB', defaultName: 'Sample' },
        { inr: 'TRANSSAMPLE2INR', label: 'TRANSSAMPLE2', sub: 'TRANSSAMPLE2SUB', defaultName: 'Sample 2' },
        { inr: 'TRANSSTOCKINR', label: 'TRANSSTOCK', sub: 'TRANSSTOCKSUB', defaultName: 'Stock' },
        { inr: 'TRANSDRAFTINR', label: 'TRANSDRAFT', sub: 'TRANSDRAFTSUB', defaultName: 'Draft' },
        { inr: 'TRANSLIAINR', label: 'TRANSLIA', sub: 'TRANSLIASUB', defaultName: 'Liaison' },
        { inr: 'TRANSWHARINR', label: 'TRANSWHAR', sub: 'TRANSWHARSUB', defaultName: 'Wharfage' },
        { inr: 'TRANSEQUIPINR', label: 'TRANSEQUIP', sub: 'TRANSEQUIPSUB', defaultName: 'Equipment' },
        { inr: 'TRANSDISCINR', label: 'TRANSDISC', sub: 'TRANSDISCSUB', defaultName: 'Discharge' },
        { inr: 'TRANSSTEVINR', label: 'TRANSSTEV', sub: 'TRANSSTEVSUB', defaultName: 'Stevedoring' },
        { inr: 'TRANSWHARENINR', label: 'TRANSWHAREN', sub: 'TRANSWHARENSUB', defaultName: 'Wharfage Energy' },
        { inr: 'TRANSTRANSINR', label: 'TRANSTRANS', sub: 'TRANSTRANSSUB', defaultName: 'Transfer' },
        { inr: 'TRANSHANDLINGINR', label: 'TRANSHANDLING', sub: 'TRANSHANDLINGSUB', defaultName: 'Handling' }
      ];

      Array.from(voucherNodes).forEach(voucher => {
        const voucherNum = voucher.getElementsByTagName('VOUCHERNUMBER')[0]?.textContent || '';
        const rawDate = voucher.getElementsByTagName('DATE')[0]?.textContent || '';
        const partyName = voucher.getElementsByTagName('PARTYNAME')[0]?.textContent || voucher.getElementsByTagName('PARTYLEDGERNAME')[0]?.textContent || '';
        const stateName = voucher.getElementsByTagName('STATENAME')[0]?.textContent || '';
        const placeOfSupply = voucher.getElementsByTagName('PLACEOFSUPPLY')[0]?.textContent || '';
        const voucherType = voucher.getElementsByTagName('VOUCHERTYPENAME')[0]?.textContent || '';
        const guid = voucher.getElementsByTagName('GUID')[0]?.textContent || '';
        const alterId = voucher.getElementsByTagName('ALTERID')[0]?.textContent || '';
        const masterId = voucher.getElementsByTagName('MASTERID')[0]?.textContent || '';

        // Format Date YYYYMMDD -> DD-MMM-YYYY
        let formattedDate = rawDate;
        if (rawDate && rawDate.length === 8) {
          const year = rawDate.substring(0, 4);
          const month = parseInt(rawDate.substring(4, 6), 10) - 1;
          const day = parseInt(rawDate.substring(6, 8), 10);
          const dateObj = new Date(year, month, day);
          if (!isNaN(dateObj.getTime())) {
            formattedDate = dateObj.toLocaleDateString('en-IN', {
              day: '2-digit',
              month: 'short',
              year: 'numeric'
            });
          }
        }

        // 2. Get stock items
        const itemNodes = voucher.getElementsByTagName('ALLINVENTORYENTRIES.LIST');
        const items = Array.from(itemNodes).map(node => {
          const rawAmount = node.getElementsByTagName('AMOUNT')[0]?.textContent || '0';
          const parsedAmount = Math.abs(parseFloat(rawAmount));
          return {
            name: node.getElementsByTagName('STOCKITEMNAME')[0]?.textContent || '',
            rate: node.getElementsByTagName('RATE')[0]?.textContent || '',
            amount: parsedAmount,
            actualQty: node.getElementsByTagName('ACTUALQTY')[0]?.textContent || '',
            billedQty: node.getElementsByTagName('BILLEDQTY')[0]?.textContent || '',
          };
        });

        // 3. Get ledger entries
        const ledgerMap = {};
        const ledgerNodes = voucher.getElementsByTagName('LEDGERENTRIES.LIST');
        const ledgers = Array.from(ledgerNodes).map(node => {
          const name = node.getElementsByTagName('LEDGERNAME')[0]?.textContent || '';
          const rawAmount = node.getElementsByTagName('AMOUNT')[0]?.textContent || '0';
          const parsedAmount = parseFloat(rawAmount);
          if (name) {
            ledgerMap[name.trim()] = parsedAmount;
          }
          return {
            name,
            amount: parsedAmount,
            isParty: node.getElementsByTagName('ISPARTYLEDGER')[0]?.textContent === 'Yes',
            deemedPositive: node.getElementsByTagName('ISDEEMEDPOSITIVE')[0]?.textContent === 'Yes',
          };
        });

        // Also gather accounting allocations under inventory entries
        const allocNodes = voucher.getElementsByTagName('ACCOUNTINGALLOCATIONS.LIST');
        Array.from(allocNodes).forEach(node => {
          const name = node.getElementsByTagName('LEDGERNAME')[0]?.textContent || '';
          const rawAmount = node.getElementsByTagName('AMOUNT')[0]?.textContent || '0';
          const parsedAmount = parseFloat(rawAmount);
          if (name) {
            ledgerMap[name.trim()] = parsedAmount;
          }
        });

        // 4. Get UDF Additional Charges
        // Parse every slot that has a label name, regardless of amount being zero.
        // Store as udfByName map keyed by the actual charge label for O(1) lookup
        // (same charge type may land in different slots across different vouchers).
        const udfs = [];
        const udfByName = {};
        udfKeys.forEach(k => {
          const name = (voucher.getElementsByTagName('UDF:' + k.label)[0]?.textContent || voucher.getElementsByTagName(k.label)[0]?.textContent || '').trim();
          if (!name) return; // skip slots with no label
          const valStr = voucher.getElementsByTagName('UDF:' + k.inr)[0]?.textContent || voucher.getElementsByTagName(k.inr)[0]?.textContent || '';
          const amt = valStr ? (parseFloat(valStr) || 0) : 0;
          const ledgerCode = voucher.getElementsByTagName('UDF:' + k.sub)[0]?.textContent || voucher.getElementsByTagName(k.sub)[0]?.textContent || '';
          const entry = { keyName: k.inr, name, ledgerCode, amount: amt };
          udfs.push(entry);
          udfByName[name] = entry; // last write wins if duplicate name (shouldn't happen)
        });

        list.push({
          udfs,
          udfByName,
          voucherNum,
          date: formattedDate,
          rawDate,
          partyName,
          stateName,
          placeOfSupply,
          voucherType,
          guid,
          alterId,
          masterId,
          items,
          ledgers,
          ledgerMap
        });
      });

      setVouchers(list);
      setSelectedVoucherIndex(0);
      setCurrentPage(1);
    } catch (err) {
      console.error(err);
      setError('Failed to extract data from XML: ' + err.message);
    }
  };

  const copyToClipboard = () => {
    navigator.clipboard.writeText(xmlText);
    setCopySuccess(true);
    setTimeout(() => setCopySuccess(false), 2000);
  };

  const downloadXml = () => {
    const blob = new Blob([xmlText], { type: 'application/xml' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = selectedFile;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const formatCurrency = (val) => {
    return Number(val).toLocaleString('en-IN', {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2
    });
  };

  const formatValue = (val, addDrCr = false) => {
    if (val === undefined || val === null) return '';
    const num = typeof val === 'number' ? val : parseFloat(val);
    if (isNaN(num)) return val;
    const absNum = Math.abs(num);
    const formatted = absNum.toLocaleString('en-IN', {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2
    });
    if (addDrCr) {
      return `${formatted} Dr`;
    }
    return formatted;
  };

  const getLedgerAmt = (v, nameKey) => {
    for (const key of Object.keys(v.ledgerMap || {})) {
      if (key.toLowerCase().includes(nameKey.toLowerCase())) {
        return v.ledgerMap[key];
      }
    }
    return undefined;
  };

  // Collect unique dropdown options across ALL vouchers
  const uniqueParticulars = Array.from(new Set(vouchers.map(v => v.partyName).filter(Boolean))).sort();
  const uniqueVoucherTypes = Array.from(new Set(vouchers.map(v => v.voucherType).filter(Boolean))).sort();

  // Filtering Logic
  const filteredVouchers = vouchers.filter(v => {
    // 1. General search filter
    const matchStr = `${v.voucherNum} ${v.partyName} ${v.voucherType}`.toLowerCase();
    if (!matchStr.includes(searchTerm.toLowerCase())) return false;

    // 2. Column-wise filters
    for (const [key, value] of Object.entries(colFilters)) {
      if (!value) continue;
      const cleanFilterVal = value.toLowerCase().trim();

      if (key === 'date') {
        // value is YYYY-MM-DD, convert to YYYYMMDD to match rawDate
        const dateMatch = value.replace(/-/g, '');
        if (v.rawDate !== dateMatch) return false;
      } else if (key === 'particulars') {
        if (v.partyName !== value) return false;
      } else if (key === 'voucherNum') {
        if (!v.voucherNum?.toLowerCase().includes(cleanFilterVal)) return false;
      } else if (key === 'voucherType') {
        if (v.voucherType !== value) return false;
      } else if (key === 'value') {
        const totalItemVal = v.items.reduce((sum, item) => sum + item.amount, 0);
        if (!formatValue(totalItemVal).toLowerCase().includes(cleanFilterVal)) return false;
      } else if (key === 'grossTotal') {
        const partyLedger = v.ledgers.find(l => l.isParty);
        const totalAmount = partyLedger ? Math.abs(partyLedger.amount) : 0;
        if (!formatValue(totalAmount).toLowerCase().includes(cleanFilterVal)) return false;
      } else if (key === 'purchaseZcol') {
        const amt = getLedgerAmt(v, 'Purchase ZCOL');
        if (amt === undefined || !formatValue(amt).toLowerCase().includes(cleanFilterVal)) return false;
      } else if (key === 'igst') {
        const amt = getLedgerAmt(v, 'IGST');
        if (amt === undefined || !formatValue(amt).toLowerCase().includes(cleanFilterVal)) return false;
      } else if (key === 'clearingCess') {
        const amt = getLedgerAmt(v, 'Clean Energy Cess') || getLedgerAmt(v, '800012');
        if (amt === undefined || !formatValue(amt).toLowerCase().includes(cleanFilterVal)) return false;
      } else if (key === 'cgst') {
        const amt = getLedgerAmt(v, 'CGST');
        if (amt === undefined || !formatValue(amt).toLowerCase().includes(cleanFilterVal)) return false;
      } else if (key === 'sgst') {
        const amt = getLedgerAmt(v, 'SGST');
        if (amt === undefined || !formatValue(amt).toLowerCase().includes(cleanFilterVal)) return false;
      } else if (key === 'cleanEnergy') {
        const amt = getLedgerAmt(v, 'Clean Engry Cess') || getLedgerAmt(v, 'Clean Entry Cess');
        if (amt === undefined || !formatValue(amt).toLowerCase().includes(cleanFilterVal)) return false;
      } else {
        // Dynamic UDF column filter
        const u = v.udfByName?.[key];
        if (!u || !u.amount || !formatValue(u.amount).toLowerCase().includes(cleanFilterVal)) return false;
      }
    }
    return true;
  });

  // Pagination Logic
  const indexOfLastRow = currentPage * rowsPerPage;
  const indexOfFirstRow = indexOfLastRow - rowsPerPage;
  const currentRows = filteredVouchers.slice(indexOfFirstRow, indexOfLastRow);
  const totalPages = Math.ceil(filteredVouchers.length / rowsPerPage);

  const handlePageChange = (pageNumber) => {
    if (pageNumber >= 1 && pageNumber <= totalPages) {
      setCurrentPage(pageNumber);
    }
  };

  // Collect unique UDF charge labels across ALL vouchers (not just current page)
  // so columns remain stable while filtering/paging.
  const activeUdfNames = new Set();
  vouchers.forEach(v => {
    v.udfs.forEach(udf => {
      if (udf.name && udf.name.trim() !== '') {
        activeUdfNames.add(udf.name.trim());
      }
    });
  });
  const activeUdfCols = Array.from(activeUdfNames).map(name => ({ name }));

  const escapeXml = (str) => {
    if (str === undefined || str === null) return '';
    return String(str).replace(/[<>&'"]/g, (c) => {
      switch (c) {
        case '<': return '&lt;';
        case '>': return '&gt;';
        case '&': return '&amp;';
        case '\'': return '&apos;';
        case '"': return '&quot;';
        default: return c;
      }
    });
  };

  const exportTableToExcel = (filename) => {
    const headers = [
      "Date", "Particulars", "Voucher No.", "Voucher Type", "Value", "Gross Total",
      "Purchase ZCOL", "IGST Tax", "800012-Clearing Cess", "CGST Tax", "SGST Tax", "Clean Engry Cess"
    ];

    activeUdfCols.forEach(udf => {
      headers.push(udf.name);
    });

    const rows = [];
    filteredVouchers.forEach(v => {
      const totalItemVal = v.items.reduce((sum, item) => sum + item.amount, 0);
      const partyLedger = v.ledgers.find(l => l.isParty);
      const totalAmount = partyLedger ? Math.abs(partyLedger.amount) : 0;

      const rowData = [
        v.date || "",
        v.partyName || "",
        v.voucherNum || "",
        v.voucherType || "",
        totalItemVal ? formatValue(totalItemVal) : "0.00",
        totalAmount ? formatValue(totalAmount) + " Cr" : "0.00 Cr",

        getLedgerAmt(v, 'Purchase ZCOL') !== undefined ? formatValue(getLedgerAmt(v, 'Purchase ZCOL'), true) : "",
        getLedgerAmt(v, 'IGST') !== undefined ? formatValue(getLedgerAmt(v, 'IGST'), true) : "",
        (getLedgerAmt(v, 'Clean Energy Cess') || getLedgerAmt(v, '800012')) !== undefined ? formatValue(getLedgerAmt(v, 'Clean Energy Cess') || getLedgerAmt(v, '800012'), true) : "",
        getLedgerAmt(v, 'CGST') !== undefined ? formatValue(getLedgerAmt(v, 'CGST'), true) : "",
        getLedgerAmt(v, 'SGST') !== undefined ? formatValue(getLedgerAmt(v, 'SGST'), true) : "",
        (getLedgerAmt(v, 'Clean Engry Cess') || getLedgerAmt(v, 'Clean Entry Cess')) !== undefined ? formatValue(getLedgerAmt(v, 'Clean Engry Cess') || getLedgerAmt(v, 'Clean Entry Cess'), true) : ""
      ];

      activeUdfCols.forEach(udf => {
        const u = v.udfByName?.[udf.name];
        rowData.push(u && u.amount ? formatValue(u.amount, true) : "");
      });

      rows.push(rowData);
    });

    // Calculate column widths based on longest text
    const colWidths = headers.map((header, colIndex) => {
      let maxLen = header.length;
      rows.forEach(row => {
        const val = row[colIndex] ? String(row[colIndex]).length : 0;
        if (val > maxLen) {
          maxLen = val;
        }
      });
      return Math.max(60, Math.min(300, maxLen * 8.5 + 20));
    });

    let xml = `<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">
 <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
  <Author>OPG Gateway</Author>
  <Created>${new Date().toISOString()}</Created>
 </DocumentProperties>
 <Styles>
  <Style ss:ID="Default" ss:Name="Normal">
   <Alignment ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/>
   </Borders>
   <Font ss:FontName="Consolas" x:CharSet="1" x:Family="Modern" ss:Size="10.5" ss:Color="#0F172A"/>
   <Interior/>
   <NumberFormat/>
   <Protection/>
  </Style>
  <Style ss:ID="Header">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#1E293B"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#1E293B"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#1E293B"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#1E293B"/>
   </Borders>
   <Font ss:FontName="Consolas" x:Family="Modern" ss:Size="10.5" ss:Color="#FFFFFF" ss:Bold="1"/>
   <Interior ss:Color="#334155" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="CellLeft">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
   <Font ss:FontName="Consolas" x:CharSet="1" x:Family="Modern" ss:Size="10.5" ss:Color="#0F172A"/>
  </Style>
  <Style ss:ID="CellRight">
   <Alignment ss:Horizontal="Right" ss:Vertical="Center"/>
   <Font ss:FontName="Consolas" x:CharSet="1" x:Family="Modern" ss:Size="10.5" ss:Color="#0F172A"/>
  </Style>
  <Style ss:ID="CellCenter">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Font ss:FontName="Consolas" x:CharSet="1" x:Family="Modern" ss:Size="10.5" ss:Color="#0F172A"/>
  </Style>
 </Styles>
 <Worksheet ss:Name="Voucher Register">
  <Table>`;

    colWidths.forEach(width => {
      xml += `\n   <Column ss:Width="${width.toFixed(1)}"/>`;
    });

    xml += `\n   <Row ss:Height="25">`;
    headers.forEach(h => {
      xml += `\n    <Cell ss:StyleID="Header"><Data ss:Type="String">${escapeXml(h)}</Data></Cell>`;
    });
    xml += `\n   </Row>`;

    rows.forEach(row => {
      xml += `\n   <Row ss:Height="20">`;
      row.forEach((cellVal, cellIdx) => {
        let styleId = 'CellLeft';
        if (cellIdx === 0 || cellIdx === 3) {
          styleId = 'CellCenter';
        } else if (cellIdx >= 5) {
          styleId = 'CellRight';
        }

        const isNum = !isNaN(parseFloat(cellVal.replace(/ Dr| Cr/g, '').replace(/,/g, ''))) && (cellIdx >= 5);
        const type = isNum ? 'Number' : 'String';
        const cleanVal = isNum ? cellVal.replace(/ Dr| Cr/g, '').replace(/,/g, '') : cellVal;

        xml += `\n    <Cell ss:StyleID="${styleId}"><Data ss:Type="${type}">${escapeXml(cleanVal)}</Data></Cell>`;
      });
      xml += `\n   </Row>`;
    });

    xml += `\n  </Table>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <PageSetup>
    <Header x:Margin="0.3"/>
    <Footer x:Margin="0.3"/>
    <PageMargins x:Bottom="0.75" x:Left="0.7" x:Right="0.7" x:Top="0.75"/>
   </PageSetup>
   <Print>
    <ValidPrinterInfo/>
    <PaperSizeIndex>9</PaperSizeIndex>
    <HorizontalResolution>600</HorizontalResolution>
    <VerticalResolution>600</VerticalResolution>
   </Print>
   <Selected/>
   <Panes>
    <Pane>
     <Number>3</Number>
     <ActiveRow>0</ActiveRow>
     <ActiveCol>0</ActiveCol>
    </Pane>
   </Panes>
   <ProtectObjects>False</ProtectObjects>
   <ProtectScenarios>False</ProtectScenarios>
  </WorksheetOptions>
 </Worksheet>
</Workbook>`;

    const blob = new Blob([xml], { type: "application/vnd.ms-excel;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${filename}.xls`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  if (loading) {
    return (
      <div className="xml-loading">
        <span className="spin" />
        <p>Fetching and parsing XML file: {selectedFile}...</p>
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
        <h3>Visualization Failed</h3>
        <p>{error}</p>
        <div style={{ marginTop: '1rem', display: 'flex', gap: '1rem' }}>
          <select
            value={selectedFile}
            onChange={(e) => setSelectedFile(e.target.value)}
            className="search-input"
            style={{ minWidth: '200px' }}
          >
            {xmlFiles.map(f => (
              <option key={f} value={f}>{f}</option>
            ))}
          </select>
          <button className="btn-primary" onClick={() => fetchXml(selectedFile)}>Retry</button>
        </div>
      </div>
    );
  }

  // Safe Index Selection
  const selectedVoucher = filteredVouchers[selectedVoucherIndex] || filteredVouchers[0] || vouchers[0];
  const partyLedger = selectedVoucher?.ledgers.find(l => l.isParty);
  const totalAmount = partyLedger ? Math.abs(partyLedger.amount) : 0;
  const totalItemVal = selectedVoucher ? selectedVoucher.items.reduce((sum, item) => sum + item.amount, 0) : 0;

  return (
    <div className="visualizer-container">
      <div className="visualizer-header">
        <div>
          <span className="visualizer-badge">XML Visualizer</span>
          <h1 className="page-title">XML Voucher Register</h1>
          <p className="page-desc">Viewing file: <strong>{selectedFile}</strong> ({vouchers.length} Vouchers)</p>
        </div>
        <div className="visualizer-actions" style={{ display: 'flex', gap: '0.75rem', alignItems: 'center' }}>
          <div style={{ display: 'flex', flexDirection: 'column', gap: '2px' }}>
            <label style={{ fontSize: '0.75rem', color: 'var(--text-2)', fontWeight: 'bold' }}>Source XML File:</label>
            <select
              value={selectedFile}
              onChange={(e) => {
                setSelectedFile(e.target.value);
                setSelectedVoucherIndex(0);
              }}
              className="search-input"
              style={{ padding: '0.4rem 0.8rem', minWidth: '250px', background: 'var(--surface-2)', color: 'var(--text-1)', border: '1px solid var(--border)' }}
            >
              {xmlFiles.map(f => (
                <option key={f} value={f}>{f}</option>
              ))}
            </select>
          </div>
          <button
            className="btn-primary"
            onClick={() => setShowTallyModal(true)}
            style={{ padding: '0.6rem 1rem', background: '#3b82f6', borderColor: '#3b82f6', color: 'white' }}
          >
            Fetch Live Tally
          </button>
          <button className="btn-secondary" onClick={() => setShowXml(!showXml)}>
            {showXml ? 'Hide XML' : 'View XML'}
          </button>
          <button className="btn-secondary" onClick={copyToClipboard}>
            {copySuccess ? '✓ Copied!' : 'Copy XML'}
          </button>
          <button className="btn-primary" onClick={downloadXml}>
            Download XML
          </button>
          <button
            className="btn-primary"
            onClick={() => exportTableToExcel(`Voucher_Register_${selectedFile.replace('.xml', '')}`)}
            style={{ background: '#16a34a', borderColor: '#16a34a' }}
          >
            Export to Excel
          </button>
        </div>
      </div>

      {/* FILTER & PAGINATION BAR */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', background: 'var(--surface)', padding: '0.75rem 1.25rem', border: '1px solid var(--border)', borderBottom: 'none', borderRadius: 'var(--radius) var(--radius) 0 0', gap: '1rem', flexWrap: 'wrap' }}>
        <div style={{ display: 'flex', gap: '1rem', alignItems: 'center' }}>
          <input
            type="text"
            placeholder="Search PO#, Party, Voucher Type..."
            value={searchTerm}
            onChange={(e) => {
              setSearchTerm(e.target.value);
              setCurrentPage(1);
              setSelectedVoucherIndex(0);
            }}
            className="search-input"
            style={{ width: '300px' }}
          />
          <span style={{ fontSize: '0.85rem', color: 'var(--text-2)' }}>
            Found: <strong>{filteredVouchers.length}</strong> vouchers
          </span>
          {Object.values(colFilters).some(v => v !== '') && (
            <button
              className="btn-secondary"
              onClick={() => {
                setColFilters({});
                setCurrentPage(1);
              }}
              style={{ padding: '0.25rem 0.6rem', fontSize: '0.8rem', marginLeft: '0.5rem', background: 'var(--accent-light)', borderColor: 'var(--accent)', color: 'var(--accent)' }}
            >
              Clear Column Filters
            </button>
          )}
        </div>

        {totalPages > 1 && (
          <div style={{ display: 'flex', gap: '0.5rem', alignItems: 'center' }}>
            <span style={{ fontSize: '0.85rem', color: 'var(--text-2)' }}>Rows per page:</span>
            <select
              value={rowsPerPage}
              onChange={(e) => {
                setRowsPerPage(Number(e.target.value));
                setCurrentPage(1);
                setSelectedVoucherIndex(0);
              }}
              className="search-input"
              style={{ padding: '0.25rem 0.5rem', marginRight: '1rem' }}
            >
              <option value={10}>10</option>
              <option value={25}>25</option>
              <option value={50}>50</option>
              <option value={100}>100</option>
            </select>

            <button
              className="btn-secondary"
              onClick={() => handlePageChange(currentPage - 1)}
              disabled={currentPage === 1}
              style={{ padding: '0.25rem 0.5rem', fontSize: '0.85rem' }}
            >
              ◀ Prev
            </button>
            <span style={{ fontSize: '0.85rem', color: 'var(--text-1)' }}>
              Page <strong>{currentPage}</strong> of {totalPages}
            </span>
            <button
              className="btn-secondary"
              onClick={() => handlePageChange(currentPage + 1)}
              disabled={currentPage === totalPages}
              style={{ padding: '0.25rem 0.5rem', fontSize: '0.85rem' }}
            >
              Next ▶
            </button>
          </div>
        )}
      </div>

      {/* TALLY REGISTER COLUMN VIEW */}
      <div className="tally-register-container" style={{ marginBottom: '1.5rem', borderRadius: '0 0 var(--radius) var(--radius)' }}>
        <div className="tally-register-window-header">
          <span className="window-title">Voucher Register</span>
          <span className="window-company">Opg Legacy Data Apr-Dec(2023)</span>
          <span className="window-close-btn">&times;</span>
        </div>
        <div className="tally-register-subheader">
          <span className="subheader-title">Purchase Order Column View</span>
          <span className="subheader-date-range">{selectedFile}</span>
        </div>
        <div className="tally-register-grid-wrap">
          <table className="tally-register-table" id="po-register-table">
            <thead>
              <tr>
                <th rowSpan="2" className="col-date">
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                    <span>Date</span>
                    <input
                      type="date"
                      className="col-filter-input"
                      value={colFilters.date || ''}
                      onChange={(e) => {
                        setColFilters({ ...colFilters, date: e.target.value });
                        setCurrentPage(1);
                      }}
                      onClick={(e) => e.stopPropagation()}
                    />
                  </div>
                </th>
                <th rowSpan="2" className="col-particulars">
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                    <span>Particulars</span>
                    <select
                      className="col-filter-input"
                      value={colFilters.particulars || ''}
                      onChange={(e) => {
                        setColFilters({ ...colFilters, particulars: e.target.value });
                        setCurrentPage(1);
                      }}
                      onClick={(e) => e.stopPropagation()}
                    >
                      <option value="">All</option>
                      {uniqueParticulars.map(p => (
                        <option key={p} value={p}>{p}</option>
                      ))}
                    </select>
                  </div>
                </th>

                <th rowSpan="2" className="col-vchno">
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                    <span>Voucher No.</span>
                    <input
                      type="text"
                      className="col-filter-input"
                      placeholder="Filter No..."
                      value={colFilters.voucherNum || ''}
                      onChange={(e) => {
                        setColFilters({ ...colFilters, voucherNum: e.target.value });
                        setCurrentPage(1);
                      }}
                      onClick={(e) => e.stopPropagation()}
                    />
                  </div>
                </th>
                <th rowSpan="2" className="col-vchtype">
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                    <span>Voucher Type</span>
                    <select
                      className="col-filter-input"
                      value={colFilters.voucherType || ''}
                      onChange={(e) => {
                        setColFilters({ ...colFilters, voucherType: e.target.value });
                        setCurrentPage(1);
                      }}
                      onClick={(e) => e.stopPropagation()}
                    >
                      <option value="">All</option>
                      {uniqueVoucherTypes.map(t => (
                        <option key={t} value={t}>{t}</option>
                      ))}
                    </select>
                  </div>
                </th>
                <th rowSpan="2" className="col-value align-right">
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '4px', alignItems: 'flex-end' }}>
                    <span>Value</span>
                    <input
                      type="text"
                      className="col-filter-input"
                      placeholder="Filter Val..."
                      value={colFilters.value || ''}
                      onChange={(e) => {
                        setColFilters({ ...colFilters, value: e.target.value });
                        setCurrentPage(1);
                      }}
                      style={{ textAlign: 'right' }}
                      onClick={(e) => e.stopPropagation()}
                    />
                  </div>
                </th>
                <th rowSpan="2" className="col-gross align-right">
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '4px', alignItems: 'flex-end' }}>
                    <span>Gross Total</span>
                    <input
                      type="text"
                      className="col-filter-input"
                      placeholder="Filter Gross..."
                      value={colFilters.grossTotal || ''}
                      onChange={(e) => {
                        setColFilters({ ...colFilters, grossTotal: e.target.value });
                        setCurrentPage(1);
                      }}
                      style={{ textAlign: 'right' }}
                      onClick={(e) => e.stopPropagation()}
                    />
                  </div>
                </th>
                <th colSpan="6" className="col-ledgers-header align-center">Ledgers &amp; Taxes Breakdown</th>
                {activeUdfCols.length > 0 && (
                  <th colSpan={activeUdfCols.length} className="col-ledgers-header align-center" style={{ background: '#e9f0f6' }}>UDF Additional Charges</th>
                )}
              </tr>
              <tr>
                <th className="col-subled align-right">
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '4px', alignItems: 'flex-end' }}>
                    <span>Purchase ZCOL</span>
                    <input
                      type="text"
                      className="col-filter-input"
                      placeholder="Filter..."
                      value={colFilters.purchaseZcol || ''}
                      onChange={(e) => {
                        setColFilters({ ...colFilters, purchaseZcol: e.target.value });
                        setCurrentPage(1);
                      }}
                      style={{ textAlign: 'right' }}
                      onClick={(e) => e.stopPropagation()}
                    />
                  </div>
                </th>
                <th className="col-subled align-right">
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '4px', alignItems: 'flex-end' }}>
                    <span>IGST Tax</span>
                    <input
                      type="text"
                      className="col-filter-input"
                      placeholder="Filter..."
                      value={colFilters.igst || ''}
                      onChange={(e) => {
                        setColFilters({ ...colFilters, igst: e.target.value });
                        setCurrentPage(1);
                      }}
                      style={{ textAlign: 'right' }}
                      onClick={(e) => e.stopPropagation()}
                    />
                  </div>
                </th>
                <th className="col-subled align-right">
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '4px', alignItems: 'flex-end' }}>
                    <span>800012-Clearing Cess</span>
                    <input
                      type="text"
                      className="col-filter-input"
                      placeholder="Filter..."
                      value={colFilters.clearingCess || ''}
                      onChange={(e) => {
                        setColFilters({ ...colFilters, clearingCess: e.target.value });
                        setCurrentPage(1);
                      }}
                      style={{ textAlign: 'right' }}
                      onClick={(e) => e.stopPropagation()}
                    />
                  </div>
                </th>
                <th className="col-subled align-right">
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '4px', alignItems: 'flex-end' }}>
                    <span>CGST Tax</span>
                    <input
                      type="text"
                      className="col-filter-input"
                      placeholder="Filter..."
                      value={colFilters.cgst || ''}
                      onChange={(e) => {
                        setColFilters({ ...colFilters, cgst: e.target.value });
                        setCurrentPage(1);
                      }}
                      style={{ textAlign: 'right' }}
                      onClick={(e) => e.stopPropagation()}
                    />
                  </div>
                </th>
                <th className="col-subled align-right">
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '4px', alignItems: 'flex-end' }}>
                    <span>SGST Tax</span>
                    <input
                      type="text"
                      className="col-filter-input"
                      placeholder="Filter..."
                      value={colFilters.sgst || ''}
                      onChange={(e) => {
                        setColFilters({ ...colFilters, sgst: e.target.value });
                        setCurrentPage(1);
                      }}
                      style={{ textAlign: 'right' }}
                      onClick={(e) => e.stopPropagation()}
                    />
                  </div>
                </th>
                <th className="col-subled align-right">
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '4px', alignItems: 'flex-end' }}>
                    <span>Clean Engry Cess</span>
                    <input
                      type="text"
                      className="col-filter-input"
                      placeholder="Filter..."
                      value={colFilters.cleanEnergy || ''}
                      onChange={(e) => {
                        setColFilters({ ...colFilters, cleanEnergy: e.target.value });
                        setCurrentPage(1);
                      }}
                      style={{ textAlign: 'right' }}
                      onClick={(e) => e.stopPropagation()}
                    />
                  </div>
                </th>

                {activeUdfCols.map(udf => (
                  <th key={udf.name} className="col-subled align-right">
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '4px', alignItems: 'flex-end' }}>
                      <span>{udf.name}</span>
                      <input
                        type="text"
                        className="col-filter-input"
                        placeholder="Filter..."
                        value={colFilters[udf.name] || ''}
                        onChange={(e) => {
                          setColFilters({ ...colFilters, [udf.name]: e.target.value });
                          setCurrentPage(1);
                        }}
                        style={{ textAlign: 'right' }}
                        onClick={(e) => e.stopPropagation()}
                      />
                    </div>
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {currentRows.map((v, pageIdx) => {
                const globalIdx = indexOfFirstRow + pageIdx;
                const isSelected = selectedVoucher === v;
                const curTotalItemVal = v.items.reduce((sum, item) => sum + item.amount, 0);
                const curPartyLedger = v.ledgers.find(l => l.isParty);
                const curTotalAmount = curPartyLedger ? Math.abs(curPartyLedger.amount) : 0;

                return (
                  <tr
                    key={globalIdx}
                    className={isSelected ? "selected-row" : ""}
                    onClick={() => setSelectedVoucherIndex(globalIdx)}
                    style={{ cursor: 'pointer' }}
                  >
                    <td className="cell-date">{v.date}</td>
                    <td className="cell-particulars font-highlight">{v.partyName || '—'}</td>
                    <td className="cell-vchno">{v.voucherNum || '—'}</td>
                    <td className="cell-vchtype">{v.voucherType || '—'}</td>
                    <td className="cell-value align-right">{formatValue(curTotalItemVal)}</td>
                    <td className="cell-gross align-right font-bold">{formatValue(curTotalAmount)} Cr</td>

                    <td className="cell-led align-right">{formatValue(getLedgerAmt(v, 'Purchase ZCOL'), true)}</td>
                    <td className="cell-led align-right">{formatValue(getLedgerAmt(v, 'IGST'), true)}</td>
                    <td className="cell-led align-right">{formatValue(getLedgerAmt(v, 'Clean Energy Cess') || getLedgerAmt(v, '800012'), true)}</td>
                    <td className="cell-led align-right">{formatValue(getLedgerAmt(v, 'CGST'), true)}</td>
                    <td className="cell-led align-right">{formatValue(getLedgerAmt(v, 'SGST'), true)}</td>
                    <td className="cell-led align-right">{formatValue(getLedgerAmt(v, 'Clean Engry Cess') || getLedgerAmt(v, 'Clean Entry Cess'), true)}</td>

                    {activeUdfCols.map(udf => {
                      const u = v.udfByName?.[udf.name];
                      return (
                        <td key={udf.name} className="cell-led align-right">
                          {u && u.amount ? formatValue(u.amount, true) : ''}
                        </td>
                      );
                    })}
                  </tr>
                );
              })}
              {currentRows.length === 0 && (
                <tr>
                  <td colSpan={12 + activeUdfCols.length} style={{ textAlign: 'center', padding: '2rem', color: 'var(--text-2)' }}>
                    No vouchers match your filter.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>

      {selectedVoucher && (
        <div style={{ border: '1px solid var(--border)', borderRadius: 'var(--radius)', background: 'var(--surface)', overflow: 'hidden' }}>
          <div style={{ background: 'var(--surface-2)', borderBottom: '1px solid var(--border)', padding: '0.75rem 1.25rem', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <h3 style={{ margin: 0, fontSize: '1rem', color: 'var(--text-1)' }}>
              Voucher Details: PO #{selectedVoucher.voucherNum}
            </h3>
            <span style={{ fontSize: '0.85rem', color: 'var(--text-2)' }}>
              Type: <strong>{selectedVoucher.voucherType}</strong>
            </span>
          </div>

          {/* 1. VOUCHER HEADER DETAILS TABLE */}
          <div className="tab-content" style={{ padding: '1.25rem' }}>
            <h3 style={{ marginBottom: '1rem', display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="18" height="18" style={{ color: 'var(--accent)' }}>
                <rect x="3" y="4" width="18" height="18" rx="2" ry="2" />
                <line x1="16" y1="2" x2="16" y2="6" />
                <line x1="8" y1="2" x2="8" y2="6" />
                <line x1="3" y1="10" x2="21" y2="10" />
              </svg>
              Voucher Information
            </h3>
            <div className="table-scroll">
              <table className="results-table">
                <thead>
                  <tr>
                    <th>Field Name</th>
                    <th>Value</th>
                    <th>Field Name</th>
                    <th>Value</th>
                  </tr>
                </thead>
                <tbody>
                  <tr>
                    <td><strong>Voucher Number</strong></td>
                    <td><code className="mat-code">#{selectedVoucher.voucherNum}</code></td>
                    <td><strong>Voucher Type</strong></td>
                    <td>{selectedVoucher.voucherType}</td>
                  </tr>
                  <tr>
                    <td><strong>Voucher Date</strong></td>
                    <td>{selectedVoucher.date}</td>
                    <td><strong>Tally GUID</strong></td>
                    <td><code className="mat-code">{selectedVoucher.guid}</code></td>
                  </tr>
                  <tr>
                    <td><strong>Master ID</strong></td>
                    <td>{selectedVoucher.masterId}</td>
                    <td><strong>Alter ID</strong></td>
                    <td>{selectedVoucher.alterId}</td>
                  </tr>
                  <tr>
                    <td><strong>Supplier Name</strong></td>
                    <td colSpan="3"><strong>{selectedVoucher.partyName}</strong></td>
                  </tr>
                  <tr>
                    <td><strong>State Name</strong></td>
                    <td>{selectedVoucher.stateName}</td>
                    <td><strong>Place of Supply</strong></td>
                    <td>{selectedVoucher.placeOfSupply}</td>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>

          {/* 2. INVENTORY ENTRIES TABLE */}
          <div className="tab-content" style={{ padding: '1.25rem' }}>
            <h3 style={{ marginBottom: '1rem', display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="18" height="18" style={{ color: 'var(--green)' }}>
                <path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z" />
                <polyline points="3.27 6.96 12 12.01 20.73 6.96" />
                <line x1="12" y1="22.08" x2="12" y2="12" />
              </svg>
              Inventory Stock Items
            </h3>
            <div className="table-scroll">
              <table className="results-table">
                <thead>
                  <tr>
                    <th>Stock Item Name</th>
                    <th>Billed Qty</th>
                    <th>Actual Qty</th>
                    <th style={{ textAlign: 'right' }}>Rate</th>
                    <th style={{ textAlign: 'right' }}>Amount</th>
                  </tr>
                </thead>
                <tbody>
                  {selectedVoucher.items.map((item, idx) => (
                    <tr key={idx}>
                      <td><strong>{item.name}</strong></td>
                      <td>{item.billedQty}</td>
                      <td>{item.actualQty}</td>
                      <td style={{ textAlign: 'right' }} className="code-val">{item.rate}</td>
                      <td style={{ textAlign: 'right' }} className="num-cell" style={{ fontWeight: '600' }}>
                        ₹ {formatCurrency(item.amount)}
                      </td>
                    </tr>
                  ))}
                  {selectedVoucher.items.length === 0 && (
                    <tr>
                      <td colSpan="5" style={{ textAlign: 'center', color: 'var(--text-2)' }}>No stock items found in this voucher.</td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>

          {/* 3. LEDGER ACCOUNTING TABLE */}
          <div className="tab-content" style={{ padding: '1.25rem' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1rem' }}>
              <h3 style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', margin: 0 }}>
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="18" height="18" style={{ color: 'var(--amber)' }}>
                  <line x1="12" y1="1" x2="12" y2="23" />
                  <path d="M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"/>
                </svg>
                Financial Ledgers &amp; Taxes
              </h3>
              <div style={{ background: 'var(--surface-2)', border: '1px solid var(--border)', borderRadius: 'var(--radius)', padding: '0.5rem 1rem', display: 'flex', gap: '1.5rem', alignItems: 'center' }}>
                <span style={{ fontSize: '0.85rem', color: 'var(--text-2)' }}>Total Voucher Value:</span>
                <strong style={{ fontSize: '1.1rem', color: 'var(--accent)' }}>₹ {formatCurrency(totalAmount)}</strong>
              </div>
            </div>
            <div className="table-scroll">
              <table className="results-table">
                <thead>
                  <tr>
                    <th>Ledger Name</th>
                    <th>Ledger Type</th>
                    <th>Deemed Positive (Debit)</th>
                    <th style={{ textAlign: 'right' }}>Amount</th>
                  </tr>
                </thead>
                <tbody>
                  {selectedVoucher.ledgers.map((ledger, idx) => {
                    const absAmt = Math.abs(ledger.amount);
                    const isCredit = ledger.amount > 0;
                    return (
                      <tr key={idx} className={ledger.isParty ? 'party-row' : ''}>
                        <td>
                          <strong>{ledger.name}</strong>
                          {ledger.isParty && <span className="row-badge">Party Ledger</span>}
                        </td>
                        <td>{ledger.isParty ? 'Sundry Creditor (Party)' : 'Tax / Purchase Account'}</td>
                        <td>
                          <span className={`status-badge ${ledger.deemedPositive ? 'badge-success' : 'badge-failed'}`}>
                            {ledger.deemedPositive ? 'Yes' : 'No'}
                          </span>
                        </td>
                        <td style={{ textAlign: 'right' }} className="num-cell">
                          <span style={{ fontWeight: ledger.isParty ? '700' : '500' }}>
                            ₹ {formatCurrency(absAmt)}
                          </span>
                          <span className="ledger-dr-cr">
                            {isCredit ? ' Cr' : ' Dr'}
                          </span>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>

          {/* 4. UDF ADDITIONAL CHARGES TABLE */}
          {selectedVoucher.udfs && selectedVoucher.udfs.length > 0 && (
            <div className="tab-content" style={{ padding: '1.25rem' }}>
              <h3 style={{ marginBottom: '1rem', display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="18" height="18" style={{ color: 'var(--accent)' }}>
                  <path d="M12 2L2 7l10 5 10-5-10-5zM2 17l10 5 10-5M2 12l10 5 10-5" />
                </svg>
                UDF Additional Charges Breakdown
              </h3>
              <div className="table-scroll">
                <table className="results-table">
                  <thead>
                    <tr>
                      <th>UDF Field Key</th>
                      <th>Charge Name / Description</th>
                      <th>Sub Ledger Code</th>
                      <th style={{ textAlign: 'right' }}>Amount</th>
                    </tr>
                  </thead>
                  <tbody>
                    {selectedVoucher.udfs.map((udf, idx) => (
                      <tr key={idx}>
                        <td><code className="mat-code">{udf.keyName}</code></td>
                        <td><strong>{udf.name}</strong></td>
                        <td>{udf.ledgerCode || '—'}</td>
                        <td style={{ textAlign: 'right' }} className="num-cell" style={{ fontWeight: '600' }}>
                          ₹ {formatCurrency(udf.amount)}
                          <span className="ledger-dr-cr"> Dr</span>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </div>
      )}

      {/* 5. COLLAPSIBLE RAW XML VIEW */}
      {showXml && (
        <div className="tab-content raw-xml-view" style={{ marginTop: '1.5rem' }}>
          <div className="xml-header">
            <h3>Tally XML Document payload</h3>
            <button className="btn-sm-outline" onClick={copyToClipboard}>
              {copySuccess ? 'Copied!' : 'Copy Code'}
            </button>
          </div>
          <pre className="xml-code">
            <code>{xmlText}</code>
          </pre>
        </div>
      )}

      {/* TALLY LIVE DATE RANGE POPUP MODAL */}
      {showTallyModal && (
        <div className="tally-modal-overlay">
          <div className="tally-modal-content">
            <div className="tally-modal-header" style={{ textAlign: 'center', marginBottom: '1.5rem' }}>
              <div style={{ background: 'var(--accent-light)', color: 'var(--accent)', width: '48px', height: '48px', borderRadius: '50%', display: 'flex', alignItems: 'center', justifyContent: 'center', margin: '0 auto 1rem auto' }}>
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" width="22" height="22">
                  <rect x="3" y="4" width="18" height="18" rx="2" ry="2" />
                  <line x1="16" y1="2" x2="16" y2="6" />
                  <line x1="8" y1="2" x2="8" y2="6" />
                  <line x1="3" y1="10" x2="21" y2="10" />
                </svg>
              </div>
              <h2 className="tally-modal-title" style={{ fontSize: '1.3rem', fontWeight: '700' }}>Fetch Live Report from Tally</h2>
              <p className="tally-modal-desc" style={{ marginTop: '0.35rem' }}>Choose the date range to query the Voucher Register</p>
            </div>

            <div style={{ display: 'flex', gap: '1rem', marginBottom: '1rem' }}>
              <div className="tally-modal-form-group" style={{ flex: 1 }}>
                <label>From Date:</label>
                <input
                  type="date"
                  value={fromDate}
                  onChange={(e) => setFromDate(e.target.value)}
                />
              </div>

              <div className="tally-modal-form-group" style={{ flex: 1 }}>
                <label>To Date:</label>
                <input
                  type="date"
                  value={toDate}
                  onChange={(e) => setToDate(e.target.value)}
                />
              </div>
            </div>

            <div className="tally-modal-presets">
              <span className="preset-label">Quick Date Presets</span>
              <div className="preset-buttons">
                <button
                  type="button"
                  className={`preset-btn ${fromDate === '2023-04-01' && toDate === '2024-03-31' ? 'active' : ''}`}
                  onClick={() => { setFromDate('2023-04-01'); setToDate('2024-03-31'); }}
                >
                  FY 2023-24
                </button>
                <button
                  type="button"
                  className={`preset-btn ${fromDate === '2022-04-01' && toDate === '2023-03-31' ? 'active' : ''}`}
                  onClick={() => { setFromDate('2022-04-01'); setToDate('2023-03-31'); }}
                >
                  FY 2022-23
                </button>
                <button
                  type="button"
                  className={`preset-btn ${fromDate === '2023-01-01' && toDate === '2023-12-31' ? 'active' : ''}`}
                  onClick={() => { setFromDate('2023-01-01'); setToDate('2023-12-31'); }}
                >
                  CY 2023
                </button>
              </div>
            </div>

            <div className="tally-modal-actions">
              <button
                className="btn-secondary"
                onClick={() => {
                  setShowTallyModal(false);
                  if (vouchers.length === 0 && xmlFiles.length > 0) {
                    fetchXml(selectedFile);
                  }
                }}
              >
                Cancel / View Local
              </button>
              <button
                className="btn-primary"
                onClick={() => fetchFromTally(fromDate, toDate)}
                style={{ background: '#3b82f6', borderColor: '#3b82f6' }}
              >
                Fetch Report
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

export default XmlVisualizer;
