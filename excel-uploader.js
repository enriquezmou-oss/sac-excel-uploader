if (typeof XLSX === 'undefined') {
  var script = document.createElement('script');
  script.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
  script.async = false;
  document.head.appendChild(script);
}

class ExcelUploaderPro extends HTMLElement {
  constructor() {
    super();
    this._shadow = this.attachShadow({ mode: 'open' });
    this._step = 1;
    this._filesQueue = [];
    this._rows = [];
    this._headers = [];
    this._currentUser = 'SAP User (Planning)';
    this._isLoading = false;
    this._loadingText = 'Processing files...';
    
    const savedCols = localStorage.getItem('sac_required_cols');
    this._requiredColumns = savedCols ? JSON.parse(savedCols) : ['ID', 'DESCRIPTION', 'H1', 'COMPANY', 'COSTCENTER', 'ASSET_CLASS', 'CAPITALIZED'];
    
    this._masterData = {
      COMPANY: ['1000', '2000', '950'],
      COSTCENTER: ['CC_IT_01', 'CC_FIN_02', '12234566', '10180000']
    };

    this._history = [];
    this._issueFilter = 'all';
    this._render();
  }

  connectedCallback() {
    this._bindEvents();
  }

  _bindEvents() {
    const shadow = this._shadow;

    const fileInput = shadow.getElementById('fileInput');
    const chooseBtn = shadow.getElementById('chooseFileBtn');
    if (chooseBtn && fileInput) {
      chooseBtn.addEventListener('click', () => fileInput.click());
      fileInput.addEventListener('change', (e) => {
        if (e.target.files.length) {
          this._handleFilesSelection(e.target.files);
        }
      });
    }

    const nextStep2Btn = shadow.getElementById('nextStep2Btn');
    const prevBtn = shadow.getElementById('prevStepBtn');
    const finishBtn = shadow.getElementById('finishBtn');
    const exitBtn = shadow.getElementById('exitBtn');
    const downloadLogBtn = shadow.getElementById('downloadLogBtn');

    if (nextStep2Btn) nextStep2Btn.addEventListener('click', () => { this._step = 3; this._render(); });
    if (prevBtn) prevBtn.addEventListener('click', () => { if (this._step > 1) { this._step--; this._render(); } });
    if (exitBtn) exitBtn.addEventListener('click', () => { this._step = 1; this._filesQueue = []; this._rows = []; this._render(); });

    if (downloadLogBtn) {
      downloadLogBtn.addEventListener('click', () => this._exportErrorLog());
    }

    shadow.querySelectorAll('.nav-tab').forEach(tab => {
      tab.addEventListener('click', (e) => {
        shadow.querySelectorAll('.nav-tab').forEach(t => t.classList.remove('active'));
        shadow.querySelectorAll('.tab-content').forEach(c => c.style.display = 'none');
        e.target.classList.add('active');
        shadow.getElementById(e.target.dataset.target).style.display = 'block';
      });
    });

    shadow.querySelectorAll('.chip-filter').forEach(chip => {
      chip.addEventListener('click', (e) => {
        this._issueFilter = e.target.dataset.filter;
        this._render();
      });
    });

    shadow.querySelectorAll('.editable-cell').forEach(cell => {
      cell.addEventListener('blur', (e) => {
        const rowIdx = parseInt(e.target.dataset.row, 10);
        const colName = e.target.dataset.col;
        const newVal = e.target.innerText.trim();
        if (this._rows[rowIdx]) {
          this._rows[rowIdx][colName] = newVal;
          this._validateRows();
          this._render();
        }
      });
    });

    const addColBtn = shadow.getElementById('addColumnBtn');
    if (addColBtn) {
      addColBtn.addEventListener('click', () => {
        const colName = prompt('Enter new mandatory field name:');
        if (colName && !this._requiredColumns.includes(colName.toUpperCase())) {
          this._requiredColumns.push(colName.toUpperCase());
          localStorage.setItem('sac_required_cols', JSON.stringify(this._requiredColumns));
          this._validateRows();
          this._render();
        }
      });
    }

    if (finishBtn) {
      finishBtn.addEventListener('click', () => {
        const errorsCount = this._rows.filter(r => Object.keys(r._errors).length > 0).length;
        if (errorsCount > 0) {
          alert('⚠️ Cannot finish. There are cells with errors that must be corrected.');
          return;
        }

        const fileNamesList = this._filesQueue.map(f => f.name).join(', ');
        this._history.unshift({
          id: 'LOAD-' + Math.floor(Math.random() * 90000 + 10000),
          user: this._currentUser,
          file: fileNamesList,
          rows: this._rows.length,
          timestamp: new Date().toLocaleString(),
          status: 'Successful'
        });

        const cleanJson = JSON.stringify(this._rows.map(r => {
          const clean = { ...r };
          delete clean._rowNum;
          delete clean._errors;
          delete clean._sourceFile;
          return clean;
        }));

        this.dispatchEvent(new CustomEvent('propertiesChanged', {
          bubbles: true, composed: true,
          detail: { properties: { validatedDataJson: cleanJson } }
        }));

        alert('🚀 Files successfully consolidated and synchronized with SAP Analytics Cloud!');
        this._step = 1;
        this._filesQueue = [];
        this._rows = [];
        this._render();
      });
    }
  }

  _handleFilesSelection(fileList) {
    if (typeof XLSX === 'undefined') return;
    this._isLoading = true;
    this._loadingText = 'Processing and unifying Excel files...';
    this._render();

    for (let i = 0; i < fileList.length; i++) {
      this._filesQueue.push(fileList[i]);
    }

    this._processAllFilesInQueue();
  }

  _processAllFilesInQueue() {
    let processedCount = 0;
    let allCombinedRows = [];
    let detectedHeaders = [];

    this._filesQueue.forEach((file, fileIndex) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const wb = XLSX.read(data, { type: 'array' });
          const ws = wb.Sheets[wb.SheetNames[0]];
          const json = XLSX.utils.sheet_to_json(ws, { header: 1 });

          if (json.length > 0) {
            if (fileIndex === 0) {
              detectedHeaders = json[0].map(h => String(h).trim());
            }
            const rows = json.slice(1).map((row, idx) => {
              const obj = { _rowNum: allCombinedRows.length + idx + 2, _sourceFile: file.name, _errors: {} };
              detectedHeaders.forEach((h, i) => {
                obj[h] = row[i] !== undefined ? row[i] : '';
              });
              return obj;
            }).filter(r => Object.values(r).some(v => v !== '' && v !== '_rowNum' && v !== '_sourceFile' && v !== '_errors'));

            allCombinedRows = allCombinedRows.concat(rows);
          }
        } catch (err) {
          console.error(`Error processing ${file.name}:`, err);
        }

        processedCount++;
        if (processedCount === this._filesQueue.length) {
          this._headers = detectedHeaders;
          this._rows = allCombinedRows;
          this._validateRows();
          this._isLoading = false;
          this._step = 2;
          this._render();
        }
      };
      reader.readAsArrayBuffer(file);
    });
  }

  _validateRows() {
    const idCounts = {};
    this._rows.forEach(r => {
      const id = String(r.ID || '').trim();
      if (id) idCounts[id] = (idCounts[id] || 0) + 1;
    });

    this._rows.forEach(row => {
      row._errors = {};
      
      this._requiredColumns.forEach(col => {
        if (!row[col] || String(row[col]).trim() === '') {
          row._errors[col] = 'Blank cell not allowed';
        }
      });

      const id = String(row.ID || '').trim();
      if (id && idCounts[id] > 1) {
        row._errors['ID'] = 'Duplicate value not allowed';
      }

      const comp = String(row.COMPANY || '').trim();
      if (comp && !this._masterData.COMPANY.includes(comp)) {
        row._errors['COMPANY'] = `Value ${comp} does not exist in master data (Allowed: ${this._masterData.COMPANY.join(', ')})`;
      }

      const cc = String(row.COSTCENTER || '').trim();
      if (cc && !this._masterData.COSTCENTER.includes(cc)) {
        row._errors['COSTCENTER'] = `Value ${cc} does not exist in master data (Allowed: ${this._masterData.COSTCENTER.join(', ')})`;
      }
    });
  }

  _exportErrorLog() {
    const errorRows = this._rows.filter(r => Object.keys(r._errors).length > 0);
    if (errorRows.length === 0) {
      alert('No errors to export.');
      return;
    }
    let csv = 'SourceFile,Row,ID,Column,Error Description\n';
    errorRows.forEach(r => {
      Object.keys(r._errors).forEach(col => {
        csv += `${r._sourceFile},${r._rowNum},${r.ID || ''},${col},"${r._errors[col]}"\n`;
      });
    });
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'MultiFile_Validation_Error_Log.csv';
    a.click();
  }

  _render() {
    const totalRows = this._rows.length;
    const errorRows = this._rows.filter(r => Object.keys(r._errors).length > 0).length;
    const validRows = totalRows - errorRows;
    const validPercent = totalRows > 0 ? Math.round((validRows / totalRows) * 100) : 100;

    this._shadow.innerHTML = `
      <style>
        :host { display: block; font-family: "72", "72full", Arial, Helvetica, sans-serif; font-size: 13px; color: #32363A; background: #F5F6F7; padding: 16px; box-sizing: border-box; height: 100%; overflow: auto; }
        .main-card { background: #FFFFFF; border: 1px solid #D1D7DD; border-radius: 8px; box-shadow: 0 4px 12px rgba(0,0,0,0.08); overflow: hidden; position: relative; }
        
        .busy-overlay { position: absolute; top: 0; left: 0; right: 0; bottom: 0; background: rgba(255,255,255,0.85); display: flex; flex-direction: column; align-items: center; justify-content: center; z-index: 100; }
        .spinner { width: 40px; height: 40px; border: 4px solid #E5E9EC; border-top-color: #0070F2; border-radius: 50%; animation: spin 0.8s linear infinite; margin-bottom: 10px; }
        @keyframes spin { to { transform: rotate(360deg); } }

        .sap-header { display: flex; align-items: center; justify-content: space-between; background: #354A5F; color: #FFFFFF; padding: 12px 20px; }
        .sap-brand { display: flex; align-items: center; gap: 10px; }
        .sap-logo { width: 28px; height: 28px; background: #0070F2; border-radius: 4px; display: flex; align-items: center; justify-content: center; font-weight: 800; color: #FFF; font-size: 14px; letter-spacing: -1px; box-shadow: inset 0 0 0 2px rgba(255,255,255,0.4); }
        .sap-title { font-size: 15px; font-weight: 700; }

        .nav-tabs { display: flex; background: #FFFFFF; border-bottom: 1px solid #D1D7DD; padding: 0 16px; }
        .nav-tab { padding: 10px 18px; cursor: pointer; font-weight: 600; color: #6A6D70; border-bottom: 3px solid transparent; font-size: 13px; }
        .nav-tab.active { color: #0070F2; border-bottom-color: #0070F2; background: rgba(0,112,242,0.04); }

        .tab-content { padding: 20px; background: #F5F6F7; }

        .wizard-steps { display: flex; justify-content: space-between; background: #FFF; border: 1px solid #D1D7DD; border-radius: 6px; padding: 12px 20px; margin-bottom: 16px; }
        .step-item { display: flex; align-items: center; gap: 8px; font-weight: 600; color: #89919A; font-size: 12px; }
        .step-item.active { color: #0070F2; }
        .step-num { width: 22px; height: 22px; border-radius: 50%; background: #E5E9EC; color: #6A6D70; display: flex; align-items: center; justify-content: center; font-size: 11px; }
        .step-item.active .step-num { background: #0070F2; color: #FFF; }

        .btn { padding: 7px 16px; font-size: 12px; font-weight: 600; border-radius: 4px; cursor: pointer; border: 1px solid #0070F2; background: #0070F2; color: #FFF; transition: background 0.2s; }
        .btn:hover { background: #005CB9; }
        .btn-ghost { background: #FFFFFF; color: #0070F2; border-color: #0070F2; }
        .btn-ghost:hover { background: rgba(0,112,242,0.05); }
        .btn-danger { background: #FFF; color: #BB0000; border-color: #BB0000; }
        .btn-danger:hover { background: #FEEAEA; }

        .drop-zone { border: 2px dashed #0070F2; border-radius: 8px; padding: 40px; text-align: center; background: #FFFFFF; cursor: pointer; }
        .drop-zone:hover { background: rgba(0,112,242,0.02); }

        .metrics-bar { display: grid; grid-template-columns: repeat(5, 1fr); gap: 12px; margin-bottom: 12px; }
        .metric-box { background: #FFFFFF; border: 1px solid #D1D7DD; padding: 10px 14px; border-radius: 6px; }
        .metric-title { font-size: 11px; color: #6A6D70; margin-bottom: 2px; }
        .metric-val { font-size: 15px; font-weight: 700; color: #1D2D3E; }

        .dist-bar { display: flex; height: 6px; border-radius: 3px; overflow: hidden; background: #FEEAEA; margin-bottom: 16px; border: 1px solid #D1D7DD; }
        .dist-ok { background: #107E3E; height: 100%; }

        .chips-row { display: flex; gap: 8px; margin-bottom: 12px; align-items: center; justify-content: space-between; }
        .chips-left { display: flex; gap: 8px; align-items: center; }
        .chip { padding: 4px 10px; border-radius: 12px; font-size: 11px; font-weight: 600; cursor: pointer; border: 1px solid #D1D7DD; background: #FFF; color: #32363A; }
        .chip-all.active-chip { background: #354A5F; color: #FFF; border-color: #354A5F; }
        .chip-ok.active-chip { background: #107E3E; color: #FFF; border-color: #107E3E; }
        .chip-err.active-chip { background: #BB0000; color: #FFF; border-color: #BB0000; }

        .tbl-wrap { max-height: 280px; overflow: auto; border: 1px solid #D1D7DD; border-radius: 6px; background: #FFFFFF; }
        table { width: 100%; border-collapse: collapse; font-size: 12px; text-align: left; }
        th { background: #F5F6F7; padding: 8px 10px; border-bottom: 1px solid #D1D7DD; position: sticky; top: 0; font-weight: 700; color: #32363A; }
        td { padding: 8px 10px; border-bottom: 1px solid #EDF0F2; }
        td.cell-error { background: #FEEAEA; color: #BB0000; }
        .editable-cell { outline: none; border-bottom: 1px dashed #0070F2; display: inline-block; min-width: 60px; }
        .editable-cell:focus { background: #FFF; border: 1px solid #0070F2; box-shadow: 0 0 0 2px rgba(0,112,242,0.2); }

        .wizard-footer { display: flex; justify-content: space-between; margin-top: 20px; border-top: 1px solid #D1D7DD; padding-top: 12px; }
        .file-pill { display: inline-flex; align-items: center; background: #E6F0FF; color: #0064D9; padding: 4px 10px; border-radius: 12px; font-size: 11px; font-weight: 600; margin: 2px; }
      </style>

      <div class="main-card">
        ${this._isLoading ? `
          <div class="busy-overlay">
            <div class="spinner"></div>
            <div style="font-weight:600; color:#354A5F;">${this._loadingText}</div>
          </div>
        ` : ''}

        <div class="sap-header">
          <div class="sap-brand">
            <div class="sap-logo">SAP</div>
            <div class="sap-title">Excel Planning Uploader Pro (Multi-File)</div>
          </div>
          <div style="font-size: 11px; font-weight: 600; opacity: 0.85;">SAP Analytics Cloud BTP</div>
        </div>

        <div class="nav-tabs">
          <div class="nav-tab active" data-target="uploaderTab">Upload Wizard</div>
          <div class="nav-tab" data-target="monitorTab">Audit Monitor (${this._history.length})</div>
        </div>

        <div id="uploaderTab" class="tab-content">
          <div class="wizard-steps">
            <div class="step-item ${this._step === 1 ? 'active' : ''}"><div class="step-num">1</div> Multi-File Upload</div>
            <div class="step-item ${this._step === 2 ? 'active' : ''}"><div class="step-num">2</div> Consolidation & Mapping</div>
            <div class="step-item ${this._step === 3 ? 'active' : ''}"><div class="step-num">3</div> Global Validation</div>
          </div>

          ${this._step === 1 ? `
            <div class="drop-zone" id="chooseFileBtn">
              <h3 style="margin:0 0 6px 0; color:#0070F2;">Drag and drop one or more planning Excel files</h3>
              <p style="margin:0; color:#6A6D70; font-size:12px;">You can select multiple files simultaneously (.xlsx, .xls)</p>
              <button class="btn" style="margin-top:14px;" type="button">Select Files</button>
            </div>
            <input type="file" id="fileInput" style="display:none" multiple accept=".xlsx, .xls, .xlsm" />
          ` : ''}

          ${this._step === 2 ? `
            <div>
              <h3 style="margin-top:0; color:#1D2D3E;">Multiple Files Consolidation</h3>
              <p style="color:#6A6D70;">Files loaded in queue (${this._filesQueue.length}):</p>
              <div style="margin-bottom:16px;">
                ${this._filesQueue.map(f => `<span class="file-pill">📄 ${f.name}</span>`).join('')}
              </div>

              <p>Mandatory fields required by the unified model:</p>
              <div style="display:flex; flex-wrap:wrap; gap:6px; margin-bottom:16px;">
                ${this._requiredColumns.map(c => `<span style="background:#E6F0FF; color:#0064D9; padding:4px 8px; border-radius:4px; font-weight:600; font-size:11px;">${c}</span>`).join('')}
              </div>
              <button class="btn btn-ghost" id="addColumnBtn" type="button">+ Add new mandatory field</button>
            </div>
            <div class="wizard-footer">
              <button class="btn btn-ghost" id="prevStepBtn" type="button">Back</button>
              <button class="btn" id="nextStep2Btn" type="button">Validate ${totalRows} unified rows</button>
            </div>
          ` : ''}

          ${this._step === 3 ? `
            <div>
              <div class="metrics-bar">
                <div class="metric-box"><div class="metric-title">Files</div><div class="metric-val">${this._filesQueue.length}</div></div>
                <div class="metric-box"><div class="metric-title">Total Rows</div><div class="metric-val">${totalRows}</div></div>
                <div class="metric-box"><div class="metric-title">Correct</div><div class="metric-val" style="color:#107E3E;">${validRows}</div></div>
                <div class="metric-box"><div class="metric-title">With Errors</div><div class="metric-val" style="color:#BB0000;">${errorRows}</div></div>
                <div class="metric-box"><div class="metric-title">Global Status</div><div class="metric-val" style="font-size:12px; color:${errorRows > 0 ? '#BB0000':'#107E3E'}">${errorRows > 0 ? '⚠️ Review Errors' : '✅ Valid'}</div></div>
              </div>

              <div class="dist-bar" title="Data quality: ${validPercent}% valid">
                <div class="dist-ok" style="width: ${validPercent}%;"></div>
              </div>

              <div class="chips-row">
                <div class="chips-left">
                  <span style="font-weight:600; font-size:12px;">Filter cells:</span>
                  <div class="chip chip-all ${this._issueFilter === 'all' ? 'active-chip':''}" data-filter="all">All (${totalRows})</div>
                  <div class="chip chip-ok ${this._issueFilter === 'correct' ? 'active-chip':''}" data-filter="correct">Correct (${validRows})</div>
                  <div class="chip chip-err ${this._issueFilter === 'errors' ? 'active-chip':''}" data-filter="errors">With Errors (${errorRows})</div>
                </div>
                <div>
                  ${errorRows > 0 ? `<button class="btn btn-danger" id="downloadLogBtn" type="button" style="font-size:11px; padding:5px 12px;">📥 Download Error Log (CSV)</button>` : ''}
                </div>
              </div>

              <div class="tbl-wrap">
                <table>
                  <thead>
                    <tr>
                      <th style="width:30px;">#</th>
                      <th>Source File</th>
                      ${this._requiredColumns.map(h => `<th>${h}</th>`).join('')}
                    </tr>
                  </thead>
                  <tbody>
                    ${this._rows.filter(r => {
                      const hasErr = Object.keys(r._errors).length > 0;
                      if (this._issueFilter === 'correct' && hasErr) return false;
                      if (this._issueFilter === 'errors' && !hasErr) return false;
                      return true;
                    }).map(r => {
                      return `
                        <tr>
                          <td style="color:#89919A; font-size:11px;">${r._rowNum}</td>
                          <td style="font-size:11px; color:#0064D9; font-weight:600;">${r._sourceFile}</td>
                          ${this._requiredColumns.map(col => {
                            const err = r._errors[col];
                            return `
                              <td class="${err ? 'cell-error' : ''}" title="${err || 'Correct'}">
                                <span class="editable-cell" contenteditable="true" data-row="${r._rowNum - 2}" data-col="${col}">${r[col] !== undefined ? r[col] : ''}</span>
                                ${err ? `<span style="display:block; font-size:10px; color:#BB0000; font-weight:700; margin-top:2px;">⚠️ ${err}</span>` : ''}
                              </td>
                            `;
                          }).join('')}
                        </tr>
                      `;
                    }).join('')}
                  </tbody>
                </table>
              </div>
            </div>
            <div class="wizard-footer">
              <button class="btn btn-ghost" id="prevStepBtn" type="button">Back</button>
              <div style="display:flex; gap:8px;">
                <button class="btn btn-ghost" id="exitBtn" type="button">Cancel</button>
                <button class="btn" id="finishBtn" type="button" ${errorRows > 0 ? 'style="opacity:0.6; cursor:not-allowed;"':''}>Consolidate and Load into SAC</button>
              </div>
            </div>
          ` : ''}
        </div>

        <div id="monitorTab" class="tab-content" style="display:none;">
          <h3 style="margin-top:0; color:#1D2D3E;">BTP Consolidation & Audit Monitor</h3>
          <p style="color:#6A6D70; font-size:12px;">Corporate traceability of executed multi-file planning loads.</p>
          
          <div style="border: 1px solid #D1D7DD; border-radius: 6px; overflow:hidden; background:#FFF;">
            <table>
              <thead>
                <tr>
                  <th>Load ID</th>
                  <th>SAP User</th>
                  <th>Source Files</th>
                  <th>Total Rows</th>
                  <th>Timestamp</th>
                  <th>BTP Status</th>
                </tr>
              </thead>
              <tbody>
                ${this._history.length > 0 ? this._history.map(h => `
                  <tr>
                    <td><strong>${h.id}</strong></td>
                    <td>${h.user}</td>
                    <td>${h.file}</td>
                    <td>${h.rows} rows</td>
                    <td>${h.timestamp}</td>
                    <td><span style="background:#E6F9EC; color:#107E3E; padding:2px 8px; border-radius:10px; font-weight:600; font-size:11px;">${h.status}</span></td>
                  </tr>
                `).join('') : `
                  <tr><td colspan="6" style="text-align:center; padding:24px; color:#89919A;">No loads registered in the audit monitor yet.</td></tr>
                `}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    `;

    this._bindEvents();
  }
}

customElements.define('custom-excel-uploader-pro', ExcelUploaderPro);
