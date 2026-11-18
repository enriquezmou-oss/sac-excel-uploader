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
    this._filesQueue = []; // [{ file: File, comment: string }]
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
    if (fileInput) {
      fileInput.addEventListener('change', (e) => {
        if (e.target.files.length) {
          this._handleFilesSelection(e.target.files);
          // Reset input to allow selecting the same file again if needed
          e.target.value = '';
        }
      });
    }

    const nextStep2Btn = shadow.getElementById('nextStep2Btn');
    const prevBtn = shadow.getElementById('prevStepBtn');
    const finishBtn = shadow.getElementById('finishBtn');
    const exitBtn = shadow.getElementById('exitBtn');
    const downloadLogBtn = shadow.getElementById('downloadLogBtn');
    const addRowBtn = shadow.getElementById('addRowBtn');

    if (nextStep2Btn) nextStep2Btn.addEventListener('click', () => { this._step = 3; this._render(); });
    if (prevBtn) prevBtn.addEventListener('click', () => { if (this._step > 1) { this._step--; this._render(); } });
    if (exitBtn) exitBtn.addEventListener('click', () => { this._step = 1; this._filesQueue = []; this._rows = []; this._render(); });

    if (downloadLogBtn) {
      downloadLogBtn.addEventListener('click', () => this._exportErrorLog());
    }

    if (addRowBtn) {
      addRowBtn.addEventListener('click', () => {
        const newRowNum = this._rows.length + 2;
        const newRow = { _rowNum: newRowNum, _sourceFile: 'Manual Entry', _errors: {} };
        this._requiredColumns.forEach(col => {
          newRow[col] = col === 'ID' ? 'PLN_' + Math.floor(Math.random() * 900 + 100) : '';
        });
        this._rows.push(newRow);
        this._validateRows();
        this._render();
      });
    }

    shadow.querySelectorAll('.file-comment-input').forEach(input => {
      input.addEventListener('input', (e) => {
        const idx = parseInt(e.target.dataset.index, 10);
        if (this._filesQueue[idx]) {
          this._filesQueue[idx].comment = e.target.value;
        }
      });
    });

    shadow.querySelectorAll('.remove-file-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const idx = parseInt(e.target.dataset.index, 10);
        this._filesQueue.splice(idx, 1);
        if (this._filesQueue.length === 0) {
          this._step = 1;
          this._rows = [];
        } else {
          this._reprocessQueueFiles();
        }
        this._render();
      });
    });

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

        const fileSummary = this._filesQueue.map(f => `${f.file.name} (${f.comment || 'No comment'})`).join('; ');
        this._history.unshift({
          id: 'LOAD-' + Math.floor(Math.random() * 90000 + 10000),
          user: this._currentUser,
          file: fileSummary,
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
    
    let addedCount = 0;
    let duplicateNames = [];

    for (let i = 0; i < fileList.length; i++) {
      const newFile = fileList[i];
      // Validar si el nombre del fichero ya existe en la cola
      const exists = this._filesQueue.some(item => item.file.name.toLowerCase() === newFile.name.toLowerCase());
      if (exists) {
        duplicateNames.push(newFile.name);
      } else {
        this._filesQueue.push({ file: newFile, comment: '' });
        addedCount++;
      }
    }

    if (duplicateNames.length > 0) {
      alert(`⚠️ The following file(s) were skipped because they are already in the queue: ${duplicateNames.join(', ')}`);
    }

    if (addedCount > 0) {
      this._isLoading = true;
      this._loadingText = 'Processing multiple sheets & Excel files...';
      this._render();
      this._processAllFilesInQueue();
    }
  }

  _reprocessQueueFiles() {
    if (this._filesQueue.length === 0) {
      this._rows = [];
      return;
    }
    this._isLoading = true;
    this._loadingText = 'Updating files queue...';
    this._render();
    this._processAllFilesInQueue();
  }

  _processAllFilesInQueue() {
    let processedCount = 0;
    let allCombinedRows = [];
    let detectedHeaders = [];
    let headersFound = false;

    this._filesQueue.forEach((item) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const wb = XLSX.read(data, { type: 'array' });
          
          // RECORRER TODAS LAS PESTAÑAS (SHEETS) DEL FICHERO EXCEL
          wb.SheetNames.forEach(sheetName => {
            const ws = wb.Sheets[sheetName];
            const json = XLSX.utils.sheet_to_json(ws, { header: 1 });

            if (json.length > 0) {
              const sheetHeaders = json[0].map(h => String(h).trim());
              if (!headersFound && sheetHeaders.length > 0) {
                detectedHeaders = sheetHeaders;
                headersFound = true;
              }

              const rows = json.slice(1).map((row, idx) => {
                const sourceLabel = `${item.file.name} [${sheetName}]`;
                const obj = { _rowNum: allCombinedRows.length + idx + 2, _sourceFile: sourceLabel, _errors: {} };
                
                // Mapear columnas según los cabeceros detectados
                const currentHeaders = sheetHeaders.length > 0 ? sheetHeaders : detectedHeaders;
                currentHeaders.forEach((h, i) => {
                  obj[h] = row[i] !== undefined ? row[i] : '';
                });
                return obj;
              }).filter(r => Object.keys(r).some(k => k !== '_rowNum' && k !== '_sourceFile' && k !== '_errors' && r[k] !== ''));

              allCombinedRows = allCombinedRows.concat(rows);
            }
          });
        } catch (err) {
          console.error(`Error processing ${item.file.name}:`, err);
        }

        processedCount++;
        if (processedCount === this._filesQueue.length) {
          this._headers = detectedHeaders.length > 0 ? detectedHeaders : this._requiredColumns;
          this._rows = allCombinedRows;
          this._validateRows();
          this._isLoading = false;
          this._render();
        }
      };
      reader.readAsArrayBuffer(item.file);
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
        :host { display: block; font-family: "72", "72full", Arial, Helvetica, sans-serif; font-size: 13px; color: #1d2d3e; background: #edf1f5; padding: 16px; box-sizing: border-box; height: 100%; overflow: auto; }
        .main-card { background: #FFFFFF; border: 1px solid #cdd6df; border-radius: 10px; box-shadow: 0 6px 20px rgba(0,0,0,0.06); overflow: hidden; position: relative; }
        
        .busy-overlay { position: absolute; top: 0; left: 0; right: 0; bottom: 0; background: rgba(255,255,255,0.85); display: flex; flex-direction: column; align-items: center; justify-content: center; z-index: 100; }
        .spinner { width: 40px; height: 40px; border: 4px solid #E5E9EC; border-top-color: #0070F2; border-radius: 50%; animation: spin 0.8s linear infinite; margin-bottom: 10px; }
        @keyframes spin { to { transform: rotate(360deg); } }

        .sap-header { display: flex; align-items: center; justify-content: space-between; background: linear-gradient(135deg, #18314f 0%, #29465b 100%); color: #FFFFFF; padding: 14px 22px; border-bottom: 3px solid #0070F2; }
        .sap-brand { display: flex; align-items: center; gap: 12px; }
        .sap-logo-svg { width: 44px; height: 22px; display: block; filter: drop-shadow(0 1px 2px rgba(0,0,0,0.2)); }
        .sap-title { font-size: 16px; font-weight: 700; letter-spacing: -0.2px; }

        .nav-tabs { display: flex; background: #FFFFFF; border-bottom: 1px solid #d1d7dd; padding: 0 20px; }
        .nav-tab { padding: 12px 20px; cursor: pointer; font-weight: 600; color: #555c64; border-bottom: 3px solid transparent; font-size: 13px; transition: all 0.2s; }
        .nav-tab:hover { color: #0070F2; }
        .nav-tab.active { color: #0070F2; border-bottom-color: #0070F2; background: rgba(0,112,242,0.03); }

        .tab-content { padding: 24px; background: #f5f7fa; }

        .wizard-steps { display: flex; justify-content: space-between; background: #FFF; border: 1px solid #d1d7dd; border-radius: 8px; padding: 14px 22px; margin-bottom: 20px; box-shadow: 0 2px 6px rgba(0,0,0,0.02); }
        .step-item { display: flex; align-items: center; gap: 10px; font-weight: 600; color: #7a828a; font-size: 12px; }
        .step-item.active { color: #0070F2; }
        .step-num { width: 24px; height: 24px; border-radius: 50%; background: #e5e9ec; color: #555c64; display: flex; align-items: center; justify-content: center; font-size: 11px; font-weight: 700; }
        .step-item.active .step-num { background: #0070F2; color: #FFF; box-shadow: 0 0 0 3px rgba(0,112,242,0.2); }

        .btn { padding: 8px 18px; font-size: 12px; font-weight: 600; border-radius: 6px; cursor: pointer; border: 1px solid #0070F2; background: #0070F2; color: #FFF; transition: all 0.2s; box-shadow: 0 2px 4px rgba(0,112,242,0.2); }
        .btn:hover { background: #005cb9; box-shadow: 0 4px 8px rgba(0,92,185,0.3); }
        .btn-ghost { background: #FFFFFF; color: #0070F2; border-color: #cdd6df; }
        .btn-ghost:hover { background: rgba(0,112,242,0.04); border-color: #0070F2; }
        .btn-danger { background: #FFF; color: #bb0000; border-color: #f5c6c6; }
        .btn-danger:hover { background: #feeaea; }

        /* ZONA DE CARGA CON BOTÓN NATIVO VISIBLE Y ACCESIBLE */
        .drop-zone { border: 2px dashed #0070F2; border-radius: 10px; padding: 40px; text-align: center; background: #FFFFFF; box-shadow: 0 2px 8px rgba(0,0,0,0.02); display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 12px; }
        .file-input-native { margin-top: 8px; font-size: 13px; color: #1d2d3e; }

        .metrics-bar { display: grid; grid-template-columns: repeat(5, 1fr); gap: 14px; margin-bottom: 16px; }
        .metric-box { background: #FFFFFF; border: 1px solid #cdd6df; padding: 12px 16px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.02); }
        .metric-title { font-size: 11px; color: #6a6d70; margin-bottom: 4px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px; }
        .metric-val { font-size: 16px; font-weight: 700; color: #1d2d3e; }

        .dist-bar { display: flex; height: 8px; border-radius: 4px; overflow: hidden; background: #feeaea; margin-bottom: 20px; border: 1px solid #cdd6df; }
        .dist-ok { background: #107e3e; height: 100%; transition: width 0.4s ease; }

        .chips-row { display: flex; gap: 10px; margin-bottom: 14px; align-items: center; justify-content: space-between; }
        .chips-left { display: flex; gap: 8px; align-items: center; }
        .chip { padding: 5px 12px; border-radius: 14px; font-size: 11px; font-weight: 600; cursor: pointer; border: 1px solid #cdd6df; background: #FFF; color: #32363a; transition: all 0.2s; }
        .chip:hover { border-color: #0070F2; }
        .chip-all.active-chip { background: #354a5f; color: #FFF; border-color: #354a5f; }
        .chip-ok.active-chip { background: #107e3e; color: #FFF; border-color: #107e3e; }
        .chip-err.active-chip { background: #bb0000; color: #FFF; border-color: #bb0000; }

        .tbl-wrap { max-height: 320px; overflow: auto; border: 1px solid #cdd6df; border-radius: 8px; background: #FFFFFF; box-shadow: 0 2px 6px rgba(0,0,0,0.02); }
        table { width: 100%; border-collapse: collapse; font-size: 12px; text-align: left; }
        th { background: #f0f3f6; padding: 10px 12px; border-bottom: 2px solid #cdd6df; position: sticky; top: 0; font-weight: 700; color: #1d2d3e; }
        td { padding: 10px 12px; border-bottom: 1px solid #edf0f2; }
        td.cell-error { background: #feeaea; color: #bb0000; }
        .editable-cell { outline: none; border-bottom: 1px dashed #0070F2; display: inline-block; min-width: 60px; padding: 2px 4px; border-radius: 3px; }
        .editable-cell:focus { background: #FFF; border: 1px solid #0070F2; box-shadow: 0 0 0 2px rgba(0,112,242,0.2); }

        .wizard-footer { display: flex; justify-content: space-between; margin-top: 24px; border-top: 1px solid #cdd6df; padding-top: 16px; }
        
        .file-card-item { background: #fff; border: 1px solid #cdd6df; border-radius: 8px; padding: 12px 16px; margin-bottom: 10px; display: flex; align-items: center; justify-content: space-between; box-shadow: 0 2px 4px rgba(0,0,0,0.02); }
        .file-info { display: flex; align-items: center; gap: 10px; font-weight: 600; color: #0064d9; font-size: 12px; }
        .file-comment-input { border: 1px solid #cdd6df; border-radius: 4px; padding: 5px 10px; font-size: 12px; width: 300px; outline: none; }
        .file-comment-input:focus { border-color: #0070F2; box-shadow: 0 0 0 2px rgba(0,112,242,0.2); }
        .remove-file-btn { background: #feeaea; border: 1px solid #f5c6c6; color: #bb0000; font-weight: 700; cursor: pointer; font-size: 11px; padding: 4px 8px; border-radius: 4px; }
        .remove-file-btn:hover { background: #fcd5d5; }
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
            <svg class="sap-logo-svg" viewBox="0 0 100 48" xmlns="http://www.w3.org/2000/svg">
              <path d="M0 0h100v48H0z" fill="#0070F2"/>
              <path d="M28.2 35.8c-2.4 0-4.3-.6-5.8-1.7-1.5-1.1-2.2-2.7-2.2-4.9 0-1.8.6-3.2 1.9-4.2 1.3-1 3.2-1.7 5.7-2.1l4.4-.7c3-.5 5.2-1.4 6.7-2.6 1.5-1.2 2.3-2.9 2.3-5.2 0-3.1-1.3-5.5-3.8-7.2C35-0.6 31.5-1.4 27.2-1.4c-3.7 0-6.9.7-9.5 2.2l2.3 5.4c2.1-1.2 4.6-1.8 7.4-1.8 2 0 3.5.4 4.5 1.2 1 .8 1.5 1.9 1.5 3.3 0 1.5-.6 2.6-1.9 3.4-1.3.8-3.3 1.4-6 1.9l-4.2.7c-3.1.5-5.4 1.4-7 2.7-1.6 1.3-2.4 3.1-2.4 5.5 0 3 1.2 5.3 3.6 7 2.4 1.7 5.7 2.6 9.8 2.6 3.8 0 7.2-.7 10.2-2.2l-2.4-5.3c-2.4 1.3-5.2 2-8.5 2zm26.6-35h-7.6V35h7.6v-13h6.4l5.6 13h8.3l-6.5-14c3.9-1.2 6-3.9 6-8.1 0-4.2-2.5-7.1-7.5-8.7-1.9-.6-4.4-.9-7.5-.9H54.8v34.8zm7.6 6.3V14h5.2c2.2 0 3.7.4 4.5 1.2.8.8 1.2 1.9 1.2 3.3 0 1.4-.4 2.5-1.2 3.3-.8.8-2.3 1.2-4.5 1.2h-5.2zM79 35h7.6V1.5H79V35z" fill="#FFFFFF"/>
            </svg>
            <div class="sap-title">Excel Planning Uploader Pro (Multi-File)</div>
          </div>
          <div style="font-size: 11px; font-weight: 600; opacity: 0.9;">SAP Analytics Cloud BTP</div>
        </div>

        <div class="nav-tabs">
          <div class="nav-tab active" data-target="uploaderTab">Upload Wizard</div>
          <div class="nav-tab" data-target="monitorTab">Audit Monitor (${this._history.length})</div>
        </div>

        <div id="uploaderTab" class="tab-content">
          <div class="wizard-steps">
            <div class="step-item ${this._step === 1 ? 'active' : ''}"><div class="step-num">1</div> Multi-File & Multi-Sheet Upload</div>
            <div class="step-item ${this._step === 2 ? 'active' : ''}"><div class="step-num">2</div> Consolidation & Comments</div>
            <div class="step-item ${this._step === 3 ? 'active' : ''}"><div class="step-num">3</div> Global Validation & Edit</div>
          </div>

          ${this._step === 1 ? `
            <div class="drop-zone">
              <h3 style="margin:0; color:#0070F2;">Select one or multiple planning Excel files</h3>
              <p style="margin:0; color:#6a6d70; font-size:12px;">All sheets from each workbook will be automatically parsed and consolidated (.xlsx, .xls)</p>
              <input type="file" id="fileInput" class="file-input-native" multiple accept=".xlsx, .xls, .xlsm" />
            </div>
          ` : ''}

          ${this._step === 2 ? `
            <div>
              <h3 style="margin-top:0; color:#1d2d3e;">Files Queue & Business Comments</h3>
              <p style="color:#6a6d70; margin-bottom:14px;">Add a descriptive comment for each uploaded file for corporate traceability:</p>
              
              <div style="margin-bottom:20px;">
                ${this._filesQueue.map((item, idx) => `
                  <div class="file-card-item">
                    <div class="file-info">📄 ${item.file.name}</div>
                    <div style="display:flex; align-items:center; gap:10px;">
                      <input type="text" class="file-comment-input" data-index="${idx}" placeholder="Add comment/notes for this file..." value="${item.comment || ''}" />
                      <button class="remove-file-btn" data-index="${idx}" type="button">Remove</button>
                    </div>
                  </div>
                `).join('')}
              </div>

              <p>Mandatory fields required by the unified model:</p>
              <div style="display:flex; flex-wrap:wrap; gap:6px; margin-bottom:16px;">
                ${this._requiredColumns.map(c => `<span style="background:#e6f0ff; color:#0064d9; padding:5px 10px; border-radius:4px; font-weight:600; font-size:11px; border:1px solid #b3d1ff;">${c}</span>`).join('')}
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
                <div class="metric-box"><div class="metric-title">Correct</div><div class="metric-val" style="color:#107e3e;">${validRows}</div></div>
                <div class="metric-box"><div class="metric-title">With Errors</div><div class="metric-val" style="color:#bb0000;">${errorRows}</div></div>
                <div class="metric-box"><div class="metric-title">Global Status</div><div class="metric-val" style="font-size:12px; color:${errorRows > 0 ? '#bb0000':'#107e3e'}">${errorRows > 0 ? '⚠️ Review Errors' : '✅ Valid'}</div></div>
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
                <div style="display:flex; gap:8px; align-items:center;">
                  <button class="btn btn-ghost" id="addRowBtn" type="button" style="font-size:11px; padding:6px 12px;">+ Add New Row</button>
                  ${errorRows > 0 ? `<button class="btn btn-danger" id="downloadLogBtn" type="button" style="font-size:11px; padding:6px 14px;">📥 Download Error Log (CSV)</button>` : ''}
                </div>
              </div>

              <div class="tbl-wrap">
                <table>
                  <thead>
                    <tr>
                      <th style="width:30px;">#</th>
                      <th>Source File & Sheet</th>
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
                          <td style="color:#7a828a; font-size:11px;">${r._rowNum}</td>
                          <td style="font-size:11px; color:#0064d9; font-weight:600;">${r._sourceFile}</td>
                          ${this._requiredColumns.map(col => {
                            const err = r._errors[col];
                            return `
                              <td class="${err ? 'cell-error' : ''}" title="${err || 'Correct'}">
                                <span class="editable-cell" contenteditable="true" data-row="${r._rowNum - 2}" data-col="${col}">${r[col] !== undefined ? r[col] : ''}</span>
                                ${err ? `<span style="display:block; font-size:10px; color:#bb0000; font-weight:700; margin-top:2px;">⚠️ ${err}</span>` : ''}
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
              <div style="display:flex; gap:10px;">
                <button class="btn btn-ghost" id="exitBtn" type="button">Cancel</button>
                <button class="btn" id="finishBtn" type="button" ${errorRows > 0 ? 'style="opacity:0.6; cursor:not-allowed;"':''}>Consolidate and Load into SAC</button>
              </div>
            </div>
          ` : ''}
        </div>

        <div id="monitorTab" class="tab-content" style="display:none;">
          <h3 style="margin-top:0; color:#1d2d3e;">BTP Consolidation & Audit Monitor</h3>
          <p style="color:#6a6d70; font-size:12px;">Corporate traceability of executed multi-file planning loads.</p>
          
          <div style="border: 1px solid #cdd6df; border-radius: 8px; overflow:hidden; background:#FFF; box-shadow: 0 2px 6px rgba(0,0,0,0.02);">
            <table>
              <thead>
                <tr>
                  <th>Load ID</th>
                  <th>SAP User</th>
                  <th>Source Files & Comments</th>
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
                    <td><span style="background:#e6f9ec; color:#107e3e; padding:3px 10px; border-radius:12px; font-weight:600; font-size:11px;">${h.status}</span></td>
                  </tr>
                `).join('') : `
                  <tr><td colspan="6" style="text-align:center; padding:28px; color:#7a828a;">No loads registered in the audit monitor yet.</td></tr>
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
