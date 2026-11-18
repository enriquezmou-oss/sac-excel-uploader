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
    this._rows = [];
    this._headers = [];
    this._fileName = 'No file chosen';
    this._sheetName = 'Sheet1';
    this._searchQuery = '';
    this._showPreview = true;
    this._requiredColumns = ['ID', 'DESCRIPTION', 'H1', 'COMPANY', 'COSTCENTER', 'ASSET_CLASS', 'CAPITALIZED'];
    this._render();
  }

  connectedCallback() {
    this._setupListeners();
  }

  _setupListeners() {
    const fileInput = this._shadow.getElementById('fileInput');
    const uploadBtn = this._shadow.getElementById('uploadBtn');
    const clearBtn = this._shadow.getElementById('clearBtn');
    const searchInput = this._shadow.getElementById('searchInput');
    const downloadTemplateBtn = this._shadow.getElementById('downloadTemplateBtn');
    const downloadErrorLogBtn = this._shadow.getElementById('downloadErrorLogBtn');
    const togglePreviewBtn = this._shadow.getElementById('togglePreviewBtn');

    if (uploadBtn && fileInput) {
      uploadBtn.addEventListener('click', () => fileInput.click());
    }

    if (fileInput) {
      fileInput.addEventListener('change', (e) => {
        if (e.target.files.length) {
          this._fileName = e.target.files[0].name;
          this._processFile(e.target.files[0]);
        }
      });
    }

    if (clearBtn) {
      clearBtn.addEventListener('click', () => {
        this._rows = [];
        this._headers = [];
        this._fileName = 'No file chosen';
        this._sheetName = 'Sheet1';
        this._render();
      });
    }

    if (searchInput) {
      searchInput.addEventListener('input', (e) => {
        this._searchQuery = e.target.value.toLowerCase();
        this._renderTableBody();
      });
    }

    if (downloadTemplateBtn) {
      downloadTemplateBtn.addEventListener('click', () => this._downloadTemplate());
    }

    if (downloadErrorLogBtn) {
      downloadErrorLogBtn.addEventListener('click', () => this._downloadErrorLog());
    }

    if (togglePreviewBtn) {
      togglePreviewBtn.addEventListener('click', () => {
        this._showPreview = !this._showPreview;
        const body = this._shadow.getElementById('previewBody');
        if (body) body.style.display = this._showPreview ? 'block' : 'none';
      });
    }
  }

  _downloadTemplate() {
    if (typeof XLSX === 'undefined') return;
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([this._requiredColumns]);
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, 'Asset_Master_Template.xlsx');
  }

  _downloadErrorLog() {
    const errors = this._rows.filter(r => r._errors && r._errors.length > 0);
    if (errors.length === 0) {
      alert('No hay errores para descargar.');
      return;
    }
    let csv = 'Row,ID,Errors\n';
    errors.forEach(r => {
      csv += `${r._rowNum},${r.ID || ''},"${r._errors.join('; ')}"\n`;
    });
    const blob = new Blob([csv], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'Validation_Error_Log.csv';
    a.click();
  }

  _processFile(file) {
    if (typeof XLSX === 'undefined') {
      alert('SheetJS cargando, intenta de nuevo.');
      return;
    }

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        this._sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[this._sheetName];
        const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        if (json.length === 0) throw new Error('El archivo está vacío');

        this._headers = json[0].map(h => String(h).trim());
        const rawRows = json.slice(1).map((row, idx) => {
          const obj = { _rowNum: idx + 2, _errors: {} };
          this._headers.forEach((h, i) => {
            obj[h] = row[i] !== undefined ? row[i] : '';
          });
          return obj;
        }).filter(r => Object.values(r).some(v => v !== '' && v !== '_rowNum' && v !== '_errors'));

        // Validaciones de negocio idénticas al diseño
        const idCounts = {};
        rawRows.forEach(r => {
          const id = String(r.ID || '').trim();
          if (id) idCounts[id] = (idCounts[id] || 0) + 1;
        });

        rawRows.forEach(row => {
          row._errors = {};
          // 1. Columnas obligatorias / celdas vacías
          this._requiredColumns.forEach(col => {
            if (row[col] === '' || row[col] === null || row[col] === undefined) {
              row._errors[col] = 'Blank cell not allowed';
            }
          });

          // 2. Duplicados en ID
          const id = String(row.ID || '').trim();
          if (id && idCounts[id] > 1) {
            row._errors['ID'] = 'Duplicate value not allowed';
          }

          // 3. Validación de Cost Center (ejemplo: si no es numérico de 8 dígitos)
          const cc = String(row.COSTCENTER || '').trim();
          if (cc && !/^\d{8}$/.test(cc)) {
            row._errors['COSTCENTER'] = 'Invalid Cost Center';
          }
        });

        this._rows = rawRows;
        this._render();
      } catch (err) {
        alert('Error al procesar: ' + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  }

  _render() {
    const rowsRead = this._rows.length;
    const invalidRowsCount = this._rows.filter(r => Object.keys(r._errors).length > 0).length;
    const validRowsCount = rowsRead - invalidRowsCount;
    const isInvalid = invalidRowsCount > 0;

    this._shadow.innerHTML = `
      <style>
        :host { display: block; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; font-size: 13px; color: #333; background: #fff; border: 1px solid #d1d5db; border-radius: 6px; padding: 16px; box-sizing: border-box; width: 100%; height: 100%; overflow: auto; }
        .top-bar { display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px; }
        .title-area h2 { margin: 0; font-size: 16px; font-weight: 700; color: #1e3a8a; }
        .title-area p { margin: 2px 0 0; font-size: 11px; color: #6b7280; }
        .top-actions { display: flex; gap: 8px; align-items: center; }
        
        .btn { padding: 6px 12px; font-size: 12px; font-weight: 600; border-radius: 4px; cursor: pointer; border: 1px solid #d1d5db; background: #fff; color: #374151; }
        .btn:hover { background: #f9fafb; }
        .btn-primary { background: #2563eb; color: #fff; border-color: #2563eb; }
        .btn-primary:hover { background: #1d4ed8; }
        .badge-err { background: #fee2e2; color: #991b1b; border: 1px solid #fca5a5; padding: 4px 10px; border-radius: 4px; font-weight: 600; font-size: 12px; }
        .badge-ok { background: #d1fae5; color: #065f46; border: 1px solid #6ee7b7; padding: 4px 10px; border-radius: 4px; font-weight: 600; font-size: 12px; }

        .control-row { display: flex; justify-content: space-between; align-items: center; background: #f9fafb; padding: 12px; border: 1px solid #e5e7eb; border-radius: 6px; margin-bottom: 16px; gap: 12px; }
        .control-left { display: flex; gap: 8px; align-items: center; }
        .file-box { border: 1px solid #d1d5db; background: #fff; padding: 4px 8px; border-radius: 4px; font-size: 12px; color: #374151; min-width: 180px; }
        .search-box { padding: 6px 10px; border: 1px solid #d1d5db; border-radius: 4px; font-size: 12px; width: 220px; }

        .metrics-grid { display: grid; grid-template-columns: repeat(6, 1fr); gap: 1px; background: #e5e7eb; border: 1px solid #e5e7eb; border-radius: 6px; overflow: hidden; margin-bottom: 16px; }
        .metric-card { background: #fff; padding: 10px 14px; }
        .metric-label { font-size: 11px; color: #6b7280; margin-bottom: 2px; }
        .metric-val { font-size: 15px; font-weight: 700; color: #1f2937; }

        .progress-section { margin-bottom: 16px; }
        .progress-header { display: flex; justify-content: space-between; font-size: 11px; color: #4b5563; margin-bottom: 4px; font-weight: 600; }
        .progress-bar-bg { background: #e5e7eb; height: 6px; border-radius: 3px; overflow: hidden; }
        .progress-bar-fill { background: #2563eb; height: 100%; width: ${rowsRead > 0 ? '100%' : '0%'}; }

        .preview-header-row { display: flex; justify-content: space-between; align-items: center; border-bottom: 2px solid #e5e7eb; padding-bottom: 6px; margin-bottom: 8px; font-weight: 700; color: #1f2937; }
        
        .tbl-container { overflow-x: auto; max-height: 280px; border: 1px solid #e5e7eb; border-radius: 4px; }
        table { width: 100%; border-collapse: collapse; font-size: 12px; text-align: left; }
        thead th { background: #f3f4f6; padding: 8px 10px; font-weight: 600; border-bottom: 1px solid #e5e7eb; color: #374151; position: sticky; top: 0; }
        tbody td { padding: 8px 10px; border-bottom: 1px solid #f3f4f6; color: #1f2937; }
        tbody tr:hover { background: #f9fafb; }
        
        td.cell-error { background: #fee2e2 !important; color: #991b1b; }
        .err-msg { font-size: 10px; color: #991b1b; display: block; margin-top: 2px; font-weight: 600; }

        .footer-note { font-size: 11px; color: #6b7280; margin-top: 12px; }
      </style>

      <div class="top-bar">
        <div class="title-area">
          <h2>Excel Asset Upload</h2>
          <p>Upload and validate Excel file</p>
        </div>
        <div class="top-actions">
          <button class="btn" id="downloadTemplateBtn" type="button">Download Template</button>
          <div class="${isInvalid ? 'badge-err' : 'badge-ok'}">${isInvalid ? 'Validation Error' : 'Validation Valid'}</div>
          <button class="btn" id="downloadErrorLogBtn" type="button">Download Error Log</button>
          <button class="btn" id="togglePreviewBtn" type="button">Toggle Preview</button>
        </div>
      </div>

      <div class="control-row">
        <div class="control-left">
          <button class="btn btn-primary" id="uploadBtn" type="button">Choose File</button>
          <span class="file-box">${this._fileName}</span>
          <button class="btn btn-primary" id="uploadExecBtn" style="display:none">Upload</button>
          <button class="btn" id="clearBtn" type="button">Clear</button>
        </div>
        <div>
          <input type="text" class="search-box" id="searchInput" placeholder="Search in preview..." value="${this._searchQuery}" />
        </div>
      </div>

      <div style="font-size: 11px; color: #4b5563; margin-bottom: 8px; font-weight: 600;">
        Required Columns: ${this._requiredColumns.join(', ')}
      </div>

      <div class="metrics-grid">
        <div class="metric-card"><div class="metric-label">Rows Read</div><div class="metric-val">${rowsRead}</div></div>
        <div class="metric-card"><div class="metric-label">Valid Rows</div><div class="metric-val">${validRowsCount}</div></div>
        <div class="metric-card"><div class="metric-label">Invalid Rows</div><div class="metric-val">${invalidRowsCount}</div></div>
        <div class="metric-card"><div class="metric-label">Sheet</div><div class="metric-val" style="font-size: 13px;">${this._sheetName}</div></div>
        <div class="metric-card"><div class="metric-label">Columns</div><div class="metric-val">${this._headers.length || this._requiredColumns.length}</div></div>
        <div class="metric-card"><div class="metric-label">Validation</div><div class="metric-val" style="font-size: 13px; color: ${isInvalid ? '#991b1b' : '#065f46'};">${isInvalid ? 'Invalid' : 'Valid'}</div></div>
      </div>

      <div class="progress-section">
        <div class="progress-header"><span>Completed</span><span>100%</span></div>
        <div class="progress-bar-bg"><div class="progress-bar-fill"></div></div>
      </div>

      <div id="previewBody">
        <div class="preview-header-row">
          <span>Preview</span>
          <span style="font-size: 11px; color: #6b7280; font-weight: normal;">${rowsRead} row(s) shown</span>
        </div>
        <div class="tbl-container">
          <table>
            <thead>
              <tr>
                <th style="width: 40px;">#</th>
                ${this._requiredColumns.map(h => `<th>${h}</th>`).join('')}
              </tr>
            </thead>
            <tbody id="tableBody">
              <!-- Renderizado dinámico -->
            </tbody>
          </table>
        </div>
      </div>

      <div class="footer-note">Supported template: Sheet1 with required business columns</div>
      <input type="file" id="fileInput" style="display:none" accept=".xlsx, .xls, .xlsm" />
    `;

    this._setupListeners();
    this._renderTableBody();
  }

  _renderTableBody() {
    const tbody = this._shadow.getElementById('tableBody');
    if (!tbody) return;

    const filtered = this._rows.filter(r => {
      if (!this._searchQuery) return true;
      return Object.values(r).some(val => String(val).toLowerCase().includes(this._searchQuery));
    });

    if (filtered.length === 0) {
      tbody.innerHTML = `<tr><td colspan="${this._requiredColumns.length + 1}" style="text-align: center; padding: 24px; color: #9ca3af;">No records found. Upload an Excel file to preview data.</td></tr>`;
      return;
    }

    tbody.innerHTML = filtered.map(r => {
      return `
        <tr>
          <td style="color: #6b7280; font-size: 11px;">${r._rowNum}</td>
          ${this._requiredColumns.map(col => {
            const hasErr = r._errors && r._errors[col];
            return `
              <td class="${hasErr ? 'cell-error' : ''}">
                ${r[col] !== undefined ? r[col] : ''}
                ${hasErr ? `<span class="err-msg">${r._errors[col]}</span>` : ''}
              </td>
            `;
          }).join('')}
        </tr>
      `;
    }).join('');
  }
}

customElements.define('custom-excel-uploader-pro', ExcelUploaderPro);
