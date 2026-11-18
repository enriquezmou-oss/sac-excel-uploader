class ExcelUploaderPro extends HTMLElement {
  constructor() {
    super();
    this._shadow = this.attachShadow({ mode: 'open' });
    this._state = 'idle';
    this._activeTab = 'upload';
    this._rows = [];
    this._headers = [];
    this._structErrors = [];
    this._page = 0;
    this._filter = 'all';
    this._file = null;
    this._validatedDataJson = '[]';
    this._expectedSheetName = 'Sheet1';
    this._requiredColumns = 'ID,DESCRIPTION,H1,COMPANY,COSTCENTER,ASSET_CLASS,CAPITALIZED';
    this._uniqueKeyColumns = 'ID';
    this._numericColumns = '';
    this._customRules = [];
    this._history = [];
    this._render();
  }

  connectedCallback() {
    this._setupEvents();
  }

  set expectedSheetName(v) { this._expectedSheetName = v; }
  get expectedSheetName() { return this._expectedSheetName; }

  set requiredColumns(v) { this._requiredColumns = v; }
  get requiredColumns() { return this._requiredColumns; }

  set uniqueKeyColumns(v) { this._uniqueKeyColumns = v; }
  get uniqueKeyColumns() { return this._uniqueKeyColumns; }

  set numericColumns(v) { this._numericColumns = v; }
  get numericColumns() { return this._numericColumns; }

  set validatedDataJson(v) { this._validatedDataJson = v; }
  get validatedDataJson() { return this._validatedDataJson; }

  _setupEvents() {
    const dz = this._shadow.getElementById('dropZone');
    const fi = this._shadow.getElementById('fileInput');
    if (!dz || !fi) return;

    dz.addEventListener('dragover', (e) => { e.preventDefault(); dz.classList.add('dragover'); });
    dz.addEventListener('dragleave', () => dz.classList.remove('dragover'));
    dz.addEventListener('drop', (e) => {
      e.preventDefault();
      dz.classList.remove('dragover');
      if (e.dataTransfer.files.length) this._processFile(e.dataTransfer.files[0]);
    });
    dz.addEventListener('click', () => fi.click());
    fi.addEventListener('change', (e) => {
      if (e.target.files.length) this._processFile(e.target.files[0]);
    });

    this._shadow.querySelectorAll('.wgt-tab').forEach(tab => {
      tab.addEventListener('click', () => {
        if (tab.classList.contains('disabled')) return;
        this._activeTab = tab.dataset.tab;
        this._render();
      });
    });
  }

  _processFile(file) {
    this._file = file;
    this._state = 'parsing';
    this._render();

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = this._expectedSheetName || workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        if (!worksheet) {
          throw new Error(`No se encontró la pestaña "${sheetName}" en el Excel.`);
        }

        const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        if (json.length === 0) throw new Error('El archivo Excel está vacío.');

        this._headers = json[0].map(h => String(h).trim());
        this._rows = json.slice(1).map((row, idx) => {
          const obj = { _rowNum: idx + 2, _errors: [] };
          this._headers.forEach((h, i) => {
            obj[h] = row[i] !== undefined ? row[i] : '';
          });
          return obj;
        }).filter(row => Object.values(row).some(v => v !== ''));

        this._validateStructure();
        this._state = 'reviewing';
        this._activeTab = 'data';
        this._render();
      } catch (err) {
        this._state = 'idle';
        alert('Error al leer el Excel: ' + err.message);
        this._render();
      }
    };
    reader.readAsArrayBuffer(file);
  }

  _validateStructure() {
    this._structErrors = [];
    const reqCols = this._requiredColumns.split(',').map(c => c.trim()).filter(Boolean);
    
    reqCols.forEach(col => {
      if (!this._headers.includes(col)) {
        this._structErrors.push(`Falta la columna obligatoria: "${col}"`);
      }
    });

    this._rows.forEach(row => {
      row._errors = [];
      reqCols.forEach(col => {
        if (row[col] === '' || row[col] === null || row[col] === undefined) {
          row._errors.push(`Columna ${col} vacía`);
        }
      });
    });
  }

  _emitProp(prop, value) {
    this.dispatchEvent(new CustomEvent('propertiesChanged', {
      bubbles: true, composed: true,
      detail: { properties: { [prop]: value } }
    }));
  }

  _render() {
    const hasData = this._rows.length > 0;
    const errCount = this._rows.filter(r => r._errors.length > 0).length;

    this._shadow.innerHTML = `
      <style>
        :host { display: block; font-family: "72", Arial, sans-serif; font-size: 13px; color: #1D2D3E; background: #F5F6F7; border: 1px solid #D1D7DD; border-radius: 4px; overflow: hidden; height: 100%; box-sizing: border-box; }
        .wgt-header { display: flex; align-items: center; justify-content: space-between; padding: 12px 16px; background: #fff; border-bottom: 1px solid #D1D7DD; }
        .wgt-title { font-size: 15px; font-weight: 700; color: #1D2D3E; }
        .wgt-status { font-size: 12px; padding: 3px 10px; border-radius: 10px; font-weight: 600; }
        .wgt-status.idle { background: #EFF1F2; color: #6A6D70; }
        .wgt-status.parsing { background: #E6F0FF; color: #0064D9; }
        .wgt-status.reviewing { background: #FEF6E0; color: #A35200; }
        .wgt-status.done { background: #E6F9EC; color: #107E3E; }
        .wgt-tabs { display: flex; background: #fff; border-bottom: 1px solid #D1D7DD; padding: 0 12px; }
        .wgt-tab { padding: 8px 14px; font-size: 13px; font-weight: 600; color: #6A6D70; cursor: pointer; border-bottom: 2px solid transparent; margin-bottom: -1px; }
        .wgt-tab.active { color: #0070F2; border-bottom-color: #0070F2; }
        .wgt-tab.disabled { color: #BCC3CA; pointer-events: none; }
        .wgt-content { padding: 16px; background: #F5F6F7; height: calc(100% - 110px); overflow-y: auto; box-sizing: border-box; }
        .wgt-pane { display: none; }
        .wgt-pane.active { display: block; }
        .drop-zone { border: 2px dashed #89919A; border-radius: 8px; padding: 40px 24px; text-align: center; cursor: pointer; background: #fff; transition: all .2s; }
        .drop-zone:hover { border-color: #0070F2; background: #EFF5FF; }
        .file-input { display: none; }
        .tbl-wrap { overflow: auto; border: 1px solid #D1D7DD; border-radius: 4px; background: #fff; max-height: 280px; }
        table { width: 100%; border-collapse: collapse; font-size: 12px; }
        thead th { position: sticky; top: 0; background: #F5F6F7; font-weight: 700; padding: 8px 10px; border-bottom: 1px solid #D1D7DD; text-align: left; }
        tbody td { padding: 6px 10px; border-bottom: 1px solid #EDF0F2; }
        tbody tr.row-err { background: #FFF5F5; }
        td.cell-err { background: #FFEAEA; color: #9E0000; font-weight: 600; }
        .btn { padding: 6px 18px; border-radius: 4px; font-size: 13px; font-weight: 600; cursor: pointer; border: 1px solid #0070F2; background: #0070F2; color: #fff; }
        .btn:hover { background: #0057C0; }
      </style>

      <div class="wgt-header">
        <div class="wgt-title">Excel Planning Uploader Pro</div>
        <div class="wgt-status ${this._state}">Estado: ${this._state.toUpperCase()}</div>
      </div>

      <div class="wgt-tabs">
        <div class="wgt-tab ${this._activeTab === 'upload' ? 'active' : ''}" data-tab="upload">1. Carga</div>
        <div class="wgt-tab ${this._activeTab === 'data' ? 'active' : ''} ${!hasData ? 'disabled' : ''}" data-tab="data">2. Datos (${this._rows.length})</div>
        <div class="wgt-tab ${this._activeTab === 'history' ? 'active' : ''}" data-tab="history">3. Historial</div>
      </div>

      <div class="wgt-content">
        <div class="wgt-pane ${this._activeTab === 'upload' ? 'active' : ''}">
          <div class="drop-zone" id="dropZone">
            <div style="font-size: 15px; font-weight: 600; margin-bottom: 6px;">Arrastra tu fichero Excel de planificación aquí</div>
            <div style="font-size: 12px; color: #6A6D70; margin-bottom: 14px;">o haz clic para buscarlo en tu equipo (.xlsx)</div>
            <button class="btn" type="button">Seleccionar archivo</button>
            <input type="file" id="fileInput" class="file-input" accept=".xlsx, .xls" />
          </div>
        </div>

        <div class="wgt-pane ${this._activeTab === 'data' ? 'active' : ''}">
          ${this._structErrors.length > 0 ? `<div style="background: #FFEAEA; color: #9E0000; padding: 10px; border-radius: 4px; margin-bottom: 12px; font-size: 12px;"><strong>Errores de Estructura:</strong><br/>${this._structErrors.join('<br/>')}</div>` : ''}
          <div style="margin-bottom: 8px; font-size: 12px; font-weight: 600;">Registros leídos: ${this._rows.length} | Con errores: ${errCount}</div>
          <div class="tbl-wrap">
            <table>
              <thead>
                <tr>
                  <th style="width: 40px;">#</th>
                  ${this._headers.map(h => `<th>${h}</th>`).join('')}
                </tr>
              </thead>
              <tbody>
                ${this._rows.map((r, idx) => `
                  <tr class="${r._errors.length > 0 ? 'row-err' : ''}">
                    <td style="color: #89919A; font-size: 11px;">${r._rowNum}</td>
                    ${this._headers.map(h => `<td class="${r._errors.some(e => e.includes(h)) ? 'cell-err' : ''}" title="${r._errors.join(', ')}">${r[h]}</td>`).join('')}
                  </tr>
                `).join('')}
              </tbody>
            </table>
          </div>
        </div>

        <div class="wgt-pane ${this._activeTab === 'history' ? 'active' : ''}">
          <div style="text-align: center; color: #89919A; padding: 24px;">No hay importaciones previas en esta sesión.</div>
        </div>
      </div>
    `;

    this._setupEvents();
  }
}

customElements.define('custom-excel-uploader-pro', ExcelUploaderPro);
