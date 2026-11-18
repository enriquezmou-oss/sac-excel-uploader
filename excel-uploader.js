// Carga dinámica segura de SheetJS
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
    this._state = 'idle'; // idle | parsing | reviewing | done
    this._activeTab = 'upload';
    this._rows = [];
    this._headers = [];
    this._structErrors = [];
    this._file = null;
    this._validatedDataJson = '[]';
    this._expectedSheetName = 'Sheet1';
    this._requiredColumns = 'ID,DESCRIPTION,H1,COMPANY,COSTCENTER,ASSET_CLASS,CAPITALIZED';
    this._uniqueKeyColumns = 'ID';
    this._numericColumns = '';
    this._customRules = [
      { name: 'Empresa válida', column: 'COMPANY', operator: '=', value: '1000', severity: 'error' },
      { name: 'ID no vacío', column: 'ID', operator: 'no-vacío', value: '', severity: 'error' }
    ];
    this._history = [];
    this._render();
  }

  connectedCallback() {
    this._setupEventListeners();
  }

  // Getters y Setters para propiedades expuestas en SAC
  set expectedSheetName(v) { this._expectedSheetName = v; }
  get expectedSheetName() { return this._expectedSheetName; }

  set requiredColumns(v) { this._requiredColumns = v; this._validateStructure(); this._render(); }
  get requiredColumns() { return this._requiredColumns; }

  set uniqueKeyColumns(v) { this._uniqueKeyColumns = v; }
  get uniqueKeyColumns() { return this._uniqueKeyColumns; }

  set numericColumns(v) { this._numericColumns = v; }
  get numericColumns() { return this._numericColumns; }

  get validatedDataJson() { return this._validatedDataJson; }

  _setupEventListeners() {
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

    // Gestión de pestañas
    this._shadow.querySelectorAll('.wgt-tab').forEach(tab => {
      tab.addEventListener('click', () => {
        if (tab.classList.contains('disabled')) return;
        this._activeTab = tab.dataset.tab;
        this._render();
      });
    });

    // Botón añadir regla
    const addRuleBtn = this._shadow.getElementById('addRuleBtn');
    if (addRuleBtn) {
      addRuleBtn.addEventListener('click', () => {
        const cols = this._requiredColumns ? this._requiredColumns.split(',').map(c => c.trim()) : [];
        this._customRules.push({ name: 'Regla personalizada', column: cols[0] || 'ID', operator: 'no-vacío', value: '', severity: 'error' });
        this._validateStructure();
        this._render();
      });
    }

    // Eliminar regla
    this._shadow.querySelectorAll('.rule-rm').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const idx = parseInt(e.target.dataset.idx, 10);
        this._customRules.splice(idx, 1);
        this._validateStructure();
        this._render();
      });
    });
  }

  _processFile(file) {
    this._file = file;
    this._state = 'parsing';
    this._render();

    if (typeof XLSX === 'undefined') {
      alert('La librería SheetJS está cargando. Por favor, inténtalo de nuevo en un segundo.');
      this._state = 'idle';
      this._render();
      return;
    }

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = this._expectedSheetName || workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        if (!worksheet) {
          throw new Error(`No se encontró la pestaña "${sheetName}" en el libro de Excel.`);
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

        const errCount = this._rows.filter(r => r._errors.length > 0).length;
        this._history.unshift({
          time: new Date().toLocaleTimeString(),
          file: file.name,
          rows: this._rows.length,
          errors: errCount,
          status: errCount > 0 ? 'Con incidencias' : 'Válido'
        });

        // Preparar y emitir datos validados para SAC
        this._validatedDataJson = JSON.stringify(this._rows.map(r => {
          const clean = { ...r };
          delete clean._rowNum;
          delete clean._errors;
          return clean;
        }));

        this.dispatchEvent(new CustomEvent('propertiesChanged', {
          bubbles: true, composed: true,
          detail: { properties: { validatedDataJson: this._validatedDataJson } }
        }));

        this._render();
      } catch (err) {
        this._state = 'idle';
        alert('Error al procesar el Excel: ' + err.message);
        this._render();
      }
    };
    reader.readAsArrayBuffer(file);
  }

  _validateStructure() {
    this._structErrors = [];
    const reqCols = this._requiredColumns ? this._requiredColumns.split(',').map(c => c.trim()).filter(Boolean) : [];
    
    reqCols.forEach(col => {
      if (!this._headers.length || !this._headers.includes(col)) {
        if (this._headers.length > 0) {
          this._structErrors.push(`Falta la columna obligatoria: "${col}"`);
        }
      }
    });

    this._rows.forEach(row => {
      row._errors = [];
      // Validar obligatorias vacías
      reqCols.forEach(col => {
        if (row[col] === '' || row[col] === null || row[col] === undefined) {
          row._errors.push(`Columna ${col} vacía`);
        }
      });

      // Validar reglas personalizadas
      this._customRules.forEach(rule => {
        if (rule.column && row[rule.column] !== undefined) {
          const val = String(row[rule.column]).trim();
          if (rule.operator === 'no-vacío' && !val) {
            row._errors.push(`Regla [${rule.name}]: ${rule.column} requerido`);
          } else if (rule.operator === '=' && val !== rule.value) {
            row._errors.push(`Regla [${rule.name}]: ${rule.column} debe ser ${rule.value}`);
          }
        }
      });
    });
  }

  _getReadableState() {
    switch (this._state) {
      case 'parsing':   return '⏳ Leyendo y validando fichero...';
      case 'reviewing': return '👀 Pendiente de revisión';
      case 'done':      return '✅ Importado con éxito';
      case 'idle':
      default:          return '⏸️ Esperando archivo Excel';
    }
  }

  _render() {
    const hasData = this._rows.length > 0;
    const errCount = this._rows.filter(r => r._errors.length > 0).length;
    const okCount = this._rows.length - errCount;
    const readableState = this._getReadableState();
    const cols = this._requiredColumns ? this._requiredColumns.split(',').map(c => c.trim()) : [];

    this._shadow.innerHTML = `
      <style>
        :host { display: block; font-family: "72", Arial, sans-serif; font-size: 13px; color: #1D2D3E; background: #F5F6F7; border: 1px solid #D1D7DD; border-radius: 4px; overflow: hidden; height: 100%; box-sizing: border-box; }
        .wgt-header { display: flex; align-items: center; justify-content: space-between; padding: 12px 16px; background: #fff; border-bottom: 1px solid #D1D7DD; }
        .wgt-title { font-size: 15px; font-weight: 700; color: #1D2D3E; }
        .wgt-status { font-size: 12px; padding: 4px 10px; border-radius: 12px; font-weight: 600; }
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
        .drop-zone.dragover { border-color: #107E3E; background: #E6F9EC; }
        .file-input { display: none; }

        .val-summary { display: flex; gap: 8px; margin-bottom: 12px; }
        .val-chip { display: flex; align-items: center; gap: 6px; padding: 4px 12px; border-radius: 14px; font-size: 12px; font-weight: 700; border: 1px solid; }
        .vc-ok { background: #E6F9EC; border-color: #AFD9B8; color: #0A6C2B; }
        .vc-err { background: #FFEAEA; border-color: #F4ADAD; color: #9E0000; }
        .val-dot { width: 8px; height: 8px; border-radius: 50%; }
        .vc-ok .val-dot { background: #107E3E; }
        .vc-err .val-dot { background: #BB0000; }

        .tbl-wrap { overflow: auto; border: 1px solid #D1D7DD; border-radius: 4px; background: #fff; max-height: 240px; }
        table { width: 100%; border-collapse: collapse; font-size: 12px; }
        thead th { position: sticky; top: 0; background: #F5F6F7; font-weight: 700; padding: 8px 10px; border-bottom: 1px solid #D1D7DD; text-align: left; }
        tbody td { padding: 6px 10px; border-bottom: 1px solid #EDF0F2; }
        tbody tr.row-err { background: #FFF5F5; }
        td.cell-err { background: #FFEAEA; color: #9E0000; font-weight: 600; }

        .rule-item { display: flex; align-items: center; gap: 8px; padding: 8px 12px; background: #fff; border: 1px solid #D1D7DD; border-radius: 4px; margin-bottom: 8px; font-size: 12px; }
        .rule-item input, .rule-item select { padding: 4px 6px; border: 1px solid #BCC8D4; border-radius: 4px; font-size: 12px; }
        .rule-rm { margin-left: auto; background: none; border: none; color: #BB0000; font-weight: bold; cursor: pointer; font-size: 16px; }

        .btn { padding: 6px 18px; border-radius: 4px; font-size: 13px; font-weight: 600; cursor: pointer; border: 1px solid #0070F2; background: #0070F2; color: #fff; }
        .btn:hover { background: #0057C0; }
        .btn-ghost { background: #fff; color: #0070F2; }
      </style>

      <div class="wgt-header">
        <div class="wgt-title">Excel Planning Uploader Pro</div>
        <div class="wgt-status ${this._state}">Estado: ${readableState}</div>
      </div>

      <div class="wgt-tabs">
        <div class="wgt-tab ${this._activeTab === 'upload' ? 'active' : ''}" data-tab="upload">1. Carga</div>
        <div class="wgt-tab ${this._activeTab === 'data' ? 'active' : ''} ${!hasData ? 'disabled' : ''}" data-tab="data">2. Datos (${this._rows.length})</div>
        <div class="wgt-tab ${this._activeTab === 'rules' ? 'active' : ''}" data-tab="rules">3. Reglas (${this._customRules.length})</div>
        <div class="wgt-tab ${this._activeTab === 'history' ? 'active' : ''}" data-tab="history">4. Historial</div>
      </div>

      <div class="wgt-content">
        <!-- PESTAÑA 1: CARGA -->
        <div class="wgt-pane ${this._activeTab === 'upload' ? 'active' : ''}">
          <div class="drop-zone" id="dropZone">
            <div style="font-size: 15px; font-weight: 600; margin-bottom: 6px;">Arrastra tu fichero Excel de planificación aquí</div>
            <div style="font-size: 12px; color: #6A6D70; margin-bottom: 14px;">Columnas requeridas: ${this._requiredColumns}</div>
            <button class="btn" type="button">Seleccionar archivo</button>
            <input type="file" id="fileInput" class="file-input" accept=".xlsx, .xls" />
          </div>
        </div>

        <!-- PESTAÑA 2: DATOS Y VALIDACIONES -->
        <div class="wgt-pane ${this._activeTab === 'data' ? 'active' : ''}">
          ${this._structErrors.length > 0 ? `<div style="background: #FFEAEA; color: #9E0000; padding: 10px; border-radius: 4px; margin-bottom: 12px; font-size: 12px;"><strong>Avisos de Estructura:</strong><br/>${this._structErrors.join('<br/>')}</div>` : ''}
          
          <div class="val-summary">
            <div class="val-chip vc-ok"><span class="val-dot"></span> Válidas: ${okCount}</div>
            <div class="val-chip vc-err"><span class="val-dot"></span> Con incidencias: ${errCount}</div>
          </div>

          <div class="tbl-wrap">
            <table>
              <thead>
                <tr>
                  <th style="width: 40px;">#</th>
                  ${this._headers.map(h => `<th>${h}</th>`).join('')}
                </tr>
              </thead>
              <tbody>
                ${this._rows.map((r) => `
                  <tr class="${r._errors.length > 0 ? 'row-err' : ''}">
                    <td style="color: #89919A; font-size: 11px;">${r._rowNum}</td>
                    ${this._headers.map(h => `<td class="${r._errors.some(e => e.includes(h)) ? 'cell-err' : ''}" title="${r._errors.join(', ')}">${r[h]}</td>`).join('')}
                  </tr>
                `).join('')}
              </tbody>
            </table>
          </div>
        </div>

        <!-- PESTAÑA 3: REGLAS -->
        <div class="wgt-pane ${this._activeTab === 'rules' ? 'active' : ''}">
          <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px;">
            <strong style="font-size: 12px;">Reglas de validación activas</strong>
            <button class="btn btn-ghost" id="addRuleBtn" type="button" style="font-size: 11px; padding: 4px 10px;">+ Añadir regla</button>
          </div>
          <div>
            ${this._customRules.map((rule, idx) => `
              <div class="rule-item">
                <input type="text" value="${rule.name}" style="width: 110px;" disabled />
                <select disabled><option>${rule.column}</option></select>
                <select disabled><option>${rule.operator}</option></select>
                <input type="text" value="${rule.value}" placeholder="Valor" style="width: 60px;" disabled />
                <button class="rule-rm" data-idx="${idx}" type="button" title="Eliminar regla">×</button>
              </div>
            `).join('')}
            ${this._customRules.length === 0 ? '<div style="text-align: center; color: #89919A; padding: 16px;">No hay reglas configuradas.</div>' : ''}
          </div>
        </div>

        <!-- PESTAÑA 4: HISTORIAL -->
        <div class="wgt-pane ${this._activeTab === 'history' ? 'active' : ''}">
          ${this._history.length > 0 ? this._history.map(h => `
            <div style="background: #fff; border: 1px solid #D1D7DD; border-radius: 4px; padding: 10px; margin-bottom: 6px; font-size: 12px; display: flex; justify-content: space-between; align-items: center;">
              <div><strong>${h.file}</strong> (${h.rows} filas)</div>
              <div style="color: ${h.errors > 0 ? '#BB0000' : '#107E3E'}; font-weight: bold;">${h.status}</div>
              <div style="color: #89919A; font-size: 11px;">${h.time}</div>
            </div>
          `).join('') : '<div style="text-align: center; color: #89919A; padding: 24px;">No hay importaciones previas registradas.</div>'}
        </div>
      </div>
    `;

    this._setupEventListeners();
  }
}

customElements.define('custom-excel-uploader-pro', ExcelUploaderPro);
