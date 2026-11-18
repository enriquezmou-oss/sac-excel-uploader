class ExcelUploaderPro extends HTMLElement {
  constructor() {
    super();
    this._shadow = this.attachShadow({ mode: 'open' });
    this._step = 1;
    this._activeTab = 'uploaderTab'; // 'uploaderTab' | 'rulesTab' | 'monitorTab'
    this._filesQueue = [];
    this._rows = [];
    this._headers = [];
    this._currentUser = 'SAP User (Planning)';
    this._isLoading = false;
    this._loadingText = 'Processing files...';
    this._xlsxReady = typeof XLSX !== 'undefined';

    // Configuración de Formato, Tabulador (Delimitador) y Números
    this._parseConfig = {
      format: 'AUTO',       // 'AUTO' | 'XLSX' | 'CSV_TXT'
      delimiter: 'AUTO',    // 'AUTO' | '\t' | ';' | ',' | '|'
      numberFormat: 'STD',  // 'STD' (1,234.56) | 'EU' (1.234,56 -> 1234.56)
      headerRow: 1          // 1-based index
    };

    const defaultCols = ['ID', 'DESCRIPTION', 'H1', 'COMPANY', 'COSTCENTER', 'ASSET_CLASS', 'CAPITALIZED'];
    const savedCols = this._storageGet('sac_required_cols');
    this._requiredColumns = Array.isArray(savedCols) && savedCols.length > 0 ? savedCols : defaultCols;

    // Motor de Reglas de Validación Configurables por el Usuario
    const defaultRules = [
      { id: 'RULE-1', col: 'ID', type: 'UNIQUE', param: '', min: '', max: '', active: true, desc: 'No duplicate IDs allowed across consolidated sheets' },
      { id: 'RULE-2', col: 'COMPANY', type: 'IN_LIST', param: '1000, 2000, 950', min: '', max: '', active: true, desc: 'Must match SAP Master Data Company Codes' },
      { id: 'RULE-3', col: 'COSTCENTER', type: 'IN_LIST', param: 'CC_IT_01, CC_FIN_02, 12234566, 10180000', min: '', max: '', active: true, desc: 'Must match SAP Master Data Cost Centers' },
      { id: 'RULE-4', col: 'H1', type: 'NUMERIC_RANGE', param: '', min: '0', max: '', active: true, desc: 'Planning amount H1 must be numeric (>= 0)' }
    ];
    const savedRules = this._storageGet('sac_validation_rules_v1');
    this._validationRules = Array.isArray(savedRules) ? savedRules : defaultRules;

    // Estado del formulario de nueva regla
    this._newRuleForm = {
      col: this._requiredColumns[0] || 'ID',
      type: 'IN_LIST',
      param: '',
      min: '',
      max: '',
      desc: ''
    };

    const savedHistory = this._storageGet('sac_audit_history_v2');
    this._history = Array.isArray(savedHistory) ? savedHistory : [];

    this._currentSessionLogs = [];
    this._selectedAuditId = null;
    this._auditDetailSubTab = 'dataTab';
    this._auditSearchQuery = '';

    this._issueFilter = 'all';
    this._searchQuery = '';
    this._toast = null;
    this._toastTimer = null;
    this._showAddColModal = false;
    this._showInlineRuleCreator = false;
    this._bulkCol = this._requiredColumns.includes('COMPANY') ? 'COMPANY' : (this._requiredColumns[0] || 'ID');
    this._bulkVal = '';

    this._sapLogoUri = 'data:image/svg+xml;utf8,' + encodeURIComponent(`
      <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 116 58">
        <defs>
          <linearGradient id="sapGrad" x1="0%" y1="0%" x2="0%" y2="100%">
            <stop offset="0%" stop-color="#00B8F1"/>
            <stop offset="50%" stop-color="#0083CA"/>
            <stop offset="100%" stop-color="#0052A5"/>
          </linearGradient>
        </defs>
        <polygon points="0,0 116,0 58,58 0,58" fill="url(#sapGrad)"/>
        <text x="7" y="42" font-family="Arial Black, Arial, Helvetica, sans-serif" font-weight="900" font-size="33" letter-spacing="-1.5" fill="#FFFFFF">SAP</text>
      </svg>
    `);

    this._ensureXLSX();
    this._render();
  }

  disconnectedCallback() {
    if (this._toastTimer) {
      clearTimeout(this._toastTimer);
      this._toastTimer = null;
    }
  }

  _storageGet(key) {
    try {
      const raw = localStorage.getItem(key);
      return raw ? JSON.parse(raw) : null;
    } catch (e) {
      return null;
    }
  }

  _storageSet(key, val) {
    try {
      if (val === null) localStorage.removeItem(key);
      else localStorage.setItem(key, JSON.stringify(val));
    } catch (e) {}
  }

  _ensureXLSX() {
    if (typeof XLSX !== 'undefined') {
      this._xlsxReady = true;
      return;
    }
    let script = document.querySelector('script[data-sac-xlsx]');
    if (!script) {
      script = document.createElement('script');
      script.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
      script.async = true;
      script.setAttribute('data-sac-xlsx', 'true');
      document.head.appendChild(script);
    }
    script.addEventListener('load', () => {
      this._xlsxReady = true;
      this._render();
    });
  }

  _esc(v) {
    return String(v ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  _logAction(type, target, detail, oldVal = '', newVal = '') {
    this._currentSessionLogs.unshift({
      time: new Date().toLocaleTimeString(),
      type,
      target,
      detail,
      oldVal: String(oldVal ?? ''),
      newVal: String(newVal ?? '')
    });
  }

  _saveHistory() {
    this._storageSet('sac_audit_history_v2', this._history.slice(0, 30));
  }

  _saveRules() {
    this._storageSet('sac_validation_rules_v1', this._validationRules);
  }

  _notify(msg, type = 'info') {
    this._toast = { msg, type };
    this._render();
    if (this._toastTimer) clearTimeout(this._toastTimer);
    this._toastTimer = setTimeout(() => {
      this._toast = null;
      const banner = this._shadow.querySelector('.toast-banner');
      if (banner) banner.remove();
    }, 3800);
  }

  _recordCellModification(row, colName, oldVal, newVal) {
    if (!row._editedCols) row._editedCols = {};
    const initialFrom = row._editedCols[colName] !== undefined
      ? row._editedCols[colName].from
      : oldVal;

    if (initialFrom === newVal) {
      delete row._editedCols[colName];
    } else {
      row._editedCols[colName] = { from: initialFrom, to: newVal };
    }
  }

  // Normalización de valores según configuración de formato numérico (Europeo vs Estándar)
  _normalizeCellValue(val) {
    if (val === undefined || val === null) return '';
    let str = String(val).trim();
    if (this._parseConfig.numberFormat === 'EU' && str !== '') {
      // Detecta patrón europeo tipo 1.234,56 o -123,45
      if (/^-?\d{1,3}(\.\d{3})*(,\d+)$/.test(str) || /^-?\d+(,\d+)$/.test(str)) {
        str = str.replace(/\./g, '').replace(',', '.');
      }
    }
    return str;
  }

  // Reconstruye this._rows combinando filas manuales + las pestañas seleccionadas de cada fichero
  _syncRowsFromActiveSheets() {
    const manualRows = this._rows.filter(r => r._fileOwner === '__MANUAL__');
    const sheetRows = [];

    this._filesQueue.forEach(fileItem => {
      (fileItem.sheets || []).forEach(sh => {
        if (sh.selected) {
          sh.rows.forEach(r => {
            this._requiredColumns.forEach(reqCol => {
              if (r[reqCol] === undefined) r[reqCol] = '';
            });
            sheetRows.push(r);
          });
        }
      });
    });

    this._rows = [...manualRows, ...sheetRows];
    this._rows.forEach((r, idx) => { r._rowNum = idx + 2; });
    this._validateRows();
  }

  _bindEvents() {
    const shadow = this._shadow;

    const fileInput = shadow.getElementById('fileInput');
    const dropZone = shadow.getElementById('dropZone');

    if (fileInput) {
      fileInput.addEventListener('change', (e) => {
        if (e.target.files && e.target.files.length) {
          this._handleFilesSelection(Array.from(e.target.files));
          e.target.value = '';
        }
      });
    }

    if (dropZone) {
      dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('drag-active');
      });
      dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('drag-active');
      });
      dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('drag-active');
        if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length) {
          this._handleFilesSelection(Array.from(e.dataTransfer.files));
        }
      });
    }

    // Configuración de Formato, Tabulador (Delimitador) y Fila de Cabecera
    const cfgFormat = shadow.getElementById('cfgFormat');
    const cfgDelimiter = shadow.getElementById('cfgDelimiter');
    const cfgNumFormat = shadow.getElementById('cfgNumFormat');
    const cfgHeaderRow = shadow.getElementById('cfgHeaderRow');
    const reparseFilesBtn = shadow.getElementById('reparseFilesBtn');

    if (cfgFormat) {
      cfgFormat.addEventListener('change', (e) => {
        this._parseConfig.format = e.target.value;
      });
    }
    if (cfgDelimiter) {
      cfgDelimiter.addEventListener('change', (e) => {
        this._parseConfig.delimiter = e.target.value;
      });
    }
    if (cfgNumFormat) {
      cfgNumFormat.addEventListener('change', (e) => {
        this._parseConfig.numberFormat = e.target.value;
      });
    }
    if (cfgHeaderRow) {
      cfgHeaderRow.addEventListener('change', (e) => {
        this._parseConfig.headerRow = Math.max(1, parseInt(e.target.value, 10) || 1);
      });
    }
    if (reparseFilesBtn) {
      reparseFilesBtn.addEventListener('click', () => {
        this._reparseAllFilesWithNewConfig();
      });
    }

    // Navegación del Stepper
    shadow.querySelectorAll('.step-item').forEach(stepEl => {
      stepEl.addEventListener('click', (e) => {
        const targetStep = parseInt(e.currentTarget.dataset.step, 10);
        if (targetStep === 1 || this._filesQueue.length > 0 || this._rows.length > 0) {
          this._step = targetStep;
          this._render();
        } else {
          this._notify('Please upload at least one Excel/CSV file first.', 'info');
        }
      });
    });

    // Selección / Deselección de Pestañas (Sheets) por cada fichero Excel
    shadow.querySelectorAll('.sheet-toggle-chk').forEach(chk => {
      chk.addEventListener('change', (e) => {
        const fIdx = parseInt(e.currentTarget.dataset.fidx, 10);
        const sIdx = parseInt(e.currentTarget.dataset.sidx, 10);
        const fileItem = this._filesQueue[fIdx];
        if (fileItem && fileItem.sheets && fileItem.sheets[sIdx]) {
          const sh = fileItem.sheets[sIdx];
          sh.selected = e.currentTarget.checked;
          this._logAction(
            sh.selected ? 'SHEET_ENABLED' : 'SHEET_DISABLED',
            `${fileItem.file.name} [${sh.name}]`,
            `${sh.selected ? 'Included' : 'Excluded'} sheet (${sh.rows.length} rows) in consolidation`,
            sh.selected ? 'Excluded' : 'Included',
            sh.selected ? 'Included' : 'Excluded'
          );
          this._syncRowsFromActiveSheets();
          this._render();
        }
      });
    });

    shadow.querySelectorAll('.select-all-sheets-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const fIdx = parseInt(e.currentTarget.dataset.fidx, 10);
        const mode = e.currentTarget.dataset.mode === 'all';
        const fileItem = this._filesQueue[fIdx];
        if (fileItem && fileItem.sheets) {
          fileItem.sheets.forEach(sh => { sh.selected = mode; });
          this._logAction(
            'SHEET_BULK_TOGGLE',
            fileItem.file.name,
            `${mode ? 'Selected all' : 'Deselected all'} sheets in workbook`,
            '',
            mode ? 'ALL_SELECTED' : 'NONE_SELECTED'
          );
          this._syncRowsFromActiveSheets();
          this._render();
        }
      });
    });

    // Botón para vaciar todos los ficheros de la cola
    const clearAllFilesBtn = shadow.getElementById('clearAllFilesBtn');
    if (clearAllFilesBtn) {
      clearAllFilesBtn.addEventListener('click', () => {
        this._logAction('QUEUE_CLEARED', `All Files (${this._filesQueue.length})`, 'Removed all workbooks from queue', `${this._filesQueue.length} files`, '0 files');
        this._filesQueue = [];
        this._rows = this._rows.filter(r => r._fileOwner === '__MANUAL__');
        if (this._rows.length === 0) this._step = 1;
        this._validateRows();
        this._notify('All uploaded files removed from queue.', 'info');
      });
    }

    const downloadTemplateBtn = shadow.getElementById('downloadTemplateBtn');
    if (downloadTemplateBtn) {
      downloadTemplateBtn.addEventListener('click', () => this._downloadTemplate());
    }

    const exportCleanBtn = shadow.getElementById('exportCleanBtn');
    if (exportCleanBtn) {
      exportCleanBtn.addEventListener('click', () => this._exportCleanExcel());
    }

    const nextStep3Btn = shadow.getElementById('nextStep3Btn');
    const prevBtn = shadow.getElementById('prevStepBtn');
    const finishBtn = shadow.getElementById('finishBtn');
    const exitBtn = shadow.getElementById('exitBtn');
    const downloadLogBtn = shadow.getElementById('downloadLogBtn');
    const addRowBtn = shadow.getElementById('addRowBtn');

    if (nextStep3Btn) nextStep3Btn.addEventListener('click', () => { this._step = 3; this._render(); });
    if (prevBtn) prevBtn.addEventListener('click', () => { if (this._step > 1) { this._step--; this._render(); } });
    if (exitBtn) {
      exitBtn.addEventListener('click', () => {
        this._step = 1;
        this._filesQueue = [];
        this._rows = [];
        this._currentSessionLogs = [];
        this._searchQuery = '';
        this._render();
      });
    }

    if (downloadLogBtn) {
      downloadLogBtn.addEventListener('click', () => this._exportErrorLog());
    }

    if (addRowBtn) {
      addRowBtn.addEventListener('click', () => {
        const newId = 'PLN_' + Math.floor(Math.random() * 9000 + 1000);
        const newRow = { _rowNum: 2, _sourceFile: 'Manual Entry', _fileOwner: '__MANUAL__', _isManual: true, _editedCols: {}, _errors: {} };
        this._requiredColumns.forEach(col => {
          newRow[col] = col === 'ID' ? newId : '';
        });
        this._rows.unshift(newRow);
        this._rows.forEach((r, idx) => { r._rowNum = idx + 2; });
        this._logAction('ROW_ADDED', `Row ID ${newId}`, 'Manual planning row inserted directly in grid', '', newId);
        this._validateRows();
        this._notify('New planning row added at the top of the table.', 'success');
      });
    }

    const openAddColBtn = shadow.getElementById('openAddColBtn');
    const confirmAddColBtn = shadow.getElementById('confirmAddColBtn');
    const cancelAddColBtn = shadow.getElementById('cancelAddColBtn');
    const newColInput = shadow.getElementById('newColInput');

    if (openAddColBtn) {
      openAddColBtn.addEventListener('click', () => {
        this._showAddColModal = true;
        this._render();
        const inp = this._shadow.getElementById('newColInput');
        if (inp) inp.focus();
      });
    }

    if (cancelAddColBtn) {
      cancelAddColBtn.addEventListener('click', () => {
        this._showAddColModal = false;
        this._render();
      });
    }

    if (confirmAddColBtn && newColInput) {
      const saveCol = () => {
        const upperCol = newColInput.value.trim().toUpperCase().replace(/\s+/g, '_');
        if (!upperCol) return;
        if (this._requiredColumns.includes(upperCol)) {
          this._notify(`Column '${upperCol}' already exists.`, 'error');
          return;
        }
        this._requiredColumns.push(upperCol);
        this._storageSet('sac_required_cols', this._requiredColumns);
        this._rows.forEach(r => { if (r[upperCol] === undefined) r[upperCol] = ''; });
        this._logAction('SCHEMA_ADD_COL', `Column [${upperCol}]`, 'Added new mandatory column to planning schema', '', upperCol);
        this._showAddColModal = false;
        this._validateRows();
        this._notify(`Column '${upperCol}' added to model.`, 'success');
      };
      confirmAddColBtn.addEventListener('click', saveCol);
      newColInput.addEventListener('keydown', (e) => { if (e.key === 'Enter') saveCol(); });
    }

    shadow.querySelectorAll('.remove-col-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const colToRemove = e.currentTarget.dataset.col;
        if (colToRemove === 'ID') {
          this._notify('The primary ID column cannot be removed.', 'error');
          return;
        }
        this._requiredColumns = this._requiredColumns.filter(c => c !== colToRemove);
        this._storageSet('sac_required_cols', this._requiredColumns);
        if (!this._requiredColumns.includes(this._bulkCol)) {
          this._bulkCol = this._requiredColumns[0] || 'ID';
        }
        this._logAction('SCHEMA_DEL_COL', `Column [${colToRemove}]`, 'Removed column from active planning model', colToRemove, 'DELETED');
        this._validateRows();
        this._notify(`Column '${colToRemove}' removed.`, 'info');
      });
    });

    // Eventos del Motor de Reglas de Validación
    const toggleInlineRuleBtn = shadow.getElementById('toggleInlineRuleBtn');
    if (toggleInlineRuleBtn) {
      toggleInlineRuleBtn.addEventListener('click', () => {
        this._showInlineRuleCreator = !this._showInlineRuleCreator;
        this._render();
      });
    }

    const goToRulesTabBtn = shadow.getElementById('goToRulesTabBtn');
    if (goToRulesTabBtn) {
      goToRulesTabBtn.addEventListener('click', () => {
        this._activeTab = 'rulesTab';
        this._render();
      });
    }

    const ruleColSel = shadow.getElementById('ruleColSel');
    const ruleTypeSel = shadow.getElementById('ruleTypeSel');
    const ruleParamInp = shadow.getElementById('ruleParamInp');
    const ruleMinInp = shadow.getElementById('ruleMinInp');
    const ruleMaxInp = shadow.getElementById('ruleMaxInp');
    const ruleDescInp = shadow.getElementById('ruleDescInp');
    const saveNewRuleBtn = shadow.getElementById('saveNewRuleBtn');
    const resetDefaultRulesBtn = shadow.getElementById('resetDefaultRulesBtn');

    if (ruleColSel) ruleColSel.addEventListener('change', (e) => { this._newRuleForm.col = e.target.value; });
    if (ruleTypeSel) {
      ruleTypeSel.addEventListener('change', (e) => {
        this._newRuleForm.type = e.target.value;
        this._render();
      });
    }
    if (ruleParamInp) ruleParamInp.addEventListener('input', (e) => { this._newRuleForm.param = e.target.value; });
    if (ruleMinInp) ruleMinInp.addEventListener('input', (e) => { this._newRuleForm.min = e.target.value; });
    if (ruleMaxInp) ruleMaxInp.addEventListener('input', (e) => { this._newRuleForm.max = e.target.value; });
    if (ruleDescInp) ruleDescInp.addEventListener('input', (e) => { this._newRuleForm.desc = e.target.value; });

    if (saveNewRuleBtn) {
      saveNewRuleBtn.addEventListener('click', () => this._createNewValidationRule());
    }

    if (resetDefaultRulesBtn) {
      resetDefaultRulesBtn.addEventListener('click', () => {
        this._storageSet('sac_validation_rules_v1', null);
        this._validationRules = [
          { id: 'RULE-1', col: 'ID', type: 'UNIQUE', param: '', min: '', max: '', active: true, desc: 'No duplicate IDs allowed across consolidated sheets' },
          { id: 'RULE-2', col: 'COMPANY', type: 'IN_LIST', param: '1000, 2000, 950', min: '', max: '', active: true, desc: 'Must match SAP Master Data Company Codes' },
          { id: 'RULE-3', col: 'COSTCENTER', type: 'IN_LIST', param: 'CC_IT_01, CC_FIN_02, 12234566, 10180000', min: '', max: '', active: true, desc: 'Must match SAP Master Data Cost Centers' },
          { id: 'RULE-4', col: 'H1', type: 'NUMERIC_RANGE', param: '', min: '0', max: '', active: true, desc: 'Planning amount H1 must be numeric (>= 0)' }
        ];
        this._saveRules();
        this._validateRows();
        this._notify('Default corporate validation rules restored.', 'info');
      });
    }

    shadow.querySelectorAll('.toggle-rule-chk').forEach(chk => {
      chk.addEventListener('change', (e) => {
        const rId = e.currentTarget.dataset.ruleId;
        const rule = this._validationRules.find(r => r.id === rId);
        if (rule) {
          rule.active = e.currentTarget.checked;
          this._saveRules();
          this._logAction('RULE_TOGGLED', `${rule.col} [${rule.type}]`, `Validation rule ${rule.active ? 'enabled' : 'disabled'}`, !rule.active, rule.active);
          this._validateRows();
          this._render();
        }
      });
    });

    shadow.querySelectorAll('.delete-rule-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const rId = e.currentTarget.dataset.ruleId;
        const rule = this._validationRules.find(r => r.id === rId);
        this._validationRules = this._validationRules.filter(r => r.id !== rId);
        this._saveRules();
        if (rule) {
          this._logAction('RULE_DELETED', `${rule.col} [${rule.type}]`, 'Deleted custom validation rule', rule.param || rule.type, 'DELETED');
        }
        this._validateRows();
        this._notify('Validation rule removed.', 'info');
      });
    });

    shadow.querySelectorAll('.remove-row-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const origIdx = parseInt(e.currentTarget.dataset.origIdx, 10);
        if (!isNaN(origIdx) && this._rows[origIdx]) {
          const removed = this._rows[origIdx];
          this._logAction('ROW_DELETED', `Row #${removed._rowNum} (ID: ${removed.ID || 'N/A'})`, `Deleted row from ${removed._sourceFile}`, String(removed.ID || 'Row'), 'DELETED');
          this._rows.splice(origIdx, 1);
          this._rows.forEach((r, idx) => { r._rowNum = idx + 2; });
          this._validateRows();
          this._render();
        }
      });
    });

    shadow.querySelectorAll('.dup-row-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const origIdx = parseInt(e.currentTarget.dataset.origIdx, 10);
        if (!isNaN(origIdx) && this._rows[origIdx]) {
          const source = this._rows[origIdx];
          const oldId = source.ID || 'N/A';
          const newId = oldId ? `${oldId}_COPY` : 'PLN_COPY';
          const clone = {
            ...source,
            ID: newId,
            _fileOwner: '__MANUAL__',
            _isManual: true,
            _editedCols: { ...(source._editedCols || {}) },
            _errors: {}
          };
          this._rows.splice(origIdx + 1, 0, clone);
          this._rows.forEach((r, idx) => { r._rowNum = idx + 2; });
          this._logAction('ROW_CLONED', `Row #${source._rowNum}`, `Cloned planning row from ID '${oldId}'`, oldId, newId);
          this._validateRows();
          this._notify('Row duplicated. Verify the new ID.', 'info');
        }
      });
    });

    const searchInput = shadow.getElementById('searchInput');
    if (searchInput) {
      searchInput.addEventListener('input', (e) => {
        const pos = e.target.selectionStart;
        this._searchQuery = e.target.value;
        this._render();
        const restored = this._shadow.getElementById('searchInput');
        if (restored) {
          restored.focus();
          restored.setSelectionRange(pos, pos);
        }
      });
    }

    const bulkColSelect = shadow.getElementById('bulkColSelect');
    const bulkValInput = shadow.getElementById('bulkValInput');
    const applyBulkEmptyBtn = shadow.getElementById('applyBulkEmptyBtn');
    const applyBulkAllBtn = shadow.getElementById('applyBulkAllBtn');

    if (bulkColSelect) {
      bulkColSelect.addEventListener('change', (e) => { this._bulkCol = e.target.value; });
    }
    if (bulkValInput) {
      bulkValInput.addEventListener('input', (e) => { this._bulkVal = e.target.value; });
    }
    if (applyBulkEmptyBtn) {
      applyBulkEmptyBtn.addEventListener('click', () => this._applyBulkFill(true));
    }
    if (applyBulkAllBtn) {
      applyBulkAllBtn.addEventListener('click', () => this._applyBulkFill(false));
    }

    shadow.querySelectorAll('.file-comment-input').forEach(input => {
      input.addEventListener('input', (e) => {
        const idx = parseInt(e.target.dataset.index, 10);
        if (this._filesQueue[idx]) {
          this._filesQueue[idx].comment = e.target.value;
        }
      });
      input.addEventListener('blur', (e) => {
        const idx = parseInt(e.target.dataset.index, 10);
        if (this._filesQueue[idx]) {
          const prevLogged = this._filesQueue[idx]._lastLoggedComment || '';
          const newComment = (this._filesQueue[idx].comment || '').trim();
          this._filesQueue[idx].comment = newComment;
          if (prevLogged !== newComment) {
            this._filesQueue[idx]._lastLoggedComment = newComment;
            this._logAction('FILE_COMMENT', this._filesQueue[idx].file.name, 'Updated business traceability comment', prevLogged || '(Empty)', newComment || '(Empty)');
          }
        }
      });
    });

    shadow.querySelectorAll('.remove-file-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const idx = parseInt(e.currentTarget.dataset.index, 10);
        const removedFile = this._filesQueue[idx];
        if (removedFile) {
          const fileName = removedFile.file.name;
          this._logAction('FILE_REMOVED', fileName, 'Removed workbook and its sheets from consolidation queue', fileName, 'REMOVED');
          this._filesQueue.splice(idx, 1);
          this._syncRowsFromActiveSheets();
        }
        if (this._filesQueue.length === 0 && this._rows.length === 0) {
          this._step = 1;
        }
        this._render();
      });
    });

    shadow.querySelectorAll('.nav-tab').forEach(tab => {
      tab.addEventListener('click', (e) => {
        this._activeTab = e.currentTarget.dataset.target;
        this._render();
      });
    });

    shadow.querySelectorAll('.chip').forEach(chip => {
      chip.addEventListener('click', (e) => {
        this._issueFilter = e.currentTarget.dataset.filter;
        this._render();
      });
    });

    shadow.querySelectorAll('.editable-cell').forEach(cell => {
      cell.addEventListener('blur', (e) => {
        const origIdx = parseInt(e.target.dataset.origIdx, 10);
        const colName = e.target.dataset.col;
        const newVal = e.target.innerText.trim();
        if (!isNaN(origIdx) && this._rows[origIdx]) {
          const row = this._rows[origIdx];
          const oldVal = String(row[colName] ?? '').trim();
          if (oldVal !== newVal) {
            row[colName] = newVal;
            this._recordCellModification(row, colName, oldVal, newVal);
            this._logAction(
              'CELL_EDIT',
              `Row #${row._rowNum} [${colName}]`,
              `Inline value modification in ${row._sourceFile}`,
              oldVal || '(Blank)',
              newVal || '(Blank)'
            );
            this._validateRows();
            this._render();
          }
        }
      });
      cell.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
          e.preventDefault();
          e.target.blur();
        }
      });
    });

    shadow.querySelectorAll('.audit-row-link').forEach(rowEl => {
      rowEl.addEventListener('click', (e) => {
        this._selectedAuditId = e.currentTarget.dataset.loadId;
        this._auditDetailSubTab = 'dataTab';
        this._auditSearchQuery = '';
        this._render();
      });
    });

    const backToAuditListBtn = shadow.getElementById('backToAuditListBtn');
    if (backToAuditListBtn) {
      backToAuditListBtn.addEventListener('click', () => {
        this._selectedAuditId = null;
        this._render();
      });
    }

    const clearHistoryBtn = shadow.getElementById('clearHistoryBtn');
    if (clearHistoryBtn) {
      clearHistoryBtn.addEventListener('click', () => {
        this._history = [];
        this._selectedAuditId = null;
        this._storageSet('sac_audit_history_v2', null);
        this._notify('Audit history cleared.', 'info');
      });
    }

    shadow.querySelectorAll('.audit-subtab').forEach(subBtn => {
      subBtn.addEventListener('click', (e) => {
        this._auditDetailSubTab = e.currentTarget.dataset.subtab;
        this._render();
      });
    });

    const auditSearchInput = shadow.getElementById('auditSearchInput');
    if (auditSearchInput) {
      auditSearchInput.addEventListener('input', (e) => {
        const pos = e.target.selectionStart;
        this._auditSearchQuery = e.target.value;
        this._render();
        const restored = this._shadow.getElementById('auditSearchInput');
        if (restored) {
          restored.focus();
          restored.setSelectionRange(pos, pos);
        }
      });
    }

    const exportAuditExcelBtn = shadow.getElementById('exportAuditExcelBtn');
    if (exportAuditExcelBtn) {
      exportAuditExcelBtn.addEventListener('click', () => {
        const rec = this._history.find(h => h.id === this._selectedAuditId);
        if (rec) this._exportAuditRecordExcel(rec);
      });
    }

    const exportAuditJsonBtn = shadow.getElementById('exportAuditJsonBtn');
    if (exportAuditJsonBtn) {
      exportAuditJsonBtn.addEventListener('click', () => {
        const rec = this._history.find(h => h.id === this._selectedAuditId);
        if (rec) this._exportAuditRecordJson(rec);
      });
    }

    if (finishBtn) {
      finishBtn.addEventListener('click', () => {
        const errorsCount = this._rows.filter(r => Object.keys(r._errors).length > 0).length;
        if (errorsCount > 0) {
          this._notify(`Cannot consolidate: ${errorsCount} row(s) still contain errors.`, 'error');
          return;
        }
        if (this._rows.length === 0) {
          this._notify('No active data rows to load. Select at least one sheet or add rows.', 'error');
          return;
        }

        const loadId = 'LOAD-' + Math.floor(10000 + Math.random() * 90000);
        const fileSummary = this._filesQueue.length > 0
          ? this._filesQueue.map(f => {
              const activeSheets = (f.sheets || []).filter(s => s.selected).map(s => s.name).join(', ');
              return `${f.file.name} [Sheets: ${activeSheets || 'None'}] (${f.comment || 'No comment'})`;
            }).join('; ')
          : 'Manual Planning Entry';

        const filesMetadata = this._filesQueue.map(f => ({
          name: f.file.name,
          sizeKB: Math.max(1, Math.round(f.file.size / 1024)),
          sheets: (f.sheets || []).filter(s => s.selected).length,
          totalSheets: (f.sheets || []).length,
          selectedSheetNames: (f.sheets || []).filter(s => s.selected).map(s => s.name),
          rows: (f.sheets || []).filter(s => s.selected).reduce((acc, s) => acc + s.rows.length, 0),
          comment: f.comment || 'No business comment provided'
        }));

        const snapshotRows = this._rows.map(r => {
          const copy = {
            _rowNum: r._rowNum,
            _sourceFile: r._sourceFile,
            _isManual: !!r._isManual,
            _editedCols: { ...(r._editedCols || {}) }
          };
          this._requiredColumns.forEach(c => { copy[c] = r[c] ?? ''; });
          return copy;
        });

        const cleanPayloadArray = this._rows.map(r => {
          const clean = {};
          this._requiredColumns.forEach(c => { clean[c] = r[c] ?? ''; });
          return clean;
        });

        const manualEditsCount = this._currentSessionLogs.filter(
          l => l.type === 'CELL_EDIT' || l.type === 'BULK_FILL' || l.type === 'ROW_ADDED' || l.type === 'ROW_CLONED' || l.type === 'ROW_DELETED'
        ).length;

        const auditRecord = {
          id: loadId,
          user: this._currentUser,
          file: fileSummary,
          filesDetails: filesMetadata,
          columns: [...this._requiredColumns],
          rows: this._rows.length,
          editsCount: manualEditsCount,
          timestamp: new Date().toLocaleString(),
          status: 'Synchronized',
          snapshotRows,
          changeLog: [...this._currentSessionLogs],
          payloadJson: JSON.stringify(cleanPayloadArray, null, 2)
        };

        this._history.unshift(auditRecord);
        this._saveHistory();

        this.dispatchEvent(new CustomEvent('propertiesChanged', {
          bubbles: true,
          composed: true,
          detail: { properties: { validatedDataJson: JSON.stringify(cleanPayloadArray) } }
        }));

        this._step = 1;
        this._filesQueue = [];
        this._rows = [];
        this._currentSessionLogs = [];
        this._activeTab = 'monitorTab';
        this._selectedAuditId = loadId;
        this._auditDetailSubTab = 'dataTab';
        this._notify(`Load ${loadId} synchronized! Inspecting full audit trail.`, 'success');
      });
    }
  }

  _createNewValidationRule() {
    const col = this._newRuleForm.col || this._requiredColumns[0] || 'ID';
    const type = this._newRuleForm.type || 'IN_LIST';
    const param = (this._newRuleForm.param || '').trim();
    const min = (this._newRuleForm.min || '').trim();
    const max = (this._newRuleForm.max || '').trim();
    let desc = (this._newRuleForm.desc || '').trim();

    if ((type === 'IN_LIST' || type === 'STARTS_WITH' || type === 'REGEX') && !param) {
      this._notify('Please specify the validation parameter / allowed values.', 'error');
      return;
    }

    if (!desc) {
      if (type === 'IN_LIST') desc = `${col} must be in list: ${param}`;
      if (type === 'NUMERIC_RANGE') desc = `${col} must be numeric ${min !== '' ? '>= ' + min : ''} ${max !== '' ? '<= ' + max : ''}`.trim();
      if (type === 'STARTS_WITH') desc = `${col} must start with '${param}'`;
      if (type === 'REGEX') desc = `${col} must match pattern ${param}`;
      if (type === 'UNIQUE') desc = `${col} must contain unique values (no duplicates)`;
    }

    const newRule = {
      id: 'RULE-' + Math.floor(1000 + Math.random() * 9000),
      col,
      type,
      param,
      min,
      max,
      active: true,
      desc
    };

    this._validationRules.push(newRule);
    this._saveRules();
    this._logAction('RULE_CREATED', `${col} [${type}]`, desc, '', param || `${min}..${max}` || type);
    this._newRuleForm.param = '';
    this._newRuleForm.min = '';
    this._newRuleForm.max = '';
    this._newRuleForm.desc = '';
    this._showInlineRuleCreator = false;
    this._validateRows();
    this._notify(`Validation rule created for column '${col}'.`, 'success');
  }

  _applyBulkFill(onlyEmpty) {
    const col = this._requiredColumns.includes(this._bulkCol)
      ? this._bulkCol
      : (this._requiredColumns[0] || 'ID');
    this._bulkCol = col;
    const val = this._bulkVal.trim();
    if (!col || !val) {
      this._notify('Enter a value to apply in Bulk Fill.', 'error');
      return;
    }
    let updated = 0;
    const targetRows = this._getFilteredRows();
    targetRows.forEach(({ row }) => {
      const current = String(row[col] ?? '').trim();
      if (!onlyEmpty || current === '') {
        if (current !== val) {
          this._recordCellModification(row, col, current || '(Blank)', val);
          row[col] = val;
          updated++;
        }
      }
    });
    if (updated > 0) {
      this._logAction(
        'BULK_FILL',
        `Column [${col}] (${updated} rows)`,
        onlyEmpty ? 'Bulk filled empty cells' : 'Bulk overwrote visible cells',
        onlyEmpty ? '(Empty cells)' : '(Multiple values)',
        val
      );
    }
    this._validateRows();
    this._notify(`Bulk updated ${updated} row(s) in column '${col}'.`, 'success');
  }

  async _handleFilesSelection(fileList) {
    if (typeof XLSX === 'undefined') {
      this._notify('Excel engine is still loading, please wait a second...', 'error');
      return;
    }

    const newItemsToParse = [];
    const duplicateNames = [];

    for (const newFile of fileList) {
      const exists = this._filesQueue.some(item => item.file.name.toLowerCase() === newFile.name.toLowerCase());
      if (exists) {
        duplicateNames.push(newFile.name);
      } else {
        const ext = (newFile.name.split('.').pop() || '').toUpperCase();
        const queueItem = {
          file: newFile,
          ext,
          comment: '',
          _lastLoggedComment: '',
          sheets: [], // [{ name, selected, rows: [] }]
          sheetCount: 0,
          rowCount: 0,
          parsed: false
        };
        this._filesQueue.push(queueItem);
        newItemsToParse.push(queueItem);
        this._logAction('FILE_UPLOAD', newFile.name, `File loaded (${Math.max(1, Math.round(newFile.size / 1024))} KB, Format: ${ext})`, '', 'QUEUED');
      }
    }

    if (newItemsToParse.length > 0) {
      await this._parseFilesInQueue(newItemsToParse);
    }

    if (duplicateNames.length > 0) {
      this._notify(`Skipped duplicate file(s): ${duplicateNames.join(', ')}`, 'error');
    }
  }

  async _reparseAllFilesWithNewConfig() {
    if (this._filesQueue.length === 0) return;
    this._logAction(
      'PARSE_CONFIG_CHANGED',
      `Format: ${this._parseConfig.format}, Delimiter: ${this._parseConfig.delimiter === '\t' ? 'TAB' : this._parseConfig.delimiter}`,
      `Re-parsing workbooks with Header Row #${this._parseConfig.headerRow} and Number Format ${this._parseConfig.numberFormat}`,
      '',
      'REPARSED'
    );
    await this._parseFilesInQueue(this._filesQueue);
    this._notify('All files re-parsed with the new format & delimiter settings.', 'success');
  }

  _readFileAsArrayBuffer(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => resolve(e.target.result);
      reader.onerror = (err) => reject(err);
      reader.readAsArrayBuffer(file);
    });
  }

  _readFileAsText(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => resolve(e.target.result);
      reader.onerror = (err) => reject(err);
      reader.readAsText(file, 'utf-8');
    });
  }

  // Parser determinista para ficheros CSV / TXT con soporte explícito de Tabulador (\t), Punto y Coma (;), Coma (,) o Pipe (|)
  _parseDelimitedTextToAOA(text, delimiterSetting) {
    let sep = delimiterSetting;
    const lines = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n').filter(l => l.trim() !== '');
    if (lines.length === 0) return [];

    if (sep === 'AUTO') {
      const sample = lines[0];
      const counts = {
        '\t': (sample.match(/\t/g) || []).length,
        ';': (sample.match(/;/g) || []).length,
        ',': (sample.match(/,/g) || []).length,
        '|': (sample.match(/\|/g) || []).length
      };
      sep = Object.keys(counts).reduce((a, b) => (counts[a] >= counts[b] ? a : b), ',');
      if (counts[sep] === 0) sep = ',';
    }

    return lines.map(line => {
      const result = [];
      let cur = '';
      let inQuotes = false;
      for (let i = 0; i < line.length; i++) {
        const ch = line[i];
        if (ch === '"') {
          inQuotes = !inQuotes;
        } else if (ch === sep && !inQuotes) {
          result.push(cur.trim());
          cur = '';
        } else {
          cur += ch;
        }
      }
      result.push(cur.trim());
      return result;
    });
  }

  async _parseFilesInQueue(itemsToParse) {
    this._isLoading = true;
    this._loadingText = 'Parsing sheets, formats & delimiters...';
    this._render();

    let detectedHeaders = [...this._headers];
    const headerIdx = Math.max(0, (this._parseConfig.headerRow || 1) - 1);

    for (const item of itemsToParse) {
      try {
        const prevSelectionMap = {};
        (item.sheets || []).forEach(s => { prevSelectionMap[s.name] = s.selected; });
        item.sheets = [];
        let fileRowsCount = 0;

        const isTextFile = ['CSV', 'TXT', 'TSV'].includes(item.ext) || this._parseConfig.format === 'CSV_TXT';

        if (isTextFile) {
          const text = await this._readFileAsText(item.file);
          const aoa = this._parseDelimitedTextToAOA(text, this._parseConfig.delimiter);
          const sheetName = 'Delimited_Data';

          if (aoa.length > headerIdx) {
            const sheetHeaders = aoa[headerIdx].map(h => String(h ?? '').trim().toUpperCase()).filter(Boolean);
            if (detectedHeaders.length === 0 && sheetHeaders.length > 0) {
              detectedHeaders = sheetHeaders;
            }
            const currentHeaders = sheetHeaders.length > 0 ? sheetHeaders : detectedHeaders;
            const rows = aoa.slice(headerIdx + 1).map(row => {
              const obj = {
                _rowNum: 0,
                _sourceFile: `${item.file.name} [${sheetName}]`,
                _fileOwner: item.file.name,
                _sheetOwner: sheetName,
                _isManual: false,
                _editedCols: {},
                _errors: {}
              };
              currentHeaders.forEach((h, i) => {
                obj[h] = this._normalizeCellValue(row[i]);
              });
              this._requiredColumns.forEach(reqCol => {
                if (obj[reqCol] === undefined) obj[reqCol] = '';
              });
              return obj;
            }).filter(r => Object.keys(r).some(k => !k.startsWith('_') && String(r[k]).trim() !== ''));

            item.sheets.push({
              name: sheetName,
              selected: prevSelectionMap[sheetName] !== undefined ? prevSelectionMap[sheetName] : true,
              rows
            });
            fileRowsCount += rows.length;
          }
        } else {
          const buffer = await this._readFileAsArrayBuffer(item.file);
          const data = new Uint8Array(buffer);
          const wb = XLSX.read(data, { type: 'array' });

          wb.SheetNames.forEach(sheetName => {
            const ws = wb.Sheets[sheetName];
            const json = XLSX.utils.sheet_to_json(ws, { header: 1 });

            if (json.length > headerIdx) {
              const sheetHeaders = (json[headerIdx] || []).map(h => String(h ?? '').trim().toUpperCase()).filter(Boolean);
              if (detectedHeaders.length === 0 && sheetHeaders.length > 0) {
                detectedHeaders = sheetHeaders;
              }

              const currentHeaders = sheetHeaders.length > 0 ? sheetHeaders : detectedHeaders;
              const rows = json.slice(headerIdx + 1).map((row) => {
                const sourceLabel = `${item.file.name} [${sheetName}]`;
                const obj = {
                  _rowNum: 0,
                  _sourceFile: sourceLabel,
                  _fileOwner: item.file.name,
                  _sheetOwner: sheetName,
                  _isManual: false,
                  _editedCols: {},
                  _errors: {}
                };
                currentHeaders.forEach((h, i) => {
                  obj[h] = this._normalizeCellValue(row[i]);
                });
                this._requiredColumns.forEach(reqCol => {
                  if (obj[reqCol] === undefined) obj[reqCol] = '';
                });
                return obj;
              }).filter(r => Object.keys(r).some(k => !k.startsWith('_') && String(r[k]).trim() !== ''));

              item.sheets.push({
                name: sheetName,
                selected: prevSelectionMap[sheetName] !== undefined ? prevSelectionMap[sheetName] : true,
                rows
              });
              fileRowsCount += rows.length;
            } else {
              item.sheets.push({
                name: sheetName,
                selected: false,
                rows: []
              });
            }
          });
        }

        item.sheetCount = item.sheets.length;
        item.rowCount = fileRowsCount;
        item.parsed = true;
      } catch (err) {
        console.error(`Error parsing ${item.file.name}:`, err);
        this._notify(`Error reading ${item.file.name}`, 'error');
      }
    }

    this._headers = detectedHeaders.length > 0 ? detectedHeaders : this._requiredColumns;
    this._syncRowsFromActiveSheets();
    if (this._step === 1 && this._filesQueue.length > 0) this._step = 2;
    this._isLoading = false;
    this._render();
  }

  _validateRows() {
    const activeRules = this._validationRules.filter(r => r.active);

    // Pre-calcular conteos para reglas de tipo UNIQUE
    const uniqueCols = activeRules.filter(r => r.type === 'UNIQUE').map(r => r.col);
    const valueCountsPerCol = {};
    uniqueCols.forEach(col => {
      valueCountsPerCol[col] = {};
      this._rows.forEach(r => {
        const val = String(r[col] ?? '').trim();
        if (val) valueCountsPerCol[col][val] = (valueCountsPerCol[col][val] || 0) + 1;
      });
    });

    this._rows.forEach(row => {
      row._errors = {};

      // 1. Campos obligatorios del esquema
      this._requiredColumns.forEach(col => {
        if (row[col] === undefined || String(row[col]).trim() === '') {
          row._errors[col] = 'Mandatory field';
        }
      });

      // 2. Evaluación dinámica de todas las reglas creadas por el usuario
      activeRules.forEach(rule => {
        const col = rule.col;
        if (!this._requiredColumns.includes(col)) return;
        const rawVal = String(row[col] ?? '').trim();
        if (!rawVal) return; // Si está vacío, ya salta Mandatory field

        if (rule.type === 'UNIQUE') {
          if (valueCountsPerCol[col] && valueCountsPerCol[col][rawVal] > 1) {
            row._errors[col] = 'Duplicate value';
          }
        } else if (rule.type === 'IN_LIST') {
          const allowed = rule.param.split(',').map(s => s.trim()).filter(Boolean);
          if (allowed.length > 0 && !allowed.includes(rawVal)) {
            row._errors[col] = `Allowed: ${allowed.join(', ')}`;
          }
        } else if (rule.type === 'NUMERIC_RANGE') {
          const num = Number(rawVal);
          if (isNaN(num)) {
            row._errors[col] = 'Must be numeric';
          } else {
            if (rule.min !== '' && !isNaN(Number(rule.min)) && num < Number(rule.min)) {
              row._errors[col] = `Min allowed: ${rule.min}`;
            }
            if (rule.max !== '' && !isNaN(Number(rule.max)) && num > Number(rule.max)) {
              row._errors[col] = `Max allowed: ${rule.max}`;
            }
          }
        } else if (rule.type === 'STARTS_WITH') {
          if (!rawVal.toUpperCase().startsWith(rule.param.toUpperCase())) {
            row._errors[col] = `Must start with '${rule.param}'`;
          }
        } else if (rule.type === 'REGEX') {
          try {
            const re = new RegExp(rule.param);
            if (!re.test(rawVal)) {
              row._errors[col] = `Invalid format (${rule.param})`;
            }
          } catch (e) {}
        }
      });
    });
  }

  _downloadTemplate() {
    if (typeof XLSX === 'undefined') return;
    const sampleRow = this._requiredColumns.map(col => {
      if (col === 'ID') return 'PLN_1001';
      if (col === 'COMPANY') return '1000';
      if (col === 'COSTCENTER') return 'CC_IT_01';
      if (col === 'H1') return '25000';
      return 'SAMPLE';
    });
    const ws = XLSX.utils.aoa_to_sheet([this._requiredColumns, sampleRow]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Planning_Template');
    XLSX.writeFile(wb, 'SAC_Planning_Template.xlsx');
    this._notify('Official Excel template downloaded.', 'success');
  }

  _exportCleanExcel() {
    if (typeof XLSX === 'undefined' || this._rows.length === 0) return;
    const header = ['SOURCE_FILE', ...this._requiredColumns];
    const data = this._rows.map(r => [
      r._sourceFile,
      ...this._requiredColumns.map(c => r[c] ?? '')
    ]);
    const ws = XLSX.utils.aoa_to_sheet([header, ...data]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Consolidated_Data');
    XLSX.writeFile(wb, 'SAC_Consolidated_Planning.xlsx');
    this._notify('Consolidated data exported to Excel.', 'success');
  }

  _exportAuditRecordExcel(rec) {
    if (typeof XLSX === 'undefined' || !rec || !rec.snapshotRows) return;
    const cols = rec.columns || this._requiredColumns;
    const header = ['ROW_NUM', 'SOURCE_WORKBOOK_SHEET', 'ORIGIN_TYPE', ...cols];
    const data = rec.snapshotRows.map(r => [
      r._rowNum,
      r._sourceFile,
      r._isManual ? 'MANUAL / CLONED' : (Object.keys(r._editedCols || {}).length > 0 ? 'EDITED IN WIDGET' : 'ORIGINAL EXCEL'),
      ...cols.map(c => r[c] ?? '')
    ]);
    const ws = XLSX.utils.aoa_to_sheet([header, ...data]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Committed_Snapshot');

    if (rec.changeLog && rec.changeLog.length > 0) {
      const logHeader = ['TIME', 'ACTION_TYPE', 'TARGET', 'DESCRIPTION', 'OLD_VALUE', 'NEW_VALUE'];
      const logData = rec.changeLog.map(l => [l.time, l.type, l.target, l.detail, l.oldVal, l.newVal]);
      const wsLog = XLSX.utils.aoa_to_sheet([logHeader, ...logData]);
      XLSX.utils.book_append_sheet(wb, wsLog, 'Audit_Change_Log');
    }

    XLSX.writeFile(wb, `${rec.id}_SAC_Audit_Snapshot.xlsx`);
    this._notify(`Audit snapshot ${rec.id} exported to Excel.`, 'success');
  }

  _exportAuditRecordJson(rec) {
    if (!rec) return;
    const blob = new Blob([JSON.stringify(rec, null, 2)], { type: 'application/json;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${rec.id}_BTP_Audit_Certificate.json`;
    a.click();
    URL.revokeObjectURL(url);
    this._notify(`Audit certificate ${rec.id} downloaded.`, 'success');
  }

  _exportErrorLog() {
    const errorRows = this._rows.filter(r => Object.keys(r._errors).length > 0);
    if (errorRows.length === 0) return;
    let csv = 'SourceFile,Row,ID,Column,Error Description\n';
    errorRows.forEach(r => {
      Object.keys(r._errors).forEach(col => {
        csv += `"${this._esc(r._sourceFile)}",${r._rowNum},"${this._esc(r.ID || '')}","${col}","${this._esc(r._errors[col])}"\n`;
      });
    });
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'SAC_Validation_Error_Log.csv';
    a.click();
    URL.revokeObjectURL(url);
  }

  _getFilteredRows() {
    const q = this._searchQuery.trim().toLowerCase();
    return this._rows
      .map((row, origIdx) => ({ row, origIdx }))
      .filter(({ row }) => {
        const hasErr = Object.keys(row._errors).length > 0;
        if (this._issueFilter === 'correct' && hasErr) return false;
        if (this._issueFilter === 'errors' && !hasErr) return false;
        if (q) {
          const matchSource = String(row._sourceFile ?? '').toLowerCase().includes(q);
          const matchCols = this._requiredColumns.some(col =>
            String(row[col] ?? '').toLowerCase().includes(q)
          );
          return matchSource || matchCols;
        }
        return true;
      });
  }

  _renderRuleCreatorForm() {
    const t = this._newRuleForm.type;
    return `
      <div style="background:#F8FAFC; border:1px solid #B3D4FF; border-radius:8px; padding:14px; margin-top:10px;">
        <div style="font-weight:700; color:#005CB9; margin-bottom:10px; font-size:12px;">➕ Create New Column Validation Rule</div>
        <div style="display:flex; flex-wrap:wrap; gap:10px; align-items:flex-end;">
          <div>
            <label style="display:block; font-size:10px; font-weight:700; color:#556B82; margin-bottom:3px;">TARGET COLUMN</label>
            <select id="ruleColSel" class="input-std">
              ${this._requiredColumns.map(c => `<option value="${this._esc(c)}" ${this._newRuleForm.col === c ? 'selected' : ''}>${this._esc(c)}</option>`).join('')}
            </select>
          </div>
          <div>
            <label style="display:block; font-size:10px; font-weight:700; color:#556B82; margin-bottom:3px;">RULE TYPE</label>
            <select id="ruleTypeSel" class="input-std">
              <option value="IN_LIST" ${t === 'IN_LIST' ? 'selected' : ''}>Allowed List / Master Data (IN_LIST)</option>
              <option value="NUMERIC_RANGE" ${t === 'NUMERIC_RANGE' ? 'selected' : ''}>Numeric & Min/Max Range</option>
              <option value="STARTS_WITH" ${t === 'STARTS_WITH' ? 'selected' : ''}>Starts With Prefix</option>
              <option value="REGEX" ${t === 'REGEX' ? 'selected' : ''}>Custom Regex Pattern</option>
              <option value="UNIQUE" ${t === 'UNIQUE' ? 'selected' : ''}>Unique Values (No Duplicates)</option>
            </select>
          </div>

          ${t === 'IN_LIST' || t === 'STARTS_WITH' || t === 'REGEX' ? `
            <div style="flex:1; min-width:200px;">
              <label style="display:block; font-size:10px; font-weight:700; color:#556B82; margin-bottom:3px;">
                ${t === 'IN_LIST' ? 'ALLOWED VALUES (COMMA SEPARATED)' : (t === 'STARTS_WITH' ? 'REQUIRED PREFIX (E.G. PLN_)' : 'REGEX PATTERN (E.G. ^[A-Z0-9]+$)')}
              </label>
              <input type="text" id="ruleParamInp" class="input-std" style="width:100%; box-sizing:border-box;" placeholder="${t === 'IN_LIST' ? 'e.g. 1000, 2000, 3000' : (t === 'STARTS_WITH' ? 'e.g. CC_' : 'e.g. ^[0-9]{4}$')}" value="${this._esc(this._newRuleForm.param)}" />
            </div>
          ` : ''}

          ${t === 'NUMERIC_RANGE' ? `
            <div>
              <label style="display:block; font-size:10px; font-weight:700; color:#556B82; margin-bottom:3px;">MIN VALUE (OPTIONAL)</label>
              <input type="number" id="ruleMinInp" class="input-std" style="width:110px;" placeholder="e.g. 0" value="${this._esc(this._newRuleForm.min)}" />
            </div>
            <div>
              <label style="display:block; font-size:10px; font-weight:700; color:#556B82; margin-bottom:3px;">MAX VALUE (OPTIONAL)</label>
              <input type="number" id="ruleMaxInp" class="input-std" style="width:110px;" placeholder="e.g. 999999" value="${this._esc(this._newRuleForm.max)}" />
            </div>
          ` : ''}

          <div style="flex:1; min-width:180px;">
            <label style="display:block; font-size:10px; font-weight:700; color:#556B82; margin-bottom:3px;">CUSTOM MESSAGE / DESCRIPTION</label>
            <input type="text" id="ruleDescInp" class="input-std" style="width:100%; box-sizing:border-box;" placeholder="Optional description..." value="${this._esc(this._newRuleForm.desc)}" />
          </div>

          <div>
            <button class="btn" id="saveNewRuleBtn" type="button">+ Save & Apply Rule</button>
          </div>
        </div>
      </div>
    `;
  }

  _renderValidationRulesTab() {
    return `
      <div>
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:14px; flex-wrap:wrap; gap:10px;">
          <div>
            <h3 style="margin:0; color:#1D2D3E;">⚙️ Corporate Validation Rules & Master Data Engine</h3>
            <p style="margin:4px 0 0 0; color:#556B82; font-size:12px;">Create, toggle, or remove real-time validation rules applied across all consolidated Excel sheets.</p>
          </div>
          <button class="btn btn-ghost" id="resetDefaultRulesBtn" type="button">↺ Restore Default Corporate Rules</button>
        </div>

        ${this._renderRuleCreatorForm()}

        <div style="margin-top:18px; border:1px solid #D9E0E6; border-radius:8px; overflow:hidden; background:#FFFFFF;">
          <table>
            <thead>
              <tr>
                <th style="width:60px; text-align:center;">Active</th>
                <th style="width:130px;">Column</th>
                <th style="width:150px;">Rule Type</th>
                <th>Validation Parameters / Master Data</th>
                <th>Business Description</th>
                <th style="width:80px; text-align:center;">Action</th>
              </tr>
            </thead>
            <tbody>
              ${this._validationRules.length > 0 ? this._validationRules.map(r => `
                <tr style="${!r.active ? 'opacity:0.55; background:#F8FAFC;' : ''}">
                  <td style="text-align:center;">
                    <input type="checkbox" class="toggle-rule-chk" data-rule-id="${this._esc(r.id)}" ${r.active ? 'checked' : ''} style="cursor:pointer; width:15px; height:15px;" />
                  </td>
                  <td><span class="col-pill">${this._esc(r.col)}</span></td>
                  <td><strong style="color:#1D2D3E; font-size:11px;">${this._esc(r.type)}</strong></td>
                  <td style="font-family:monospace; font-size:11px; color:#005CB9; font-weight:700;">
                    ${r.type === 'NUMERIC_RANGE'
                      ? `Min: ${r.min !== '' ? this._esc(r.min) : '-∞'} | Max: ${r.max !== '' ? this._esc(r.max) : '+∞'}`
                      : (r.param ? this._esc(r.param) : 'All rows in table')}
                  </td>
                  <td style="color:#556B82;">${this._esc(r.desc)}</td>
                  <td style="text-align:center;">
                    <button class="icon-btn icon-btn-del delete-rule-btn" data-rule-id="${this._esc(r.id)}" type="button" title="Delete rule">✕</button>
                  </td>
                </tr>
              `).join('') : `
                <tr><td colspan="6" style="text-align:center; padding:28px; color:#556B82;">No custom validation rules configured. Use the form above to add one.</td></tr>
              `}
            </tbody>
          </table>
        </div>
      </div>
    `;
  }

  _renderFormatConfigBar(showReparseBtn = false) {
    return `
      <div class="toolbar-box" style="background:#F8FAFC; border:1px solid #B3D4FF;">
        <div style="display:flex; align-items:center; gap:14px; flex-wrap:wrap;">
          <span style="font-weight:800; color:#005CB9; font-size:12px;">⚙️ File Format & Tabulador Settings:</span>
          <div>
            <label style="font-size:10px; font-weight:700; color:#556B82; margin-right:4px;">FORMAT:</label>
            <select id="cfgFormat" class="input-std">
              <option value="AUTO" ${this._parseConfig.format === 'AUTO' ? 'selected' : ''}>Auto-Detect (.xlsx, .xls, .csv, .txt)</option>
              <option value="XLSX" ${this._parseConfig.format === 'XLSX' ? 'selected' : ''}>Excel Workbook (.xlsx / .xls / .xlsm)</option>
              <option value="CSV_TXT" ${this._parseConfig.format === 'CSV_TXT' ? 'selected' : ''}>Flat File / CSV / Text (.csv / .txt)</option>
            </select>
          </div>
          <div>
            <label style="font-size:10px; font-weight:700; color:#556B82; margin-right:4px;">TABULADOR / DELIMITER:</label>
            <select id="cfgDelimiter" class="input-std">
              <option value="AUTO" ${this._parseConfig.delimiter === 'AUTO' ? 'selected' : ''}>Auto-Detect</option>
              <option value="&#9;" ${this._parseConfig.delimiter === '\t' ? 'selected' : ''}>Tabulador (TAB \t)</option>
              <option value=";" ${this._parseConfig.delimiter === ';' ? 'selected' : ''}>Semicolon (;)</option>
              <option value="," ${this._parseConfig.delimiter === ',' ? 'selected' : ''}>Comma (,)</option>
              <option value="|" ${this._parseConfig.delimiter === '|' ? 'selected' : ''}>Pipe (|)</option>
            </select>
          </div>
          <div>
            <label style="font-size:10px; font-weight:700; color:#556B82; margin-right:4px;">NUMBER FORMAT:</label>
            <select id="cfgNumFormat" class="input-std">
              <option value="STD" ${this._parseConfig.numberFormat === 'STD' ? 'selected' : ''}>Standard (1,234.56)</option>
              <option value="EU" ${this._parseConfig.numberFormat === 'EU' ? 'selected' : ''}>European (1.234,56 → 1234.56)</option>
            </select>
          </div>
          <div>
            <label style="font-size:10px; font-weight:700; color:#556B82; margin-right:4px;">HEADER ROW:</label>
            <input type="number" id="cfgHeaderRow" class="input-std" style="width:52px;" min="1" max="20" value="${this._parseConfig.headerRow}" />
          </div>
        </div>
        ${showReparseBtn ? `
          <button class="btn btn-ghost" id="reparseFilesBtn" type="button" style="font-size:11px; padding:6px 12px;">↻ Apply & Re-Parse Files</button>
        ` : ''}
      </div>
    `;
  }

  _renderAuditMonitor() {
    if (this._selectedAuditId) {
      const rec = this._history.find(h => h.id === this._selectedAuditId);
      if (!rec) {
        this._selectedAuditId = null;
      } else {
        const cols = rec.columns || this._requiredColumns;
        const snapshotRows = rec.snapshotRows || [];
        const changeLog = rec.changeLog || [];
        const filesDetails = rec.filesDetails || [];
        const q = this._auditSearchQuery.trim().toLowerCase();

        const filteredSnapshot = snapshotRows.filter(r => {
          if (!q) return true;
          const matchSrc = String(r._sourceFile ?? '').toLowerCase().includes(q);
          const matchCol = cols.some(c => String(r[c] ?? '').toLowerCase().includes(q));
          return matchSrc || matchCol;
        });

        const modifiedRowsCount = snapshotRows.filter(
          r => r._isManual || Object.keys(r._editedCols || {}).length > 0
        ).length;

        return `
          <div class="audit-detail-view">
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:16px; flex-wrap:wrap; gap:10px;">
              <div style="display:flex; align-items:center; gap:10px;">
                <button class="btn btn-ghost" id="backToAuditListBtn" type="button">← Back to Audit Log</button>
                <div>
                  <div style="font-size:11px; color:#556B82; font-weight:600;">BTP Audit Monitor  /  Execution Drill-Down</div>
                  <div style="font-size:16px; font-weight:800; color:#1D2D3E;">Load ID: ${this._esc(rec.id)} <span style="background:#E6F9EC; color:#107E3E; padding:3px 10px; border-radius:12px; font-size:11px; margin-left:6px;">✓ ${this._esc(rec.status)}</span></div>
                </div>
              </div>
              <div style="display:flex; gap:8px;">
                <button class="btn btn-ghost" id="exportAuditExcelBtn" type="button">📥 Download Snapshot (.xlsx)</button>
                <button class="btn btn-ghost" id="exportAuditJsonBtn" type="button">📜 Audit Certificate (.json)</button>
              </div>
            </div>

            <div class="metrics-bar">
              <div class="metric-box"><div class="metric-title">Executed By</div><div class="metric-val" style="font-size:13px;">${this._esc(rec.user)}</div></div>
              <div class="metric-box"><div class="metric-title">Timestamp</div><div class="metric-val" style="font-size:13px;">${this._esc(rec.timestamp)}</div></div>
              <div class="metric-box kpi-ok"><div class="metric-title">Committed Rows</div><div class="metric-val" style="color:#107E3E;">${rec.rows}</div></div>
              <div class="metric-box"><div class="metric-title">Rows Modified / Added</div><div class="metric-val" style="color:#0064D9;">${modifiedRowsCount}</div></div>
              <div class="metric-box"><div class="metric-title">Tracked User Actions</div><div class="metric-val">${changeLog.length} events</div></div>
            </div>

            <div class="subtabs-bar">
              <button class="audit-subtab ${this._auditDetailSubTab === 'dataTab' ? 'active' : ''}" data-subtab="dataTab" type="button">
                📊 Consolidated Data Snapshot (${snapshotRows.length})
              </button>
              <button class="audit-subtab ${this._auditDetailSubTab === 'changesTab' ? 'active' : ''}" data-subtab="changesTab" type="button">
                🛠️ User Modifications & Activity Log (${changeLog.length})
              </button>
              <button class="audit-subtab ${this._auditDetailSubTab === 'filesTab' ? 'active' : ''}" data-subtab="filesTab" type="button">
                📁 Source Workbooks & BTP Payload (${filesDetails.length})
              </button>
            </div>

            ${this._auditDetailSubTab === 'dataTab' ? `
              <div class="toolbar-box" style="margin-bottom:10px;">
                <div style="display:flex; align-items:center; gap:8px;">
                  <span style="font-weight:700; font-size:12px;">🔍 Inspect Committed Rows:</span>
                  <input type="text" id="auditSearchInput" class="input-std" style="width:240px;" placeholder="Search ID, cost center, sheet..." value="${this._esc(this._auditSearchQuery)}" />
                </div>
                <div style="display:flex; gap:12px; font-size:11px; font-weight:600; color:#556B82;">
                  <span><span style="display:inline-block; width:10px; height:10px; background:#FFF8E6; border:1px solid #F59E0B; border-radius:2px; margin-right:4px;"></span>✎ Edited in Widget</span>
                  <span><span style="display:inline-block; width:10px; height:10px; background:#EBF5FF; border:1px solid #0070F2; border-radius:2px; margin-right:4px;"></span>★ Added / Cloned Row</span>
                </div>
              </div>

              <div class="tbl-wrap">
                <table>
                  <thead>
                    <tr>
                      <th style="width:32px;">#</th>
                      <th style="width:120px;">Traceability State</th>
                      <th>Source Workbook & Sheet</th>
                      ${cols.map(c => `<th>${this._esc(c)}</th>`).join('')}
                    </tr>
                  </thead>
                  <tbody>
                    ${filteredSnapshot.length > 0 ? filteredSnapshot.map(r => {
                      const editedCols = r._editedCols || {};
                      const hasEdits = Object.keys(editedCols).length > 0;
                      return `
                        <tr style="${r._isManual ? 'background:#F5F9FF;' : ''}">
                          <td style="color:#556B82; font-weight:600; font-size:11px;">${r._rowNum}</td>
                          <td>
                            ${r._isManual
                              ? `<span style="background:#EBF5FF; color:#005CB9; border:1px solid #B3D4FF; padding:2px 8px; border-radius:10px; font-size:10px; font-weight:700;">★ New / Cloned</span>`
                              : (hasEdits
                                ? `<span style="background:#FFF8E6; color:#B45309; border:1px solid #FCD34D; padding:2px 8px; border-radius:10px; font-size:10px; font-weight:700;">✎ Edited (${Object.keys(editedCols).length})</span>`
                                : `<span style="background:#F0F3F6; color:#556B82; padding:2px 8px; border-radius:10px; font-size:10px; font-weight:600;">Original Excel</span>`)}
                          </td>
                          <td style="font-size:11px; color:#0064D9; font-weight:700;">${this._esc(r._sourceFile)}</td>
                          ${cols.map(col => {
                            const editInfo = editedCols[col];
                            return `
                              <td style="${editInfo ? 'background:#FFFBEB; border-left:3px solid #F59E0B;' : ''}">
                                <div style="font-weight:${editInfo ? '700' : '400'}; color:#1D2D3E;">${this._esc(r[col] ?? '')}</div>
                                ${editInfo ? `<div style="font-size:10px; color:#B45309; margin-top:2px;">Was: <strike>${this._esc(editInfo.from)}</strike></div>` : ''}
                              </td>
                            `;
                          }).join('')}
                        </tr>
                      `;
                    }).join('') : `
                      <tr><td colspan="${cols.length + 3}" style="text-align:center; padding:24px; color:#556B82;">No rows match the search filter.</td></tr>
                    `}
                  </tbody>
                </table>
              </div>
            ` : ''}

            ${this._auditDetailSubTab === 'changesTab' ? `
              <div class="tbl-wrap">
                <table>
                  <thead>
                    <tr>
                      <th style="width:90px;">Time</th>
                      <th style="width:140px;">Action Type</th>
                      <th style="width:180px;">Target Element</th>
                      <th>Audit Description</th>
                      <th style="width:140px;">Previous Value</th>
                      <th style="width:140px;">Committed Value</th>
                    </tr>
                  </thead>
                  <tbody>
                    ${changeLog.length > 0 ? changeLog.map(log => {
                      let badgeStyle = 'background:#EBF5FF; color:#005CB9;';
                      if (log.type === 'CELL_EDIT' || log.type === 'BULK_FILL') badgeStyle = 'background:#FFF8E6; color:#B45309;';
                      if (log.type === 'ROW_DELETED' || log.type === 'FILE_REMOVED' || log.type === 'SCHEMA_DEL_COL' || log.type === 'SHEET_DISABLED') badgeStyle = 'background:#FEEAEA; color:#BB0000;';
                      if (log.type === 'ROW_ADDED' || log.type === 'ROW_CLONED' || log.type === 'SHEET_ENABLED') badgeStyle = 'background:#E6F9EC; color:#107E3E;';
                      return `
                        <tr>
                          <td style="font-size:11px; color:#556B82; font-weight:600;">${this._esc(log.time)}</td>
                          <td><span style="${badgeStyle} padding:3px 8px; border-radius:6px; font-weight:700; font-size:10px;">${this._esc(log.type)}</span></td>
                          <td style="font-weight:700; color:#1D2D3E;">${this._esc(log.target)}</td>
                          <td style="color:#32363A;">${this._esc(log.detail)}</td>
                          <td style="color:#BB0000; font-family:monospace; font-size:11px;">${this._esc(log.oldVal || '—')}</td>
                          <td style="color:#107E3E; font-weight:700; font-family:monospace; font-size:11px;">${this._esc(log.newVal || '—')}</td>
                        </tr>
                      `;
                    }).join('') : `
                      <tr><td colspan="6" style="text-align:center; padding:28px; color:#556B82;">No manual edits or modifications were required during this load.</td></tr>
                    `}
                  </tbody>
                </table>
              </div>
            ` : ''}

            ${this._auditDetailSubTab === 'filesTab' ? `
              <div style="display:grid; grid-template-columns: 1fr 1fr; gap:16px;">
                <div style="background:#FFFFFF; border:1px solid #D9E0E6; border-radius:8px; padding:16px;">
                  <h4 style="margin:0 0 12px 0; color:#1D2D3E;">Consolidated Source Workbooks</h4>
                  ${filesDetails.length > 0 ? filesDetails.map(f => `
                    <div class="file-card-item" style="margin-bottom:8px;">
                      <div>
                        <div style="font-weight:700; color:#0064D9;">📄 ${this._esc(f.name)} <span style="font-weight:400; color:#556B82;">(${f.sizeKB || 1} KB)</span></div>
                        <div style="font-size:11px; color:#556B82; margin-top:2px;">
                          ${f.sheets} of ${f.totalSheets || f.sheets} sheet(s) selected (${this._esc((f.selectedSheetNames || []).join(', '))}) • ${f.rows} rows committed
                        </div>
                        <div style="font-size:12px; color:#1D2D3E; margin-top:6px; background:#F8FAFC; padding:6px 10px; border-radius:6px; border-left:3px solid #0070F2;">
                          💬 <strong>Business Comment:</strong> ${this._esc(f.comment)}
                        </div>
                      </div>
                    </div>
                  `).join('') : `<p style="color:#556B82;">Manual entry load without Excel workbooks.</p>`}
                </div>

                <div style="background:#FFFFFF; border:1px solid #D9E0E6; border-radius:8px; padding:16px;">
                  <h4 style="margin:0 0 12px 0; color:#1D2D3E;">Synchronized SAC JSON Payload</h4>
                  <pre style="background:#132337; color:#74D8FF; padding:12px; border-radius:6px; font-size:11px; max-height:220px; overflow:auto; margin:0;">${this._esc(rec.payloadJson || '[]')}</pre>
                </div>
              </div>
            ` : ''}
          </div>
        `;
      }
    }

    return `
      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:14px;">
        <div>
          <h3 style="margin:0; color:#1D2D3E;">BTP Consolidation & Audit Monitor</h3>
          <p style="margin:4px 0 0 0; color:#556B82; font-size:12px;">Click on any execution row below to inspect its committed rows, cell-by-cell modifications, and source workbooks.</p>
        </div>
        ${this._history.length > 0 ? `<button class="btn btn-danger" id="clearHistoryBtn" type="button" style="font-size:11px; padding:6px 12px;">🗑️ Clear History</button>` : ''}
      </div>

      <div style="border: 1px solid #D9E0E6; border-radius: 8px; overflow:hidden; background:#FFFFFF; box-shadow: 0 2px 6px rgba(0,0,0,0.02);">
        <table>
          <thead>
            <tr>
              <th>Load ID</th>
              <th>SAP User</th>
              <th>Workbooks, Selected Sheets & Comments</th>
              <th>Committed Rows</th>
              <th>User Edits</th>
              <th>Timestamp</th>
              <th>BTP Status</th>
              <th style="text-align:right;">Inspection</th>
            </tr>
          </thead>
          <tbody>
            ${this._history.length > 0 ? this._history.map(h => `
              <tr class="audit-row-link" data-load-id="${this._esc(h.id)}" style="cursor:pointer;">
                <td><strong style="color:#0064D9; text-decoration:underline;">${this._esc(h.id)}</strong></td>
                <td>${this._esc(h.user)}</td>
                <td>${this._esc(h.file)}</td>
                <td><strong>${h.rows}</strong> rows</td>
                <td>
                  ${(h.editsCount || 0) > 0
                    ? `<span style="background:#FFF8E6; color:#B45309; border:1px solid #FCD34D; padding:2px 8px; border-radius:10px; font-weight:700; font-size:11px;">✎ ${h.editsCount} change(s)</span>`
                    : `<span style="color:#556B82; font-size:11px;">Direct (0 edits)</span>`}
                </td>
                <td>${this._esc(h.timestamp)}</td>
                <td><span style="background:#E6F9EC; color:#107E3E; padding:4px 10px; border-radius:12px; font-weight:700; font-size:11px;">✓ ${this._esc(h.status)}</span></td>
                <td style="text-align:right;">
                  <button class="btn btn-ghost" type="button" style="padding:4px 10px; font-size:11px;">🔍 View Details →</button>
                </td>
              </tr>
            `).join('') : `
              <tr><td colspan="8" style="text-align:center; padding:36px; color:#556B82;">No planning loads registered yet. Complete an upload in the Wizard to inspect its full audit trail here.</td></tr>
            `}
          </tbody>
        </table>
      </div>
    `;
  }

  _render() {
    const totalRows = this._rows.length;
    const errorRows = this._rows.filter(r => Object.keys(r._errors).length > 0).length;
    const validRows = totalRows - errorRows;
    const validPercent = totalRows > 0 ? Math.round((validRows / totalRows) * 100) : 100;
    const filteredRows = this._getFilteredRows();
    const activeRulesCount = this._validationRules.filter(r => r.active).length;

    this._shadow.innerHTML = `
      <style>
        :host {
          display: block;
          font-family: "72", "72full", -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif;
          font-size: 13px;
          color: #1D2D3E;
          background: #F5F6F7;
          padding: 16px;
          box-sizing: border-box;
          height: 100%;
          overflow: auto;
        }

        .main-card {
          background: #FFFFFF;
          border: 1px solid #D9E0E6;
          border-radius: 12px;
          box-shadow: 0 8px 24px rgba(29, 45, 62, 0.07);
          overflow: hidden;
          position: relative;
        }

        .busy-overlay {
          position: absolute;
          inset: 0;
          background: rgba(255, 255, 255, 0.9);
          backdrop-filter: blur(2px);
          display: flex;
          flex-direction: column;
          align-items: center;
          justify-content: center;
          z-index: 100;
        }
        .spinner {
          width: 40px;
          height: 40px;
          border: 4px solid #E5E9EC;
          border-top-color: #0070F2;
          border-radius: 50%;
          animation: spin 0.75s linear infinite;
          margin-bottom: 12px;
        }
        @keyframes spin { to { transform: rotate(360deg); } }

        .sap-header {
          display: flex;
          align-items: center;
          justify-content: space-between;
          background: linear-gradient(90deg, #132337 0%, #1F3652 100%);
          color: #FFFFFF;
          padding: 12px 24px;
          border-bottom: 3px solid #0070F2;
        }
        .sap-brand {
          display: flex;
          align-items: center;
          gap: 14px;
        }
        .sap-logo-img {
          height: 28px;
          width: auto;
          display: block;
          filter: drop-shadow(0 2px 4px rgba(0, 0, 0, 0.25));
        }
        .sap-title-wrap {
          display: flex;
          flex-direction: column;
        }
        .sap-title {
          font-size: 15px;
          font-weight: 700;
          letter-spacing: -0.2px;
          color: #FFFFFF;
        }
        .sap-subtitle {
          font-size: 11px;
          color: #A9B8C8;
          font-weight: 500;
        }
        .sap-badge-env {
          background: rgba(0, 184, 241, 0.16);
          border: 1px solid rgba(0, 184, 241, 0.45);
          color: #74D8FF;
          padding: 4px 10px;
          border-radius: 20px;
          font-size: 11px;
          font-weight: 600;
        }

        .toast-banner {
          display: flex;
          align-items: center;
          justify-content: space-between;
          padding: 10px 24px;
          font-weight: 600;
          font-size: 12px;
          color: #FFFFFF;
        }
        .toast-info { background: #0070F2; }
        .toast-success { background: #107E3E; }
        .toast-error { background: #BB0000; }

        .nav-tabs {
          display: flex;
          background: #FFFFFF;
          border-bottom: 1px solid #D9E0E6;
          padding: 0 24px;
          gap: 8px;
        }
        .nav-tab {
          padding: 13px 18px;
          cursor: pointer;
          font-weight: 600;
          color: #556B82;
          border-bottom: 3px solid transparent;
          font-size: 13px;
          transition: all 0.2s;
        }
        .nav-tab:hover { color: #0070F2; }
        .nav-tab.active {
          color: #0070F2;
          border-bottom-color: #0070F2;
          background: linear-gradient(180deg, transparent 0%, rgba(0,112,242,0.04) 100%);
        }

        .tab-content {
          padding: 24px;
          background: #F5F6F7;
        }

        .subtabs-bar {
          display: flex;
          gap: 8px;
          margin-bottom: 14px;
          border-bottom: 1px solid #D9E0E6;
          padding-bottom: 10px;
        }
        .audit-subtab {
          padding: 8px 14px;
          border-radius: 6px;
          border: 1px solid #D9E0E6;
          background: #FFFFFF;
          color: #556B82;
          font-weight: 700;
          font-size: 12px;
          cursor: pointer;
          transition: all 0.15s;
        }
        .audit-subtab:hover {
          border-color: #0070F2;
          color: #0070F2;
        }
        .audit-subtab.active {
          background: #0070F2;
          color: #FFFFFF;
          border-color: #0070F2;
          box-shadow: 0 2px 6px rgba(0, 112, 242, 0.2);
        }

        .wizard-steps {
          display: grid;
          grid-template-columns: repeat(3, 1fr);
          gap: 12px;
          margin-bottom: 20px;
        }
        .step-item {
          display: flex;
          align-items: center;
          gap: 12px;
          background: #FFFFFF;
          border: 1px solid #D9E0E6;
          border-radius: 10px;
          padding: 12px 16px;
          cursor: pointer;
          transition: all 0.2s ease;
        }
        .step-item:hover {
          border-color: #0070F2;
          transform: translateY(-1px);
          box-shadow: 0 4px 12px rgba(0, 112, 242, 0.08);
        }
        .step-item.active {
          border-color: #0070F2;
          background: #F0F6FF;
          box-shadow: 0 2px 8px rgba(0, 112, 242, 0.12);
        }
        .step-item.completed {
          border-color: #107E3E;
        }
        .step-num {
          width: 26px;
          height: 26px;
          border-radius: 50%;
          background: #E5E9EC;
          color: #556B82;
          display: flex;
          align-items: center;
          justify-content: center;
          font-size: 12px;
          font-weight: 700;
          flex-shrink: 0;
        }
        .step-item.active .step-num {
          background: #0070F2;
          color: #FFFFFF;
          box-shadow: 0 0 0 3px rgba(0, 112, 242, 0.2);
        }
        .step-item.completed .step-num {
          background: #107E3E;
          color: #FFFFFF;
        }
        .step-label-title {
          font-weight: 700;
          font-size: 12px;
          color: #1D2D3E;
        }
        .step-label-sub {
          font-size: 11px;
          color: #556B82;
        }

        .btn {
          padding: 8px 16px;
          font-size: 12px;
          font-weight: 600;
          border-radius: 6px;
          cursor: pointer;
          border: 1px solid #0070F2;
          background: #0070F2;
          color: #FFFFFF;
          transition: all 0.15s ease;
          display: inline-flex;
          align-items: center;
          gap: 6px;
        }
        .btn:hover { background: #005CB9; border-color: #005CB9; }
        .btn-ghost {
          background: #FFFFFF;
          color: #0064D9;
          border-color: #BCC8D4;
        }
        .btn-ghost:hover {
          background: #F0F6FF;
          border-color: #0070F2;
        }
        .btn-danger {
          background: #FFFFFF;
          color: #BB0000;
          border-color: #F5C6C6;
        }
        .btn-danger:hover {
          background: #FEEAEA;
          border-color: #BB0000;
        }

        .drop-zone {
          border: 2px dashed #0070F2;
          border-radius: 12px;
          padding: 36px 24px;
          text-align: center;
          background: #FFFFFF;
          display: flex;
          flex-direction: column;
          align-items: center;
          gap: 12px;
          transition: all 0.2s ease;
          box-shadow: 0 2px 8px rgba(0,0,0,0.02);
        }
        .drop-zone.drag-active {
          background: #F0F6FF;
          border-color: #005CB9;
          transform: scale(1.005);
        }
        .upload-icon-circle {
          width: 54px;
          height: 54px;
          border-radius: 50%;
          background: #EBF5FF;
          color: #0070F2;
          display: flex;
          align-items: center;
          justify-content: center;
          font-size: 24px;
          margin-bottom: 4px;
        }
        .file-label-btn {
          padding: 9px 20px;
          background: #0070F2;
          color: #FFFFFF;
          font-weight: 700;
          font-size: 13px;
          border-radius: 6px;
          cursor: pointer;
          box-shadow: 0 2px 6px rgba(0, 112, 242, 0.25);
          transition: background 0.2s;
          display: inline-block;
        }
        .file-label-btn:hover { background: #005CB9; }

        .metrics-bar {
          display: grid;
          grid-template-columns: repeat(5, 1fr);
          gap: 12px;
          margin-bottom: 14px;
        }
        .metric-box {
          background: #FFFFFF;
          border: 1px solid #D9E0E6;
          border-left: 4px solid #0070F2;
          padding: 12px 14px;
          border-radius: 8px;
          box-shadow: 0 2px 4px rgba(0,0,0,0.02);
        }
        .metric-box.kpi-ok { border-left-color: #107E3E; }
        .metric-box.kpi-err { border-left-color: #BB0000; }
        .metric-title {
          font-size: 10px;
          color: #556B82;
          margin-bottom: 4px;
          font-weight: 700;
          text-transform: uppercase;
          letter-spacing: 0.4px;
        }
        .metric-val {
          font-size: 18px;
          font-weight: 800;
          color: #1D2D3E;
        }

        .dist-bar {
          display: flex;
          height: 8px;
          border-radius: 4px;
          overflow: hidden;
          background: #FEEAEA;
          margin-bottom: 16px;
          border: 1px solid #D9E0E6;
        }
        .dist-ok {
          background: linear-gradient(90deg, #107E3E, #16A34A);
          height: 100%;
          transition: width 0.35s ease;
        }

        .toolbar-box {
          background: #FFFFFF;
          border: 1px solid #D9E0E6;
          border-radius: 8px;
          padding: 10px 14px;
          margin-bottom: 12px;
          display: flex;
          flex-wrap: wrap;
          gap: 12px;
          align-items: center;
          justify-content: space-between;
        }
        .input-std {
          border: 1px solid #BCC8D4;
          border-radius: 6px;
          padding: 6px 10px;
          font-size: 12px;
          color: #1D2D3E;
          outline: none;
          transition: border 0.15s;
        }
        .input-std:focus {
          border-color: #0070F2;
          box-shadow: 0 0 0 2px rgba(0, 112, 242, 0.15);
        }

        .chips-row {
          display: flex;
          flex-wrap: wrap;
          gap: 10px;
          margin-bottom: 12px;
          align-items: center;
          justify-content: space-between;
        }
        .chips-left { display: flex; gap: 6px; align-items: center; }
        .chip {
          padding: 5px 14px;
          border-radius: 16px;
          font-size: 11px;
          font-weight: 700;
          cursor: pointer;
          border: 1px solid #D9E0E6;
          background: #FFFFFF;
          color: #32363A;
          transition: all 0.15s;
        }
        .chip:hover { border-color: #0070F2; }
        .chip-all.active-chip { background: #1D2D3E; color: #FFFFFF; border-color: #1D2D3E; }
        .chip-ok.active-chip { background: #107E3E; color: #FFFFFF; border-color: #107E3E; }
        .chip-err.active-chip { background: #BB0000; color: #FFFFFF; border-color: #BB0000; }

        .tbl-wrap {
          max-height: 360px;
          overflow: auto;
          border: 1px solid #D9E0E6;
          border-radius: 8px;
          background: #FFFFFF;
          box-shadow: 0 2px 6px rgba(0,0,0,0.02);
        }
        table {
          width: 100%;
          border-collapse: collapse;
          font-size: 12px;
          text-align: left;
        }
        th {
          background: #F0F3F6;
          padding: 10px 12px;
          border-bottom: 2px solid #D9E0E6;
          position: sticky;
          top: 0;
          font-weight: 700;
          color: #1D2D3E;
          white-space: nowrap;
          z-index: 5;
        }
        tbody tr:hover { background: #F0F6FF; }
        td {
          padding: 8px 12px;
          border-bottom: 1px solid #EDF0F2;
          vertical-align: top;
        }
        td.cell-error { background: #FFF2F2; }
        .editable-cell {
          outline: none;
          border: 1px solid transparent;
          border-bottom: 1px dashed #0070F2;
          display: inline-block;
          min-width: 65px;
          padding: 3px 6px;
          border-radius: 4px;
          transition: all 0.15s;
        }
        .editable-cell:hover {
          background: #F0F6FF;
          border-color: #B3D1FF;
        }
        .editable-cell:focus {
          background: #FFFFFF;
          border: 1px solid #0070F2;
          box-shadow: 0 0 0 2px rgba(0, 112, 242, 0.2);
        }

        .th-content {
          display: flex;
          align-items: center;
          justify-content: space-between;
          gap: 6px;
        }
        .remove-col-btn {
          background: none;
          border: none;
          color: #BB0000;
          font-weight: 700;
          cursor: pointer;
          font-size: 11px;
          padding: 1px 4px;
          border-radius: 3px;
        }
        .remove-col-btn:hover { background: #FEEAEA; }

        .row-actions { display: flex; gap: 4px; justify-content: center; }
        .icon-btn {
          border: 1px solid #D9E0E6;
          background: #FFFFFF;
          border-radius: 4px;
          cursor: pointer;
          font-size: 11px;
          padding: 3px 7px;
          font-weight: 700;
          transition: all 0.15s;
        }
        .icon-btn:hover { border-color: #0070F2; color: #0070F2; background: #F0F6FF; }
        .icon-btn-del { border-color: #F5C6C6; color: #BB0000; background: #FEEAEA; }
        .icon-btn-del:hover { background: #FCD5D5; }

        .wizard-footer {
          display: flex;
          justify-content: space-between;
          margin-top: 20px;
          border-top: 1px solid #D9E0E6;
          padding-top: 16px;
        }
        .file-card-item {
          background: #FFFFFF;
          border: 1px solid #D9E0E6;
          border-left: 4px solid #0070F2;
          border-radius: 8px;
          padding: 14px 16px;
          margin-bottom: 12px;
        }
        .sheet-pill-chk {
          display: inline-flex;
          align-items: center;
          gap: 6px;
          padding: 4px 10px;
          border-radius: 6px;
          font-size: 11px;
          font-weight: 600;
          cursor: pointer;
          border: 1px solid #BCC8D4;
          background: #F8FAFC;
          color: #32363A;
          user-select: none;
        }
        .sheet-pill-chk.sheet-active {
          background: #EBF5FF;
          border-color: #0070F2;
          color: #005CB9;
        }
        .col-pill {
          display: inline-flex;
          align-items: center;
          gap: 6px;
          background: #EBF5FF;
          color: #005CB9;
          padding: 5px 12px;
          border-radius: 6px;
          font-weight: 700;
          font-size: 11px;
          border: 1px solid #B3D4FF;
        }
      </style>

      <div class="main-card">
        ${this._isLoading ? `
          <div class="busy-overlay">
            <div class="spinner"></div>
            <div style="font-weight:700; color:#1D2D3E;">${this._esc(this._loadingText)}</div>
          </div>
        ` : ''}

        <div class="sap-header">
          <div class="sap-brand">
            <img class="sap-logo-img" src="${this._sapLogoUri}" alt="SAP Logo" />
            <div class="sap-title-wrap">
              <div class="sap-title">Excel Planning Uploader Pro</div>
              <div class="sap-subtitle">Multi-Workbook, Sheet Selector & Custom Validation Rules Engine</div>
            </div>
          </div>
          <div class="sap-badge-env">SAP Analytics Cloud • BTP</div>
        </div>

        ${this._toast ? `
          <div class="toast-banner toast-${this._toast.type}">
            <span>${this._esc(this._toast.msg)}</span>
          </div>
        ` : ''}

        <div class="nav-tabs">
          <div class="nav-tab ${this._activeTab === 'uploaderTab' ? 'active' : ''}" data-target="uploaderTab">Planning Upload Wizard</div>
          <div class="nav-tab ${this._activeTab === 'rulesTab' ? 'active' : ''}" data-target="rulesTab">⚙️ Validation Rules (${activeRulesCount})</div>
          <div class="nav-tab ${this._activeTab === 'monitorTab' ? 'active' : ''}" data-target="monitorTab">BTP Audit Monitor (${this._history.length})</div>
        </div>

        <div id="uploaderTab" class="tab-content" style="display:${this._activeTab === 'uploaderTab' ? 'block' : 'none'};">
          <div class="wizard-steps">
            <div class="step-item ${this._step === 1 ? 'active' : (this._filesQueue.length > 0 ? 'completed' : '')}" data-step="1">
              <div class="step-num">${this._filesQueue.length > 0 && this._step > 1 ? '✓' : '1'}</div>
              <div>
                <div class="step-label-title">Select Files & Format</div>
                <div class="step-label-sub">Multi-Excel, CSV & Tabulador</div>
              </div>
            </div>
            <div class="step-item ${this._step === 2 ? 'active' : (this._step > 2 ? 'completed' : '')}" data-step="2">
              <div class="step-num">${this._step > 2 ? '✓' : '2'}</div>
              <div>
                <div class="step-label-title">Sheets & Validation Rules</div>
                <div class="step-label-sub">Pick tabs, columns & rules</div>
              </div>
            </div>
            <div class="step-item ${this._step === 3 ? 'active' : ''}" data-step="3">
              <div class="step-num">3</div>
              <div>
                <div class="step-label-title">Interactive Grid & Sync</div>
                <div class="step-label-sub">Inline edits, bulk fill & load</div>
              </div>
            </div>
          </div>

          ${this._step === 1 ? `
            ${this._renderFormatConfigBar(false)}
            <div class="drop-zone" id="dropZone">
              <div class="upload-icon-circle">☁️</div>
              <h3 style="margin:0; color:#1D2D3E; font-size:16px;">Drag & drop two or more Excel / CSV files here</h3>
              <p style="margin:0; color:#556B82; font-size:12px; max-width:560px;">
                Select multiple files simultaneously or add them incrementally. You will be able to choose which sheets (pestañas) to consolidate in the next step (.xlsx, .xls, .xlsm, .csv, .txt).
              </p>
              <div style="margin-top:8px;">
                <label class="file-label-btn" for="fileInput">
                  📂 Browse & Select Multiple Files
                </label>
                <input type="file" id="fileInput" multiple accept=".xlsx, .xls, .xlsm, .csv, .txt" ${!this._xlsxReady ? 'disabled' : ''} style="display:none;" />
              </div>
              <div style="margin-top:6px;">
                <button class="btn btn-ghost" id="downloadTemplateBtn" type="button" ${!this._xlsxReady ? 'disabled' : ''}>
                  📥 Download Official Planning Template (.xlsx)
                </button>
              </div>
            </div>
          ` : ''}

          ${this._step === 2 ? `
            <div>
              ${this._renderFormatConfigBar(true)}

              <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px; flex-wrap:wrap; gap:8px;">
                <div>
                  <h3 style="margin:0; color:#1D2D3E;">Loaded Workbooks & Sheet Selector (${this._filesQueue.length} file(s))</h3>
                  <span style="font-size:11px; color:#556B82;">Check or uncheck the sheets (pestañas) inside each Excel file that you want to include in the consolidation:</span>
                </div>
                <div style="display:flex; gap:8px;">
                  <label class="btn" for="fileInput" style="cursor:pointer;">
                    + Add More Excel / CSV Files
                    <input type="file" id="fileInput" multiple accept=".xlsx, .xls, .xlsm, .csv, .txt" style="display:none;" />
                  </label>
                  ${this._filesQueue.length > 0 ? `
                    <button class="btn btn-danger" id="clearAllFilesBtn" type="button">🗑️ Remove All Files</button>
                  ` : ''}
                </div>
              </div>

              <div style="margin-bottom:20px;">
                ${this._filesQueue.map((item, fIdx) => {
                  const activeSheetCount = (item.sheets || []).filter(s => s.selected).length;
                  const activeRowCount = (item.sheets || []).filter(s => s.selected).reduce((acc, s) => acc + s.rows.length, 0);
                  return `
                    <div class="file-card-item">
                      <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px;">
                        <div>
                          <div style="font-weight:800; color:#0064D9; font-size:13px;">
                            📄 ${this._esc(item.file.name)}
                            <span style="background:#EBF5FF; color:#005CB9; font-size:10px; padding:2px 6px; border-radius:4px; margin-left:6px;">${this._esc(item.ext)}</span>
                          </div>
                          <div style="font-size:11px; color:#556B82; margin-top:2px;">
                            <strong>${activeSheetCount}</strong> of ${item.sheetCount} sheet(s) selected • <strong>${activeRowCount}</strong> active planning rows
                          </div>
                        </div>
                        <div style="display:flex; align-items:center; gap:8px;">
                          <input type="text" class="input-std file-comment-input" style="width:250px;" data-index="${fIdx}" placeholder="Business comment for audit log..." value="${this._esc(item.comment || '')}" />
                          <button class="btn btn-danger remove-file-btn" data-index="${fIdx}" type="button" style="padding:6px 12px;">Remove File</button>
                        </div>
                      </div>

                      <!-- SELECTOR DE PESTAÑAS (SHEETS) DEL EXCEL -->
                      <div style="margin-top:10px; padding-top:10px; border-top:1px dashed #D9E0E6; display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:8px;">
                        <div style="display:flex; align-items:center; flex-wrap:wrap; gap:6px;">
                          <span style="font-size:11px; font-weight:700; color:#1D2D3E; margin-right:4px;">📑 Excel Sheets:</span>
                          ${(item.sheets || []).map((sh, sIdx) => `
                            <label class="sheet-pill-chk ${sh.selected ? 'sheet-active' : ''}">
                              <input type="checkbox" class="sheet-toggle-chk" data-fidx="${fIdx}" data-sidx="${sIdx}" ${sh.selected ? 'checked' : ''} style="margin:0; cursor:pointer;" />
                              <span>${this._esc(sh.name)} <strong>(${sh.rows.length})</strong></span>
                            </label>
                          `).join('')}
                        </div>
                        <div style="display:flex; gap:4px;">
                          <button class="icon-btn select-all-sheets-btn" data-fidx="${fIdx}" data-mode="all" type="button">Select All Tabs</button>
                          <button class="icon-btn select-all-sheets-btn" data-fidx="${fIdx}" data-mode="none" type="button">None</button>
                        </div>
                      </div>
                    </div>
                  `;
                }).join('')}
              </div>

              <!-- ESQUEMA DE COLUMNAS Y REGLAS DE VALIDACIÓN EN PASO 2 -->
              <div style="display:grid; grid-template-columns: 1fr 1fr; gap:14px;">
                <div style="background:#FFFFFF; border:1px solid #D9E0E6; border-radius:8px; padding:16px;">
                  <div style="font-weight:700; color:#1D2D3E; margin-bottom:10px;">Active Model Columns (Required Schema):</div>
                  <div style="display:flex; flex-wrap:wrap; gap:8px; margin-bottom:14px;">
                    ${this._requiredColumns.map(c => `
                      <span class="col-pill">
                        ${this._esc(c)}
                        ${c !== 'ID' ? `<button class="remove-col-btn" data-col="${this._esc(c)}" title="Remove column">✕</button>` : ''}
                      </span>
                    `).join('')}
                  </div>

                  ${this._showAddColModal ? `
                    <div style="display:flex; gap:8px; align-items:center;">
                      <input type="text" id="newColInput" class="input-std" placeholder="NEW_COLUMN_NAME" />
                      <button class="btn" id="confirmAddColBtn" type="button">Save Column</button>
                      <button class="btn btn-ghost" id="cancelAddColBtn" type="button">Cancel</button>
                    </div>
                  ` : `
                    <button class="btn btn-ghost" id="openAddColBtn" type="button">+ Add Mandatory Column</button>
                  `}
                </div>

                <div style="background:#FFFFFF; border:1px solid #D9E0E6; border-radius:8px; padding:16px;">
                  <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px;">
                    <div style="font-weight:700; color:#1D2D3E;">⚙️ Active Validation Rules (${activeRulesCount}):</div>
                    <div style="display:flex; gap:6px;">
                      <button class="btn btn-ghost" id="toggleInlineRuleBtn" type="button" style="padding:4px 10px; font-size:11px;">
                        ${this._showInlineRuleCreator ? 'Close Form' : '+ New Rule'}
                      </button>
                      <button class="btn btn-ghost" id="goToRulesTabBtn" type="button" style="padding:4px 10px; font-size:11px;">
                        Manage All Rules →
                      </button>
                    </div>
                  </div>
                  <div style="display:flex; flex-wrap:wrap; gap:6px;">
                    ${this._validationRules.map(r => `
                      <span style="display:inline-flex; align-items:center; gap:6px; background:${r.active ? '#E6F9EC' : '#F0F3F6'}; color:${r.active ? '#107E3E' : '#556B82'}; border:1px solid ${r.active ? '#A7F3D0' : '#D9E0E6'}; padding:4px 10px; border-radius:6px; font-size:11px; font-weight:600;">
                        <input type="checkbox" class="toggle-rule-chk" data-rule-id="${this._esc(r.id)}" ${r.active ? 'checked' : ''} style="margin:0; cursor:pointer;" />
                        <strong>${this._esc(r.col)}</strong>: ${this._esc(r.type)} ${r.param ? `(${this._esc(r.param)})` : ''}
                        <button class="remove-col-btn delete-rule-btn" data-rule-id="${this._esc(r.id)}" title="Delete rule">✕</button>
                      </span>
                    `).join('')}
                  </div>
                </div>
              </div>

              ${this._showInlineRuleCreator ? this._renderRuleCreatorForm() : ''}
            </div>
            <div class="wizard-footer">
              <button class="btn btn-ghost" id="prevStepBtn" type="button">← Back to Upload</button>
              <button class="btn" id="nextStep3Btn" type="button">Validate & Edit ${totalRows} Consolidated Rows →</button>
            </div>
          ` : ''}

          ${this._step === 3 ? `
            <div>
              <div class="metrics-bar">
                <div class="metric-box"><div class="metric-title">Workbooks</div><div class="metric-val">${this._filesQueue.length}</div></div>
                <div class="metric-box"><div class="metric-title">Total Rows</div><div class="metric-val">${totalRows}</div></div>
                <div class="metric-box kpi-ok"><div class="metric-title">Valid Rows</div><div class="metric-val" style="color:#107E3E;">${validRows}</div></div>
                <div class="metric-box kpi-err"><div class="metric-title">With Errors</div><div class="metric-val" style="color:#BB0000;">${errorRows}</div></div>
                <div class="metric-box ${errorRows > 0 ? 'kpi-err' : 'kpi-ok'}"><div class="metric-title">Readiness</div><div class="metric-val" style="font-size:14px; color:${errorRows > 0 ? '#BB0000' : '#107E3E'}">${validPercent}% (${errorRows > 0 ? 'Fix Errors' : 'Ready'})</div></div>
              </div>

              <div class="dist-bar">
                <div class="dist-ok" style="width: ${validPercent}%;"></div>
              </div>

              <div class="toolbar-box">
                <div style="display:flex; align-items:center; gap:8px;">
                  <span style="font-weight:700; font-size:12px;">🔍 Quick Search:</span>
                  <input type="text" id="searchInput" class="input-std" style="width:200px;" placeholder="Filter by ID, text, sheet..." value="${this._esc(this._searchQuery)}" />
                </div>

                <div style="display:flex; align-items:center; gap:6px; flex-wrap:wrap;">
                  <span style="font-weight:700; font-size:12px;">⚡ Bulk Fill:</span>
                  <select id="bulkColSelect" class="input-std">
                    ${this._requiredColumns.map(c => `<option value="${this._esc(c)}" ${this._bulkCol === c ? 'selected' : ''}>${this._esc(c)}</option>`).join('')}
                  </select>
                  <input type="text" id="bulkValInput" class="input-std" style="width:120px;" placeholder="Value (e.g. 1000)" value="${this._esc(this._bulkVal)}" />
                  <button class="btn btn-ghost" id="applyBulkEmptyBtn" type="button" style="padding:5px 10px; font-size:11px;">Fill Empty</button>
                  <button class="btn btn-ghost" id="applyBulkAllBtn" type="button" style="padding:5px 10px; font-size:11px;">Overwrite Visible</button>
                </div>
              </div>

              ${this._showAddColModal ? `
                <div class="toolbar-box" style="justify-content:flex-start;">
                  <span style="font-weight:700;">New Column Name:</span>
                  <input type="text" id="newColInput" class="input-std" placeholder="E.g. CURRENCY" />
                  <button class="btn" id="confirmAddColBtn" type="button">Add Column</button>
                  <button class="btn btn-ghost" id="cancelAddColBtn" type="button">Cancel</button>
                </div>
              ` : ''}

              ${this._showInlineRuleCreator ? this._renderRuleCreatorForm() : ''}

              <div class="chips-row">
                <div class="chips-left">
                  <div class="chip chip-all ${this._issueFilter === 'all' ? 'active-chip' : ''}" data-filter="all">All Rows (${totalRows})</div>
                  <div class="chip chip-ok ${this._issueFilter === 'correct' ? 'active-chip' : ''}" data-filter="correct">Valid (${validRows})</div>
                  <div class="chip chip-err ${this._issueFilter === 'errors' ? 'active-chip' : ''}" data-filter="errors">With Errors (${errorRows})</div>
                </div>
                <div style="display:flex; gap:6px; align-items:center; flex-wrap:wrap;">
                  <button class="btn btn-ghost" id="addRowBtn" type="button" style="font-size:11px; padding:6px 12px;">+ Add Row</button>
                  <button class="btn btn-ghost" id="openAddColBtn" type="button" style="font-size:11px; padding:6px 12px;">+ Add Column</button>
                  <button class="btn btn-ghost" id="toggleInlineRuleBtn" type="button" style="font-size:11px; padding:6px 12px;">⚙️ + Validation Rule</button>
                  <button class="btn btn-ghost" id="exportCleanBtn" type="button" style="font-size:11px; padding:6px 12px;">📤 Export Excel (.xlsx)</button>
                  ${errorRows > 0 ? `<button class="btn btn-danger" id="downloadLogBtn" type="button" style="font-size:11px; padding:6px 12px;">📥 Error Log (.csv)</button>` : ''}
                </div>
              </div>

              <div class="tbl-wrap">
                <table>
                  <thead>
                    <tr>
                      <th style="width:28px;">#</th>
                      <th style="width:76px; text-align:center;">Actions</th>
                      <th>Source Workbook & Sheet</th>
                      ${this._requiredColumns.map(h => `
                        <th>
                          <div class="th-content">
                            <span>${this._esc(h)}</span>
                            ${h !== 'ID' ? `<button class="remove-col-btn" data-col="${this._esc(h)}" title="Delete column">✕</button>` : ''}
                          </div>
                        </th>
                      `).join('')}
                    </tr>
                  </thead>
                  <tbody>
                    ${filteredRows.map(({ row, origIdx }) => `
                      <tr>
                        <td style="color:#556B82; font-size:11px; font-weight:600;">${row._rowNum}</td>
                        <td>
                          <div class="row-actions">
                            <button class="icon-btn dup-row-btn" data-orig-idx="${origIdx}" title="Clone row" type="button">⧉</button>
                            <button class="icon-btn icon-btn-del remove-row-btn" data-orig-idx="${origIdx}" title="Delete row" type="button">✕</button>
                          </div>
                        </td>
                        <td style="font-size:11px; color:#0064D9; font-weight:700;">${this._esc(row._sourceFile)}</td>
                        ${this._requiredColumns.map(col => {
                          const err = row._errors[col];
                          const editedInfo = row._editedCols && row._editedCols[col];
                          return `
                            <td class="${err ? 'cell-error' : ''}" style="${!err && editedInfo ? 'background:#FFFBEB;' : ''}">
                              <span class="editable-cell" contenteditable="true" data-orig-idx="${origIdx}" data-col="${this._esc(col)}">${this._esc(row[col] ?? '')}</span>
                              ${err ? `<span style="display:block; font-size:10px; color:#BB0000; font-weight:700; margin-top:2px;">⚠️ ${this._esc(err)}</span>` : ''}
                              ${!err && editedInfo ? `<span style="display:block; font-size:9px; color:#B45309; font-weight:600;">✎ Edited (was: ${this._esc(editedInfo.from || 'Blank')})</span>` : ''}
                            </td>
                          `;
                        }).join('')}
                      </tr>
                    `).join('')}
                  </tbody>
                </table>
              </div>
            </div>
            <div class="wizard-footer">
              <button class="btn btn-ghost" id="prevStepBtn" type="button">← Back</button>
              <div style="display:flex; gap:10px;">
                <button class="btn btn-ghost" id="exitBtn" type="button">Reset Queue</button>
                <button class="btn" id="finishBtn" type="button" ${errorRows > 0 ? 'style="opacity:0.55;"' : ''}>🚀 Consolidate & Synchronize with SAC</button>
              </div>
            </div>
          ` : ''}
        </div>

        <div id="rulesTab" class="tab-content" style="display:${this._activeTab === 'rulesTab' ? 'block' : 'none'};">
          ${this._renderValidationRulesTab()}
        </div>

        <div id="monitorTab" class="tab-content" style="display:${this._activeTab === 'monitorTab' ? 'block' : 'none'};">
          ${this._renderAuditMonitor()}
        </div>
      </div>
    `;

    this._bindEvents();
  }
}

if (!customElements.get('custom-excel-uploader-pro')) {
  customElements.define('custom-excel-uploader-pro', ExcelUploaderPro);
}
