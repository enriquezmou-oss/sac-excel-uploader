class ExcelUploaderPro extends HTMLElement {
  constructor() {
    super();
    this._shadow = this.attachShadow({ mode: 'open' });
    this._step = 1;
    this._activeTab = 'uploaderTab';
    this._filesQueue = [];
    this._rows = [];
    this._headers = [];
    this._currentUser = 'SAP User (Planning)';
    this._isLoading = false;
    this._loadingText = 'Processing files...';
    this._xlsxReady = typeof XLSX !== 'undefined';
    this._showHelpModal = false;

    // Configuración de Carga y Formato
    this._loadStrategy = 'FULL'; // 'FULL' | 'DELTA'
    this._parseConfig = {
      format: 'AUTO',
      delimiter: 'AUTO',
      numberFormat: 'STD',
      headerRow: 1
    };

    const defaultCols = ['ID', 'DESCRIPTION', 'H1', 'COMPANY', 'COSTCENTER', 'ASSET_CLASS', 'CAPITALIZED'];
    const savedCols = this._storageGet('sac_required_cols');
    this._requiredColumns = Array.isArray(savedCols) && savedCols.length > 0 ? savedCols : defaultCols;

    // Configuración de exportación del Audit Monitor
    this._auditExportConfig = {
      open: false,
      numCols: this._requiredColumns.length,
      selectedCols: [...this._requiredColumns],
      includeComments: true,
      includeAuditDetail: true,
      format: 'XLSX'
    };

    // Reglas con semáforo, alcance temporal y archivos objetivo ('ALL' o nombres específicos)
    const defaultRules = [
      {
        id: 'RULE-1',
        col: 'ID',
        severity: 'BLOCKING',
        semaphore: 'RED',
        type: 'UNIQUE',
        param: '',
        min: '',
        max: '',
        autoFix: "value + '_' + row._rowNum",
        scopeFiles: 'ALL',
        startDate: '',
        endDate: '',
        active: true,
        desc: 'No duplicate IDs allowed across consolidated sheets'
      },
      {
        id: 'RULE-2',
        col: 'COMPANY',
        severity: 'BLOCKING',
        semaphore: 'RED',
        type: 'IN_LIST',
        param: '1000, 2000, 950',
        min: '',
        max: '',
        autoFix: "'1000'",
        scopeFiles: 'ALL',
        startDate: '',
        endDate: '',
        active: true,
        desc: 'Must match SAP Master Data Company Codes'
      },
      {
        id: 'RULE-3',
        col: 'COSTCENTER',
        severity: 'BLOCKING',
        semaphore: 'RED',
        type: 'IN_LIST',
        param: 'CC_IT_01, CC_FIN_02, 12234566, 10180000',
        min: '',
        max: '',
        autoFix: "'CC_IT_01'",
        scopeFiles: 'ALL',
        startDate: '',
        endDate: '',
        active: true,
        desc: 'Must match SAP Master Data Cost Centers'
      },
      {
        id: 'RULE-4',
        col: 'H1',
        severity: 'PREVENTIVE',
        semaphore: 'AMBER',
        type: 'NUMERIC_RANGE',
        param: '',
        min: '0',
        max: '100000',
        autoFix: 'num < 0 ? Math.abs(num) : 100000',
        scopeFiles: 'ALL',
        startDate: '',
        endDate: '',
        active: true,
        desc: 'Preventive budget threshold: H1 between 0 and 100,000'
      },
      {
        id: 'RULE-5',
        col: 'CAPITALIZED',
        severity: 'CHECK',
        semaphore: 'PURPLE',
        type: 'ASSISTED_CODE',
        param: "Number(row.H1 || 0) <= 50000 || String(value).toUpperCase() === 'Y' || String(value).toUpperCase() === 'X'",
        min: '',
        max: '',
        autoFix: "'Y'",
        scopeFiles: 'ALL',
        startDate: '',
        endDate: '',
        active: true,
        desc: 'Check: If H1 > 50,000, asset should be CAPITALIZED (Y/X)'
      }
    ];

    const savedRules = this._storageGet('sac_validation_rules_v3');
    this._validationRules = Array.isArray(savedRules) ? savedRules : defaultRules;

    this._newRuleForm = {
      col: this._requiredColumns[0] || 'ID',
      severity: 'BLOCKING',
      semaphore: 'RED',
      type: 'IN_LIST',
      param: '',
      min: '',
      max: '',
      autoFix: '',
      scopeFiles: 'ALL',
      startDate: '',
      endDate: '',
      desc: ''
    };

    const savedHistory = this._storageGet('sac_audit_history_v3');
    this._history = Array.isArray(savedHistory) ? savedHistory : [];

    this._currentSessionLogs = [];
    this._selectedAuditId = null;
    this._auditDetailSubTab = 'dataTab';
    this._auditSearchQuery = '';

    this._issueFilter = 'all';
    this._searchQuery = '';
    this._acknowledgeNonBlocking = false;
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
    this._storageSet('sac_audit_history_v3', this._history.slice(0, 30));
  }

  _saveRules() {
    this._storageSet('sac_validation_rules_v3', this._validationRules);
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

  _syncExportColumnsWithSchema(availableCols) {
    const validSelected = (this._auditExportConfig.selectedCols || []).filter(c => availableCols.includes(c));
    if (validSelected.length === 0) {
      this._auditExportConfig.selectedCols = [...availableCols];
    } else {
      this._auditExportConfig.selectedCols = validSelected;
    }
    this._auditExportConfig.numCols = this._auditExportConfig.selectedCols.length;
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

  _normalizeCellValue(val) {
    if (val === undefined || val === null) return '';
    let str = String(val).trim();
    if (this._parseConfig.numberFormat === 'EU' && str !== '') {
      if (/^-?\d{1,3}(\.\d{3})*(,\d+)$/.test(str) || /^-?\d+(,\d+)$/.test(str)) {
        str = str.replace(/\./g, '').replace(',', '.');
      }
    }
    return str;
  }

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

  _evalAssistedCode(expr, row, colName) {
    const value = String(row[colName] ?? '').trim();
    const num = Number(value);
    const fn = new Function('value', 'num', 'row', `"use strict"; return (${expr});`);
    return fn(value, num, row);
  }

  _applyCellAutoFix(origIdx, colName, ruleId) {
    const row = this._rows[origIdx];
    if (!row) return;
    const alerts = (row._alerts && row._alerts[colName]) || [];
    const targetAlert = ruleId ? alerts.find(a => a.ruleId === ruleId) : alerts.find(a => a.autoFix);
    if (!targetAlert || !targetAlert.autoFix) return;

    try {
      const oldVal = String(row[colName] ?? '').trim();
      const computed = this._evalAssistedCode(targetAlert.autoFix, row, colName);
      const newVal = String(computed ?? '').trim();
      if (oldVal !== newVal) {
        row[colName] = newVal;
        this._recordCellModification(row, colName, oldVal || '(Blank)', newVal);
        this._logAction(
          'AUTO_FIX_CODE',
          `Row #${row._rowNum} [${colName}]`,
          `Assisted auto-fix (${targetAlert.autoFix}) applied for rule ${targetAlert.ruleId}`,
          oldVal || '(Blank)',
          newVal
        );
        this._validateRows();
        this._notify(`Auto-fixed Row #${row._rowNum} [${colName}] → '${newVal}'`, 'success');
      }
    } catch (e) {
      this._notify(`Assisted auto-fix error: ${e.message}`, 'error');
    }
  }

  _applyAllAutoFixes() {
    let fixedCount = 0;
    this._rows.forEach(row => {
      this._requiredColumns.forEach(col => {
        const alerts = (row._alerts && row._alerts[col]) || [];
        const fixable = alerts.find(a => a.autoFix);
        if (fixable) {
          try {
            const oldVal = String(row[col] ?? '').trim();
            const computed = this._evalAssistedCode(fixable.autoFix, row, col);
            const newVal = String(computed ?? '').trim();
            if (oldVal !== newVal) {
              row[col] = newVal;
              this._recordCellModification(row, col, oldVal || '(Blank)', newVal);
              fixedCount++;
            }
          } catch (e) {}
        }
      });
    });

    if (fixedCount > 0) {
      this._logAction(
        'BULK_AUTO_FIX',
        `Grid Auto-Fix (${fixedCount} cells)`,
        'Executed Assisted Code Auto-Fix across all fixable alerts',
        `${fixedCount} alerts`,
        'FIXED'
      );
      this._validateRows();
      this._notify(`Assisted Code repaired ${fixedCount} cell(s) automatically!`, 'success');
    } else {
      this._notify('No cells with Assisted Auto-Fix rules pending.', 'info');
    }
  }

  _acknowledgeCheckCell(origIdx, colName, ruleId) {
    const row = this._rows[origIdx];
    if (!row) return;
    if (!row._checkedRules) row._checkedRules = {};
    const key = `${colName}__${ruleId}`;
    row._checkedRules[key] = true;
    this._logAction(
      'CHECK_VERIFIED',
      `Row #${row._rowNum} [${colName}]`,
      `User manually verified & approved check alert (${ruleId})`,
      'Pending Check',
      'Verified ✓'
    );
    this._validateRows();
    this._render();
  }

  _bindEvents() {
    const shadow = this._shadow;

    const toggleHelpBtn = shadow.getElementById('toggleHelpBtn');
    const closeHelpModalBtn = shadow.getElementById('closeHelpModalBtn');
    if (toggleHelpBtn) {
      toggleHelpBtn.addEventListener('click', () => {
        this._showHelpModal = !this._showHelpModal;
        this._render();
      });
    }
    if (closeHelpModalBtn) {
      closeHelpModalBtn.addEventListener('click', () => {
        this._showHelpModal = false;
        this._render();
      });
    }

    const loadStrategySel = shadow.getElementById('loadStrategySel');
    if (loadStrategySel) {
      loadStrategySel.addEventListener('change', (e) => {
        this._loadStrategy = e.target.value;
        this._logAction('STRATEGY_CHANGE', 'Load Strategy', `Switched planning mode to ${this._loadStrategy}`, '', this._loadStrategy);
      });
    }

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

    const cfgFormat = shadow.getElementById('cfgFormat');
    const cfgDelimiter = shadow.getElementById('cfgDelimiter');
    const cfgNumFormat = shadow.getElementById('cfgNumFormat');
    const cfgHeaderRow = shadow.getElementById('cfgHeaderRow');
    const reparseFilesBtn = shadow.getElementById('reparseFilesBtn');

    if (cfgFormat) cfgFormat.addEventListener('change', (e) => { this._parseConfig.format = e.target.value; });
    if (cfgDelimiter) cfgDelimiter.addEventListener('change', (e) => { this._parseConfig.delimiter = e.target.value; });
    if (cfgNumFormat) cfgNumFormat.addEventListener('change', (e) => { this._parseConfig.numberFormat = e.target.value; });
    if (cfgHeaderRow) {
      cfgHeaderRow.addEventListener('change', (e) => {
        this._parseConfig.headerRow = Math.max(1, parseInt(e.target.value, 10) || 1);
      });
    }
    if (reparseFilesBtn) {
      reparseFilesBtn.addEventListener('click', () => this._reparseAllFilesWithNewConfig());
    }

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

    shadow.querySelectorAll('.sheet-toggle-chk').forEach(chk => {
      chk.addEventListener('change', (e) => {
        const fIdx = parseInt(e.currentTarget.dataset.fidx, 10);
        const sIdx = parseInt(e.currentTarget.dataset.sidx, 10);
        const fileItem = this._filesQueue[fIdx];
        if (fileItem && fileItem.sheets && fileItem.sheets[sIdx]) {
          const sh = fileItem.sheets[sIdx];
          sh.selected = e.currentTarget.checked;
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
          this._syncRowsFromActiveSheets();
          this._render();
        }
      });
    });

    const clearAllFilesBtn = shadow.getElementById('clearAllFilesBtn');
    if (clearAllFilesBtn) {
      clearAllFilesBtn.addEventListener('click', () => {
        this._filesQueue = [];
        this._rows = this._rows.filter(r => r._fileOwner === '__MANUAL__');
        if (this._rows.length === 0) this._step = 1;
        this._validateRows();
        this._notify('All uploaded files removed from queue.', 'info');
      });
    }

    const downloadTemplateBtn = shadow.getElementById('downloadTemplateBtn');
    if (downloadTemplateBtn) downloadTemplateBtn.addEventListener('click', () => this._downloadTemplate());

    const exportCleanBtn = shadow.getElementById('exportCleanBtn');
    if (exportCleanBtn) exportCleanBtn.addEventListener('click', () => this._exportCleanExcel());

    const nextStep3Btn = shadow.getElementById('nextStep3Btn');
    const prevBtn = shadow.getElementById('prevStepBtn');
    const finishBtn = shadow.getElementById('finishBtn');
    const exitBtn = shadow.getElementById('exitBtn');
    const downloadLogBtn = shadow.getElementById('downloadLogBtn');
    const addRowBtn = shadow.getElementById('addRowBtn');
    const autoFixAllBtn = shadow.getElementById('autoFixAllBtn');

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

    if (downloadLogBtn) downloadLogBtn.addEventListener('click', () => this._exportErrorLog());
    if (autoFixAllBtn) autoFixAllBtn.addEventListener('click', () => this._applyAllAutoFixes());

    if (addRowBtn) {
      addRowBtn.addEventListener('click', () => {
        const newId = 'PLN_' + Math.floor(Math.random() * 9000 + 1000);
        const newRow = { _rowNum: 2, _sourceFile: 'Manual Entry', _fileOwner: '__MANUAL__', _isManual: true, _editedCols: {}, _checkedRules: {}, _errors: {}, _alerts: {} };
        this._requiredColumns.forEach(col => {
          newRow[col] = col === 'ID' ? newId : '';
        });
        this._rows.unshift(newRow);
        this._rows.forEach((r, idx) => { r._rowNum = idx + 2; });
        this._validateRows();
        this._notify('New planning row added.', 'success');
      });
    }

    shadow.querySelectorAll('.cell-autofix-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const origIdx = parseInt(e.currentTarget.dataset.origIdx, 10);
        const col = e.currentTarget.dataset.col;
        const ruleId = e.currentTarget.dataset.ruleId;
        this._applyCellAutoFix(origIdx, col, ruleId);
      });
    });

    shadow.querySelectorAll('.cell-check-ok-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const origIdx = parseInt(e.currentTarget.dataset.origIdx, 10);
        const col = e.currentTarget.dataset.col;
        const ruleId = e.currentTarget.dataset.ruleId;
        this._acknowledgeCheckCell(origIdx, col, ruleId);
      });
    });

    const ackNonBlockingChk = shadow.getElementById('ackNonBlockingChk');
    if (ackNonBlockingChk) {
      ackNonBlockingChk.addEventListener('change', (e) => {
        this._acknowledgeNonBlocking = e.currentTarget.checked;
        this._render();
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
        this._auditExportConfig.selectedCols.push(upperCol);
        this._auditExportConfig.numCols = this._auditExportConfig.selectedCols.length;
        this._rows.forEach(r => { if (r[upperCol] === undefined) r[upperCol] = ''; });
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
        this._syncExportColumnsWithSchema(this._requiredColumns);
        if (!this._requiredColumns.includes(this._bulkCol)) {
          this._bulkCol = this._requiredColumns[0] || 'ID';
        }
        this._validateRows();
        this._notify(`Column '${colToRemove}' removed.`, 'info');
      });
    });

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
    const ruleSeveritySel = shadow.getElementById('ruleSeveritySel');
    const ruleSemaphoreSel = shadow.getElementById('ruleSemaphoreSel');
    const ruleTypeSel = shadow.getElementById('ruleTypeSel');
    const ruleParamInp = shadow.getElementById('ruleParamInp');
    const ruleMinInp = shadow.getElementById('ruleMinInp');
    const ruleMaxInp = shadow.getElementById('ruleMaxInp');
    const ruleAutoFixInp = shadow.getElementById('ruleAutoFixInp');
    const ruleScopeFilesInp = shadow.getElementById('ruleScopeFilesInp');
    const ruleStartDateInp = shadow.getElementById('ruleStartDateInp');
    const ruleEndDateInp = shadow.getElementById('ruleEndDateInp');
    const ruleDescInp = shadow.getElementById('ruleDescInp');
    const saveNewRuleBtn = shadow.getElementById('saveNewRuleBtn');
    const resetDefaultRulesBtn = shadow.getElementById('resetDefaultRulesBtn');

    if (ruleColSel) ruleColSel.addEventListener('change', (e) => { this._newRuleForm.col = e.target.value; });
    if (ruleSeveritySel) ruleSeveritySel.addEventListener('change', (e) => { this._newRuleForm.severity = e.target.value; });
    if (ruleSemaphoreSel) ruleSemaphoreSel.addEventListener('change', (e) => { this._newRuleForm.semaphore = e.target.value; });
    if (ruleTypeSel) {
      ruleTypeSel.addEventListener('change', (e) => {
        this._newRuleForm.type = e.target.value;
        this._render();
      });
    }
    if (ruleParamInp) ruleParamInp.addEventListener('input', (e) => { this._newRuleForm.param = e.target.value; });
    if (ruleMinInp) ruleMinInp.addEventListener('input', (e) => { this._newRuleForm.min = e.target.value; });
    if (ruleMaxInp) ruleMaxInp.addEventListener('input', (e) => { this._newRuleForm.max = e.target.value; });
    if (ruleAutoFixInp) ruleAutoFixInp.addEventListener('input', (e) => { this._newRuleForm.autoFix = e.target.value; });
    if (ruleScopeFilesInp) ruleScopeFilesInp.addEventListener('input', (e) => { this._newRuleForm.scopeFiles = e.target.value; });
    if (ruleStartDateInp) ruleStartDateInp.addEventListener('change', (e) => { this._newRuleForm.startDate = e.target.value; });
    if (ruleEndDateInp) ruleEndDateInp.addEventListener('change', (e) => { this._newRuleForm.endDate = e.target.value; });
    if (ruleDescInp) ruleDescInp.addEventListener('input', (e) => { this._newRuleForm.desc = e.target.value; });

    shadow.querySelectorAll('.assisted-snippet-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const preset = e.currentTarget.dataset.preset;
        if (preset === 'capex_check') {
          this._newRuleForm.col = this._requiredColumns.includes('CAPITALIZED') ? 'CAPITALIZED' : this._requiredColumns[0];
          this._newRuleForm.severity = 'CHECK';
          this._newRuleForm.semaphore = 'PURPLE';
          this._newRuleForm.type = 'ASSISTED_CODE';
          this._newRuleForm.param = "Number(row.H1 || 0) <= 50000 || String(value).toUpperCase() === 'Y'";
          this._newRuleForm.autoFix = "'Y'";
          this._newRuleForm.desc = 'Chequeo CAPEX: Si H1 > 50.000, CAPITALIZED debe ser Y';
        } else if (preset === 'budget_prev') {
          this._newRuleForm.col = this._requiredColumns.includes('H1') ? 'H1' : this._requiredColumns[0];
          this._newRuleForm.severity = 'PREVENTIVE';
          this._newRuleForm.semaphore = 'AMBER';
          this._newRuleForm.type = 'ASSISTED_CODE';
          this._newRuleForm.param = 'num >= 0 && num <= 100000';
          this._newRuleForm.autoFix = 'num < 0 ? Math.abs(num) : 100000';
          this._newRuleForm.desc = 'Alerta Preventiva: Importe H1 entre 0 y 100.000';
        } else if (preset === 'cc_company_block') {
          this._newRuleForm.col = this._requiredColumns.includes('COSTCENTER') ? 'COSTCENTER' : this._requiredColumns[0];
          this._newRuleForm.severity = 'BLOCKING';
          this._newRuleForm.semaphore = 'RED';
          this._newRuleForm.type = 'ASSISTED_CODE';
          this._newRuleForm.param = "row.COMPANY !== '1000' || String(value).startsWith('CC_')";
          this._newRuleForm.autoFix = "'CC_' + String(value).trim()";
          this._newRuleForm.desc = 'Bloqueante: Para Sociedad 1000 el Centro de Coste empieza por CC_';
        }
        this._render();
      });
    });

    if (saveNewRuleBtn) saveNewRuleBtn.addEventListener('click', () => this._createNewValidationRule());

    if (resetDefaultRulesBtn) {
      resetDefaultRulesBtn.addEventListener('click', () => {
        this._storageSet('sac_validation_rules_v3', null);
        window.location.reload();
      });
    }

    shadow.querySelectorAll('.toggle-rule-chk').forEach(chk => {
      chk.addEventListener('change', (e) => {
        const rId = e.currentTarget.dataset.ruleId;
        const rule = this._validationRules.find(r => r.id === rId);
        if (rule) {
          rule.active = e.currentTarget.checked;
          this._saveRules();
          this._validateRows();
          this._render();
        }
      });
    });

    shadow.querySelectorAll('.delete-rule-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const rId = e.currentTarget.dataset.ruleId;
        this._validationRules = this._validationRules.filter(r => r.id !== rId);
        this._saveRules();
        this._validateRows();
        this._notify('Validation rule removed.', 'info');
      });
    });

    // Copiar regla a otro fichero o a todos los ficheros
    shadow.querySelectorAll('.copy-rule-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const rId = e.currentTarget.dataset.ruleId;
        const targetFile = prompt('Introduce el nombre exacto del fichero destino o escribe "ALL" para aplicarlo a todos los excels:', 'ALL');
        if (targetFile !== null) {
          const rule = this._validationRules.find(r => r.id === rId);
          if (rule) {
            rule.scopeFiles = targetFile.trim().toUpperCase() === 'ALL' ? 'ALL' : targetFile.trim();
            this._saveRules();
            this._validateRows();
            this._notify(`Rule scope updated to: ${rule.scopeFiles}`, 'success');
            this._render();
          }
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

    if (bulkColSelect) bulkColSelect.addEventListener('change', (e) => { this._bulkCol = e.target.value; });
    if (bulkValInput) bulkValInput.addEventListener('input', (e) => { this._bulkVal = e.target.value; });
    if (applyBulkEmptyBtn) applyBulkEmptyBtn.addEventListener('click', () => this._applyBulkFill(true));
    if (applyBulkAllBtn) applyBulkAllBtn.addEventListener('click', () => this._applyBulkFill(false));

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
        const idx = parseInt(e.currentTarget.dataset.index, 10);
        const removedFile = this._filesQueue[idx];
        if (removedFile) {
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
        const rec = this._history.find(h => h.id === this._selectedAuditId);
        const availCols = (rec && rec.columns) || this._requiredColumns;
        this._syncExportColumnsWithSchema(availCols);
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
        this._storageSet('sac_audit_history_v3', null);
        this._notify('Audit history cleared.', 'info');
      });
    }

    const toggleAuditExportCfgBtn = shadow.getElementById('toggleAuditExportCfgBtn');
    if (toggleAuditExportCfgBtn) {
      toggleAuditExportCfgBtn.addEventListener('click', () => {
        this._auditExportConfig.open = !this._auditExportConfig.open;
        const rec = this._selectedAuditId ? this._history.find(h => h.id === this._selectedAuditId) : null;
        const availCols = (rec && rec.columns) || this._requiredColumns;
        this._syncExportColumnsWithSchema(availCols);
        this._render();
      });
    }

    const auditNumColsInp = shadow.getElementById('auditNumColsInp');
    if (auditNumColsInp) {
      auditNumColsInp.addEventListener('change', (e) => {
        const rec = this._selectedAuditId ? this._history.find(h => h.id === this._selectedAuditId) : null;
        const availCols = (rec && rec.columns) || this._requiredColumns;
        const count = Math.max(1, Math.min(availCols.length, parseInt(e.target.value, 10) || availCols.length));
        this._auditExportConfig.numCols = count;
        this._auditExportConfig.selectedCols = availCols.slice(0, count);
        this._render();
      });
    }

    shadow.querySelectorAll('.audit-exp-col-chk').forEach(chk => {
      chk.addEventListener('change', (e) => {
        const col = e.currentTarget.dataset.col;
        const rec = this._selectedAuditId ? this._history.find(h => h.id === this._selectedAuditId) : null;
        const availCols = (rec && rec.columns) || this._requiredColumns;
        if (e.currentTarget.checked) {
          if (!this._auditExportConfig.selectedCols.includes(col)) {
            this._auditExportConfig.selectedCols.push(col);
          }
        } else {
          if (this._auditExportConfig.selectedCols.length <= 1) {
            this._notify('At least 1 column must be selected for download.', 'error');
            e.currentTarget.checked = true;
            return;
          }
          this._auditExportConfig.selectedCols = this._auditExportConfig.selectedCols.filter(c => c !== col);
        }
        this._auditExportConfig.selectedCols = availCols.filter(c => this._auditExportConfig.selectedCols.includes(c));
        this._auditExportConfig.numCols = this._auditExportConfig.selectedCols.length;
        this._render();
      });
    });

    const auditIncCommentsChk = shadow.getElementById('auditIncCommentsChk');
    if (auditIncCommentsChk) {
      auditIncCommentsChk.addEventListener('change', (e) => {
        this._auditExportConfig.includeComments = e.currentTarget.checked;
      });
    }

    const auditIncDetailChk = shadow.getElementById('auditIncDetailChk');
    if (auditIncDetailChk) {
      auditIncDetailChk.addEventListener('change', (e) => {
        this._auditExportConfig.includeAuditDetail = e.currentTarget.checked;
      });
    }

    const auditExpFormatSel = shadow.getElementById('auditExpFormatSel');
    if (auditExpFormatSel) {
      auditExpFormatSel.addEventListener('change', (e) => {
        this._auditExportConfig.format = e.target.value;
      });
    }

    const runConfiguredAuditDownloadBtn = shadow.getElementById('runConfiguredAuditDownloadBtn');
    if (runConfiguredAuditDownloadBtn) {
      runConfiguredAuditDownloadBtn.addEventListener('click', () => {
        if (this._selectedAuditId) {
          const rec = this._history.find(h => h.id === this._selectedAuditId);
          if (rec) this._downloadConfiguredAuditReport([rec], `${rec.id}_Custom_Audit_Report`);
        } else {
          this._downloadConfiguredAuditReport(this._history, 'Global_Audit_Monitor_Report');
        }
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

    const exportAuditJsonBtn = shadow.getElementById('exportAuditJsonBtn');
    if (exportAuditJsonBtn) {
      exportAuditJsonBtn.addEventListener('click', () => {
        const rec = this._history.find(h => h.id === this._selectedAuditId);
        if (rec) this._exportAuditRecordJson(rec);
      });
    }

    if (finishBtn) {
      finishBtn.addEventListener('click', () => {
        const stats = this._getAlertStats();
        if (stats.blockingRows > 0) {
          this._notify(`🛑 Cannot synchronize: ${stats.blockingRows} row(s) have BLOCKING alerts.`, 'error');
          return;
        }
        if ((stats.preventiveRows > 0 || stats.checkRows > 0) && !this._acknowledgeNonBlocking) {
          this._notify(`⚠️ Please confirm review of Preventive and Check alerts before synchronizing.`, 'error');
          return;
        }
        if (this._rows.length === 0) {
          this._notify('No active data rows to load.', 'error');
          return;
        }

        const loadId = 'LOAD-' + Math.floor(10000 + Math.random() * 90000);
        const fileSummary = this._filesQueue.length > 0
          ? this._filesQueue.map(f => `${f.file.name} [Strategy: ${this._loadStrategy}]`).join('; ')
          : `Manual Planning [Strategy: ${this._loadStrategy}]`;

        const filesMetadata = this._filesQueue.map(f => ({
          name: f.file.name,
          strategy: this._loadStrategy,
          sizeKB: Math.max(1, Math.round(f.file.size / 1024)),
          sheets: (f.sheets || []).filter(s => s.selected).length,
          rows: (f.sheets || []).filter(s => s.selected).reduce((acc, s) => acc + s.rows.length, 0),
          comment: f.comment || 'No comment'
        }));

        const snapshotRows = this._rows.map(r => {
          const copy = {
            _rowNum: r._rowNum,
            _sourceFile: r._sourceFile,
            _fileOwner: r._fileOwner || '',
            _loadStrategy: this._loadStrategy,
            _isManual: !!r._isManual,
            _editedCols: { ...(r._editedCols || {}) }
          };
          this._requiredColumns.forEach(c => { copy[c] = r[c] ?? ''; });
          return copy;
        });

        const cleanPayloadArray = this._rows.map(r => {
          const clean = { _loadStrategy: this._loadStrategy };
          this._requiredColumns.forEach(c => { clean[c] = r[c] ?? ''; });
          return clean;
        });

        const auditRecord = {
          id: loadId,
          user: this._currentUser,
          loadStrategy: this._loadStrategy,
          file: fileSummary,
          filesDetails: filesMetadata,
          columns: [...this._requiredColumns],
          rows: this._rows.length,
          timestamp: new Date().toLocaleString(),
          status: stats.preventiveRows > 0 || stats.checkRows > 0 ? `${this._loadStrategy} (Approved Alerts)` : `${this._loadStrategy} Success`,
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
        this._acknowledgeNonBlocking = false;
        this._activeTab = 'monitorTab';
        this._selectedAuditId = loadId;
        this._syncExportColumnsWithSchema(auditRecord.columns);
        this._auditDetailSubTab = 'dataTab';
        this._notify(`Load ${loadId} (${this._loadStrategy}) synchronized successfully!`, 'success');
      });
    }
  }

  _createNewValidationRule() {
    const col = this._newRuleForm.col || this._requiredColumns[0] || 'ID';
    const severity = this._newRuleForm.severity || 'BLOCKING';
    const semaphore = this._newRuleForm.semaphore || 'RED';
    const type = this._newRuleForm.type || 'IN_LIST';
    const param = (this._newRuleForm.param || '').trim();
    const min = (this._newRuleForm.min || '').trim();
    const max = (this._newRuleForm.max || '').trim();
    const autoFix = (this._newRuleForm.autoFix || '').trim();
    const scopeFiles = (this._newRuleForm.scopeFiles || 'ALL').trim();
    const startDate = (this._newRuleForm.startDate || '').trim();
    const endDate = (this._newRuleForm.endDate || '').trim();
    let desc = (this._newRuleForm.desc || '').trim();

    if ((type === 'IN_LIST' || type === 'STARTS_WITH' || type === 'REGEX' || type === 'ASSISTED_CODE') && !param) {
      this._notify('Please specify validation parameter or assisted code formula.', 'error');
      return;
    }

    if (!desc) {
      desc = `${col} rule [${type}]`;
    }

    const newRule = {
      id: 'RULE-' + Math.floor(1000 + Math.random() * 9000),
      col,
      severity,
      semaphore,
      type,
      param,
      min,
      max,
      autoFix,
      scopeFiles,
      startDate,
      endDate,
      active: true,
      desc
    };

    this._validationRules.push(newRule);
    this._saveRules();
    this._newRuleForm.param = '';
    this._newRuleForm.min = '';
    this._newRuleForm.max = '';
    this._newRuleForm.autoFix = '';
    this._newRuleForm.desc = '';
    this._showInlineRuleCreator = false;
    this._validateRows();
    this._notify(`Rule created for column '${col}' (Scope: ${scopeFiles}).`, 'success');
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
    this._validateRows();
    this._notify(`Bulk updated ${updated} row(s) in column '${col}'.`, 'success');
  }

  async _handleFilesSelection(fileList) {
    if (typeof XLSX === 'undefined') {
      this._notify('Excel engine is still loading...', 'error');
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
          sheets: [],
          sheetCount: 0,
          rowCount: 0,
          parsed: false
        };
        this._filesQueue.push(queueItem);
        newItemsToParse.push(queueItem);
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
    await this._parseFilesInQueue(this._filesQueue);
    this._notify('All files re-parsed with new settings.', 'success');
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
    this._loadingText = 'Parsing sheets & data...';
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
                _checkedRules: {},
                _errors: {},
                _alerts: {}
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
                  _checkedRules: {},
                  _errors: {},
                  _alerts: {}
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

  _pushAlert(row, col, alertObj) {
    if (!row._alerts[col]) row._alerts[col] = [];
    row._alerts[col].push(alertObj);
    row._errors[col] = `[${alertObj.severity}] ${alertObj.msg}`;
  }

  _validateRows() {
    const activeRules = this._validationRules.filter(r => r.active);
    const now = new Date();

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
      row._alerts = {};
      if (!row._checkedRules) row._checkedRules = {};

      this._requiredColumns.forEach(col => {
        if (row[col] === undefined || String(row[col]).trim() === '') {
          this._pushAlert(row, col, {
            ruleId: 'SYS-MANDATORY',
            severity: 'BLOCKING',
            semaphore: 'RED',
            msg: 'Mandatory field',
            autoFix: ''
          });
        }
      });

      activeRules.forEach(rule => {
        const col = rule.col;
        if (!this._requiredColumns.includes(col)) return;

        // Validar alcance por fichero (scopeFiles)
        const scope = (rule.scopeFiles || 'ALL').trim().toUpperCase();
        if (scope !== 'ALL') {
          const owner = String(row._fileOwner || '').trim().toLowerCase();
          const matchTarget = scope.split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
          if (!matchTarget.includes(owner)) {
            return; // La regla no aplica a este fichero
          }
        }

        // Validar carácter temporal (startDate / endDate)
        if (rule.startDate) {
          const start = new Date(rule.startDate);
          if (!isNaN(start.getTime()) && now < start) return;
        }
        if (rule.endDate) {
          const end = new Date(rule.endDate);
          if (!isNaN(end.getTime()) && now > end) return;
        }

        const rawVal = String(row[col] ?? '').trim();
        if (!rawVal && rule.type !== 'ASSISTED_CODE') return;

        const severity = rule.severity || 'BLOCKING';
        const semaphore = rule.semaphore || (severity === 'BLOCKING' ? 'RED' : (severity === 'PREVENTIVE' ? 'AMBER' : (severity === 'CHECK' ? 'PURPLE' : 'BLUE')));
        const checkKey = `${col}__${rule.id}`;
        if ((severity === 'CHECK' || severity === 'PREVENTIVE') && row._checkedRules[checkKey]) {
          return;
        }

        let failed = false;
        let failMsg = rule.desc || 'Validation failed';

        if (rule.type === 'UNIQUE') {
          if (rawVal && valueCountsPerCol[col] && valueCountsPerCol[col][rawVal] > 1) {
            failed = true;
            failMsg = 'Duplicate value not allowed';
          }
        } else if (rule.type === 'IN_LIST') {
          const allowed = rule.param.split(',').map(s => s.trim()).filter(Boolean);
          if (allowed.length > 0 && !allowed.includes(rawVal)) {
            failed = true;
            failMsg = `Allowed: ${allowed.join(', ')}`;
          }
        } else if (rule.type === 'NUMERIC_RANGE') {
          const num = Number(rawVal);
          if (isNaN(num)) {
            failed = true;
            failMsg = 'Must be numeric';
          } else {
            if (rule.min !== '' && !isNaN(Number(rule.min)) && num < Number(rule.min)) {
              failed = true;
              failMsg = `Min threshold: ${rule.min}`;
            }
            if (rule.max !== '' && !isNaN(Number(rule.max)) && num > Number(rule.max)) {
              failed = true;
              failMsg = `Max threshold: ${rule.max}`;
            }
          }
        } else if (rule.type === 'STARTS_WITH') {
          if (!rawVal.toUpperCase().startsWith(rule.param.toUpperCase())) {
            failed = true;
            failMsg = `Must start with '${rule.param}'`;
          }
        } else if (rule.type === 'REGEX') {
          try {
            const re = new RegExp(rule.param);
            if (!re.test(rawVal)) {
              failed = true;
              failMsg = `Invalid pattern (${rule.param})`;
            }
          } catch (e) {}
        } else if (rule.type === 'ASSISTED_CODE') {
          try {
            const isValid = Boolean(this._evalAssistedCode(rule.param, row, col));
            if (!isValid) {
              failed = true;
              failMsg = rule.desc || `Formula failed: ${rule.param}`;
            }
          } catch (e) {
            failed = true;
            failMsg = `Code syntax error: ${e.message}`;
          }
        }

        if (failed) {
          this._pushAlert(row, col, {
            ruleId: rule.id,
            severity,
            semaphore,
            msg: failMsg,
            autoFix: rule.autoFix || ''
          });
        }
      });
    });
  }

  _getRowSeverities(row) {
    const set = new Set();
    Object.values(row._alerts || {}).forEach(arr => {
      (arr || []).forEach(a => set.add(a.severity));
    });
    return set;
  }

  _getAlertStats() {
    let blockingRows = 0;
    let preventiveRows = 0;
    let checkRows = 0;
    let infoRows = 0;
    let fixableCells = 0;

    this._rows.forEach(r => {
      const sevs = this._getRowSeverities(r);
      if (sevs.has('BLOCKING')) blockingRows++;
      if (sevs.has('PREVENTIVE')) preventiveRows++;
      if (sevs.has('CHECK')) checkRows++;
      if (sevs.has('INFO')) infoRows++;

      Object.values(r._alerts || {}).forEach(arr => {
        if ((arr || []).some(a => a.autoFix)) fixableCells++;
      });
    });

    const cleanRows = this._rows.filter(r => this._getRowSeverities(r).size === 0).length;
    return { blockingRows, preventiveRows, checkRows, infoRows, cleanRows, fixableCells };
  }

  _downloadTemplate() {
    if (typeof XLSX === 'undefined') return;
    const sampleRow = this._requiredColumns.map(col => {
      if (col === 'ID') return 'PLN_1001';
      if (col === 'COMPANY') return '1000';
      if (col === 'COSTCENTER') return 'CC_IT_01';
      if (col === 'H1') return '25000';
      if (col === 'CAPITALIZED') return 'Y';
      return 'SAMPLE_DESC';
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

  _downloadConfiguredAuditReport(recordsList, filePrefix) {
    if (!recordsList || recordsList.length === 0) {
      this._notify('No audit records available to download.', 'error');
      return;
    }

    const cfg = this._auditExportConfig;
    const chosenCols = cfg.selectedCols && cfg.selectedCols.length > 0
      ? cfg.selectedCols
      : this._requiredColumns.slice(0, cfg.numCols || this._requiredColumns.length);

    const baseHeaders = ['LOAD_ID', 'LOAD_STRATEGY', 'TIMESTAMP', 'SAP_USER', 'ROW_NUM', 'SOURCE_WORKBOOK_SHEET'];
    if (cfg.includeComments) baseHeaders.push('BUSINESS_COMMENTS');
    if (cfg.includeAuditDetail) baseHeaders.push('AUDIT_TRACEABILITY_STATE', 'CELL_MODIFICATION_HISTORY');
    const finalHeaders = [...baseHeaders, ...chosenCols];

    const dataRows = [];
    const allChangeLogs = [];

    recordsList.forEach(rec => {
      const snap = rec.snapshotRows || [];
      snap.forEach(r => {
        const editedCols = r._editedCols || {};
        const editSummary = Object.keys(editedCols)
          .map(c => `${c}: '${editedCols[c].from}' -> '${editedCols[c].to}'`)
          .join(' | ');
        const stateLabel = r._isManual
          ? 'NEW / CLONED IN WIDGET'
          : (Object.keys(editedCols).length > 0 ? `EDITED (${Object.keys(editedCols).length} cols)` : 'ORIGINAL EXCEL');

        const rowArr = [
          rec.id,
          rec.loadStrategy || 'FULL',
          rec.timestamp,
          rec.user,
          r._rowNum,
          r._sourceFile
        ];

        if (cfg.includeComments) rowArr.push(r._fileComment || rec.file || 'No comment');
        if (cfg.includeAuditDetail) rowArr.push(stateLabel, editSummary || 'No changes');

        chosenCols.forEach(col => { rowArr.push(r[col] ?? ''); });
        dataRows.push(rowArr);
      });

      if (cfg.includeAuditDetail && rec.changeLog) {
        rec.changeLog.forEach(l => {
          allChangeLogs.push([rec.id, rec.loadStrategy || 'FULL', rec.timestamp, rec.user, l.time, l.type, l.target, l.detail, l.oldVal, l.newVal]);
        });
      }
    });

    if (cfg.format === 'CSV' || typeof XLSX === 'undefined') {
      const escapeCsv = (v) => `"${String(v ?? '').replace(/"/g, '""')}"`;
      let csvContent = finalHeaders.map(escapeCsv).join(',') + '\n';
      dataRows.forEach(r => { csvContent += r.map(escapeCsv).join(',') + '\n'; });
      const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${filePrefix}_${chosenCols.length}Cols.csv`;
      a.click();
      URL.revokeObjectURL(url);
      this._notify('Audit report downloaded as CSV.', 'success');
      return;
    }

    const wb = XLSX.utils.book_new();
    const wsData = XLSX.utils.aoa_to_sheet([finalHeaders, ...dataRows]);
    XLSX.utils.book_append_sheet(wb, wsData, 'Audit_Snapshot_Data');

    if (cfg.includeComments) {
      const filesHeader = ['LOAD_ID', 'LOAD_STRATEGY', 'TIMESTAMP', 'WORKBOOK_NAME', 'SIZE_KB', 'SELECTED_SHEETS', 'COMMITTED_ROWS', 'BUSINESS_COMMENT'];
      const filesRows = [];
      recordsList.forEach(rec => {
        (rec.filesDetails || []).forEach(f => {
          filesRows.push([
            rec.id,
            rec.strategy || rec.loadStrategy || 'FULL',
            rec.timestamp,
            f.name,
            f.sizeKB || 1,
            (f.selectedSheetNames || []).join(', '),
            f.rows,
            f.comment || 'No comment'
          ]);
        });
      });
      if (filesRows.length > 0) {
        const wsFiles = XLSX.utils.aoa_to_sheet([filesHeader, ...filesRows]);
        XLSX.utils.book_append_sheet(wb, wsFiles, 'Workbooks_&_Comments');
      }
    }

    if (cfg.includeAuditDetail && allChangeLogs.length > 0) {
      const logHeader = ['LOAD_ID', 'LOAD_STRATEGY', 'LOAD_DATE', 'SAP_USER', 'EVENT_TIME', 'ACTION_TYPE', 'TARGET', 'DESCRIPTION', 'PREVIOUS_VALUE', 'NEW_VALUE'];
      const wsLog = XLSX.utils.aoa_to_sheet([logHeader, ...allChangeLogs]);
      XLSX.utils.book_append_sheet(wb, wsLog, 'Audit_Activity_Log');
    }

    XLSX.writeFile(wb, `${filePrefix}_${chosenCols.length}Cols.xlsx`);
    this._notify('Configurable Audit Report exported successfully.', 'success');
  }

  _exportAuditRecordJson(rec) {
    if (!rec) return;
    const blob = new Blob([JSON.stringify(rec, null, 2)], { type: 'application/json;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${rec.id}_Audit_Certificate.json`;
    a.click();
    URL.revokeObjectURL(url);
    this._notify(`Audit certificate ${rec.id} downloaded.`, 'success');
  }

  _exportErrorLog() {
    let csv = 'SourceFile,Row,ID,Column,Alert Severity,Semaphore,Rule ID,Alert Description\n';
    let count = 0;
    this._rows.forEach(r => {
      Object.keys(r._alerts || {}).forEach(col => {
        (r._alerts[col] || []).forEach(a => {
          count++;
          csv += `"${this._esc(r._sourceFile)}",${r._rowNum},"${this._esc(r.ID || '')}","${col}","${a.severity}","${a.semaphore}","${a.ruleId}","${this._esc(a.msg)}"\n`;
        });
      });
    });
    if (count === 0) return;
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'SAC_Validation_Alerts_Log.csv';
    a.click();
    URL.revokeObjectURL(url);
  }

  _getFilteredRows() {
    const q = this._searchQuery.trim().toLowerCase();
    return this._rows
      .map((row, origIdx) => ({ row, origIdx }))
      .filter(({ row }) => {
        const sevs = this._getRowSeverities(row);
        if (this._issueFilter === 'correct' && sevs.size > 0) return false;
        if (['BLOCKING', 'PREVENTIVE', 'CHECK', 'INFO'].includes(this._issueFilter) && !sevs.has(this._issueFilter)) {
          return false;
        }
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

  _renderSemaphoreDot(sem) {
    let bg = '#107E3E'; // Green
    if (sem === 'RED') bg = '#BB0000';
    else if (sem === 'AMBER') bg = '#E9730C';
    else if (sem === 'PURPLE') bg = '#6D28D9';
    else if (sem === 'BLUE') bg = '#0070F2';
    return `<span style="display:inline-block; width:10px; height:10px; border-radius:50%; background:${bg}; margin-right:4px; box-shadow:0 0 4px ${bg};" title="Semaphore: ${sem}"></span>`;
  }

  _renderSeverityBadge(sev, sem) {
    const dot = this._renderSemaphoreDot(sem);
    if (sev === 'BLOCKING') return `<span class="sev-badge sev-blocking">${dot} Bloqueante</span>`;
    if (sev === 'PREVENTIVE') return `<span class="sev-badge sev-preventive">${dot} Preventiva</span>`;
    if (sev === 'CHECK') return `<span class="sev-badge sev-check">${dot} Chequeo</span>`;
    return `<span class="sev-badge sev-info">${dot} Informativa</span>`;
  }

  _renderHelpModal() {
    if (!this._showHelpModal) return '';
    return `
      <div class="help-modal-overlay">
        <div class="help-modal-card">
          <div style="display:flex; justify-content:space-between; align-items:center; border-bottom:1px solid #D9E0E6; padding-bottom:12px; margin-bottom:14px;">
            <div style="display:flex; align-items:center; gap:10px;">
              <span style="background:#0070F2; color:#FFF; width:28px; height:28px; border-radius:50%; display:inline-flex; align-items:center; justify-content:center; font-weight:800; font-size:15px;">?</span>
              <div>
                <h3 style="margin:0; color:#1D2D3E; font-size:16px;">Guía de Uso — Excel Planning Uploader Pro</h3>
                <span style="font-size:11px; color:#556B82;">Consolidación Full/Delta, Pestañas, Reglas con Alcance Temporal, Semáforos y Audit Monitor</span>
              </div>
            </div>
            <button class="btn" id="closeHelpModalBtn" type="button">✕ Cerrar Ayuda</button>
          </div>

          <div style="display:grid; grid-template-columns: 1fr 1fr; gap:12px; max-height:420px; overflow:auto; padding-right:4px;">
            <div class="help-step-box">
              <div class="help-step-title">1️⃣ Carga Full vs. Delta & Formatos</div>
              <p style="margin:4px 0 0 0; font-size:12px; color:#32363A; line-height:1.45;">
                Selecciona la estrategia de carga (<strong>FULL</strong> para reemplazo total o <strong>DELTA</strong> para ajustes incrementales). Arrastra múltiples Excel/CSV, configura tabuladores y decimales europeos.
              </p>
            </div>

            <div class="help-step-box">
              <div class="help-step-title">2️⃣ Selector de Pestañas y Comentarios</div>
              <p style="margin:4px 0 0 0; font-size:12px; color:#32363A; line-height:1.45;">
                En el Paso 2 marca qué pestañas (sheets) exactas de cada workbook consolidar, añade comentarios de negocio y gestiona el esquema de columnas.
              </p>
            </div>

            <div class="help-step-box">
              <div class="help-step-title">3️⃣ Reglas con Semáforos, Alcance Temporal y Auto-Fix</div>
              <p style="margin:4px 0 0 0; font-size:12px; color:#32363A; line-height:1.45;">
                Crea reglas configurando un color semafórico (🟢🟡🔴🟣), aplícalas a ficheros específicos (<code>ALL</code> o nombre de fichero) o dales un carácter temporal (fechas inicio/fin). Utiliza <strong>Código Asistido</strong> y <strong>✨ Auto-Fix</strong>.
              </p>
            </div>

            <div class="help-step-box">
              <div class="help-step-title">4️⃣ Audit Monitor y Descarga a Medida</div>
              <p style="margin:4px 0 0 0; font-size:12px; color:#32363A; line-height:1.45;">
                Haz clic en cualquier carga del <strong>Audit Monitor</strong> para ver instantáneas y cambios. Usa el botón <strong>⚙️ Configure & Download Audit</strong> para seleccionar el número exacto de columnas, comentarios y detalle de auditoría a exportar.
              </p>
            </div>
          </div>
        </div>
      </div>
    `;
  }

  _renderAuditExportConfigurator(availableCols, isSingleRecord = false) {
    if (!this._auditExportConfig.open) return '';
    const cfg = this._auditExportConfig;
    return `
      <div style="background:#F0F6FF; border:1px solid #0070F2; border-radius:8px; padding:14px 16px; margin-bottom:14px;">
        <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px; margin-bottom:10px;">
          <div style="font-weight:800; color:#005CB9; font-size:13px;">
            ⚙️ Configurar Descarga de Auditoría ${isSingleRecord ? `(Carga ${this._esc(this._selectedAuditId)})` : '(Historial Global)'}
          </div>
          <div style="display:flex; align-items:center; gap:10px; flex-wrap:wrap;">
            <div>
              <label style="font-size:10px; font-weight:700; color:#556B82; margin-right:4px;">NÚMERO DE COLUMNAS (1 - ${availableCols.length}):</label>
              <input type="number" id="auditNumColsInp" class="input-std" style="width:58px; font-weight:700;" min="1" max="${availableCols.length}" value="${cfg.numCols}" />
            </div>
            <div>
              <label style="font-size:10px; font-weight:700; color:#556B82; margin-right:4px;">FORMATO:</label>
              <select id="auditExpFormatSel" class="input-std">
                <option value="XLSX" ${cfg.format === 'XLSX' ? 'selected' : ''}>Excel Workbook (.xlsx)</option>
                <option value="CSV" ${cfg.format === 'CSV' ? 'selected' : ''}>CSV Flat File (.csv)</option>
              </select>
            </div>
            <button class="btn btn-success" id="runConfiguredAuditDownloadBtn" type="button">
              📥 Descargar Informe Configurado (${cfg.selectedCols.length} cols)
            </button>
          </div>
        </div>

        <div style="display:flex; flex-wrap:wrap; gap:16px; align-items:center; padding-top:8px; border-top:1px dashed #B3D4FF; margin-bottom:10px;">
          <label style="display:inline-flex; align-items:center; gap:6px; font-size:12px; font-weight:700; color:#1D2D3E; cursor:pointer;">
            <input type="checkbox" id="auditIncCommentsChk" ${cfg.includeComments ? 'checked' : ''} style="cursor:pointer; width:15px; height:15px;" />
            <span>💬 Incorporar Comentarios de Negocio y Metadatos de Ficheros</span>
          </label>
          <label style="display:inline-flex; align-items:center; gap:6px; font-size:12px; font-weight:700; color:#1D2D3E; cursor:pointer;">
            <input type="checkbox" id="auditIncDetailChk" ${cfg.includeAuditDetail ? 'checked' : ''} style="cursor:pointer; width:15px; height:15px;" />
            <span>🛠️ Incorporar Detalle de Auditoría (Estado de Trazabilidad, Valores Previos y Log de Cambios)</span>
          </label>
        </div>

        <div style="display:flex; align-items:center; flex-wrap:wrap; gap:6px;">
          <span style="font-size:11px; font-weight:700; color:#556B82; margin-right:4px;">Columnas seleccionadas (${cfg.selectedCols.length}/${availableCols.length}):</span>
          ${availableCols.map(col => {
            const isChecked = cfg.selectedCols.includes(col);
            return `
              <label class="sheet-pill-chk ${isChecked ? 'sheet-active' : ''}">
                <input type="checkbox" class="audit-exp-col-chk" data-col="${this._esc(col)}" ${isChecked ? 'checked' : ''} style="margin:0; cursor:pointer;" />
                <span>${this._esc(col)}</span>
              </label>
            `;
          }).join('')}
        </div>
      </div>
    `;
  }

  _renderRuleCreatorForm() {
    const t = this._newRuleForm.type;
    const sev = this._newRuleForm.severity;
    const sem = this._newRuleForm.semaphore;
    return `
      <div style="background:#F8FAFC; border:1px solid #B3D4FF; border-radius:8px; padding:16px; margin-top:10px;">
        <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:8px; margin-bottom:12px;">
          <div style="font-weight:800; color:#005CB9; font-size:13px;">➕ Create Validation Rule, Semaphore, Scope & Assisted Code</div>
          <div style="display:flex; gap:6px; flex-wrap:wrap; align-items:center;">
            <span style="font-size:11px; font-weight:700; color:#556B82;">💻 Templates:</span>
            <button class="icon-btn assisted-snippet-btn" data-preset="cc_company_block" type="button">🛑 Sociedad/CeCo</button>
            <button class="icon-btn assisted-snippet-btn" data-preset="budget_prev" type="button">⚠️ Presupuesto</button>
            <button class="icon-btn assisted-snippet-btn" data-preset="capex_check" type="button">🔎 CAPEX</button>
          </div>
        </div>

        <div style="display:grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap:10px; margin-bottom:10px;">
          <div>
            <label style="display:block; font-size:10px; font-weight:700; color:#556B82; margin-bottom:3px;">COLUMN</label>
            <select id="ruleColSel" class="input-std" style="width:100%;">
              ${this._requiredColumns.map(c => `<option value="${this._esc(c)}" ${this._newRuleForm.col === c ? 'selected' : ''}>${this._esc(c)}</option>`).join('')}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:10px; font-weight:700; color:#556B82; margin-bottom:3px;">SEVERITY</label>
            <select id="ruleSeveritySel" class="input-std" style="width:100%; font-weight:700;">
              <option value="BLOCKING" ${sev === 'BLOCKING' ? 'selected' : ''}>🛑 Bloqueante</option>
              <option value="PREVENTIVE" ${sev === 'PREVENTIVE' ? 'selected' : ''}>⚠️ Preventiva</option>
              <option value="CHECK" ${sev === 'CHECK' ? 'selected' : ''}>🔎 Chequeo</option>
              <option value="INFO" ${sev === 'INFO' ? 'selected' : ''}>ℹ️ Informativa</option>
            </select>
          </div>

          <div>
            <label style="display:block; font-size:10px; font-weight:700; color:#556B82; margin-bottom:3px;">🚦 SEMAPHORE</label>
            <select id="ruleSemaphoreSel" class="input-std" style="width:100%; font-weight:700;">
              <option value="RED" ${sem === 'RED' ? 'selected' : ''}>🔴 Rojo (Red)</option>
              <option value="AMBER" ${sem === 'AMBER' ? 'selected' : ''}>🟡 Ámbar (Amber)</option>
              <option value="PURPLE" ${sem === 'PURPLE' ? 'selected' : ''}>🟣 Púrpura (Purple)</option>
              <option value="BLUE" ${sem === 'BLUE' ? 'selected' : ''}>🔵 Azul (Blue)</option>
            </select>
          </div>

          <div>
            <label style="display:block; font-size:10px; font-weight:700; color:#556B82; margin-bottom:3px;">LOGIC TYPE</label>
            <select id="ruleTypeSel" class="input-std" style="width:100%;">
              <option value="IN_LIST" ${t === 'IN_LIST' ? 'selected' : ''}>Master Data List</option>
              <option value="NUMERIC_RANGE" ${t === 'NUMERIC_RANGE' ? 'selected' : ''}>Numeric Range</option>
              <option value="STARTS_WITH" ${t === 'STARTS_WITH' ? 'selected' : ''}>Starts With Prefix</option>
              <option value="REGEX" ${t === 'REGEX' ? 'selected' : ''}>Regex Pattern</option>
              <option value="UNIQUE" ${t === 'UNIQUE' ? 'selected' : ''}>Unique (No Dupes)</option>
              <option value="ASSISTED_CODE" ${t === 'ASSISTED_CODE' ? 'selected' : ''}>💻 Assisted Code</option>
            </select>
          </div>
        </div>

        <div style="display:flex; flex-wrap:wrap; gap:10px; align-items:flex-end; margin-bottom:10px;">
          ${t === 'IN_LIST' || t === 'STARTS_WITH' || t === 'REGEX' || t === 'ASSISTED_CODE' ? `
            <div style="flex:2; min-width:220px;">
              <label style="display:block; font-size:10px; font-weight:700; color:#005CB9; margin-bottom:3px;">PARAMETER / FORMULA</label>
              <input type="text" id="ruleParamInp" class="input-std" style="width:100%; box-sizing:border-box;" placeholder="e.g. 1000, 2000 or JS expression" value="${this._esc(this._newRuleForm.param)}" />
            </div>
          ` : ''}

          ${t === 'NUMERIC_RANGE' ? `
            <div>
              <label style="display:block; font-size:10px; font-weight:700; color:#556B82; margin-bottom:3px;">MIN</label>
              <input type="number" id="ruleMinInp" class="input-std" style="width:90px;" placeholder="0" value="${this._esc(this._newRuleForm.min)}" />
            </div>
            <div>
              <label style="display:block; font-size:10px; font-weight:700; color:#556B82; margin-bottom:3px;">MAX</label>
              <input type="number" id="ruleMaxInp" class="input-std" style="width:90px;" placeholder="100000" value="${this._esc(this._newRuleForm.max)}" />
            </div>
          ` : ''}

          <div style="flex:1; min-width:170px;">
            <label style="display:block; font-size:10px; font-weight:700; color:#107E3E; margin-bottom:3px;">✨ AUTO-FIX CODE</label>
            <input type="text" id="ruleAutoFixInp" class="input-std" style="width:100%; box-sizing:border-box; font-family:monospace;" placeholder="e.g. '1000'" value="${this._esc(this._newRuleForm.autoFix)}" />
          </div>
        </div>

        <!-- ALCANCE Y CARÁCTER TEMPORAL -->
        <div style="display:flex; flex-wrap:wrap; gap:10px; align-items:flex-end; padding-top:8px; border-top:1px dashed #B3D4FF;">
          <div style="flex:1; min-width:180px;">
            <label style="display:block; font-size:10px; font-weight:700; color:#556B82; margin-bottom:3px;">SCOPE (FILE NAME OR 'ALL')</label>
            <input type="text" id="ruleScopeFilesInp" class="input-std" style="width:100%; box-sizing:border-box;" placeholder="ALL or file1.xlsx, file2.xlsx" value="${this._esc(this._newRuleForm.scopeFiles || 'ALL')}" />
          </div>
          <div>
            <label style="display:block; font-size:10px; font-weight:700; color:#556B82; margin-bottom:3px;">TEMPORAL START DATE</label>
            <input type="date" id="ruleStartDateInp" class="input-std" value="${this._esc(this._newRuleForm.startDate)}" />
          </div>
          <div>
            <label style="display:block; font-size:10px; font-weight:700; color:#556B82; margin-bottom:3px;">TEMPORAL END DATE</label>
            <input type="date" id="ruleEndDateInp" class="input-std" value="${this._esc(this._newRuleForm.endDate)}" />
          </div>
          <div style="flex:1; min-width:180px;">
            <label style="display:block; font-size:10px; font-weight:700; color:#556B82; margin-bottom:3px;">DESCRIPTION</label>
            <input type="text" id="ruleDescInp" class="input-std" style="width:100%; box-sizing:border-box;" placeholder="Rule description..." value="${this._esc(this._newRuleForm.desc)}" />
          </div>
          <div>
            <button class="btn" id="saveNewRuleBtn" type="button">+ Save Rule</button>
          </div>
        </div>
      </div>
    `;
  }

  _renderValidationRulesTab() {
    const activeRulesCount = this._validationRules.filter(r => r.active).length;
    return `
      <div>
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:14px; flex-wrap:wrap; gap:10px;">
          <div>
            <h3 style="margin:0; color:#1D2D3E;">⚙️ Validation Rules, Semaphores, Scope & Temporal Validity</h3>
            <p style="margin:4px 0 0 0; color:#556B82; font-size:12px;">Manage rule application scope (ALL or specific workbooks), date validity windows, semaphore indicators and auto-fixes.</p>
          </div>
          <button class="btn btn-ghost" id="resetDefaultRulesBtn" type="button">↺ Restore Default Rules</button>
        </div>

        ${this._renderRuleCreatorForm()}

        <div style="margin-top:18px; border:1px solid #D9E0E6; border-radius:8px; overflow:hidden; background:#FFFFFF;">
          <table>
            <thead>
              <tr>
                <th style="width:40px; text-align:center;">On</th>
                <th style="width:100px;">Column</th>
                <th style="width:120px;">Semaphore</th>
                <th style="width:120px;">Severity</th>
                <th style="width:120px;">Scope</th>
                <th style="width:140px;">Temporal Validity</th>
                <th>Logic & Assisted Code</th>
                <th style="width:60px; text-align:center;">Copy</th>
                <th style="width:60px; text-align:center;">Del</th>
              </tr>
            </thead>
            <tbody>
              ${this._validationRules.length > 0 ? this._validationRules.map(r => `
                <tr style="${!r.active ? 'opacity:0.55; background:#F8FAFC;' : ''}">
                  <td style="text-align:center;">
                    <input type="checkbox" class="toggle-rule-chk" data-rule-id="${this._esc(r.id)}" ${r.active ? 'checked' : ''} style="cursor:pointer; width:15px; height:15px;" />
                  </td>
                  <td><span class="col-pill">${this._esc(r.col)}</span></td>
                  <td>${this._renderSemaphoreDot(r.semaphore || 'RED')} <strong style="font-size:11px;">${this._esc(r.semaphore || 'RED')}</strong></td>
                  <td>${this._renderSeverityBadge(r.severity || 'BLOCKING', r.semaphore || 'RED')}</td>
                  <td><code style="background:#F0F6FF; color:#005CB9; padding:2px 6px; border-radius:4px; font-size:11px;">${this._esc(r.scopeFiles || 'ALL')}</code></td>
                  <td style="font-size:10px; color:#556B82;">
                    ${r.startDate || r.endDate ? `${r.startDate || 'Any'} → ${r.endDate || 'Any'}` : 'Permanent'}
                  </td>
                  <td style="font-family:monospace; font-size:11px; color:#1D2D3E;">
                    <strong>${this._esc(r.type)}</strong>: ${this._esc(r.param || r.desc)}
                  </td>
                  <td style="text-align:center;">
                    <button class="icon-btn copy-rule-btn" data-rule-id="${this._esc(r.id)}" type="button" title="Copiar / Asignar alcance">📋</button>
                  </td>
                  <td style="text-align:center;">
                    <button class="icon-btn icon-btn-del delete-rule-btn" data-rule-id="${this._esc(r.id)}" type="button" title="Delete rule">✕</button>
                  </td>
                </tr>
              `).join('') : `
                <tr><td colspan="9" style="text-align:center; padding:28px; color:#556B82;">No validation rules configured.</td></tr>
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
                <button class="btn btn-ghost" id="backToAuditListBtn" type="button">← Back to Audit Monitor</button>
                <div>
                  <div style="font-size:11px; color:#556B82; font-weight:600;">Audit Monitor  /  Execution Drill-Down</div>
                  <div style="font-size:16px; font-weight:800; color:#1D2D3E;">Load ID: ${this._esc(rec.id)} <span style="background:#E6F9EC; color:#107E3E; padding:3px 10px; border-radius:12px; font-size:11px; margin-left:6px;">✓ ${this._esc(rec.status)}</span></div>
                </div>
              </div>
              <div style="display:flex; gap:8px;">
                <button class="btn" id="toggleAuditExportCfgBtn" type="button">
                  ⚙️ ${this._auditExportConfig.open ? 'Close Export Config' : 'Configure & Download Audit'}
                </button>
                <button class="btn btn-ghost" id="exportAuditJsonBtn" type="button">📜 Audit Certificate (.json)</button>
              </div>
            </div>

            ${this._renderAuditExportConfigurator(cols, true)}

            <div class="metrics-bar" style="grid-template-columns: repeat(5, 1fr);">
              <div class="metric-box"><div class="metric-title">Strategy</div><div class="metric-val" style="font-size:13px; color:#0064D9;">${this._esc(rec.loadStrategy || 'FULL')}</div></div>
              <div class="metric-box"><div class="metric-title">Timestamp</div><div class="metric-val" style="font-size:13px;">${this._esc(rec.timestamp)}</div></div>
              <div class="metric-box kpi-ok"><div class="metric-title">Committed Rows</div><div class="metric-val" style="color:#107E3E;">${rec.rows}</div></div>
              <div class="metric-box"><div class="metric-title">Rows Modified</div><div class="metric-val" style="color:#0064D9;">${modifiedRowsCount}</div></div>
              <div class="metric-box"><div class="metric-title">Tracked Actions</div><div class="metric-val">${changeLog.length} events</div></div>
            </div>

            <div class="subtabs-bar">
              <button class="audit-subtab ${this._auditDetailSubTab === 'dataTab' ? 'active' : ''}" data-subtab="dataTab" type="button">
                📊 Consolidated Data Snapshot (${snapshotRows.length})
              </button>
              <button class="audit-subtab ${this._auditDetailSubTab === 'changesTab' ? 'active' : ''}" data-subtab="changesTab" type="button">
                🛠️ User Edits, Auto-Fixes & Checks (${changeLog.length})
              </button>
              <button class="audit-subtab ${this._auditDetailSubTab === 'filesTab' ? 'active' : ''}" data-subtab="filesTab" type="button">
                📁 Workbooks & SAC Payload (${filesDetails.length})
              </button>
            </div>

            ${this._auditDetailSubTab === 'dataTab' ? `
              <div class="toolbar-box" style="margin-bottom:10px;">
                <div style="display:flex; align-items:center; gap:8px;">
                  <span style="font-weight:700; font-size:12px;">🔍 Inspect Committed Rows:</span>
                  <input type="text" id="auditSearchInput" class="input-std" style="width:240px;" placeholder="Search ID, cost center, sheet..." value="${this._esc(this._auditSearchQuery)}" />
                </div>
              </div>

              <div class="tbl-wrap">
                <table>
                  <thead>
                    <tr>
                      <th style="width:32px;">#</th>
                      <th style="width:120px;">Traceability State</th>
                      <th>Source Workbook & Sheet</th>
                      <th>Business Comment</th>
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
                          <td style="font-size:11px; color:#556B82;">${this._esc(r._fileComment || 'No comment')}</td>
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
                      <tr><td colspan="${cols.length + 4}" style="text-align:center; padding:24px; color:#556B82;">No rows match the search filter.</td></tr>
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
                      if (['CELL_EDIT', 'BULK_FILL', 'AUTO_FIX_CODE', 'BULK_AUTO_FIX'].includes(log.type)) badgeStyle = 'background:#FFF8E6; color:#B45309;';
                      if (['ROW_DELETED', 'FILE_REMOVED', 'SCHEMA_DEL_COL', 'SHEET_DISABLED'].includes(log.type)) badgeStyle = 'background:#FEEAEA; color:#BB0000;';
                      if (['ROW_ADDED', 'ROW_CLONED', 'SHEET_ENABLED', 'CHECK_VERIFIED'].includes(log.type)) badgeStyle = 'background:#E6F9EC; color:#107E3E;';
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
                          Strategy: <strong>${this._esc(f.strategy || rec.loadStrategy || 'FULL')}</strong> • ${f.sheets} of ${f.totalSheets || f.sheets} sheet(s) selected • ${f.rows} rows
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
      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:14px; flex-wrap:wrap; gap:10px;">
        <div>
          <h3 style="margin:0; color:#1D2D3E;">Audit Monitor</h3>
          <p style="margin:4px 0 0 0; color:#556B82; font-size:12px;">Click on any execution row below to inspect committed rows, load strategy (Full/Delta), edits, and source workbooks, or configure custom exports.</p>
        </div>
        <div style="display:flex; gap:8px;">
          ${this._history.length > 0 ? `
            <button class="btn" id="toggleAuditExportCfgBtn" type="button" style="font-size:11px; padding:6px 12px;">
              ⚙️ ${this._auditExportConfig.open ? 'Close Export Config' : 'Configure & Download Audit'}
            </button>
            <button class="btn btn-danger" id="clearHistoryBtn" type="button" style="font-size:11px; padding:6px 12px;">🗑️ Clear History</button>
          ` : ''}
        </div>
      </div>

      ${this._renderAuditExportConfigurator(this._requiredColumns, false)}

      <div style="border: 1px solid #D9E0E6; border-radius: 8px; overflow:hidden; background:#FFFFFF; box-shadow: 0 2px 6px rgba(0,0,0,0.02);">
        <table>
          <thead>
            <tr>
              <th>Load ID</th>
              <th>Strategy</th>
              <th>SAP User</th>
              <th>Workbooks & Comments</th>
              <th>Rows</th>
              <th>Edits</th>
              <th>Timestamp</th>
              <th>Status</th>
              <th style="text-align:right;">Inspection</th>
            </tr>
          </thead>
          <tbody>
            ${this._history.length > 0 ? this._history.map(h => `
              <tr class="audit-row-link" data-load-id="${this._esc(h.id)}" style="cursor:pointer;">
                <td><strong style="color:#0064D9; text-decoration:underline;">${this._esc(h.id)}</strong></td>
                <td><span style="background:#EBF5FF; color:#005CB9; padding:2px 8px; border-radius:10px; font-weight:700; font-size:11px;">${this._esc(h.loadStrategy || 'FULL')}</span></td>
                <td>${this._esc(h.user)}</td>
                <td>${this._esc(h.file)}</td>
                <td><strong>${h.rows}</strong> rows</td>
                <td>
                  ${(h.editsCount || 0) > 0
                    ? `<span style="background:#FFF8E6; color:#B45309; border:1px solid #FCD34D; padding:2px 8px; border-radius:10px; font-weight:700; font-size:11px;">✎ ${h.editsCount} action(s)</span>`
                    : `<span style="color:#556B82; font-size:11px;">Direct (0 edits)</span>`}
                </td>
                <td>${this._esc(h.timestamp)}</td>
                <td><span style="background:#E6F9EC; color:#107E3E; padding:4px 10px; border-radius:12px; font-weight:700; font-size:11px;">✓ ${this._esc(h.status)}</span></td>
                <td style="text-align:right;">
                  <button class="btn btn-ghost" type="button" style="padding:4px 10px; font-size:11px;">🔍 View Details →</button>
                </td>
              </tr>
            `).join('') : `
              <tr><td colspan="9" style="text-align:center; padding:36px; color:#556B82;">No planning loads registered yet. Complete an upload in the Wizard to inspect its full audit trail here.</td></tr>
            `}
          </tbody>
        </table>
      </div>
    `;
  }

  _render() {
    const totalRows = this._rows.length;
    const stats = this._getAlertStats();
    const validPercent = totalRows > 0 ? Math.round(((totalRows - stats.blockingRows) / totalRows) * 100) : 100;
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

        .help-modal-overlay {
          position: fixed;
          inset: 0;
          background: rgba(19, 35, 55, 0.62);
          backdrop-filter: blur(2px);
          display: flex;
          align-items: center;
          justify-content: center;
          z-index: 250;
          padding: 20px;
        }
        .help-modal-card {
          background: #FFFFFF;
          border: 1px solid #D9E0E6;
          border-top: 4px solid #0070F2;
          border-radius: 12px;
          max-width: 820px;
          width: 100%;
          padding: 20px 24px;
          box-shadow: 0 16px 40px rgba(0,0,0,0.22);
        }
        .help-step-box {
          background: #F8FAFC;
          border: 1px solid #D9E0E6;
          border-left: 4px solid #0070F2;
          border-radius: 8px;
          padding: 12px 14px;
        }
        .help-step-title {
          font-weight: 800;
          color: #005CB9;
          font-size: 13px;
        }

        .sap-header {
          display: flex;
          align-items: center;
          justify-content: space-between;
          background: linear-gradient(90deg, #132337 0%, #1F3652 100%);
          color: #FFFFFF;
          padding: 12px 24px;
          border-bottom: 3px solid #0070F2;
        }
        .sap-brand { display: flex; align-items: center; gap: 14px; }
        .sap-logo-img { height: 28px; width: auto; display: block; filter: drop-shadow(0 2px 4px rgba(0, 0, 0, 0.25)); }
        .sap-title-wrap { display: flex; flex-direction: column; }
        .sap-title { font-size: 15px; font-weight: 700; letter-spacing: -0.2px; color: #FFFFFF; }
        .sap-subtitle { font-size: 11px; color: #A9B8C8; font-weight: 500; }
        .sap-badge-env {
          background: rgba(0, 184, 241, 0.16);
          border: 1px solid rgba(0, 184, 241, 0.45);
          color: #74D8FF;
          padding: 4px 10px;
          border-radius: 20px;
          font-size: 11px;
          font-weight: 600;
        }
        .help-icon-btn {
          background: rgba(255, 255, 255, 0.14);
          border: 1px solid rgba(255, 255, 255, 0.35);
          color: #FFFFFF;
          padding: 5px 12px;
          border-radius: 20px;
          font-size: 12px;
          font-weight: 700;
          cursor: pointer;
          display: inline-flex;
          align-items: center;
          gap: 6px;
          transition: all 0.15s;
        }
        .help-icon-btn:hover { background: #0070F2; border-color: #74D8FF; }

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

        .tab-content { padding: 24px; background: #F5F6F7; }

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
        }
        .audit-subtab.active { background: #0070F2; color: #FFFFFF; border-color: #0070F2; }

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
        .step-item:hover { border-color: #0070F2; transform: translateY(-1px); }
        .step-item.active { border-color: #0070F2; background: #F0F6FF; }
        .step-item.completed { border-color: #107E3E; }
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
        }
        .step-item.active .step-num { background: #0070F2; color: #FFFFFF; }
        .step-item.completed .step-num { background: #107E3E; color: #FFFFFF; }
        .step-label-title { font-weight: 700; font-size: 12px; color: #1D2D3E; }
        .step-label-sub { font-size: 11px; color: #556B82; }

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
        .btn:hover { background: #005CB9; }
        .btn-ghost { background: #FFFFFF; color: #0064D9; border-color: #BCC8D4; }
        .btn-ghost:hover { background: #F0F6FF; border-color: #0070F2; }
        .btn-danger { background: #FFFFFF; color: #BB0000; border-color: #F5C6C6; }
        .btn-danger:hover { background: #FEEAEA; border-color: #BB0000; }
        .btn-success { background: #107E3E; border-color: #107E3E; color: #FFFFFF; }
        .btn-success:hover { background: #0C6230; }

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
        }
        .drop-zone.drag-active { background: #F0F6FF; border-color: #005CB9; }
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
        }
        .file-label-btn {
          padding: 9px 20px;
          background: #0070F2;
          color: #FFFFFF;
          font-weight: 700;
          font-size: 13px;
          border-radius: 6px;
          cursor: pointer;
          display: inline-block;
        }

        .metrics-bar {
          display: grid;
          grid-template-columns: repeat(6, 1fr);
          gap: 10px;
          margin-bottom: 14px;
        }
        .metric-box {
          background: #FFFFFF;
          border: 1px solid #D9E0E6;
          border-left: 4px solid #0070F2;
          padding: 10px 12px;
          border-radius: 8px;
        }
        .metric-box.kpi-ok { border-left-color: #107E3E; }
        .metric-box.kpi-err { border-left-color: #BB0000; }
        .metric-box.kpi-prev { border-left-color: #E9730C; }
        .metric-box.kpi-chk { border-left-color: #6D28D9; }
        .metric-title { font-size: 10px; color: #556B82; margin-bottom: 4px; font-weight: 700; text-transform: uppercase; }
        .metric-val { font-size: 17px; font-weight: 800; color: #1D2D3E; }

        .dist-bar {
          display: flex;
          height: 8px;
          border-radius: 4px;
          overflow: hidden;
          background: #FEEAEA;
          margin-bottom: 16px;
          border: 1px solid #D9E0E6;
        }
        .dist-ok { background: linear-gradient(90deg, #107E3E, #16A34A); height: 100%; transition: width 0.35s ease; }

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
        }
        .input-std:focus { border-color: #0070F2; box-shadow: 0 0 0 2px rgba(0, 112, 242, 0.15); }

        .chips-row {
          display: flex;
          flex-wrap: wrap;
          gap: 10px;
          margin-bottom: 12px;
          align-items: center;
          justify-content: space-between;
        }
        .chips-left { display: flex; gap: 6px; align-items: center; flex-wrap: wrap; }
        .chip {
          padding: 5px 12px;
          border-radius: 16px;
          font-size: 11px;
          font-weight: 700;
          cursor: pointer;
          border: 1px solid #D9E0E6;
          background: #FFFFFF;
          color: #32363A;
        }
        .chip-all.active-chip { background: #1D2D3E; color: #FFFFFF; border-color: #1D2D3E; }
        .chip-ok.active-chip { background: #107E3E; color: #FFFFFF; border-color: #107E3E; }
        .chip-err.active-chip { background: #BB0000; color: #FFFFFF; border-color: #BB0000; }
        .chip-prev.active-chip { background: #E9730C; color: #FFFFFF; border-color: #E9730C; }
        .chip-chk.active-chip { background: #6D28D9; color: #FFFFFF; border-color: #6D28D9; }
        .chip-info.active-chip { background: #0070F2; color: #FFFFFF; border-color: #0070F2; }

        .sev-badge {
          display: inline-flex;
          align-items: center;
          gap: 4px;
          padding: 3px 8px;
          border-radius: 12px;
          font-size: 10px;
          font-weight: 800;
        }
        .sev-blocking { background: #FEEAEA; color: #BB0000; border: 1px solid #F5C6C6; }
        .sev-preventive { background: #FFF4E5; color: #B45309; border: 1px solid #FCD34D; }
        .sev-check { background: #F5F3FF; color: #6D28D9; border: 1px solid #DDD6FE; }
        .sev-info { background: #EBF5FF; color: #005CB9; border: 1px solid #B3D4FF; }

        td.cell-sev-BLOCKING { background: #FFF2F2; border-left: 3px solid #BB0000; }
        td.cell-sev-PREVENTIVE { background: #FFFBEB; border-left: 3px solid #E9730C; }
        td.cell-sev-CHECK { background: #FAF5FF; border-left: 3px solid #6D28D9; }
        td.cell-sev-INFO { background: #F0F7FF; border-left: 3px solid #0070F2; }

        .alert-line {
          display: flex;
          align-items: center;
          justify-content: space-between;
          gap: 6px;
          margin-top: 3px;
          padding: 2px 5px;
          border-radius: 4px;
          font-size: 10px;
          font-weight: 700;
        }
        .alert-line-BLOCKING { color: #BB0000; background: rgba(187,0,0,0.06); }
        .alert-line-PREVENTIVE { color: #B45309; background: rgba(233,115,12,0.08); }
        .alert-line-CHECK { color: #6D28D9; background: rgba(109,40,217,0.08); }
        .alert-line-INFO { color: #005CB9; background: rgba(0,112,242,0.08); }

        .mini-fix-btn {
          border: 1px solid #107E3E;
          background: #E6F9EC;
          color: #107E3E;
          font-size: 9px;
          font-weight: 800;
          padding: 1px 5px;
          border-radius: 4px;
          cursor: pointer;
          white-space: nowrap;
        }
        .mini-fix-btn:hover { background: #107E3E; color: #FFFFFF; }

        .mini-chk-btn {
          border: 1px solid #6D28D9;
          background: #F5F3FF;
          color: #6D28D9;
          font-size: 9px;
          font-weight: 800;
          padding: 1px 5px;
          border-radius: 4px;
          cursor: pointer;
          white-space: nowrap;
        }
        .mini-chk-btn:hover { background: #6D28D9; color: #FFFFFF; }

        .tbl-wrap {
          max-height: 360px;
          overflow: auto;
          border: 1px solid #D9E0E6;
          border-radius: 8px;
          background: #FFFFFF;
        }
        table { width: 100%; border-collapse: collapse; font-size: 12px; text-align: left; }
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
        tbody tr:hover { background: #F8FAFC; }
        td { padding: 8px 10px; border-bottom: 1px solid #EDF0F2; vertical-align: top; }
        .editable-cell {
          outline: none;
          border: 1px solid transparent;
          border-bottom: 1px dashed #0070F2;
          display: inline-block;
          min-width: 65px;
          padding: 2px 5px;
          border-radius: 4px;
        }
        .editable-cell:focus { background: #FFFFFF; border: 1px solid #0070F2; }

        .th-content { display: flex; align-items: center; justify-content: space-between; gap: 6px; }
        .remove-col-btn { background: none; border: none; color: #BB0000; font-weight: 700; cursor: pointer; font-size: 11px; }

        .row-actions { display: flex; gap: 4px; justify-content: center; }
        .icon-btn {
          border: 1px solid #D9E0E6;
          background: #FFFFFF;
          border-radius: 4px;
          cursor: pointer;
          font-size: 11px;
          padding: 3px 7px;
          font-weight: 700;
        }
        .icon-btn:hover { border-color: #0070F2; color: #0070F2; background: #F0F6FF; }
        .icon-btn-del { border-color: #F5C6C6; color: #BB0000; background: #FEEAEA; }

        .wizard-footer {
          display: flex;
          justify-content: space-between;
          align-items: center;
          margin-top: 20px;
          border-top: 1px solid #D9E0E6;
          padding-top: 16px;
          flex-wrap: wrap;
          gap: 10px;
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
        }
        .sheet-pill-chk.sheet-active { background: #EBF5FF; border-color: #0070F2; color: #005CB9; }
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

        ${this._renderHelpModal()}

        <div class="sap-header">
          <div class="sap-brand">
            <img class="sap-logo-img" src="${this._sapLogoUri}" alt="SAP Logo" />
            <div class="sap-title-wrap">
              <div class="sap-title">Excel Planning Uploader Pro</div>
              <div class="sap-subtitle">Multi-Sheet Consolidation, Alert Semaphores & Full/Delta Engine</div>
            </div>
          </div>
          <div style="display:flex; align-items:center; gap:10px;">
            <span class="sap-badge-env">SAP Analytics Cloud</span>
            <button class="help-icon-btn" id="toggleHelpBtn" type="button" title="Cómo funciona este widget">
              <span style="background:#0070F2; color:#FFF; width:18px; height:18px; border-radius:50%; display:inline-flex; align-items:center; justify-content:center; font-size:11px; font-weight:900;">?</span>
              <span>Help</span>
            </button>
          </div>
        </div>

        ${this._toast ? `
          <div class="toast-banner toast-${this._toast.type}">
            <span>${this._esc(this._toast.msg)}</span>
          </div>
        ` : ''}

        <div class="nav-tabs">
          <div class="nav-tab ${this._activeTab === 'uploaderTab' ? 'active' : ''}" data-target="uploaderTab">Planning Upload Wizard</div>
          <div class="nav-tab ${this._activeTab === 'rulesTab' ? 'active' : ''}" data-target="rulesTab">⚙️ Validation Rules & Assisted Code (${activeRulesCount})</div>
          <div class="nav-tab ${this._activeTab === 'monitorTab' ? 'active' : ''}" data-target="monitorTab">Audit Monitor (${this._history.length})</div>
        </div>

        <div id="uploaderTab" class="tab-content" style="display:${this._activeTab === 'uploaderTab' ? 'block' : 'none'};">
          <div class="wizard-steps">
            <div class="step-item ${this._step === 1 ? 'active' : (this._filesQueue.length > 0 ? 'completed' : '')}" data-step="1">
              <div class="step-num">${this._filesQueue.length > 0 && this._step > 1 ? '✓' : '1'}</div>
              <div>
                <div class="step-label-title">Full/Delta & Files</div>
                <div class="step-label-sub">Strategy, multi-Excel & CSV</div>
              </div>
            </div>
            <div class="step-item ${this._step === 2 ? 'active' : (this._step > 2 ? 'completed' : '')}" data-step="2">
              <div class="step-num">${this._step > 2 ? '✓' : '2'}</div>
              <div>
                <div class="step-label-title">Sheets & Semaphores</div>
                <div class="step-label-sub">Tabs, scope & rules</div>
              </div>
            </div>
            <div class="step-item ${this._step === 3 ? 'active' : ''}" data-step="3">
              <div class="step-num">3</div>
              <div>
                <div class="step-label-title">Grid & Auto-Fix</div>
                <div class="step-label-sub">Resolution & SAC sync</div>
              </div>
            </div>
          </div>

          ${this._step === 1 ? `
            <div class="toolbar-box" style="background:#EBF5FF; border:1px solid #B3D4FF; margin-bottom:14px;">
              <div style="display:flex; align-items:center; gap:12px; width:100%; justify-content:space-between; flex-wrap:wrap;">
                <div>
                  <strong style="color:#005CB9; font-size:13px;">🔄 Planning Load Strategy:</strong>
                  <div style="font-size:11px; color:#556B82;">Choose whether this upload replaces current model data or applies incremental delta changes.</div>
                </div>
                <select id="loadStrategySel" class="input-std" style="font-weight:800; font-size:13px; color:#005CB9; padding:8px 12px;">
                  <option value="FULL" ${this._loadStrategy === 'FULL' ? 'selected' : ''}>FULL LOAD (Replace / Complete Model)</option>
                  <option value="DELTA" ${this._loadStrategy === 'DELTA' ? 'selected' : ''}>DELTA LOAD (Incremental / Adjustments)</option>
                </select>
              </div>
            </div>

            ${this._renderFormatConfigBar(false)}
            <div class="drop-zone" id="dropZone">
              <div class="upload-icon-circle">☁️</div>
              <h3 style="margin:0; color:#1D2D3E; font-size:16px;">Drag & drop your planning files here</h3>
              <p style="margin:0; color:#556B82; font-size:12px; max-width:560px;">
                Support for multiple workbooks, sheets, CSV and text files with configurable delimiters.
              </p>
              <div style="margin-top:8px;">
                <label class="file-label-btn" for="fileInput">
                  📂 Browse & Select Files
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
                  <h3 style="margin:0; color:#1D2D3E;">Loaded Workbooks & Sheet Selector (${this._filesQueue.length} file(s)) — Strategy: <span style="color:#005CB9;">${this._loadStrategy}</span></h3>
                  <span style="font-size:11px; color:#556B82;">Check or uncheck the sheets inside each file to consolidate:</span>
                </div>
                <div style="display:flex; gap:8px;">
                  <label class="btn" for="fileInput" style="cursor:pointer;">
                    + Add More Files
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

                      <div style="margin-top:10px; padding-top:10px; border-top:1px dashed #D9E0E6; display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:8px;">
                        <div style="display:flex; align-items:center; flex-wrap:wrap; gap:6px;">
                          <span style="font-size:11px; font-weight:700; color:#1D2D3E; margin-right:4px;">📑 Sheets:</span>
                          ${(item.sheets || []).map((sh, sIdx) => `
                            <label class="sheet-pill-chk ${sh.selected ? 'sheet-active' : ''}">
                              <input type="checkbox" class="sheet-toggle-chk" data-fidx="${fidx || fIdx}" data-sidx="${sIdx}" ${sh.selected ? 'checked' : ''} style="margin:0; cursor:pointer;" />
                              <span>${this._esc(sh.name)} <strong>(${sh.rows.length})</strong></span>
                            </label>
                          `).join('')}
                        </div>
                        <div style="display:flex; gap:4px;">
                          <button class="icon-btn select-all-sheets-btn" data-fidx="${fIdx}" data-mode="all" type="button">Select All</button>
                          <button class="icon-btn select-all-sheets-btn" data-fidx="${fIdx}" data-mode="none" type="button">None</button>
                        </div>
                      </div>
                    </div>
                  `;
                }).join('')}
              </div>

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
                    <div style="font-weight:700; color:#1D2D3E;">⚙️ Active Rules & Semaphores (${activeRulesCount}):</div>
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
                      <span style="display:inline-flex; align-items:center; gap:6px; background:#F8FAFC; border:1px solid #D9E0E6; padding:4px 8px; border-radius:6px; font-size:11px; font-weight:600;">
                        <input type="checkbox" class="toggle-rule-chk" data-rule-id="${this._esc(r.id)}" ${r.active ? 'checked' : ''} style="margin:0; cursor:pointer;" />
                        ${this._renderSemaphoreDot(r.semaphore || 'RED')}
                        <strong>${this._esc(r.col)}</strong>: ${this._esc(r.type)}
                        <button class="remove-col-btn delete-rule-btn" data-rule-id="${this._esc(r.id)}" title="Delete rule">✕</button>
                      </span>
                    `).join('')}
                  </div>
                </div>
              </div>

              ${this._showInlineRuleCreator ? this._renderRuleCreatorForm() : ''}
            </div>
            <div class="wizard-footer">
              <button class="btn btn-ghost" id="prevStepBtn" type="button">← Back</button>
              <button class="btn" id="nextStep3Btn" type="button">Validate & Edit ${totalRows} Rows →</button>
            </div>
          ` : ''}

          ${this._step === 3 ? `
            <div>
              <div class="metrics-bar">
                <div class="metric-box"><div class="metric-title">Strategy</div><div class="metric-val" style="font-size:14px; color:#0064D9;">${this._loadStrategy}</div></div>
                <div class="metric-box kpi-ok"><div class="metric-title">Clean Rows</div><div class="metric-val" style="color:#107E3E;">${stats.cleanRows}</div></div>
                <div class="metric-box kpi-err"><div class="metric-title">🛑 Bloqueantes</div><div class="metric-val" style="color:#BB0000;">${stats.blockingRows}</div></div>
                <div class="metric-box kpi-prev"><div class="metric-title">⚠️ Preventivas</div><div class="metric-val" style="color:#E9730C;">${stats.preventiveRows}</div></div>
                <div class="metric-box kpi-chk"><div class="metric-title">🔎 Chequeos</div><div class="metric-val" style="color:#6D28D9;">${stats.checkRows}</div></div>
                <div class="metric-box"><div class="metric-title">ℹ️ Informativas</div><div class="metric-val" style="color:#0070F2;">${stats.infoRows}</div></div>
              </div>

              <div class="dist-bar">
                <div class="dist-ok" style="width: ${validPercent}%;"></div>
              </div>

              <div class="toolbar-box">
                <div style="display:flex; align-items:center; gap:8px;">
                  <span style="font-weight:700; font-size:12px;">🔍 Quick Search:</span>
                  <input type="text" id="searchInput" class="input-std" style="width:190px;" placeholder="Filter by ID, text, sheet..." value="${this._esc(this._searchQuery)}" />
                </div>

                <div style="display:flex; align-items:center; gap:6px; flex-wrap:wrap;">
                  <span style="font-weight:700; font-size:12px;">⚡ Bulk Fill:</span>
                  <select id="bulkColSelect" class="input-std">
                    ${this._requiredColumns.map(c => `<option value="${this._esc(c)}" ${this._bulkCol === c ? 'selected' : ''}>${this._esc(c)}</option>`).join('')}
                  </select>
                  <input type="text" id="bulkValInput" class="input-std" style="width:110px;" placeholder="Value (e.g. 1000)" value="${this._esc(this._bulkVal)}" />
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
                  <div class="chip chip-all ${this._issueFilter === 'all' ? 'active-chip' : ''}" data-filter="all">All (${totalRows})</div>
                  <div class="chip chip-ok ${this._issueFilter === 'correct' ? 'active-chip' : ''}" data-filter="correct">Clean (${stats.cleanRows})</div>
                  <div class="chip chip-err ${this._issueFilter === 'BLOCKING' ? 'active-chip' : ''}" data-filter="BLOCKING">🛑 Bloqueantes (${stats.blockingRows})</div>
                  <div class="chip chip-prev ${this._issueFilter === 'PREVENTIVE' ? 'active-chip' : ''}" data-filter="PREVENTIVE">⚠️ Preventivas (${stats.preventiveRows})</div>
                  <div class="chip chip-chk ${this._issueFilter === 'CHECK' ? 'active-chip' : ''}" data-filter="CHECK">🔎 Chequeos (${stats.checkRows})</div>
                  <div class="chip chip-info ${this._issueFilter === 'INFO' ? 'active-chip' : ''}" data-filter="INFO">ℹ️ Info (${stats.infoRows})</div>
                </div>
                <div style="display:flex; gap:6px; align-items:center; flex-wrap:wrap;">
                  ${stats.fixableCells > 0 ? `
                    <button class="btn btn-success" id="autoFixAllBtn" type="button" style="font-size:11px; padding:6px 12px;">
                      ✨ Auto-Fix All (${stats.fixableCells})
                    </button>
                  ` : ''}
                  <button class="btn btn-ghost" id="addRowBtn" type="button" style="font-size:11px; padding:6px 12px;">+ Add Row</button>
                  <button class="btn btn-ghost" id="openAddColBtn" type="button" style="font-size:11px; padding:6px 12px;">+ Add Column</button>
                  <button class="btn btn-ghost" id="toggleInlineRuleBtn" type="button" style="font-size:11px; padding:6px 12px;">⚙️ + Rule / Code</button>
                  <button class="btn btn-ghost" id="exportCleanBtn" type="button" style="font-size:11px; padding:6px 12px;">📤 Export (.xlsx)</button>
                  ${(stats.blockingRows + stats.preventiveRows + stats.checkRows + stats.infoRows) > 0 ? `
                    <button class="btn btn-danger" id="downloadLogBtn" type="button" style="font-size:11px; padding:6px 12px;">📥 Alerts Log (.csv)</button>
                  ` : ''}
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
                          const cellAlerts = (row._alerts && row._alerts[col]) || [];
                          const editedInfo = row._editedCols && row._editedCols[col];
                          let topSev = '';
                          if (cellAlerts.some(a => a.severity === 'BLOCKING')) topSev = 'BLOCKING';
                          else if (cellAlerts.some(a => a.severity === 'PREVENTIVE')) topSev = 'PREVENTIVE';
                          else if (cellAlerts.some(a => a.severity === 'CHECK')) topSev = 'CHECK';
                          else if (cellAlerts.some(a => a.severity === 'INFO')) topSev = 'INFO';

                          return `
                            <td class="${topSev ? 'cell-sev-' + topSev : ''}" style="${!topSev && editedInfo ? 'background:#FFFBEB;' : ''}">
                              <span class="editable-cell" contenteditable="true" data-orig-idx="${origIdx}" data-col="${this._esc(col)}">${this._esc(row[col] ?? '')}</span>
                              ${cellAlerts.map(a => {
                                const dotHtml = this._renderSemaphoreDot(a.semaphore);
                                return `
                                  <div class="alert-line alert-line-${a.severity}">
                                    <span>${dotHtml} ${this._esc(a.msg)}</span>
                                    <span style="display:inline-flex; gap:3px;">
                                      ${a.autoFix ? `<button class="mini-fix-btn cell-autofix-btn" data-orig-idx="${origIdx}" data-col="${this._esc(col)}" data-rule-id="${this._esc(a.ruleId)}" type="button" title="Auto-Fix">✨ Fix</button>` : ''}
                                      ${(a.severity === 'CHECK' || a.severity === 'PREVENTIVE') ? `<button class="mini-chk-btn cell-check-ok-btn" data-orig-idx="${origIdx}" data-col="${this._esc(col)}" data-rule-id="${this._esc(a.ruleId)}" type="button" title="Approve">✓ OK</button>` : ''}
                                    </span>
                                  </div>
                                `;
                              }).join('')}
                              ${cellAlerts.length === 0 && editedInfo ? `<span style="display:block; font-size:9px; color:#B45309; font-weight:600; margin-top:2px;">✎ Edited (was: ${this._esc(editedInfo.from || 'Blank')})</span>` : ''}
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
              <div style="display:flex; align-items:center; gap:14px; flex-wrap:wrap;">
                ${stats.blockingRows === 0 && (stats.preventiveRows > 0 || stats.checkRows > 0) ? `
                  <label style="display:flex; align-items:center; gap:6px; background:#FFF8E6; border:1px solid #FCD34D; color:#B45309; padding:6px 12px; border-radius:6px; font-size:11px; font-weight:700; cursor:pointer;">
                    <input type="checkbox" id="ackNonBlockingChk" ${this._acknowledgeNonBlocking ? 'checked' : ''} style="margin:0; cursor:pointer;" />
                    <span>Confirm review of ${stats.preventiveRows} preventive and ${stats.checkRows} check alerts</span>
                  </label>
                ` : ''}
                <button class="btn btn-ghost" id="exitBtn" type="button">Reset Queue</button>
                <button class="btn" id="finishBtn" type="button" ${stats.blockingRows > 0 ? 'style="opacity:0.55;"' : ''}>
                  🚀 Consolidate & Synchronize with SAC (${this._loadStrategy})
                </button>
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
