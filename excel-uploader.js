/**
 * ExcelUploaderPro — SAP Analytics Cloud Custom Widget
 * Version: 2.5.0
 * Requires: SheetJS (loaded via CDN automatically if not present)
 * Tag:      <custom-excel-uploader-pro>
 *
 * Improvements in 2.2.0 (vs 2.1.0):
 *   • Removed DSP Artifact toggle from UI (sync always active)
 *   • Cleaner Full/Delta strategy bar with active badge
 *   • Visible scrollbars on all table wrappers and pre blocks
 *
 * Improvements in 2.1.0 (vs Gemini-generated v2):
 *   • Compiled-rule function cache — ASSISTED_CODE rules compiled once, reused on every validate pass
 *   • requestAnimationFrame debounce on search inputs — no full re-render on every keystroke
 *   • crypto.randomUUID() for collision-free rule IDs (fallback to Math.random)
 *   • Early exit in _validateRows() when grid is empty
 *   • User feedback (toast) when export functions have nothing to export
 *   • Cache invalidated automatically when rules are saved
 */
class ExcelUploaderPro extends HTMLElement {
  constructor() {
    super();
    this._shadow = this.attachShadow({ mode: 'open' });
    this._step = 1;
    this._activeTab = 'uploaderTab';
    this._filesQueue = [];
    this._rows = [];
    this._headers = [];
    this._currentUser = 'SAP User (Planning)'; // overridden by manifest property setter below
    this._isLoading = false;
    this._loadingText = 'Processing files...';
    this._xlsxReady = typeof XLSX !== 'undefined';
    this._showHelpModal = false;

    this._loadStrategy  = 'FULL';
    this._dspArtifactSync = true;
    this._targetVersion = this._storageGet('sac_target_version') || '';
    this._targetDate    = this._storageGet('sac_target_date')    || '';
    this._parseConfig = { format: 'AUTO', delimiter: 'AUTO', numberFormat: 'STD', headerRow: 1 };

    const defaultCols = ['ID', 'DESCRIPTION', 'H1', 'COMPANY', 'COSTCENTER', 'ASSET_CLASS', 'CAPITALIZED'];
    const savedCols = this._storageGet('sac_required_cols');
    this._requiredColumns = Array.isArray(savedCols) && savedCols.length > 0 ? savedCols : defaultCols;

    this._auditExportConfig = {
      open: false,
      numCols: this._requiredColumns.length,
      selectedCols: [...this._requiredColumns],
      includeComments: true,
      includeAuditDetail: true,
      format: 'XLSX'
    };

    const savedRules = this._storageGet('sac_validation_rules_v9');
    this._validationRules = Array.isArray(savedRules) ? savedRules : this._getDefaultRules();

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

    const savedHistory = this._storageGet('sac_audit_history_v9');
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
    this._editingScopeRuleId = null;
    this._bulkCol = this._requiredColumns.includes('COMPANY') ? 'COMPANY' : (this._requiredColumns[0] || 'ID');
    this._bulkVal = '';

    // Transform & Convert engine
    this._transformMap     = this._storageGet('sac_transform_map_v1') || {};
    this._convTables       = this._storageGet('sac_conv_tables_v1')   || {};
    this._convTabActiveCol = this._requiredColumns[0] || 'ID';

    const svgLogo = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 116 58"><defs><linearGradient id="sapGrad" x1="0%" y1="0%" x2="0%" y2="100%"><stop offset="0%" stop-color="#00B8F1"/><stop offset="50%" stop-color="#0083CA"/><stop offset="100%" stop-color="#0052A5"/></linearGradient></defs><polygon points="0,0 116,0 58,58 0,58" fill="url(#sapGrad)"/><text x="7" y="42" font-family="Arial Black, Arial, sans-serif" font-weight="900" font-size="33" letter-spacing="-1.5" fill="#FFFFFF">SAP</text></svg>';
    this._sapLogoUri = 'data:image/svg+xml;utf8,' + encodeURIComponent(svgLogo);

    this._compiledRules = {};
    this._searchRaf = null;
    this._auditSearchRaf = null;
    // ── SAC DIS connection ────────────────────────────────────────────────────
    this._sacTenantUrl    = this._storageGet('sac_dis_tenant_url') || '';
    this._sacModelId      = this._storageGet('sac_dis_model_id')   || 'PLAN_CAPEX_PLANNING';
    this._sacImportType   = this._storageGet('sac_import_type')    || 'privateFactData';
    this._sacModelName    = '';
    this._sacModelState   = '';
    this._sacModels       = [];
    this._sacModelsLoading = false;
    this._showModelPicker  = false;
    this._ensureXLSX();
    if (this._sacModelId) this._fetchModelInfo();
    this._render();
  }

  _getDefaultRules() {
    return [
      { id: 'RULE-1', col: 'ID', severity: 'BLOCKING', semaphore: 'RED', type: 'UNIQUE', param: '', min: '', max: '', autoFix: "value + '_' + row._rowNum", scopeFiles: 'ALL', startDate: '', endDate: '', active: true, desc: 'No duplicate IDs allowed across consolidated sheets' },
      { id: 'RULE-2', col: 'COMPANY', severity: 'BLOCKING', semaphore: 'RED', type: 'IN_LIST', param: '1000, 2000, 950', min: '', max: '', autoFix: "'1000'", scopeFiles: 'ALL', startDate: '', endDate: '', active: true, desc: 'Must match SAP Master Data Company Codes' },
      { id: 'RULE-3', col: 'COSTCENTER', severity: 'BLOCKING', semaphore: 'RED', type: 'IN_LIST', param: 'CC_IT_01, CC_FIN_02, 12234566, 10180000', min: '', max: '', autoFix: "'CC_IT_01'", scopeFiles: 'ALL', startDate: '', endDate: '', active: true, desc: 'Must match SAP Master Data Cost Centers' },
      { id: 'RULE-4', col: 'H1', severity: 'PREVENTIVE', semaphore: 'AMBER', type: 'NUMERIC_RANGE', param: '', min: '0', max: '100000', autoFix: 'num < 0 ? Math.abs(num) : 100000', scopeFiles: 'ALL', startDate: '', endDate: '', active: true, desc: 'Preventive budget threshold: H1 between 0 and 100,000' },
      { id: 'RULE-5', col: 'CAPITALIZED', severity: 'CHECK', semaphore: 'PURPLE', type: 'ASSISTED_CODE', param: "Number(row.H1 || 0) <= 50000 || String(value).toUpperCase() === 'Y' || String(value).toUpperCase() === 'X'", min: '', max: '', autoFix: "'Y'", scopeFiles: 'ALL', startDate: '', endDate: '', active: true, desc: 'Check: If H1 > 50,000, asset should be marked as CAPITALIZED (Y/X)' },
      { id: 'RULE-6', col: 'DESCRIPTION', severity: 'INFO', semaphore: 'BLUE', type: 'ASSISTED_CODE', param: 'String(value).trim().length >= 5', min: '', max: '', autoFix: "String(value).trim() + ' (PLANNING)'", scopeFiles: 'ALL', startDate: '', endDate: '', active: true, desc: 'Info: Recommended descriptive text of at least 5 characters' }
    ];
  }

  disconnectedCallback() {
    if (this._toastTimer) clearTimeout(this._toastTimer);
    if (this._searchRaf) cancelAnimationFrame(this._searchRaf);
    if (this._auditSearchRaf) cancelAnimationFrame(this._auditSearchRaf);
  }

  // ── SAC manifest property: currentUser ───────────────────────────────────────
  // SAC sets this after the element is connected. The setter syncs _currentUser
  // and triggers a re-render so the audit log always shows the real SAC user.
  get currentUser() { return this._currentUser; }
  set currentUser(val) {
    const v = String(val || '').trim();
    if (v && v !== this._currentUser) {
      this._currentUser = v;
      if (this._shadow && this._shadow.innerHTML) this._render();
    }
  }

  // ── SAC DIS connection properties (set by SAC manifest or Analytics Designer) ─
  get sacModelId() { return this._sacModelId; }
  set sacModelId(val) {
    const v = String(val || '').trim();
    // If SAC passes empty string (old story with blank default), keep localStorage or built-in default
    const effective = v || this._storageGet('sac_dis_model_id') || 'PLAN_CAPEX_PLANNING';
    if (effective !== this._sacModelId) {
      this._sacModelId    = effective;
      this._sacModelName  = '';
      this._sacModelState = 'connecting';
      this._storageSet('sac_dis_model_id', effective);
      if (this._shadow && this._shadow.innerHTML) this._fetchModelInfo();
    }
  }

  get sacTenantUrl() { return this._sacTenantUrl; }
  set sacTenantUrl(val) {
    const v = String(val || '').trim().replace(/\/+$/, '');
    if (v !== this._sacTenantUrl) {
      this._sacTenantUrl = v;
      this._storageSet('sac_dis_tenant_url', v || null);
    }
  }

  _storageGet(key) {
    try { const raw = localStorage.getItem(key); return raw ? JSON.parse(raw) : null; } catch (e) { return null; }
  }

  _storageSet(key, val) {
    try { if (val === null) localStorage.removeItem(key); else localStorage.setItem(key, JSON.stringify(val)); } catch (e) {
      if (e && e.name !== 'QuotaExceededError') console.warn('[ExcelUploaderPro] localStorage write failed:', e.message);
    }
  }

  _ensureXLSX() {
    if (typeof XLSX !== 'undefined') { this._xlsxReady = true; return; }
    let script = document.querySelector('script[data-sac-xlsx]');
    if (!script) {
      script = document.createElement('script');
      script.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
      script.async = true;
      script.setAttribute('data-sac-xlsx', 'true');
      document.head.appendChild(script);
    }
    script.addEventListener('load', () => { this._xlsxReady = true; this._render(); });
  }

  _esc(v) {
    return String(v ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  _logAction(type, target, detail, oldVal = '', newVal = '') {
    this._currentSessionLogs.unshift({ time: new Date().toLocaleTimeString(), type, target, detail, oldVal: String(oldVal ?? ''), newVal: String(newVal ?? '') });
  }

  _saveHistory() { this._storageSet('sac_audit_history_v9', this._history.slice(0, 30)); }
  _saveRules() { this._storageSet('sac_validation_rules_v9', this._validationRules); this._compiledRules = {}; }

  _notify(msg, type = 'info') {
    this._toast = { msg, type };
    this._render();
    if (this._toastTimer) clearTimeout(this._toastTimer);
    this._toastTimer = setTimeout(() => {
      this._toast = null;
      const banner = this._shadow.querySelector('.toast-banner');
      if (banner) banner.remove();
    }, 3600);
  }

  _syncExportColumnsWithSchema(availableCols) {
    const validSelected = (this._auditExportConfig.selectedCols || []).filter(c => availableCols.includes(c));
    this._auditExportConfig.selectedCols = validSelected.length > 0 ? validSelected : [...availableCols];
    this._auditExportConfig.numCols = this._auditExportConfig.selectedCols.length;
  }

  _recordCellModification(row, colName, oldVal, newVal) {
    if (!row._editedCols) row._editedCols = {};
    const initialFrom = row._editedCols[colName] !== undefined ? row._editedCols[colName].from : oldVal;
    if (initialFrom === newVal) delete row._editedCols[colName];
    else row._editedCols[colName] = { from: initialFrom, to: newVal };
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
            this._requiredColumns.forEach(reqCol => { if (r[reqCol] === undefined) r[reqCol] = ''; });
            sheetRows.push(r);
          });
        }
      });
    });
    this._rows = [...manualRows, ...sheetRows];
    this._rows.forEach((r, idx) => { r._rowNum = idx + 2; });
    // Auto-run transform pipeline if any rules are configured
    const hasTrf = Object.keys(this._transformMap).some(d => {
      const m = this._transformMap[d] || {};
      return (m.fixed && String(m.fixed).trim()) || (m.source && m.source !== d);
    });
    const hasConv = Object.keys(this._convTables).some(k => !k.includes('__') && (this._convTables[k] || []).length > 0);
    if (hasTrf || hasConv) this._applyTransformAndConvert();
    this._validateRows();
  }

  _evalAssistedCode(expr, row, colName, cacheKey) {
    const value = String(row[colName] ?? '').trim();
    const num = Number(value);
    // Always cache: use the provided key or fall back to the expression string itself
    const key = cacheKey || ('__expr__' + expr);
    if (!this._compiledRules[key]) {
      this._compiledRules[key] = new Function('value', 'num', 'row', '"use strict"; return (' + expr + ');');
    }
    return this._compiledRules[key](value, num, row);
  }

  _applyCellAutoFix(origIdx, colName, ruleId) {
    const row = this._rows[origIdx];
    if (!row) return;
    const alerts = (row._alerts && row._alerts[colName]) || [];
    const targetAlert = ruleId ? alerts.find(a => a.ruleId === ruleId) : alerts.find(a => a.autoFix);
    if (!targetAlert || !targetAlert.autoFix) return;
    try {
      const oldVal = String(row[colName] ?? '').trim();
      const newVal = String(this._evalAssistedCode(targetAlert.autoFix, row, colName, targetAlert.ruleId + '_f') ?? '').trim();
      if (oldVal !== newVal) {
        row[colName] = newVal;
        this._recordCellModification(row, colName, oldVal || '(Blank)', newVal);
        this._logAction('AUTO_FIX_CODE', 'Row #' + row._rowNum + ' [' + colName + ']', 'Assisted auto-fix (' + targetAlert.autoFix + ') applied', oldVal || '(Blank)', newVal);
        this._validateRows();
        this._notify('Auto-fixed Row #' + row._rowNum + ' [' + colName + '] -> ' + newVal, 'success');
      }
    } catch (e) { this._notify('Assisted auto-fix error: ' + e.message, 'error'); }
  }

  _applyAllAutoFixes() {
    let fixedCount = 0;
    let failedCount = 0;
    this._rows.forEach(row => {
      this._requiredColumns.forEach(col => {
        const fixable = ((row._alerts && row._alerts[col]) || []).find(a => a.autoFix);
        if (fixable) {
          try {
            const oldVal = String(row[col] ?? '').trim();
            const newVal = String(this._evalAssistedCode(fixable.autoFix, row, col, fixable.ruleId + '_f') ?? '').trim();
            if (oldVal !== newVal) {
              row[col] = newVal;
              this._recordCellModification(row, col, oldVal || '(Blank)', newVal);
              fixedCount++;
            }
          } catch (e) {
            failedCount++;
            console.warn('[ExcelUploaderPro] Auto-fix error — Row #' + row._rowNum + ' [' + col + ']:', e.message);
          }
        }
      });
    });
    if (fixedCount > 0) {
      this._logAction('BULK_AUTO_FIX', 'Grid Auto-Fix (' + fixedCount + ' cells)', 'Executed Assisted Code Auto-Fix across all fixable alerts', fixedCount + ' alerts', 'FIXED');
      this._validateRows();
      const failMsg = failedCount > 0 ? ' (' + failedCount + ' fix(es) failed — see console)' : '';
      this._notify('Assisted Code repaired ' + fixedCount + ' cell(s) automatically!' + failMsg, 'success');
    } else {
      this._notify('No cells with Assisted Auto-Fix rules pending.', 'info');
    }
  }

  _acknowledgeCheckCell(origIdx, colName, ruleId) {
    const row = this._rows[origIdx];
    if (!row) return;
    if (!row._checkedRules) row._checkedRules = {};
    row._checkedRules[colName + '__' + ruleId] = true;
    this._logAction('CHECK_VERIFIED', 'Row #' + row._rowNum + ' [' + colName + ']', 'User verified & approved check alert (' + ruleId + ')', 'Pending Check', 'Verified');
    this._validateRows();
    this._render();
  }


  // ── Persist transform config ─────────────────────────────────────────────────
  _saveTransformMap()   { this._storageSet('sac_transform_map_v1',  this._transformMap); }
  _saveConvTables()     { this._storageSet('sac_conv_tables_v1',    this._convTables);   }
  _saveTargetConfig()   { this._storageSet('sac_target_version', this._targetVersion); this._storageSet('sac_target_date', this._targetDate); }

  // ── Fetch real model metadata from SAC DIS ───────────────────────────────
  // Called whenever sacModelId changes; updates only the header subtitle inline.
  async _fetchModelInfo() {
    if (!this._sacModelId) {
      this._sacModelName  = '';
      this._sacModelState = '';
      this._updateHeaderSubtitle();
      return;
    }
    this._sacModelState = 'connecting';
    this._updateHeaderSubtitle();
    try {
      const base = (this._sacTenantUrl || window.location.origin).replace(/\/+$/, '');
      const resp = await fetch(base + '/api/v1/dataimport/models/' + encodeURIComponent(this._sacModelId), {
        method: 'GET',
        credentials: 'include',
        headers: { 'x-csrf-token': 'fetch', 'Accept': 'application/json' }
      });
      if (resp.status === 404) {
        // SAC DIS metadata endpoint not available — model is configured, treat as ready
        this._sacModelName  = this._sacModelId;
        this._sacModelState = 'ok';
      } else if (!resp.ok) {
        throw new Error('HTTP ' + resp.status + (resp.status === 401 || resp.status === 403 ? ' — Sin permisos DIS' : ''));
      } else {
        const meta = await resp.json().catch(() => ({}));
        this._sacModelName  = meta.description || meta.name || this._sacModelId;
        this._sacModelState = 'ok';
      }
    } catch (err) {
      this._sacModelName  = err.message || 'Error de conexión';
      this._sacModelState = 'error';
    }
    this._updateHeaderSubtitle();
  }

  // Updates only the header subtitle span — avoids full re-render
  _updateHeaderSubtitle() {
    const sub = this._shadow && this._shadow.querySelector('.sap-subtitle');
    if (!sub) return;
    const badge =
      this._sacModelState === 'connecting' ? ' &nbsp;|&nbsp; <span style="color:#A9B8C8;font-weight:600;">⏳ Conectando…</span>' :
      this._sacModelState === 'ok'         ? ' &nbsp;|&nbsp; <span style="color:#5BBFFF;font-weight:700;">📊 ' + this._esc(this._sacModelName || this._sacModelId) + '</span>' :
      this._sacModelState === 'error'      ? ' &nbsp;|&nbsp; <span style="color:#E67E22;font-weight:600;" title="' + this._esc(this._sacModelName) + '">⚠️ ' + this._esc(this._sacModelName.slice(0, 60)) + '</span>' :
      '';
    sub.innerHTML = 'Multi-Sheet Consolidation · Transform &amp; Convert · Alert Semaphores · v2.5.0' + badge;
  }

  // ── Fetch list of all models from SAC DIS ─────────────────────────────────
  async _fetchSacModels() {
    this._sacModelsLoading = true;
    this._showModelPicker  = true;
    this._sacModels        = [];
    this._render();
    try {
      const base = (this._sacTenantUrl || window.location.origin).replace(/\/+$/, '');
      const resp = await fetch(base + '/api/v1/dataimport/models', {
        method: 'GET', credentials: 'include',
        headers: { 'Accept': 'application/json', 'x-sap-dis-request-source': 'fileUploadWidget' }
      });
      if (!resp.ok) throw new Error('HTTP ' + resp.status);
      const data = await resp.json().catch(() => ({}));
      // Handle OData v2, v4, plain array formats
      const raw = Array.isArray(data) ? data
                : (data?.d?.results || data?.value || data?.models || []);
      this._sacModels = raw.map(m => ({
        id:         m.id         || m.modelId      || m.resourceId || '',
        name:       m.name       || m.modelName    || m.description || m.id || '',
        importType: m.importType || m.type         || ''
      })).filter(m => m.id);
    } catch (err) {
      this._sacModels = [];
      this._notify('No se pudo obtener modelos de SAC: ' + err.message, 'error');
    }
    this._sacModelsLoading = false;
    this._render();
  }

  // ── SAC Data Import Service (DIS) REST API — official SAP 5-step flow ───────
  // 1. CSRF  : GET /models (with x-sap-dis-request-source: fileUploadWidget)
  // 2. Create: POST /models/{id}/factData → { jobID }
  // 3. Upload: POST /jobs/{jobId}  Content-Type: text/csv
  // 4. Run   : POST /jobs/{jobId}/run
  // 5. Poll  : GET /jobs/{jobId}/status until COMPLETED / FAILED
  async _callDISApi(rows) {
    const base       = (this._sacTenantUrl || window.location.origin).replace(/\/+$/, '');
    const DIS        = base + '/api/v1/dataimport';
    const DIS_HEADER = { 'x-sap-dis-request-source': 'fileUploadWidget' };

    const _setLoading = (txt) => { this._loadingText = txt; this._render(); };

    // ── 1. Fetch CSRF token from GET /models list ─────────────────────────
    _setLoading('🔑 Obteniendo CSRF token de SAC DIS…');
    const csrfResp = await fetch(DIS + '/models', {
      method: 'GET', credentials: 'include',
      headers: { 'Content-Type': 'application/json', 'x-csrf-token': 'fetch', ...DIS_HEADER }
    });
    const csrfToken = csrfResp.headers.get('x-csrf-token') || csrfResp.headers.get('X-Csrf-Token') || '';

    // ── 2. Create import job: POST /models/{id}/factData ─────────────────
    _setLoading('📋 Creando job de importación en SAC…');
    const defaultValues = {};
    if (this._targetVersion) defaultValues['Version'] = this._targetVersion;
    if (this._targetDate)    defaultValues['Date']    = String(this._targetDate);

    const jobCreateResp = await fetch(DIS + '/models/' + encodeURIComponent(this._sacModelId) + '/' + (this._sacImportType || 'privateFactData'), {
      method: 'POST', credentials: 'include',
      headers: { 'Content-Type': 'application/json', 'x-csrf-token': csrfToken, ...DIS_HEADER },
      body: JSON.stringify({ Mapping: {}, DefaultValues: defaultValues, JobSettings: {} })
    });
    if (!jobCreateResp.ok) {
      let msg = 'Create job HTTP ' + jobCreateResp.status;
      try { const b = await jobCreateResp.json(); msg = b?.error?.message || b?.message || JSON.stringify(b); } catch (_) {}
      throw new Error(msg);
    }
    const jobCreateBody = await jobCreateResp.json().catch(() => ({}));
    const jobId = jobCreateBody.jobID || jobCreateBody.jobId || jobCreateBody.id;
    if (!jobId) throw new Error('SAC DIS no devolvió jobID: ' + JSON.stringify(jobCreateBody));

    // ── 3. Upload CSV data to job: POST /jobs/{jobId} ─────────────────────
    _setLoading('📤 Subiendo ' + rows.length + ' filas a SAC DIS…');

    // Build columns: required columns, with Version/Date prepended when set
    const cols = [...this._requiredColumns];
    if (this._targetVersion && !cols.includes('Version')) cols.unshift('Version');
    if (this._targetDate    && !cols.includes('Date'))    cols.unshift('Date');

    const csvLines = [cols.join(',')];
    for (const r of rows) {
      const vals = cols.map(col => {
        let v;
        if      (col === 'Version') v = this._targetVersion || '';
        else if (col === 'Date')    v = this._targetDate ? String(this._targetDate) : '';
        else                        v = r[col] ?? '';
        v = String(v).trim();
        if (v.includes(',') || v.includes('"') || v.includes('\n')) {
          v = '"' + v.replace(/"/g, '""') + '"';
        }
        return v;
      });
      csvLines.push(vals.join(','));
    }
    const csvBody = csvLines.join('\n');

    const csvResp = await fetch(DIS + '/jobs/' + jobId, {
      method: 'POST', credentials: 'include',
      headers: { 'Content-Type': 'text/csv', 'x-csrf-token': csrfToken, ...DIS_HEADER },
      body: csvBody
    });
    if (!csvResp.ok) {
      let msg = 'Upload CSV HTTP ' + csvResp.status;
      try { const b = await csvResp.json(); msg = b?.error?.message || b?.message || JSON.stringify(b); } catch (_) {}
      throw new Error(msg);
    }

    // ── 4. Run the job: POST /jobs/{jobId}/run ────────────────────────────
    _setLoading('🚀 Aplicando datos al modelo ' + this._sacModelId + '…');
    const runResp = await fetch(DIS + '/jobs/' + jobId + '/run', {
      method: 'POST', credentials: 'include',
      headers: { 'x-csrf-token': csrfToken, ...DIS_HEADER }
    });
    if (!runResp.ok) {
      let msg = 'Run HTTP ' + runResp.status;
      try { const b = await runResp.json(); msg = b?.error?.message || b?.message || msg; } catch (_) {}
      throw new Error(msg);
    }

    // ── 5. Poll GET /jobs/{jobId}/status until COMPLETED ─────────────────
    _setLoading('⏳ Esperando confirmación de SAC DIS…');
    return await this._pollJobStatus(
      DIS + '/jobs/' + jobId + '/status',
      ['COMPLETED', 'COMPLETED_WITH_FAILURES', 'COMPLETE'],
      ['FAILED'],
      120, 3000, _setLoading, DIS_HEADER
    );
  }

  // Polls a DIS /status URL until one of the target statuses is reached.
  async _pollJobStatus(statusUrl, targetStatuses, failStatuses, maxAttempts, intervalMs, _setLoading, extraHeaders) {
    for (let i = 0; i < maxAttempts; i++) {
      await new Promise(r => setTimeout(r, intervalMs));
      const r = await fetch(statusUrl, {
        method: 'GET', credentials: 'include',
        headers: { 'Accept': 'application/json', ...(extraHeaders || {}) }
      });
      if (!r.ok) continue;
      const body   = await r.json().catch(() => ({}));
      const status = (body.jobStatus || body.status || '').toUpperCase();
      if (targetStatuses.some(s => status === s || status.startsWith(s))) return body;
      if ((failStatuses || []).some(s => status === s.toUpperCase())) {
        throw new Error('SAC DIS job ' + status + ': ' + (body.message || body.errorMessage || JSON.stringify(body)));
      }
      if (_setLoading) _setLoading('⏳ SAC DIS: ' + (status || 'procesando') + '…');
    }
    throw new Error('Timeout esperando SAC DIS (URL: ' + statusUrl + ')');
  }

  // ── Core conversion resolver ─────────────────────────────────────────────────
  _applyConversionToValue(rawVal, rules) {
    for (let i = 0; i < rules.length; i++) {
      const rule = rules[i];
      const mode = (rule.mode || 'EXACT').toUpperCase();
      const src  = String(rule.source || '');
      const tgt  = String(rule.target || '');

      if (mode === 'EXACT' && rawVal === src) {
        return { out: tgt, hit: true, rule };
      }
      if (mode === 'SLICE') {
        const start = rule.sliceStart !== undefined && rule.sliceStart !== '' ? Number(rule.sliceStart) : 0;
        const end   = rule.sliceEnd   !== undefined && rule.sliceEnd   !== '' ? Number(rule.sliceEnd)   : undefined;
        return { out: rawVal.slice(start, end), hit: true, rule };
      }
      if (mode === 'REGEX') {
        try {
          const re = new RegExp(src);
          if (re.test(rawVal)) return { out: rawVal.replace(re, tgt), hit: true, rule };
        } catch (e) { /* invalid regex — skip rule */ }
      }
      if (mode === 'PASSTHROUGH' || src === '*') {
        return { out: rawVal, hit: false, rule };
      }
      if (mode === 'DEFAULT') {
        return { out: tgt, hit: true, rule };
      }
    }
    return { out: rawVal, hit: false, noMatch: true };
  }

  // ── Full Transform & Convert pipeline ────────────────────────────────────────
  _applyTransformAndConvert() {
    const headers = this._headers;

    this._rows.forEach(row => {
      if (!row._originalValues) row._originalValues = {};
      if (!row._convLog)        row._convLog        = {};

      // 1. Column Mapping — remap or inject fixed values
      this._requiredColumns.forEach(dim => {
        const map = this._transformMap[dim];
        if (!map) return;
        const fixedVal = String(map.fixed || '').trim();
        const srcCol   = String(map.source || '').trim();
        if (fixedVal !== '') {
          if (!row[dim] || String(row[dim]).trim() === '') {
            row._convLog[dim] = { from: row[dim] || '(Empty)', to: fixedVal, mode: 'FIXED' };
            row[dim] = fixedVal;
          }
        } else if (srcCol && srcCol !== dim && headers.includes(srcCol)) {
          const srcVal = String(row[srcCol] ?? '').trim();
          if (srcVal !== '') {
            row._originalValues[dim] = row[dim];
            row._convLog[dim] = { from: row[dim] || '(Empty)', to: srcVal, mode: 'REMAP' };
            row[dim] = srcVal;
          }
        }
      });

      // 2. Conversion Rules — per-dimension lookup
      this._requiredColumns.forEach(col => {
        const rules = this._convTables[col];
        if (!rules || rules.length === 0) return;
        const rawVal = String(row[col] ?? '').trim();
        if (!rawVal) return;

        const { out, hit, noMatch, rule } = this._applyConversionToValue(rawVal, rules);

        if (hit && out !== rawVal) {
          row._originalValues[col] = rawVal;
          row._convLog[col] = { from: rawVal, to: out, mode: rule ? rule.mode : '—' };
          row[col] = out;
        }

        if (noMatch) {
          const fallback = String(this._convTables[col + '__fallback'] || 'PASSTHROUGH').toUpperCase();
          const defVal   = String(this._convTables[col + '__default']  || '').trim();
          if (fallback === 'ALERT') {
            this._pushAlert(row, col, {
              ruleId: 'CONV-NO-MATCH', severity: 'BLOCKING', semaphore: 'RED',
              msg: 'No conversion match: ' + rawVal, autoFix: ''
            });
          } else if (fallback === 'DEFAULT' && defVal) {
            row._originalValues[col] = rawVal;
            row._convLog[col] = { from: rawVal, to: defVal, mode: 'DEFAULT' };
            row[col] = defVal;
          } else if (fallback === 'REJECT') {
            row._rejected = true;
          }
        }
      });
    });

    // Remove rejected rows and renumber
    const before = this._rows.length;
    this._rows = this._rows.filter(r => !r._rejected);
    this._rows.forEach((r, idx) => { r._rowNum = idx + 2; });
    if (this._rows.length < before) {
      this._logAction('CONV_REJECT', 'Conversion Engine',
        (before - this._rows.length) + ' row(s) rejected (no conversion match)',
        before + ' rows', this._rows.length + ' rows');
    }
  }

  // ── Import conversion file (Excel / CSV / TXT) ───────────────────────────────
  async _importConversionFile(file) {
    const dim = this._convTabActiveCol;
    if (!dim) return;
    try {
      const ext = (file.name.split('.').pop() || '').toUpperCase();
      let importedRules = [];

      const parseAOA = (aoa) => {
        if (aoa.length < 2) return [];
        return aoa.slice(1).map(r => ({
          mode:   String(r[2] || 'EXACT').trim().toUpperCase() || 'EXACT',
          source: String(r[0] || '').trim(),
          target: String(r[1] || '').trim()
        })).filter(r => r.source);
      };

      if (['CSV', 'TXT'].includes(ext)) {
        const text = await this._readFileAsText(file);
        importedRules = parseAOA(this._parseDelimitedTextToAOA(text, 'AUTO'));
      } else {
        if (typeof XLSX === 'undefined') { this._notify('XLSX engine not ready.', 'error'); return; }
        const buffer = await this._readFileAsArrayBuffer(file);
        const wb  = XLSX.read(new Uint8Array(buffer), { type: 'array' });
        const aoa = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 });
        importedRules = parseAOA(aoa);
      }

      if (importedRules.length === 0) { this._notify('No valid rows in conversion file (expected: SOURCE, TARGET, MODE).', 'error'); return; }
      if (!this._convTables[dim]) this._convTables[dim] = [];
      this._convTables[dim].push(...importedRules);
      this._saveConvTables();
      this._logAction('CONV_IMPORT', 'Conversion Engine [' + dim + ']',
        'Imported ' + importedRules.length + ' rule(s) from ' + file.name, '', importedRules.length + ' rules');
      this._notify('Imported ' + importedRules.length + ' rule(s) for ' + dim + '.', 'success');
      this._render();
    } catch (e) {
      this._notify('Import error: ' + e.message, 'error');
    }
  }


  // ── Transform & Convert tab renderer ─────────────────────────────────────────
  _renderTransformConvertTab() {
    const availHeaders = this._headers.length > 0 ? this._headers : this._requiredColumns;
    const activeDim    = this._requiredColumns.includes(this._convTabActiveCol)
      ? this._convTabActiveCol : (this._requiredColumns[0] || 'ID');
    const activeRules  = this._convTables[activeDim] || [];
    const fallback     = String(this._convTables[activeDim + '__fallback'] || 'PASSTHROUGH');
    const defVal       = String(this._convTables[activeDim + '__default']  || '');
    const convertDims  = this._requiredColumns.filter(d => (this._convTables[d] || []).length > 0).length;
    const mappedDims   = this._requiredColumns.filter(d => {
      const m = this._transformMap[d] || {};
      return (m.fixed && m.fixed.trim()) || (m.source && m.source !== d);
    }).length;

    return `
      <div>
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:16px;flex-wrap:wrap;gap:10px;">
          <div>
            <h3 style="margin:0;color:#1D2D3E;">🔄 Transform &amp; Convert Engine</h3>
            <p style="margin:4px 0 0;color:#556B82;font-size:12px;">
              Map source columns → model dimensions and apply conversion rules (EXACT · SLICE · REGEX · PASSTHROUGH) <strong>before</strong> validation. <span style="color:#B45309;">⚠️ SLICE: to get last 6 chars of a 10-digit code → Start=4, End=empty (not End=6).</span>
              ${mappedDims > 0 || convertDims > 0 ? '<span style="color:#107E3E;font-weight:700;">Active: ' + mappedDims + ' mapping(s) · ' + convertDims + ' dimension(s) with conversion rules.</span>' : '<span style="color:#788FA6;">No rules configured yet.</span>'}
            </p>
          </div>
          <button class="btn btn-success" id="runTransformBtn" type="button" ${this._rows.length === 0 ? 'disabled style="opacity:0.55;"' : ''}>
            ▶ Apply to ${this._rows.length} Row(s)
          </button>
        </div>

        <!-- Section 1: Column Mapping -->
        <div style="background:#FFFFFF;border:1px solid #D9E0E6;border-radius:8px;padding:16px;margin-bottom:16px;">
          <div style="font-weight:800;color:#005CB9;font-size:13px;margin-bottom:12px;">📐 Column Mapping — Rename or inject fixed values</div>
          <div style="border:1px solid #D9E0E6;border-radius:8px;overflow:hidden;">
            <table>
              <thead>
                <tr>
                  <th style="width:130px;">Model Dimension</th>
                  <th>Source Column (from file)</th>
                  <th style="width:200px;">Fixed Override</th>
                  <th style="width:100px;text-align:center;">Status</th>
                </tr>
              </thead>
              <tbody>
                ${this._requiredColumns.map(dim => {
                  const map = this._transformMap[dim] || {};
                  const hasFixed = map.fixed && String(map.fixed).trim() !== '';
                  const srcMatch = map.source && map.source !== dim && availHeaders.includes(map.source);
                  const status   = hasFixed ? 'FIXED' : (srcMatch ? 'REMAP' : 'PASS');
                  const statusColor = status === 'REMAP' ? '#107E3E' : (status === 'FIXED' ? '#005CB9' : '#788FA6');
                  const statusBg    = status === 'REMAP' ? '#E6F9EC' : (status === 'FIXED' ? '#EBF5FF' : '#F0F3F6');
                  return '<tr>' +
                    '<td><span class="col-pill">' + this._esc(dim) + '</span></td>' +
                    '<td><select class="input-std trf-src-sel" data-dim="' + this._esc(dim) + '" style="width:100%;max-width:280px;">' +
                      '<option value="">— Same name (' + this._esc(dim) + ') —</option>' +
                      availHeaders.filter(h => h !== dim).map(h =>
                        '<option value="' + this._esc(h) + '"' + (map.source === h ? ' selected' : '') + '>' + this._esc(h) + '</option>'
                      ).join('') +
                    '</select></td>' +
                    '<td><input type="text" class="input-std trf-fixed-inp" data-dim="' + this._esc(dim) + '" style="width:100%;box-sizing:border-box;" placeholder="e.g. 202601 or 1000" value="' + this._esc(map.fixed || '') + '" /></td>' +
                    '<td style="text-align:center;"><span style="background:' + statusBg + ';color:' + statusColor + ';padding:3px 10px;border-radius:12px;font-size:10px;font-weight:800;">' + status + '</span></td>' +
                  '</tr>';
                }).join('')}
              </tbody>
            </table>
          </div>
        </div>

        <!-- Section 2: Conversion Rules -->
        <div style="background:#FFFFFF;border:1px solid #D9E0E6;border-radius:8px;padding:16px;">
          <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:12px;flex-wrap:wrap;gap:8px;">
            <div style="font-weight:800;color:#005CB9;font-size:13px;">🔁 Conversion Rules — per Dimension</div>
            <div style="display:flex;gap:6px;align-items:center;">
              <span style="font-size:10px;color:#556B82;font-weight:600;">File format: SOURCE_VALUE, TARGET_VALUE, MODE (optional)</span>
              <button class="btn btn-ghost" id="convImportBtn" type="button" style="font-size:11px;padding:6px 10px;">📂 Import File</button>
              <input type="file" id="convFileInput" accept=".xlsx,.csv,.txt" style="display:none;" />
            </div>
          </div>

          <div style="display:flex;flex-wrap:wrap;gap:6px;margin-bottom:14px;">
            ${this._requiredColumns.map(dim => {
              const ruleCount = (this._convTables[dim] || []).length;
              const isActive  = dim === activeDim;
              return '<button type="button" class="conv-dim-btn" data-dim="' + this._esc(dim) + '" style="padding:5px 14px;border-radius:16px;font-size:11px;font-weight:700;cursor:pointer;border:1px solid ' + (isActive ? '#0070F2' : '#D9E0E6') + ';background:' + (isActive ? '#0070F2' : '#FFFFFF') + ';color:' + (isActive ? '#FFFFFF' : '#32363A') + ';transition:all 0.15s;">' +
                this._esc(dim) + (ruleCount > 0 ? ' <span style="background:' + (isActive ? 'rgba(255,255,255,0.3)' : '#E5E9EC') + ';padding:0 5px;border-radius:8px;">' + ruleCount + '</span>' : '') +
              '</button>';
            }).join('')}
          </div>

          <div style="display:flex;align-items:center;gap:10px;margin-bottom:10px;flex-wrap:wrap;padding:8px 10px;background:#F8FAFC;border-radius:8px;border:1px solid #D9E0E6;">
            <span style="font-weight:700;font-size:12px;color:#1D2D3E;">Rules for <span class="col-pill" style="font-size:11px;">${this._esc(activeDim)}</span></span>
            <label style="font-size:11px;font-weight:700;color:#556B82;">Fallback:</label>
            <select class="input-std conv-fallback-sel" data-dim="${this._esc(activeDim)}" style="font-size:11px;">
              <option value="PASSTHROUGH" ${fallback === 'PASSTHROUGH' ? 'selected' : ''}>Keep Original (PASSTHROUGH)</option>
              <option value="ALERT"       ${fallback === 'ALERT'       ? 'selected' : ''}>🔴 Blocking Alert</option>
              <option value="DEFAULT"     ${fallback === 'DEFAULT'     ? 'selected' : ''}>Use Default Value</option>
              <option value="REJECT"      ${fallback === 'REJECT'      ? 'selected' : ''}>❌ Reject Row</option>
            </select>
            ${fallback === 'DEFAULT' ? '<input type="text" class="input-std conv-default-inp" data-dim="' + this._esc(activeDim) + '" style="width:110px;" placeholder="Default value" value="' + this._esc(defVal) + '" />' : ''}
          </div>

          ${activeRules.length > 0 ? `
            <div style="border:1px solid #D9E0E6;border-radius:8px;overflow:hidden;margin-bottom:10px;">
              <table>
                <thead>
                  <tr>
                    <th style="width:130px;">Mode</th>
                    <th>Source Pattern</th>
                    <th>Target Value</th>
                    <th style="width:76px;">Slice From</th>
                    <th style="width:76px;">Slice To</th>
                    <th style="width:48px;text-align:center;">Del</th>
                  </tr>
                </thead>
                <tbody>
                  ${activeRules.map((rule, rIdx) => {
                    const isSlice  = rule.mode === 'SLICE';
                    const noSrc    = isSlice || rule.mode === 'PASSTHROUGH' || rule.mode === 'DEFAULT';
                    const noTgt    = rule.mode === 'PASSTHROUGH';
                    return '<tr>' +
                      '<td><select class="input-std conv-rule-mode" data-dim="' + this._esc(activeDim) + '" data-ridx="' + rIdx + '" style="width:100%;font-size:11px;">' +
                        ['EXACT','SLICE','REGEX','PASSTHROUGH','DEFAULT'].map(m =>
                          '<option value="' + m + '"' + (rule.mode === m ? ' selected' : '') + '>' + m + '</option>'
                        ).join('') +
                      '</select></td>' +
                      '<td><input type="text" class="input-std conv-rule-src" data-dim="' + this._esc(activeDim) + '" data-ridx="' + rIdx + '" style="width:100%;box-sizing:border-box;font-family:monospace;font-size:11px;' + (noSrc ? 'opacity:0.4;' : '') + '" value="' + this._esc(rule.source || '') + '" placeholder="e.g. 0000123456 or ^0+" ' + (noSrc ? 'disabled' : '') + ' /></td>' +
                      '<td><input type="text" class="input-std conv-rule-tgt" data-dim="' + this._esc(activeDim) + '" data-ridx="' + rIdx + '" style="width:100%;box-sizing:border-box;font-family:monospace;font-size:11px;' + (noTgt ? 'opacity:0.4;' : '') + '" value="' + this._esc(rule.target || '') + '" placeholder="e.g. 123456" ' + (noTgt ? 'disabled' : '') + ' /></td>' +
                      '<td><input type="number" class="input-std conv-rule-start" data-dim="' + this._esc(activeDim) + '" data-ridx="' + rIdx + '" style="width:64px;font-size:11px;' + (!isSlice ? 'opacity:0.4;' : '') + '" value="' + (rule.sliceStart !== undefined ? rule.sliceStart : '') + '" placeholder="0" ' + (!isSlice ? 'disabled' : '') + ' /></td>' +
                      '<td><input type="number" class="input-std conv-rule-end"   data-dim="' + this._esc(activeDim) + '" data-ridx="' + rIdx + '" style="width:64px;font-size:11px;' + (!isSlice ? 'opacity:0.4;' : '') + '" value="' + (rule.sliceEnd   !== undefined ? rule.sliceEnd   : '') + '" placeholder="6" '  + (!isSlice ? 'disabled' : '') + ' /></td>' +
                      '<td style="text-align:center;"><button class="icon-btn icon-btn-del conv-del-rule-btn" data-dim="' + this._esc(activeDim) + '" data-ridx="' + rIdx + '" type="button">✕</button></td>' +
                    '</tr>';
                  }).join('')}
                </tbody>
              </table>
            </div>
          ` : '<div style="text-align:center;padding:18px;color:#788FA6;font-size:12px;border:1px dashed #D9E0E6;border-radius:8px;margin-bottom:10px;">No rules for <strong>' + this._esc(activeDim) + '</strong>. Add one below or import a file.</div>'}

          <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;background:#F8FAFC;padding:10px 12px;border-radius:8px;border:1px solid #D9E0E6;">
            <span style="font-size:11px;font-weight:700;color:#556B82;white-space:nowrap;">+ Add Rule:</span>
            <select class="input-std" id="newConvMode" style="font-size:11px;width:120px;">
              <option value="EXACT">EXACT</option>
              <option value="SLICE">SLICE</option>
              <option value="REGEX">REGEX</option>
              <option value="PASSTHROUGH">PASSTHROUGH</option>
              <option value="DEFAULT">DEFAULT</option>
            </select>
            <input type="text" class="input-std" id="newConvSrc" placeholder="Source pattern" style="width:150px;font-family:monospace;font-size:11px;" />
            <input type="text" class="input-std" id="newConvTgt" placeholder="Target value"   style="width:130px;font-family:monospace;font-size:11px;" />
            <input type="number" class="input-std" id="newConvStart" placeholder="Start (0=first)" style="width:88px;font-size:11px;" title="SLICE start index (0 = first char). To get LAST 6 of 10 digits: Start=4, End empty." />
            <input type="number" class="input-std" id="newConvEnd"   placeholder="End (empty=last)" style="width:106px;font-size:11px;" title="SLICE end index — leave empty to take all chars from Start to end. Example last 6 of 10: Start=4, End empty." />
            <button class="btn" id="addConvRuleBtn" type="button" data-dim="${this._esc(activeDim)}">+ Add</button>
          </div>
        </div>
      </div>
    `;
  }

  // Updates only the badge span inline — avoids full re-render (which resets focus)
  _refreshBadge(shadow) {
    const badge = shadow.querySelector('.sac-target-badge');
    if (!badge) return;
    if (!this._targetVersion || !this._targetDate) {
      badge.innerHTML = '⚠️ Configura Version y Period antes de sincronizar';
      badge.style.cssText = 'background:#FFF8E6;color:#B45309;border:1px solid #FCD34D;padding:5px 14px;border-radius:20px;font-size:11px;font-weight:700;';
      return;
    }
    const base = '✅ ' + this._esc(this._targetVersion) + ' · ' + this._esc(this._targetDate);
    if (this._sacModelId) {
      badge.innerHTML = base + ' → <strong style="color:#5BBFFF;">' + this._esc(this._sacModelId) + '</strong>';
      badge.style.cssText = 'background:#E6F9EC;color:#107E3E;border:1px solid #A3E6BF;padding:5px 14px;border-radius:20px;font-size:11px;font-weight:700;';
    } else {
      badge.innerHTML = base + ' <span style="color:#B45309;font-weight:600;">(sin Model ID — solo evento)</span>';
      badge.style.cssText = 'background:#E6F9EC;color:#107E3E;border:1px solid #A3E6BF;padding:5px 14px;border-radius:20px;font-size:11px;font-weight:700;';
    }
  }

  _bindEvents() {
    const shadow = this._shadow;

    const toggleHelpBtn = shadow.getElementById('toggleHelpBtn');
    const closeHelpModalBtn = shadow.getElementById('closeHelpModalBtn');
    if (toggleHelpBtn) toggleHelpBtn.addEventListener('click', () => { this._showHelpModal = !this._showHelpModal; this._render(); });
    if (closeHelpModalBtn) closeHelpModalBtn.addEventListener('click', () => { this._showHelpModal = false; this._render(); });

    shadow.querySelectorAll('.strategy-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const strat = e.currentTarget.dataset.strategy;
        if (strat && this._loadStrategy !== strat) {
          const prev = this._loadStrategy;
          this._loadStrategy = strat;
          this._logAction('STRATEGY_CHANGE', 'Load Strategy', 'Planning mode switched from ' + prev + ' to ' + strat, prev, strat);
          this._render();
        }
      });
    });

    // SAC Artifact sync is always active

    // ── SAC Load Target: Version & Period ─────────────────────────────────────
    const targetVersionInp = shadow.getElementById('targetVersionInp');
    const targetDateInp    = shadow.getElementById('targetDateInp');
    if (targetVersionInp) {
      targetVersionInp.addEventListener('input', (e) => {
        this._targetVersion = e.target.value.trim();
        this._saveTargetConfig();
        this._refreshBadge(shadow);
      });
      targetVersionInp.addEventListener('change', () => this._render());
    }
    if (targetDateInp) {
      targetDateInp.addEventListener('input', (e) => {
        this._targetDate = e.target.value.trim();
        this._targetDate = e.target.value.trim();
        this._saveTargetConfig();
        this._refreshBadge(shadow);
      });
      targetDateInp.addEventListener('change', () => this._render());
    }

    const sacModelIdInp = shadow.getElementById('sacModelIdInp');
    if (sacModelIdInp) {
      sacModelIdInp.addEventListener('input', (e) => {
        const v = e.target.value.trim();
        this._sacModelId    = v;
        this._sacModelName  = '';
        this._sacModelState = v ? 'connecting' : '';
        this._storageSet('sac_dis_model_id', v || null);
        this._refreshBadge(shadow);
      });
      sacModelIdInp.addEventListener('change', () => this._fetchModelInfo());
    }

    const browseModelsBtn = shadow.getElementById('browseModelsBtn');
    if (browseModelsBtn) {
      browseModelsBtn.addEventListener('click', () => this._fetchSacModels());
    }

    const closeModelPickerBtn = shadow.getElementById('closeModelPickerBtn');
    if (closeModelPickerBtn) {
      closeModelPickerBtn.addEventListener('click', () => { this._showModelPicker = false; this._render(); });
    }

    shadow.querySelectorAll('.sac-model-pick-row').forEach(row => {
      row.addEventListener('mouseenter', () => { row.style.background = '#EBF5FF'; });
      row.addEventListener('mouseleave', () => { row.style.background = ''; });
      row.addEventListener('click', () => {
        const modelId    = row.dataset.modelId;
        const importType = row.dataset.importType || 'privateFactData';
        this._sacModelId    = modelId;
        this._sacImportType = importType;
        this._sacModelName  = '';
        this._sacModelState = 'connecting';
        this._showModelPicker = false;
        this._storageSet('sac_dis_model_id', modelId);
        this._storageSet('sac_import_type',  importType);
        this._fetchModelInfo();
        this._render();
      });
    });

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
      dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('drag-active'); });
      dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-active'));
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
    if (cfgHeaderRow) cfgHeaderRow.addEventListener('change', (e) => { this._parseConfig.headerRow = Math.max(1, parseInt(e.target.value, 10) || 1); });
    if (reparseFilesBtn) reparseFilesBtn.addEventListener('click', () => this._reparseAllFilesWithNewConfig());

    shadow.querySelectorAll('.step-item').forEach(stepEl => {
      stepEl.addEventListener('click', (e) => {
        const targetStep = parseInt(e.currentTarget.dataset.step, 10);
        if (targetStep === 1 || this._filesQueue.length > 0 || this._rows.length > 0) {
          this._step = targetStep;
          this._render();
        } else {
          this._notify('Please upload at least one Excel or CSV file first.', 'info');
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
          this._logAction(sh.selected ? 'SHEET_ENABLED' : 'SHEET_DISABLED', fileItem.file.name + ' [' + sh.name + ']', (sh.selected ? 'Included' : 'Excluded') + ' sheet (' + sh.rows.length + ' rows)', !sh.selected, sh.selected);
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
          this._logAction('SHEET_BULK_TOGGLE', fileItem.file.name, (mode ? 'Selected all' : 'Deselected all') + ' sheets', '', mode ? 'ALL' : 'NONE');
          this._syncRowsFromActiveSheets();
          this._render();
        }
      });
    });

    const clearAllFilesBtn = shadow.getElementById('clearAllFilesBtn');
    if (clearAllFilesBtn) {
      clearAllFilesBtn.addEventListener('click', () => {
        this._logAction('QUEUE_CLEARED', 'All Files (' + this._filesQueue.length + ')', 'Removed all workbooks from queue', this._filesQueue.length + ' files', '0 files');
        this._filesQueue = [];
        this._rows = this._rows.filter(r => r._fileOwner === '__MANUAL__');
        if (this._rows.length === 0) this._step = 1;
        this._validateRows();
        this._notify('All files removed from queue.', 'info');
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
        this._step = 1; this._filesQueue = []; this._rows = []; this._currentSessionLogs = []; this._searchQuery = ''; this._render();
      });
    }
    if (downloadLogBtn) downloadLogBtn.addEventListener('click', () => this._exportErrorLog());
    if (autoFixAllBtn) autoFixAllBtn.addEventListener('click', () => this._applyAllAutoFixes());

    if (addRowBtn) {
      addRowBtn.addEventListener('click', () => {
        const newId = 'PLN_' + Math.floor(Math.random() * 9000 + 1000);
        const newRow = { _rowNum: 2, _sourceFile: 'Manual Entry', _fileOwner: '__MANUAL__', _isManual: true, _editedCols: {}, _checkedRules: {}, _errors: {}, _alerts: {} };
        this._requiredColumns.forEach(col => { newRow[col] = col === 'ID' ? newId : ''; });
        this._rows.unshift(newRow);
        this._rows.forEach((r, idx) => { r._rowNum = idx + 2; });
        this._logAction('ROW_ADDED', 'Row ID ' + newId, 'Manual planning row inserted directly in grid', '', newId);
        this._validateRows();
        this._notify('New planning row added.', 'success');
      });
    }

    shadow.querySelectorAll('.remove-row-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const origIdx = parseInt(e.currentTarget.dataset.origIdx, 10);
        if (!isNaN(origIdx) && this._rows[origIdx]) {
          const removed = this._rows[origIdx];
          this._logAction('ROW_DELETED', 'Row #' + removed._rowNum + ' (ID: ' + (removed.ID || 'N/A') + ')', 'Deleted row from ' + removed._sourceFile, String(removed.ID || 'Row'), 'DELETED');
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
          const newId = oldId ? oldId + '_COPY' : 'PLN_COPY';
          const clone = { ...source, ID: newId, _fileOwner: '__MANUAL__', _isManual: true, _editedCols: { ...(source._editedCols || {}) }, _checkedRules: {}, _errors: {}, _alerts: {} };
          this._rows.splice(origIdx + 1, 0, clone);
          this._rows.forEach((r, idx) => { r._rowNum = idx + 2; });
          this._logAction('ROW_CLONED', 'Row #' + source._rowNum, 'Cloned planning row from ID ' + oldId, oldId, newId);
          this._validateRows();
          this._notify('Row duplicated. Verify the new ID.', 'info');
        }
      });
    });

    shadow.querySelectorAll('.cell-autofix-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        this._applyCellAutoFix(parseInt(e.currentTarget.dataset.origIdx, 10), e.currentTarget.dataset.col, e.currentTarget.dataset.ruleId);
      });
    });

    shadow.querySelectorAll('.cell-check-ok-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        this._acknowledgeCheckCell(parseInt(e.currentTarget.dataset.origIdx, 10), e.currentTarget.dataset.col, e.currentTarget.dataset.ruleId);
      });
    });

    const ackNonBlockingChk = shadow.getElementById('ackNonBlockingChk');
    if (ackNonBlockingChk) ackNonBlockingChk.addEventListener('change', (e) => { this._acknowledgeNonBlocking = e.currentTarget.checked; this._render(); });

    const openAddColBtn = shadow.getElementById('openAddColBtn');
    const confirmAddColBtn = shadow.getElementById('confirmAddColBtn');
    const cancelAddColBtn = shadow.getElementById('cancelAddColBtn');
    const newColInput = shadow.getElementById('newColInput');

    if (openAddColBtn) {
      openAddColBtn.addEventListener('click', () => {
        this._showAddColModal = true; this._render();
        const inp = this._shadow.getElementById('newColInput');
        if (inp) inp.focus();
      });
    }
    if (cancelAddColBtn) cancelAddColBtn.addEventListener('click', () => { this._showAddColModal = false; this._render(); });
    if (confirmAddColBtn && newColInput) {
      const saveCol = () => {
        const upperCol = newColInput.value.trim().toUpperCase().replace(/\s+/g, '_');
        if (!upperCol) return;
        if (this._requiredColumns.includes(upperCol)) { this._notify('Column ' + upperCol + ' already exists.', 'error'); return; }
        this._requiredColumns.push(upperCol);
        this._storageSet('sac_required_cols', this._requiredColumns);
        this._auditExportConfig.selectedCols.push(upperCol);
        this._auditExportConfig.numCols = this._auditExportConfig.selectedCols.length;
        this._rows.forEach(r => { if (r[upperCol] === undefined) r[upperCol] = ''; });
        this._logAction('SCHEMA_ADD_COL', 'Column [' + upperCol + ']', 'Added mandatory column to schema', '', upperCol);
        this._showAddColModal = false;
        this._validateRows();
        this._notify('Column ' + upperCol + ' added to model.', 'success');
      };
      confirmAddColBtn.addEventListener('click', saveCol);
      newColInput.addEventListener('keydown', (e) => { if (e.key === 'Enter') saveCol(); });
    }

    shadow.querySelectorAll('.remove-col-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const colToRemove = e.currentTarget.dataset.col;
        if (colToRemove === 'ID') { this._notify('Primary ID column cannot be removed.', 'error'); return; }
        this._requiredColumns = this._requiredColumns.filter(c => c !== colToRemove);
        this._storageSet('sac_required_cols', this._requiredColumns);
        this._syncExportColumnsWithSchema(this._requiredColumns);
        if (!this._requiredColumns.includes(this._bulkCol)) this._bulkCol = this._requiredColumns[0] || 'ID';
        this._logAction('SCHEMA_DEL_COL', 'Column [' + colToRemove + ']', 'Removed column from active planning model', colToRemove, 'DELETED');
        this._validateRows();
        this._notify('Column ' + colToRemove + ' removed.', 'info');
      });
    });

    const toggleInlineRuleBtn = shadow.getElementById('toggleInlineRuleBtn');
    if (toggleInlineRuleBtn) toggleInlineRuleBtn.addEventListener('click', () => { this._showInlineRuleCreator = !this._showInlineRuleCreator; this._render(); });

    const goToRulesTabBtn = shadow.getElementById('goToRulesTabBtn');
    if (goToRulesTabBtn) goToRulesTabBtn.addEventListener('click', () => { this._activeTab = 'rulesTab'; this._render(); });

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
    if (ruleSeveritySel) {
      ruleSeveritySel.addEventListener('change', (e) => {
        this._newRuleForm.severity = e.target.value;
        this._newRuleForm.semaphore = e.target.value === 'BLOCKING' ? 'RED' : (e.target.value === 'PREVENTIVE' ? 'AMBER' : (e.target.value === 'CHECK' ? 'PURPLE' : 'BLUE'));
        this._render();
      });
    }
    if (ruleSemaphoreSel) ruleSemaphoreSel.addEventListener('change', (e) => { this._newRuleForm.semaphore = e.target.value; });
    if (ruleTypeSel) ruleTypeSel.addEventListener('change', (e) => { this._newRuleForm.type = e.target.value; this._render(); });
    if (ruleParamInp) ruleParamInp.addEventListener('input', (e) => { this._newRuleForm.param = e.target.value; });
    if (ruleMinInp) ruleMinInp.addEventListener('input', (e) => { this._newRuleForm.min = e.target.value; });
    if (ruleMaxInp) ruleMaxInp.addEventListener('input', (e) => { this._newRuleForm.max = e.target.value; });
    if (ruleAutoFixInp) ruleAutoFixInp.addEventListener('input', (e) => { this._newRuleForm.autoFix = e.target.value; });
    if (ruleScopeFilesInp) ruleScopeFilesInp.addEventListener('input', (e) => { this._newRuleForm.scopeFiles = e.target.value; });
    if (ruleStartDateInp) ruleStartDateInp.addEventListener('change', (e) => { this._newRuleForm.startDate = e.target.value; });
    if (ruleEndDateInp) ruleEndDateInp.addEventListener('change', (e) => { this._newRuleForm.endDate = e.target.value; });
    if (ruleDescInp) ruleDescInp.addEventListener('input', (e) => { this._newRuleForm.desc = e.target.value; });

    shadow.querySelectorAll('.scope-quick-pill').forEach(pill => {
      pill.addEventListener('click', (e) => {
        const val = e.currentTarget.dataset.scopeVal;
        if (val === 'ALL') this._newRuleForm.scopeFiles = 'ALL';
        else {
          const cur = (this._newRuleForm.scopeFiles || 'ALL').trim();
          if (cur.toUpperCase() === 'ALL' || !cur) this._newRuleForm.scopeFiles = val;
          else {
            const parts = cur.split(',').map(s => s.trim()).filter(Boolean);
            if (!parts.includes(val)) parts.push(val);
            this._newRuleForm.scopeFiles = parts.join(', ');
          }
        }
        this._render();
      });
    });

    shadow.querySelectorAll('.assisted-snippet-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const preset = e.currentTarget.dataset.preset;
        if (preset === 'capex_check') {
          Object.assign(this._newRuleForm, { col: this._requiredColumns.includes('CAPITALIZED') ? 'CAPITALIZED' : this._requiredColumns[0], severity: 'CHECK', semaphore: 'PURPLE', type: 'ASSISTED_CODE', param: "Number(row.H1 || 0) <= 50000 || String(value).toUpperCase() === 'Y'", autoFix: "'Y'", desc: 'CAPEX Check: If H1 > 50,000, CAPITALIZED must be Y' });
        } else if (preset === 'budget_prev') {
          Object.assign(this._newRuleForm, { col: this._requiredColumns.includes('H1') ? 'H1' : this._requiredColumns[0], severity: 'PREVENTIVE', semaphore: 'AMBER', type: 'ASSISTED_CODE', param: 'num >= 0 && num <= 100000', autoFix: 'num < 0 ? Math.abs(num) : 100000', desc: 'Preventive Alert: H1 budget between 0 and 100,000' });
        } else if (preset === 'cc_company_block') {
          Object.assign(this._newRuleForm, { col: this._requiredColumns.includes('COSTCENTER') ? 'COSTCENTER' : this._requiredColumns[0], severity: 'BLOCKING', semaphore: 'RED', type: 'ASSISTED_CODE', param: "row.COMPANY !== '1000' || String(value).startsWith('CC_')", autoFix: "'CC_' + String(value).trim()", desc: 'Blocking: For Company 1000, Cost Center must start with CC_' });
        } else if (preset === 'upper_info') {
          Object.assign(this._newRuleForm, { col: this._requiredColumns.includes('DESCRIPTION') ? 'DESCRIPTION' : this._requiredColumns[0], severity: 'INFO', semaphore: 'BLUE', type: 'ASSISTED_CODE', param: 'String(value) === String(value).toUpperCase()', autoFix: 'String(value).toUpperCase().trim()', desc: 'Info: Recommended uppercase normalized description' });
        }
        this._render();
      });
    });

    if (saveNewRuleBtn) saveNewRuleBtn.addEventListener('click', () => this._createNewValidationRule());
    if (resetDefaultRulesBtn) {
      resetDefaultRulesBtn.addEventListener('click', () => {
        this._validationRules = this._getDefaultRules();
        this._saveRules();
        this._logAction('RULES_RESET', 'Validation Rules Engine', 'Restored default corporate validation rules', 'Custom', 'Default');
        this._validateRows();
        this._notify('Default corporate validation rules restored.', 'info');
      });
    }

    shadow.querySelectorAll('.toggle-rule-chk').forEach(chk => {
      chk.addEventListener('change', (e) => {
        const rule = this._validationRules.find(r => r.id === e.currentTarget.dataset.ruleId);
        if (rule) {
          rule.active = e.currentTarget.checked;
          this._saveRules();
          this._logAction('RULE_TOGGLED', rule.col + ' [' + rule.severity + ']', 'Validation rule ' + (rule.active ? 'enabled' : 'disabled'), !rule.active, rule.active);
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
        if (rule) this._logAction('RULE_DELETED', rule.col + ' [' + rule.severity + ']', 'Deleted validation rule', rule.param || rule.type, 'DELETED');
        this._validateRows();
        this._notify('Validation rule removed.', 'info');
      });
    });

    shadow.querySelectorAll('.copy-rule-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const sourceRule = this._validationRules.find(r => r.id === e.currentTarget.dataset.ruleId);
        if (sourceRule) {
          const clonedRule = { ...sourceRule, id: (function() { try { return 'RULE-' + crypto.randomUUID().split('-')[0].toUpperCase(); } catch(e) { return 'RULE-' + Math.floor(100000 + Math.random() * 900000); } })(), desc: sourceRule.desc + ' (Copy)' };
          this._validationRules.push(clonedRule);
          this._editingScopeRuleId = clonedRule.id;
          this._saveRules();
          this._logAction('RULE_COPIED', clonedRule.col + ' [' + clonedRule.id + ']', 'Cloned rule from ' + sourceRule.id, sourceRule.id, clonedRule.id);
          this._validateRows();
          this._notify('Rule cloned (' + clonedRule.id + '). Edit its target workbook scope inline.', 'success');
        }
      });
    });

    shadow.querySelectorAll('.edit-scope-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const rId = e.currentTarget.dataset.ruleId;
        this._editingScopeRuleId = this._editingScopeRuleId === rId ? null : rId;
        this._render();
      });
    });

    shadow.querySelectorAll('.inline-scope-save-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const rId = e.currentTarget.dataset.ruleId;
        const inp = shadow.getElementById('scopeInp_' + rId);
        const rule = this._validationRules.find(r => r.id === rId);
        if (rule && inp) {
          const oldScope = rule.scopeFiles || 'ALL';
          const newScope = inp.value.trim() || 'ALL';
          rule.scopeFiles = newScope.toUpperCase() === 'ALL' ? 'ALL' : newScope;
          this._editingScopeRuleId = null;
          this._saveRules();
          this._logAction('RULE_SCOPE_UPDATED', rule.col + ' [' + rule.id + ']', 'Updated target workbook scope', oldScope, rule.scopeFiles);
          this._validateRows();
          this._notify('Rule scope updated to: ' + rule.scopeFiles, 'success');
        }
      });
    });

    const searchInput = shadow.getElementById('searchInput');
    if (searchInput) {
      searchInput.addEventListener('input', (e) => {
        const pos = e.target.selectionStart;
        this._searchQuery = e.target.value;
        if (this._searchRaf) cancelAnimationFrame(this._searchRaf);
        this._searchRaf = requestAnimationFrame(() => {
          this._render();
          const restored = this._shadow.getElementById('searchInput');
          if (restored) { restored.focus(); restored.setSelectionRange(pos, pos); }
        });
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
        if (this._filesQueue[idx]) this._filesQueue[idx].comment = e.target.value;
      });
      input.addEventListener('blur', (e) => {
        const idx = parseInt(e.target.dataset.index, 10);
        if (this._filesQueue[idx]) {
          const prevLogged = this._filesQueue[idx]._lastLoggedComment || '';
          const newComment = (this._filesQueue[idx].comment || '').trim();
          this._filesQueue[idx].comment = newComment;
          if (prevLogged !== newComment) {
            this._filesQueue[idx]._lastLoggedComment = newComment;
            this._logAction('FILE_COMMENT', this._filesQueue[idx].file.name, 'Updated business comment', prevLogged || '(Empty)', newComment || '(Empty)');
          }
        }
      });
    });

    shadow.querySelectorAll('.remove-file-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const idx = parseInt(e.currentTarget.dataset.index, 10);
        const removedFile = this._filesQueue[idx];
        if (removedFile) {
          this._logAction('FILE_REMOVED', removedFile.file.name, 'Removed workbook from queue', removedFile.file.name, 'REMOVED');
          this._filesQueue.splice(idx, 1);
          this._syncRowsFromActiveSheets();
        }
        if (this._filesQueue.length === 0 && this._rows.length === 0) this._step = 1;
        this._render();
      });
    });

    shadow.querySelectorAll('.nav-tab').forEach(tab => {
      tab.addEventListener('click', (e) => { this._activeTab = e.currentTarget.dataset.target; this._render(); });
    });

    shadow.querySelectorAll('.chip').forEach(chip => {
      chip.addEventListener('click', (e) => { this._issueFilter = e.currentTarget.dataset.filter; this._render(); });
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
            this._logAction('CELL_EDIT', 'Row #' + row._rowNum + ' [' + colName + ']', 'Inline modification in ' + row._sourceFile, oldVal || '(Blank)', newVal || '(Blank)');
            this._validateRows();
            this._render();
          }
        }
      });
      cell.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); e.target.blur(); } });
    });

    shadow.querySelectorAll('.audit-row-link').forEach(rowEl => {
      rowEl.addEventListener('click', (e) => {
        this._selectedAuditId = e.currentTarget.dataset.loadId;
        const rec = this._history.find(h => h.id === this._selectedAuditId);
        this._syncExportColumnsWithSchema((rec && rec.columns) || this._requiredColumns);
        this._auditDetailSubTab = 'dataTab';
        this._auditSearchQuery = '';
        this._render();
      });
    });

    const backToAuditListBtn = shadow.getElementById('backToAuditListBtn');
    if (backToAuditListBtn) backToAuditListBtn.addEventListener('click', () => { this._selectedAuditId = null; this._render(); });

    const clearHistoryBtn = shadow.getElementById('clearHistoryBtn');
    if (clearHistoryBtn) {
      clearHistoryBtn.addEventListener('click', () => {
        this._history = []; this._selectedAuditId = null; this._storageSet('sac_audit_history_v9', null);
        this._notify('Audit history cleared.', 'info');
      });
    }

    shadow.querySelectorAll('.direct-audit-exp-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const fmt = e.currentTarget.dataset.fmt;
        const targetId = e.currentTarget.dataset.loadId || this._selectedAuditId;
        if (targetId) {
          const rec = this._history.find(h => h.id === targetId);
          if (rec) this._downloadConfiguredAuditReport([rec], rec.id + '_Audit_Export', fmt);
        } else {
          this._downloadConfiguredAuditReport(this._history, 'Global_Audit_Monitor', fmt);
        }
      });
    });

    const toggleAuditExportCfgBtn = shadow.getElementById('toggleAuditExportCfgBtn');
    if (toggleAuditExportCfgBtn) {
      toggleAuditExportCfgBtn.addEventListener('click', () => {
        this._auditExportConfig.open = !this._auditExportConfig.open;
        const rec = this._selectedAuditId ? this._history.find(h => h.id === this._selectedAuditId) : null;
        this._syncExportColumnsWithSchema((rec && rec.columns) || this._requiredColumns);
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
          if (!this._auditExportConfig.selectedCols.includes(col)) this._auditExportConfig.selectedCols.push(col);
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
    if (auditIncCommentsChk) auditIncCommentsChk.addEventListener('change', (e) => { this._auditExportConfig.includeComments = e.currentTarget.checked; });

    const auditIncDetailChk = shadow.getElementById('auditIncDetailChk');
    if (auditIncDetailChk) auditIncDetailChk.addEventListener('change', (e) => { this._auditExportConfig.includeAuditDetail = e.currentTarget.checked; });

    const auditExpFormatSel = shadow.getElementById('auditExpFormatSel');
    if (auditExpFormatSel) auditExpFormatSel.addEventListener('change', (e) => { this._auditExportConfig.format = e.target.value; });

    const runConfiguredAuditDownloadBtn = shadow.getElementById('runConfiguredAuditDownloadBtn');
    if (runConfiguredAuditDownloadBtn) {
      runConfiguredAuditDownloadBtn.addEventListener('click', () => {
        if (this._selectedAuditId) {
          const rec = this._history.find(h => h.id === this._selectedAuditId);
          if (rec) this._downloadConfiguredAuditReport([rec], rec.id + '_Custom_Audit_Report', this._auditExportConfig.format);
        } else {
          this._downloadConfiguredAuditReport(this._history, 'Global_Audit_Monitor_Report', this._auditExportConfig.format);
        }
      });
    }

    shadow.querySelectorAll('.audit-subtab').forEach(subBtn => {
      subBtn.addEventListener('click', (e) => { this._auditDetailSubTab = e.currentTarget.dataset.subtab; this._render(); });
    });

    const auditSearchInput = shadow.getElementById('auditSearchInput');
    if (auditSearchInput) {
      auditSearchInput.addEventListener('input', (e) => {
        const pos = e.target.selectionStart;
        this._auditSearchQuery = e.target.value;
        if (this._auditSearchRaf) cancelAnimationFrame(this._auditSearchRaf);
        this._auditSearchRaf = requestAnimationFrame(() => {
          this._render();
          const restored = this._shadow.getElementById('auditSearchInput');
          if (restored) { restored.focus(); restored.setSelectionRange(pos, pos); }
        });
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
      finishBtn.addEventListener('click', async () => {
        const stats = this._getAlertStats();
        if (!this._targetVersion || !this._targetDate) {
          this._notify('⚠️ Configura la Version y el Period (🎯 SAC Load Target) antes de sincronizar.', 'error');
          return;
        }
        if (stats.blockingRows > 0) { this._notify('Cannot synchronize: ' + stats.blockingRows + ' row(s) have BLOCKING alerts.', 'error'); return; }
        if ((stats.preventiveRows > 0 || stats.checkRows > 0) && !this._acknowledgeNonBlocking) {
          this._notify('Please confirm review of preventive and check alerts before synchronizing.', 'error'); return;
        }
        if (this._rows.length === 0) { this._notify('No active data rows to load.', 'error'); return; }

        const loadId = 'LOAD-' + Math.floor(10000 + Math.random() * 90000);
        const fileSummary = this._filesQueue.length > 0
          ? this._filesQueue.map(f => {
              const activeSheets = (f.sheets || []).filter(s => s.selected).map(s => s.name).join(', ');
              return f.file.name + ' [Sheets: ' + (activeSheets || 'None') + '] (' + (f.comment || 'No comment') + ')';
            }).join('; ')
          : 'Manual Planning Entry [Strategy: ' + this._loadStrategy + ']';

        const filesMetadata = this._filesQueue.map(f => ({
          name: f.file.name,
          strategy: this._loadStrategy,
          sizeKB: Math.max(1, Math.round(f.file.size / 1024)),
          sheets: (f.sheets || []).filter(s => s.selected).length,
          totalSheets: (f.sheets || []).length,
          selectedSheetNames: (f.sheets || []).filter(s => s.selected).map(s => s.name),
          rows: (f.sheets || []).filter(s => s.selected).reduce((acc, s) => acc + s.rows.length, 0),
          comment: f.comment || 'No business comment provided'
        }));

        const commentByFileMap = {};
        this._filesQueue.forEach(f => { commentByFileMap[f.file.name] = f.comment || 'No comment'; });

        const snapshotRows = this._rows.map(r => {
          const copy = {
            _rowNum: r._rowNum,
            _sourceFile: r._sourceFile,
            _fileOwner: r._fileOwner || '',
            _fileComment: commentByFileMap[r._fileOwner] || (r._isManual ? 'Manual Entry' : 'No comment'),
            _loadStrategy: this._loadStrategy,
            _isManual: !!r._isManual,
            _editedCols: { ...(r._editedCols || {}) }
          };
          this._requiredColumns.forEach(c => { copy[c] = r[c] ?? ''; });
          return copy;
        });

        const cleanPayloadArray = this._rows.map(r => {
          const clean = { _loadStrategy: this._loadStrategy, _dspArtifactSync: this._dspArtifactSync, _targetVersion: this._targetVersion, _targetDate: this._targetDate };
          this._requiredColumns.forEach(c => { clean[c] = r[c] ?? ''; });
          return clean;
        });

        const manualEditsCount = this._currentSessionLogs.filter(
          l => ['CELL_EDIT', 'BULK_FILL', 'AUTO_FIX_CODE', 'BULK_AUTO_FIX', 'ROW_ADDED', 'ROW_CLONED', 'ROW_DELETED', 'CHECK_VERIFIED'].includes(l.type)
        ).length;

        const appliedRulesSnapshot = this._validationRules.filter(r => r.active).map(r => ({
          id: r.id, col: r.col, severity: r.severity, semaphore: r.semaphore, type: r.type,
          scopeFiles: r.scopeFiles || 'ALL',
          validity: (r.startDate || r.endDate) ? (r.startDate || 'Any') + ' to ' + (r.endDate || 'Any') : 'Permanent'
        }));

        const auditRecord = {
          id: loadId,
          user: this._currentUser,
          loadStrategy: this._loadStrategy,
          dspArtifact: this._dspArtifactSync,
          semaphoreStatus: stats.preventiveRows > 0 ? 'AMBER' : (stats.checkRows > 0 ? 'PURPLE' : 'GREEN'),
          file: fileSummary,
          filesDetails: filesMetadata,
          appliedRules: appliedRulesSnapshot,
          columns: [...this._requiredColumns],
          rows: this._rows.length,
          editsCount: manualEditsCount,
          timestamp: new Date().toLocaleString(),
          status: this._sacModelId
            ? this._loadStrategy + ' (DIS Pendiente…)'
            : (stats.preventiveRows > 0 || stats.checkRows > 0 ? this._loadStrategy + ' (Approved Alerts)' : this._loadStrategy + ' Synchronized'),
          disStatus: this._sacModelId ? 'PENDING' : null,
          snapshotRows,
          changeLog: [...this._currentSessionLogs],
          payloadJson: JSON.stringify(cleanPayloadArray, null, 2)
        };

        this._history.unshift(auditRecord);
        this._saveHistory();

        this.dispatchEvent(new CustomEvent('propertiesChanged', {
          bubbles: true, composed: true,
          detail: { properties: { validatedDataJson: JSON.stringify(cleanPayloadArray) } }
        }));

        // ── SAC DIS write-back ─────────────────────────────────────────────
        if (this._sacModelId) {
          this._isLoading = true;
          this._loadingText = '📤 Iniciando importación a ' + this._sacModelId + '…';
          this._render();
          let disResult;
          try {
            disResult = await this._callDISApi(this._rows);
            this._isLoading = false;
          } catch (err) {
            this._isLoading = false;
            // Update audit record with real DIS error — do NOT show "Synchronized"
            const idx = this._history.findIndex(h => h.id === loadId);
            if (idx !== -1) {
              this._history[idx].status   = this._loadStrategy + ' ❌ Error DIS';
              this._history[idx].disStatus = 'ERROR';
              this._history[idx].disError  = err.message;
              this._saveHistory();
            }
            this._notify('❌ SAC DIS Error: ' + err.message, 'error');
            this._render();
            return;
          }

          // Update audit with actual DIS result (rows written vs. rejected)
          const jobStatus  = (disResult?.jobStatus || disResult?.status || '').toUpperCase();
          const written    = disResult?.importedRecords ?? disResult?.writtenRecords ?? null;
          const failed     = disResult?.failedRecords   ?? disResult?.rejectedRecords ?? 0;
          const idx = this._history.findIndex(h => h.id === loadId);
          if (idx !== -1) {
            if (failed > 0 || jobStatus === 'COMPLETED_WITH_FAILURES') {
              const writtenTxt = written != null ? written + ' grabadas, ' : '';
              this._history[idx].status    = this._loadStrategy + ' ⚠️ DIS: ' + writtenTxt + failed + ' rechazadas';
              this._history[idx].disStatus = 'PARTIAL';
            } else {
              const writtenTxt = written != null ? ' (' + written + ' filas)' : '';
              this._history[idx].status    = this._loadStrategy + ' Synchronized ✅' + writtenTxt;
              this._history[idx].disStatus = 'OK';
            }
            this._history[idx].disWrittenRows = written;
            this._history[idx].disFailedRows  = failed;
            this._saveHistory();
          }
        }

        // Fire dataReady AFTER DIS completes so SAC table refresh sees the new data
        this.dispatchEvent(new CustomEvent('dataReady', { bubbles: true, composed: true }));

        this._step = 1; this._filesQueue = []; this._rows = []; this._currentSessionLogs = [];
        this._acknowledgeNonBlocking = false; this._activeTab = 'monitorTab'; this._selectedAuditId = loadId;
        this._syncExportColumnsWithSchema(auditRecord.columns);
        this._auditDetailSubTab = 'dataTab';

        // Final notification — only "success" if DIS actually wrote the data
        const finalRec = this._history.find(h => h.id === loadId);
        if (!this._sacModelId || finalRec?.disStatus === 'OK') {
          this._notify('Load ' + loadId + ' (' + this._loadStrategy + ') synchronized successfully with SAP Analytics Cloud!', 'success');
        } else if (finalRec?.disStatus === 'PARTIAL') {
          this._notify('⚠️ Load ' + loadId + ': ' + (finalRec.disFailedRows || '?') + ' filas rechazadas por SAC DIS. Revisa la auditoría.', 'warning');
        }
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
    const scopeFiles = (this._newRuleForm.scopeFiles || 'ALL').trim() || 'ALL';
    const startDate = (this._newRuleForm.startDate || '').trim();
    const endDate = (this._newRuleForm.endDate || '').trim();
    let desc = (this._newRuleForm.desc || '').trim();

    if ((type === 'IN_LIST' || type === 'STARTS_WITH' || type === 'REGEX' || type === 'ASSISTED_CODE') && !param) {
      this._notify('Please specify validation parameter or assisted code formula.', 'error'); return;
    }
    if (!desc) desc = col + ' rule [' + type + '] ' + param;

    const newRule = { id: (function() { try { return 'RULE-' + crypto.randomUUID().split('-')[0].toUpperCase(); } catch(e) { return 'RULE-' + Math.floor(100000 + Math.random() * 900000); } })(), col, severity, semaphore, type, param, min, max, autoFix, scopeFiles, startDate, endDate, active: true, desc };
    this._validationRules.push(newRule);
    this._saveRules();
    this._logAction('RULE_CREATED', col + ' [' + severity + ' / ' + semaphore + ']', desc + ' (Scope: ' + scopeFiles + ')', '', param || type);
    Object.assign(this._newRuleForm, { param: '', min: '', max: '', autoFix: '', desc: '' });
    this._showInlineRuleCreator = false;
    this._validateRows();
    this._notify('Rule created for column ' + col + ' (Scope: ' + scopeFiles + ').', 'success');
  }

  _applyBulkFill(onlyEmpty) {
    const col = this._requiredColumns.includes(this._bulkCol) ? this._bulkCol : (this._requiredColumns[0] || 'ID');
    this._bulkCol = col;
    const val = this._bulkVal.trim();
    if (!col || !val) { this._notify('Enter a value to apply in Bulk Fill.', 'error'); return; }
    let updated = 0;
    this._getFilteredRows().forEach(({ row }) => {
      const current = String(row[col] ?? '').trim();
      if (!onlyEmpty || current === '') {
        if (current !== val) {
          this._recordCellModification(row, col, current || '(Blank)', val);
          row[col] = val;
          updated++;
        }
      }
    });
    if (updated > 0) this._logAction('BULK_FILL', 'Column [' + col + '] (' + updated + ' rows)', onlyEmpty ? 'Bulk filled empty cells' : 'Bulk overwrote visible cells', onlyEmpty ? '(Empty)' : '(Multiple)', val);
    this._validateRows();
    this._notify('Bulk updated ' + updated + ' row(s) in column ' + col + '.', 'success');
  }

  async _handleFilesSelection(fileList) {
    if (typeof XLSX === 'undefined') { this._notify('Excel engine is still loading...', 'error'); return; }
    const newItemsToParse = [];
    const duplicateNames = [];
    for (const newFile of fileList) {
      if (this._filesQueue.some(item => item.file.name.toLowerCase() === newFile.name.toLowerCase())) {
        duplicateNames.push(newFile.name);
      } else {
        const ext = (newFile.name.split('.').pop() || '').toUpperCase();
        const queueItem = { file: newFile, ext, comment: '', _lastLoggedComment: '', sheets: [], sheetCount: 0, rowCount: 0, parsed: false };
        this._filesQueue.push(queueItem);
        newItemsToParse.push(queueItem);
        this._logAction('FILE_UPLOAD', newFile.name, 'Workbook loaded (' + Math.max(1, Math.round(newFile.size / 1024)) + ' KB, Format: ' + ext + ')', '', 'QUEUED');
      }
    }
    if (newItemsToParse.length > 0) await this._parseFilesInQueue(newItemsToParse);
    if (duplicateNames.length > 0) this._notify('Skipped duplicate file(s): ' + duplicateNames.join(', '), 'error');
  }

  async _reparseAllFilesWithNewConfig() {
    if (this._filesQueue.length === 0) return;
    this._logAction('PARSE_CONFIG_CHANGED', 'Format: ' + this._parseConfig.format, 'Re-parsing workbooks with Header Row #' + this._parseConfig.headerRow, '', 'REPARSED');
    await this._parseFilesInQueue(this._filesQueue);
    this._notify('All files re-parsed with new settings.', 'success');
  }

  _readFileAsArrayBuffer(file) {
    return new Promise((resolve, reject) => {
      const r = new FileReader(); r.onload = (e) => resolve(e.target.result); r.onerror = reject; r.readAsArrayBuffer(file);
    });
  }

  _readFileAsText(file) {
    return new Promise((resolve, reject) => {
      const r = new FileReader(); r.onload = (e) => resolve(e.target.result); r.onerror = reject; r.readAsText(file, 'utf-8');
    });
  }

  _parseDelimitedTextToAOA(text, delimiterSetting) {
    const TAB_CH = String.fromCharCode(9);
    const LF_CH = String.fromCharCode(10);
    const CR_CH = String.fromCharCode(13);
    let sep = delimiterSetting === 'TAB' ? TAB_CH : delimiterSetting;
    const normalized = String(text ?? '').split(CR_CH + LF_CH).join(LF_CH).split(CR_CH).join(LF_CH);
    const lines = normalized.split(LF_CH).filter(l => l.trim() !== '');
    if (lines.length === 0) return [];
    if (sep === 'AUTO') {
      const s = lines[0];
      const counts = {};
      counts[TAB_CH] = s.split(TAB_CH).length - 1;
      counts[';'] = s.split(';').length - 1;
      counts[','] = s.split(',').length - 1;
      counts['|'] = s.split('|').length - 1;
      sep = Object.keys(counts).reduce((a, b) => (counts[a] >= counts[b] ? a : b), ',');
      if (counts[sep] === 0) sep = ',';
    }
    return lines.map(line => {
      const res = []; let cur = ''; let inQ = false;
      for (let i = 0; i < line.length; i++) {
        const ch = line[i];
        if (ch === '"') inQ = !inQ;
        else if (ch === sep && !inQ) { res.push(cur.trim()); cur = ''; }
        else cur += ch;
      }
      res.push(cur.trim());
      return res;
    });
  }

  async _parseFilesInQueue(itemsToParse) {
    this._isLoading = true; this._loadingText = 'Parsing sheets & data...'; this._render();
    let detectedHeaders = [...this._headers];
    const headerIdx = Math.max(0, (this._parseConfig.headerRow || 1) - 1);

    for (const item of itemsToParse) {
      try {
        const prevSelectionMap = {};
        (item.sheets || []).forEach(s => { prevSelectionMap[s.name] = s.selected; });
        item.sheets = [];
        let fileRowsCount = 0;
        const isTextFile = ['CSV', 'TXT', 'TSV'].includes(item.ext) || this._parseConfig.format === 'CSV_TXT';

        const buildSheetRows = (aoa, sheetName) => {
          if (aoa.length <= headerIdx) { item.sheets.push({ name: sheetName, selected: false, rows: [] }); return; }
          const sheetHeaders = (aoa[headerIdx] || []).map(h => String(h ?? '').trim().toUpperCase()).filter(Boolean);
          if (detectedHeaders.length === 0 && sheetHeaders.length > 0) detectedHeaders = sheetHeaders;
          const currentHeaders = sheetHeaders.length > 0 ? sheetHeaders : detectedHeaders;
          const rows = aoa.slice(headerIdx + 1).map(row => {
            const obj = { _rowNum: 0, _sourceFile: item.file.name + ' [' + sheetName + ']', _fileOwner: item.file.name, _sheetOwner: sheetName, _isManual: false, _editedCols: {}, _checkedRules: {}, _errors: {}, _alerts: {} };
            currentHeaders.forEach((h, i) => { obj[h] = this._normalizeCellValue(row[i]); });
            this._requiredColumns.forEach(reqCol => { if (obj[reqCol] === undefined) obj[reqCol] = ''; });
            return obj;
          }).filter(r => Object.keys(r).some(k => !k.startsWith('_') && String(r[k]).trim() !== ''));
          item.sheets.push({ name: sheetName, selected: prevSelectionMap[sheetName] !== undefined ? prevSelectionMap[sheetName] : true, rows });
          fileRowsCount += rows.length;
        };

        if (isTextFile) {
          const text = await this._readFileAsText(item.file);
          buildSheetRows(this._parseDelimitedTextToAOA(text, this._parseConfig.delimiter), 'Delimited_Data');
        } else {
          const buffer = await this._readFileAsArrayBuffer(item.file);
          const wb = XLSX.read(new Uint8Array(buffer), { type: 'array' });
          wb.SheetNames.forEach(sheetName => {
            buildSheetRows(XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1 }), sheetName);
          });
        }
        item.sheetCount = item.sheets.length;
        item.rowCount = fileRowsCount;
        item.parsed = true;
      } catch (err) {
        this._notify('Error reading ' + item.file.name, 'error');
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
    row._errors[col] = '[' + alertObj.severity + '] ' + alertObj.msg;
  }

  _validateRows() {
    if (this._rows.length === 0) return;
    const activeRules = this._validationRules.filter(r => r.active);
    const todayIso = new Date().toISOString().slice(0, 10);
    const valueCountsPerCol = {};
    activeRules.filter(r => r.type === 'UNIQUE').forEach(r => {
      valueCountsPerCol[r.col] = {};
      this._rows.forEach(row => {
        const val = String(row[r.col] ?? '').trim();
        if (val) valueCountsPerCol[r.col][val] = (valueCountsPerCol[r.col][val] || 0) + 1;
      });
    });

    this._rows.forEach(row => {
      row._errors = {}; row._alerts = {};
      if (!row._checkedRules) row._checkedRules = {};

      this._requiredColumns.forEach(col => {
        if (row[col] === undefined || String(row[col]).trim() === '') {
          this._pushAlert(row, col, { ruleId: 'SYS-MANDATORY', severity: 'BLOCKING', semaphore: 'RED', msg: 'Mandatory field', autoFix: '' });
        }
      });

      activeRules.forEach(rule => {
        const col = rule.col;
        if (!this._requiredColumns.includes(col)) return;
        const scope = (rule.scopeFiles || 'ALL').trim();
        if (scope.toUpperCase() !== 'ALL') {
          const owner = String(row._fileOwner || '').trim().toLowerCase();
          const targets = scope.split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
          if (!targets.some(t => owner === t || owner.includes(t))) return;
        }
        if (rule.startDate && todayIso < rule.startDate) return;
        if (rule.endDate && todayIso > rule.endDate) return;

        const rawVal = String(row[col] ?? '').trim();
        if (!rawVal && rule.type !== 'ASSISTED_CODE') return;
        const severity = rule.severity || 'BLOCKING';
        const semaphore = rule.semaphore || (severity === 'BLOCKING' ? 'RED' : (severity === 'PREVENTIVE' ? 'AMBER' : (severity === 'CHECK' ? 'PURPLE' : 'BLUE')));
        if ((severity === 'CHECK' || severity === 'PREVENTIVE') && row._checkedRules[col + '__' + rule.id]) return;

        let failed = false;
        let failMsg = rule.desc || 'Validation failed';

        if (rule.type === 'UNIQUE' && rawVal && valueCountsPerCol[col] && valueCountsPerCol[col][rawVal] > 1) {
          failed = true; failMsg = 'Duplicate value not allowed';
        } else if (rule.type === 'IN_LIST') {
          const allowed = rule.param.split(',').map(s => s.trim()).filter(Boolean);
          if (allowed.length > 0 && !allowed.includes(rawVal)) { failed = true; failMsg = 'Allowed: ' + allowed.join(', '); }
        } else if (rule.type === 'NUMERIC_RANGE') {
          const num = Number(rawVal);
          if (isNaN(num)) { failed = true; failMsg = 'Must be numeric'; }
          else {
            if (rule.min !== '' && !isNaN(Number(rule.min)) && num < Number(rule.min)) { failed = true; failMsg = 'Min threshold: ' + rule.min; }
            if (rule.max !== '' && !isNaN(Number(rule.max)) && num > Number(rule.max)) { failed = true; failMsg = 'Max threshold: ' + rule.max; }
          }
        } else if (rule.type === 'STARTS_WITH' && !rawVal.toUpperCase().startsWith(rule.param.toUpperCase())) {
          failed = true; failMsg = 'Must start with ' + rule.param;
        } else if (rule.type === 'REGEX') {
          try { if (!(new RegExp(rule.param)).test(rawVal)) { failed = true; failMsg = 'Invalid pattern (' + rule.param + ')'; } } catch (e) {}
        } else if (rule.type === 'ASSISTED_CODE') {
          try { if (!this._evalAssistedCode(rule.param, row, col, rule.id + '_p')) { failed = true; failMsg = rule.desc || ('Formula failed: ' + rule.param); } }
          catch (e) { failed = true; failMsg = 'Code syntax error: ' + e.message; }
        }
        if (failed) this._pushAlert(row, col, { ruleId: rule.id, severity, semaphore, msg: failMsg, autoFix: rule.autoFix || '' });
      });
    });
  }

  _getRowSeverities(row) {
    const set = new Set();
    Object.values(row._alerts || {}).forEach(arr => (arr || []).forEach(a => set.add(a.severity)));
    return set;
  }

  _getAlertStats() {
    let blockingRows = 0, preventiveRows = 0, checkRows = 0, infoRows = 0, fixableCells = 0;
    this._rows.forEach(r => {
      const sevs = this._getRowSeverities(r);
      if (sevs.has('BLOCKING')) blockingRows++;
      if (sevs.has('PREVENTIVE')) preventiveRows++;
      if (sevs.has('CHECK')) checkRows++;
      if (sevs.has('INFO')) infoRows++;
      Object.values(r._alerts || {}).forEach(arr => { if ((arr || []).some(a => a.autoFix)) fixableCells++; });
    });
    const cleanRows = this._rows.filter(r => this._getRowSeverities(r).size === 0).length;
    return { blockingRows, preventiveRows, checkRows, infoRows, cleanRows, fixableCells };
  }

  _downloadTemplate() {
    if (typeof XLSX === 'undefined') return;
    const sampleRow = this._requiredColumns.map(col => col === 'ID' ? 'PLN_1001' : (col === 'COMPANY' ? '1000' : (col === 'COSTCENTER' ? 'CC_IT_01' : (col === 'H1' ? '25000' : (col === 'CAPITALIZED' ? 'Y' : 'SAMPLE_DESC')))));
    const ws = XLSX.utils.aoa_to_sheet([this._requiredColumns, sampleRow]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Planning_Template');
    XLSX.writeFile(wb, 'SAC_Planning_Template.xlsx');
    this._notify('Official Excel template downloaded.', 'success');
  }

  _exportCleanExcel() {
    if (typeof XLSX === 'undefined') return;
    if (this._rows.length === 0) { this._notify('No data rows to export.', 'info'); return; }
    const header = ['SOURCE_FILE', ...this._requiredColumns];
    const data = this._rows.map(r => [r._sourceFile, ...this._requiredColumns.map(c => r[c] ?? '')]);
    const ws = XLSX.utils.aoa_to_sheet([header, ...data]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Consolidated_Data');
    XLSX.writeFile(wb, 'SAC_Consolidated_Planning.xlsx');
    this._notify('Consolidated data exported to Excel.', 'success');
  }

  _downloadConfiguredAuditReport(recordsList, filePrefix, forcedFormat = null) {
    const TAB_CH = String.fromCharCode(9);
    const LF_CH = String.fromCharCode(10);
    const cfg = this._auditExportConfig;
    const format = forcedFormat || cfg.format || 'XLSX';
    const chosenCols = cfg.selectedCols && cfg.selectedCols.length > 0 ? cfg.selectedCols : this._requiredColumns.slice(0, cfg.numCols || this._requiredColumns.length);

    let effectiveRecords = recordsList && recordsList.length > 0 ? recordsList : [];
    if (effectiveRecords.length === 0 && this._rows.length > 0) {
      effectiveRecords = [{
        id: 'LIVE-DRAFT',
        loadStrategy: this._loadStrategy,
        semaphoreStatus: 'DRAFT',
        timestamp: new Date().toLocaleString(),
        user: this._currentUser,
        file: 'Current Uncommitted Session',
        snapshotRows: this._rows,
        changeLog: this._currentSessionLogs,
        filesDetails: []
      }];
    }

    const baseHeaders = ['LOAD_ID', 'LOAD_STRATEGY', 'SEMAPHORE_STATUS', 'TIMESTAMP', 'SAP_USER', 'ROW_NUM', 'SOURCE_WORKBOOK_SHEET'];
    if (cfg.includeComments) baseHeaders.push('BUSINESS_COMMENTS');
    if (cfg.includeAuditDetail) baseHeaders.push('AUDIT_TRACEABILITY_STATE', 'CELL_MODIFICATION_HISTORY');
    const finalHeaders = [...baseHeaders, ...chosenCols];

    const dataRows = [];
    const allChangeLogs = [];

    effectiveRecords.forEach(rec => {
      (rec.snapshotRows || []).forEach(r => {
        const editedCols = r._editedCols || {};
        const editSummary = Object.keys(editedCols).map(c => c + ': ' + editedCols[c].from + ' -> ' + editedCols[c].to).join(' | ');
        const stateLabel = r._isManual ? 'NEW / CLONED IN WIDGET' : (Object.keys(editedCols).length > 0 ? 'EDITED (' + Object.keys(editedCols).length + ' cols)' : 'ORIGINAL EXCEL');
        const rowArr = [rec.id, rec.loadStrategy || 'FULL', rec.semaphoreStatus || 'GREEN', rec.timestamp, rec.user, r._rowNum, r._sourceFile];
        if (cfg.includeComments) rowArr.push(r._fileComment || rec.file || 'No comment');
        if (cfg.includeAuditDetail) rowArr.push(stateLabel, editSummary || 'No changes');
        chosenCols.forEach(col => rowArr.push(r[col] ?? ''));
        dataRows.push(rowArr);
      });
      if (cfg.includeAuditDetail && rec.changeLog) {
        rec.changeLog.forEach(l => allChangeLogs.push([rec.id, rec.loadStrategy || 'FULL', rec.timestamp, rec.user, l.time, l.type, l.target, l.detail, l.oldVal, l.newVal]));
      }
    });

    if (format === 'CSV' || format === 'TXT' || typeof XLSX === 'undefined') {
      const isTxt = format === 'TXT';
      const sep = isTxt ? TAB_CH : ',';
      const ext = isTxt ? 'txt' : 'csv';
      const fmtCell = (v) => isTxt ? String(v ?? '').replace(/\s+/g, ' ') : '"' + String(v ?? '').replace(/"/g, '""') + '"';

      let content = finalHeaders.map(fmtCell).join(sep) + LF_CH;
      dataRows.forEach(r => { content += r.map(fmtCell).join(sep) + LF_CH; });

      if (cfg.includeAuditDetail && allChangeLogs.length > 0) {
        content += LF_CH + '--- AUDIT ACTIVITY LOG ---' + LF_CH;
        const logHeaders = ['LOAD_ID', 'LOAD_STRATEGY', 'LOAD_DATE', 'SAP_USER', 'EVENT_TIME', 'ACTION_TYPE', 'TARGET', 'DESCRIPTION', 'PREVIOUS_VALUE', 'NEW_VALUE'];
        content += logHeaders.map(fmtCell).join(sep) + LF_CH;
        allChangeLogs.forEach(l => { content += l.map(fmtCell).join(sep) + LF_CH; });
      }

      const blob = new Blob([content], { type: 'text/plain;charset=utf-8;' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filePrefix + '_' + chosenCols.length + 'Cols.' + ext;
      a.click();
      URL.revokeObjectURL(url);
      this._notify('Audit report exported as .' + ext.toUpperCase() + ' (' + chosenCols.length + ' columns).', 'success');
      return;
    }

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([finalHeaders, ...dataRows]), 'Audit_Snapshot_Data');

    if (cfg.includeComments) {
      const filesHeader = ['LOAD_ID', 'LOAD_STRATEGY', 'TIMESTAMP', 'WORKBOOK_NAME', 'SIZE_KB', 'SELECTED_SHEETS', 'COMMITTED_ROWS', 'BUSINESS_COMMENT'];
      const filesRows = [];
      effectiveRecords.forEach(rec => {
        (rec.filesDetails || []).forEach(f => filesRows.push([rec.id, rec.strategy || rec.loadStrategy || 'FULL', rec.timestamp, f.name, f.sizeKB || 1, (f.selectedSheetNames || []).join(', '), f.rows, f.comment || 'No comment']));
      });
      if (filesRows.length > 0) XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([filesHeader, ...filesRows]), 'Workbooks_&_Comments');
    }

    if (cfg.includeAuditDetail && allChangeLogs.length > 0) {
      const logHeader = ['LOAD_ID', 'LOAD_STRATEGY', 'LOAD_DATE', 'SAP_USER', 'EVENT_TIME', 'ACTION_TYPE', 'TARGET', 'DESCRIPTION', 'PREVIOUS_VALUE', 'NEW_VALUE'];
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([logHeader, ...allChangeLogs]), 'Audit_Activity_Log');
    }

    XLSX.writeFile(wb, filePrefix + '_' + chosenCols.length + 'Cols.xlsx');
    this._notify('Audit Report exported as .XLSX (' + chosenCols.length + ' columns).', 'success');
  }

  _exportAuditRecordJson(rec) {
    if (!rec) return;
    const blob = new Blob([JSON.stringify(rec, null, 2)], { type: 'application/json;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a'); a.href = url; a.download = rec.id + '_Audit_Certificate.json'; a.click();
    URL.revokeObjectURL(url);
    this._notify('Audit certificate ' + rec.id + ' downloaded.', 'success');
  }

  _exportErrorLog() {
    const LF_CH = String.fromCharCode(10);
    const csvCell = (v) => '"' + String(v ?? '').replace(/"/g, '""') + '"';
    let csv = 'SourceFile,Row,ID,Column,Alert Severity,Semaphore,Rule ID,Alert Description' + LF_CH;
    let count = 0;
    this._rows.forEach(r => {
      Object.keys(r._alerts || {}).forEach(col => {
        (r._alerts[col] || []).forEach(a => {
          count++;
          csv += [r._sourceFile, r._rowNum, r.ID || '', col, a.severity, a.semaphore, a.ruleId, a.msg].map(csvCell).join(',') + LF_CH;
        });
      });
    });
    if (count === 0) { this._notify('No validation alerts to export — grid looks clean!', 'info'); return; }
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a'); a.href = url; a.download = 'SAC_Validation_Alerts_Log.csv'; a.click();
    URL.revokeObjectURL(url);
    this._notify('Validation alerts exported: ' + count + ' row(s).', 'success');
  }

  _getFilteredRows() {
    const q = this._searchQuery.trim().toLowerCase();
    return this._rows.map((row, origIdx) => ({ row, origIdx })).filter(({ row }) => {
      const sevs = this._getRowSeverities(row);
      if (this._issueFilter === 'correct' && sevs.size > 0) return false;
      if (['BLOCKING', 'PREVENTIVE', 'CHECK', 'INFO'].includes(this._issueFilter) && !sevs.has(this._issueFilter)) return false;
      if (q) {
        return String(row._sourceFile ?? '').toLowerCase().includes(q) || this._requiredColumns.some(col => String(row[col] ?? '').toLowerCase().includes(q));
      }
      return true;
    });
  }

  _renderSemaphoreDot(sem) {
    const colors = { RED: '#BB0000', AMBER: '#E9730C', PURPLE: '#6D28D9', BLUE: '#0070F2', GREEN: '#107E3E' };
    const bg = colors[sem] || '#107E3E';
    return '<span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:' + bg + ';margin-right:4px;box-shadow:0 0 4px ' + bg + ';" title="Semaphore: ' + sem + '"></span>';
  }

  _renderSeverityBadge(sev, sem) {
    const dot = this._renderSemaphoreDot(sem);
    if (sev === 'BLOCKING') return '<span class="sev-badge sev-blocking">' + dot + ' Blocking</span>';
    if (sev === 'PREVENTIVE') return '<span class="sev-badge sev-preventive">' + dot + ' Preventive</span>';
    if (sev === 'CHECK') return '<span class="sev-badge sev-check">' + dot + ' Check</span>';
    return '<span class="sev-badge sev-info">' + dot + ' Info</span>';
  }

  _renderHelpModal() {
    if (!this._showHelpModal) return '';
    return `
      <div class="help-modal-overlay">
        <div class="help-modal-card">
          <div style="display:flex;justify-content:space-between;align-items:center;border-bottom:1px solid #D9E0E6;padding-bottom:12px;margin-bottom:14px;">
            <div style="display:flex;align-items:center;gap:10px;">
              <span style="background:#0070F2;color:#FFF;width:28px;height:28px;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;font-weight:800;">?</span>
              <div>
                <h3 style="margin:0;color:#1D2D3E;font-size:16px;">📘 Interactive Help Guide — Excel Planning Uploader Pro</h3>
                <span style="font-size:11px;color:#556B82;">Full/Delta Load, Sheet Selector, Scoped Rules, Semaphores & Multi-Format Audit Exports</span>
              </div>
            </div>
            <button class="btn" id="closeHelpModalBtn" type="button">✕ Close Help</button>
          </div>
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;max-height:420px;overflow:auto;">
            <div class="help-step-box">
              <div class="help-step-title">🔄 1. Full vs. Delta Load & File Formats</div>
              <p style="margin:4px 0 0;font-size:12px;line-height:1.45;">Select <strong>FULL LOAD</strong> (complete replacement) or <strong>DELTA LOAD</strong> (incremental adjustments). Drag & drop multiple Excel/CSV/TXT files and configure delimiters and European decimals.</p>
            </div>
            <div class="help-step-box">
              <div class="help-step-title">📑 2. Sheet Selector & Business Comments</div>
              <p style="margin:4px 0 0;font-size:12px;line-height:1.45;">In Step 2, check/uncheck which individual sheets (tabs) inside each workbook to consolidate, attach business audit comments, and add/remove model columns.</p>
            </div>
            <div class="help-step-box">
              <div class="help-step-title">🚦 3. Semaphores, File Scope & Temporal Rules</div>
              <p style="margin:4px 0 0;font-size:12px;line-height:1.45;">Assign color semaphores (🟢🟡🔴🟣), clone rules (<strong>📋 Clone</strong>), target specific workbooks (<strong>ALL</strong> or filenames), set validity dates, and run <strong>Assisted Code Auto-Fix</strong> (<strong>✨</strong>).</p>
            </div>
            <div class="help-step-box">
              <div class="help-step-title">🛡️ 4. Audit Monitor & XLSX / CSV / TXT Exports</div>
              <p style="margin:4px 0 0;font-size:12px;line-height:1.45;">In <strong>🛡️ Audit Monitor</strong>, export directly to <strong>.XLSX</strong>, <strong>.CSV</strong>, or <strong>.TXT</strong>, or open <strong>⚙️ Configure Export</strong> to pick exact column counts, comments, and cell modification logs.</p>
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
      <div style="background:#F0F6FF;border:1px solid #0070F2;border-radius:8px;padding:14px 16px;margin-bottom:14px;">
        <div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px;margin-bottom:10px;">
          <div style="font-weight:800;color:#005CB9;font-size:13px;">
            ⚙️ Custom Audit Export Configuration ${isSingleRecord ? '(Load ' + this._esc(this._selectedAuditId) + ')' : '(Global Audit Log)'}
          </div>
          <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;">
            <div>
              <label style="font-size:10px;font-weight:700;color:#556B82;margin-right:4px;">COLUMNS COUNT (1 - ${availableCols.length}):</label>
              <input type="number" id="auditNumColsInp" class="input-std" style="width:56px;font-weight:700;" min="1" max="${availableCols.length}" value="${cfg.numCols}" />
            </div>
            <div>
              <label style="font-size:10px;font-weight:700;color:#556B82;margin-right:4px;">EXPORT FORMAT:</label>
              <select id="auditExpFormatSel" class="input-std">
                <option value="XLSX" ${cfg.format === 'XLSX' ? 'selected' : ''}>📊 Excel Workbook (.xlsx)</option>
                <option value="CSV" ${cfg.format === 'CSV' ? 'selected' : ''}>📄 CSV Comma-Separated (.csv)</option>
                <option value="TXT" ${cfg.format === 'TXT' ? 'selected' : ''}>📝 Text Tab-Delimited (.txt)</option>
              </select>
            </div>
            <button class="btn btn-success" id="runConfiguredAuditDownloadBtn" type="button">
              📥 Download Custom Audit (${cfg.selectedCols.length} cols • .${cfg.format})
            </button>
          </div>
        </div>
        <div style="display:flex;flex-wrap:wrap;gap:16px;align-items:center;padding-top:8px;border-top:1px dashed #B3D4FF;margin-bottom:10px;">
          <label style="display:inline-flex;align-items:center;gap:6px;font-size:12px;font-weight:700;color:#1D2D3E;cursor:pointer;">
            <input type="checkbox" id="auditIncCommentsChk" ${cfg.includeComments ? 'checked' : ''} />
            <span>💬 Include Business Comments & Workbook Metadata</span>
          </label>
          <label style="display:inline-flex;align-items:center;gap:6px;font-size:12px;font-weight:700;color:#1D2D3E;cursor:pointer;">
            <input type="checkbox" id="auditIncDetailChk" ${cfg.includeAuditDetail ? 'checked' : ''} />
            <span>🛠️ Include Full Audit Trail (Row State, Previous Values & Activity Log)</span>
          </label>
        </div>
        <div style="display:flex;align-items:center;flex-wrap:wrap;gap:6px;">
          <span style="font-size:11px;font-weight:700;color:#556B82;margin-right:4px;">Selected Columns (${cfg.selectedCols.length}/${availableCols.length}):</span>
          ${availableCols.map(col => {
            const isChecked = cfg.selectedCols.includes(col);
            return '<label class="sheet-pill-chk ' + (isChecked ? 'sheet-active' : '') + '"><input type="checkbox" class="audit-exp-col-chk" data-col="' + this._esc(col) + '" ' + (isChecked ? 'checked' : '') + ' style="margin:0;" /><span>' + this._esc(col) + '</span></label>';
          }).join('')}
        </div>
      </div>
    `;
  }

  _renderRuleCreatorForm() {
    const t = this._newRuleForm.type, sev = this._newRuleForm.severity, sem = this._newRuleForm.semaphore;
    const loadedFileNames = this._filesQueue.map(f => f.file.name);
    return `
      <div style="background:#F8FAFC;border:1px solid #B3D4FF;border-radius:8px;padding:16px;margin-top:10px;">
        <div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px;margin-bottom:12px;">
          <div style="font-weight:800;color:#005CB9;font-size:13px;">➕ Create Validation Rule, Semaphore, Workbook Scope & Temporal Validity</div>
          <div style="display:flex;gap:6px;flex-wrap:wrap;align-items:center;">
            <span style="font-size:11px;font-weight:700;color:#556B82;">💻 1-Click Templates:</span>
            <button class="icon-btn assisted-snippet-btn" data-preset="cc_company_block" type="button">🛑 Company/CC</button>
            <button class="icon-btn assisted-snippet-btn" data-preset="budget_prev" type="button">⚠️ Budget</button>
            <button class="icon-btn assisted-snippet-btn" data-preset="capex_check" type="button">🔎 CAPEX</button>
            <button class="icon-btn assisted-snippet-btn" data-preset="upper_info" type="button">ℹ️ Uppercase</button>
          </div>
        </div>
        <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px;margin-bottom:10px;">
          <div>
            <label style="display:block;font-size:10px;font-weight:700;color:#556B82;margin-bottom:3px;">🧱 COLUMN</label>
            <select id="ruleColSel" class="input-std" style="width:100%;">${this._requiredColumns.map(c => '<option value="' + this._esc(c) + '" ' + (this._newRuleForm.col === c ? 'selected' : '') + '>' + this._esc(c) + '</option>').join('')}</select>
          </div>
          <div>
            <label style="display:block;font-size:10px;font-weight:700;color:#556B82;margin-bottom:3px;">⚡ ALERT SEVERITY</label>
            <select id="ruleSeveritySel" class="input-std" style="width:100%;font-weight:700;">
              <option value="BLOCKING" ${sev === 'BLOCKING' ? 'selected' : ''}>🛑 Blocking (Blocks Sync)</option>
              <option value="PREVENTIVE" ${sev === 'PREVENTIVE' ? 'selected' : ''}>⚠️ Preventive (Sign-off)</option>
              <option value="CHECK" ${sev === 'CHECK' ? 'selected' : ''}>🔎 Check (Verify Cell)</option>
              <option value="INFO" ${sev === 'INFO' ? 'selected' : ''}>ℹ️ Info (Guidance)</option>
            </select>
          </div>
          <div>
            <label style="display:block;font-size:10px;font-weight:700;color:#556B82;margin-bottom:3px;">🚦 SEMAPHORE COLOR</label>
            <select id="ruleSemaphoreSel" class="input-std" style="width:100%;font-weight:700;">
              <option value="RED" ${sem === 'RED' ? 'selected' : ''}>🔴 Red Signal</option>
              <option value="AMBER" ${sem === 'AMBER' ? 'selected' : ''}>🟡 Amber Signal</option>
              <option value="PURPLE" ${sem === 'PURPLE' ? 'selected' : ''}>🟣 Purple Signal</option>
              <option value="BLUE" ${sem === 'BLUE' ? 'selected' : ''}>🔵 Blue Signal</option>
            </select>
          </div>
          <div>
            <label style="display:block;font-size:10px;font-weight:700;color:#556B82;margin-bottom:3px;">🧠 LOGIC TYPE</label>
            <select id="ruleTypeSel" class="input-std" style="width:100%;">
              <option value="IN_LIST" ${t === 'IN_LIST' ? 'selected' : ''}>Master Data List</option>
              <option value="NUMERIC_RANGE" ${t === 'NUMERIC_RANGE' ? 'selected' : ''}>Numeric Range</option>
              <option value="STARTS_WITH" ${t === 'STARTS_WITH' ? 'selected' : ''}>Starts With Prefix</option>
              <option value="REGEX" ${t === 'REGEX' ? 'selected' : ''}>Regex Pattern</option>
              <option value="UNIQUE" ${t === 'UNIQUE' ? 'selected' : ''}>Unique (No Duplicates)</option>
              <option value="ASSISTED_CODE" ${t === 'ASSISTED_CODE' ? 'selected' : ''}>💻 Assisted Code (JS)</option>
            </select>
          </div>
        </div>
        <div style="display:flex;flex-wrap:wrap;gap:10px;align-items:flex-end;margin-bottom:10px;">
          ${t === 'IN_LIST' || t === 'STARTS_WITH' || t === 'REGEX' || t === 'ASSISTED_CODE' ? `
            <div style="flex:2;min-width:220px;">
              <label style="display:block;font-size:10px;font-weight:700;color:#005CB9;margin-bottom:3px;">PARAMETER / ASSISTED FORMULA (value, num, row.COL)</label>
              <input type="text" id="ruleParamInp" class="input-std" style="width:100%;box-sizing:border-box;" placeholder="e.g. 1000, 2000 or JS expression" value="${this._esc(this._newRuleForm.param)}" />
            </div>
          ` : ''}
          ${t === 'NUMERIC_RANGE' ? `
            <div><label style="display:block;font-size:10px;font-weight:700;color:#556B82;margin-bottom:3px;">MIN VALUE</label><input type="number" id="ruleMinInp" class="input-std" style="width:90px;" placeholder="0" value="${this._esc(this._newRuleForm.min)}" /></div>
            <div><label style="display:block;font-size:10px;font-weight:700;color:#556B82;margin-bottom:3px;">MAX VALUE</label><input type="number" id="ruleMaxInp" class="input-std" style="width:90px;" placeholder="100000" value="${this._esc(this._newRuleForm.max)}" /></div>
          ` : ''}
          <div style="flex:1;min-width:170px;">
            <label style="display:block;font-size:10px;font-weight:700;color:#107E3E;margin-bottom:3px;">✨ AUTO-FIX EXPRESSION</label>
            <input type="text" id="ruleAutoFixInp" class="input-std" style="width:100%;box-sizing:border-box;font-family:monospace;" placeholder="e.g. '1000' or value.trim()" value="${this._esc(this._newRuleForm.autoFix)}" />
          </div>
        </div>
        <div style="display:flex;flex-wrap:wrap;gap:10px;align-items:flex-end;padding-top:8px;border-top:1px dashed #B3D4FF;">
          <div style="flex:1;min-width:200px;">
            <label style="display:block;font-size:10px;font-weight:700;color:#556B82;margin-bottom:3px;">📂 TARGET WORKBOOKS SCOPE ('ALL' OR FILENAMES)</label>
            <input type="text" id="ruleScopeFilesInp" class="input-std" style="width:100%;box-sizing:border-box;" placeholder="ALL or file1.xlsx, file2.xlsx" value="${this._esc(this._newRuleForm.scopeFiles || 'ALL')}" />
            ${loadedFileNames.length > 0 ? '<div style="margin-top:4px;display:flex;gap:4px;flex-wrap:wrap;"><button type="button" class="icon-btn scope-quick-pill" data-scope-val="ALL" style="font-size:10px;padding:1px 6px;">ALL</button>' + loadedFileNames.map(fn => '<button type="button" class="icon-btn scope-quick-pill" data-scope-val="' + this._esc(fn) + '" style="font-size:10px;padding:1px 6px;">+' + this._esc(fn) + '</button>').join('') + '</div>' : ''}
          </div>
          <div><label style="display:block;font-size:10px;font-weight:700;color:#556B82;margin-bottom:3px;">📅 VALID FROM</label><input type="date" id="ruleStartDateInp" class="input-std" value="${this._esc(this._newRuleForm.startDate)}" /></div>
          <div><label style="display:block;font-size:10px;font-weight:700;color:#556B82;margin-bottom:3px;">📅 VALID UNTIL</label><input type="date" id="ruleEndDateInp" class="input-std" value="${this._esc(this._newRuleForm.endDate)}" /></div>
          <div style="flex:1;min-width:180px;"><label style="display:block;font-size:10px;font-weight:700;color:#556B82;margin-bottom:3px;">💬 DESCRIPTION</label><input type="text" id="ruleDescInp" class="input-std" style="width:100%;box-sizing:border-box;" placeholder="Rule description..." value="${this._esc(this._newRuleForm.desc)}" /></div>
          <div><button class="btn" id="saveNewRuleBtn" type="button">💾 Save Rule</button></div>
        </div>
      </div>
    `;
  }

  _renderValidationRulesTab() {
    return `
      <div>
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:14px;flex-wrap:wrap;gap:10px;">
          <div>
            <h3 style="margin:0;color:#1D2D3E;">⚙️ Validation Rules, Semaphores, File Scope & Temporal Validity</h3>
            <p style="margin:4px 0 0;color:#556B82;font-size:12px;">Clone rules across workbooks, assign multi-file scopes, set temporal validity windows, and configure color semaphores.</p>
          </div>
          <button class="btn btn-ghost" id="resetDefaultRulesBtn" type="button">↺ Restore Default Rules</button>
        </div>
        ${this._renderRuleCreatorForm()}
        <div style="margin-top:18px;border:1px solid #D9E0E6;border-radius:8px;overflow:hidden;background:#FFFFFF;">
          <table>
            <thead>
              <tr>
                <th style="width:55px;text-align:center;">Active</th>
                <th style="width:100px;">🧱 Column</th>
                <th style="width:110px;">🚦 Signal</th>
                <th style="width:115px;">⚡ Severity</th>
                <th style="width:180px;">📂 Workbook Scope</th>
                <th style="width:140px;">📅 Temporal Validity</th>
                <th>💻 Logic & Assisted Code</th>
                <th style="width:110px;text-align:center;">Actions</th>
              </tr>
            </thead>
            <tbody>
              ${this._validationRules.map(r => `
                <tr style="${!r.active ? 'opacity:0.55;background:#F8FAFC;' : ''}">
                  <td style="text-align:center;"><label class="fiori-switch"><input type="checkbox" class="toggle-rule-chk" data-rule-id="${this._esc(r.id)}" ${r.active ? 'checked' : ''} /><span class="fiori-slider"></span></label></td>
                  <td><span class="col-pill">${this._esc(r.col)}</span></td>
                  <td>${this._renderSemaphoreDot(r.semaphore || 'RED')} <strong style="font-size:11px;">${this._esc(r.semaphore || 'RED')}</strong></td>
                  <td>${this._renderSeverityBadge(r.severity || 'BLOCKING', r.semaphore || 'RED')}</td>
                  <td>
                    ${this._editingScopeRuleId === r.id ? `
                      <div style="display:flex;gap:4px;"><input type="text" id="scopeInp_${this._esc(r.id)}" class="input-std" style="width:115px;padding:3px 6px;font-size:11px;" value="${this._esc(r.scopeFiles || 'ALL')}" /><button class="btn inline-scope-save-btn" data-rule-id="${this._esc(r.id)}" type="button" style="padding:3px 8px;font-size:10px;">Save</button></div>
                    ` : `
                      <div style="display:flex;align-items:center;gap:6px;"><code style="background:#F0F6FF;color:#005CB9;padding:2px 6px;border-radius:4px;font-size:11px;">${this._esc(r.scopeFiles || 'ALL')}</code><button class="icon-btn edit-scope-btn" data-rule-id="${this._esc(r.id)}" type="button" style="padding:1px 5px;font-size:10px;">✎</button></div>
                    `}
                  </td>
                  <td style="font-size:11px;color:#556B82;">${r.startDate || r.endDate ? '📅 ' + this._esc(r.startDate || 'Any') + ' -> ' + this._esc(r.endDate || 'Any') : '♾️ Permanent'}</td>
                  <td style="font-family:monospace;font-size:11px;color:#1D2D3E;"><strong>${this._esc(r.type)}</strong>: ${this._esc(r.param || r.desc)}${r.autoFix ? '<div style="color:#107E3E;font-size:10px;">✨ Auto-Fix: ' + this._esc(r.autoFix) + '</div>' : ''}</td>
                  <td style="text-align:center;"><div class="row-actions"><button class="icon-btn copy-rule-btn" data-rule-id="${this._esc(r.id)}" type="button">📋 Clone</button><button class="icon-btn icon-btn-del delete-rule-btn" data-rule-id="${this._esc(r.id)}" type="button">✕</button></div></td>
                </tr>
              `).join('')}
            </tbody>
          </table>
        </div>
      </div>
    `;
  }

  _renderFormatConfigBar(showReparseBtn = false) {
    return `
      <div class="toolbar-box" style="background:#F8FAFC;border:1px solid #B3D4FF;">
        <div style="display:flex;align-items:center;gap:14px;flex-wrap:wrap;">
          <span style="font-weight:800;color:#005CB9;font-size:12px;">🛠️ Format & Delimiter Settings:</span>
          <div>
            <label style="font-size:10px;font-weight:700;color:#556B82;margin-right:4px;">📄 FORMAT:</label>
            <select id="cfgFormat" class="input-std">
              <option value="AUTO" ${this._parseConfig.format === 'AUTO' ? 'selected' : ''}>Auto-Detect (.xlsx, .xls, .csv, .txt)</option>
              <option value="XLSX" ${this._parseConfig.format === 'XLSX' ? 'selected' : ''}>Excel Workbook (.xlsx / .xls / .xlsm)</option>
              <option value="CSV_TXT" ${this._parseConfig.format === 'CSV_TXT' ? 'selected' : ''}>Flat File / CSV / Text (.csv / .txt)</option>
            </select>
          </div>
          <div>
            <label style="font-size:10px;font-weight:700;color:#556B82;margin-right:4px;">⇥ TABULADOR / DELIMITER:</label>
            <select id="cfgDelimiter" class="input-std">
              <option value="AUTO" ${this._parseConfig.delimiter === 'AUTO' ? 'selected' : ''}>Auto-Detect</option>
              <option value="TAB" ${this._parseConfig.delimiter === 'TAB' ? 'selected' : ''}>Tab (TAB)</option>
              <option value=";" ${this._parseConfig.delimiter === ';' ? 'selected' : ''}>Semicolon (;)</option>
              <option value="," ${this._parseConfig.delimiter === ',' ? 'selected' : ''}>Comma (,)</option>
              <option value="|" ${this._parseConfig.delimiter === '|' ? 'selected' : ''}>Pipe (|)</option>
            </select>
          </div>
          <div>
            <label style="font-size:10px;font-weight:700;color:#556B82;margin-right:4px;">🔢 DECIMALS:</label>
            <select id="cfgNumFormat" class="input-std">
              <option value="STD" ${this._parseConfig.numberFormat === 'STD' ? 'selected' : ''}>Standard (1,234.56)</option>
              <option value="EU" ${this._parseConfig.numberFormat === 'EU' ? 'selected' : ''}>European (1.234,56 -> 1234.56)</option>
            </select>
          </div>
          <div>
            <label style="font-size:10px;font-weight:700;color:#556B82;margin-right:4px;">📏 HEADER ROW:</label>
            <input type="number" id="cfgHeaderRow" class="input-std" style="width:52px;" min="1" max="20" value="${this._parseConfig.headerRow}" />
          </div>
        </div>
        ${showReparseBtn ? '<button class="btn btn-ghost" id="reparseFilesBtn" type="button" style="font-size:11px;padding:6px 12px;">↻ Apply & Re-Parse Files</button>' : ''}
      </div>
    `;
  }

  _renderAuditExportButtonsBar(loadId = '') {
    return `
      <div style="display:flex;gap:6px;flex-wrap:wrap;align-items:center;">
        <button class="btn btn-ghost direct-audit-exp-btn" data-fmt="XLSX" data-load-id="${this._esc(loadId)}" type="button" style="font-size:11px;padding:6px 10px;">📊 Export .XLSX</button>
        <button class="btn btn-ghost direct-audit-exp-btn" data-fmt="CSV" data-load-id="${this._esc(loadId)}" type="button" style="font-size:11px;padding:6px 10px;">📄 Export .CSV</button>
        <button class="btn btn-ghost direct-audit-exp-btn" data-fmt="TXT" data-load-id="${this._esc(loadId)}" type="button" style="font-size:11px;padding:6px 10px;">📝 Export .TXT</button>
        <button class="btn" id="toggleAuditExportCfgBtn" type="button" style="font-size:11px;padding:6px 12px;">
          ⚙️ ${this._auditExportConfig.open ? 'Close Custom Export' : 'Configure Export'}
        </button>
      </div>
    `;
  }

  _renderAuditMonitor() {
    if (this._selectedAuditId) {
      const rec = this._history.find(h => h.id === this._selectedAuditId);
      if (!rec) { this._selectedAuditId = null; }
      else {
        const cols = rec.columns || this._requiredColumns;
        const snapshotRows = rec.snapshotRows || [];
        const changeLog = rec.changeLog || [];
        const filesDetails = rec.filesDetails || [];
        const appliedRules = rec.appliedRules || [];
        const q = this._auditSearchQuery.trim().toLowerCase();
        const filteredSnapshot = snapshotRows.filter(r => !q || String(r._sourceFile ?? '').toLowerCase().includes(q) || cols.some(c => String(r[c] ?? '').toLowerCase().includes(q)));
        const modifiedRowsCount = snapshotRows.filter(r => r._isManual || Object.keys(r._editedCols || {}).length > 0).length;

        return `
          <div class="audit-detail-view">
            <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:16px;flex-wrap:wrap;gap:10px;">
              <div style="display:flex;align-items:center;gap:10px;">
                <button class="btn btn-ghost" id="backToAuditListBtn" type="button">← Back to Audit Monitor</button>
                <div>
                  <div style="font-size:11px;color:#556B82;font-weight:600;">🛡️ Audit Monitor / Execution Drill-Down</div>
                  <div style="font-size:16px;font-weight:800;color:#1D2D3E;">Load ID: ${this._esc(rec.id)} <span style="background:#E6F9EC;color:#107E3E;padding:3px 10px;border-radius:12px;font-size:11px;margin-left:6px;">${this._renderSemaphoreDot(rec.semaphoreStatus || 'GREEN')} ${this._esc(rec.status)}</span></div>
                </div>
              </div>
              <div style="display:flex;gap:6px;flex-wrap:wrap;">
                ${this._renderAuditExportButtonsBar(rec.id)}
                <button class="btn btn-ghost" id="exportAuditJsonBtn" type="button" style="font-size:11px;padding:6px 10px;">📜 .JSON</button>
              </div>
            </div>

            ${this._renderAuditExportConfigurator(cols, true)}

            <div class="metrics-bar" style="grid-template-columns:repeat(5,1fr);">
              <div class="metric-box"><div class="metric-title">🔄 Load Strategy</div><div class="metric-val" style="font-size:13px;color:#0064D9;">${this._esc(rec.loadStrategy || 'FULL')}</div></div>
              <div class="metric-box"><div class="metric-title">🕒 Timestamp</div><div class="metric-val" style="font-size:13px;">${this._esc(rec.timestamp)}</div></div>
              <div class="metric-box kpi-ok"><div class="metric-title">✅ Committed Rows</div><div class="metric-val" style="color:#107E3E;">${rec.rows}</div></div>
              <div class="metric-box"><div class="metric-title">✎ Rows Modified</div><div class="metric-val" style="color:#0064D9;">${modifiedRowsCount}</div></div>
              <div class="metric-box"><div class="metric-title">🛠️ Tracked Actions</div><div class="metric-val">${changeLog.length} events</div></div>
            </div>

            ${rec.disStatus ? `<div style="margin:8px 0 4px;padding:10px 14px;border-radius:8px;font-size:12px;font-weight:600;border:1px solid ${rec.disStatus === 'OK' ? '#A3E6BF' : rec.disStatus === 'PARTIAL' ? '#FFD580' : '#F5B8B8'};background:${rec.disStatus === 'OK' ? '#E6F9EC' : rec.disStatus === 'PARTIAL' ? '#FFF8E1' : '#FFF0F0'};color:${rec.disStatus === 'OK' ? '#107E3E' : rec.disStatus === 'PARTIAL' ? '#7A5A00' : '#BB0000'};">
              ${rec.disStatus === 'OK' ? '✅ SAC DIS: Datos grabados en el modelo' + (rec.disWrittenRows != null ? ' (' + rec.disWrittenRows + ' filas)' : '') : ''}
              ${rec.disStatus === 'PARTIAL' ? '⚠️ SAC DIS: ' + (rec.disWrittenRows != null ? rec.disWrittenRows + ' grabadas, ' : '') + (rec.disFailedRows || '?') + ' rechazadas por el modelo' : ''}
              ${rec.disStatus === 'ERROR' ? '❌ SAC DIS Error: ' + this._esc(rec.disError || 'Error desconocido') : ''}
              ${rec.disStatus === 'PENDING' ? '⏳ SAC DIS: Pendiente de confirmación' : ''}
            </div>` : ''}

            <div class="subtabs-bar">
              <button class="audit-subtab ${this._auditDetailSubTab === 'dataTab' ? 'active' : ''}" data-subtab="dataTab" type="button">📊 Consolidated Data Snapshot (${snapshotRows.length})</button>
              <button class="audit-subtab ${this._auditDetailSubTab === 'changesTab' ? 'active' : ''}" data-subtab="changesTab" type="button">🛠️ User Edits, Auto-Fixes & Checks (${changeLog.length})</button>
              <button class="audit-subtab ${this._auditDetailSubTab === 'filesTab' ? 'active' : ''}" data-subtab="filesTab" type="button">📁 Workbooks, Rules & SAC Payload (${filesDetails.length})</button>
            </div>

            ${this._auditDetailSubTab === 'dataTab' ? `
              <div class="toolbar-box" style="margin-bottom:10px;">
                <div style="display:flex;align-items:center;gap:8px;"><span style="font-weight:700;font-size:12px;">🔍 Inspect Committed Rows:</span><input type="text" id="auditSearchInput" class="input-std" style="width:240px;" placeholder="Search ID, cost center, sheet..." value="${this._esc(this._auditSearchQuery)}" /></div>
              </div>
              <div class="tbl-wrap">
                <table>
                  <thead><tr><th style="width:32px;">#</th><th style="width:120px;">🏷️ Traceability</th><th>📂 Source Workbook & Sheet</th><th>💬 Business Comment</th>${cols.map(c => '<th>' + this._esc(c) + '</th>').join('')}</tr></thead>
                  <tbody>
                    ${filteredSnapshot.map(r => {
                      const editedCols = r._editedCols || {};
                      const hasEdits = Object.keys(editedCols).length > 0;
                      return `
                        <tr style="${r._isManual ? 'background:#F5F9FF;' : ''}">
                          <td style="color:#556B82;font-weight:600;font-size:11px;">${r._rowNum}</td>
                          <td>${r._isManual ? '<span style="background:#EBF5FF;color:#005CB9;border:1px solid #B3D4FF;padding:2px 8px;border-radius:10px;font-size:10px;font-weight:700;">★ New / Cloned</span>' : (hasEdits ? '<span style="background:#FFF8E6;color:#B45309;border:1px solid #FCD34D;padding:2px 8px;border-radius:10px;font-size:10px;font-weight:700;">✎ Edited (' + Object.keys(editedCols).length + ')</span>' : '<span style="background:#F0F3F6;color:#556B82;padding:2px 8px;border-radius:10px;font-size:10px;font-weight:600;">Original Excel</span>')}</td>
                          <td style="font-size:11px;color:#0064D9;font-weight:700;">${this._esc(r._sourceFile)}</td>
                          <td style="font-size:11px;color:#556B82;">${this._esc(r._fileComment || 'No comment')}</td>
                          ${cols.map(col => {
                            const editInfo = editedCols[col];
                            return '<td style="' + (editInfo ? 'background:#FFFBEB;border-left:3px solid #F59E0B;' : '') + '"><div style="font-weight:' + (editInfo ? '700' : '400') + ';">' + this._esc(r[col] ?? '') + '</div>' + (editInfo ? '<div style="font-size:10px;color:#B45309;">Was: <strike>' + this._esc(editInfo.from) + '</strike></div>' : '') + '</td>';
                          }).join('')}
                        </tr>
                      `;
                    }).join('')}
                  </tbody>
                </table>
              </div>
            ` : ''}

            ${this._auditDetailSubTab === 'changesTab' ? `
              <div class="tbl-wrap">
                <table>
                  <thead><tr><th style="width:90px;">🕒 Time</th><th style="width:140px;">⚡ Action Type</th><th style="width:180px;">🎯 Target</th><th>📝 Audit Description</th><th style="width:140px;">⏪ Previous Value</th><th style="width:140px;">⏩ Committed Value</th></tr></thead>
                  <tbody>
                    ${changeLog.length > 0 ? changeLog.map(log => `
                      <tr>
                        <td style="font-size:11px;color:#556B82;font-weight:600;">${this._esc(log.time)}</td>
                        <td><span style="background:#FFF8E6;color:#B45309;padding:3px 8px;border-radius:6px;font-weight:700;font-size:10px;">${this._esc(log.type)}</span></td>
                        <td style="font-weight:700;">${this._esc(log.target)}</td>
                        <td>${this._esc(log.detail)}</td>
                        <td style="color:#BB0000;font-family:monospace;font-size:11px;">${this._esc(log.oldVal || '—')}</td>
                        <td style="color:#107E3E;font-weight:700;font-family:monospace;font-size:11px;">${this._esc(log.newVal || '—')}</td>
                      </tr>
                    `).join('') : '<tr><td colspan="6" style="text-align:center;padding:28px;color:#556B82;">No manual edits were required during this load.</td></tr>'}
                  </tbody>
                </table>
              </div>
            ` : ''}

            ${this._auditDetailSubTab === 'filesTab' ? `
              <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;">
                <div style="background:#FFFFFF;border:1px solid #D9E0E6;border-radius:8px;padding:16px;">
                  <h4 style="margin:0 0 12px;color:#1D2D3E;">📂 Consolidated Source Workbooks & Scoped Rules</h4>
                  ${filesDetails.map(f => '<div class="file-card-item" style="margin-bottom:8px;"><div><div style="font-weight:700;color:#0064D9;">📄 ' + this._esc(f.name) + ' (' + (f.sizeKB || 1) + ' KB)</div><div style="font-size:11px;color:#556B82;">Strategy: <strong>' + this._esc(f.strategy || rec.loadStrategy) + '</strong> • ' + f.sheets + ' sheet(s) • ' + f.rows + ' rows</div><div style="font-size:12px;margin-top:6px;background:#F8FAFC;padding:6px 10px;border-radius:6px;border-left:3px solid #0070F2;">💬 <strong>Comment:</strong> ' + this._esc(f.comment) + '</div></div></div>').join('')}
                  ${appliedRules.length > 0 ? '<div style="margin-top:12px;padding-top:12px;border-top:1px solid #D9E0E6;"><div style="font-weight:700;font-size:12px;margin-bottom:6px;">🚦 Active Validation Rules Executed (' + appliedRules.length + '):</div><div style="display:flex;flex-wrap:wrap;gap:6px;">' + appliedRules.map(ar => '<span style="background:#F8FAFC;border:1px solid #D9E0E6;padding:3px 8px;border-radius:6px;font-size:10px;font-weight:600;">' + this._renderSemaphoreDot(ar.semaphore) + ' <strong>' + this._esc(ar.col) + '</strong> [' + this._esc(ar.scopeFiles) + ']</span>').join('') + '</div></div>' : ''}
                </div>
                <div style="background:#FFFFFF;border:1px solid #D9E0E6;border-radius:8px;padding:16px;">
                  <h4 style="margin:0 0 12px;color:#1D2D3E;">💻 Synchronized SAC JSON Payload</h4>
                  <pre style="background:#132337;color:#74D8FF;padding:12px;border-radius:6px;font-size:11px;max-height:220px;overflow:auto;margin:0;">${this._esc(rec.payloadJson || '[]')}</pre>
                </div>
              </div>
            ` : ''}
          </div>
        `;
      }
    }

    return `
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:14px;flex-wrap:wrap;gap:10px;">
        <div>
          <h3 style="margin:0;color:#1D2D3E;">🛡️ Audit Monitor & Multi-Format Export Center</h3>
          <p style="margin:4px 0 0;color:#556B82;font-size:12px;">Click on any execution row to inspect committed rows and modifications, or export audit reports in <strong>.XLSX</strong>, <strong>.CSV</strong>, or <strong>.TXT</strong>.</p>
        </div>
        <div style="display:flex;gap:8px;flex-wrap:wrap;align-items:center;">
          ${this._renderAuditExportButtonsBar('')}
          ${this._history.length > 0 ? '<button class="btn btn-danger" id="clearHistoryBtn" type="button" style="font-size:11px;padding:6px 12px;">🗑️ Clear History</button>' : ''}
        </div>
      </div>

      ${this._renderAuditExportConfigurator(this._requiredColumns, false)}

      <div style="border:1px solid #D9E0E6;border-radius:8px;overflow:hidden;background:#FFFFFF;">
        <table>
          <thead>
            <tr>
              <th>🆔 Load ID</th>
              <th>🔄 Strategy</th>
              <th>👤 SAP User</th>
              <th>📂 Workbooks, Selected Sheets & Comments</th>
              <th>📊 Rows</th>
              <th>✎ User Edits</th>
              <th>🕒 Timestamp</th>
              <th>🚦 Semaphore & Status</th>
              <th style="text-align:right;">📥 Quick Export & Inspect</th>
            </tr>
          </thead>
          <tbody>
            ${this._history.length > 0 ? this._history.map(h => `
              <tr class="audit-row-link" data-load-id="${this._esc(h.id)}" style="cursor:pointer;">
                <td><strong style="color:#0064D9;text-decoration:underline;">${this._esc(h.id)}</strong></td>
                <td><span style="background:#EBF5FF;color:#005CB9;padding:2px 8px;border-radius:10px;font-weight:700;font-size:11px;">${this._esc(h.loadStrategy || 'FULL')}</span></td>
                <td>${this._esc(h.user)}</td>
                <td>${this._esc(h.file)}</td>
                <td><strong>${h.rows}</strong> rows</td>
                <td>${(h.editsCount || 0) > 0 ? '<span style="background:#FFF8E6;color:#B45309;border:1px solid #FCD34D;padding:2px 8px;border-radius:10px;font-weight:700;font-size:11px;">✎ ' + h.editsCount + ' action(s)</span>' : '<span style="color:#556B82;font-size:11px;">Direct (0 edits)</span>'}</td>
                <td>${this._esc(h.timestamp)}</td>
                <td>
                  ${h.disStatus === 'ERROR'
                    ? `<span style="background:#FFF0F0;color:#BB0000;padding:4px 10px;border-radius:12px;font-weight:700;font-size:11px;display:inline-flex;align-items:center;">❌ ${this._esc(h.status)}</span>
                       <div style="font-size:10px;color:#BB0000;margin-top:3px;max-width:260px;white-space:normal;line-height:1.3;">${this._esc(h.disError || 'Error DIS desconocido')}</div>`
                    : h.disStatus === 'PARTIAL'
                    ? `<span style="background:#FFF8E1;color:#7A5A00;padding:4px 10px;border-radius:12px;font-weight:700;font-size:11px;display:inline-flex;align-items:center;">⚠️ ${this._esc(h.status)}</span>`
                    : h.disStatus === 'PENDING'
                    ? `<span style="background:#F0F4FF;color:#0052CC;padding:4px 10px;border-radius:12px;font-weight:700;font-size:11px;display:inline-flex;align-items:center;">⏳ ${this._esc(h.status)}</span>`
                    : `<span style="background:#E6F9EC;color:#107E3E;padding:4px 10px;border-radius:12px;font-weight:700;font-size:11px;display:inline-flex;align-items:center;">${this._renderSemaphoreDot(h.semaphoreStatus || 'GREEN')} ${this._esc(h.status)}</span>`
                  }
                </td>
                <td style="text-align:right;white-space:nowrap;">
                  <button class="icon-btn direct-audit-exp-btn" data-fmt="XLSX" data-load-id="${this._esc(h.id)}" type="button" title="Export Excel .XLSX">📊 .XLSX</button>
                  <button class="icon-btn direct-audit-exp-btn" data-fmt="CSV" data-load-id="${this._esc(h.id)}" type="button" title="Export .CSV">📄 .CSV</button>
                  <button class="icon-btn direct-audit-exp-btn" data-fmt="TXT" data-load-id="${this._esc(h.id)}" type="button" title="Export .TXT">📝 .TXT</button>
                  <button class="btn btn-ghost" type="button" style="padding:4px 10px;font-size:11px;margin-left:4px;">🔍 Details →</button>
                </td>
              </tr>
            `).join('') : `
              <tr><td colspan="9" style="text-align:center;padding:36px;color:#556B82;">No planning loads registered yet. Use the export buttons above to download a template/draft audit report or complete an upload in the Wizard.</td></tr>
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

    let globalSemaphore = 'GREEN', globalSemaphoreText = 'GREEN (Ready)';
    if (stats.blockingRows > 0) { globalSemaphore = 'RED'; globalSemaphoreText = 'RED (Blocking)'; }
    else if (stats.preventiveRows > 0) { globalSemaphore = 'AMBER'; globalSemaphoreText = 'AMBER (Sign-Off)'; }
    else if (stats.checkRows > 0) { globalSemaphore = 'PURPLE'; globalSemaphoreText = 'PURPLE (Check)'; }

    this._shadow.innerHTML = `
      <style>
        :host { display:block; font-family:"72","72full",-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif; font-size:13px; color:#1D2D3E; background:#F5F6F7; padding:16px; box-sizing:border-box; height:100%; overflow:auto; }
        .main-card { background:#FFFFFF; border:1px solid #D9E0E6; border-radius:12px; box-shadow:0 8px 24px rgba(29,45,62,0.07); overflow:hidden; position:relative; }
        .busy-overlay { position:absolute; inset:0; background:rgba(255,255,255,0.9); backdrop-filter:blur(2px); display:flex; flex-direction:column; align-items:center; justify-content:center; z-index:100; }
        .spinner { width:40px; height:40px; border:4px solid #E5E9EC; border-top-color:#0070F2; border-radius:50%; animation:spin 0.75s linear infinite; margin-bottom:12px; }
        @keyframes spin { to { transform:rotate(360deg); } }
        .help-modal-overlay { position:fixed; inset:0; background:rgba(19,35,55,0.62); backdrop-filter:blur(2px); display:flex; align-items:center; justify-content:center; z-index:250; padding:20px; }
        .help-modal-card { background:#FFFFFF; border:1px solid #D9E0E6; border-top:4px solid #0070F2; border-radius:12px; max-width:820px; width:100%; padding:20px 24px; box-shadow:0 16px 40px rgba(0,0,0,0.22); }
        .help-step-box { background:#F8FAFC; border:1px solid #D9E0E6; border-left:4px solid #0070F2; border-radius:8px; padding:12px 14px; }
        .help-step-title { font-weight:800; color:#005CB9; font-size:13px; }
        .sap-header { display:flex; align-items:center; justify-content:space-between; background:linear-gradient(90deg,#132337 0%,#1F3652 100%); color:#FFFFFF; padding:12px 24px; border-bottom:3px solid #0070F2; }
        .sap-brand { display:flex; align-items:center; gap:14px; }
        .sap-logo-img { height:28px; width:auto; display:block; filter:drop-shadow(0 2px 4px rgba(0,0,0,0.25)); }
        .sap-title-wrap { display:flex; flex-direction:column; }
        .sap-title { font-size:15px; font-weight:700; color:#FFFFFF; }
        .sap-subtitle { font-size:11px; color:#A9B8C8; font-weight:500; }
        .sap-badge-env { background:rgba(0,184,241,0.16); border:1px solid rgba(0,184,241,0.45); color:#74D8FF; padding:4px 10px; border-radius:20px; font-size:11px; font-weight:600; }
        .help-icon-btn { background:rgba(255,255,255,0.14); border:1px solid rgba(255,255,255,0.35); color:#FFFFFF; padding:5px 12px; border-radius:20px; font-size:12px; font-weight:700; cursor:pointer; display:inline-flex; align-items:center; gap:6px; }
        .help-icon-btn:hover { background:#0070F2; border-color:#74D8FF; }
        .fiori-switch-wrap { display:inline-flex; align-items:center; gap:8px; cursor:pointer; user-select:none; }
        .fiori-switch { position:relative; display:inline-block; width:38px; height:20px; }
        .fiori-switch input { opacity:0; width:0; height:0; }
        .fiori-slider { position:absolute; cursor:pointer; inset:0; background-color:#788FA6; transition:0.2s; border-radius:20px; }
        .fiori-slider:before { position:absolute; content:""; height:14px; width:14px; left:3px; bottom:3px; background-color:#FFFFFF; transition:0.2s; border-radius:50%; box-shadow:0 1px 3px rgba(0,0,0,0.25); }
        .fiori-switch input:checked + .fiori-slider { background-color:#0070F2; }
        .fiori-switch input:checked + .fiori-slider:before { transform:translateX(18px); }
        .toast-banner { display:flex; align-items:center; justify-content:space-between; padding:10px 24px; font-weight:600; font-size:12px; color:#FFFFFF; }
        .toast-info { background:#0070F2; } .toast-success { background:#107E3E; } .toast-error { background:#BB0000; }
        .nav-tabs { display:flex; background:#FFFFFF; border-bottom:1px solid #D9E0E6; padding:0 24px; gap:8px; }
        .nav-tab { padding:13px 18px; cursor:pointer; font-weight:700; color:#556B82; border-bottom:3px solid transparent; font-size:13px; transition:all 0.2s; display:inline-flex; align-items:center; gap:7px; }
        .nav-tab:hover { color:#0070F2; }
        .nav-tab.active { color:#0070F2; border-bottom-color:#0070F2; background:linear-gradient(180deg,transparent 0%,rgba(0,112,242,0.05) 100%); }
        .tab-content { padding:24px; background:#F5F6F7; }
        .subtabs-bar { display:flex; gap:8px; margin-bottom:14px; border-bottom:1px solid #D9E0E6; padding-bottom:10px; }
        .audit-subtab { padding:8px 14px; border-radius:6px; border:1px solid #D9E0E6; background:#FFFFFF; color:#556B82; font-weight:700; font-size:12px; cursor:pointer; }
        .audit-subtab.active { background:#0070F2; color:#FFFFFF; border-color:#0070F2; }
        .wizard-steps { display:grid; grid-template-columns:repeat(3,1fr); gap:12px; margin-bottom:20px; }
        .step-item { display:flex; align-items:center; gap:12px; background:#FFFFFF; border:1px solid #D9E0E6; border-radius:10px; padding:12px 16px; cursor:pointer; transition:all 0.2s ease; }
        .step-item:hover { border-color:#0070F2; transform:translateY(-1px); }
        .step-item.active { border-color:#0070F2; background:#F0F6FF; }
        .step-item.completed { border-color:#107E3E; }
        .step-num { width:26px; height:26px; border-radius:50%; background:#E5E9EC; color:#556B82; display:flex; align-items:center; justify-content:center; font-size:12px; font-weight:700; }
        .step-item.active .step-num { background:#0070F2; color:#FFFFFF; }
        .step-item.completed .step-num { background:#107E3E; color:#FFFFFF; }
        .step-label-title { font-weight:700; font-size:12px; color:#1D2D3E; }
        .step-label-sub { font-size:11px; color:#556B82; }
        .strategy-bar-card { background:#FFFFFF; border:1px solid #D9E0E6; border-radius:10px; padding:12px 18px; margin-bottom:14px; display:flex; align-items:center; gap:16px; flex-wrap:wrap; }
        .strategy-toggle-group { display:inline-flex; background:#EBF0F5; padding:4px; border-radius:10px; gap:6px; }
        .strategy-btn { border:1px solid transparent; background:transparent; border-radius:8px; padding:8px 16px; text-align:left; cursor:pointer; transition:all 0.18s ease; display:flex; align-items:center; gap:10px; }
        .strategy-btn:hover { background:rgba(255,255,255,0.6); }
        .strategy-btn.active { background:#0070F2; color:#FFFFFF; box-shadow:0 2px 8px rgba(0,112,242,0.28); }
        .strategy-btn-title { font-size:12px; font-weight:800; display:block; color:#1D2D3E; }
        .strategy-btn.active .strategy-btn-title { color:#FFFFFF; }
        .strategy-btn-sub { font-size:10px; color:#556B82; display:block; }
        .strategy-btn.active .strategy-btn-sub { color:#D9EFFF; }
        .btn { padding:8px 16px; font-size:12px; font-weight:600; border-radius:6px; cursor:pointer; border:1px solid #0070F2; background:#0070F2; color:#FFFFFF; transition:all 0.15s ease; display:inline-flex; align-items:center; gap:6px; }
        .btn:hover { background:#005CB9; }
        .btn-ghost { background:#FFFFFF; color:#0064D9; border-color:#BCC8D4; }
        .btn-ghost:hover { background:#F0F6FF; border-color:#0070F2; }
        .btn-danger { background:#FFFFFF; color:#BB0000; border-color:#F5C6C6; }
        .btn-danger:hover { background:#FEEAEA; border-color:#BB0000; }
        .btn-success { background:#107E3E; border-color:#107E3E; color:#FFFFFF; }
        .btn-success:hover { background:#0C6230; }
        .drop-zone { border:2px dashed #0070F2; border-radius:12px; padding:36px 24px; text-align:center; background:#FFFFFF; display:flex; flex-direction:column; align-items:center; gap:12px; }
        .drop-zone.drag-active { background:#F0F6FF; border-color:#005CB9; }
        .upload-icon-circle { width:54px; height:54px; border-radius:50%; background:#EBF5FF; color:#0070F2; display:flex; align-items:center; justify-content:center; font-size:24px; }
        .file-label-btn { padding:9px 20px; background:#0070F2; color:#FFFFFF; font-weight:700; font-size:13px; border-radius:6px; cursor:pointer; display:inline-block; }
        .metrics-bar { display:grid; grid-template-columns:repeat(6,1fr); gap:10px; margin-bottom:14px; }
        .metric-box { background:#FFFFFF; border:1px solid #D9E0E6; border-left:4px solid #0070F2; padding:10px 12px; border-radius:8px; }
        .metric-box.kpi-ok { border-left-color:#107E3E; } .metric-box.kpi-err { border-left-color:#BB0000; } .metric-box.kpi-prev { border-left-color:#E9730C; } .metric-box.kpi-chk { border-left-color:#6D28D9; }
        .metric-title { font-size:10px; color:#556B82; margin-bottom:4px; font-weight:700; text-transform:uppercase; }
        .metric-val { font-size:17px; font-weight:800; color:#1D2D3E; }
        .dist-bar { display:flex; height:8px; border-radius:4px; overflow:hidden; background:#FEEAEA; margin-bottom:16px; border:1px solid #D9E0E6; }
        .dist-ok { background:linear-gradient(90deg,#107E3E,#16A34A); height:100%; transition:width 0.35s ease; }
        .toolbar-box { background:#FFFFFF; border:1px solid #D9E0E6; border-radius:8px; padding:10px 14px; margin-bottom:12px; display:flex; flex-wrap:wrap; gap:12px; align-items:center; justify-content:space-between; }
        .input-std { border:1px solid #BCC8D4; border-radius:6px; padding:6px 10px; font-size:12px; color:#1D2D3E; outline:none; }
        .input-std:focus { border-color:#0070F2; box-shadow:0 0 0 2px rgba(0,112,242,0.15); }
        .chips-row { display:flex; flex-wrap:wrap; gap:10px; margin-bottom:12px; align-items:center; justify-content:space-between; }
        .chips-left { display:flex; gap:6px; align-items:center; flex-wrap:wrap; }
        .chip { padding:5px 12px; border-radius:16px; font-size:11px; font-weight:700; cursor:pointer; border:1px solid #D9E0E6; background:#FFFFFF; color:#32363A; }
        .chip-all.active-chip { background:#1D2D3E; color:#FFFFFF; border-color:#1D2D3E; }
        .chip-ok.active-chip { background:#107E3E; color:#FFFFFF; border-color:#107E3E; }
        .chip-err.active-chip { background:#BB0000; color:#FFFFFF; border-color:#BB0000; }
        .chip-prev.active-chip { background:#E9730C; color:#FFFFFF; border-color:#E9730C; }
        .chip-chk.active-chip { background:#6D28D9; color:#FFFFFF; border-color:#6D28D9; }
        .chip-info.active-chip { background:#0070F2; color:#FFFFFF; border-color:#0070F2; }
        .sev-badge { display:inline-flex; align-items:center; gap:4px; padding:3px 8px; border-radius:12px; font-size:10px; font-weight:800; }
        .sev-blocking { background:#FEEAEA; color:#BB0000; border:1px solid #F5C6C6; }
        .sev-preventive { background:#FFF4E5; color:#B45309; border:1px solid #FCD34D; }
        .sev-check { background:#F5F3FF; color:#6D28D9; border:1px solid #DDD6FE; }
        .sev-info { background:#EBF5FF; color:#005CB9; border:1px solid #B3D4FF; }
        td.cell-sev-BLOCKING { background:#FFF2F2; border-left:3px solid #BB0000; }
        td.cell-sev-PREVENTIVE { background:#FFFBEB; border-left:3px solid #E9730C; }
        td.cell-sev-CHECK { background:#FAF5FF; border-left:3px solid #6D28D9; }
        td.cell-sev-INFO { background:#F0F7FF; border-left:3px solid #0070F2; }
        .alert-line { display:flex; align-items:center; justify-content:space-between; gap:6px; margin-top:3px; padding:2px 5px; border-radius:4px; font-size:10px; font-weight:700; }
        .alert-line-BLOCKING { color:#BB0000; background:rgba(187,0,0,0.06); }
        .alert-line-PREVENTIVE { color:#B45309; background:rgba(233,115,12,0.08); }
        .alert-line-CHECK { color:#6D28D9; background:rgba(109,40,217,0.08); }
        .alert-line-INFO { color:#005CB9; background:rgba(0,112,242,0.08); }
        .mini-fix-btn { border:1px solid #107E3E; background:#E6F9EC; color:#107E3E; font-size:9px; font-weight:800; padding:1px 5px; border-radius:4px; cursor:pointer; }
        .mini-fix-btn:hover { background:#107E3E; color:#FFFFFF; }
        .mini-chk-btn { border:1px solid #6D28D9; background:#F5F3FF; color:#6D28D9; font-size:9px; font-weight:800; padding:1px 5px; border-radius:4px; cursor:pointer; }
        .mini-chk-btn:hover { background:#6D28D9; color:#FFFFFF; }
        .tbl-wrap { max-height:360px; overflow:auto; border:1px solid #D9E0E6; border-radius:8px; background:#FFFFFF; }
        table { width:100%; border-collapse:collapse; font-size:12px; text-align:left; }
        th { background:#F0F3F6; padding:10px 12px; border-bottom:2px solid #D9E0E6; position:sticky; top:0; font-weight:700; color:#1D2D3E; white-space:nowrap; z-index:5; }
        tbody tr:hover { background:#F8FAFC; }
        td { padding:8px 10px; border-bottom:1px solid #EDF0F2; vertical-align:top; }
        .editable-cell { outline:none; border:1px solid transparent; border-bottom:1px dashed #0070F2; display:inline-block; min-width:65px; padding:2px 5px; border-radius:4px; }
        .editable-cell:focus { background:#FFFFFF; border:1px solid #0070F2; }
        .th-content { display:flex; align-items:center; justify-content:space-between; gap:6px; }
        .remove-col-btn { background:none; border:none; color:#BB0000; font-weight:700; cursor:pointer; font-size:11px; }
        .row-actions { display:flex; gap:4px; justify-content:center; }
        .icon-btn { border:1px solid #D9E0E6; background:#FFFFFF; border-radius:4px; cursor:pointer; font-size:11px; padding:3px 7px; font-weight:700; }
        .icon-btn:hover { border-color:#0070F2; color:#0070F2; background:#F0F6FF; }
        .icon-btn-del { border-color:#F5C6C6; color:#BB0000; background:#FEEAEA; }
        .wizard-footer { display:flex; justify-content:space-between; align-items:center; margin-top:20px; border-top:1px solid #D9E0E6; padding-top:16px; flex-wrap:wrap; gap:10px; }
        .file-card-item { background:#FFFFFF; border:1px solid #D9E0E6; border-left:4px solid #0070F2; border-radius:8px; padding:14px 16px; margin-bottom:12px; }
        .sheet-pill-chk { display:inline-flex; align-items:center; gap:6px; padding:4px 10px; border-radius:6px; font-size:11px; font-weight:600; cursor:pointer; border:1px solid #BCC8D4; background:#F8FAFC; }
        .sheet-pill-chk.sheet-active { background:#EBF5FF; border-color:#0070F2; color:#005CB9; }
        .col-pill { display:inline-flex; align-items:center; gap:6px; background:#EBF5FF; color:#005CB9; padding:5px 12px; border-radius:6px; font-weight:700; font-size:11px; border:1px solid #B3D4FF; }

        /* ── Visible scrollbars ─────────────────────────────────────────── */
        :host { scrollbar-width: thin; scrollbar-color: #788FA6 #E5E9EC; }
        :host::-webkit-scrollbar { width: 8px; height: 8px; }
        :host::-webkit-scrollbar-track { background: #E5E9EC; border-radius: 4px; }
        :host::-webkit-scrollbar-thumb { background: #788FA6; border-radius: 4px; }
        :host::-webkit-scrollbar-thumb:hover { background: #556B82; }
        .tbl-wrap { scrollbar-width: thin; scrollbar-color: #0070F2 #D9E0E6; }
        .tbl-wrap::-webkit-scrollbar { width: 7px; height: 7px; }
        .tbl-wrap::-webkit-scrollbar-track { background: #EDF0F2; border-radius: 4px; }
        .tbl-wrap::-webkit-scrollbar-thumb { background: #0070F2; border-radius: 4px; box-shadow: inset 0 0 0 1px rgba(0,112,242,.3); }
        .tbl-wrap::-webkit-scrollbar-thumb:hover { background: #005CB9; }
        .tbl-wrap::-webkit-scrollbar-corner { background: #EDF0F2; }
        .help-modal-card div { scrollbar-width: thin; scrollbar-color: #788FA6 #E5E9EC; }
        .help-modal-card div::-webkit-scrollbar { width: 6px; }
        .help-modal-card div::-webkit-scrollbar-track { background: #F0F3F6; border-radius: 3px; }
        .help-modal-card div::-webkit-scrollbar-thumb { background: #788FA6; border-radius: 3px; }
        pre::-webkit-scrollbar { width: 6px; height: 6px; }
        pre::-webkit-scrollbar-track { background: #0D1B2A; border-radius: 3px; }
        pre::-webkit-scrollbar-thumb { background: #0070F2; border-radius: 3px; }
        pre::-webkit-scrollbar-thumb:hover { background: #74D8FF; }
      </style>

      <div class="main-card">
        ${this._isLoading ? '<div class="busy-overlay"><div class="spinner"></div><div style="font-weight:700;color:#1D2D3E;">' + this._esc(this._loadingText) + '</div></div>' : ''}
        ${this._renderHelpModal()}

        <div class="sap-header">
          <div class="sap-brand">
            <img class="sap-logo-img" src="${this._sapLogoUri}" alt="SAP Logo" />
            <div class="sap-title-wrap">
              <div class="sap-title">Excel Planning Uploader Pro</div>
              <div class="sap-subtitle">Multi-Sheet Consolidation · Transform &amp; Convert · Alert Semaphores · v2.5.0${
                this._sacModelState === 'connecting' ? ' &nbsp;|&nbsp; <span style="color:#A9B8C8;font-weight:600;">⏳ Conectando…</span>' :
                this._sacModelState === 'ok'         ? ' &nbsp;|&nbsp; <span style="color:#5BBFFF;font-weight:700;">📊 ' + this._esc(this._sacModelName || this._sacModelId) + '</span>' :
                this._sacModelState === 'error'      ? ' &nbsp;|&nbsp; <span style="color:#E67E22;font-weight:600;" title="' + this._esc(this._sacModelName) + '">⚠️ ' + this._esc((this._sacModelName || 'Modelo no encontrado').slice(0, 60)) + '</span>' :
                ''
              }</div>
            </div>
          </div>
          <div style="display:flex;align-items:center;gap:12px;">
            <span class="sap-badge-env">☁️ SAP Analytics Cloud</span>
            <button class="help-icon-btn" id="toggleHelpBtn" type="button" title="Interactive Help Guide">
              <span style="background:#0070F2;color:#FFF;width:18px;height:18px;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;font-size:11px;font-weight:900;">?</span>
              <span>Help</span>
            </button>
          </div>
        </div>

        ${this._toast ? '<div class="toast-banner toast-' + this._toast.type + '"><span>' + this._esc(this._toast.msg) + '</span></div>' : ''}

        <div class="nav-tabs">
          <div class="nav-tab ${this._activeTab === 'uploaderTab' ? 'active' : ''}" data-target="uploaderTab">🚀 Planning Upload Wizard</div>
          <div class="nav-tab ${this._activeTab === 'rulesTab' ? 'active' : ''}" data-target="rulesTab">⚙️ Validation Rules & Assisted Code (${activeRulesCount})</div>
          <div class="nav-tab ${this._activeTab === 'transformTab' ? 'active' : ''}" data-target="transformTab">🔄 Transform &amp; Convert</div>
          <div class="nav-tab ${this._activeTab === 'monitorTab' ? 'active' : ''}" data-target="monitorTab">🛡️ Audit Monitor (${this._history.length})</div>
        </div>

        <div id="uploaderTab" class="tab-content" style="display:${this._activeTab === 'uploaderTab' ? 'block' : 'none'};">
          <div class="wizard-steps">
            <div class="step-item ${this._step === 1 ? 'active' : (this._filesQueue.length > 0 ? 'completed' : '')}" data-step="1">
              <div class="step-num">${this._filesQueue.length > 0 && this._step > 1 ? '✓' : '1'}</div>
              <div>
                <div class="step-label-title">📂 Full/Delta & Files</div>
                <div class="step-label-sub">Strategy, multi-Excel & CSV</div>
              </div>
            </div>
            <div class="step-item ${this._step === 2 ? 'active' : (this._step > 2 ? 'completed' : '')}" data-step="2">
              <div class="step-num">${this._step > 2 ? '✓' : '2'}</div>
              <div>
                <div class="step-label-title">📑 Sheets & Semaphores</div>
                <div class="step-label-sub">Tabs, scope & rules</div>
              </div>
            </div>
            <div class="step-item ${this._step === 3 ? 'active' : ''}" data-step="3">
              <div class="step-num">3</div>
              <div>
                <div class="step-label-title">✨ Grid & Auto-Fix</div>
                <div class="step-label-sub">Resolution & SAC sync</div>
              </div>
            </div>
          </div>

          <div class="strategy-bar-card">
            <div style="display:flex;align-items:center;gap:10px;">
              <span style="font-size:11px;font-weight:700;color:#556B82;text-transform:uppercase;letter-spacing:.5px;">Load Mode</span>
            </div>
            <div class="strategy-toggle-group">
              <button class="strategy-btn ${this._loadStrategy === 'FULL' ? 'active' : ''}" data-strategy="FULL" type="button">
                <span style="font-size:18px;">📦</span>
                <span>
                  <span class="strategy-btn-title">FULL LOAD</span>
                  <span class="strategy-btn-sub">Replaces the planning slice</span>
                </span>
              </button>
              <button class="strategy-btn ${this._loadStrategy === 'DELTA' ? 'active' : ''}" data-strategy="DELTA" type="button">
                <span style="font-size:18px;">⚡</span>
                <span>
                  <span class="strategy-btn-title">DELTA LOAD</span>
                  <span class="strategy-btn-sub">Appends / incremental update</span>
                </span>
              </button>
            </div>
            <div style="display:flex;align-items:center;gap:8px;margin-left:auto;">
              ${this._loadStrategy === 'FULL'
                ? '<span style="background:#EBF5FF;color:#005CB9;border:1px solid #B3D4FF;padding:5px 12px;border-radius:20px;font-size:11px;font-weight:700;">📦 Full — overwrites target slice</span>'
                : '<span style="background:#E6F9EC;color:#107E3E;border:1px solid #A3E6BF;padding:5px 12px;border-radius:20px;font-size:11px;font-weight:700;">⚡ Delta — appends records</span>'}
            </div>
            <div style="width:100%;display:flex;align-items:flex-end;gap:14px;flex-wrap:wrap;padding-top:12px;border-top:1px dashed #D9E0E6;margin-top:4px;">
              <span style="font-weight:800;color:#1D2D3E;font-size:12px;padding-bottom:2px;">🎯 SAC Load Target:</span>
              <div>
                <label style="font-size:10px;font-weight:700;color:#556B82;display:block;margin-bottom:3px;">VERSION</label>
                <input list="sacVersionList" id="targetVersionInp" class="input-std" style="width:170px;"
                  placeholder="public.Actual" value="${this._esc(this._targetVersion)}" />
                <datalist id="sacVersionList">
                  <option value="public.Actual">
                  <option value="public.Budget">
                  <option value="public.Forecast">
                  <option value="public.Plan">
                </datalist>
              </div>
              <div>
                <label style="font-size:10px;font-weight:700;color:#556B82;display:block;margin-bottom:3px;">PERIOD (YYYYMM)</label>
                <input type="text" id="targetDateInp" class="input-std" style="width:110px;"
                  placeholder="e.g. 202601" maxlength="7" value="${this._esc(this._targetDate)}" />
              </div>
              <div style="position:relative;">
                <label style="font-size:10px;font-weight:700;color:#556B82;display:block;margin-bottom:3px;">MODEL ID ⭐</label>
                <div style="display:flex;gap:4px;align-items:center;">
                  <input type="text" id="sacModelIdInp" class="input-std" style="width:180px;"
                    placeholder="PLAN_CAPEX_PLANNING" value="${this._esc(this._sacModelId)}" title="SAC planning model technical name — required for DIS write-back" />
                  <button id="browseModelsBtn" class="btn btn-ghost" type="button" style="padding:5px 9px;font-size:11px;white-space:nowrap;" title="Buscar modelos disponibles en SAC">
                    ${this._sacModelsLoading ? '⏳' : '🔍 SAC'}
                  </button>
                </div>
                ${this._showModelPicker ? `
                  <div id="modelPickerDropdown" style="position:absolute;z-index:9999;top:calc(100% + 2px);left:0;min-width:360px;max-height:280px;overflow-y:auto;background:#fff;border:1px solid #0070F2;border-radius:8px;box-shadow:0 4px 16px rgba(0,0,0,0.15);padding:4px 0;">
                    <div style="display:flex;justify-content:space-between;align-items:center;padding:6px 12px;border-bottom:1px solid #E9EEF3;background:#F8FAFC;">
                      <span style="font-weight:700;font-size:11px;color:#1D2D3E;">Modelos disponibles en SAC DIS</span>
                      <button id="closeModelPickerBtn" type="button" style="background:none;border:none;cursor:pointer;font-size:14px;color:#556B82;padding:0 2px;">✕</button>
                    </div>
                    ${this._sacModelsLoading
                      ? '<div style="padding:20px;text-align:center;color:#556B82;font-size:12px;">⏳ Cargando modelos…</div>'
                      : this._sacModels.length === 0
                        ? '<div style="padding:20px;text-align:center;color:#BB0000;font-size:12px;">Sin modelos disponibles o sin permisos DIS</div>'
                        : this._sacModels.map(m => `
                          <div class="sac-model-pick-row" data-model-id="${this._esc(m.id)}" data-import-type="${this._esc(m.importType || 'privateFactData')}"
                            style="padding:7px 12px;cursor:pointer;border-bottom:1px solid #F0F4F8;display:flex;justify-content:space-between;align-items:center;">
                            <div>
                              <div style="font-weight:700;font-size:12px;color:#0064D9;">${this._esc(m.id)}</div>
                              ${m.name && m.name !== m.id ? `<div style="font-size:10px;color:#556B82;">${this._esc(m.name)}</div>` : ''}
                            </div>
                            ${m.importType ? `<span style="font-size:10px;background:#EBF5FF;color:#005CB9;padding:2px 7px;border-radius:10px;font-weight:600;">${this._esc(m.importType)}</span>` : ''}
                          </div>`).join('')
                    }
                  </div>` : ''}
              </div>
              ${this._targetVersion && this._targetDate
                ? (this._sacModelId
                    ? '<span class="sac-target-badge" style="background:#E6F9EC;color:#107E3E;border:1px solid #A3E6BF;padding:5px 14px;border-radius:20px;font-size:11px;font-weight:700;">✅ ' + this._esc(this._targetVersion) + ' · ' + this._esc(this._targetDate) + ' → ' + this._esc(this._sacModelId) + '</span>'
                    : '<span class="sac-target-badge" style="background:#E6F9EC;color:#107E3E;border:1px solid #A3E6BF;padding:5px 14px;border-radius:20px;font-size:11px;font-weight:700;">✅ ' + this._esc(this._targetVersion) + ' · ' + this._esc(this._targetDate) + ' <span style=\'color:#B45309\'>(sin Model ID — solo evento)</span></span>'
                  )
                : '<span class="sac-target-badge" style="background:#FFF8E6;color:#B45309;border:1px solid #FCD34D;padding:5px 14px;border-radius:20px;font-size:11px;font-weight:700;">⚠️ Configura Version y Period antes de sincronizar</span>'
              }
            </div>
          </div>

          ${this._step === 1 ? `
            ${this._renderFormatConfigBar(false)}
            <div class="drop-zone" id="dropZone">
              <div class="upload-icon-circle">☁️</div>
              <h3 style="margin:0;color:#1D2D3E;font-size:16px;">📥 Drag & drop two or more planning files here</h3>
              <p style="margin:0;color:#556B82;font-size:12px;max-width:560px;">Support for multiple workbooks, individual sheet selection, CSV/TXT delimiters, and scoped validation rules with semaphores.</p>
              <div style="margin-top:8px;">
                <label class="file-label-btn" for="fileInput">📂 Browse & Select Files</label>
                <input type="file" id="fileInput" multiple accept=".xlsx, .xls, .xlsm, .csv, .txt" ${!this._xlsxReady ? 'disabled' : ''} style="display:none;" />
              </div>
              <div style="margin-top:6px;">
                <button class="btn btn-ghost" id="downloadTemplateBtn" type="button" ${!this._xlsxReady ? 'disabled' : ''}>📥 Download Official Planning Template (.xlsx)</button>
              </div>
            </div>
          ` : ''}

          ${this._step === 2 ? `
            <div>
              ${this._renderFormatConfigBar(true)}
              <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:12px;flex-wrap:wrap;gap:8px;">
                <div>
                  <h3 style="margin:0;color:#1D2D3E;">📚 Loaded Workbooks & Sheet Selector (${this._filesQueue.length} file(s)) — Mode: <span style="color:#005CB9;">${this._loadStrategy}</span></h3>
                  <span style="font-size:11px;color:#556B82;">Check or uncheck the sheets inside each workbook that you want to include in the consolidation:</span>
                </div>
                <div style="display:flex;gap:8px;">
                  <label class="btn" for="fileInput" style="cursor:pointer;">➕ Add More Files<input type="file" id="fileInput" multiple accept=".xlsx, .xls, .xlsm, .csv, .txt" style="display:none;" /></label>
                  ${this._filesQueue.length > 0 ? '<button class="btn btn-danger" id="clearAllFilesBtn" type="button">🗑️ Remove All Files</button>' : ''}
                </div>
              </div>

              <div style="margin-bottom:20px;">
                ${this._filesQueue.map((item, fIdx) => {
                  const activeSheetCount = (item.sheets || []).filter(s => s.selected).length;
                  const activeRowCount = (item.sheets || []).filter(s => s.selected).reduce((acc, s) => acc + s.rows.length, 0);
                  return `
                    <div class="file-card-item">
                      <div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px;">
                        <div>
                          <div style="font-weight:800;color:#0064D9;font-size:13px;">📄 ${this._esc(item.file.name)} <span style="background:#EBF5FF;color:#005CB9;font-size:10px;padding:2px 6px;border-radius:4px;margin-left:6px;">${this._esc(item.ext)}</span></div>
                          <div style="font-size:11px;color:#556B82;margin-top:2px;"><strong>${activeSheetCount}</strong> of ${item.sheetCount} sheet(s) selected • <strong>${activeRowCount}</strong> active planning rows</div>
                        </div>
                        <div style="display:flex;align-items:center;gap:8px;">
                          <input type="text" class="input-std file-comment-input" style="width:250px;" data-index="${fIdx}" placeholder="💬 Business audit comment..." value="${this._esc(item.comment || '')}" />
                          <button class="btn btn-danger remove-file-btn" data-index="${fIdx}" type="button" style="padding:6px 12px;">✕ Remove</button>
                        </div>
                      </div>
                      <div style="margin-top:10px;padding-top:10px;border-top:1px dashed #D9E0E6;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px;">
                        <div style="display:flex;align-items:center;flex-wrap:wrap;gap:6px;">
                          <span style="font-size:11px;font-weight:700;color:#1D2D3E;margin-right:4px;">📑 Sheets:</span>
                          ${(item.sheets || []).map((sh, sIdx) => '<label class="sheet-pill-chk ' + (sh.selected ? 'sheet-active' : '') + '"><input type="checkbox" class="sheet-toggle-chk" data-fidx="' + fIdx + '" data-sidx="' + sIdx + '" ' + (sh.selected ? 'checked' : '') + ' style="margin:0;" /><span>' + this._esc(sh.name) + ' <strong>(' + sh.rows.length + ')</strong></span></label>').join('')}
                        </div>
                        <div style="display:flex;gap:4px;">
                          <button class="icon-btn select-all-sheets-btn" data-fidx="${fIdx}" data-mode="all" type="button">☑ Select All</button>
                          <button class="icon-btn select-all-sheets-btn" data-fidx="${fIdx}" data-mode="none" type="button">☐ None</button>
                        </div>
                      </div>
                    </div>
                  `;
                }).join('')}
              </div>

              <div style="display:grid;grid-template-columns:1fr 1fr;gap:14px;">
                <div style="background:#FFFFFF;border:1px solid #D9E0E6;border-radius:8px;padding:16px;">
                  <div style="font-weight:700;color:#1D2D3E;margin-bottom:10px;">🧱 Active Model Columns (Required Schema):</div>
                  <div style="display:flex;flex-wrap:wrap;gap:8px;margin-bottom:14px;">
                    ${this._requiredColumns.map(c => '<span class="col-pill">' + this._esc(c) + (c !== 'ID' ? '<button class="remove-col-btn" data-col="' + this._esc(c) + '" title="Remove column">✕</button>' : '') + '</span>').join('')}
                  </div>
                  ${this._showAddColModal ? '<div style="display:flex;gap:8px;align-items:center;"><input type="text" id="newColInput" class="input-std" placeholder="COLUMN_NAME" /><button class="btn" id="confirmAddColBtn" type="button">Save</button><button class="btn btn-ghost" id="cancelAddColBtn" type="button">Cancel</button></div>' : '<button class="btn btn-ghost" id="openAddColBtn" type="button">➕ Add Mandatory Column</button>'}
                </div>

                <div style="background:#FFFFFF;border:1px solid #D9E0E6;border-radius:8px;padding:16px;">
                  <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;">
                    <div style="font-weight:700;color:#1D2D3E;">🚦 Active Rules, Semaphores & Scopes (${activeRulesCount}):</div>
                    <div style="display:flex;gap:6px;">
                      <button class="btn btn-ghost" id="toggleInlineRuleBtn" type="button" style="padding:4px 10px;font-size:11px;">${this._showInlineRuleCreator ? 'Close Form' : '➕ New Rule'}</button>
                      <button class="btn btn-ghost" id="goToRulesTabBtn" type="button" style="padding:4px 10px;font-size:11px;">⚙️ Rules Manager →</button>
                    </div>
                  </div>
                  <div style="display:flex;flex-wrap:wrap;gap:6px;">
                    ${this._validationRules.map(r => '<span style="display:inline-flex;align-items:center;gap:6px;background:#F8FAFC;border:1px solid #D9E0E6;padding:4px 8px;border-radius:6px;font-size:11px;font-weight:600;"><input type="checkbox" class="toggle-rule-chk" data-rule-id="' + this._esc(r.id) + '" ' + (r.active ? 'checked' : '') + ' style="margin:0;" />' + this._renderSemaphoreDot(r.semaphore || 'RED') + '<strong>' + this._esc(r.col) + '</strong> [' + this._esc(r.scopeFiles || 'ALL') + ']<button class="remove-col-btn delete-rule-btn" data-rule-id="' + this._esc(r.id) + '">✕</button></span>').join('')}
                  </div>
                </div>
              </div>
              ${this._showInlineRuleCreator ? this._renderRuleCreatorForm() : ''}
            </div>
            <div class="wizard-footer">
              <button class="btn btn-ghost" id="prevStepBtn" type="button">← Back</button>
              <button class="btn" id="nextStep3Btn" type="button">✨ Validate & Edit ${totalRows} Rows →</button>
            </div>
          ` : ''}

          ${this._step === 3 ? `
            <div>
              <div class="metrics-bar">
                <div class="metric-box"><div class="metric-title">🔄 Strategy & Signal</div><div class="metric-val" style="font-size:12px;color:#0064D9;display:flex;align-items:center;margin-top:3px;">${this._renderSemaphoreDot(globalSemaphore)} ${this._loadStrategy} • ${globalSemaphoreText}</div></div>
                <div class="metric-box kpi-ok"><div class="metric-title">🟢 Clean Rows</div><div class="metric-val" style="color:#107E3E;">${stats.cleanRows}</div></div>
                <div class="metric-box kpi-err"><div class="metric-title">🔴 Blocking</div><div class="metric-val" style="color:#BB0000;">${stats.blockingRows}</div></div>
                <div class="metric-box kpi-prev"><div class="metric-title">🟡 Preventive</div><div class="metric-val" style="color:#E9730C;">${stats.preventiveRows}</div></div>
                <div class="metric-box kpi-chk"><div class="metric-title">🟣 Check</div><div class="metric-val" style="color:#6D28D9;">${stats.checkRows}</div></div>
                <div class="metric-box"><div class="metric-title">🔵 Info</div><div class="metric-val" style="color:#0070F2;">${stats.infoRows}</div></div>
              </div>

              <div class="dist-bar"><div class="dist-ok" style="width:${validPercent}%;"></div></div>

              <div class="toolbar-box">
                <div style="display:flex;align-items:center;gap:8px;"><span style="font-weight:700;font-size:12px;">🔍 Quick Search:</span><input type="text" id="searchInput" class="input-std" style="width:190px;" placeholder="Filter by ID, text, sheet..." value="${this._esc(this._searchQuery)}" /></div>
                <div style="display:flex;align-items:center;gap:6px;flex-wrap:wrap;">
                  <span style="font-weight:700;font-size:12px;">⚡ Bulk Fill:</span>
                  <select id="bulkColSelect" class="input-std">${this._requiredColumns.map(c => '<option value="' + this._esc(c) + '" ' + (this._bulkCol === c ? 'selected' : '') + '>' + this._esc(c) + '</option>').join('')}</select>
                  <input type="text" id="bulkValInput" class="input-std" style="width:110px;" placeholder="Value (e.g. 1000)" value="${this._esc(this._bulkVal)}" />
                  <button class="btn btn-ghost" id="applyBulkEmptyBtn" type="button" style="padding:5px 10px;font-size:11px;">Fill Empty</button>
                  <button class="btn btn-ghost" id="applyBulkAllBtn" type="button" style="padding:5px 10px;font-size:11px;">Overwrite Visible</button>
                </div>
              </div>

              ${this._showAddColModal ? '<div class="toolbar-box" style="justify-content:flex-start;"><span style="font-weight:700;">🧱 Column Name:</span><input type="text" id="newColInput" class="input-std" placeholder="E.g. CURRENCY" /><button class="btn" id="confirmAddColBtn" type="button">Add Column</button><button class="btn btn-ghost" id="cancelAddColBtn" type="button">Cancel</button></div>' : ''}
              ${this._showInlineRuleCreator ? this._renderRuleCreatorForm() : ''}

              <div class="chips-row">
                <div class="chips-left">
                  <div class="chip chip-all ${this._issueFilter === 'all' ? 'active-chip' : ''}" data-filter="all">📋 All (${totalRows})</div>
                  <div class="chip chip-ok ${this._issueFilter === 'correct' ? 'active-chip' : ''}" data-filter="correct">🟢 Clean (${stats.cleanRows})</div>
                  <div class="chip chip-err ${this._issueFilter === 'BLOCKING' ? 'active-chip' : ''}" data-filter="BLOCKING">🔴 Blocking (${stats.blockingRows})</div>
                  <div class="chip chip-prev ${this._issueFilter === 'PREVENTIVE' ? 'active-chip' : ''}" data-filter="PREVENTIVE">🟡 Preventive (${stats.preventiveRows})</div>
                  <div class="chip chip-chk ${this._issueFilter === 'CHECK' ? 'active-chip' : ''}" data-filter="CHECK">🟣 Check (${stats.checkRows})</div>
                  <div class="chip chip-info ${this._issueFilter === 'INFO' ? 'active-chip' : ''}" data-filter="INFO">🔵 Info (${stats.infoRows})</div>
                </div>
                <div style="display:flex;gap:6px;align-items:center;flex-wrap:wrap;">
                  ${stats.fixableCells > 0 ? '<button class="btn btn-success" id="autoFixAllBtn" type="button" style="font-size:11px;padding:6px 12px;">✨ Auto-Fix All (' + stats.fixableCells + ')</button>' : ''}
                  <button class="btn btn-ghost" id="addRowBtn" type="button" style="font-size:11px;padding:6px 12px;">➕ Add Row</button>
                  <button class="btn btn-ghost" id="openAddColBtn" type="button" style="font-size:11px;padding:6px 12px;">🧱 Add Column</button>
                  <button class="btn btn-ghost" id="toggleInlineRuleBtn" type="button" style="font-size:11px;padding:6px 12px;">⚙️ + Rule</button>
                  <button class="btn btn-ghost" id="exportCleanBtn" type="button" style="font-size:11px;padding:6px 12px;">📤 Export (.xlsx)</button>
                  ${(stats.blockingRows + stats.preventiveRows + stats.checkRows + stats.infoRows) > 0 ? '<button class="btn btn-danger" id="downloadLogBtn" type="button" style="font-size:11px;padding:6px 12px;">📥 Alerts Log (.csv)</button>' : ''}
                </div>
              </div>

              <div class="tbl-wrap">
                <table>
                  <thead>
                    <tr>
                      <th style="width:28px;">#</th>
                      <th style="width:76px;text-align:center;">🛠️ Actions</th>
                      <th>📂 Source Workbook & Sheet</th>
                      ${this._requiredColumns.map(h => '<th><div class="th-content"><span>' + this._esc(h) + '</span>' + (h !== 'ID' ? '<button class="remove-col-btn" data-col="' + this._esc(h) + '">✕</button>' : '') + '</div></th>').join('')}
                    </tr>
                  </thead>
                  <tbody>
                    ${filteredRows.map(({ row, origIdx }) => `
                      <tr>
                        <td style="color:#556B82;font-size:11px;font-weight:600;">${row._rowNum}</td>
                        <td><div class="row-actions"><button class="icon-btn dup-row-btn" data-orig-idx="${origIdx}" title="Clone row" type="button">⧉</button><button class="icon-btn icon-btn-del remove-row-btn" data-orig-idx="${origIdx}" title="Delete row" type="button">✕</button></div></td>
                        <td style="font-size:11px;color:#0064D9;font-weight:700;">${this._esc(row._sourceFile)}</td>
                        ${this._requiredColumns.map(col => {
                          const cellAlerts = (row._alerts && row._alerts[col]) || [];
                          const editedInfo = row._editedCols && row._editedCols[col];
                          let topSev = '';
                          if (cellAlerts.some(a => a.severity === 'BLOCKING')) topSev = 'BLOCKING';
                          else if (cellAlerts.some(a => a.severity === 'PREVENTIVE')) topSev = 'PREVENTIVE';
                          else if (cellAlerts.some(a => a.severity === 'CHECK')) topSev = 'CHECK';
                          else if (cellAlerts.some(a => a.severity === 'INFO')) topSev = 'INFO';
                          return '<td class="' + (topSev ? 'cell-sev-' + topSev : '') + '" style="' + (!topSev && editedInfo ? 'background:#FFFBEB;' : '') + '">' +
                            '<span class="editable-cell" contenteditable="true" data-orig-idx="' + origIdx + '" data-col="' + this._esc(col) + '">' + this._esc(row[col] ?? '') + '</span>' +
                            cellAlerts.map(a => '<div class="alert-line alert-line-' + a.severity + '"><span>' + this._renderSemaphoreDot(a.semaphore) + ' ' + this._esc(a.msg) + '</span><span style="display:inline-flex;gap:3px;">' + (a.autoFix ? '<button class="mini-fix-btn cell-autofix-btn" data-orig-idx="' + origIdx + '" data-col="' + this._esc(col) + '" data-rule-id="' + this._esc(a.ruleId) + '" type="button">✨ Fix</button>' : '') + ((a.severity === 'CHECK' || a.severity === 'PREVENTIVE') ? '<button class="mini-chk-btn cell-check-ok-btn" data-orig-idx="' + origIdx + '" data-col="' + this._esc(col) + '" data-rule-id="' + this._esc(a.ruleId) + '" type="button">✓ OK</button>' : '') + '</span></div>').join('') +
                            (cellAlerts.length === 0 && editedInfo ? '<span style="display:block;font-size:9px;color:#B45309;font-weight:600;margin-top:2px;">✎ Edited (was: ' + this._esc(editedInfo.from || 'Blank') + ')</span>' : '') +
                            (function() { const ci = row._convLog && row._convLog[col]; return (!editedInfo && ci) ? '<span style="display:block;font-size:9px;color:#0064D9;font-weight:600;margin-top:2px;">🔄 ' + ci.mode + ': ' + ci.from + '</span>' : ''; })() +
                            '</td>';
                        }).join('')}
                      </tr>
                    `).join('')}
                  </tbody>
                </table>
              </div>
            </div>
            <div class="wizard-footer">
              <button class="btn btn-ghost" id="prevStepBtn" type="button">← Back</button>
              <div style="display:flex;align-items:center;gap:14px;flex-wrap:wrap;">
                ${stats.blockingRows === 0 && (stats.preventiveRows > 0 || stats.checkRows > 0) ? '<label style="display:flex;align-items:center;gap:6px;background:#FFF8E6;border:1px solid #FCD34D;color:#B45309;padding:6px 12px;border-radius:6px;font-size:11px;font-weight:700;cursor:pointer;"><input type="checkbox" id="ackNonBlockingChk" ' + (this._acknowledgeNonBlocking ? 'checked' : '') + ' style="margin:0;" /><span>Confirm review of ' + stats.preventiveRows + ' preventive and ' + stats.checkRows + ' check alerts</span></label>' : ''}
                <button class="btn btn-ghost" id="exitBtn" type="button">🔄 Reset Queue</button>
                <button class="btn" id="finishBtn" type="button" ${stats.blockingRows > 0 ? 'style="opacity:0.55;"' : ''}>🚀 Consolidate & Synchronize with SAC (${this._loadStrategy})</button>
              </div>
            </div>
          ` : ''}
        </div>

        <div id="transformTab" class="tab-content" style="display:${this._activeTab === 'transformTab' ? 'block' : 'none'};">
          ${this._renderTransformConvertTab()}
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

    // ── Transform & Convert handlers ──────────────────────────────────────────
    const shadow = this._shadow;
    const runTransformBtn = shadow.getElementById('runTransformBtn');
    if (runTransformBtn) {
      runTransformBtn.addEventListener('click', () => {
        this._applyTransformAndConvert();
        this._validateRows();
        this._notify('Transform & Convert applied to ' + this._rows.length + ' row(s).', 'success');
        this._render();
      });
    }

    shadow.querySelectorAll('.trf-src-sel').forEach(sel => {
      sel.addEventListener('change', (e) => {
        const dim = e.target.dataset.dim;
        if (!this._transformMap[dim]) this._transformMap[dim] = {};
        this._transformMap[dim].source = e.target.value;
        this._saveTransformMap();
      });
    });

    shadow.querySelectorAll('.trf-fixed-inp').forEach(inp => {
      inp.addEventListener('input', (e) => {
        const dim = e.target.dataset.dim;
        if (!this._transformMap[dim]) this._transformMap[dim] = {};
        this._transformMap[dim].fixed = e.target.value;
        this._saveTransformMap();
      });
    });

    shadow.querySelectorAll('.conv-dim-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        this._convTabActiveCol = e.currentTarget.dataset.dim;
        this._render();
      });
    });

    shadow.querySelectorAll('.conv-fallback-sel').forEach(sel => {
      sel.addEventListener('change', (e) => {
        const dim = e.target.dataset.dim;
        this._convTables[dim + '__fallback'] = e.target.value;
        this._saveConvTables();
        this._render();
      });
    });

    shadow.querySelectorAll('.conv-default-inp').forEach(inp => {
      inp.addEventListener('input', (e) => {
        const dim = e.target.dataset.dim;
        this._convTables[dim + '__default'] = e.target.value;
        this._saveConvTables();
      });
    });

    // Inline rule field editing
    shadow.querySelectorAll('.conv-rule-mode, .conv-rule-src, .conv-rule-tgt, .conv-rule-start, .conv-rule-end').forEach(inp => {
      inp.addEventListener('change', (e) => {
        const dim  = e.target.dataset.dim;
        const rIdx = parseInt(e.target.dataset.ridx, 10);
        if (!this._convTables[dim] || !this._convTables[dim][rIdx]) return;
        const rule = this._convTables[dim][rIdx];
        const cl   = [...e.target.classList];
        if (cl.includes('conv-rule-mode'))  { rule.mode   = e.target.value; }
        if (cl.includes('conv-rule-src'))   { rule.source = e.target.value; }
        if (cl.includes('conv-rule-tgt'))   { rule.target = e.target.value; }
        if (cl.includes('conv-rule-start')) { rule.sliceStart = e.target.value !== '' ? Number(e.target.value) : undefined; }
        if (cl.includes('conv-rule-end'))   { rule.sliceEnd   = e.target.value !== '' ? Number(e.target.value) : undefined; }
        this._saveConvTables();
        this._render();
      });
    });

    shadow.querySelectorAll('.conv-del-rule-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const dim  = e.currentTarget.dataset.dim;
        const rIdx = parseInt(e.currentTarget.dataset.ridx, 10);
        if (this._convTables[dim]) {
          this._convTables[dim].splice(rIdx, 1);
          this._saveConvTables();
          this._render();
        }
      });
    });

    const addConvRuleBtn = shadow.getElementById('addConvRuleBtn');
    if (addConvRuleBtn) {
      addConvRuleBtn.addEventListener('click', () => {
        const dim      = addConvRuleBtn.dataset.dim;
        const mode     = (shadow.getElementById('newConvMode')?.value  || 'EXACT').toUpperCase();
        const src      = (shadow.getElementById('newConvSrc')?.value   || '').trim();
        const tgt      = (shadow.getElementById('newConvTgt')?.value   || '').trim();
        const startVal = shadow.getElementById('newConvStart')?.value  || '';
        const endVal   = shadow.getElementById('newConvEnd')?.value    || '';
        if (!this._convTables[dim]) this._convTables[dim] = [];
        const newRule = { mode, source: src, target: tgt };
        if (mode === 'SLICE') {
          if (startVal !== '') newRule.sliceStart = Number(startVal);
          if (endVal   !== '') newRule.sliceEnd   = Number(endVal);
        }
        this._convTables[dim].push(newRule);
        this._saveConvTables();
        this._notify('Rule added for ' + dim + '.', 'success');
        this._render();
      });
    }

    const convImportBtn  = shadow.getElementById('convImportBtn');
    const convFileInput  = shadow.getElementById('convFileInput');
    if (convImportBtn && convFileInput) {
      convImportBtn.addEventListener('click',  () => convFileInput.click());
      convFileInput.addEventListener('change', (e) => {
        if (e.target.files && e.target.files[0]) {
          this._importConversionFile(e.target.files[0]);
          e.target.value = '';
        }
      });
    }

  }
}

if (!customElements.get('custom-excel-uploader-pro')) {
  customElements.define('custom-excel-uploader-pro', ExcelUploaderPro);
}