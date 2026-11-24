/**
 * ExcelUploaderPro — SAP Analytics Cloud Custom Widget
 * Version: 2.6.0
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
if (typeof ExcelUploaderPro === 'undefined') {

const _EPM_VERSION = '3.9.3';

const _EPM_CHANGELOG = [
  { v: '3.9.3', date: '2026-11-24', title: 'Fix: Update/Append — client-side merge (DIS v1 always Replaces)',
    items: ['Root cause confirmed: DIS v1 ignores writeBackMode/importMode/writeBack JSON body params — always replaces data for the uploaded Version+Period.',
            'Fix: for Update/Append modes, the widget now fetches existing data via OData export before uploading, merges the datasets in memory, then does a single Replace with the merged complete dataset.',
            'Update merge: existing rows not in the uploaded file are preserved; uploaded rows replace matching ones.',
            'Append merge: measure values are accumulated for matching dimension keys; new keys are inserted; all other existing rows are kept.',
            'If existing data cannot be fetched (OData endpoint unavailable), a warning is logged and Replace proceeds as fallback.',
            'All modes now use the same proven single-step CSV POST flow — the JSON job creation branch has been removed.',
            'New methods: _fetchExistingDataOData(version, date) and _mergePlanningData(uploadedRows, existingRows, mode, dimensions).'] },
  { v: '3.9.2', date: '2026-11-24', title: 'Fix: 5 bugs from code review',
    items: ['Version constant 3.9.0 → 3.9.2 (was not bumped).',
            'Critical: CSV upload was skipped if job initial status was not exactly READY_FOR_DATA — fixed with isNonReplace guard.',
            'Spanish timeout message translated to English.',
            'Hardcoded model name "PLAN_CAPEX_PLANNING" in _fetchDeltaPreview replaced with this._sacModelId.',
            'README version updated to 3.9.2.'] },
  { v: '3.9.1', date: '2026-11-24', title: 'All UI messages in English + Full load badge shows method',
    items: ['Translated all remaining Spanish strings: console logs, toasts, loading text, UI badges.',
            'Full load strategy badge now shows the active import method (Replace / Update / Append).',
            'Strategy bar badge for Full load: "📦 Full · 🗑️ Replace" / "📦 Full · ✏️ Update" / "📦 Full · ➕ Append".'] },
  { v: '3.9.0', date: '2026-11-24', title: 'Auto-load model dimensions via FPA REST · Real SAP user in Audit',
    items: ['_fetchModelDimensions() now first tries POST /sap/fpa/services/rest/fpa/model — the confirmed working endpoint on this tenant. Payload: {"action":"getDesignTime","data":{"resourceId":"{GUID}","metadata":{},"includeDatasphereInfos":true}}.',
            'New _discoverTenantUUID(): tries SAC global variables (window.__sap_umi_config), then GET /sap/fpa/services/rest/session. Calls the FPA endpoint both with and without ?tenant= param so it works across tenants.',
            'New _parseDimensionsFromFPA(): parses the DIMENSION array returned by FPA REST (metadata.id.name + data.type + data.dimensionId). VersionDim → isAuto, TimeHANA → isAuto.',
            'DIS GET endpoints kept as fallback (10 variants) in case the FPA REST endpoint is unavailable.',
            'Real SAP user auto-populated in Audit Monitor: session endpoint also returns the logged-in loginName / userName; replaces the placeholder "SAP User (Planning)" automatically without requiring manual configuration.',
            '_discoverTenantUUID() called at _fetchModelInfo() startup (non-blocking) so user and tenant are ready before the first sync.'] },
  { v: '3.8.0', date: '2026-11-24', title: 'Import Method selector + Delta data preview',
    items: ['New ⚙️ Import Method selector in the strategy bar: Replace / Update / Append. Full load offers all 3; Delta load offers Update (default) and Append.',
            'Method stored in localStorage and saved in the audit record per load.',
            'Method passed as ?method= query parameter to DIS job creation (honored by DIS versions that support it).',
            'Replace: clears the entire Version+Period region before writing. Update: updates matching keys, inserts new, preserves rest. Append: accumulates measure values for matching keys.',
            'New 🔍 Delta Preview card — appears when DELTA is selected. Click "Load Preview" to fetch existing data for the target Version+Period from SAC DIS and display it in a table.',
            'Preview gracefully degrades: if the DIS endpoint returns no data or is not supported, shows a helpful fallback message.',
            'Switching strategy auto-resets the method (DELTA → Update, FULL → keeps current unless was replaced) and clears any loaded preview.'] },
  { v: '3.7.0', date: '2026-11-24', title: 'VERSION + PERIOD in Audit Monitor · Delta behavior corrected',
    items: ['Audit Monitor list now shows VERSION and PERIOD columns for every load record.',
            'Audit Monitor detail view metrics bar expanded (5 → 7 KPI boxes) to include VERSION and PERIOD.',
            'VERSION and PERIOD are saved in the audit record at sync time, so they are always accurate even if the config changes later.',
            'Delta load description corrected: DIS Update mode updates matching dimension-key records, inserts new ones, and preserves all other existing data — it does not delete non-matching rows.',
            'Delta warning badge updated to: "Updates matching records, inserts new ones — existing data for other keys is preserved".'] },
  { v: '3.6.0', date: '2026-11-24', title: 'Publish-ready banner in Audit Monitor',
    items: ['After a successful DIS load (🟢 LOAD COMPLETE), the Audit Monitor detail view now shows a persistent green banner: "Data loaded — review the snapshot below, then publish when ready".',
            '"📢 How to Publish" button opens a step-by-step guide for publishing planning data in SAP Analytics Cloud.',
            'Publish modal covers permissions, Full vs Delta behaviour, and the SAC publish flow.'] },
  { v: '3.5.0', date: '2026-11-24', title: 'All UI messages in English + Delta load warning',
    items: ['All user-visible messages translated to English (toasts, audit status, model picker, DIS block).',
            'Delta load strategy now shows ⚠️ warning badge: "replaces matching records, does not sum amounts".',
            'Fixed misleading Delta description ("Appends / incremental update" → "Replaces matching key records").',
            'Delta console.log warning added in _callDISApi to surface the replace-not-sum behavior in devtools.'] },
  { v: '3.4.0', date: '2026-11-23', title: 'Detection of rows rejected by SAC DIS (invalidrecords)',
    items: ['After COMPLETED, the widget queries GET /jobs/{jobId}/invalidrecords.',
            'If rows are rejected: status changes to "⚠️ PARTIAL LOAD", toast shows the reason.',
            'Audit detail shows the DIS error message in a monospace block.',
            'Detects cases like "invalid dimension member" (e.g. CC_FIN_02) that SAC DIS rejects silently.'] },
  { v: '3.3.1', date: '2026-11-23', title: 'Fix: 3 bugs — double render DIS error, double render addConvRule, duplicate targetDate assignment',
    items: ['Removed redundant this._render() after this._notify() in DIS error catch (already re-rendered).',
            'Removed redundant this._render() after this._notify() in addConvRuleBtn handler.',
            'Removed duplicate assignment of this._targetDate in PERIOD field listener.'] },
  { v: '3.3.0', date: '2026-11-23', title: 'MODEL NAME always visible + audit messages with semaphore colours',
    items: ['MODEL NAME always shows the description without similarity filtering against the ID.',
            '_updateModelDesc() and template remove the "show" logic — always rendered.',
            'Load status in audit: "🟢 FULL LOAD", "⚠️ PARTIAL LOAD", "❌ DIS ERROR".',
            'Success toast: "🟢 FULL LOAD — LOAD-XXXXX · FULL — N rows written → SAC".',
            'DIS block in audit detail redesigned with semaphore signals (colour dot + thick border).'] },
  { v: '3.2.0', date: '2026-11-23', title: 'Version and Date visible in Column Mapping',
    items: ['Version and Date appear in the Column Mapping table with green background and "← SAC Load Target" label.',
            'Locked rows (🔒) — their values always come from the VERSION and PERIOD fields.',
            'Only shown when VERSION or PERIOD fields are filled.'] },
  { v: '3.1.1', date: '2026-11-23', title: 'Use SAP jobURL for data upload',
    items: ['Extracts jobURL from the job creation response body.',
            'Attempts POST to jobURL first (the canonical URL SAP returns).',
            'Fallback: /jobs/{id} and /jobs/{id}/data if jobURL fails with 404/405.'] },
  { v: '3.1.0', date: '2026-11-23', title: 'Fix DIS two-step: correct endpoint + abort if no data',
    items: ['Changes upload endpoint from /jobs/{id}/data to /jobs/{id} (canonical SAP DIS SDK endpoint).',
            'Fallback: if /jobs/{id} returns 405, automatically tries /jobs/{id}/data.',
            'Poll aborts after 15 s if job stays in READY_FOR_DATA (prevents infinite UI block).',
            'Improved log: "[DIS] data upload jobId HTTP: X" and "[DIS] data upload jobId/data HTTP: X".'] },
  { v: '3.0.1', date: '2026-11-23', title: 'Bug fixes: double render, poll match, clear button',
    items: ['Fix double _render() after Transform & Convert — _notify() already re-renders.',
            'Fix DIS poll failure detection — uses startsWith instead of strict equality (catches UPLOAD_FAILED variants).',
            '✕ Clear button now visible in mapping rows — removes custom name or exclusion for that column.'] },
  { v: '3.0.0', date: '2026-11-23', title: 'Mapping feedback + toast confirmations',
    items: ['💾 Apply Mapping shows toast: "✅ Mapping saved — X/N columns mapped" and scrolls to Step 1 if no rows loaded.',
            '⚡ Defaults shows toast: "⚡ Default mapping applied" after auto-mapping.',
            'No visible changes if 0 columns are loaded yet — toast shows how many columns were mapped.'] },
  { v: '2.9.0', date: '2026-11-23', title: 'Column Mapping simple + fix DIS two-step upload',
    items: ['New 🔗 Column Mapping panel: text inputs per file column, no model properties loaded.',
            'Fix DIS upload: detect READY_FOR_DATA after job creation → POST CSV to /jobs/{jobId}/data.',
            '⚡ Defaults: ID→CAPEX_ITEM, H1→AMOUNT, DESCRIPTION→excluded.',
            '⊘ Excl. button per row to exclude a column from DIS upload.'] },
  { v: '2.8.0', date: '2026-11-23', title: 'Expanded dimension probe: OData, FPA REST, Planning API',
    items: ['6 new endpoint variants added to dimension probe: OData $metadata (XML), SAC FPA REST, Planning REST API.',
            'XML OData $metadata parsed to extract EntityType Properties as dimensions.',
            'Total of 12 endpoint variants tried on connect — higher chance of auto-detecting model structure.'] },
  { v: '2.7.9', date: '2026-11-23', title: 'Fix: 412 on /run step handled silently',
    items: ['POST /jobs/{jobId}/run returning 412 (Precondition Failed) now skipped silently — same behaviour as 405.',
            'Job was already auto-executing on creation; widget proceeds directly to polling.'] },
  { v: '2.7.8', date: '2026-11-23', title: 'MODEL NAME editable + auto-populated from models list',
    items: ['MODEL NAME field is now editable — type the name and it is saved to localStorage.',
            'When DIS metadata endpoint returns 404, widget automatically searches the models list to find the model description.',
            'MODEL NAME auto-populated without needing to use the picker.'] },
  { v: '2.7.7', date: '2026-11-23', title: 'Fix: DESCRIPTION excluded from DIS + restored as default column',
    items: ['DESCRIPTION column now shows in the grid (useful to identify items) but is excluded from the CSV sent to DIS.',
            'Default Required Columns restored to match the Excel template: ID, DESCRIPTION, H1, COMPANY, COSTCENTER, ASSET_CLASS, CAPITALIZED.',
            'Column mapping ID→CAPEX_ITEM and H1→AMOUNT applied only to the CSV header — DIS always receives correct SAC dimension names.'] },
  { v: '2.7.6', date: '2026-11-23', title: 'Column Mapping — Excel names → SAC dimension names',
    items: ['New _columnMapping state: maps Excel column names to SAC dimension names in the CSV header.',
            'Default mapping pre-configured: ID → CAPEX_ITEM, H1 → AMOUNT.',
            'Data lookup uses both SAC name and Excel alias — values are always found regardless of Excel column names.',
            'Required Columns keep Excel names (what appears in the file); DIS always receives the correct SAC names.'] },
  { v: '2.7.5', date: '2026-11-23', title: 'Fix: auto-migrate old column names from localStorage',
    items: ['On startup, ID → CAPEX_ITEM and H1 → AMOUNT automatically in any saved schema.',
            'DESCRIPTION removed automatically if present. No manual localStorage reset needed.'] },
  { v: '2.7.4', date: '2026-11-23', title: 'Fix: default columns match PLAN_CAPEX_PLANNING model',
    items: ['Default Required Columns updated to match the actual SAC model dimensions: CAPEX_ITEM, COMPANY, COSTCENTER, ASSET_CLASS, CAPITALIZED, AMOUNT.',
            'ID → CAPEX_ITEM (dimension key). H1 → AMOUNT (measure). DESCRIPTION removed (not a DIS upload column).'] },
  { v: '2.7.3', date: '2026-11-23', title: 'What\'s New panel + single version constant',
    items: ['Added "What\'s New" button in header — shows version history inline without leaving the widget.',
            'Version string now defined as a single constant (_EPM_VERSION) so subtitle, header and modal always stay in sync.',
            'ID column can now be removed from Required Columns (✕ button was incorrectly disabled).'] },
  { v: '2.7.2', date: '2026-11-23', title: 'Fix: ID column removal + dimension probe logging',
    items: ['Removed guard that prevented deleting the "ID" column from the grid schema.',
            'Added per-endpoint diagnostic logging ([DIS] dim probe) to identify DIS API response format.'] },
  { v: '2.7.1', date: '2026-11-23', title: 'Fix: dimension fetch — CSRF endpoint chain + columnID parser',
    items: ['_fetchModelDimensions now tries 6 endpoint variants (CSRF, /dimensions, metadata) with GUID and technical name.',
            '_parseDimensions now handles SAC DIS "columns" array with "columnID" and "columnType" fields.',
            'Added SIGNEDDATA to measure types. Better fallback message when no dimensions returned.'] },
  { v: '2.7.0', date: '2026-11-23', title: 'Model Structure Validation panel',
    items: ['Inline collapsible panel showing model dimensions with mapping status (✅ auto / ✅ mapped / ❌ missing).',
            'Synchronize blocked when any ❌ missing dimension detected. "⚡ Add missing" button.',
            'Dimension data fetched automatically from SAC DIS after model connects.'] },
  { v: '2.6.1', date: '2026-11-23', title: 'Fix: 405 on /run step + MODEL NAME always visible',
    items: ['POST /jobs/{jobId}/run returning 405 now skipped silently — job auto-executes on this tenant.',
            'MODEL NAME field always rendered (placeholder "—" until description available).',
            'MODEL NAME auto-populated from DIS metadata on connect.'] },
  { v: '2.6.0', date: '2026-11-23', title: 'DIS single-step CSV + MODEL ID/NAME fields + redesigned picker',
    items: ['DIS upload now single-step: POST CSV directly to /models/{GUID}/{importType}.',
            'SAC Load Target: MODEL ID (editable) and MODEL NAME (readonly) shown side by side.',
            'Model picker: technical name as primary label, description as subtitle, GUID in monospace.',
            'Real-time search across name, description and GUID. Planning models listed first.'] },
  { v: '2.5.0', date: '2026-11-22', title: 'Critical DIS fixes — GUID resolution + CSRF',
    items: ['Automatic GUID resolution via _resolveDISModelId() — DIS requires internal GUID, not technical name.',
            'Models list called without ?importType= parameter (caused 404 on this tenant).',
            'CSRF token now fetched from model-specific endpoint. Double-load guard added.'] },
];

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
    this._showChangelog = false;
    this._showPublishModal = false;

    this._loadStrategy  = 'FULL';
    this._dspArtifactSync = true;
    this._targetVersion = this._storageGet('sac_target_version') || '';
    this._targetDate    = this._storageGet('sac_target_date')    || '';
    this._parseConfig = { format: 'AUTO', delimiter: 'AUTO', numberFormat: 'STD', headerRow: 1 };

    const defaultCols = ['ID', 'DESCRIPTION', 'H1', 'COMPANY', 'COSTCENTER', 'ASSET_CLASS', 'CAPITALIZED'];
    const savedCols = this._storageGet('sac_required_cols');
    this._requiredColumns = Array.isArray(savedCols) && savedCols.length > 0 ? savedCols : defaultCols;

    // Column mapping: Excel column name → SAC dimension name (applied to CSV header)
    const defaultMapping = { 'ID': 'CAPEX_ITEM', 'H1': 'AMOUNT' };
    const savedMapping = this._storageGet('sac_column_mapping');
    this._columnMapping = (savedMapping && typeof savedMapping === 'object') ? savedMapping : defaultMapping;

    // Columns excluded from DIS upload (shown in grid but not sent to SAC)
    const savedExcluded = this._storageGet('sac_dis_excluded_cols');
    this._disExcludedCols = Array.isArray(savedExcluded) ? savedExcluded : ['DESCRIPTION'];

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
    this._importMethod    = this._storageGet('sac_import_method')  || 'replace';
    this._deltaPreview    = { rows: [], cols: [], loading: false, error: null, loaded: false };
    this._sacModelName    = '';
    this._sacModelDesc    = this._storageGet('sac_dis_model_desc') || '';
    this._sacModelState   = '';
    this._sacDISModelId   = this._storageGet('sac_dis_guid') || '';
    this._sacModels       = [];
    this._sacModelsLoading = false;
    this._showModelPicker   = false;
    this._modelPickerFilter = '';
    this._sacModelDimensions  = [];   // [{id, description, type, isMeasure, isAuto}]
    this._sacDimLoading       = false;
    this._sacDimError         = '';
    this._sacDimPanelOpen     = true;
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
      // Discover tenant UUID + real SAP user in background (non-blocking)
      this._discoverTenantUUID(base).catch(() => {});
      const resp = await fetch(base + '/api/v1/dataimport/models/' + encodeURIComponent(this._sacModelId), {
        method: 'GET',
        credentials: 'include',
        headers: { 'x-csrf-token': 'fetch', 'Accept': 'application/json' }
      });
      if (resp.status === 404) {
        // SAC DIS metadata endpoint not available — try to find description in models list
        this._sacModelState = 'ok';
        this._sacModelName  = this._sacModelId;
        if (!this._sacModelDesc) {
          const listResp = await fetch(base + '/api/v1/dataimport/models', {
            credentials: 'include',
            headers: { 'Accept': 'application/json', 'x-sap-dis-request-source': 'fileUploadWidget' }
          }).catch(() => null);
          if (listResp && listResp.ok) {
            const listData = await listResp.json().catch(() => null);
            const list = listData?.value || listData?.d?.results || (Array.isArray(listData) ? listData : []);
            const match = list.find(m => (m.modelName || m.name || '') === this._sacModelId ||
                                         (m.id || '') === this._sacModelId ||
                                         (m.id || '') === this._sacDISModelId);
            if (match) {
              const desc = match.description || match.modelDescription || '';
              if (desc) {
                this._sacModelDesc = desc;
                this._storageSet('sac_dis_model_desc', desc);
                const descEl = this._shadow && this._shadow.querySelector('#sacModelDescDisplay');
                if (descEl) { descEl.value = desc; descEl.style.color = '#1D2D3E'; }
              }
            }
          }
        }
      } else if (!resp.ok) {
        throw new Error('HTTP ' + resp.status + (resp.status === 401 || resp.status === 403 ? ' — No DIS permissions' : ''));
      } else {
        const meta = await resp.json().catch(() => ({}));
        this._sacModelName  = meta.description || meta.name || this._sacModelId;
        this._sacModelState = 'ok';
        // populate description for MODEL NAME display field
        if (meta.description && meta.description !== this._sacModelId) {
          this._sacModelDesc = meta.description;
          this._storageSet('sac_dis_model_desc', meta.description);
        }
      }
    } catch (err) {
      this._sacModelName  = err.message || 'Connection error';
      this._sacModelState = 'error';
    }
    this._updateHeaderSubtitle();
    // update MODEL NAME input without full re-render
    const descEl = this._shadow && this._shadow.querySelector('#sacModelDescDisplay');
    if (descEl) {
      const desc = this._sacModelDesc || '';
      descEl.value = desc;
      descEl.style.color = desc ? '#1D2D3E' : '#A9B8C8';
    }
  }

  // ── Fetch model dimension structure from SAC ─────────────────────────────
  async _fetchModelDimensions() {
    const guid = this._sacDISModelId || this._sacModelId;
    const tech = this._sacModelId;
    if (!guid) return;
    this._sacDimLoading = true;
    this._sacDimError   = '';
    this._sacModelDimensions = [];
    this._render();
    try {
      const base = (this._sacTenantUrl || window.location.origin).replace(/\/+$/, '');
      let dims = [];

      // ── Priority 1: SAC FPA REST API (POST, confirmed working on this tenant) ──
      // POST /sap/fpa/services/rest/fpa/model[?tenant=uuid]
      // Body: {"action":"getDesignTime","data":{"resourceId":"{DIS_GUID}","metadata":{},"includeDatasphereInfos":true}}
      const fpaBody = JSON.stringify({ action: 'getDesignTime', data: { resourceId: guid, metadata: {}, includeDatasphereInfos: true } });
      const fpaEndpoint = base + '/sap/fpa/services/rest/fpa/model';
      const tenantUUID = await this._discoverTenantUUID(base);
      const fpaUrls = tenantUUID
        ? [fpaEndpoint + '?tenant=' + encodeURIComponent(tenantUUID), fpaEndpoint]
        : [fpaEndpoint];
      for (const url of fpaUrls) {
        try {
          const r = await fetch(url, {
            method: 'POST',
            credentials: 'include',
            headers: { 'Content-Type': 'application/json', 'Accept': 'application/json' },
            body: fpaBody
          });
          console.log('[DIS] FPA probe', r.status, url.replace(base, ''));
          if (r.ok) {
            const data = await r.json().catch(() => null);
            if (data) { dims = this._parseDimensionsFromFPA(data); }
            if (dims.length > 0) { console.log('[DIS] dimensions via FPA REST from:', url); break; }
          }
        } catch (_) { continue; }
      }

      // ── Priority 2: DIS fallback endpoints (GET) ─────────────────────────────
      if (dims.length === 0) {
        const iType = this._sacImportType || 'privateFactData';
        const attempts = [
          base + '/api/v1/dataimport/models/' + encodeURIComponent(guid) + '/' + iType,
          base + '/api/v1/dataimport/models/' + encodeURIComponent(tech) + '/' + iType,
          base + '/api/v1/dataimport/models/' + encodeURIComponent(guid) + '/dimensions',
          base + '/api/v1/dataimport/models/' + encodeURIComponent(tech) + '/dimensions',
          base + '/api/v1/dataimport/models/' + encodeURIComponent(guid),
          base + '/api/v1/dataimport/models/' + encodeURIComponent(tech),
          base + '/api/v1/dataexport/providers/sac/' + encodeURIComponent(tech) + '/$metadata',
          base + '/api/v1/dataexport/providers/sac/' + encodeURIComponent(guid) + '/$metadata',
          base + '/api/v1/planning/models/' + encodeURIComponent(tech),
          base + '/api/v1/planning/models/' + encodeURIComponent(guid),
        ];
        for (const url of attempts) {
          try {
            const r = await fetch(url, {
              credentials: 'include',
              headers: { 'x-csrf-token': 'fetch', 'Accept': 'application/json,application/xml,*/*', 'x-sap-dis-request-source': 'fileUploadWidget' }
            });
            const ct = r.headers.get('content-type') || '';
            if (ct.includes('xml')) {
              const xml = await r.text().catch(() => null);
              console.log('[DIS] dim probe', r.status, url.replace(base, ''), '→ XML', xml?.slice(0, 300));
              if (r.ok && xml) { dims = this._parseDimensionsFromXml(xml); }
            } else {
              const data = await r.json().catch(() => null);
              console.log('[DIS] dim probe', r.status, url.replace(base, ''), '→', JSON.stringify(data)?.slice(0, 200));
              if (r.ok && data) { dims = this._parseDimensions(data); }
            }
            if (dims.length > 0) { console.log('[DIS] dimensions from:', url); break; }
          } catch (_) { continue; }
        }
      }

      this._sacModelDimensions = dims;
      if (dims.length > 0) this._autoMapColumns();
      console.log('[DIS] model dimensions:', dims.map(d => d.id).join(', ') || '(none)');
      if (dims.length === 0) {
        this._sacDimError = 'Could not load model dimensions automatically. Configure column mapping manually.';
      }
    } catch (err) {
      this._sacDimError = 'Could not load model dimensions: ' + err.message;
      console.warn('[DIS] dimension fetch failed:', err.message);
    }
    this._sacDimLoading = false;
    this._render();
  }

  // Discover SAC tenant UUID (and optionally the current user) without hardcoding
  async _discoverTenantUUID(base) {
    if (this._cachedTenantUUID) return this._cachedTenantUUID;
    // Try SAC global variables first (cheapest)
    try {
      const uuid = window?.__sap_umi_config?.tenant
        || window?.sap?.umi?.config?.tenant
        || window?.sap?.authz?.tenantId
        || window?.SapAnalyticsCloud?.tenantId
        || null;
      if (uuid && /^[0-9a-f-]{36}$/i.test(uuid)) {
        this._cachedTenantUUID = uuid;
        return uuid;
      }
    } catch (_) {}
    const sessionBase = (base || window.location.origin);
    // SAC FPA session endpoint accepts GET with ?request= JSON param
    const sessionActions = [
      '{"action":"getUserInfo","data":{}}',
      '{"action":"getState","data":{}}',
      '{"action":"getUser","data":{}}',
      '{}',
    ];
    for (const action of sessionActions) {
      try {
        const url = sessionBase + '/sap/fpa/services/rest/session?request=' + encodeURIComponent(action);
        const r = await fetch(url, { credentials: 'include', headers: { 'Accept': 'application/json' } });
        if (r.ok) {
          const data = await r.json().catch(() => null);
          if (data) {
            const uuid = data?.tenantId || data?.tenant || data?.id || data?.data?.tenantId
              || data?.state?.tenantId || null;
            if (uuid && /^[0-9a-f-]{36}$/i.test(uuid)) this._cachedTenantUUID = uuid;
            const DEFAULT_USER = 'SAP User (Planning)';
            if (this._currentUser === DEFAULT_USER || !this._currentUser) {
              const u = data?.user?.loginName || data?.user?.name || data?.loginName
                || data?.userName || data?.user?.email || data?.state?.user?.loginName
                || data?.data?.user?.loginName || null;
              if (u) this._currentUser = u;
            }
            if (this._cachedTenantUUID) return this._cachedTenantUUID;
          }
        }
      } catch (_) {}
    }
    // Try POST to session endpoint as fallback
    try {
      const r = await fetch(sessionBase + '/sap/fpa/services/rest/session', {
        method: 'POST',
        credentials: 'include',
        headers: { 'Content-Type': 'application/json', 'Accept': 'application/json' },
        body: JSON.stringify({ action: 'getState', data: {} })
      });
      if (r.ok) {
        const data = await r.json().catch(() => null);
        if (data) {
          const uuid = data?.tenantId || data?.tenant || data?.id || null;
          if (uuid && /^[0-9a-f-]{36}$/i.test(uuid)) { this._cachedTenantUUID = uuid; return uuid; }
          const DEFAULT_USER = 'SAP User (Planning)';
          if ((this._currentUser === DEFAULT_USER || !this._currentUser)) {
            const u = data?.user?.loginName || data?.loginName || data?.userName || null;
            if (u) this._currentUser = u;
          }
        }
      }
    } catch (_) {}
    return null;
  }

  // Parse SAC FPA REST response: { DIMENSION: [{ metadata: { id: { name } }, data: { type, dimensionId, description } }] }
  _parseDimensionsFromFPA(data) {
    const dimArray = data?.DIMENSION;
    if (!Array.isArray(dimArray) || dimArray.length === 0) return [];
    const AUTO_TYPES    = new Set(['VERSIONDIM', 'TIMEHANA']);
    const MEASURE_TYPES = new Set(['MEASURE', 'KPI', 'AMOUNT', 'QUANTITY', 'SIGNEDDATA']);
    return dimArray.map(d => {
      const dimId   = d?.metadata?.id?.name || d?.data?.dimensionId || '';
      const dimType = (d?.data?.type || '').toUpperCase();
      const desc    = d?.data?.description || dimId;
      const isAuto  = AUTO_TYPES.has(dimType) || dimId.toUpperCase() === 'VERSION' || dimId.toUpperCase() === 'DATE';
      return {
        id:          dimId,
        description: desc,
        type:        dimType || 'GENERIC',
        isMeasure:   MEASURE_TYPES.has(dimType),
        isAuto,
      };
    }).filter(d => d.id);
  }

  _parseDimensionsFromXml(xml) {
    const dims = [];
    const seen = new Set();
    const propRx = /<Property[^>]+Name="([^"]+)"[^>]*Type="([^"]+)"[^>]*>/g;
    let m;
    while ((m = propRx.exec(xml)) !== null) {
      const id = m[1];
      // Skip: system prefix, sub-attributes (contain __), time/version sub-properties
      if (id.startsWith('_')) continue;
      if (id.includes('__')) continue;
      if (seen.has(id.toUpperCase())) continue;
      seen.add(id.toUpperCase());
      const typeRaw = m[2].replace('Edm.', '').toUpperCase();
      const isAuto    = id.toUpperCase() === 'VERSION' || id.toUpperCase() === 'DATE';
      const isMeasure = typeRaw === 'DECIMAL' || typeRaw === 'DOUBLE' || typeRaw === 'INT32' || typeRaw === 'INT64';
      dims.push({ id, description: id, type: typeRaw, isMeasure, isAuto });
    }
    return dims;
  }

  _parseDimensions(data) {
    // SAC DIS returns columns as "columns" array with "columnID" field (CSRF endpoint),
    // or as "dimensions" array with "id" field (metadata endpoint).
    const raw = Array.isArray(data)
      ? data
      : (data?.columns || data?.dimensions || data?.d?.dimensions || data?.value || data?.d?.results || []);
    const AUTO_TYPES    = new Set(['VERSION', 'DATE', 'FISCAL_YEAR', 'FISCAL_PERIOD']);
    const MEASURE_TYPES = new Set(['MEASURE', 'KPI', 'AMOUNT', 'QUANTITY', 'SIGNEDDATA']);
    return raw.map(d => {
      const typeRaw = (d.type || d.columnType || d.dimensionType || d.DimensionType || '').toUpperCase();
      const colId   = d.columnID || d.id || d.dimensionId || d.DimensionId || d.name || d.Name || '';
      return {
        id:          colId,
        description: d.description || d.columnDescription || d.Description || d.label || d.name || colId,
        type:        typeRaw || 'GENERIC',
        isMeasure:   MEASURE_TYPES.has(typeRaw),
        isAuto:      AUTO_TYPES.has(typeRaw),
      };
    }).filter(d => d.id);
  }

  _autoMapColumns() {
    const ALIASES    = { 'ID': 'CAPEX_ITEM', 'H1': 'AMOUNT', 'IMPORTE': 'AMOUNT', 'MONTO': 'AMOUNT' };
    const EXCL_NAMES = new Set(['DESCRIPTION', 'DESCRIPCION', 'DESC', 'NOTES', 'NOTA', 'COMMENT']);
    const mapping  = {};
    const excluded = [];
    for (const col of (this._requiredColumns || [])) {
      const up = col.toUpperCase();
      if (EXCL_NAMES.has(up)) { excluded.push(col); continue; }
      const alias = ALIASES[up];
      if (alias) { mapping[col] = alias; continue; }
    }
    this._columnMapping   = mapping;
    this._disExcludedCols = excluded;
    this._storageSet('sac_column_mapping', mapping);
    this._storageSet('sac_dis_excluded_cols', excluded);
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
    sub.innerHTML = 'Multi-Sheet Consolidation · Transform &amp; Convert · Alert Semaphores · v' + _EPM_VERSION + badge;
  }

  // ── Fetch list of all models from SAC DIS ─────────────────────────────────
  async _fetchSacModels() {
    this._sacModelsLoading  = true;
    this._showModelPicker   = true;
    this._modelPickerFilter = '';
    this._sacModels         = [];
    this._render();
    try {
      const base = (this._sacTenantUrl || window.location.origin).replace(/\/+$/, '');
      const resp = await fetch(base + '/api/v1/dataimport/models', {
        method: 'GET', credentials: 'include',
        headers: { 'Accept': 'application/json', 'x-sap-dis-request-source': 'fileUploadWidget' }
      });
      if (!resp.ok) throw new Error('HTTP ' + resp.status);
      const data = await resp.json().catch(() => null);
      const raw  = Array.isArray(data) ? data : (data?.d?.results || data?.value || data?.models || []);
      this._sacModels = raw.map(m => ({
        id:          m.modelID          || m.id      || m.modelId    || '',
        name:        m.modelName        || m.name    || '',
        description: m.modelDescription || m.description || m.label || '',
        importType:  m.importType       || ''
      })).filter(m => m.id);
    } catch (err) {
      this._sacModels = [];
      this._notify('Error listing SAC models: ' + err.message, 'error');
    }
    this._sacModelsLoading = false;
    this._render();
  }

  // Resolves the SAC technical name (e.g. PLAN_CAPEX_PLANNING) to the DIS internal GUID
  async _resolveDISModelId() {
    if (this._sacDISModelId) return this._sacDISModelId;
    try {
      const base = (this._sacTenantUrl || window.location.origin).replace(/\/+$/, '');
      const resp = await fetch(base + '/api/v1/dataimport/models', {
        credentials: 'include',
        headers: { 'Accept': 'application/json', 'x-sap-dis-request-source': 'fileUploadWidget' }
      });
      if (!resp.ok) return this._sacModelId;
      const data = await resp.json().catch(() => []);
      const raw  = Array.isArray(data) ? data : (data?.d?.results || data?.value || []);
      const match = raw.find(m =>
        m.modelName === this._sacModelId ||
        m.modelID   === this._sacModelId ||
        (m.modelDescription || '').includes(this._sacModelId)
      );
      if (match?.modelID) {
        this._sacDISModelId = match.modelID;
        this._storageSet('sac_dis_guid', match.modelID);
        console.log('[DIS] resolved', this._sacModelId, '→', match.modelID);
        // Also populate MODEL NAME from the matched model description
        const desc = match.modelDescription || match.description || match.label || match.modelName || '';
        if (desc && desc !== this._sacModelId) {
          this._sacModelDesc = desc;
          this._storageSet('sac_dis_model_desc', desc);
          const descEl = this._shadow && this._shadow.querySelector('#sacModelDescDisplay');
          if (descEl) { descEl.value = desc; descEl.style.color = '#1D2D3E'; }
        }
        return match.modelID;
      }
    } catch (_) {}
    console.warn('[DIS] GUID not resolved — using technical name as fallback:', this._sacModelId, '(may cause [1003])');
    return this._sacModelId; // fallback: use as-is
  }

  // SAC DIS REST API — 3-step flow:
  // 1. CSRF  : GET /models/{GUID}/{importType} with x-csrf-token: fetch
  // 2. Create + Upload: POST /models/{GUID}/{importType} with CSV body → { jobID }
  // 3. Run   : POST /jobs/{jobId}/run
  // 4. Poll  : GET /jobs/{jobId} until successful
  async _fetchDeltaPreview() {
    if (!this._sacModelId || !this._targetVersion || !this._targetDate) {
      this._deltaPreview = { rows: [], cols: [], loading: false, error: 'Set VERSION, PERIOD and MODEL ID first.', loaded: false };
      this._render(); return;
    }
    this._deltaPreview = { rows: [], cols: [], loading: true, error: null, loaded: false };
    this._render();
    const base = (this._sacTenantUrl || window.location.origin).replace(/\/+$/, '');
    const DIS  = base + '/api/v1/dataimport';
    const DIS_HEADER = { 'x-sap-sac-custom-auth': 'true' };
    let disModelId = this._sacDISModelId || '';
    if (!disModelId) { try { disModelId = await this._resolveDISModelId(); } catch (_) { disModelId = this._sacModelId; } }

    const vFilter = encodeURIComponent(this._targetVersion);
    const dFilter = encodeURIComponent(this._targetDate);
    const importType = this._sacImportType || 'privateFactData';

    const endpoints = [
      DIS + '/models/' + encodeURIComponent(disModelId) + '/' + importType + '?Version=' + vFilter + '&Date=' + dFilter + '&$format=json&$top=200',
      DIS + '/models/' + encodeURIComponent(disModelId) + '/data?$filter=Version%20eq%20%27' + vFilter + '%27%20and%20Date%20eq%20%27' + dFilter + '%27&$top=200&$format=json',
      DIS + '/models/' + encodeURIComponent(this._sacModelId) + '/' + importType + '?Version=' + vFilter + '&Date=' + dFilter + '&$format=json&$top=200',
    ];

    let rows = [], cols = [], error = null, loaded = false;
    for (const url of endpoints) {
      try {
        const resp = await fetch(url, { method: 'GET', credentials: 'include', headers: { 'Accept': 'application/json', ...DIS_HEADER } });
        if (!resp.ok) { console.log('[DeltaPreview] GET', url.slice(-60), 'HTTP', resp.status); continue; }
        const body = await resp.json().catch(() => null);
        if (!body) continue;
        const arr = body.value || body.data || body.rows || (Array.isArray(body) ? body : null);
        if (arr && arr.length > 0) {
          rows = arr;
          cols = Object.keys(arr[0]).filter(k => !k.startsWith('@') && !k.startsWith('__'));
          loaded = true;
          console.log('[DeltaPreview] found', rows.length, 'rows via', url.slice(-60));
          break;
        }
        // empty array → no data in region
        if (arr && arr.length === 0) { loaded = true; rows = []; break; }
      } catch (e) { console.warn('[DeltaPreview] endpoint failed:', e.message); }
    }
    if (!loaded) error = 'Preview not available on this DIS endpoint. Check data in SAC Modeler → ' + (this._sacModelId || 'model') + '.';
    this._deltaPreview = { rows, cols, loading: false, error, loaded };
    this._render();
  }

  async _fetchExistingDataOData(version, date) {
    if (!version && !date) return null;
    const base    = (this._sacTenantUrl || window.location.origin).replace(/\/+$/, '');
    const modelId = this._sacModelId;
    const guid    = this._sacDISModelId || modelId;
    const filterParts = [];
    if (version) filterParts.push("Version eq '" + version + "'");
    if (date)    filterParts.push("Date eq '" + date + "'");
    const filter = filterParts.join(' and ');
    const guidWithPrefix = guid.startsWith('t.5.') ? guid : ('t.5.' + guid);
    const endpoints = [
      base + '/api/v1/dataexport/providers/sac/' + encodeURIComponent(modelId) + '?$format=json&$filter=' + encodeURIComponent(filter) + '&$top=50000',
      base + '/api/v1/dataexport/providers/sac/' + encodeURIComponent(guid) + '?$format=json&$filter=' + encodeURIComponent(filter) + '&$top=50000',
      base + '/api/v1/dataexport/providers/sac/' + encodeURIComponent(guidWithPrefix) + '?$format=json&$filter=' + encodeURIComponent(filter) + '&$top=50000',
    ];
    for (const url of endpoints) {
      try {
        const r = await fetch(url, { credentials: 'include', headers: { 'Accept': 'application/json' } });
        const shortUrl = url.replace(/.*providers\/sac\//, '').split('?')[0];
        console.log('[DIS] existing data fetch:', r.status, shortUrl);
        if (!r.ok) continue;
        const data = await r.json().catch(() => null);
        if (!data) continue;
        const rows = data?.value || data?.d?.results || (Array.isArray(data) ? data : null);
        if (rows) {
          const filtered = rows.filter(row => {
            const rv = row?.Version || row?.version || '';
            const rd = row?.Date    || row?.date    || '';
            return (!version || rv === version) && (!date || rd === date);
          });
          console.log('[DIS] existing data: found', filtered.length, 'rows for Version=' + (version || '*') + ' Date=' + (date || '*'));
          return filtered;
        }
      } catch (_) {}
    }
    return null; // null = cannot fetch (vs [] = no existing data)
  }

  _mergePlanningData(uploadedRows, existingRows, mode, dimensions) {
    const measureSet = new Set((dimensions || []).filter(d => d.isMeasure).map(d => d.id.toUpperCase()));
    const autoSet    = new Set(['VERSION', 'DATE']);
    const makeKey = row => Object.keys(row)
      .filter(k => !measureSet.has(k.toUpperCase()) && !autoSet.has(k.toUpperCase()))
      .sort()
      .map(k => k + '\x00' + (row[k] ?? ''))
      .join('\x01');
    const uploadedMap = new Map(uploadedRows.map(r => [makeKey(r), r]));
    if (mode === 'update') {
      const kept = existingRows
        .filter(r => !uploadedMap.has(makeKey(r)))
        .map(r => Object.fromEntries(Object.entries(r).map(([k, v]) => [k, String(v ?? '')])));
      return [...kept, ...uploadedRows];
    }
    if (mode === 'append') {
      const existingMap = new Map(
        existingRows.map(r => [makeKey(r), Object.fromEntries(Object.entries(r).map(([k, v]) => [k, String(v ?? '')]))])
      );
      for (const uploaded of uploadedRows) {
        const key = makeKey(uploaded);
        if (existingMap.has(key)) {
          const existing = existingMap.get(key);
          for (const [col, val] of Object.entries(uploaded)) {
            if (measureSet.has(col.toUpperCase())) {
              existing[col] = String((parseFloat(existing[col] || 0) || 0) + (parseFloat(val || 0) || 0));
            }
          }
        } else {
          existingMap.set(key, { ...uploaded });
        }
      }
      return [...existingMap.values()];
    }
    return uploadedRows;
  }

  async _callDISApi(rows) {
    const base       = (this._sacTenantUrl || window.location.origin).replace(/\/+$/, '');
    const DIS        = base + '/api/v1/dataimport';
    const DIS_HEADER = { 'x-sap-dis-request-source': 'fileUploadWidget' };
    const _setLoading = (txt) => { this._loadingText = txt; this._render(); };

    // ── 0. Resolve DIS internal GUID ─────────────────────────────────────────
    _setLoading('🔍 Resolving model ID in SAC DIS…');
    const disModelId = await this._resolveDISModelId();
    console.log('[DIS] modelId:', disModelId, '| technical:', this._sacModelId, '| cached GUID:', this._sacDISModelId || '(none)');
    if (this._loadStrategy === 'DELTA') {
      console.log('[DIS] Load strategy: DELTA — DIS Update mode: updates matching dimension-key records, inserts new ones, preserves all other existing data');
    }

    // ── 1. Build SAC-name rows (enables client-side merge), then CSV ──────────
    const disExcluded = new Set(this._disExcludedCols || ['DESCRIPTION']);
    const cols = [...this._requiredColumns].filter(c => !disExcluded.has(c));
    if (this._targetVersion && !cols.includes('Version')) cols.unshift('Version');
    if (this._targetDate    && !cols.includes('Date'))    cols.unshift('Date');
    const mapping = this._columnMapping || {};
    const revMap = {};
    Object.entries(mapping).forEach(([excelCol, sacName]) => { revMap[sacName] = excelCol; });
    const sacCols = cols.map(c => mapping[c] || c);

    // Build rows as SAC-name objects first
    let sacRows = rows.map(r => {
      const obj = {};
      cols.forEach((col, i) => {
        let v;
        if      (col === 'Version') v = this._targetVersion || '';
        else if (col === 'Date')    v = this._targetDate ? String(this._targetDate) : '';
        else                        v = r[col] ?? r[revMap[col]] ?? '';
        obj[sacCols[i]] = String(v).trim();
      });
      return obj;
    });

    // ── 1b. For Update/Append: fetch existing data and merge client-side ───────
    // DIS v1 ignores writeBackMode/importMode/writeBack JSON body params — always Replace.
    // Client-side merge: fetch existing → merge in memory → Replace with merged dataset.
    const isNonReplace = this._importMethod === 'update' || this._importMethod === 'append';
    this._mergeWarningMsg = '';
    if (isNonReplace) {
      _setLoading('🔍 Reading existing data for ' + (this._importMethod === 'update' ? 'Update' : 'Append') + ' merge…');
      const existing = await this._fetchExistingDataOData(this._targetVersion, this._targetDate);
      if (existing && existing.length > 0) {
        sacRows = this._mergePlanningData(sacRows, existing, this._importMethod, this._sacModelDimensions || []);
        console.log('[DIS] client-side', this._importMethod, 'merge:', existing.length, 'existing +', rows.length, 'uploaded →', sacRows.length, 'rows in merged CSV');
      } else if (existing === null) {
        console.warn('[DIS] Cannot read existing data — Upload will Replace. Existing rows for Version=' + (this._targetVersion || '—') + ' Date=' + (this._targetDate || '—') + ' may be overwritten.');
        this._mergeWarningMsg = 'Could not read existing data to merge — existing rows for this Version+Period may have been replaced.';
      } else {
        console.log('[DIS] No existing data for Version=' + (this._targetVersion || '—') + ', Date=' + (this._targetDate || '—') + ' — uploading as new records');
      }
    }

    // Build CSV from (merged) sacRows
    const sacColsFinal = sacRows.length > 0 ? Object.keys(sacRows[0]) : sacCols;
    const csvLines = [sacColsFinal.join(',')];
    for (const r of sacRows) {
      const vals = sacColsFinal.map(col => {
        let v = String(r[col] ?? '').trim();
        if (v.includes(',') || v.includes('"') || v.includes('\n')) {
          v = '"' + v.replace(/"/g, '""') + '"';
        }
        return v;
      });
      csvLines.push(vals.join(','));
    }
    const csvBody = csvLines.join('\n');
    console.log('[DIS] CSV header:', csvLines[0]);
    console.log('[DIS] CSV rows:', csvLines.length - 1, '| first row:', csvLines[1] || '(empty)');

    // ── 2. CSRF token ─────────────────────────────────────────────────────────
    _setLoading('🔑 Fetching CSRF token from SAC DIS…');
    const importType = this._sacImportType || 'privateFactData';
    const csrfResp = await fetch(DIS + '/models/' + encodeURIComponent(disModelId) + '/' + importType, {
      method: 'GET', credentials: 'include',
      headers: { 'x-csrf-token': 'fetch', 'Accept': 'application/json', ...DIS_HEADER }
    });
    const csrfToken = csrfResp.headers.get('x-csrf-token') || csrfResp.headers.get('X-Csrf-Token') || '';
    console.log('[DIS] CSRF:', csrfToken ? '✅ obtained' : '❌ empty (HTTP ' + csrfResp.status + ')');

    // ── 3. Create job + upload data (single-step CSV POST) ───────────────────
    // Write mode (Update/Append) is handled client-side in step 1b — DIS v1 always Replaces.
    _setLoading('📤 Sending ' + sacRows.length + ' rows to SAC DIS…');
    const importTypesToTry = [importType, importType === 'privateFactData' ? 'factData' : 'privateFactData'];
    let jobCreateResp, usedImportType;

    for (const iType of importTypesToTry) {
      const r = await fetch(DIS + '/models/' + encodeURIComponent(disModelId) + '/' + iType, {
        method: 'POST', credentials: 'include',
        headers: { 'Content-Type': 'text/csv', 'x-csrf-token': csrfToken, ...DIS_HEADER },
        body: csvBody
      });
      console.log('[DIS] job create (CSV' + (isNonReplace ? ', client-merged ' + this._importMethod : '') + ')', iType, 'HTTP', r.status);
      if (r.ok) { jobCreateResp = r; usedImportType = iType; break; }
      let errBody = {};
      try { errBody = await r.clone().json(); } catch (_) {}
      const errMsg = JSON.stringify(errBody);
      console.log('[DIS] POST', iType, 'HTTP', r.status, '| error body:', errMsg.slice(0, 400));
      if (r.status !== 404 && !errMsg.includes('1704-1801-3304')) {
        const msg = errBody?.error?.message || errBody?.message || errBody?.errorDetails?.[0]?.message || 'Upload+Create HTTP ' + r.status;
        throw new Error(msg);
      }
      jobCreateResp = r;
    }
    if (!jobCreateResp?.ok) {
      let msg = 'Upload+Create HTTP ' + (jobCreateResp?.status || '?');
      try { const b = await jobCreateResp.json(); msg = b?.error?.message || b?.message || msg; } catch (_) {}
      throw new Error(msg + ' — Check the model technical name in SAC Modeler');
    }
    if (usedImportType && usedImportType !== this._sacImportType) {
      this._sacImportType = usedImportType;
      this._storageSet('sac_import_type', usedImportType);
      console.log('[DIS] auto-detected importType:', usedImportType);
    }

    const jobCreateBody = await jobCreateResp.json().catch(() => ({}));
    const jobId  = jobCreateBody.jobID || jobCreateBody.jobId || jobCreateBody.id;
    const jobUrl = jobCreateBody.jobURL || jobCreateBody.jobUrl || '';
    if (!jobId) throw new Error('SAC DIS did not return a jobID: ' + JSON.stringify(jobCreateBody));
    console.log('[DIS] jobId:', jobId, '| jobURL:', jobUrl || '(none)', '| importType:', usedImportType);

    // ── 4. Upload CSV data (two-step tenant: job created empty) ──────────────
    // Check initial job status; if READY_FOR_DATA the tenant requires a separate data upload.
    // Try POST /jobs/{jobId} (SAP DIS SDK canonical) then /jobs/{jobId}/data as fallback.
    _setLoading('📤 Uploading data to DIS job…');
    const jobCheckResp = await fetch(DIS + '/jobs/' + jobId, {
      method: 'GET', credentials: 'include',
      headers: { 'Accept': 'application/json', ...DIS_HEADER }
    });
    const jobCheckBody  = jobCheckResp.ok ? await jobCheckResp.json().catch(() => ({})) : {};
    const jobInitStatus = ((jobCheckBody.d || jobCheckBody).jobStatus || (jobCheckBody.d || jobCheckBody).status || '').toUpperCase();
    console.log('[DIS] job initial status:', jobInitStatus || '(unknown)');

    if (jobInitStatus === 'READY_FOR_DATA' || !jobInitStatus) {
      let dataUploaded = false;
      // Try SAP-provided jobURL first, then canonical /jobs/{id}, then /jobs/{id}/data
      const dataUrls = [jobUrl, DIS + '/jobs/' + jobId, DIS + '/jobs/' + jobId + '/data'].filter(Boolean);
      for (const dataUrl of dataUrls) {
        const dataResp = await fetch(dataUrl, {
          method: 'POST', credentials: 'include',
          headers: {
            'Content-Type': 'text/csv',
            'x-csrf-token': csrfToken,
            ...DIS_HEADER
          },
          body: csvBody
        });
        const label = dataUrl === jobUrl ? 'jobURL' : (dataUrl.endsWith('/data') ? 'jobId/data' : 'jobId');
        console.log('[DIS] data upload', label, 'HTTP:', dataResp.status);
        if (dataResp.ok) { dataUploaded = true; break; }
        if (dataResp.status === 405 || dataResp.status === 404) continue;
        let msg = 'Data upload HTTP ' + dataResp.status;
        try { const b = await dataResp.json(); msg = b?.error?.message || b?.message || msg; } catch (_) {}
        throw new Error(msg);
      }
      if (!dataUploaded) console.warn('[DIS] data upload: no endpoint accepted the CSV (', dataUrls.join(', '), ')');
    }

    // ── 5. Run the job ────────────────────────────────────────────────────────
    _setLoading('🚀 Applying data to model ' + this._sacModelId + '…');
    const runResp = await fetch(DIS + '/jobs/' + jobId + '/run', {
      method: 'POST', credentials: 'include',
      headers: { 'x-csrf-token': csrfToken, ...DIS_HEADER }
    });
    console.log('[DIS] run HTTP:', runResp.status);
    if (!runResp.ok && runResp.status !== 405 && runResp.status !== 412) {
      let msg = 'Run HTTP ' + runResp.status;
      try { const b = await runResp.json(); msg = b?.error?.message || b?.message || msg; } catch (_) {}
      throw new Error(msg);
    }
    if (runResp.status === 405 || runResp.status === 412) {
      console.log('[DIS] run endpoint returned', runResp.status, '— proceeding to poll');
    }

    // ── 6. Poll GET /jobs/{jobId} until successful ────────────────────────────
    _setLoading('⏳ Waiting for SAC DIS confirmation…');
    const pollResult = await this._pollJobStatus(
      DIS + '/jobs/' + jobId,
      ['SUCCESSFUL', 'SUCCESS', 'COMPLETED', 'COMPLETED_WITH_FAILURES', 'COMPLETE'],
      ['FAILED', 'UPLOADING_FAILED', 'UPLOADINGFAILED'],
      120, 3000, _setLoading, DIS_HEADER
    );

    // ── 7. Fetch invalid records to detect silently-rejected rows ─────────────
    try {
      const invResp = await fetch(DIS + '/jobs/' + jobId + '/invalidrecords', {
        method: 'GET', credentials: 'include',
        headers: { 'Accept': 'application/json', ...DIS_HEADER }
      });
      if (invResp.ok) {
        const invBody = await invResp.json().catch(() => ({}));
        const records = invBody.value || invBody.invalidRecords || invBody.records || (Array.isArray(invBody) ? invBody : []);
        if (records.length > 0) {
          console.warn('[DIS] invalidrecords (' + records.length + '):', JSON.stringify(records).slice(0, 300));
          pollResult._invalidRecords = records;
          pollResult._invalidCount   = records.length;
        }
      }
    } catch (_) {}

    return pollResult;
  }

  // Polls a DIS /status URL until one of the target statuses is reached.
  async _pollJobStatus(statusUrl, targetStatuses, failStatuses, maxAttempts, intervalMs, _setLoading, extraHeaders) {
    for (let i = 0; i < maxAttempts; i++) {
      await new Promise(r => setTimeout(r, intervalMs));
      const r = await fetch(statusUrl, {
        method: 'GET', credentials: 'include',
        headers: { 'Accept': 'application/json', ...(extraHeaders || {}) }
      });
      if (!r.ok) { console.warn('[DIS poll] HTTP ' + r.status + ' on attempt ' + (i+1)); continue; }
      const body   = await r.json().catch(() => ({}));
      // SAC DIS may return OData format { d: { status: ... } } or flat { status: ... }
      const inner  = body.d || body;
      const status = (inner.jobStatus || inner.status || body.jobStatus || body.status || '').toUpperCase();
      console.log('[DIS poll] attempt', i+1, '— status:', status || '(empty)', JSON.stringify(inner).slice(0, 120));
      if (targetStatuses.some(s => status === s || status.startsWith(s))) return body;
      if ((failStatuses || []).some(s => status.startsWith(s.toUpperCase()) || s.toUpperCase().startsWith(status))) {
        throw new Error('SAC DIS job ' + status + ': ' + (inner.message || inner.errorMessage || body.message || JSON.stringify(inner).slice(0, 200)));
      }
      if (status === 'READY_FOR_DATA' && i >= 4) {
        throw new Error('SAC DIS: job did not receive data (still READY_FOR_DATA after ' + (i+1) + ' attempts). Check DIS permissions or model technical name.');
      }
      if (_setLoading) _setLoading('⏳ SAC DIS: ' + (status || 'processing') + '…');
    }
    throw new Error('Timeout waiting for SAC DIS (URL: ' + statusUrl + ')');
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
      badge.innerHTML = '⚠️ Set Version and Period before synchronizing';
      badge.style.cssText = 'background:#FFF8E6;color:#B45309;border:1px solid #FCD34D;padding:5px 14px;border-radius:20px;font-size:11px;font-weight:700;';
      return;
    }
    const base = '✅ ' + this._esc(this._targetVersion) + ' · ' + this._esc(this._targetDate);
    if (this._sacModelId) {
      badge.innerHTML = base + ' → <strong style="color:#5BBFFF;">' + this._esc(this._sacModelId) + '</strong>';
      badge.style.cssText = 'background:#E6F9EC;color:#107E3E;border:1px solid #A3E6BF;padding:5px 14px;border-radius:20px;font-size:11px;font-weight:700;';
    } else {
      badge.innerHTML = base + ' <span style="color:#B45309;font-weight:600;">(no Model ID — event only)</span>';
      badge.style.cssText = 'background:#E6F9EC;color:#107E3E;border:1px solid #A3E6BF;padding:5px 14px;border-radius:20px;font-size:11px;font-weight:700;';
    }
  }

  _bindEvents() {
    const shadow = this._shadow;

    const toggleHelpBtn = shadow.getElementById('toggleHelpBtn');
    const closeHelpModalBtn = shadow.getElementById('closeHelpModalBtn');
    if (toggleHelpBtn) toggleHelpBtn.addEventListener('click', () => { this._showHelpModal = !this._showHelpModal; this._showChangelog = false; this._render(); });
    if (closeHelpModalBtn) closeHelpModalBtn.addEventListener('click', () => { this._showHelpModal = false; this._render(); });

    const toggleChangelogBtn = shadow.getElementById('toggleChangelogBtn');
    const closeChangelogBtn = shadow.getElementById('closeChangelogBtn');
    if (toggleChangelogBtn) toggleChangelogBtn.addEventListener('click', () => { this._showChangelog = !this._showChangelog; this._showHelpModal = false; this._render(); });
    if (closeChangelogBtn) closeChangelogBtn.addEventListener('click', () => { this._showChangelog = false; this._render(); });

    shadow.querySelectorAll('.strategy-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const strat = e.currentTarget.dataset.strategy;
        if (strat && this._loadStrategy !== strat) {
          const prev = this._loadStrategy;
          this._loadStrategy = strat;
          if (strat === 'DELTA' && this._importMethod === 'replace') {
            this._importMethod = 'update';
            this._storageSet('sac_import_method', 'update');
          } else if (strat === 'FULL' && !this._importMethod) {
            this._importMethod = 'replace';
            this._storageSet('sac_import_method', 'replace');
          }
          this._deltaPreview = { rows: [], cols: [], loading: false, error: null, loaded: false };
          this._logAction('STRATEGY_CHANGE', 'Load Strategy', 'Planning mode switched from ' + prev + ' to ' + strat, prev, strat);
          this._render();
        }
      });
    });

    // ── SAC Artifact sync is always active

    shadow.querySelectorAll('.method-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const m = e.currentTarget.dataset.method;
        if (m && this._importMethod !== m) {
          this._importMethod = m;
          this._storageSet('sac_import_method', m);
          this._deltaPreview = { rows: [], cols: [], loading: false, error: null, loaded: false };
          this._render();
        }
      });
    });

    const loadDeltaPreviewBtn = shadow.getElementById('loadDeltaPreviewBtn');
    if (loadDeltaPreviewBtn) loadDeltaPreviewBtn.addEventListener('click', () => { this._fetchDeltaPreview(); });

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

    const sacModelDescInp = shadow.getElementById('sacModelDescDisplay');
    if (sacModelDescInp) {
      sacModelDescInp.addEventListener('change', (e) => {
        const v = e.target.value.trim();
        this._sacModelDesc = v;
        this._storageSet('sac_dis_model_desc', v || null);
      });
    }

    const browseModelsBtn = shadow.getElementById('browseModelsBtn');
    if (browseModelsBtn) {
      browseModelsBtn.addEventListener('click', () => this._fetchSacModels());
    }

    const sacDimPanelHeader = shadow.getElementById('sacDimPanelHeader');
    if (sacDimPanelHeader) {
      sacDimPanelHeader.addEventListener('click', () => {
        this._sacDimPanelOpen = !this._sacDimPanelOpen;
        this._render();
      });
    }

    const autoMapAllBtn = shadow.getElementById('autoMapAllBtn');
    if (autoMapAllBtn) {
      autoMapAllBtn.addEventListener('click', () => {
        this._autoMapColumns();
        const mappedCnt  = Object.keys(this._columnMapping || {}).length;
        const exclCnt    = (this._disExcludedCols || []).length;
        this._notify('⚡ Default mapping applied — ' + mappedCnt + ' mapped, ' + exclCnt + ' excluded', 'info');
      });
    }

    const applyMappingBtn = shadow.getElementById('applyMappingBtn');
    if (applyMappingBtn) {
      applyMappingBtn.addEventListener('click', () => {
        const inputs = [...shadow.querySelectorAll('[data-mapping-col]')];
        const newMapping = {};
        for (const inp of inputs) {
          const col = inp.dataset.mappingCol;
          const val = inp.value.trim();
          if (val) newMapping[col] = val;
        }
        this._columnMapping = newMapping;
        this._storageSet('sac_column_mapping', newMapping);
        const mappedCnt = Object.keys(newMapping).length;
        const totalCols = (this._requiredColumns || []).length;
        const exclCnt   = (this._disExcludedCols || []).length;
        const msg = totalCols === 0
          ? '💾 Mapping saved — load a file to see columns'
          : '✅ Mapping saved — ' + mappedCnt + '/' + (totalCols - exclCnt) + ' columns mapped' + (exclCnt > 0 ? ', ' + exclCnt + ' excluded' : '') + ' — you can proceed!';
        this._notify(msg, mappedCnt > 0 || exclCnt > 0 ? 'success' : 'info');
      });
    }

    shadow.querySelectorAll('[data-clear-col]').forEach(btn => {
      btn.addEventListener('click', () => {
        const col = btn.dataset.clearCol;
        const m = { ...(this._columnMapping || {}) };
        delete m[col];
        const ex = (this._disExcludedCols || []).filter(c => c !== col);
        this._columnMapping   = m;
        this._disExcludedCols = ex;
        this._storageSet('sac_column_mapping', m);
        this._storageSet('sac_dis_excluded_cols', ex);
        this._render();
      });
    });

    shadow.querySelectorAll('[data-toggle-excl-col]').forEach(btn => {
      btn.addEventListener('click', () => {
        const col = btn.dataset.toggleExclCol;
        const ex  = new Set(this._disExcludedCols || []);
        if (ex.has(col)) ex.delete(col); else ex.add(col);
        this._disExcludedCols = [...ex];
        this._storageSet('sac_dis_excluded_cols', [...ex]);
        this._render();
      });
    });

    const closeModelPickerBtn = shadow.getElementById('closeModelPickerBtn');
    if (closeModelPickerBtn) {
      closeModelPickerBtn.addEventListener('click', () => { this._showModelPicker = false; this._render(); });
    }

    shadow.querySelectorAll('.sac-model-pick-row').forEach(row => {
      row.addEventListener('mouseenter', () => { row.style.background = '#EBF5FF'; });
      row.addEventListener('mouseleave', () => { row.style.background = ''; });
      row.addEventListener('click', () => {
        const disGuid    = row.dataset.modelId;    // internal GUID
        const modelName  = row.dataset.modelName;  // technical name (preferred)
        const modelDesc  = row.dataset.modelDesc  || '';
        const importType = row.dataset.importType || 'privateFactData';
        this._sacModelId    = modelName || disGuid;  // store technical name
        this._sacDISModelId = disGuid;               // cache GUID to skip resolve step
        this._sacImportType = importType;
        this._sacModelDesc  = modelDesc;
        this._sacModelName  = '';
        this._sacModelState = 'connecting';
        this._showModelPicker = false;
        this._modelPickerFilter = '';
        this._storageSet('sac_dis_model_id',   modelName || disGuid);
        this._storageSet('sac_dis_guid',        disGuid);
        this._storageSet('sac_import_type',     importType);
        this._storageSet('sac_dis_model_desc',  modelDesc);
        this._fetchModelInfo();
        this._render();
      });
    });

    const modelPickerFilterInp = shadow.getElementById('modelPickerFilterInp');
    if (modelPickerFilterInp) {
      modelPickerFilterInp.addEventListener('input', (e) => {
        this._modelPickerFilter = e.target.value;
        this._render();
        // restore focus after re-render
        const inp = this._shadow.getElementById('modelPickerFilterInp');
        if (inp) { inp.focus(); inp.setSelectionRange(inp.value.length, inp.value.length); }
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

    const publishReadyBtn = shadow.getElementById('publishReadyBtn');
    if (publishReadyBtn) publishReadyBtn.addEventListener('click', () => { this._showPublishModal = true; this._render(); });

    const closePublishModalBtn = shadow.getElementById('closePublishModalBtn');
    if (closePublishModalBtn) closePublishModalBtn.addEventListener('click', () => { this._showPublishModal = false; this._render(); });

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
          this._notify('⚠️ Set the Version and Period (🎯 SAC Load Target) before synchronizing.', 'error');
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
          targetVersion: this._targetVersion || '',
          targetDate: this._targetDate || '',
          importMethod: this._importMethod || 'replace',
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
          this._loadingText = '📤 Starting import to ' + this._sacModelId + '…';
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
              this._history[idx].status   = '❌ DIS ERROR';
              this._history[idx].disStatus = 'ERROR';
              this._history[idx].disError  = err.message;
              this._saveHistory();
            }
            this._notify('❌ SAC DIS Error: ' + err.message, 'error');
            return;
          }

          // Update audit with actual DIS result (rows written vs. rejected)
          const jobStatus  = (disResult?.jobStatus || disResult?.status || '').toUpperCase();
          const written    = disResult?.importedRecords ?? disResult?.writtenRecords ?? null;
          const failed     = (disResult?.failedRecords ?? disResult?.rejectedRecords ?? 0) + (disResult?._invalidCount ?? 0);
          const invRecs    = disResult?._invalidRecords || [];
          const idx = this._history.findIndex(h => h.id === loadId);
          if (idx !== -1) {
            if (failed > 0 || jobStatus === 'COMPLETED_WITH_FAILURES') {
              const writtenTxt = written != null ? written + ' written, ' : '';
              this._history[idx].status    = '⚠️ PARTIAL LOAD — ' + writtenTxt + failed + ' rejected';
              this._history[idx].disStatus = 'PARTIAL';
              if (invRecs.length > 0) {
                // Store first error reason for display in audit detail
                const firstErr = invRecs[0];
                const reason   = firstErr?.errorMessage || firstErr?.message || firstErr?.error || JSON.stringify(firstErr).slice(0, 120);
                this._history[idx].disError = failed + ' row(s) rejected by SAC DIS: ' + reason;
              }
            } else {
              const writtenTxt = written != null ? ' — ' + written + ' rows' : '';
              this._history[idx].status    = '🟢 LOAD COMPLETE' + writtenTxt;
              this._history[idx].disStatus = 'OK';
            }
            this._history[idx].disWrittenRows = written;
            this._history[idx].disFailedRows  = failed > 0 ? failed : null;
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
          const infoRows = finalRec?.disWrittenRows != null ? ' — ' + finalRec.disWrittenRows + ' rows written' : '';
          const deltaNote = this._loadStrategy === 'DELTA' ? ' · Delta: matching records updated, new ones inserted, other existing data preserved' : '';
          this._notify('🟢 LOAD COMPLETE — ' + loadId + ' · ' + this._loadStrategy + infoRows + ' → SAP Analytics Cloud ✅' + deltaNote, 'success');
        } else if (finalRec?.disStatus === 'PARTIAL') {
          const reason = finalRec.disError ? ' · ' + finalRec.disError.slice(0, 120) : '';
          this._notify('⚠️ PARTIAL LOAD — ' + loadId + ': ' + (finalRec.disWrittenRows || '0') + ' written, ' + (finalRec.disFailedRows || '?') + ' rejected. Check the audit log.' + reason, 'warning');
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

  _renderChangelogModal() {
    if (!this._showChangelog) return '';
    const rows = _EPM_CHANGELOG.map(entry => {
      const isLatest = entry.v === _EPM_VERSION;
      return `
        <div style="border:1px solid ${isLatest ? '#107E3E' : '#D9E0E6'};border-radius:8px;padding:12px 14px;background:${isLatest ? '#F5FAF5' : '#FAFBFC'};">
          <div style="display:flex;align-items:center;gap:8px;margin-bottom:6px;">
            <span style="background:${isLatest ? '#107E3E' : '#556B82'};color:#FFF;border-radius:12px;padding:2px 10px;font-size:11px;font-weight:700;">v${entry.v}</span>
            <span style="font-size:12px;color:#8A9BB0;">${entry.date}</span>
            ${isLatest ? '<span style="background:#107E3E;color:#FFF;border-radius:8px;padding:1px 7px;font-size:10px;font-weight:600;">LATEST</span>' : ''}
          </div>
          <div style="font-size:13px;font-weight:600;color:#1D2D3E;margin-bottom:5px;">${this._esc(entry.title)}</div>
          <ul style="margin:0;padding-left:16px;">
            ${entry.items.map(i => `<li style="font-size:12px;color:#3D4F5F;line-height:1.5;">${this._esc(i)}</li>`).join('')}
          </ul>
        </div>`;
    }).join('');
    return `
      <div class="help-modal-overlay">
        <div class="help-modal-card">
          <div style="display:flex;justify-content:space-between;align-items:center;border-bottom:1px solid #D9E0E6;padding-bottom:12px;margin-bottom:14px;">
            <div style="display:flex;align-items:center;gap:10px;">
              <span style="background:#107E3E;color:#FFF;width:28px;height:28px;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;font-weight:800;font-size:14px;">★</span>
              <div>
                <h3 style="margin:0;color:#1D2D3E;font-size:16px;">What's New — Excel Planning Uploader Pro</h3>
                <span style="font-size:11px;color:#556B82;">Version history and improvements</span>
              </div>
            </div>
            <button class="btn" id="closeChangelogBtn" type="button">✕ Close</button>
          </div>
          <div style="display:flex;flex-direction:column;gap:10px;max-height:420px;overflow:auto;">
            ${rows}
          </div>
        </div>
      </div>`;
  }

  _renderPublishModal() {
    if (!this._showPublishModal) return '';
    return `
      <div class="help-modal-overlay">
        <div class="help-modal-card" style="max-width:520px;">
          <div style="display:flex;justify-content:space-between;align-items:center;border-bottom:1px solid #D9E0E6;padding-bottom:12px;margin-bottom:14px;">
            <div style="display:flex;align-items:center;gap:10px;">
              <span style="background:#107E3E;color:#FFF;width:28px;height:28px;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;font-weight:800;font-size:16px;">📢</span>
              <div>
                <h3 style="margin:0;color:#1D2D3E;font-size:16px;">Publish Planning Data in SAC</h3>
                <span style="font-size:11px;color:#556B82;">Make the imported data visible to all story users</span>
              </div>
            </div>
            <button class="btn" id="closePublishModalBtn" type="button">✕ Close</button>
          </div>
          <div style="display:flex;flex-direction:column;gap:14px;">
            <div style="background:#E6F9EC;border:1px solid #A3D9B1;border-radius:8px;padding:12px 16px;">
              <div style="font-size:12px;font-weight:700;color:#107E3E;margin-bottom:4px;">✅ Your data has been loaded successfully</div>
              <div style="font-size:12px;color:#1D6B2E;">The rows are now in the planning model's private workspace. Publish to share them with all users of this story.</div>
            </div>
            <div>
              <div style="font-size:13px;font-weight:700;color:#1D2D3E;margin-bottom:10px;">Steps to publish in SAP Analytics Cloud:</div>
              <ol style="margin:0;padding-left:20px;display:flex;flex-direction:column;gap:10px;">
                <li style="font-size:12px;color:#3D4F5F;line-height:1.6;">In your SAC <strong>Story</strong>, click the <strong>Edit</strong> button in the top toolbar to enter edit mode (if not already editing).</li>
                <li style="font-size:12px;color:#3D4F5F;line-height:1.6;">In the top menu bar, click <strong>Data</strong> → <strong>Publish Data Changes</strong>.<br><span style="color:#556B82;font-size:11px;">Alternatively: click the model name in a planning table → <em>Publish</em>.</span></li>
                <li style="font-size:12px;color:#3D4F5F;line-height:1.6;">SAC will show a confirmation dialog with the planning version and timestamp. Review and click <strong>Publish</strong>.</li>
                <li style="font-size:12px;color:#3D4F5F;line-height:1.6;">The data is now available in the <strong>public</strong> area of the model and visible to all story viewers.</li>
              </ol>
            </div>
            <div style="background:#FFF8E1;border:1px solid #FFD54F;border-radius:8px;padding:12px 16px;">
              <div style="font-size:12px;font-weight:700;color:#7A5A00;margin-bottom:4px;">⚠️ Important notes</div>
              <ul style="margin:0;padding-left:16px;display:flex;flex-direction:column;gap:4px;">
                <li style="font-size:12px;color:#5A4200;line-height:1.5;">You need <em>Create and Edit Stories</em> and <em>Planning</em> permissions to publish.</li>
                <li style="font-size:12px;color:#5A4200;line-height:1.5;">If you loaded with <strong>DELTA</strong> strategy, publishing updates only the matching key records and inserts new ones — non-matching rows in the model are preserved.</li>
                <li style="font-size:12px;color:#5A4200;line-height:1.5;">If the Publish option is greyed out, contact your SAC administrator to verify your permissions on model <strong>PLAN_CAPEX_PLANNING</strong>.</li>
              </ul>
            </div>
          </div>
        </div>
      </div>`;
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

            <div class="metrics-bar" style="grid-template-columns:repeat(7,1fr);">
              <div class="metric-box"><div class="metric-title">🔄 Load Strategy</div><div class="metric-val" style="font-size:13px;color:#0064D9;">${this._esc(rec.loadStrategy || 'FULL')}</div></div>
              <div class="metric-box"><div class="metric-title">🏷️ Version</div><div class="metric-val" style="font-size:13px;color:#1D3A6E;">${this._esc(rec.targetVersion || '—')}</div></div>
              <div class="metric-box"><div class="metric-title">📅 Period</div><div class="metric-val" style="font-size:13px;color:#5B21B6;">${this._esc(rec.targetDate || '—')}</div></div>
              <div class="metric-box"><div class="metric-title">🕒 Timestamp</div><div class="metric-val" style="font-size:13px;">${this._esc(rec.timestamp)}</div></div>
              <div class="metric-box kpi-ok"><div class="metric-title">✅ Committed Rows</div><div class="metric-val" style="color:#107E3E;">${rec.rows}</div></div>
              <div class="metric-box"><div class="metric-title">✎ Rows Modified</div><div class="metric-val" style="color:#0064D9;">${modifiedRowsCount}</div></div>
              <div class="metric-box"><div class="metric-title">🛠️ Tracked Actions</div><div class="metric-val">${changeLog.length} events</div></div>
            </div>

            ${rec.disStatus ? `<div style="margin:8px 0 4px;padding:12px 16px;border-radius:8px;font-size:13px;font-weight:700;border:2px solid ${rec.disStatus === 'OK' ? '#107E3E' : rec.disStatus === 'PARTIAL' ? '#E9730C' : '#BB0000'};background:${rec.disStatus === 'OK' ? '#E6F9EC' : rec.disStatus === 'PARTIAL' ? '#FFF8E1' : '#FFF0F0'};color:${rec.disStatus === 'OK' ? '#107E3E' : rec.disStatus === 'PARTIAL' ? '#7A5A00' : '#BB0000'};">
              <div style="display:flex;align-items:center;gap:10px;">
                ${rec.disStatus === 'OK' ? this._renderSemaphoreDot('GREEN') + ' SAC DIS: LOAD COMPLETE' + (rec.disWrittenRows != null ? ' — <strong>' + rec.disWrittenRows + ' rows written to model</strong>' : '') : ''}
                ${rec.disStatus === 'PARTIAL' ? this._renderSemaphoreDot('AMBER') + ' SAC DIS: PARTIAL LOAD — ' + (rec.disWrittenRows != null ? '<strong>' + rec.disWrittenRows + ' written</strong>, ' : '') + '<strong>' + (rec.disFailedRows || '?') + ' rejected</strong> by model' : ''}
                ${rec.disStatus === 'ERROR' ? this._renderSemaphoreDot('RED') + ' SAC DIS ERROR: ' + this._esc(rec.disError || 'Unknown error') : ''}
                ${rec.disStatus === 'PENDING' ? this._renderSemaphoreDot('BLUE') + ' SAC DIS: Pending confirmation' : ''}
              </div>
              ${(rec.disStatus === 'PARTIAL' || rec.disStatus === 'ERROR') && rec.disError ? `<div style="margin-top:6px;padding:8px 10px;border-radius:6px;background:rgba(0,0,0,0.05);font-size:11px;font-weight:600;font-family:monospace;white-space:pre-wrap;line-height:1.5;">${this._esc(rec.disError)}</div>` : ''}
            </div>` : ''}

            ${rec.disStatus === 'OK' ? `
            <div style="margin:8px 0 4px;padding:14px 18px;border-radius:8px;background:linear-gradient(135deg,#E6F9EC,#EBF7FF);border:2px solid #107E3E;display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap;">
              <div style="display:flex;align-items:center;gap:12px;">
                <span style="font-size:22px;line-height:1;">📋</span>
                <div>
                  <div style="font-size:13px;font-weight:800;color:#107E3E;">Data loaded — review the snapshot below, then publish when ready</div>
                  <div style="font-size:11px;color:#1D6B2E;margin-top:3px;">Once you have verified the rows, publish the planning version in SAP Analytics Cloud to make it visible to all users.</div>
                </div>
              </div>
              <button class="btn" id="publishReadyBtn" type="button" style="background:#107E3E;color:#FFF;padding:8px 16px;font-size:12px;font-weight:700;border-radius:6px;white-space:nowrap;border:none;cursor:pointer;">📢 How to Publish</button>
            </div>
            ` : ''}

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
              <th>🏷️ Version</th>
              <th>📅 Period</th>
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
                <td><span style="background:#F0F4FF;color:#1D3A6E;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600;">${this._esc(h.targetVersion || '—')}</span></td>
                <td><span style="background:#F5F0FF;color:#5B21B6;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600;">${this._esc(h.targetDate || '—')}</span></td>
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
        .toast-info { background:#0070F2; } .toast-success { background:#107E3E; } .toast-error { background:#BB0000; } .toast-warning { background:#E76500; }
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
        ${this._renderChangelogModal()}
        ${this._renderPublishModal()}

        <div class="sap-header">
          <div class="sap-brand">
            <img class="sap-logo-img" src="${this._sapLogoUri}" alt="SAP Logo" />
            <div class="sap-title-wrap">
              <div class="sap-title">Excel Planning Uploader Pro</div>
              <div class="sap-subtitle">Multi-Sheet Consolidation · Transform &amp; Convert · Alert Semaphores · v${_EPM_VERSION}${
                this._sacModelState === 'connecting' ? ' &nbsp;|&nbsp; <span style="color:#A9B8C8;font-weight:600;">⏳ Conectando…</span>' :
                this._sacModelState === 'ok'         ? ' &nbsp;|&nbsp; <span style="color:#5BBFFF;font-weight:700;">📊 ' + this._esc(this._sacModelName || this._sacModelId) + '</span>' :
                this._sacModelState === 'error'      ? ' &nbsp;|&nbsp; <span style="color:#E67E22;font-weight:600;" title="' + this._esc(this._sacModelName) + '">⚠️ ' + this._esc((this._sacModelName || 'Modelo no encontrado').slice(0, 60)) + '</span>' :
                ''
              }</div>
            </div>
          </div>
          <div style="display:flex;align-items:center;gap:12px;">
            <span class="sap-badge-env">☁️ SAP Analytics Cloud</span>
            <button class="help-icon-btn" id="toggleChangelogBtn" type="button" title="What's New — version history">
              <span style="background:#107E3E;color:#FFF;width:18px;height:18px;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;font-size:11px;font-weight:900;">★</span>
              <span>What's New</span>
            </button>
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
                  <span class="strategy-btn-sub">Replaces matching key records</span>
                </span>
              </button>
            </div>
            <div style="display:flex;align-items:center;gap:8px;margin-left:auto;">
              ${this._loadStrategy === 'FULL'
                ? `<span style="background:#EBF5FF;color:#005CB9;border:1px solid #B3D4FF;padding:5px 12px;border-radius:20px;font-size:11px;font-weight:700;">📦 Full · ${this._importMethod === 'update' ? '✏️ Update' : this._importMethod === 'append' ? '➕ Append' : '🗑️ Replace'}</span>`
                : '<span style="background:#FFF3CD;color:#8A5700;border:1px solid #FFCC66;padding:5px 12px;border-radius:20px;font-size:11px;font-weight:700;">⚠️ Delta — updates matching records, inserts new ones</span>'}
            </div>
            ${(() => {
              const METHODS = this._loadStrategy === 'DELTA'
                ? [
                    { id: 'update',  icon: '✏️', label: 'Update',  desc: 'Updates matching dimension-key records, inserts new ones, preserves all other existing data.' },
                    { id: 'append',  icon: '➕', label: 'Append',  desc: 'Accumulates measure values for matching keys (adds to existing amounts), inserts new records.' },
                  ]
                : [
                    { id: 'replace', icon: '🗑️', label: 'Replace', desc: 'Clears all data for Version+Period, then writes the file. Use for complete monthly re-upload.' },
                    { id: 'update',  icon: '✏️', label: 'Update',  desc: 'Updates matching dimension-key records, inserts new ones, preserves all other existing data.' },
                    { id: 'append',  icon: '➕', label: 'Append',  desc: 'Accumulates measure values for matching keys (adds to existing amounts), inserts new records.' },
                  ];
              const active = METHODS.find(m => m.id === this._importMethod) || METHODS[0];
              return `
              <div style="width:100%;padding-top:10px;border-top:1px dashed #D9E0E6;margin-top:4px;">
                <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;">
                  <span style="font-size:11px;font-weight:700;color:#556B82;text-transform:uppercase;letter-spacing:.5px;white-space:nowrap;">⚙️ Import Method</span>
                  <div style="display:flex;gap:6px;">
                    ${METHODS.map(m => `
                      <button class="method-btn" data-method="${m.id}" type="button"
                        style="padding:5px 14px;border-radius:20px;border:2px solid ${this._importMethod === m.id ? '#0064D9' : '#D9E0E6'};background:${this._importMethod === m.id ? '#EBF5FF' : '#FFF'};color:${this._importMethod === m.id ? '#0064D9' : '#556B82'};font-size:11px;font-weight:${this._importMethod === m.id ? '800' : '600'};cursor:pointer;transition:all .15s;">
                        ${m.icon} ${m.label}
                      </button>`).join('')}
                  </div>
                  <span style="font-size:11px;color:#556B82;font-style:italic;min-width:0;">${active.desc}</span>
                </div>
              </div>`;
            })()}
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
                <div style="display:flex;gap:10px;align-items:flex-end;">
                  <div>
                    <label style="font-size:10px;font-weight:700;color:#556B82;display:block;margin-bottom:3px;">MODEL ID ⭐</label>
                    <div style="display:flex;gap:4px;align-items:center;">
                      <input type="text" id="sacModelIdInp" class="input-std" style="width:180px;"
                        placeholder="PLAN_CAPEX_PLANNING" value="${this._esc(this._sacModelId)}" title="SAC planning model technical name — required for DIS write-back" />
                      <button id="browseModelsBtn" class="btn btn-ghost" type="button" style="padding:5px 9px;font-size:11px;white-space:nowrap;" title="Browse available models in SAC">
                        ${this._sacModelsLoading ? '⏳' : '🔍 SAC'}
                      </button>
                    </div>
                  </div>
                  <div>
                    <label style="font-size:10px;font-weight:700;color:#556B82;display:block;margin-bottom:3px;">MODEL NAME</label>
                    <input type="text" id="sacModelDescDisplay"
                      style="width:200px;background:#FFF;color:${this._sacModelDesc ? '#1D2D3E' : '#A9B8C8'};border:1px solid #D9E0E6;border-radius:4px;padding:5px 8px;font-size:12px;font-family:inherit;"
                      value="${this._esc(this._sacModelDesc || '')}"
                      placeholder="Auto-fills on connect…"
                      title="Model friendly name — editable, saved locally" />
                  </div>
                </div>
                ${this._showModelPicker ? (() => {
                  const filterTerm = (this._modelPickerFilter || '').toLowerCase();
                  const matches = (m) =>
                    (m.description || '').toLowerCase().includes(filterTerm) ||
                    (m.name        || '').toLowerCase().includes(filterTerm) ||
                    (m.id          || '').toLowerCase().includes(filterTerm);
                  const filtered = filterTerm ? this._sacModels.filter(matches) : this._sacModels;

                  // group: planning (privateFactData) first, then rest
                  const planning = filtered.filter(m => m.importType === 'privateFactData');
                  const other    = filtered.filter(m => m.importType !== 'privateFactData');

                  const renderRow = (m) => {
                    const isActive = (m.name || m.id) === this._sacModelId;
                    const badge    = m.importType === 'privateFactData'
                      ? '<span style="flex-shrink:0;font-size:10px;background:#E8F5E9;color:#2E7D32;padding:2px 7px;border-radius:10px;font-weight:600;">Planning</span>'
                      : m.importType === 'factData'
                        ? '<span style="flex-shrink:0;font-size:10px;background:#EBF5FF;color:#005CB9;padding:2px 7px;border-radius:10px;font-weight:600;">Actual</span>'
                        : m.importType
                          ? `<span style="flex-shrink:0;font-size:10px;background:#F5F5F5;color:#556B82;padding:2px 7px;border-radius:10px;font-weight:600;">${this._esc(m.importType)}</span>`
                          : '';
                    return `<div class="sac-model-pick-row"
                        data-model-id="${this._esc(m.id)}"
                        data-model-name="${this._esc(m.name || m.id)}"
                        data-model-desc="${this._esc(m.description || '')}"
                        data-import-type="${this._esc(m.importType || 'privateFactData')}"
                        style="padding:7px 12px;cursor:pointer;border-bottom:1px solid #F0F4F8;display:flex;justify-content:space-between;align-items:center;gap:8px;${isActive ? 'background:#EBF5FF;' : ''}">
                        <div style="min-width:0;flex:1;">
                          <div style="font-weight:${isActive ? '800' : '700'};font-size:12px;color:${isActive ? '#003F8A' : '#0064D9'};white-space:nowrap;overflow:hidden;text-overflow:ellipsis;" title="${this._esc(m.name || m.id)}">${isActive ? '✓ ' : ''}${this._esc(m.name || m.id)}</div>
                          ${m.description ? `<div style="font-size:11px;color:#556B82;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;" title="${this._esc(m.description)}">${this._esc(m.description)}</div>` : ''}
                          <div style="font-size:9px;color:#B0BEC5;font-family:monospace;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;" title="${this._esc(m.id)}">${this._esc(m.id)}</div>
                        </div>
                        ${badge}
                      </div>`;
                  };

                  const renderGroup = (title, color, items) => items.length === 0 ? '' : `
                    <div style="padding:4px 12px;font-size:10px;font-weight:700;color:${color};background:#FAFBFC;border-bottom:1px solid #F0F4F8;letter-spacing:.5px;text-transform:uppercase;">${title} (${items.length})</div>
                    ${items.map(renderRow).join('')}`;

                  return `
                  <div id="modelPickerDropdown" style="position:absolute;z-index:9999;top:calc(100% + 2px);left:0;min-width:420px;max-height:360px;display:flex;flex-direction:column;background:#fff;border:1px solid #0070F2;border-radius:8px;box-shadow:0 4px 16px rgba(0,0,0,0.15);">
                    <div style="display:flex;justify-content:space-between;align-items:center;padding:6px 12px;border-bottom:1px solid #E9EEF3;background:#F8FAFC;border-radius:8px 8px 0 0;flex-shrink:0;">
                      <span style="font-weight:700;font-size:11px;color:#1D2D3E;">Modelos SAC DIS — ${filtered.length} de ${this._sacModels.length}</span>
                      <button id="closeModelPickerBtn" type="button" style="background:none;border:none;cursor:pointer;font-size:14px;color:#556B82;padding:0 2px;">✕</button>
                    </div>
                    <div style="padding:6px 10px;border-bottom:1px solid #E9EEF3;flex-shrink:0;">
                      <input type="text" id="modelPickerFilterInp" placeholder="🔍 Search by description, technical name or ID…"
                        value="${this._esc(this._modelPickerFilter)}"
                        style="width:100%;box-sizing:border-box;border:1px solid #C0CDD8;border-radius:4px;padding:5px 8px;font-size:12px;outline:none;font-family:inherit;" />
                    </div>
                    <div style="overflow-y:auto;flex:1;">
                    ${this._sacModelsLoading
                      ? '<div style="padding:20px;text-align:center;color:#556B82;font-size:12px;">⏳ Loading models…</div>'
                      : filtered.length === 0
                        ? `<div style="padding:20px;text-align:center;color:#BB0000;font-size:12px;">${filterTerm ? 'No results for "' + this._esc(filterTerm) + '"' : 'No models available or no DIS permissions'}</div>`
                        : renderGroup('📋 Planning', '#2E7D32', planning) + renderGroup('📊 Actual / Otros', '#005CB9', other)
                    }
                    </div>
                  </div>`;
                })() : ''}
              </div>
            </div>
            ${/* ── Column Mapping Panel ── */ ''}
            ${(() => {
              const mapping  = this._columnMapping || {};
              const excluded = new Set(this._disExcludedCols || []);
              const reqCols  = this._requiredColumns || [];
              if (reqCols.length === 0) return '';

              const mappedCnt  = reqCols.filter(c => !excluded.has(c) && mapping[c]).length;
              const excludedCnt= reqCols.filter(c => excluded.has(c)).length;
              const sameCnt    = reqCols.length - mappedCnt - excludedCnt;

              const hBg    = '#F0F4F8';
              const hLabel = `🔗 Column Mapping — ${mappedCnt > 0 ? mappedCnt + ' renamed · ' : ''}${excludedCnt > 0 ? excludedCnt + ' excluded · ' : ''}${sameCnt} same name`;

              return `<div style="width:100%;margin-top:10px;border:1px solid #D9E0E6;border-radius:6px;overflow:hidden;background:#FAFBFC;">
                <div id="sacDimPanelHeader"
                  style="display:flex;align-items:center;justify-content:space-between;padding:7px 12px;background:${hBg};cursor:pointer;user-select:none;">
                  <span style="font-size:11px;font-weight:700;color:#1D2D3E;">${hLabel}</span>
                  <span style="font-size:11px;color:#556B82;">${this._sacDimPanelOpen ? '▲ collapse' : '▼ expand'}</span>
                </div>
                ${this._sacDimPanelOpen ? `<div>
                  <table style="width:100%;border-collapse:collapse;font-size:12px;">
                    <thead>
                      <tr style="background:#EBF2FA;border-bottom:2px solid #C5D9EB;">
                        <th style="padding:6px 12px;text-align:left;font-weight:700;color:#1D2D3E;font-size:11px;width:40%;">📄 File Column</th>
                        <th style="padding:6px 4px;text-align:center;color:#9CB0C4;font-size:13px;width:5%;">→</th>
                        <th style="padding:6px 8px;text-align:left;font-weight:700;color:#1D2D3E;font-size:11px;">📊 SAC Dimension Name</th>
                        <th style="width:32px;padding:6px 4px;"></th>
                      </tr>
                    </thead>
                    <tbody>
                      ${/* Auto-filled rows: Version and Date from SAC Load Target */ ''}
                      ${(this._targetVersion || this._targetDate) ? `
                        ${this._targetVersion ? `<tr style="border-bottom:1px solid #F0F4F8;background:#F0FAF4;">
                          <td style="padding:5px 12px;">
                            <span style="display:inline-block;background:#D6F5E3;color:#0A5C2B;font-family:monospace;font-weight:700;font-size:11px;padding:2px 8px;border-radius:4px;border:1px solid #A3E6BF;">Version</span>
                          </td>
                          <td style="padding:5px 4px;text-align:center;color:#9CB0C4;font-size:17px;">→</td>
                          <td style="padding:3px 8px;">
                            <span style="display:inline-flex;align-items:center;gap:5px;font-size:11px;color:#0A5C2B;font-weight:600;">
                              <span style="background:#107E3E;color:#fff;border-radius:3px;padding:1px 6px;font-size:10px;font-family:monospace;">${this._esc(this._targetVersion)}</span>
                              <span style="color:#556B82;font-weight:400;">← SAC Load Target</span>
                            </span>
                          </td>
                          <td style="padding:3px 6px;text-align:center;color:#556B82;font-size:12px;" title="Auto-filled from VERSION field">🔒</td>
                        </tr>` : ''}
                        ${this._targetDate ? `<tr style="border-bottom:1px solid #F0F4F8;background:#F0FAF4;">
                          <td style="padding:5px 12px;">
                            <span style="display:inline-block;background:#D6F5E3;color:#0A5C2B;font-family:monospace;font-weight:700;font-size:11px;padding:2px 8px;border-radius:4px;border:1px solid #A3E6BF;">Date</span>
                          </td>
                          <td style="padding:5px 4px;text-align:center;color:#9CB0C4;font-size:17px;">→</td>
                          <td style="padding:3px 8px;">
                            <span style="display:inline-flex;align-items:center;gap:5px;font-size:11px;color:#0A5C2B;font-weight:600;">
                              <span style="background:#107E3E;color:#fff;border-radius:3px;padding:1px 6px;font-size:10px;font-family:monospace;">${this._esc(String(this._targetDate))}</span>
                              <span style="color:#556B82;font-weight:400;">← SAC Load Target</span>
                            </span>
                          </td>
                          <td style="padding:3px 6px;text-align:center;color:#556B82;font-size:12px;" title="Auto-filled from PERIOD field">🔒</td>
                        </tr>` : ''}
                      ` : ''}
                      ${reqCols.map((col, i) => {
                        const isExcl = excluded.has(col);
                        const val    = mapping[col] || '';
                        return `<tr style="border-bottom:1px solid #F0F4F8;background:${i % 2 === 0 ? '#fff' : '#F7FAFD'};">
                          <td style="padding:5px 12px;">
                            <span style="display:inline-block;background:#E8F4FF;color:#003F8A;font-family:monospace;font-weight:700;font-size:11px;padding:2px 8px;border-radius:4px;border:1px solid #B8D8F8;">${this._esc(col)}</span>
                          </td>
                          <td style="padding:5px 4px;text-align:center;color:#9CB0C4;font-size:17px;">→</td>
                          <td style="padding:3px 8px;">
                            ${isExcl
                              ? `<span style="display:inline-block;background:#EBF5FF;color:#005CB9;font-size:11px;font-weight:700;padding:3px 10px;border-radius:4px;border:1px solid #B3D4F0;">⊘ Excluded from DIS</span>`
                              : `<input type="text" data-mapping-col="${this._esc(col)}" value="${this._esc(val)}"
                                  placeholder="(same name as file)"
                                  style="width:100%;border:1px solid #C0CDD8;border-radius:4px;padding:4px 8px;font-size:11px;font-family:monospace;background:#FFF;color:#1D2D3E;outline:none;box-sizing:border-box;" />`
                            }
                          </td>
                          <td style="padding:3px 6px;text-align:center;white-space:nowrap;">
                            <button type="button" data-toggle-excl-col="${this._esc(col)}" title="${isExcl ? 'Include in DIS' : 'Exclude from DIS'}"
                              style="background:none;border:1px solid ${isExcl ? '#B3D4F0' : '#DDE3E9'};border-radius:3px;cursor:pointer;padding:2px 5px;font-size:10px;color:${isExcl ? '#005CB9' : '#556B82'};line-height:1;white-space:nowrap;">${isExcl ? '+ Include' : '⊘ Excl.'}</button>
                            ${val || isExcl ? `<button type="button" data-clear-col="${this._esc(col)}" title="Clear mapping/exclusion" style="background:none;border:1px solid #FFBDBC;border-radius:3px;cursor:pointer;padding:2px 5px;font-size:10px;color:#BB0000;margin-left:3px;line-height:1;">✕</button>` : ''}
                          </td>
                        </tr>`;
                      }).join('')}
                    </tbody>
                  </table>
                  <div style="padding:7px 10px;display:flex;align-items:center;gap:8px;border-top:1px solid #D9E0E6;background:#F0F4F8;">
                    <span style="font-size:10px;color:#556B82;">Leave blank = use file column name as-is in CSV header.</span>
                    <div style="margin-left:auto;display:flex;gap:8px;">
                      <button id="autoMapAllBtn" class="btn btn-secondary" type="button" style="font-size:11px;padding:4px 12px;white-space:nowrap;" title="Apply default aliases: ID→CAPEX_ITEM, H1→AMOUNT, DESCRIPTION→excluded">⚡ Defaults</button>
                      <button id="applyMappingBtn" class="btn btn-primary" type="button" style="font-size:11px;padding:4px 12px;white-space:nowrap;">💾 Apply</button>
                    </div>
                  </div>
                </div>` : ''}
              </div>`;
            })()}
            <div style="width:100%;display:flex;align-items:flex-end;gap:14px;flex-wrap:wrap;padding-top:10px;">
              ${this._targetVersion && this._targetDate
                ? (this._sacModelId
                    ? '<span class="sac-target-badge" style="background:#E6F9EC;color:#107E3E;border:1px solid #A3E6BF;padding:5px 14px;border-radius:20px;font-size:11px;font-weight:700;">✅ ' + this._esc(this._targetVersion) + ' · ' + this._esc(this._targetDate) + ' → ' + this._esc(this._sacModelId) + '</span>'
                    : '<span class="sac-target-badge" style="background:#E6F9EC;color:#107E3E;border:1px solid #A3E6BF;padding:5px 14px;border-radius:20px;font-size:11px;font-weight:700;">✅ ' + this._esc(this._targetVersion) + ' · ' + this._esc(this._targetDate) + ' <span style=\'color:#B45309\'>(no Model ID — event only)</span></span>'
                  )
                : '<span class="sac-target-badge" style="background:#FFF8E6;color:#B45309;border:1px solid #FCD34D;padding:5px 14px;border-radius:20px;font-size:11px;font-weight:700;">⚠️ Set Version and Period before synchronizing</span>'
              }
            </div>
          </div>

          ${this._loadStrategy === 'DELTA' ? (() => {
            const dp = this._deltaPreview;
            const canPreview = !!(this._sacModelId && this._targetVersion && this._targetDate);
            return `
            <div style="margin-bottom:14px;border:2px solid #FFC947;border-radius:10px;background:#FFFDF0;overflow:hidden;">
              <div style="display:flex;align-items:center;justify-content:space-between;padding:10px 16px;background:#FFF8E1;gap:10px;flex-wrap:wrap;">
                <div style="display:flex;align-items:center;gap:10px;">
                  <span style="font-size:18px;">🔍</span>
                  <div>
                    <div style="font-size:13px;font-weight:800;color:#7A5A00;">Delta Preview — existing data in this region</div>
                    <div style="font-size:11px;color:#9A7000;">
                      ${canPreview
                        ? 'Version: <strong>' + this._esc(this._targetVersion) + '</strong> · Period: <strong>' + this._esc(this._targetDate) + '</strong> · Model: <strong>' + this._esc(this._sacModelId) + '</strong>'
                        : 'Set VERSION, PERIOD and MODEL ID to enable preview.'}
                    </div>
                  </div>
                </div>
                <button id="loadDeltaPreviewBtn" class="btn" type="button"
                  ${canPreview ? '' : 'disabled'}
                  style="background:${canPreview ? '#FFC947' : '#D9D9D9'};color:${canPreview ? '#3D2800' : '#888'};font-weight:700;font-size:12px;padding:7px 16px;border-radius:6px;border:none;cursor:${canPreview ? 'pointer' : 'default'};">
                  ${dp.loading ? '⏳ Loading…' : '📊 Load Preview'}
                </button>
              </div>
              ${dp.loading ? '<div style="padding:20px;text-align:center;color:#9A7000;font-size:12px;">Fetching data from SAC DIS…</div>' : ''}
              ${dp.error ? `<div style="padding:12px 16px;font-size:12px;color:#BB0000;background:#FFF0F0;border-top:1px solid #FFCCCC;">⚠️ ${this._esc(dp.error)}</div>` : ''}
              ${dp.loaded && !dp.loading && !dp.error ? (
                dp.rows.length === 0
                  ? `<div style="padding:14px 16px;font-size:12px;color:#556B82;border-top:1px solid #FFE082;">
                      <strong>No existing data</strong> found for Version <em>${this._esc(this._targetVersion)}</em> · Period <em>${this._esc(this._targetDate)}</em>.
                      This delta load will insert all rows as new records.
                    </div>`
                  : `<div style="padding:8px 16px 14px;border-top:1px solid #FFE082;">
                      <div style="font-size:11px;color:#9A7000;font-weight:700;margin-bottom:6px;">${dp.rows.length} existing row(s) — these will be updated or preserved depending on key match:</div>
                      <div class="tbl-wrap" style="max-height:220px;overflow:auto;">
                        <table>
                          <thead><tr>${dp.cols.map(c => '<th>' + this._esc(c) + '</th>').join('')}</tr></thead>
                          <tbody>${dp.rows.map(r => '<tr>' + dp.cols.map(c => '<td>' + this._esc(r[c] ?? '') + '</td>').join('') + '</tr>').join('')}</tbody>
                        </table>
                      </div>
                    </div>`
              ) : ''}
            </div>`;
          })() : ''}

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
                      ${this._requiredColumns.map(h => '<th><div class="th-content"><span>' + this._esc(h) + '</span><button class="remove-col-btn" data-col="' + this._esc(h) + '">✕</button></div></th>').join('')}
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
} // end double-load guard