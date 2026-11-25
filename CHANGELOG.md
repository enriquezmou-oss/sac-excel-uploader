# Changelog — Excel Planning Uploader Pro

All notable changes to this project are documented here.
Format: **[version] — date — summary**

---

## [3.9.8] — 2026-11-25

### Audit Monitor: strategy+method column, SAP user detection, SAC Help link

- **Audit Monitor — Strategy column** — now shows the import method alongside the load strategy: `FULL 🗑️ Replace`, `FULL ✏️ Update`, `FULL ➕ Append`. The audit detail panel also shows the combined value. Works for both new and historical records.
- **SAP User auto-detection enhanced** — added SAC-specific global paths (`sap.sac.app.accessControl.currentUser`, `sap.sac.app.context.user`, `sap.sac.DataImport.user`) and two additional REST endpoints. User discovery is now retried a second time after the DIS authentication succeeds, improving the chance of resolving the actual logged-in user.
- **SAP Analytics Cloud badge** — now a clickable link to `help.sap.com/docs/SAP_ANALYTICS_CLOUD` (opens in new tab).

---



### 4 bug fixes + import method warnings per data type

- **Fix Bug 1 (duplicate click listener)** — `document.addEventListener('click', …)` inside `_bindEvents()` was registered on every `_render()`, accumulating duplicate listeners with every UI update. Fixed with a `_docClickBound` one-time guard. Also fixed a state desync: clicking outside the model picker now sets `this._showModelPicker = false` (previously only hid the DOM element without updating state).
- **Fix Bug 2 (dead Delta Preview block)** — the Delta Preview card in the render template was conditioned on `this._loadStrategy === 'DELTA'`, which is always `'FULL'` since the Full/Delta toggle was removed in v3.9.5. The entire permanently-false block is removed, along with the unreachable DELTA branch in `_callDISApi`.
- **Fix Bug 3 (orphaned merge code)** — `_mergePlanningData()` method and the `_mergeWarningMsg` toast handler were left over from the v3.9.3 client-side merge approach (removed in v3.9.5). Both are now deleted.
- **Fix Bug 4 (phantom audit record on cancel)** — the audit record was added to `_history` and `propertiesChanged` was dispatched **before** `_callDISApi` was called. If the user cancelled the confirmation dialog, a phantom record appeared in the Audit Monitor with no actual upload. Fix: `_callDISApi` now returns `{ cancelled: true }` on cancel; the caller immediately removes the phantom record before returning.
- **Feature: import method warnings** — the Append button is now visually disabled (greyed out, `cursor:not-allowed`, tooltip) when `_sacImportType === 'privateFactData'`. An amber warning banner explains: *"Append is only available for public fact data (factData). For private planning versions SAC DIS will Replace matching rows regardless."* All three method descriptions updated with clearer language about DIS behaviour.

---

## [3.9.6] — 2026-11-24

### SAC story bridge: data check via Table_1

- **New property `existingDataInfo`** — a SAC Application script can inject the row count from any table widget (e.g. Table_1) into the uploader before an Upload/Append operation. The widget reads this value in priority over the OData Export endpoint.
- **Confirmation dialog enhanced** — if `existingDataInfo` carries `{rowCount: N, hasData: true/false}`, the dialog shows `📊 N existing row(s) found in model (via story table)` or `✅ No existing data found` before asking the user to confirm or cancel.
- **OData fallback kept** — if `existingDataInfo` is not set (older stories without script), the widget still tries the OData Export endpoint as before (returns 400 on most tenants without OAuth, so the dialog shows no row count — but still asks for confirmation).
- **SAC script to wire up** — in SAC Application Builder, add this to `Table_1 → onResultChanged`:
  ```javascript
  // Replace 'ExcelUploaderPro_1' with your widget's name in the story
  ExcelUploaderPro_1.existingDataInfo = JSON.stringify({
    rowCount: Table_1.getDataBinding().getData().rows.length,
    hasData:  Table_1.getDataBinding().getData().rows.length > 0
  });
  ```

---

## [3.9.5] — 2026-11-24

### Update/Append: data existence check with confirmation dialog

- **Simpler approach** — instead of client-side merging (which required OData reads that returned 400), Update/Append modes now check whether data already exists for the target Version+Date and show a confirmation dialog if it does.
- **Confirmation dialog** — if existing rows are found, a dialog shows the row count (`⚠️ Found N rows for Version: X / Date: Y`) and asks the user to proceed or cancel. If the user cancels, the upload is aborted.
- **No existing data** — if no rows are found (or the check endpoint fails), the upload proceeds immediately without any interruption.
- **OData endpoint fix** — `_fetchExistingDataOData` now tries the OData export endpoint without `$filter` first (SAC returns 400 when `$filter` is present on some tenants). Client-side filtering handles Version+Date matching after fetching all rows.
- **Removed client-side merge** — `_mergePlanningData` is no longer called. DIS v1 always replaces matching dimension-key rows; merge was unreliable when OData reads returned 400.

---

## [3.9.4] — 2026-11-24

### 9 bug fixes: merge safety, double upload, regex, ID column, Audit colspan, blur-before-sync

- **Fix 1 (double upload)** — `_callDISApi` step 4 CSV re-upload was incorrectly triggered when the job status GET failed (due to `|| !jobInitStatus` condition). Now fires only when status is explicitly `READY_FOR_DATA`.
- **Fix 2A (Append case mismatch)** — OData may return column names in different casing than the uploaded rows. `_mergePlanningData` now uses a case-insensitive `existKey` lookup when accumulating measure values in Append mode.
- **Fix 2B (empty dimensions guard)** — If `_sacModelDimensions` is empty at merge time, all columns would have been treated as dimension keys, corrupting the merge key. Added early-return guard: merge skipped, uploaded rows returned as-is, and a console warning emitted.
- **Fix 3 (merge warning never shown)** — `_mergeWarningMsg` was set but never surfaced. Now shown as a warning toast 3.5 s after the success toast, so the user is informed if existing data could not be fetched before an Update/Append operation.
- **Fix 4 (dimensions never auto-fetched)** — `_fetchModelDimensions()` was never called automatically; dimensions were always empty at first sync. Now called in both branches of `_fetchModelInfo()` after model state becomes `ok`.
- **Fix 5 (ID column deletion blocked)** — `if (colToRemove === 'ID') return` guard prevented the ID column from being removed from the Required Columns grid. Guard removed; `_columnMapping` and `_disExcludedCols` are also cleaned up on removal.
- **Fix 6 (Audit Monitor empty-state colspan)** — Empty-state row `colspan` corrected from 9 to 11 to match the actual 11-column table.
- **Fix 7A (last cell edit lost)** — `document.activeElement.blur()` now called at the start of the Synchronize handler so any in-progress cell edit is committed before the grid rows are read.
- **Fix 7B (model picker never closes)** — Added a `document.click` listener so the model picker dropdown closes when the user clicks outside it.
- **Fix 8 (REGEX errors swallowed)** — The `catch` block for `new RegExp(rule.param)` was empty; invalid regex patterns silently passed validation. Now reports `'Invalid regex: ' + e.message` in the cell as a failed validation.

---

## [3.9.3] — 2026-11-24

### Fix: Update/Append — client-side merge (DIS v1 always Replaces)

- **Root cause confirmed** — DIS v1 ignores `writeBackMode`/`importMode`/`writeBack` JSON body params for `privateFactData`. The API accepts them but always replaces data for the uploaded Version+Period.
- **Fix** — for Update/Append, the widget now:
  1. Fetches existing data for the target Version+Period via OData export (`/api/v1/dataexport/providers/sac/{modelId}`)
  2. Merges the datasets in memory (client-side)
  3. Uploads the complete merged dataset via single-step Replace
- **Update merge** — existing rows not in the uploaded file are preserved; uploaded rows replace matching dimension-key combinations.
- **Append merge** — measure values are accumulated for matching dimension keys; new keys are inserted; all other existing rows are kept.
- **Fallback** — if the OData endpoint is unavailable (404), a warning is logged and the upload proceeds as Replace.
- **Simplified DIS flow** — all modes now use the proven single-step CSV POST. The JSON job creation branch has been removed.
- **New methods** — `_fetchExistingDataOData(version, date)` and `_mergePlanningData(uploadedRows, existingRows, mode, dimensions)`.

---

## [3.9.2] — 2026-11-24

### Fix: import method (Update/Append) always replaced — DIS write mode enforced

- **Root cause** — `?method=UPDATE` URL param is ignored by SAC DIS v1 (always replaces).
- **Fix** — for Update/Append, POST a JSON settings body to create an empty job. Write mode specified in `writeBackMode`/`importMode`/`writeBack` fields. Replace keeps single-step CSV POST (unchanged).
- **Headers** — write mode headers (`x-sap-dis-write-mode`, `x-sap-dis-import-mode`) added to the CSV upload step.
- **Run step** — write mode JSON body also sent on the `/run` call for non-replace modes.
- **Step 4 guard** — CSV upload now always triggered for Update/Append (previously skipped if initial job status wasn't `READY_FOR_DATA`).

---

## [3.9.1] — 2026-11-24

### All UI messages in English + Full load badge shows method

- All remaining Spanish strings translated to English (console logs, toasts, loading text, `_pollJobStatus` timeout error).
- Full load strategy badge now shows the active import method: `📦 Full · 🗑️ Replace`, `📦 Full · ✏️ Update`, or `📦 Full · ➕ Append`.

---

## [3.9.0] — 2026-11-24

### Auto-load model dimensions via FPA REST · Real SAP user in Audit Monitor

- **FPA REST as first priority** — `_fetchModelDimensions()` now POSTs to `/sap/fpa/services/rest/fpa/model` as the first attempt. Payload: `{"action":"getDesignTime","data":{"resourceId":"{DIS_GUID}","metadata":{},"includeDatasphereInfos":true}}`. This endpoint is confirmed working on this tenant and returns the full `DIMENSION` array.
- **`_parseDimensionsFromFPA()`** — new parser for FPA REST response format: `DIMENSION[].metadata.id.name` → dimension ID; `DIMENSION[].data.type` → type (`VersionDim` → isAuto, `TimeHANA` → isAuto).
- **`_discoverTenantUUID()`** — new discovery helper: tries SAC global variables (`window.__sap_umi_config`, `window.sap.umi.config`) first, then `GET /sap/fpa/services/rest/session`. Tenant UUID cached in `_cachedTenantUUID`. FPA endpoint called both with `?tenant=` and without, so it works across tenants.
- **DIS GET fallback** — 10 DIS endpoint variants kept as fallback in case FPA REST is unavailable.
- **Real SAP user in Audit Monitor** — `_discoverTenantUUID()` also parses `loginName`/`userName` from the session response. If `currentUser` is still the default placeholder ("SAP User (Planning)"), it is replaced automatically with the actual logged-in user. No configuration needed.
- **Non-blocking startup** — `_discoverTenantUUID()` is called at the start of `_fetchModelInfo()` as a background task so user and tenant UUID are ready before the first sync.

---

## [3.8.0] — 2026-11-24

### What's New panel + single version constant

- **What's New button** — new ★ button in the header (next to Help) opens a version history panel inline, without leaving the widget. Lists all versions with date, title, and bullet-point improvements. Latest version highlighted in green.
- **Single version constant** — `_EPM_VERSION` now defined once at the top of the file. Subtitle, Help modal header, and changelog panel all reference it — no more manual updates in multiple places.
- **Help / What's New mutual exclusion** — opening one automatically closes the other.

---

## [2.7.2] — 2026-11-23

### Fix: ID column removal + dimension probe logging

- **ID column ✕ button** — removed guard that prevented deleting the `ID` column from the grid schema. All columns can now be removed.
- **Dimension probe logging** — added per-endpoint `[DIS] dim probe` console logs to diagnose which DIS API endpoint format is active on the tenant.

---

## [2.7.1] — 2026-11-23

### Fix: Model Structure Validation — dimensions not loading on this tenant

The `GET /api/v1/dataimport/models/{GUID}` endpoint returns 404 on this SAC tenant, causing the validation panel to show 0 dimensions.

**Root cause:** `_fetchModelDimensions()` was calling the model metadata endpoint, which is not supported here. The correct endpoint on this tenant is the same CSRF endpoint used for the upload flow: `GET /api/v1/dataimport/models/{GUID}/{importType}`, which returns a `columns` array in its response body listing all valid import columns.

**Changes:**
- `_fetchModelDimensions()` now tries six endpoint variants in order, stopping at the first that returns parseable dimension data: CSRF endpoint with GUID, CSRF endpoint with technical name, `/dimensions` with GUID, `/dimensions` with technical name, metadata with GUID, metadata with technical name.
- `_parseDimensions()` now handles the SAC DIS `columns` array format (field name `columnID` instead of `id`, field `columnType` instead of `type`, field `columnDescription` instead of `description`).
- Added `SIGNEDDATA` to the set of measure types.
- Error message when 0 dimensions are returned is now more actionable: *"DIS did not return column metadata for this model. Configure Required Columns manually."*

**Why this matters:** The dimensions panel can now show users which columns the SAC model accepts. When columns like `ID`, `DESCRIPTION`, or `H1` are in Required Columns but not in the model, the panel marks them ❌ missing and blocks Synchronize — preventing the HTTP 400 "Request columns did not match import type" error.

---

## [2.7.0] — 2026-11-23

### Model Structure Validation

New inline panel in the SAC Load Target section that fetches and displays the planning model's dimension structure directly from SAC DIS, validates it against the configured Required Columns, and blocks Synchronize if any dimension is unmapped.

**How it works:**
1. When the model connects successfully, `_fetchModelDimensions()` is called automatically.
2. It calls `GET /api/v1/dataimport/models/{GUID}`, falling back to `GET /api/v1/dataimport/models/{GUID}/dimensions` if the first returns an error.
3. The response is parsed into `_sacModelDimensions`: `[{ id, description, type, isMeasure, isAuto }]`.
4. The panel renders a table showing every dimension with its ID, description, type, and mapping status.

**Status logic per dimension:**
- ✅ **auto** — `VERSION` and `DATE` types: handled automatically via the VERSION / PERIOD fields.
- ✅ **mapped** — dimension ID is present in Required Columns (case-insensitive match).
- ❌ **missing** — dimension not in Required Columns and not auto-handled.

**Actions:**
- **⚡ Add missing to Required Columns** — appends all unmapped dimension IDs in one click and re-renders.
- **Synchronize blocked** — if any ❌ missing dimensions exist when Synchronize is clicked, the operation is blocked with a toast listing them. The panel is forced open.

**Collapsible panel** — header shows `✅ All dimensions mapped` or `⚠️ N dimensions unmapped`. Click to expand/collapse.

**Graceful fallbacks** — if the DIS endpoint fails or returns no dimensions, the panel shows the error and allows manual Required Columns configuration.

---

## [2.6.1] — 2026-11-23

- **Fix DIS run step (405)** — `/jobs/{jobId}/run` returns 405 on this tenant (job auto-executes on creation). Detected and skipped silently; proceeds directly to polling.
- **MODEL NAME field always visible** — renders permanently with `Model name: —` placeholder until a description is fetched.
- **MODEL NAME auto-populated on connect** — `_fetchModelInfo()` saves the description to `localStorage` on DIS connect.

---

## [2.6.0] — 2026-11-23

- **Critical DIS fix** — single-step CSV upload: `POST /models/{GUID}/{importType}` with `Content-Type: text/csv`. The separate `/jobs/{jobId}` upload endpoint (which returned 405) removed.
- **MODEL ID and MODEL NAME fields** — two side-by-side input fields. MODEL ID editable; MODEL NAME readonly, filtered to avoid showing redundant info.
- **Model Picker redesigned** — technical name as primary label, description as subtitle, GUID in monospace. `data-model-desc` added.
- **Real-time search** — filters across technical name, description, and GUID simultaneously.
- **Model grouping** — Planning (privateFactData) first, Actual/Others second.

---

## [2.5.0] — 2026-11-22

- **Critical DIS fix** — automatic GUID resolution via `_resolveDISModelId()`.
- **Model browser fix** — `GET /api/v1/dataimport/models` without `?importType=` parameter.
- **CSRF fix** — token fetched from the model-specific endpoint.
- **Double-load guard** — prevents `SyntaxError: Identifier already declared` on stale cache.
- **CSS fix** — `.toast-warning { background:#E76500 }` (SAP orange).
- Auto-detection of `importType`: `privateFactData` → `factData` fallback.

---

## [2.2.1]

- **Bug fix** — `_evalAssistedCode`: fallback uses expression string as cache key.
- **Bug fix** — `_applyAllAutoFixes`: empty `catch {}` replaced with `console.warn`.
- **Bug fix** — `_storageSet`: only suppresses `QuotaExceededError`.

---

## [2.2.0]

- Removed DSP Artifact toggle from the UI.
- Redesigned Full / Delta strategy selection bar.
- Improved table scrollbars in SAP Blue.
- `requestAnimationFrame` debounce on grid and audit search inputs.

---

## [2.1.0]

- Compiled rule function cache (`_compiledRules`).
- `crypto.randomUUID()` for collision-free rule IDs.
- Early exit in `_validateRows()` when grid is empty.

---

## [2.0.0]

Initial production release. Multi-file/multi-sheet consolidation, Full/Delta load strategy, Validation engine (6 rule types, 4 severity levels), Assisted Code Auto-Fix, Audit Monitor, Help Guide overlay.

---

## Versioning convention

| Change type | Increment |
|-------------|-----------|
| Bug fix with no visible UX impact | Patch `x.x.P` |
| New feature or visible improvement | Minor `x.M.0` |
| Architecture change or breaking change | Major `M.0.0` |
