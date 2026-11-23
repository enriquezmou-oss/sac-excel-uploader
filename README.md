# Excel Planning Uploader Pro — SAP Analytics Cloud Custom Widget

> Enterprise-grade Web Component for SAP Analytics Cloud (SAC) **v2.7.3** — Multi-sheet Excel/CSV/TXT consolidation, configurable validation semaphores, Assisted Code Auto-Fix, Full/Delta load engine, **direct DIS write-back to SAC planning models**, drill-down Audit Monitor with multi-format exports, and SAP Fiori Horizon UX.

**Vendor:** Jose Enriquez SixSevenGroup

---

## What It Does

**Excel Planning Uploader Pro** lets SAC planning users upload one or more Excel workbooks (or CSV/TXT files), select individual sheets, validate every cell against configurable business rules before the data reaches SAP Analytics Cloud, and produce a full audit trail of every action taken during the load.

---

## Key Features

### 🔄 Full / Delta Load Strategy
Choose between two ingestion modes, displayed as visual Fiori cards:

| Mode | Behaviour |
|------|-----------|
| **FULL LOAD** | Replaces the entire target planning slice |
| **DELTA LOAD** | Appends or updates only the changed records |

The selected strategy is recorded in every audit record and in the JSON payload dispatched to SAC.

### 📑 Multi-File & Multi-Sheet Consolidation
- Drag & drop (or browse) multiple `.xlsx`, `.xls`, `.xlsm`, `.csv`, or `.txt` files simultaneously
- Sheet-level selection per workbook — include or exclude individual tabs
- Business audit comment per file, persisted in the audit trail
- Re-parse with updated settings (delimiter, number format, header row) without re-uploading

### 🛠️ Format & Delimiter Configuration
- Auto-detect format (Excel / Flat file)
- Delimiter: Auto, Tab, Semicolon, Comma, Pipe
- Number format: Standard `1,234.56` or European `1.234,56 → 1234.56`
- Configurable header row offset (1–20)

### 🚦 Validation Rules Engine
Four severity levels, each with a configurable colour semaphore:

| Severity | Semaphore | Behaviour |
|----------|-----------|-----------|
| **Blocking** | 🔴 Red | Prevents synchronisation until resolved |
| **Preventive** | 🟡 Amber | Requires explicit sign-off before sync |
| **Check** | 🟣 Purple | Cell-level verification by the user |
| **Info** | 🔵 Blue | Guidance only, no block |

**Rule types:** `IN_LIST`, `NUMERIC_RANGE`, `STARTS_WITH`, `REGEX`, `UNIQUE`, `ASSISTED_CODE`

**Per-rule options:**
- `scopeFiles`: apply to `ALL` workbooks or specific filenames
- `startDate` / `endDate`: temporal validity window
- `autoFix`: JavaScript expression to repair the cell automatically (`✨`)
- Rule cloning (`📋 Clone`) and inline scope editor — no browser popups

### ✨ Assisted Code Auto-Fix
Write arbitrary JavaScript expressions that receive `value`, `num`, and the full `row` object. Compiled rule functions are **cached per rule ID** — they are compiled with `new Function()` only once per session and reused on every subsequent validation pass, regardless of dataset size.

### 🧱 Dynamic Schema Management
- Add or remove mandatory columns at runtime
- Schema persisted in `localStorage`; new columns backfilled with empty strings on existing rows

### ✏️ Inline Grid Editing
- `contenteditable` cells in Step 3
- Bulk Fill: fill empty cells or overwrite all visible rows in a column
- Add manual row, clone row (`⧉`), delete row — all tracked in the audit log

### 🛡️ Audit Monitor
Every successful synchronisation creates an immutable audit record containing:
- Load ID, strategy, timestamp, SAP user
- Committed row snapshot with per-cell edit history
- User actions log (cell edits, auto-fixes, rule changes, bulk fills, verified checks)
- Applied validation rules snapshot
- Full JSON payload dispatched to SAC

#### Export formats
| Format | Output |
|--------|--------|
| `.xlsx` | Multi-sheet Excel: Data Snapshot, Workbooks & Comments, Audit Activity Log |
| `.csv` | Comma-separated flat file |
| `.txt` | Tab-delimited text file |
| `.json` | Full audit certificate (single record) |

Custom export configurator: select exact column count, toggle business comments, toggle full audit trail.

### ❓ Interactive Help Guide & What's New
Built-in overlays in the header — **? Help** shows step-by-step documentation; **★ What's New** shows version history with improvements. Both open inline without leaving the widget.

---

## Technical Notes

| Item | Detail |
|------|--------|
| **Custom element tag** | `<custom-excel-uploader-pro>` |
| **Shadow DOM** | `mode: 'open'` |
| **XLSX library** | SheetJS `0.18.5`, loaded from CDN automatically if not already present |
| **Persistence** | `localStorage` keys: `sac_required_cols`, `sac_validation_rules_v9`, `sac_audit_history_v9`, `sac_dis_model_id`, `sac_dis_guid`, `sac_dis_model_desc` |
| **SAC integration** | Dispatches `propertiesChanged` custom event with `validatedDataJson` property |
| **DIS write-back** | GUID resolution → CSRF → POST CSV to model endpoint → auto-poll until `SUCCESSFUL` |
| **CDN** | jsDelivr pinned to commit hash — immutable, no surprise cache |

---

## SAC DIS Write-Back

The widget writes directly to a native SAC planning model using the **Data Import Service (DIS) REST API**:

1. **GUID resolution** — the technical name (e.g. `PLAN_CAPEX_PLANNING`) is resolved automatically to the internal DIS GUID by calling `GET /api/v1/dataimport/models`. The GUID is cached in `localStorage` to avoid a round-trip on every sync.
2. **CSRF token** — fetched from the model-specific endpoint: `GET /models/{GUID}/{importType}` with `x-csrf-token: fetch`.
3. **Create job + upload CSV** — `POST /models/{GUID}/privateFactData` with `Content-Type: text/csv` and the CSV data in the body. Returns `{ jobID }`. The job is created and data uploaded in a single HTTP call.
4. **Run** (optional) — `POST /jobs/{jobId}/run`. Some SAC tenants auto-execute the job on creation; if this endpoint returns 405, it is skipped silently.
5. **Poll** — `GET /jobs/{jobId}` polled every 3 s until status is `SUCCESSFUL` (or a failure status is detected).

**Authentication:** SAC session cookies (`credentials: 'include'`). No OAuth Bearer token required.

### SAC Load Target — MODEL ID & MODEL NAME fields

- **MODEL ID** — enter the technical name of the planning model (e.g. `PLAN_CAPEX_PLANNING`). The widget resolves it to the DIS internal GUID automatically.
- **MODEL NAME** — read-only display field below MODEL ID. Shows the model's friendly description once the DIS connection is established, or after selecting a model from the picker. Persists across page reloads via `localStorage`.

### Model Picker (🔍 SAC button)
Click **🔍 SAC** next to the MODEL ID field to browse available models:
- **Technical name** shown as the primary label (blue, bold)
- **Description** shown as a subtitle (grey)
- **Internal GUID** shown in small monospace text
- Grouped: **📋 Planning** (privateFactData) first · **📊 Actual / Others** (factData)
- Real-time search filtering by technical name, description, or GUID
- Active model marked with ✓ and blue background

---

## Installation

1. Copy `excel-uploader.js` to your SAC custom widget project.
2. Register the widget in your SAC tenant (Custom Widgets panel) by uploading `webcomponent.json`.
3. Add the `<custom-excel-uploader-pro>` tag to your SAC story or application.

SheetJS is fetched from CDN on first load. If your SAC tenant blocks external CDN requests, bundle `xlsx.full.min.js` locally and load it before this widget.

---

## Changelog

### v2.7.3 — 2026-11-23
- **What's New button** — ★ button in header opens version history panel inline. Latest version highlighted in green.
- **Single version constant** — `_EPM_VERSION` defined once; subtitle, header, and changelog panel always in sync.
- **Mutual exclusion** — opening What's New closes Help and vice versa.

### v2.7.2 — 2026-11-23
- **ID column removal** — ✕ button on the `ID` column in the Required Columns grid was incorrectly disabled; fixed.
- **Dimension probe logging** — per-endpoint `[DIS] dim probe` console logs added for tenant diagnostics.

### v2.7.1 — 2026-11-23
- **Dimension fetch fix** — `_fetchModelDimensions()` tries 6 endpoint variants (CSRF, /dimensions, metadata) with GUID and technical name. `_parseDimensions()` now handles `columnID` / `columnType` SAC DIS field names.

### v2.6.1 — 2026-11-23
- **Fix DIS run step (405)** — the `/jobs/{jobId}/run` endpoint returns 405 on this SAC tenant because the job is auto-executed on creation. The widget now detects this and skips the run call silently, proceeding directly to polling.
- **MODEL NAME field always visible** — previously only appeared when a description was already stored. Now always rendered below MODEL ID with a `Model name: —` placeholder until the description is fetched.
- **MODEL NAME auto-populated on connect** — `_fetchModelInfo()` now saves the model description to `localStorage` (`sac_dis_model_desc`) when it receives metadata from DIS, so the field fills in automatically without needing to use the picker.

### v2.6.0 — 2026-11-23
- **Critical DIS fix** — single-step CSV upload: data is POSTed directly to `/models/{GUID}/{importType}` with `Content-Type: text/csv`. The separate `/jobs/{jobId}` upload endpoint (which returned 405) has been removed from the flow.
- **MODEL NAME display field** — new read-only field below MODEL ID showing the model's friendly description. Populated via picker selection or on DIS connect. Persists in `localStorage`.
- **Model Picker redesigned** — technical name is now the primary label (previously showed description or GUID). Description shown as subtitle. Internal GUID shown in small monospace. `data-model-desc` attribute added to pass description through to the click handler.
- **Real-time search** — filters simultaneously across technical name, description, and GUID.
- **Model grouping** — Planning (privateFactData) listed first, Actual/Others (factData) second.
- **Active model indicator** — selected model shown with ✓ and blue background.
- Selecting a model from the picker saves the technical name to MODEL ID and caches the GUID to avoid extra resolution on the next sync.

### v2.5.0 — 2026-11-22
- **Critical DIS fix** — automatic GUID resolution via `_resolveDISModelId()`. DIS requires the internal GUID, not the technical name, for most endpoints.
- **Model browser fix** — `GET /api/v1/dataimport/models` called without `?importType=` parameter (caused 404 on this tenant).
- **CSRF fix** — token now fetched from the model-specific endpoint, not the models list.
- **Double-load guard** — `if (typeof ExcelUploaderPro === 'undefined')` prevents `SyntaxError: Identifier already declared` when uploading a new `webcomponent.json` while the previous one is still cached.
- **CSS fix** — `.toast-warning { background:#E76500 }` (SAP orange).
- Auto-detection of `importType`: tries `privateFactData` first, falls back to `factData` on error `[1704-1801-3304]` or 404.

### v2.2.1
- **Bug fix** — `_evalAssistedCode`: fallback path uses expression string as cache key.
- **Bug fix** — `_applyAllAutoFixes`: empty `catch {}` replaced with `console.warn`.
- **Bug fix** — `_storageSet`: only suppresses `QuotaExceededError`.

### v2.2.0
- Removed DSP Artifact toggle from UI.
- Redesigned Full / Delta strategy bar.
- Improved table scrollbars in SAP Blue.
- `requestAnimationFrame` debounce on grid and audit search inputs.

### v2.1.0
- Compiled rule function cache (`_compiledRules`).
- `crypto.randomUUID()` for collision-free rule IDs.
- Early exit in `_validateRows()` when grid is empty.

### v2.0.0
- Multi-file / multi-sheet consolidation.
- Full / Delta load strategy.
- Validation engine: 6 rule types, 4 severity levels.
- Assisted Code Auto-Fix.
- Audit Monitor with XLSX / CSV export.
- Help Guide overlay.
