 Excel Planning Uploader Pro — SAP Analytics Cloud Custom Widget

> Enterprise-grade Web Component for SAP Analytics Cloud (SAC) **v2.6.0** — Multi-sheet Excel/CSV/TXT consolidation, configurable validation semaphores, Assisted Code Auto-Fix, Full/Delta load engine, **direct DIS write-back to SAC planning models**, drill-down Audit Monitor with multi-format exports, and SAP Fiori Horizon UX.

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

### ❓ Interactive Help Guide
Built-in overlay (`? Help` button in the header) — step-by-step documentation without leaving the widget.

---

## Technical Notes

| Item | Detail |
|------|--------|
| **Custom element tag** | `<custom-excel-uploader-pro>` |
| **Shadow DOM** | `mode: 'open'` |
| **XLSX library** | SheetJS `0.18.5`, loaded from CDN automatically if not already present |
| **Persistence** | `localStorage` keys: `sac_required_cols`, `sac_validation_rules_v9`, `sac_audit_history_v9` |
| **SAC integration** | Dispatches `propertiesChanged` custom event with `validatedDataJson` property |
| **DIS write-back** | `GET /api/v1/dataimport/models` → GUID resolution → CSRF → POST job → CSV upload → run → poll |
| **CDN** | jsDelivr pinned por commit hash — inmutable, sin caché sorpresa |

---

## SAC DIS Write-Back

El widget escribe directamente en un modelo de planificación SAC nativo mediante la **Data Import Service (DIS) REST API**:

1. **Resolución de GUID** — el nombre técnico (`PLAN_CAPEX_PLANNING`) se resuelve automáticamente al GUID interno (`C9qfcbu7dll3q7ao78qsq5e9h4o`) llamando a `GET /api/v1/dataimport/models`.
2. **CSRF token** — se obtiene del endpoint del modelo específico (`/models/{GUID}/{importType}` con `x-csrf-token: fetch`).
3. **Creación de job** — `POST /models/{GUID}/privateFactData` → devuelve `jobId`.
4. **Upload CSV** — `PUT /jobs/{jobId}` con los datos en CSV.
5. **Run & poll** — `POST /jobs/{jobId}/run` → `GET /jobs/{jobId}` hasta `successful`.

**Autenticación:** cookies de sesión SAC (`credentials: 'include'`). No requiere OAuth Bearer token.

### Model Picker
Pulsa **🔍 SAC** junto al campo MODEL ID para explorar los modelos disponibles:
- Agrupados: **📋 Planning** (privateFactData) · **📊 Actual / Otros** (factData)
- Búsqueda en tiempo real por descripción, nombre técnico o GUID
- El modelo activo aparece marcado con ✓

---

## Installation

1. Copy `excel-uploader.js` to your SAC custom widget project.
2. Register the widget in your SAC tenant (Custom Widgets panel).
3. Add the `<custom-excel-uploader-pro>` tag to your story or application.

SheetJS is fetched from CDN on first load. If your SAC tenant blocks external CDN requests, bundle `xlsx.full.min.js` locally and load it before this widget.

---

## Changelog

### v2.6.0 — 2026-11-23
- **Model Picker rediseñado** — nombre técnico como título (antes mostraba GUID críptico)
- **Búsqueda en tiempo real** — filtra por descripción, nombre técnico o GUID
- **Grupos** — 📋 Planning (privateFactData) primero · 📊 Actual/Otros (factData) después
- **Modelo activo** marcado con ✓ y fondo azul
- Al seleccionar se guarda el nombre técnico (no el GUID) en MODEL ID
- GUID cacheado en `_sacDISModelId` — evita round-trip extra en el siguiente Synchronize

### v2.5.0 — 2026-11-22
- **Fix crítico DIS**: resolución automática de GUID interno (`_resolveDISModelId`)
- **Fix model browser**: `GET /api/v1/dataimport/models` sin parámetros `?importType=`
- **Fix CSRF**: token obtenido del endpoint del modelo, no del listado
- **Fix double-load**: guard `if (typeof ExcelUploaderPro === 'undefined')` evita SyntaxError
- **Fix CSS**: `.toast-warning { background:#E76500 }` (naranja SAP)
- Auto-detección de `importType`: `privateFactData` → `factData` como fallback

### v2.2.1
- **Bug fix** — `_evalAssistedCode`: fallback path uses expression string as cache key
- **Bug fix** — `_applyAllAutoFixes`: empty `catch {}` replaced with `console.warn`
- **Bug fix** — `_storageSet`: only suppresses `QuotaExceededError`

### v2.2.0
- Removed DSP Artifact toggle from UI
- Redesigned Full / Delta strategy bar
- Improved table scrollbars in SAP Blue
- `requestAnimationFrame` debounce on grid and audit search inputs

### v2.1.0
- Compiled rule function cache (`_compiledRules`)
- `crypto.randomUUID()` for collision-free rule IDs
- Early exit in `_validateRows()` when grid is empty

### v2.0.0
- Multi-file / multi-sheet consolidation
- Full / Delta load strategy
- Validation engine: 6 rule types, 4 severity levels
- Assisted Code Auto-Fix
- Audit Monitor with XLSX / CSV export
- Help Guide overlay

