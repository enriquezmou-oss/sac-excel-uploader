 File Planning Uploader Pro — SAP Analytics Cloud Custom Widget

> Enterprise-grade Web Component for SAP Analytics Cloud (SAC) v2.2.1 — Multi-sheet Excel/CSV/TXT consolidation, configurable validation semaphores, Assisted Code Auto-Fix, Full/Delta load engine, drill-down Audit Monitor with multi-format exports, and SAP Fiori Horizon UX.

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

---

## Installation

1. Copy `excel-uploader.js` to your SAC custom widget project.
2. Register the widget in your SAC tenant (Custom Widgets panel).
3. Add the `<custom-excel-uploader-pro>` tag to your story or application.

SheetJS is fetched from CDN on first load. If your SAC tenant blocks external CDN requests, bundle `xlsx.full.min.js` locally and load it before this widget.

---

## Changelog

### v2.2.1
- **Bug fix** — `_evalAssistedCode`: fallback path (no `cacheKey`) now uses the expression string itself as cache key — `new Function()` always compiled at most once per unique expression
- **Bug fix** — `_applyAllAutoFixes`: empty `catch {}` replaced with `console.warn` + failed count surfaced in the toast notification
- **Bug fix** — `_storageSet`: only suppresses `QuotaExceededError`; other localStorage failures now emit `console.warn`
- DSP Artifact references removed from code comments

### v2.2.0
- **Removed** DSP Artifact toggle from UI — SAC Artifact sync remains always active internally
- **Redesigned** Full / Delta strategy bar: cleaner layout with active-mode badge
- **Improved** table scrollbars: visible 7 px scrollbar in SAP Blue (`#0070F2`) on `.tbl-wrap`, host scroll, JSON `<pre>` block, and help modal
- **Performance** — `requestAnimationFrame` debounce on both grid search and audit search inputs (no full re-render on every keystroke)
- **UX** — `disconnectedCallback` now cancels pending RAF tasks on element removal

### v2.1.0
- Compiled rule function cache (`_compiledRules`) — ASSISTED_CODE rules compiled once per session
- `crypto.randomUUID()` for collision-free rule IDs (fallback to `Math.random`)
- Early exit in `_validateRows()` when grid is empty
- User toast feedback when export functions have nothing to export
- Cache invalidated automatically whenever rules are saved

### v2.0.0 *(Gemini-generated baseline)*
- Multi-file / multi-sheet consolidation
- Full / Delta load strategy with DSP Artifact toggle
- Validation engine: 6 rule types, 4 severity levels, scoped rules, temporal validity
- Assisted Code Auto-Fix with `new Function()`
- Audit Monitor with XLSX / CSV export
- Help Guide overlay

