# Changelog — Excel Planning Uploader Pro

All notable changes to this project are documented here.
Format: **[version] — date — summary**

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
