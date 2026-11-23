# Changelog — Excel Planning Uploader Pro

All notable changes to this project are documented here.
Format: **[version] — date — summary**

---

## [2.6.1] — 2026-11-23

### DIS run step — 405 handled gracefully
The `/jobs/{jobId}/run` endpoint returns 405 on this SAC tenant because the job is auto-executed on creation when the CSV is POSTed to the model endpoint. The widget now detects the 405, logs `[DIS] run endpoint not available (405) — job auto-runs on this tenant`, and proceeds directly to polling. No error is thrown.

### MODEL NAME field always visible
Previously the description field below MODEL ID was hidden when no description was stored. It now renders permanently with a `Model name: —` placeholder, so the layout is stable from the first render.

### MODEL NAME auto-populated on DIS connect
`_fetchModelInfo()` now extracts the model description from the DIS metadata response and saves it to `localStorage` (`sac_dis_model_desc`). The MODEL NAME field fills in automatically whenever the widget establishes a DIS connection — no picker interaction required.

---

## [2.6.0] — 2026-11-23

### Critical DIS fix — single-step CSV upload
SAC DIS does not expose a separate `POST /jobs/{jobId}` or `PUT /jobs/{jobId}` data upload endpoint (both return 405). The correct flow is to send the CSV directly in the body of the job-creation call:

```
POST /api/v1/dataimport/models/{GUID}/{importType}
Content-Type: text/csv

<csv data>
```

This returns `{ jobID }` immediately. Data is uploaded and the job is created in one HTTP call. The previous two-step approach (create job, then upload data) has been removed.

### MODEL NAME display field
New read-only field rendered below the MODEL ID input. Shows the model's friendly description (`modelDescription` from DIS metadata). Populated on picker selection or automatically on DIS connect. Persisted in `localStorage` as `sac_dis_model_desc`.

### Model Picker redesigned
- **Technical name** (`modelName`) is now the primary label — bold, blue. Previously showed the description or GUID as the first line.
- **Description** shown as a secondary subtitle in grey (11 px).
- **Internal GUID** shown as a third line in small monospace.
- `data-model-desc` attribute added to each row element so the description is passed to the click handler and stored on model selection.
- Active model highlighted with ✓ prefix and light blue background.
- Models grouped: **📋 Planning** (`privateFactData`) first, **📊 Actual / Others** (`factData`) second.
- Real-time search filters across technical name, description, and GUID simultaneously.

---

## [2.5.0] — 2026-11-22

### Critical DIS fixes

**GUID resolution** — SAC DIS requires the internal GUID for most endpoints, not the technical name. `_resolveDISModelId()` calls `GET /api/v1/dataimport/models`, finds the matching entry, caches the GUID in `_sacDISModelId` and `localStorage` (`sac_dis_guid`), and returns it before every import. If resolution fails, the technical name is used as a fallback with a console warning.

**Model browser endpoint** — `GET /api/v1/dataimport/models` must be called without `?importType=` query parameters. This parameter caused a 404 on the target tenant.

**CSRF token** — the token must be fetched from the model-specific endpoint (`GET /models/{GUID}/{importType}` with header `x-csrf-token: fetch`), not from the models list endpoint. Using the wrong endpoint returned an empty token, causing all subsequent write calls to fail.

**Double-load guard** — added `if (typeof ExcelUploaderPro === 'undefined')` check at the top of the script to prevent `SyntaxError: Identifier 'ExcelUploaderPro' has already been declared` when SAC serves a stale cached version of the script alongside a freshly uploaded `webcomponent.json`.

**CSS fix** — `.toast-warning { background:#E76500; }` (SAP Orange) was missing.

**importType auto-detection** — tries `privateFactData` first; falls back to `factData` if the response is a 404 or contains error code `[1704-1801-3304]`. The detected type is persisted in `localStorage` (`sac_import_type`).

---

## [2.2.1]

- **Bug fix** — `_evalAssistedCode`: the fallback path was not storing compiled functions in the cache. Changed to use the expression string as the cache key.
- **Bug fix** — `_applyAllAutoFixes`: inner `catch {}` was swallowing exceptions silently. Replaced with `console.warn`.
- **Bug fix** — `_storageSet`: was catching all errors from `localStorage.setItem`. Now only suppresses `QuotaExceededError`.

---

## [2.2.0]

- Removed DSP Artifact toggle from the UI.
- Redesigned Full / Delta strategy selection bar (visual Fiori cards).
- Improved table scrollbars in SAP Blue.
- `requestAnimationFrame` debounce added to grid search and audit search inputs to prevent excessive re-renders on fast typing.

---

## [2.1.0]

- Compiled rule function cache (`_compiledRules`) — `new Function()` called once per rule ID per session; subsequent validation passes reuse the cached function.
- `crypto.randomUUID()` for collision-free rule IDs.
- Early exit in `_validateRows()` when the grid is empty.

---

## [2.0.0]

Initial production release.

- Multi-file / multi-sheet consolidation.
- Full / Delta load strategy.
- Validation engine: 6 rule types (`IN_LIST`, `NUMERIC_RANGE`, `STARTS_WITH`, `REGEX`, `UNIQUE`, `ASSISTED_CODE`), 4 severity levels.
- Assisted Code Auto-Fix with inline expression editor.
- Audit Monitor with XLSX / CSV / TXT / JSON export.
- Interactive Help Guide overlay.

---

## Versioning convention

| Change type | Increment |
|-------------|-----------|
| Bug fix with no visible UX impact | Patch `x.x.P` |
| New feature or visible improvement | Minor `x.M.0` |
| Architecture change or breaking change | Major `M.0.0` |
