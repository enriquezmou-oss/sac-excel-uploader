# Changelog — Excel Planning Uploader Pro

All notable changes to this project are documented here.
Format: **[version] — date — summary**

---

## [2.6.0] — 2026-11-23

### Fix crítico DIS — flujo de un solo paso
- **Bug fix**: SAC DIS no tiene endpoint `POST/PUT /jobs/{jobId}` para subir datos — devuelve 405. El flujo correcto es enviar el CSV directamente en el cuerpo del `POST /models/{GUID}/{importType}` (`Content-Type: text/csv`).
- El widget ahora crea el job y sube los datos en una única llamada HTTP, igual que hace el uploader nativo de SAC.

### MODEL NAME — descripción amistosa
- Nuevo campo de texto de solo lectura debajo del input MODEL ID que muestra la descripción del modelo (`_sacModelDesc`) al seleccionar en el picker.
- Persiste en `localStorage` (`sac_dis_model_desc`) entre recargas.

### Model Picker rediseñado
- Nombre técnico del modelo como título principal (antes mostraba la descripción, que a veces coincidía con el GUID)
- Descripción humana (`modelDescription`) como subtítulo (texto gris 11 px)
- GUID en gris monoespacio como tercera línea auxiliar
- Modelos agrupados: **📋 Planning** (privateFactData) primero, **📊 Actual / Otros** (factData) después
- Contador `X de Y` modelos en la cabecera del dropdown
- Modelo activo marcado con ✓ y fondo azul
- Campo de búsqueda en tiempo real: filtra por nombre técnico, descripción y GUID simultáneamente
- Al seleccionar un modelo se guarda el nombre técnico en MODEL ID (no el GUID)
- El GUID queda cacheado en `_sacDISModelId` para evitar un round-trip extra en el siguiente Synchronize

---

## [2.5.0] — 2026-11-22

### Correcciones críticas DIS (Data Import Service)
- **Fix principal**: DIS requiere GUID interno (`C9qfcbu7dll3q7ao78qsq5e9h4o`), no el nombre técnico (`PLAN_CAPEX_PLANNING`). Implementado `_resolveDISModelId()` que resuelve el GUID automáticamente antes de cada importación.
- **Fix model browser**: `GET /api/v1/dataimport/models` sin parámetros `?importType=` — ese parámetro causaba 404 en este tenant.
- **Fix CSRF**: el token CSRF debe pedirse al endpoint del modelo específico (`/models/{GUID}/{importType}` con `x-csrf-token: fetch`), no al listado `/models`.
- **Fix double-load**: guard `if (typeof ExcelUploaderPro === 'undefined')` evita `SyntaxError: Identifier already declared` al subir un nuevo `webcomponent.json` mientras el anterior sigue cacheado.
- **Fix CSS**: añadido `.toast-warning { background:#E76500; }` (naranja SAP).
- Auto-detección de `importType`: intenta `privateFactData` primero, cae a `factData` si recibe error `[1704-1801-3304]` o 404.

---

## [2.4.x y anteriores] — historial previo

Versiones subidas manualmente a SAC antes del control de versiones formalizado.
Funcionalidades incluidas: Planning Upload Wizard, Validation Rules & Assisted Code, Transform & Convert, Audit Monitor, SAC Load Target con badge de estado, Synchronize con DIS write-back.

---

## Convención de versiones

| Cambio | Incremento |
|--------|-----------|
| Corrección de bug sin impacto en UX | Patch `x.x.P` |
| Nueva funcionalidad o mejora visible | Minor `x.M.0` |
| Cambio de arquitectura o breaking change | Major `M.0.0` |
