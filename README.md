# Excel Planning Uploader Pro (SAP Analytics Cloud BTP)

> Enterprise-grade custom web component for SAP Analytics Cloud (SAC) featuring multi-file consolidation, master data cross-validation, inline cell editing, SAP Fiori Horizon design, and full audit logging.

---

## 🚀 Architecture & Technical Stack

This widget is built as a **Native Web Component (Custom Element)** with encapsulated **Shadow DOM**, ensuring zero style leakage and seamless integration into SAC Analytic Stories.

* **Local File Parsing**: Powered by **SheetJS (xlsx.full.min.js)** via official CDN.
* **UI Design System**: Built with **SAP Fiori Horizon / SAP 72 Design Guidelines**.
* **Multi-File Support**: Simultaneous selection and unification of multiple Excel planning templates.
* **Smart Cross-Validation**: Real-time validation against simulated SAP SAC Master Data dictionaries (`COMPANY`, `COSTCENTER`).
* **Data Governance**: Persistent mandatory fields via `localStorage` and interactive error logging export (`.csv`).
* **Audit & Compliance**: Built-in BTP Audit Monitor tracking user sessions, rows processed, and timestamps.

---

## 🛠️ Main Features

1. **Multi-File Upload Wizard**: Drag-and-drop or select multiple `.xlsx` planning files simultaneously.
2. **Consolidation & Mapping**: Unified table preview with source file tracking (`Source File` column).
3. **Inline Cell Correction**: Click-to-edit cells with instant re-validation and visual error tooltips.
4. **Audit Monitor**: Dedicated tab tracking import history and corporate traceability.

---

## 📦 Deployment & Integration

* Managed via GitHub (`enriquezmou-oss/sac-excel-uploader`).
* Custom Web Component integrated into SAP Analytics Cloud stories.
