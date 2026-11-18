# Excel Planning Uploader Pro (SAP Analytics Cloud BTP)

> Enterprise-grade custom web component for SAP Analytics Cloud (SAC) featuring multi-file consolidation, file-level business comments, interactive row addition & inline cell editing, master data cross-validation, SAP Fiori Horizon design, and full audit logging.

---

## 🚀 Architecture & Technical Stack

This widget is built as a **Native Web Component (Custom Element)** with encapsulated **Shadow DOM**, ensuring zero style leakage and seamless integration into SAC Analytic Stories.

* **Local File Parsing**: Powered by **SheetJS (xlsx.full.min.js)** via official CDN.
* **UI Design System**: Built with **SAP Fiori Horizon / SAP 72 Design Guidelines** and official SVG branding.
* **Multi-File Support**: Simultaneous selection, consolidation, and individual file queue management.
* **Smart Cross-Validation**: Real-time validation against simulated SAP SAC Master Data dictionaries (`COMPANY`, `COSTCENTER`).
* **Interactive Data Governance**: Inline cell correction, manual row addition (`+ Add New Row`), and customizable mandatory fields via `localStorage`.
* **Traceability & Comments**: File-level business notes and built-in BTP Audit Monitor tracking user sessions, logs, and timestamps.

---

## 🛠️ Main Features

1. **Multi-File Upload Wizard**: Drag-and-drop or select multiple `.xlsx` planning files simultaneously with individual deletion options.
2. **Consolidation & Business Comments**: Unified table preview with source file tracking and custom notes per file.
3. **Interactive Cell Editing & Row Addition**: Click-to-edit cells with instant re-validation and visual tooltips, plus manual row insertion.
4. **Audit Monitor**: Dedicated tab tracking import history, file summaries, and corporate traceability.

---

## 📦 Deployment & Integration

* Managed via Personal GitHub (`enriquezmou-oss/sac-excel-uploader`).
* Custom Web Component integrated into SAP Analytics Cloud stories.
