(function () {
  const template = document.createElement("template");
  template.innerHTML = `
    <style>
      :host { display: block; font-family: "72", Arial, sans-serif; padding: 12px; color: #32363a; font-size: 12px; }
      .group { margin-bottom: 14px; }
      label { display: block; font-weight: 700; margin-bottom: 4px; color: #1d2d3e; }
      .hint { font-size: 11px; color: #64748b; margin-bottom: 4px; }
      input[type="text"], input[type="number"] {
        width: 100%; box-sizing: border-box; padding: 6px 8px;
        border: 1px solid #bcc8d4; border-radius: 4px; font-size: 12px;
      }
      .checkbox-row { display: flex; align-items: center; gap: 8px; margin-top: 8px; }
      hr { border: none; border-top: 1px solid #e2e8f0; margin: 14px 0; }
    </style>

    <div class="group">
      <label>Expected Sheet Name (leave blank for first sheet)</label>
      <input type="text" id="expectedSheetName" placeholder="Sheet1" />
    </div>
    <div class="group">
      <label>Required Columns (comma-separated)</label>
      <div class="hint">Defines the exact template headers</div>
      <input type="text" id="requiredColumns" />
    </div>
    <div class="group checkbox-row">
      <input type="checkbox" id="strictColumns" />
      <label for="strictColumns" style="margin:0;">Block unmapped / extra columns</label>
    </div>
    <div class="group checkbox-row">
      <input type="checkbox" id="strictOrder" />
      <label for="strictOrder" style="margin:0;">Enforce exact column order</label>
    </div>

    <hr />

    <div class="group">
      <label>Unique Key Columns (Duplicate check)</label>
      <input type="text" id="uniqueKeyColumns" />
    </div>
    <div class="group">
      <label>Numeric Columns (Amounts / Quantities)</label>
      <input type="text" id="numericColumns" placeholder="e.g., AMOUNT,QUANTITY" />
    </div>
    <div class="group">
      <label>Period Columns (YYYYMM format)</label>
      <input type="text" id="periodColumns" placeholder="e.g., CAPITALIZED,DATE" />
    </div>
    <div class="group">
      <label>Rows per Page (Pagination)</label>
      <input type="number" id="pageSize" min="10" max="500" />
    </div>
    <div class="group checkbox-row">
      <input type="checkbox" id="allowPartialUpload" />
      <label for="allowPartialUpload" style="margin:0;">Allow partial upload (skip invalid rows)</label>
    </div>
    <div class="group checkbox-row">
      <input type="checkbox" id="allowUserRules" />
      <label for="allowUserRules" style="margin:0;">Allow business users to create custom rules</label>
    </div>

    <hr />

    <div class="group">
      <label>SAC Planning Model ID ⭐</label>
      <div class="hint">Technical name of the target planning model (e.g. SAP__BR_GEN_PLANNING). Required for DIS write-back.</div>
      <input type="text" id="sacModelId" placeholder="SAP__BR_GEN_PLANNING" />
    </div>
    <div class="group">
      <label>SAC Tenant URL (optional)</label>
      <div class="hint">Leave blank — auto-detected from browser. Only set if running outside SAC context.</div>
      <input type="text" id="sacTenantUrl" placeholder="https://mytenant.eu10.sapanalytics.cloud" />
    </div>
  `;

  class ExcelUploaderBuilder extends HTMLElement {
    constructor() {
      super();
      this._shadowRoot = this.attachShadow({ mode: "open" });
      this._shadowRoot.appendChild(template.content.cloneNode(true));
      this._bindInputs();
    }

    _bindInputs() {
      ["expectedSheetName", "requiredColumns", "uniqueKeyColumns", "numericColumns", "periodColumns", "sacModelId", "sacTenantUrl"].forEach(id => {
        this._shadowRoot.getElementById(id).addEventListener("change", (e) => {
          this._emitChange(id, e.target.value);
        });
      });

      this._shadowRoot.getElementById("pageSize").addEventListener("change", (e) => {
        this._emitChange("pageSize", parseInt(e.target.value, 10) || 50);
      });

      ["strictColumns", "strictOrder", "allowPartialUpload", "allowUserRules"].forEach(id => {
        this._shadowRoot.getElementById(id).addEventListener("change", (e) => {
          this._emitChange(id, e.target.checked);
        });
      });
    }

    _emitChange(prop, value) {
      this.dispatchEvent(new CustomEvent("propertiesChanged", {
        detail: { properties: { [prop]: value } }
      }));
    }

    set expectedSheetName(v) { this._shadowRoot.getElementById("expectedSheetName").value = v || ""; }
    set requiredColumns(v) { this._shadowRoot.getElementById("requiredColumns").value = v || ""; }
    set uniqueKeyColumns(v) { this._shadowRoot.getElementById("uniqueKeyColumns").value = v || ""; }
    set numericColumns(v) { this._shadowRoot.getElementById("numericColumns").value = v || ""; }
    set periodColumns(v) { this._shadowRoot.getElementById("periodColumns").value = v || ""; }
    set pageSize(v) { this._shadowRoot.getElementById("pageSize").value = v || 50; }
    set strictColumns(v) { this._shadowRoot.getElementById("strictColumns").checked = !!v; }
    set strictOrder(v) { this._shadowRoot.getElementById("strictOrder").checked = !!v; }
    set allowPartialUpload(v) { this._shadowRoot.getElementById("allowPartialUpload").checked = !!v; }
    set allowUserRules(v) { this._shadowRoot.getElementById("allowUserRules").checked = !!v; }
    set sacModelId(v) { this._shadowRoot.getElementById("sacModelId").value = v || ""; }
    set sacTenantUrl(v) { this._shadowRoot.getElementById("sacTenantUrl").value = v || ""; }
  }

  customElements.define("custom-excel-uploader-builder", ExcelUploaderBuilder);
})();
