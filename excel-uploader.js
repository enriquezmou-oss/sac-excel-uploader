class ExcelUploaderPro extends HTMLElement {
  constructor() {
    super();
    this._shadow = this.attachShadow({ mode: 'open' });
    this._state = 'idle';
    this._activeTab = 'upload';
    this._rows = [];
    this._headers = [];
    this._structErrors = [];
    this._page = 0;
    this._filter = 'all';
    this._file = null;
    this._validatedDataJson = '[]';
    this._expectedSheetName = 'Sheet1';
    this._requiredColumns = 'ID,DESCRIPTION,H1,COMPANY,COSTCENTER,ASSET_CLASS,CAPITALIZED';
    this._uniqueKeyColumns = 'ID';
    this._numericColumns = '';
    this._customRules = [];
    this._history = [];
  }

  connectedCallback() {
    this._render();
  }

  _addRule() {
    var cols = this._requiredColumns ? this._requiredColumns.split(',') : [];
    this._customRules.push({ name: 'Nueva regla', column: cols[0] || '', operator: 'no-vacío', value: '', severity: 'error' });
    this._emitProp('customRulesJson', JSON.stringify(this._customRules));
    this._activeTab = 'rules';
    this._render();
  }

  _emitProp(prop, value) {
    this.dispatchEvent(new CustomEvent('propertiesChanged', {
      bubbles: true, composed: true,
      detail: { properties: { [prop]: value } }
    }));
  }

  _render() {
    this._shadow.innerHTML = `
      <style>
        :host { display: block; font-family: Arial, sans-serif; padding: 16px; background: #F5F6F7; border: 1px solid #D1D7DD; border-radius: 4px; }
        .title { font-size: 16px; font-weight: bold; color: #1D2D3E; margin-bottom: 12px; }
        .status { padding: 4px 8px; background: #E6F0FF; color: #0064D9; border-radius: 10px; font-size: 12px; display: inline-block; }
      </style>
      <div class="title">Excel Planning Uploader Pro</div>
      <div class="status">Estado: ${this._state}</div>
    `;
  }
}

customElements.define('custom-excel-uploader-pro', ExcelUploaderPro);
