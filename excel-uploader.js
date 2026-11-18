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
  }

  connectedCallback() {
    this._render();
  }

  _addRule() {
    this._render();
  }

  _emitProp(prop, value) {
    this.dispatchEvent(new CustomEvent('propertiesChanged', {
      bubbles: true, composed: true,
      detail: { properties: { [prop]: value } }
    }));
  }

  _render() {
    this._shadow.innerHTML = `<style>:host { display: block; padding: 16px; font-family: Arial; }</style><div>Excel Planning Uploader Pro Loaded</div>`;
  }
}

customElements.define('custom-excel-uploader-pro', ExcelUploaderPro);
