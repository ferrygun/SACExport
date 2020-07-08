(function () {
    let tmpl = document.createElement("template");
    tmpl.innerHTML = `
          <style>
              fieldset {
                  margin-bottom: 10px;
                  border: 1px solid #afafaf;
                  border-radius: 3px;
              }
              table {
                  width: 100%;
              }
              input, textarea, select {
                  font-family: "72",Arial,Helvetica,sans-serif;
                  width: 100%;
                  padding: 4px;
                  box-sizing: border-box;
                  border: 1px solid #bfbfbf;
              }
              input[type=checkbox] {
                  width: inherit;
                  margin: 6px 3px 6px 0;
                  vertical-align: middle;
              }
          </style>
          <form id="form" autocomplete="off">
            <fieldset>
              <legend>General</legend>
              <table>
                <tr>
                  <td><label for="serverURL">Server URL</label></td>
                  <td><input id="serverURL" name="serverURL" type="text"></td>
                </tr>
                <tr>
                  <td><label for="filename">Filename</label></td>
                  <td><input id="filename" name="filename" type="text"></td>
                </tr>
              </table>
            </fieldset>
            <fieldset>
              <legend>Display</legend>
              <table>
                <tr>
                  <td><label for="showComponentSelector">Show UI Selector</label></td>
                  <td><input id="showComponentSelector" name="showComponentSelector" type="checkbox"></td>
                </tr>
              </table>
            </fieldset>
            <button type="submit" hidden>Submit</button>
          </form>
        `;

    class ExportAps extends HTMLElement {
        constructor() {
            super();
            this._shadowRoot = this.attachShadow({ mode: "open" });
            this._shadowRoot.appendChild(tmpl.content.cloneNode(true));

            let form = this._shadowRoot.getElementById("form");
            form.addEventListener("submit", this._submit.bind(this));
            form.addEventListener("change", this._change.bind(this));           
        }

        connectedCallback() {
        }

        _submit(e) {
            e.preventDefault();
            let properties = {};
            for (let name of ExportAps.observedAttributes) {
                properties[name] = this[name];
            }
            this._firePropertiesChanged(properties);
            return false;
        }
        _change(e) {
            this._changeProperty(e.target.name);
        }
        _changeProperty(name) {
            let properties = {};
            properties[name] = this[name];
            this._firePropertiesChanged(properties);
        }

        _firePropertiesChanged(properties) {
            this.dispatchEvent(new CustomEvent("propertiesChanged", {
                detail: {
                    properties: properties
                }
            }));
        }

        get serverURL() {
            return this._getValue("serverURL");
        }
        set serverURL(value) {
            this._setValue("serverURL", value);
        }

        get showComponentSelector() {
            return this._getBooleanValue("showComponentSelector");
        }
        set showComponentSelector(value) {
            this._setBooleanValue("showComponentSelector", value);
        }

        get filename() {
            return this._getValue("filename");
        }
        set filename(value) {
            this._setValue("filename", value);
        }

        get metadata() {
            return this._metadata;
        }
        set metadata(value) {
            this._metadata = value;
        }

        _getValue(id) {
            return this._shadowRoot.getElementById(id).value;
        }
        _setValue(id, value) {
            this._shadowRoot.getElementById(id).value = value;
        }

        _getBooleanValue(id) {
            return this._shadowRoot.getElementById(id).checked;
        }
        _setBooleanValue(id, value) {
            this._shadowRoot.getElementById(id).checked = value;
        }

        static get observedAttributes() {
            return [
                "serverURL",
                "filename",
                "showComponentSelector",
                "metadata"
            ];
        }

        attributeChangedCallback(name, oldValue, newValue) {
            if (oldValue != newValue) {
                this[name] = newValue;
            }
        }
    }
    customElements.define("com-fd-djaja-sap-sac-export-aps", ExportAps);
})();