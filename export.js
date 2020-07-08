(function () {
    let tmpl = document.createElement("template");
    tmpl.innerHTML = `
      <style>
      </style>
      <div id="export_div" name="export_div" class="">
         <slot name="export_button"></slot>
         <form id="form" method="post" accept-charset="utf-8" action="">
            <input id="export_settings_json" name="export_settings_json" type="hidden">
        </form>
      </div>
    `;

    class Export extends HTMLElement {

        constructor() {
            super();

            this._shadowRoot = this.attachShadow({ mode: "open" });
            this._shadowRoot.appendChild(tmpl.content.cloneNode(true));

            this._id = createGuid();

            this._shadowRoot.querySelector("#export_div").id = this._id + "_export_div";
            this._shadowRoot.querySelector("#form").id = this._id + "_form";

            this.settings = this._shadowRoot.querySelector("#export_settings_json");
            this.settings.id = this._id + "_export_settings_json";

            this._cPPT_text = "HTML";
            
            this._cExport_text = "Export";
            this._cExport_icon = "sap-icon://download";

            this._showIcons = true;
            this._showTexts = false;
            this._showComponentSelector = false;
            this._enablePPT = true;

            this._export_settings = {};
            this._export_settings.filename = "";
            this._export_settings.ppt_exclude = "";          
            this._export_settings.server_urls = "";
          
            this._updateSettings();

            this._renderExportButton();
        }

        connectedCallback() {
            // try detect components in edit mode
            try {
                if (window.commonApp) {
                    let outlineContainer = commonApp.getShell().findElements(true, ele => ele.hasStyleClass && ele.hasStyleClass("sapAppBuildingOutline"))[0]; // sId: "__container0"
                    if (outlineContainer && outlineContainer.getReactProps) {
                        let parseReactState = state => {
                            let components = {};

                            let globalState = state.globalState;
                            let instances = globalState.instances;
                            let app = instances.app["[{\"app\":\"MAIN_APPLICATION\"}]"];
                            let names = app.names;

                            for (let key in names) {
                                let name = names[key];

                                let obj = JSON.parse(key).pop();
                                let type = Object.keys(obj)[0];
                                let id = obj[type];

                                components[id] = {
                                    type: type,
                                    name: name
                                };
                            }

                            let metadata = JSON.stringify({
                                components: components,
                                vars: app.globalVars
                            });

                            if (metadata != this.metadata) {
                                this.metadata = metadata;

                                this.dispatchEvent(new CustomEvent("propertiesChanged", {
                                    detail: {
                                        properties: {
                                            metadata: metadata
                                        }
                                    }
                                }));
                            }
                        };

                        let subscribeReactStore = store => {
                            this._subscription = store.subscribe({
                                effect: state => {
                                    parseReactState(state);
                                    return { result: 1 };
                                }
                            });
                        };

                        let props = outlineContainer.getReactProps();
                        if (props) {
                            subscribeReactStore(props.store);
                        } else {
                            let oldRenderReactComponent = outlineContainer.renderReactComponent;
                            outlineContainer.renderReactComponent = e => {
                                let props = outlineContainer.getReactProps();
                                subscribeReactStore(props.store);

                                oldRenderReactComponent.call(outlineContainer, e);
                            }
                        }
                    }
                }
            } catch (e) {
            }
        }

        disconnectedCallback() {
            if (this._subscription) { // react store subscription
                this._subscription();
                this._subscription = null;
            }
        }

        onCustomWidgetBeforeUpdate(changedProperties) {
            if ("designMode" in changedProperties) {
                this._designMode = changedProperties["designMode"];
            }
        }

        onCustomWidgetAfterUpdate(changedProperties) {
            this._pptMenuItem.setVisible(this.enablePpt);
            this._pptMenuItem.setText(this.showTexts ? this._cPPT_text : null);
            this._pptMenuItem.setIcon(this.showIcons ? this._cPPT_icon : null);

            this._exportButton.setVisible(this.showTexts || this.showIcons);
            this._exportButton.setText(this.showTexts ? this._cExport_text : null);
            this._exportButton.setIcon(this.showIcons ? this._cExport_icon : null);
            if (this._designMode) {
                this._exportButton.setEnabled(false);
            }
        }

        _renderExportButton() {
            let menu = new sap.m.Menu({
                title: this._cExport_text,
                itemSelected: oEvent => {
                    let oItem = oEvent.getParameter("item");
                    if (!this.showComponentSelector) {
                        this.doExport(oItem.getKey());
                    } else {
                        let ltab = new sap.m.IconTabBar({
                            expandable: false
                        });

                        let lcomponent_box;
                        if (this.showComponentSelector) {
                            lcomponent_box = new sap.ui.layout.form.SimpleForm({
                                layout: sap.ui.layout.form.SimpleFormLayout.ResponsiveGridLayout,
                                columnsM: 2,
                                columnsL: 4
                            });

                            let components = this.metadata ? JSON.parse(this.metadata)["components"] : {};

                            if (this["_initialVisibleComponents" + oItem.getKey()] == null) {
                                this["_initialVisibleComponents" + oItem.getKey()] = this[oItem.getKey().toLowerCase() + "SelectedWidgets"] ? JSON.parse(this[oItem.getKey().toLowerCase() + "SelectedWidgets"]) : [];
                            }

                            if (this["_initialVisibleComponents" + oItem.getKey()].length == 0) {
                                let linitial = [];
                                for (let componentId in components) {
                                    let component = components[componentId];
                                    var lcomp = {};
                                    lcomp.component = component.name;
                                    lcomp.isExcluded = false;
                                    lcomp.type = component.type;
                                    linitial.push(lcomp);
                                }
                                this[oItem.getKey().toLowerCase() + "SelectedWidgets"] = JSON.stringify(linitial);
                            }
                            for (let componentId in components) {
                                let component = components[componentId];

                                if (component.type === "sdk_com_fd_djaja_sap_sac_export__0") {
                                    continue;
                                }

                                if (this["_initialVisibleComponents" + oItem.getKey()].length == 0 || this["_initialVisibleComponents" + oItem.getKey()].some(v => v.component == component.name && !v.isExcluded)) {
                                    let ltext = component.name.replace(/_/g, " ");

                                    lcomponent_box.addContent(new sap.m.CheckBox({
                                        id: component.name,
                                        text: ltext,
                                        selected: true,
                                        select: oEvent => {
                                            let visibleComponents = [];
                                            let objIndex = -1;

                                            if (this[oItem.getKey().toLowerCase() + "SelectedWidgets"] != "") {
                                                visibleComponents = JSON.parse(this[oItem.getKey().toLowerCase() + "SelectedWidgets"]);
                                                objIndex = visibleComponents.findIndex(v => v.component == oEvent.getParameter("id"));
                                            }
                                            if (objIndex > -1) {
                                                visibleComponents[objIndex].isExcluded = !oEvent.getParameter("selected");
                                            } else {

                                                visibleComponents.push({
                                                    component: oEvent.getParameter("id"),
                                                    isExcluded: !oEvent.getParameter("selected"),
                                                    type: oEvent.getParameter("type")
                                                });
                                            }

                                            console.log("----------------visibleComponents--------------------");
                                            console.log(visibleComponents);

                                            this[oItem.getKey().toLowerCase() + "SelectedWidgets"] = JSON.stringify(visibleComponents);
                                        }
                                    }));
                                }
                            }

                            ltab.addItem(new sap.m.IconTabFilter({
                                key: "components",
                                text: "Select UI",
                                icon: "",
                                content: [
                                    lcomponent_box
                                ]
                            }));
                        }

                        let dialog = new sap.m.Dialog({
                            title: "Export",
                            contentWidth: "500px",
                            contentHeight: "400px",
                            draggable: true,
                            resizable: true,
                            content: [
                                ltab
                            ],
                            beginButton: new sap.m.Button({
                                text: "Submit",
                                press: () => {
                                    this._updateSettings();
                                    this.doExport(oItem.getKey());
                                    dialog.close();
                                }
                            }),
                            endButton: new sap.m.Button({
                                text: "Cancel",
                                press: () => {
                                    dialog.close();
                                }
                            }),
                            afterClose: () => {
                                if (lcomponent_box != null) { lcomponent_box.destroy(); }
                                ltab.destroy();
                                dialog.destroy();
                            }
                        });

                        dialog.open();
                    }
                }
            });

            this._pptMenuItem = new sap.m.MenuItem({ key: "PPT" });
            menu.addItem(this._pptMenuItem);

            let buttonSlot = document.createElement("div");
            buttonSlot.slot = "export_button";
            this.appendChild(buttonSlot);

            this._exportButton = new sap.m.MenuButton({ menu: menu, visible: false });
            this._exportButton.placeAt(buttonSlot);
        }

        // DISPLAY
        getButtonIconVisible() {
            return this.showIcons;
        }
        setButtonIconVisible(value) {
            this._setValue("showIcons", value);
        }

        get showIcons() {
            return this._showIcons;
        }
        set showIcons(value) {
            this._showIcons = value;
        }

        getButtonTextVisible() {
            return this.showTexts;
        }
        setButtonTextVisible(value) {
            this._setValue("showTexts", value);
        }

        get showTexts() {
            return this._showTexts;
        }
        set showTexts(value) {
            this._showTexts = value;
        }

        get showComponentSelector() {
            return this._showComponentSelector;
        }
        set showComponentSelector(value) {
            this._showComponentSelector = value;
        }

        get enablePpt() {
            return this._enablePPT;
        }
        set enablePpt(value) {
            this._enablePPT = value;
        }


        // SETTINGS
        getServerUrl() {
            return this.serverURL;
        }
        setServerUrl(value) {
            this._setValue("serverURL", value);
        }

        get serverURL() {
            return this._export_settings.server_urls;
        }
        set serverURL(value) {
            this._export_settings.server_urls = value;
            this._updateSettings();
        }

        getFilename() {
            return this.filename;
        }
        setFilename(value) {
            this._setValue("filename", value);
        }

        get filename() {
            return this._export_settings.filename;
        }
        set filename(value) {
            this._export_settings.filename = value;
            this._updateSettings();
        }

        get pptSelectedWidgets() {
            return this._export_settings.ppt_exclude;
        }
        set pptSelectedWidgets(value) {
            this._export_settings.ppt_exclude = value;
            this._updateSettings();
        }

        get metadata() {
            return this._export_settings.metadata;
        }
        set metadata(value) {
            this._export_settings.metadata = value;
            this._updateSettings();
        }

        static get observedAttributes() {
            return [
                "metadata"
            ];
        }

        // METHODS
        _updateSettings() {
            this.settings.value = JSON.stringify(this._export_settings);
        }

        _setValue(name, value) {
            this[name] = value;

            let properties = {};
            properties[name] = this[name];
            this.dispatchEvent(new CustomEvent("propertiesChanged", {
                detail: {
                    properties: properties
                }
            }));
        }

        doExport(format, overrideSettings) {
            let settings = JSON.parse(JSON.stringify(this._export_settings));

            setTimeout(() => {
                this._doExport(format, settings, overrideSettings);
            }, 200);
        }

        _doExport(format, settings, overrideSettings) {
            if (this._designMode) {
                return false;
            }

            if (overrideSettings) {
                let set = JSON.parse(overrideSettings);
                set.forEach(s => {
                    settings[s.name] = s.value;
                });
            }

            settings.format = format;
            settings.URL = location.protocol + "//" + location.host;
            settings.dashboard = location.href;
            settings.title = document.title;
            settings.cookie = document.cookie;
            settings.scroll_width = document.body.scrollWidth;
            settings.scroll_height = document.body.scrollHeight;

            // try detect runtime settings
            if (window.sap && sap.fpa && sap.fpa.ui && sap.fpa.ui.infra) {
                if (sap.fpa.ui.infra.common) {
                    let context = sap.fpa.ui.infra.common.getContext();

                    let app = context.getAppArgument();
                    settings.appid = app.appId;

                    let user = context.getUser();
                    settings.sac_user = user.getUsername();

                    if (settings.lng == "") {
                        settings.lng = context.getLanguage();
                    }
                }
                if (sap.fpa.ui.infra.service && sap.fpa.ui.infra.service.AjaxHelper) {
                    settings.tenant_URL = sap.fpa.ui.infra.service.AjaxHelper.getTenantUrl(false); // true for PUBLIC_FQDN
                }
            }

            this.dispatchEvent(new CustomEvent("onStart", {
                detail: {
                    settings: settings
                }
            }));

            let sendHtml = true;
            if (settings.application_array && settings.oauth) {
                sendHtml = false;
            }
            if (sendHtml) {
                // add settings to html so they can be serialized
                // NOTE: this is not "promise" save!
                this.settings.value = JSON.stringify(settings);

                getHtml(settings).then(html => {
                    this._updateSettings(); // reset settings
                    this._createExportForm(settings, html);

                }, reason => {
                    console.error("Error in getHtml:", reason);
                });
            } else {
                this._createExportForm(settings, null);
            }
        }

        _createExportForm(settings, content) {
            this.dispatchEvent(new CustomEvent("onSend", {
                detail: {
                    settings: settings
                }
            }));

            let form = document.createElement("form");

            let settingsEl = form.appendChild(document.createElement("input"));
            settingsEl.name = "export_settings_json";
            settingsEl.type = "hidden";
            settingsEl.value = JSON.stringify(settings);

            if (content) {
                let contentEl = form.appendChild(document.createElement("input"));
                contentEl.name = "export_content";
                contentEl.type = "hidden";
                contentEl.value = content;
            }

            let host = settings.server_urls;
            let url = host + "";

            this._submitExport(host, url, form, settings);
        }

        _submitExport(host, exportUrl, form, settings) {
            this._serviceMessage = "";
            var that = this;

            // handle response types
            let callback = (error, filename, blob) => {
                if (error) {
                    this._serviceMessage = error;
                    this.dispatchEvent(new CustomEvent("onError", {
                        detail: {
                            error: error,
                            settings: settings
                        }
                    }));

                    console.error("Export failed:", error);
                } else if (filename) {
                    if (filename.indexOf("E:") === 0) {
                        callback(new Error(filename)); // error...
                        return;
                    }

                    this._serviceMessage = "Export has been produced";
                    this.dispatchEvent(new CustomEvent("onReturn", {
                        detail: {
                            filename: filename,
                            settings: settings
                        }
                    }));

                    if (blob) { // download blob
                        
                        let downloadUrl = URL.createObjectURL(blob);                        
                        let a = document.createElement("a");
                        a.download = filename;
                        a.href = downloadUrl;
                        document.body.appendChild(a);
                        a.click();

                        setTimeout(() => {
                            document.body.removeChild(a);
                            URL.revokeObjectURL(downloadUrl);
                        }, 0);                        

                    } 
                }
            };

            if (exportUrl.indexOf(location.protocol) == 0 || exportUrl.indexOf("http:") == 0) { // same protocol => use fetch?
                fetch(exportUrl, {
                    method: "POST",
                    mode: "cors",
                    body: new FormData(form),
                    headers: {
                        "X-Requested-With": "XMLHttpRequest"
                    }
                }).then(response => {
                    if (response.ok) {
                        return response.blob().then(blob => {
                            console.log(that.getFilename());
                            callback(null, that.getFilename(), blob);
                        });
                        return response.text().then(text => {
                            callback(null, text);
                        });
                    } else {
                        throw new Error(response.status + ": " + response.statusText);
                    }
                }).catch(reason => {
                    callback(reason);
                });
            } else { // use form with blank target...
                form.action = exportUrl;
                form.target = "_blank";
                form.method = "POST";
                form.acceptCharset = "utf-8";
                this._shadowRoot.appendChild(form);

                form.submit();
                form.remove();

                callback(null, "I:Export running in separate tab");
            }
        }

    }
    customElements.define("com-fd-djaja-sap-sac-export", Export);

    // PUBLIC API
    window.getHtml = getHtml;

    // UTILS
    const cssUrlRegExp = /url\(["']?(.*?)["']?\)/i;
    const contentDispositionFilenameRegExp = /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/i;
    const startsWithHttpRegExp = /^http/i;

    function createGuid() {
        return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, c => {
            let r = Math.random() * 16 | 0, v = c === "x" ? r : (r & 0x3 | 0x8);
            return v.toString(16);
        });
    }

    function getHtml(settings) {
        let html = [];
        let promises = [];
        cloneNode(document.documentElement, html, promises, settings || {});
        return Promise.all(promises).then(() => {
            if (document.doctype && typeof XMLSerializer != "undefined") { // <!DOCTYPE html>
                html.unshift(new XMLSerializer().serializeToString(document.doctype));
            }

            return html.join("");
        });
    }

    function cloneNode(node, html, promises, settings) {
        if (node.nodeType == 8) return; // COMMENT
        if (node.tagName == "SCRIPT") return; // SCRIPT

        if (node.nodeType == 3) { // TEXT
            html.push(escapeText(node.nodeValue));
            return;
        }

        let name = node.localName;
        let content = null;
        let attributes = Object.create(null);
        for (let i = 0; i < node.attributes.length; i++) {
            let attribute = node.attributes[i];
            attributes[attribute.name] = attribute.value;
        }


        switch (node.tagName) {
            case "INPUT":
                attributes["value"] = node.value;
                delete attributes["checked"];
                if (node.checked) {
                    attributes["checked"] = "checked";
                }
                break;
            case "OPTION":
                delete attributes["selected"];
                if (node.selected) {
                    attributes["selected"] = "selected";
                }
                break;
            case "TEXTAREA":
                content = node.value;
                break;
            case "CANVAS":
                name = "img";
                attributes["src"] = node.toDataURL("image/png");
                break;
            case "IMG":
                if (node.src && !node.src.includes("data:")) {
                    attributes["src"] = getUrlAsDataUrl(node.src).then(d => d, () => node.src);
                }
                break;
            case "LINK":
                if (node.rel == "preload") {
                    return "";
                }
            // fallthrough
            case "STYLE":
                let sheet = node.sheet;
                if (sheet) {
                    let shadowHost = null;
                    let parent = node.parentNode;
                    while (parent) {
                        if (parent.host) {
                            shadowHost = parent.host;
                            break;
                        }
                        parent = parent.parentNode;
                    }

                    if (shadowHost || settings.parse_css) {
                        if (sheet.href) { // download external stylesheets
                            name = "style";
                            attributes = { "type": "text/css" };
                            content = fetch(sheet.href).then(r => r.text()).then(t => {
                                let style = document.createElement("style");
                                style.type = "text/css";
                                style.appendChild(document.createTextNode(t));
                                document.body.appendChild(style);
                                style.sheet.disabled = true;
                                return getCssText(style.sheet, sheet.href, shadowHost).then(r => {
                                    document.body.removeChild(style);
                                    return r;
                                });
                            }, reason => {
                                return "";
                            });
                        } else {
                            content = getCssText(sheet, document.baseURI, shadowHost);
                        }
                    }
                }
                break;
        }

        if (settings.parse_css) {
            if (attributes["style"]) {
                let style = attributes["style"];
                if (style.includes("url") && !style.includes("data:")) {
                    let url = cssUrlRegExp.exec(style)[1];
                    if (url) {
                        attributes["style"] = getUrlAsDataUrl(toAbsoluteUrl(document.baseURI, url)).then(d => style.replace(url, d), () => style);
                    }
                }
            }
        }

        html.push("<");
        html.push(name);
        for (let name in attributes) {
            let value = attributes[name];

            html.push(" ");
            html.push(name);
            html.push("=\"");
            if (value.then) {
                let index = html.length;
                html.push(""); // placeholder
                promises.push(value.then(v => html[index] = escapeAttributeValue(v)));
            } else {
                html.push(escapeAttributeValue(value));
            }
            html.push("\"");
        }
        html.push(">");
        let isEmpty = true;
        if (content) {
            if (content.then) {
                let index = html.length;
                html.push(""); // placeholder
                promises.push(content.then(c => html[index] = escapeText(c)));
            } else {
                html.push(escapeText(content));
            }
            isEmpty = false;
        } else {
            let child = node.firstChild;
            if ((!child || node.tagName == "COM-BIEXCELLENCE-OPENBI-SAP-SAC-EXPORT") && node.shadowRoot) { // shadowRoot
                child = node.shadowRoot.firstChild;
            }
            while (child) {
                html.push(cloneNode(child, html, promises, settings));
                child = child.nextSibling;
                isEmpty = false;
            }
        }
        if (isEmpty && !new RegExp("</" + node.tagName + ">$", "i").test(node.outerHTML)) {
            // no end tag
        } else {
            html.push("</");
            html.push(name);
            html.push(">");
        }
    }

    function getCssText(sheet, baseUrl, shadowHost) {
        return parseCssRules(sheet.rules, baseUrl, shadowHost);
    }

    function parseCssRules(rules, baseUrl, shadowHost) {
        let promises = [];
        let css = [];

        for (let i = 0; i < rules.length; i++) {
            let rule = rules[i];

            if (rule.type == CSSRule.MEDIA_RULE) { // media query
                css.push("@media ");
                css.push(rule.conditionText);
                css.push("{");

                let index = css.length;
                css.push(""); // placeholder
                promises.push(parseCssRules(rule.cssRules, baseUrl, shadowHost).then(c => css[index] = c));

                css.push("}");
            } else if (rule.type == CSSRule.IMPORT_RULE) { // @import
                let index = css.length;
                css.push(""); // placeholder
                promises.push(getCssText(rule.styleSheet, baseUrl, shadowHost).then(c => css[index] = c));
            } else if (rule.type == CSSRule.STYLE_RULE) {
                if (shadowHost) { // prefix with shadow host name...
                    css.push(shadowHost.localName);
                    css.push(" ");
                    css.push(rule.selectorText.split(",").join("," + shadowHost.localName));
                } else {
                    css.push(rule.selectorText);
                }
                css.push(" {");
                for (let j = 0; j < rule.style.length; j++) {
                    let name = rule.style[j]
                    let value = rule.style[name];
                    css.push(name);
                    css.push(":");
                    if (name.startsWith("background") && value && value.includes("url") && !value.includes("data:")) {
                        let url = cssUrlRegExp.exec(value)[1];
                        if (url) {
                            let index = css.length;
                            css.push(""); // placeholder
                            promises.push(getUrlAsDataUrl(toAbsoluteUrl(baseUrl, url)).then(d => css[index] = "url(" + d + ")", () => css[index] = value));
                        }
                    }
                    css.push(value);
                    css.push(";");
                }
                css.push("}");
            } else if (rule.type == CSSRule.FONT_FACE_RULE) {
                css.push("@font-face {");
                for (let j = 0; j < rule.style.length; j++) {
                    let name = rule.style[j]
                    let value = rule.style[name];
                    css.push(name);
                    css.push(":");
                    if (name == "src" && value && value.includes("url") && !value.includes("data:")) {
                        let url = cssUrlRegExp.exec(value)[1];
                        if (url) {
                            let index = css.length;
                            css.push(""); // placeholder
                            promises.push(getUrlAsDataUrl(toAbsoluteUrl(baseUrl, url)).then(d => css[index] = "url(" + d + ")", () => css[index] = value));
                        }
                    } else {
                        css.push(value);
                    }
                    css.push(";");
                }
                css.push("}");
            } else {
                css.push(rule.cssText);
            }
        }

        return Promise.all(promises).then(() => css.join(""));
    }

    function toAbsoluteUrl(baseUrl, url) {
        if (startsWithHttpRegExp.test(url) || url.startsWith("//")) { // already absolute
            return url;
        }

        let index = baseUrl.lastIndexOf("/");
        if (index > 8) {
            baseUrl = baseUrl.substring(0, index);
        }
        baseUrl += "/";

        if (url.startsWith("/")) {
            return baseUrl.substring(0, baseUrl.indexOf("/", 8)) + url;
        }
        return baseUrl + url;
    }

    function getUrlAsDataUrl(url) {
        return fetch(url).then(r => r.blob()).then(b => {
            return new Promise((resolve, reject) => {
                let fileReader = new FileReader();
                fileReader.onload = () => {
                    resolve(fileReader.result);
                };
                fileReader.onerror = () => {
                    reject(new Error("Failed to convert URL to data URL: " + url));
                };
                fileReader.readAsDataURL(b);
            });
        });
    }

    function escapeText(text) {
        return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
    }

    function escapeAttributeValue(value) {
        return value.replace(/"/g, "&quot;");
    }
})();