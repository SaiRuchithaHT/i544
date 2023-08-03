// modules are defined as an array
// [ module function, map of requires ]
//
// map of requires is short require name -> numeric require
//
// anything defined in a previous bundle is accessed via the
// orig method which is the require for previous bundles

(function (modules, entry, mainEntry, parcelRequireName, globalName) {
  /* eslint-disable no-undef */
  var globalObject =
    typeof globalThis !== 'undefined'
      ? globalThis
      : typeof self !== 'undefined'
      ? self
      : typeof window !== 'undefined'
      ? window
      : typeof global !== 'undefined'
      ? global
      : {};
  /* eslint-enable no-undef */

  // Save the require from previous bundle to this closure if any
  var previousRequire =
    typeof globalObject[parcelRequireName] === 'function' &&
    globalObject[parcelRequireName];

  var cache = previousRequire.cache || {};
  // Do not use `require` to prevent Webpack from trying to bundle this call
  var nodeRequire =
    typeof module !== 'undefined' &&
    typeof module.require === 'function' &&
    module.require.bind(module);

  function newRequire(name, jumped) {
    if (!cache[name]) {
      if (!modules[name]) {
        // if we cannot find the module within our internal map or
        // cache jump to the current global require ie. the last bundle
        // that was added to the page.
        var currentRequire =
          typeof globalObject[parcelRequireName] === 'function' &&
          globalObject[parcelRequireName];
        if (!jumped && currentRequire) {
          return currentRequire(name, true);
        }

        // If there are other bundles on this page the require from the
        // previous one is saved to 'previousRequire'. Repeat this as
        // many times as there are bundles until the module is found or
        // we exhaust the require chain.
        if (previousRequire) {
          return previousRequire(name, true);
        }

        // Try the node require function if it exists.
        if (nodeRequire && typeof name === 'string') {
          return nodeRequire(name);
        }

        var err = new Error("Cannot find module '" + name + "'");
        err.code = 'MODULE_NOT_FOUND';
        throw err;
      }

      localRequire.resolve = resolve;
      localRequire.cache = {};

      var module = (cache[name] = new newRequire.Module(name));

      modules[name][0].call(
        module.exports,
        localRequire,
        module,
        module.exports,
        this
      );
    }

    return cache[name].exports;

    function localRequire(x) {
      var res = localRequire.resolve(x);
      return res === false ? {} : newRequire(res);
    }

    function resolve(x) {
      var id = modules[name][1][x];
      return id != null ? id : x;
    }
  }

  function Module(moduleName) {
    this.id = moduleName;
    this.bundle = newRequire;
    this.exports = {};
  }

  newRequire.isParcelRequire = true;
  newRequire.Module = Module;
  newRequire.modules = modules;
  newRequire.cache = cache;
  newRequire.parent = previousRequire;
  newRequire.register = function (id, exports) {
    modules[id] = [
      function (require, module) {
        module.exports = exports;
      },
      {},
    ];
  };

  Object.defineProperty(newRequire, 'root', {
    get: function () {
      return globalObject[parcelRequireName];
    },
  });

  globalObject[parcelRequireName] = newRequire;

  for (var i = 0; i < entry.length; i++) {
    newRequire(entry[i]);
  }

  if (mainEntry) {
    // Expose entry point to Node, AMD or browser globals
    // Based on https://github.com/ForbesLindesay/umd/blob/master/template.js
    var mainExports = newRequire(mainEntry);

    // CommonJS
    if (typeof exports === 'object' && typeof module !== 'undefined') {
      module.exports = mainExports;

      // RequireJS
    } else if (typeof define === 'function' && define.amd) {
      define(function () {
        return mainExports;
      });

      // <script>
    } else if (globalName) {
      this[globalName] = mainExports;
    }
  }
})({"5uw0q":[function(require,module,exports) {
var global = arguments[3];
var HMR_HOST = null;
var HMR_PORT = null;
var HMR_SECURE = false;
var HMR_ENV_HASH = "d6ea1d42532a7575";
module.bundle.HMR_BUNDLE_ID = "40af12ebbf1e63c0";
"use strict";
/* global HMR_HOST, HMR_PORT, HMR_ENV_HASH, HMR_SECURE, chrome, browser, __parcel__import__, __parcel__importScripts__, ServiceWorkerGlobalScope */ /*::
import type {
  HMRAsset,
  HMRMessage,
} from '@parcel/reporter-dev-server/src/HMRServer.js';
interface ParcelRequire {
  (string): mixed;
  cache: {|[string]: ParcelModule|};
  hotData: {|[string]: mixed|};
  Module: any;
  parent: ?ParcelRequire;
  isParcelRequire: true;
  modules: {|[string]: [Function, {|[string]: string|}]|};
  HMR_BUNDLE_ID: string;
  root: ParcelRequire;
}
interface ParcelModule {
  hot: {|
    data: mixed,
    accept(cb: (Function) => void): void,
    dispose(cb: (mixed) => void): void,
    // accept(deps: Array<string> | string, cb: (Function) => void): void,
    // decline(): void,
    _acceptCallbacks: Array<(Function) => void>,
    _disposeCallbacks: Array<(mixed) => void>,
  |};
}
interface ExtensionContext {
  runtime: {|
    reload(): void,
    getURL(url: string): string;
    getManifest(): {manifest_version: number, ...};
  |};
}
declare var module: {bundle: ParcelRequire, ...};
declare var HMR_HOST: string;
declare var HMR_PORT: string;
declare var HMR_ENV_HASH: string;
declare var HMR_SECURE: boolean;
declare var chrome: ExtensionContext;
declare var browser: ExtensionContext;
declare var __parcel__import__: (string) => Promise<void>;
declare var __parcel__importScripts__: (string) => Promise<void>;
declare var globalThis: typeof self;
declare var ServiceWorkerGlobalScope: Object;
*/ var OVERLAY_ID = "__parcel__error__overlay__";
var OldModule = module.bundle.Module;
function Module(moduleName) {
    OldModule.call(this, moduleName);
    this.hot = {
        data: module.bundle.hotData[moduleName],
        _acceptCallbacks: [],
        _disposeCallbacks: [],
        accept: function(fn) {
            this._acceptCallbacks.push(fn || function() {});
        },
        dispose: function(fn) {
            this._disposeCallbacks.push(fn);
        }
    };
    module.bundle.hotData[moduleName] = undefined;
}
module.bundle.Module = Module;
module.bundle.hotData = {};
var checkedAssets /*: {|[string]: boolean|} */ , assetsToDispose /*: Array<[ParcelRequire, string]> */ , assetsToAccept /*: Array<[ParcelRequire, string]> */ ;
function getHostname() {
    return HMR_HOST || (location.protocol.indexOf("http") === 0 ? location.hostname : "localhost");
}
function getPort() {
    return HMR_PORT || location.port;
}
// eslint-disable-next-line no-redeclare
var parent = module.bundle.parent;
if ((!parent || !parent.isParcelRequire) && typeof WebSocket !== "undefined") {
    var hostname = getHostname();
    var port = getPort();
    var protocol = HMR_SECURE || location.protocol == "https:" && !/localhost|127.0.0.1|0.0.0.0/.test(hostname) ? "wss" : "ws";
    var ws = new WebSocket(protocol + "://" + hostname + (port ? ":" + port : "") + "/");
    // Web extension context
    var extCtx = typeof chrome === "undefined" ? typeof browser === "undefined" ? null : browser : chrome;
    // Safari doesn't support sourceURL in error stacks.
    // eval may also be disabled via CSP, so do a quick check.
    var supportsSourceURL = false;
    try {
        (0, eval)('throw new Error("test"); //# sourceURL=test.js');
    } catch (err) {
        supportsSourceURL = err.stack.includes("test.js");
    }
    // $FlowFixMe
    ws.onmessage = async function(event /*: {data: string, ...} */ ) {
        checkedAssets = {} /*: {|[string]: boolean|} */ ;
        assetsToAccept = [];
        assetsToDispose = [];
        var data /*: HMRMessage */  = JSON.parse(event.data);
        if (data.type === "update") {
            // Remove error overlay if there is one
            if (typeof document !== "undefined") removeErrorOverlay();
            let assets = data.assets.filter((asset)=>asset.envHash === HMR_ENV_HASH);
            // Handle HMR Update
            let handled = assets.every((asset)=>{
                return asset.type === "css" || asset.type === "js" && hmrAcceptCheck(module.bundle.root, asset.id, asset.depsByBundle);
            });
            if (handled) {
                console.clear();
                // Dispatch custom event so other runtimes (e.g React Refresh) are aware.
                if (typeof window !== "undefined" && typeof CustomEvent !== "undefined") window.dispatchEvent(new CustomEvent("parcelhmraccept"));
                await hmrApplyUpdates(assets);
                // Dispose all old assets.
                let processedAssets = {} /*: {|[string]: boolean|} */ ;
                for(let i = 0; i < assetsToDispose.length; i++){
                    let id = assetsToDispose[i][1];
                    if (!processedAssets[id]) {
                        hmrDispose(assetsToDispose[i][0], id);
                        processedAssets[id] = true;
                    }
                }
                // Run accept callbacks. This will also re-execute other disposed assets in topological order.
                processedAssets = {};
                for(let i = 0; i < assetsToAccept.length; i++){
                    let id = assetsToAccept[i][1];
                    if (!processedAssets[id]) {
                        hmrAccept(assetsToAccept[i][0], id);
                        processedAssets[id] = true;
                    }
                }
            } else fullReload();
        }
        if (data.type === "error") {
            // Log parcel errors to console
            for (let ansiDiagnostic of data.diagnostics.ansi){
                let stack = ansiDiagnostic.codeframe ? ansiDiagnostic.codeframe : ansiDiagnostic.stack;
                console.error("\uD83D\uDEA8 [parcel]: " + ansiDiagnostic.message + "\n" + stack + "\n\n" + ansiDiagnostic.hints.join("\n"));
            }
            if (typeof document !== "undefined") {
                // Render the fancy html overlay
                removeErrorOverlay();
                var overlay = createErrorOverlay(data.diagnostics.html);
                // $FlowFixMe
                document.body.appendChild(overlay);
            }
        }
    };
    ws.onerror = function(e) {
        console.error(e.message);
    };
    ws.onclose = function() {
        console.warn("[parcel] \uD83D\uDEA8 Connection to the HMR server was lost");
    };
}
function removeErrorOverlay() {
    var overlay = document.getElementById(OVERLAY_ID);
    if (overlay) {
        overlay.remove();
        console.log("[parcel] ✨ Error resolved");
    }
}
function createErrorOverlay(diagnostics) {
    var overlay = document.createElement("div");
    overlay.id = OVERLAY_ID;
    let errorHTML = '<div style="background: black; opacity: 0.85; font-size: 16px; color: white; position: fixed; height: 100%; width: 100%; top: 0px; left: 0px; padding: 30px; font-family: Menlo, Consolas, monospace; z-index: 9999;">';
    for (let diagnostic of diagnostics){
        let stack = diagnostic.frames.length ? diagnostic.frames.reduce((p, frame)=>{
            return `${p}
<a href="/__parcel_launch_editor?file=${encodeURIComponent(frame.location)}" style="text-decoration: underline; color: #888" onclick="fetch(this.href); return false">${frame.location}</a>
${frame.code}`;
        }, "") : diagnostic.stack;
        errorHTML += `
      <div>
        <div style="font-size: 18px; font-weight: bold; margin-top: 20px;">
          🚨 ${diagnostic.message}
        </div>
        <pre>${stack}</pre>
        <div>
          ${diagnostic.hints.map((hint)=>"<div>\uD83D\uDCA1 " + hint + "</div>").join("")}
        </div>
        ${diagnostic.documentation ? `<div>📝 <a style="color: violet" href="${diagnostic.documentation}" target="_blank">Learn more</a></div>` : ""}
      </div>
    `;
    }
    errorHTML += "</div>";
    overlay.innerHTML = errorHTML;
    return overlay;
}
function fullReload() {
    if ("reload" in location) location.reload();
    else if (extCtx && extCtx.runtime && extCtx.runtime.reload) extCtx.runtime.reload();
}
function getParents(bundle, id) /*: Array<[ParcelRequire, string]> */ {
    var modules = bundle.modules;
    if (!modules) return [];
    var parents = [];
    var k, d, dep;
    for(k in modules)for(d in modules[k][1]){
        dep = modules[k][1][d];
        if (dep === id || Array.isArray(dep) && dep[dep.length - 1] === id) parents.push([
            bundle,
            k
        ]);
    }
    if (bundle.parent) parents = parents.concat(getParents(bundle.parent, id));
    return parents;
}
function updateLink(link) {
    var href = link.getAttribute("href");
    if (!href) return;
    var newLink = link.cloneNode();
    newLink.onload = function() {
        if (link.parentNode !== null) // $FlowFixMe
        link.parentNode.removeChild(link);
    };
    newLink.setAttribute("href", // $FlowFixMe
    href.split("?")[0] + "?" + Date.now());
    // $FlowFixMe
    link.parentNode.insertBefore(newLink, link.nextSibling);
}
var cssTimeout = null;
function reloadCSS() {
    if (cssTimeout) return;
    cssTimeout = setTimeout(function() {
        var links = document.querySelectorAll('link[rel="stylesheet"]');
        for(var i = 0; i < links.length; i++){
            // $FlowFixMe[incompatible-type]
            var href /*: string */  = links[i].getAttribute("href");
            var hostname = getHostname();
            var servedFromHMRServer = hostname === "localhost" ? new RegExp("^(https?:\\/\\/(0.0.0.0|127.0.0.1)|localhost):" + getPort()).test(href) : href.indexOf(hostname + ":" + getPort());
            var absolute = /^https?:\/\//i.test(href) && href.indexOf(location.origin) !== 0 && !servedFromHMRServer;
            if (!absolute) updateLink(links[i]);
        }
        cssTimeout = null;
    }, 50);
}
function hmrDownload(asset) {
    if (asset.type === "js") {
        if (typeof document !== "undefined") {
            let script = document.createElement("script");
            script.src = asset.url + "?t=" + Date.now();
            if (asset.outputFormat === "esmodule") script.type = "module";
            return new Promise((resolve, reject)=>{
                var _document$head;
                script.onload = ()=>resolve(script);
                script.onerror = reject;
                (_document$head = document.head) === null || _document$head === void 0 || _document$head.appendChild(script);
            });
        } else if (typeof importScripts === "function") {
            // Worker scripts
            if (asset.outputFormat === "esmodule") return import(asset.url + "?t=" + Date.now());
            else return new Promise((resolve, reject)=>{
                try {
                    importScripts(asset.url + "?t=" + Date.now());
                    resolve();
                } catch (err) {
                    reject(err);
                }
            });
        }
    }
}
async function hmrApplyUpdates(assets) {
    global.parcelHotUpdate = Object.create(null);
    let scriptsToRemove;
    try {
        // If sourceURL comments aren't supported in eval, we need to load
        // the update from the dev server over HTTP so that stack traces
        // are correct in errors/logs. This is much slower than eval, so
        // we only do it if needed (currently just Safari).
        // https://bugs.webkit.org/show_bug.cgi?id=137297
        // This path is also taken if a CSP disallows eval.
        if (!supportsSourceURL) {
            let promises = assets.map((asset)=>{
                var _hmrDownload;
                return (_hmrDownload = hmrDownload(asset)) === null || _hmrDownload === void 0 ? void 0 : _hmrDownload.catch((err)=>{
                    // Web extension bugfix for Chromium
                    // https://bugs.chromium.org/p/chromium/issues/detail?id=1255412#c12
                    if (extCtx && extCtx.runtime && extCtx.runtime.getManifest().manifest_version == 3) {
                        if (typeof ServiceWorkerGlobalScope != "undefined" && global instanceof ServiceWorkerGlobalScope) {
                            extCtx.runtime.reload();
                            return;
                        }
                        asset.url = extCtx.runtime.getURL("/__parcel_hmr_proxy__?url=" + encodeURIComponent(asset.url + "?t=" + Date.now()));
                        return hmrDownload(asset);
                    }
                    throw err;
                });
            });
            scriptsToRemove = await Promise.all(promises);
        }
        assets.forEach(function(asset) {
            hmrApply(module.bundle.root, asset);
        });
    } finally{
        delete global.parcelHotUpdate;
        if (scriptsToRemove) scriptsToRemove.forEach((script)=>{
            if (script) {
                var _document$head2;
                (_document$head2 = document.head) === null || _document$head2 === void 0 || _document$head2.removeChild(script);
            }
        });
    }
}
function hmrApply(bundle /*: ParcelRequire */ , asset /*:  HMRAsset */ ) {
    var modules = bundle.modules;
    if (!modules) return;
    if (asset.type === "css") reloadCSS();
    else if (asset.type === "js") {
        let deps = asset.depsByBundle[bundle.HMR_BUNDLE_ID];
        if (deps) {
            if (modules[asset.id]) {
                // Remove dependencies that are removed and will become orphaned.
                // This is necessary so that if the asset is added back again, the cache is gone, and we prevent a full page reload.
                let oldDeps = modules[asset.id][1];
                for(let dep in oldDeps)if (!deps[dep] || deps[dep] !== oldDeps[dep]) {
                    let id = oldDeps[dep];
                    let parents = getParents(module.bundle.root, id);
                    if (parents.length === 1) hmrDelete(module.bundle.root, id);
                }
            }
            if (supportsSourceURL) // Global eval. We would use `new Function` here but browser
            // support for source maps is better with eval.
            (0, eval)(asset.output);
            // $FlowFixMe
            let fn = global.parcelHotUpdate[asset.id];
            modules[asset.id] = [
                fn,
                deps
            ];
        } else if (bundle.parent) hmrApply(bundle.parent, asset);
    }
}
function hmrDelete(bundle, id) {
    let modules = bundle.modules;
    if (!modules) return;
    if (modules[id]) {
        // Collect dependencies that will become orphaned when this module is deleted.
        let deps = modules[id][1];
        let orphans = [];
        for(let dep in deps){
            let parents = getParents(module.bundle.root, deps[dep]);
            if (parents.length === 1) orphans.push(deps[dep]);
        }
        // Delete the module. This must be done before deleting dependencies in case of circular dependencies.
        delete modules[id];
        delete bundle.cache[id];
        // Now delete the orphans.
        orphans.forEach((id)=>{
            hmrDelete(module.bundle.root, id);
        });
    } else if (bundle.parent) hmrDelete(bundle.parent, id);
}
function hmrAcceptCheck(bundle /*: ParcelRequire */ , id /*: string */ , depsByBundle /*: ?{ [string]: { [string]: string } }*/ ) {
    if (hmrAcceptCheckOne(bundle, id, depsByBundle)) return true;
    // Traverse parents breadth first. All possible ancestries must accept the HMR update, or we'll reload.
    let parents = getParents(module.bundle.root, id);
    let accepted = false;
    while(parents.length > 0){
        let v = parents.shift();
        let a = hmrAcceptCheckOne(v[0], v[1], null);
        if (a) // If this parent accepts, stop traversing upward, but still consider siblings.
        accepted = true;
        else {
            // Otherwise, queue the parents in the next level upward.
            let p = getParents(module.bundle.root, v[1]);
            if (p.length === 0) {
                // If there are no parents, then we've reached an entry without accepting. Reload.
                accepted = false;
                break;
            }
            parents.push(...p);
        }
    }
    return accepted;
}
function hmrAcceptCheckOne(bundle /*: ParcelRequire */ , id /*: string */ , depsByBundle /*: ?{ [string]: { [string]: string } }*/ ) {
    var modules = bundle.modules;
    if (!modules) return;
    if (depsByBundle && !depsByBundle[bundle.HMR_BUNDLE_ID]) {
        // If we reached the root bundle without finding where the asset should go,
        // there's nothing to do. Mark as "accepted" so we don't reload the page.
        if (!bundle.parent) return true;
        return hmrAcceptCheck(bundle.parent, id, depsByBundle);
    }
    if (checkedAssets[id]) return true;
    checkedAssets[id] = true;
    var cached = bundle.cache[id];
    assetsToDispose.push([
        bundle,
        id
    ]);
    if (!cached || cached.hot && cached.hot._acceptCallbacks.length) {
        assetsToAccept.push([
            bundle,
            id
        ]);
        return true;
    }
}
function hmrDispose(bundle /*: ParcelRequire */ , id /*: string */ ) {
    var cached = bundle.cache[id];
    bundle.hotData[id] = {};
    if (cached && cached.hot) cached.hot.data = bundle.hotData[id];
    if (cached && cached.hot && cached.hot._disposeCallbacks.length) cached.hot._disposeCallbacks.forEach(function(cb) {
        cb(bundle.hotData[id]);
    });
    delete bundle.cache[id];
}
function hmrAccept(bundle /*: ParcelRequire */ , id /*: string */ ) {
    // Execute the module.
    bundle(id);
    // Run the accept callbacks in the new version of the module.
    var cached = bundle.cache[id];
    if (cached && cached.hot && cached.hot._acceptCallbacks.length) cached.hot._acceptCallbacks.forEach(function(cb) {
        var assetsToAlsoAccept = cb(function() {
            return getParents(module.bundle.root, id);
        });
        if (assetsToAlsoAccept && assetsToAccept.length) {
            assetsToAlsoAccept.forEach(function(a) {
                hmrDispose(a[0], a[1]);
            });
            // $FlowFixMe[method-unbinding]
            assetsToAccept.push.apply(assetsToAccept, assetsToAlsoAccept);
        }
    });
}

},{}],"CnCET":[function(require,module,exports) {
var parcelHelpers = require("@parcel/transformer-js/src/esmodule-helpers.js");
var _appJs = require("./app.js");
var _appJsDefault = parcelHelpers.interopDefault(_appJs);
const WS_URL = "https://zdu.binghamton.edu:2345";
window.addEventListener("DOMContentLoaded", async ()=>{
    await (0, _appJsDefault.default)(WS_URL);
});

},{"./app.js":"6wtUX","@parcel/transformer-js/src/esmodule-helpers.js":"gkKU3"}],"6wtUX":[function(require,module,exports) {
var parcelHelpers = require("@parcel/transformer-js/src/esmodule-helpers.js");
parcelHelpers.defineInteropFlag(exports);
parcelHelpers.export(exports, "default", ()=>makeApp);
var _cs544JsUtils = require("cs544-js-utils");
var _utilsJs = require("./utils.js");
var _ssWsJs = require("./ss-ws.js");
var _ssWsJsDefault = parcelHelpers.interopDefault(_ssWsJs);
var _spreadsheetJs = require("./spreadsheet.js");
var _spreadsheetJsDefault = parcelHelpers.interopDefault(_spreadsheetJs);
async function makeApp(wsUrl) {
    makeTopLevelUI(wsUrl);
    setupLoadFormHandler();
}
function setupLoadFormHandler() {
    const errors = new (0, _utilsJs.Errors)();
    const wsUrlInput = document.querySelector("#ws-url");
    const ssNameInput = document.querySelector("#ss-name");
    let ws;
    let ssName;
    const ssForm = document.querySelector("#ss-form");
    ssForm.addEventListener("submit", async (ev)=>{
        ev.preventDefault();
        errors.clear();
        const wsUrl = wsUrlInput.value.trim();
        const ssName = ssNameInput.value.trim();
        if (wsUrl.length === 0 || ssName.length === 0) {
            const msg = "both the Web Services Url and Spreadsheet Name must be specified";
            errors.display([
                new (0, _cs544JsUtils.Err)(msg, {
                    code: "BAD_REQ"
                })
            ]);
        } else {
            const ws = (0, _ssWsJsDefault.default).make(wsUrl);
            await (0, _spreadsheetJsDefault.default)(ws, ssName);
        }
    });
}
/** Add UI corresponding to following HTML structure to #app

  <form class="form" id="ss-form">

    <label for="ws-url">Web Services Url</label>
    <input name="ws-url" id="ws-url">

    <label for="ss-name">Spreadsheet Name</label>
    <input name="ss-name" id="ss-name">

    <label></label>
    <button type="submit">Load Spreadsheet</button>
    
  </form>

  <ul class="error" id="errors"></ul>
    
  <div id="ss">
    <!-- innerHTML of this div replaced by spreadsheet table -->
  </div>
*/ function makeTopLevelUI(wsUrl) {
    function makeLoadForm(wsUrl) {
        const form = (0, _utilsJs.makeElement)("form", {
            class: "form",
            id: "ss-form"
        });
        form.append((0, _utilsJs.makeElement)("label", {
            for: "ws-url"
        }, "Web Services URL"));
        form.append((0, _utilsJs.makeElement)("input", {
            name: "ws-url",
            id: "ws-url",
            value: wsUrl
        }));
        form.append((0, _utilsJs.makeElement)("label", {
            for: "ss-name"
        }, "Spreadsheet Name"));
        form.append((0, _utilsJs.makeElement)("input", {
            name: "ss-name",
            id: "ss-name"
        }));
        form.append((0, _utilsJs.makeElement)("label"));
        form.append((0, _utilsJs.makeElement)("button", {
            type: "submit"
        }, "Load Spreadsheet"));
        return form;
    }
    const app = document.querySelector("#app");
    app.append(makeLoadForm(wsUrl));
    app.append((0, _utilsJs.makeElement)("ul", {
        class: "error",
        id: "errors"
    }));
    //spreadsheet table should be rendered within this div
    app.append((0, _utilsJs.makeElement)("div", {
        id: "ss"
    }));
}

},{"cs544-js-utils":"8WQYV","./utils.js":"3cCYB","./ss-ws.js":"a5j3O","./spreadsheet.js":"9ZFep","@parcel/transformer-js/src/esmodule-helpers.js":"gkKU3"}],"8WQYV":[function(require,module,exports) {
var parcelHelpers = require("@parcel/transformer-js/src/esmodule-helpers.js");
parcelHelpers.defineInteropFlag(exports);
var _errorsJs = require("./lib/errors.js"); //export { default as makeValidator} from './lib/validator.js';
parcelHelpers.exportAll(_errorsJs, exports);

},{"./lib/errors.js":"aGjnO","@parcel/transformer-js/src/esmodule-helpers.js":"gkKU3"}],"aGjnO":[function(require,module,exports) {
// Immutable API
/** throw exception with msg and args; use when impossible conditions occur */ var parcelHelpers = require("@parcel/transformer-js/src/esmodule-helpers.js");
parcelHelpers.defineInteropFlag(exports);
parcelHelpers.export(exports, "panic", ()=>panic);
parcelHelpers.export(exports, "Err", ()=>Err);
/** A Result is either a success result identified by isOk=true,
 *  or an error result identified by isOk=false.  A success
 *  result has the success value in its 'val' property; an
 *  error result will have one or more 'Err' objects in its
 *  'errors' property.
 */ parcelHelpers.export(exports, "OkResult", ()=>OkResult);
parcelHelpers.export(exports, "ErrResult", ()=>ErrResult);
/** factory function for a success result */ parcelHelpers.export(exports, "okResult", ()=>okResult);
parcelHelpers.export(exports, "VOID_RESULT", ()=>VOID_RESULT);
/** factory function for an error result initialized to contain
 *  a single error as per arg0, args.
 *    errResult(msg: string, code?: string, widget?: string)
 *    errResult(msg: string, options: ErrOptions)
 *    errResult(err: Err)
 *    errResult(err: ErrResult, options: ErrOptions)
 *    errResult(errObj: object, options: ErrOptions)
 */ parcelHelpers.export(exports, "errResult", ()=>errResult) /*
//demo program

function safeDiv(num: number, denom: number) : Result<number> {
  if (denom === 0) return errResult('zero denominator');
  return okResult(num/denom);
}

function demo(result: Result<number>) : Result<string> {
  if (!result.isOk) return result as Result<string>;
  const v = result.val + 1;
  return result.chain((val: number) => okResult('x'.repeat(v*val)))
               .chain((str: string) => okResult(str + 'aaa'));
}

console.log(safeDiv(1, 0));
console.log(safeDiv(1, 2));
console.log(demo(errResult('some error', 'ERR_CODE')));
console.log(demo(okResult(2)));
*/ ;
function panic(msg, ...args) {
    throw new Error(msg + args.map((a)=>JSON.stringify(a)).join(", "));
}
const DEFAULT_ERR_CODE = "UNKNOWN";
class Err {
    message;
    options;
    constructor(message, options){
        this.message = message;
        this.options = options;
    }
}
/** Convenience error building function.  Possible arguments:
 *     error(msg: string, code?: string, widget?: string)
 *     error(msg: string, options: ErrOptions)
 *     error(err: Err)
 *     error(err: Error, options?: ErrOptions)
 *     error(errObj: object, options?: ErrOptions)
 */ function error(arg0, ...args) {
    let options = {
        code: DEFAULT_ERR_CODE
    };
    if (typeof arg0 === "string") {
        const msg = arg0;
        if (args.length === 0) return new Err(msg, {
            code: DEFAULT_ERR_CODE
        });
        else if (args.length === 1 && typeof args[0] === "object") return new Err(msg, {
            code: DEFAULT_ERR_CODE,
            ...args[0]
        });
        else if (args.length === 1 && typeof args[0] === "string") return new Err(msg, {
            code: args[0]
        });
        else if (args.length === 2 && typeof args[0] === "string" && typeof args[1] === "string") return new Err(msg, {
            code: args[0],
            widget: args[1]
        });
        else panic(`bad error args`, [
            arg0,
            ...args
        ]);
    } else if (arg0 instanceof Err) return arg0;
    else if (arg0 instanceof Error) return new Err(arg0.message, args.length > 0 ? args[0] : {
        code: DEFAULT_ERR_CODE
    });
    else if (typeof arg0 === "object") return new Err(arg0.toString(), args.length > 0 ? args[0] : {
        code: DEFAULT_ERR_CODE
    });
    else panic(`bad error args`, [
        arg0,
        ...args
    ]);
}
class OkResult {
    isOk = true;
    val;
    constructor(v){
        this.val = v;
    }
    /** return result of applying fn on val */ chain(fn) {
        return fn(this.val);
    }
}
class ErrResult {
    isOk = false;
    errors;
    constructor(errors = []){
        this.errors = errors;
    }
    /** Possible arguments
     *   addError(ErrResult errResult)
     *   addError(msg: string, code?: string, widget?: string)
     *   addError(msg: string, options: ErrOptions)
     *   addError(err: Err)
     *   addError(err: Error, options?: ErrOptions)
     *   addError(errObj: object, options?: ErrOptions)
     */ addError(arg0, ...args) {
        const errors = arg0 instanceof ErrResult ? arg0.errors : [
            error(arg0, ...args)
        ];
        return new ErrResult(this.errors.concat(errors));
    }
    /** ignore fn, simply returning this error result */ chain(_fn) {
        return this;
    }
}
function okResult(v) {
    return new OkResult(v);
}
const VOID_RESULT = okResult(undefined);
function errResult(arg0, ...args) {
    return new ErrResult().addError(arg0, ...args);
}

},{"@parcel/transformer-js/src/esmodule-helpers.js":"gkKU3"}],"gkKU3":[function(require,module,exports) {
exports.interopDefault = function(a) {
    return a && a.__esModule ? a : {
        default: a
    };
};
exports.defineInteropFlag = function(a) {
    Object.defineProperty(a, "__esModule", {
        value: true
    });
};
exports.exportAll = function(source, dest) {
    Object.keys(source).forEach(function(key) {
        if (key === "default" || key === "__esModule" || dest.hasOwnProperty(key)) return;
        Object.defineProperty(dest, key, {
            enumerable: true,
            get: function() {
                return source[key];
            }
        });
    });
    return dest;
};
exports.export = function(dest, destName, get) {
    Object.defineProperty(dest, destName, {
        enumerable: true,
        get: get
    });
};

},{}],"3cCYB":[function(require,module,exports) {
var parcelHelpers = require("@parcel/transformer-js/src/esmodule-helpers.js");
parcelHelpers.defineInteropFlag(exports);
parcelHelpers.export(exports, "Errors", ()=>Errors);
/** Return a new DOM element with specified tagName, attributes
 *  given by object attrs and initial contained text.
 *  Note that .append(TextOrElement) can be called on the returned
 *  element to append string text or a new DOM element to it.
 */ parcelHelpers.export(exports, "makeElement", ()=>makeElement);
class Errors {
    errors;
    constructor(){
        this.errors = document.querySelector("#errors");
    }
    display(errors) {
        errors.forEach((err)=>{
            this.errors.append(makeElement("li", {}, err.message));
        });
    }
    clear() {
        this.errors.innerHTML = "";
    }
}
function makeElement(tagName, attrs = {}, text = "") {
    const element = document.createElement(tagName);
    for (const [k, v] of Object.entries(attrs))element.setAttribute(k, v);
    if (text.length > 0) element.append(text);
    return element;
}

},{"@parcel/transformer-js/src/esmodule-helpers.js":"gkKU3"}],"a5j3O":[function(require,module,exports) {
var parcelHelpers = require("@parcel/transformer-js/src/esmodule-helpers.js");
parcelHelpers.defineInteropFlag(exports);
var _cs544JsUtils = require("cs544-js-utils");
class SpreadsheetWs {
    apiUrl;
    constructor(url){
        this.apiUrl = `${url}/api`;
    }
    static make(url) {
        return new SpreadsheetWs(url);
    }
    /** return { expr, value } object for cell cellId in
     *  spreadsheet ssName.
     */ async query(ssName, cellId) {
        const url = makeURL(`${this.apiUrl}/${ssName}/${cellId}`);
        return await doFetchJson("GET", url);
    }
    /** remove formula for cell cellId in spreadsheet ssName.
     *  Return Updates object mapping cellId's to the updated value.
     */ async remove(ssName, cellId) {
        const url = makeURL(`${this.apiUrl}/${ssName}/${cellId}`);
        return await doFetchJson("DELETE", url);
    }
    /** copy formula from cell srcCellId in spreadsheet ssName to
     *  cell destCellId.  Update relative references in the formula.
     *  If the formula is empty, then delete cell destCellId.
     *  Return Updates object mapping cellId's to the updated value.
     */ async copy(ssName, destCellId, srcCellId) {
        const url = makeURL(`${this.apiUrl}/${ssName}/${destCellId}`, {
            srcCellId
        });
        return await doFetchJson("PATCH", url);
    }
    /** set formula for cell cellId in spreadsheet ssName to expr.
     *  Return Updates object mapping cellId's to the updated value.
     */ async evaluate(ssName, cellId, expr) {
        const url = makeURL(`${this.apiUrl}/${ssName}/${cellId}`, {
            expr
        });
        return await doFetchJson("PATCH", url);
    }
    /** return list of [cellId, expr] pairs for spreadsheet ssName */ async dump(ssName) {
        const url = makeURL(`${this.apiUrl}/${ssName}`);
        return await doFetchJson("GET", url);
    }
    /** return list of [cellId, expr, value] triples for spreadsheet ssName */ async dumpWithValues(ssName) {
        const url = makeURL(`${this.apiUrl}/${ssName}`, {
            doValues: "true"
        });
        return await doFetchJson("GET", url);
    }
    /** load spreadsheet ssName with dump of list of [ cellId, expr ]. */ async load(ssName, dump) {
        const url = makeURL(`${this.apiUrl}/${ssName}`);
        return await doFetchJson("PUT", url, dump);
    }
    /** clear out all contents of spreadsheet ssName */ async clear(ssName) {
        const url = makeURL(`${this.apiUrl}/${ssName}`);
        return await doFetchJson("DELETE", url);
    }
}
exports.default = SpreadsheetWs;
/** A utility function used to extend a url with properly encoded
 *  query parameters
 */ function makeURL(url, queryParams = {}) {
    const urlObj = new URL(url);
    Object.entries(queryParams).forEach(([k, v])=>{
        urlObj.searchParams.append(k, v.toString());
    });
    return urlObj;
}
/** Return a Result for dispatching HTTP method to url.  If jsonBody
 *  is specified, then it should be sent as the JSONrequest body
 *  (along with a suitable MIME-type).
 *
 *  This function should convert the response envelope used for
 *  the web services into a Result.  Specifically:
 *
 *    + The response should return an error Result if there is a fetch
 *      error or if the response JSON contains errors.
 *
 *    + If there are no errors then the function should return the
 *      response result within an ok Result.
 */ async function doFetchJson(method, url, jsonBody) {
    try {
        const response = await fetch(url.toString(), {
            method,
            headers: {
                "Content-Type": "application/json"
            },
            body: jsonBody ? JSON.stringify(jsonBody) : undefined
        });
        const data = await response.json();
        if (response.ok) return (0, _cs544JsUtils.okResult)(data.result);
        else return (0, _cs544JsUtils.errResult)(data.errors[0].message);
    } catch (err) {
        return (0, _cs544JsUtils.errResult)([
            {
                message: err.toString()
            }
        ]);
    }
}

},{"cs544-js-utils":"8WQYV","@parcel/transformer-js/src/esmodule-helpers.js":"gkKU3"}],"9ZFep":[function(require,module,exports) {
var parcelHelpers = require("@parcel/transformer-js/src/esmodule-helpers.js");
parcelHelpers.defineInteropFlag(exports);
parcelHelpers.export(exports, "default", ()=>make);
var _utilsJs = require("./utils.js");
const [N_ROWS, N_COLS] = [
    10,
    10
];
async function make(ws, ssName) {
    return await Spreadsheet.make(ws, ssName);
}
class Spreadsheet {
    ws;
    ssName;
    errors;
    currentFocusedCell = null;
    copySourceCell = null;
    errorOccurred = false;
    constructor(ws, ssName){
        this.ws = ws;
        this.ssName = ssName;
        this.errors = new (0, _utilsJs.Errors)();
        this.currentFocusedCell = null;
        this.copySourceCell = null;
        this.makeEmptySS();
        this.addListeners();
    //TODO: initialize added instance variables
    }
    static async make(ws, ssName) {
        const ss = new Spreadsheet(ws, ssName);
        await ss.load();
        return ss;
    }
    /** add listeners for different events on table elements */ addListeners() {
        const clearButton = document.querySelector("#clear");
        clearButton?.addEventListener("click", this.clearSpreadsheet);
        const cells = document.querySelectorAll(".cell");
        cells.forEach((cell)=>{
            cell.addEventListener("focusin", this.focusCell);
            cell.addEventListener("focusout", this.blurCell);
            cell.addEventListener("copy", this.copyCell);
            cell.addEventListener("paste", this.pasteCell);
        });
    }
    /** listener for a click event on #clear button */ clearSpreadsheet = async (ev)=>{
        const result = await this.ws.clear(this.ssName);
        if (result.isOk) this.clearCellData();
        else this.errors.display(result.errors);
    };
    /** clear cell data when the "Clear" button is clicked */ clearCellData() {
        const cells = document.querySelectorAll(".cell");
        cells.forEach((cell)=>{
            cell.textContent = "";
            cell.removeAttribute("data-expr");
            cell.removeAttribute("data-value");
        });
    }
    /** listener for the focus event on a spreadsheet data cell */ focusCell = (ev)=>{
        const cell = ev.target;
        this.currentFocusedCell = cell.id;
        const expr = cell.getAttribute("data-expr");
        cell.textContent = expr || cell.getAttribute("data-value") || "";
        this.errors.clear();
    };
    /** listener for a blur event on a spreadsheet data cell */ blurCell = async (ev)=>{
        const cell = ev.target;
        const cellId = cell.id;
        const content = cell.textContent?.trim() || "";
        const prevExpr = cell.getAttribute("data-expr");
        const prevValue = cell.getAttribute("data-value");
        cell.dataset.expr = content;
        if (content === "") {
            const result = await this.ws.remove(this.ssName, cellId);
            if (result.isOk) {
                cell.textContent = "";
                cell.removeAttribute("data-expr");
                cell.removeAttribute("data-value");
                this.updateCells(result.val);
            } else {
                this.errors.display(result.errors);
                this.errorOccurred = true;
                cell.setAttribute("data-prev-expr", prevExpr || "");
                cell.setAttribute("data-prev-value", prevValue || "");
            }
        } else {
            const result = await this.ws.evaluate(this.ssName, cellId, content);
            if (result.isOk) {
                cell.textContent = cell.dataset.value || "";
                this.updateCells(result.val);
            } else {
                this.errors.display(result.errors);
                this.errorOccurred = true;
                cell.setAttribute("data-prev-expr", prevExpr || "");
                cell.setAttribute("data-prev-value", prevValue || "");
            }
        }
        this.currentFocusedCell = null;
        if (this.errorOccurred) {
            this.resetCellContent(cell);
            this.errorOccurred = false;
        }
    };
    /** reset cell content to its previous value (if error occurred during evaluation) */ resetCellContent(cell) {
        const prevExpr = cell.getAttribute("data-prev-expr");
        const prevValue = cell.getAttribute("data-prev-value");
        cell.textContent = prevValue || prevExpr || "";
        cell.setAttribute("data-expr", prevExpr || "");
        cell.setAttribute("data-value", prevValue || "");
    }
    /** listener for a copy event on a spreadsheet data cell */ copyCell = (ev)=>{
        const cell = ev.target;
        const currentCopySourceCell = document.querySelector(".is-copy-source");
        if (currentCopySourceCell) currentCopySourceCell.classList.remove("is-copy-source");
        this.copySourceCell = cell.id;
        cell.classList.add("is-copy-source");
    };
    /** listener for a paste event on a spreadsheet data cell */ pasteCell = async (ev)=>{
        ev.preventDefault();
        const destinationcell = ev.target;
        const destinationCellId = destinationcell.id;
        // Check if we have a source cell from which content is copied
        if (this.copySourceCell) {
            const sourceCellId = this.copySourceCell;
            const result = await this.ws.copy(this.ssName, destinationCellId, sourceCellId);
            if (result.isOk) {
                const updates = result.val;
                const res2 = await this.ws.query(this.ssName, destinationCellId);
                if (res2.isOk) {
                    const destExpressionVal = res2.val;
                    const destExpression = destExpressionVal.expr;
                    destinationcell.setAttribute("data-exp", destExpression);
                    this.updateCells(result.val);
                    destinationcell.textContent = destExpression;
                }
            } else this.errors.display(result.errors);
            // Clear the copy source cell
            const copySourceCellElement = document.getElementById(this.copySourceCell);
            if (copySourceCellElement) {
                copySourceCellElement.classList.remove("is-copy-source");
                this.copySourceCell = null;
            }
        }
    };
    /** update cells with new values after successful evaluation */ updateCells(updates) {
        const cells = document.querySelectorAll(".cell");
        cells.forEach((cell)=>{
            const cellId = cell.id;
            if (updates.hasOwnProperty(cellId) && this.currentFocusedCell != cellId) {
                const value = updates[cellId];
                cell.setAttribute("data-value", value.toString());
                cell.textContent = value.toString();
            }
        });
    }
    /** load existing data from the spreadsheet */ async load() {
        const result = await this.ws.dumpWithValues(this.ssName);
        if (result.isOk) {
            const data = result.val;
            const cells = document.querySelectorAll(".cell");
            cells.forEach((cell)=>{
                const cellId = cell.id;
                const cellData = data.find((item)=>item[0] === cellId);
                if (cellData) {
                    const [_, expr, value] = cellData;
                    cell.textContent = value.toString();
                    cell.setAttribute("data-expr", expr);
                    cell.setAttribute("data-value", value.toString());
                }
            });
        } else this.errors.display(result.errors);
    }
    makeEmptySS() {
        const ssDiv = document.querySelector("#ss");
        ssDiv.innerHTML = "";
        const ssTable = (0, _utilsJs.makeElement)("table");
        const header = (0, _utilsJs.makeElement)("tr");
        const clearCell = (0, _utilsJs.makeElement)("td");
        const clear = (0, _utilsJs.makeElement)("button", {
            id: "clear",
            type: "button"
        }, "Clear");
        clearCell.append(clear);
        header.append(clearCell);
        const A = "A".charCodeAt(0);
        for(let i = 0; i < N_COLS; i++)header.append((0, _utilsJs.makeElement)("th", {}, String.fromCharCode(A + i)));
        ssTable.append(header);
        for(let i = 0; i < N_ROWS; i++){
            const row = (0, _utilsJs.makeElement)("tr");
            row.append((0, _utilsJs.makeElement)("th", {}, (i + 1).toString()));
            const a = "a".charCodeAt(0);
            for(let j = 0; j < N_COLS; j++){
                const colId = String.fromCharCode(a + j);
                const id = colId + (i + 1);
                const cell = (0, _utilsJs.makeElement)("td", {
                    id,
                    class: "cell",
                    contentEditable: "true"
                });
                row.append(cell);
            }
            ssTable.append(row);
        }
        ssDiv.append(ssTable);
    }
}

},{"./utils.js":"3cCYB","@parcel/transformer-js/src/esmodule-helpers.js":"gkKU3"}]},["5uw0q","CnCET"], "CnCET", "parcelRequireeeb6")

//# sourceMappingURL=index.bf1e63c0.js.map
