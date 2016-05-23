var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var definePropertySave;
var isDefinePropertySupported = false;
try {
    Object.defineProperty({}, "myTestProperty", {
        get: function () {
            return this.desc;
        },
        set: function (val) {
            this.desc = val;
        }
    });
    isDefinePropertySupported = true;
}
catch (e) {
    definePropertySave = Object.defineProperty;
    Object.defineProperty = function () { };
}
var OfficeExt;
(function (OfficeExt) {
    var MicrosoftAjaxFactory = (function () {
        function MicrosoftAjaxFactory() {
        }
        MicrosoftAjaxFactory.prototype.isMsAjaxLoaded = function () {
            if (typeof (Sys) !== 'undefined' && typeof (Type) !== 'undefined' &&
                Sys.StringBuilder && typeof (Sys.StringBuilder) === "function" &&
                Type.registerNamespace && typeof (Type.registerNamespace) === "function" &&
                Type.registerClass && typeof (Type.registerClass) === "function" &&
                typeof (Function._validateParams) === "function") {
                return true;
            }
            else {
                return false;
            }
        };
        MicrosoftAjaxFactory.prototype.loadMsAjaxFull = function (callback) {
            var msAjaxCDNPath = (window.location.protocol.toLowerCase() === 'https:' ? 'https:' : 'http:') + '//ajax.aspnetcdn.com/ajax/3.5/MicrosoftAjax.js';
            OSF.OUtil.loadScript(msAjaxCDNPath, callback);
        };
        Object.defineProperty(MicrosoftAjaxFactory.prototype, "msAjaxError", {
            get: function () {
                if (this._msAjaxError == null && this.isMsAjaxLoaded()) {
                    this._msAjaxError = Error;
                }
                return this._msAjaxError;
            },
            set: function (errorClass) {
                this._msAjaxError = errorClass;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(MicrosoftAjaxFactory.prototype, "msAjaxSerializer", {
            get: function () {
                if (this._msAjaxSerializer == null && this.isMsAjaxLoaded()) {
                    this._msAjaxSerializer = Sys.Serialization.JavaScriptSerializer;
                }
                return this._msAjaxSerializer;
            },
            set: function (serializerClass) {
                this._msAjaxSerializer = serializerClass;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(MicrosoftAjaxFactory.prototype, "msAjaxString", {
            get: function () {
                if (this._msAjaxString == null && this.isMsAjaxLoaded()) {
                    this._msAjaxSerializer = String;
                }
                return this._msAjaxString;
            },
            set: function (stringClass) {
                this._msAjaxString = stringClass;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(MicrosoftAjaxFactory.prototype, "msAjaxDebug", {
            get: function () {
                if (this._msAjaxDebug == null && this.isMsAjaxLoaded()) {
                    this._msAjaxDebug = Sys.Debug;
                }
                return this._msAjaxDebug;
            },
            set: function (debugClass) {
                this._msAjaxDebug = debugClass;
            },
            enumerable: true,
            configurable: true
        });
        return MicrosoftAjaxFactory;
    })();
    OfficeExt.MicrosoftAjaxFactory = MicrosoftAjaxFactory;
})(OfficeExt || (OfficeExt = {}));
var OsfMsAjaxFactory = new OfficeExt.MicrosoftAjaxFactory();
var OSF = OSF || {};
var OfficeExt;
(function (OfficeExt) {
    var SafeStorage = (function () {
        function SafeStorage(_internalStorage) {
            this._internalStorage = _internalStorage;
        }
        SafeStorage.prototype.getItem = function (key) {
            try {
                return this._internalStorage && this._internalStorage.getItem(key);
            }
            catch (e) {
                return null;
            }
        };
        SafeStorage.prototype.setItem = function (key, data) {
            try {
                this._internalStorage && this._internalStorage.setItem(key, data);
            }
            catch (e) {
            }
        };
        SafeStorage.prototype.clear = function () {
            try {
                this._internalStorage && this._internalStorage.clear();
            }
            catch (e) {
            }
        };
        SafeStorage.prototype.removeItem = function (key) {
            try {
                this._internalStorage && this._internalStorage.removeItem(key);
            }
            catch (e) {
            }
        };
        SafeStorage.prototype.getKeysWithPrefix = function (keyPrefix) {
            var keyList = [];
            try {
                var len = this._internalStorage && this._internalStorage.length || 0;
                for (var i = 0; i < len; i++) {
                    var key = this._internalStorage.key(i);
                    if (key.indexOf(keyPrefix) === 0) {
                        keyList.push(key);
                    }
                }
            }
            catch (e) {
            }
            return keyList;
        };
        return SafeStorage;
    })();
    OfficeExt.SafeStorage = SafeStorage;
})(OfficeExt || (OfficeExt = {}));
OSF.OUtil = (function () {
    var _uniqueId = -1;
    var _xdmInfoKey = '&_xdm_Info=';
    var _serializerVersionKey = '&_serializer_version=';
    var _xdmSessionKeyPrefix = '_xdm_';
    var _serializerVersionKeyPrefix = '_serializer_version=';
    var _fragmentSeparator = '#';
    var _loadedScripts = {};
    var _defaultScriptLoadingTimeout = 30000;
    var _safeSessionStorage = null;
    var _safeLocalStorage = null;
    var _rndentropy = new Date().getTime();
    function _random() {
        var nextrand = 0x7fffffff * (Math.random());
        nextrand ^= _rndentropy ^ ((new Date().getMilliseconds()) << Math.floor(Math.random() * (31 - 10)));
        return nextrand.toString(16);
    }
    ;
    function _getSessionStorage() {
        if (!_safeSessionStorage) {
            try {
                var sessionStorage = window.sessionStorage;
            }
            catch (ex) {
                sessionStorage = null;
            }
            _safeSessionStorage = new OfficeExt.SafeStorage(sessionStorage);
        }
        return _safeSessionStorage;
    }
    ;
    return {
        set_entropy: function OSF_OUtil$set_entropy(entropy) {
            if (typeof entropy == "string") {
                for (var i = 0; i < entropy.length; i += 4) {
                    var temp = 0;
                    for (var j = 0; j < 4 && i + j < entropy.length; j++) {
                        temp = (temp << 8) + entropy.charCodeAt(i + j);
                    }
                    _rndentropy ^= temp;
                }
            }
            else if (typeof entropy == "number") {
                _rndentropy ^= entropy;
            }
            else {
                _rndentropy ^= 0x7fffffff * Math.random();
            }
            _rndentropy &= 0x7fffffff;
        },
        extend: function OSF_OUtil$extend(child, parent) {
            var F = function () { };
            F.prototype = parent.prototype;
            child.prototype = new F();
            child.prototype.constructor = child;
            child.uber = parent.prototype;
            if (parent.prototype.constructor === Object.prototype.constructor) {
                parent.prototype.constructor = parent;
            }
        },
        setNamespace: function OSF_OUtil$setNamespace(name, parent) {
            if (parent && name && !parent[name]) {
                parent[name] = {};
            }
        },
        unsetNamespace: function OSF_OUtil$unsetNamespace(name, parent) {
            if (parent && name && parent[name]) {
                delete parent[name];
            }
        },
        loadScript: function OSF_OUtil$loadScript(url, callback, timeoutInMs) {
            if (url && callback) {
                var doc = window.document;
                var _loadedScriptEntry = _loadedScripts[url];
                if (!_loadedScriptEntry) {
                    var script = doc.createElement("script");
                    script.type = "text/javascript";
                    _loadedScriptEntry = { loaded: false, pendingCallbacks: [callback], timer: null };
                    _loadedScripts[url] = _loadedScriptEntry;
                    var onLoadCallback = function OSF_OUtil_loadScript$onLoadCallback() {
                        if (_loadedScriptEntry.timer != null) {
                            clearTimeout(_loadedScriptEntry.timer);
                            delete _loadedScriptEntry.timer;
                        }
                        _loadedScriptEntry.loaded = true;
                        var pendingCallbackCount = _loadedScriptEntry.pendingCallbacks.length;
                        for (var i = 0; i < pendingCallbackCount; i++) {
                            var currentCallback = _loadedScriptEntry.pendingCallbacks.shift();
                            currentCallback();
                        }
                    };
                    var onLoadError = function OSF_OUtil_loadScript$onLoadError() {
                        delete _loadedScripts[url];
                        if (_loadedScriptEntry.timer != null) {
                            clearTimeout(_loadedScriptEntry.timer);
                            delete _loadedScriptEntry.timer;
                        }
                        var pendingCallbackCount = _loadedScriptEntry.pendingCallbacks.length;
                        for (var i = 0; i < pendingCallbackCount; i++) {
                            var currentCallback = _loadedScriptEntry.pendingCallbacks.shift();
                            currentCallback();
                        }
                    };
                    if (script.readyState) {
                        script.onreadystatechange = function () {
                            if (script.readyState == "loaded" || script.readyState == "complete") {
                                script.onreadystatechange = null;
                                onLoadCallback();
                            }
                        };
                    }
                    else {
                        script.onload = onLoadCallback;
                    }
                    script.onerror = onLoadError;
                    timeoutInMs = timeoutInMs || _defaultScriptLoadingTimeout;
                    _loadedScriptEntry.timer = setTimeout(onLoadError, timeoutInMs);
                    script.src = url;
                    doc.getElementsByTagName("head")[0].appendChild(script);
                }
                else if (_loadedScriptEntry.loaded) {
                    callback();
                }
                else {
                    _loadedScriptEntry.pendingCallbacks.push(callback);
                }
            }
        },
        loadCSS: function OSF_OUtil$loadCSS(url) {
            if (url) {
                var doc = window.document;
                var link = doc.createElement("link");
                link.type = "text/css";
                link.rel = "stylesheet";
                link.href = url;
                doc.getElementsByTagName("head")[0].appendChild(link);
            }
        },
        parseEnum: function OSF_OUtil$parseEnum(str, enumObject) {
            var parsed = enumObject[str.trim()];
            if (typeof (parsed) == 'undefined') {
                OsfMsAjaxFactory.msAjaxDebug.trace("invalid enumeration string:" + str);
                throw OsfMsAjaxFactory.msAjaxError.argument("str");
            }
            return parsed;
        },
        delayExecutionAndCache: function OSF_OUtil$delayExecutionAndCache() {
            var obj = { calc: arguments[0] };
            return function () {
                if (obj.calc) {
                    obj.val = obj.calc.apply(this, arguments);
                    delete obj.calc;
                }
                return obj.val;
            };
        },
        getUniqueId: function OSF_OUtil$getUniqueId() {
            _uniqueId = _uniqueId + 1;
            return _uniqueId.toString();
        },
        formatString: function OSF_OUtil$formatString() {
            var args = arguments;
            var source = args[0];
            return source.replace(/{(\d+)}/gm, function (match, number) {
                var index = parseInt(number, 10) + 1;
                return args[index] === undefined ? '{' + number + '}' : args[index];
            });
        },
        generateConversationId: function OSF_OUtil$generateConversationId() {
            return [_random(), _random(), (new Date()).getTime().toString()].join('_');
        },
        getFrameNameAndConversationId: function OSF_OUtil$getFrameNameAndConversationId(cacheKey, frame) {
            var frameName = _xdmSessionKeyPrefix + cacheKey + this.generateConversationId();
            frame.setAttribute("name", frameName);
            return this.generateConversationId();
        },
        addXdmInfoAsHash: function OSF_OUtil$addXdmInfoAsHash(url, xdmInfoValue) {
            return OSF.OUtil.addInfoAsHash(url, _xdmInfoKey, xdmInfoValue);
        },
        addSerializerVersionAsHash: function OSF_OUtil$addSerializerVersionAsHash(url, serializerVersion) {
            return OSF.OUtil.addInfoAsHash(url, _serializerVersionKey, serializerVersion);
        },
        addInfoAsHash: function OSF_OUtil$addInfoAsHash(url, keyName, infoValue) {
            url = url.trim() || '';
            var urlParts = url.split(_fragmentSeparator);
            var urlWithoutFragment = urlParts.shift();
            var fragment = urlParts.join(_fragmentSeparator);
            return [urlWithoutFragment, _fragmentSeparator, fragment, keyName, infoValue].join('');
        },
        parseXdmInfo: function OSF_OUtil$parseXdmInfo(skipSessionStorage) {
            return OSF.OUtil.parseXdmInfoWithGivenFragment(skipSessionStorage, window.location.hash);
        },
        parseXdmInfoWithGivenFragment: function OSF_OUtil$parseXdmInfoWithGivenFragment(skipSessionStorage, fragment) {
            return OSF.OUtil.parseInfoWithGivenFragment(_xdmInfoKey, _xdmSessionKeyPrefix, skipSessionStorage, fragment);
        },
        parseSerializerVersion: function OSF_OUtil$parseSerializerVersion(skipSessionStorage) {
            return OSF.OUtil.parseSerializerVersionWithGivenFragment(skipSessionStorage, window.location.hash);
        },
        parseSerializerVersionWithGivenFragment: function OSF_OUtil$parseSerializerVersionWithGivenFragment(skipSessionStorage, fragment) {
            return parseInt(OSF.OUtil.parseInfoWithGivenFragment(_serializerVersionKey, _serializerVersionKeyPrefix, skipSessionStorage, fragment));
        },
        parseInfoWithGivenFragment: function OSF_OUtil$parseInfoWithGivenFragment(infoKey, infoKeyPrefix, skipSessionStorage, fragment) {
            var fragmentParts = fragment.split(infoKey);
            var xdmInfoValue = fragmentParts.length > 1 ? fragmentParts[fragmentParts.length - 1] : null;
            var osfSessionStorage = _getSessionStorage();
            if (!skipSessionStorage && osfSessionStorage) {
                var sessionKeyStart = window.name.indexOf(infoKeyPrefix);
                if (sessionKeyStart > -1) {
                    var sessionKeyEnd = window.name.indexOf(";", sessionKeyStart);
                    if (sessionKeyEnd == -1) {
                        sessionKeyEnd = window.name.length;
                    }
                    var sessionKey = window.name.substring(sessionKeyStart, sessionKeyEnd);
                    if (xdmInfoValue) {
                        osfSessionStorage.setItem(sessionKey, xdmInfoValue);
                    }
                    else {
                        xdmInfoValue = osfSessionStorage.getItem(sessionKey);
                    }
                }
            }
            return xdmInfoValue;
        },
        getConversationId: function OSF_OUtil$getConversationId() {
            var searchString = window.location.search;
            var conversationId = null;
            if (searchString) {
                var index = searchString.indexOf("&");
                conversationId = index > 0 ? searchString.substring(1, index) : searchString.substr(1);
                if (conversationId && conversationId.charAt(conversationId.length - 1) === '=') {
                    conversationId = conversationId.substring(0, conversationId.length - 1);
                    if (conversationId) {
                        conversationId = decodeURIComponent(conversationId);
                    }
                }
            }
            return conversationId;
        },
        getInfoItems: function OSF_OUtil$getInfoItems(strInfo) {
            var items = strInfo.split("$");
            if (typeof items[1] == "undefined") {
                items = strInfo.split("|");
            }
            return items;
        },
        getConversationUrl: function OSF_OUtil$getConversationUrl() {
            var conversationUrl = '';
            var xdmInfoValue = OSF.OUtil.parseXdmInfo(true);
            if (xdmInfoValue) {
                var items = OSF.OUtil.getInfoItems(xdmInfoValue);
                if (items != undefined && items.length >= 3) {
                    conversationUrl = items[2];
                }
            }
            return conversationUrl;
        },
        validateParamObject: function OSF_OUtil$validateParamObject(params, expectedProperties, callback) {
            var e = Function._validateParams(arguments, [{ name: "params", type: Object, mayBeNull: false },
                { name: "expectedProperties", type: Object, mayBeNull: false },
                { name: "callback", type: Function, mayBeNull: true }
            ]);
            if (e)
                throw e;
            for (var p in expectedProperties) {
                e = Function._validateParameter(params[p], expectedProperties[p], p);
                if (e)
                    throw e;
            }
        },
        writeProfilerMark: function OSF_OUtil$writeProfilerMark(text) {
            if (window.msWriteProfilerMark) {
                window.msWriteProfilerMark(text);
                OsfMsAjaxFactory.msAjaxDebug.trace(text);
            }
        },
        outputDebug: function OSF_OUtil$outputDebug(text) {
            if (typeof (OsfMsAjaxFactory) !== 'undefined' && OsfMsAjaxFactory.msAjaxDebug && OsfMsAjaxFactory.msAjaxDebug.trace) {
                OsfMsAjaxFactory.msAjaxDebug.trace(text);
            }
        },
        defineNondefaultProperty: function OSF_OUtil$defineNondefaultProperty(obj, prop, descriptor, attributes) {
            descriptor = descriptor || {};
            for (var nd in attributes) {
                var attribute = attributes[nd];
                if (descriptor[attribute] == undefined) {
                    descriptor[attribute] = true;
                }
            }
            Object.defineProperty(obj, prop, descriptor);
            return obj;
        },
        defineNondefaultProperties: function OSF_OUtil$defineNondefaultProperties(obj, descriptors, attributes) {
            descriptors = descriptors || {};
            for (var prop in descriptors) {
                OSF.OUtil.defineNondefaultProperty(obj, prop, descriptors[prop], attributes);
            }
            return obj;
        },
        defineEnumerableProperty: function OSF_OUtil$defineEnumerableProperty(obj, prop, descriptor) {
            return OSF.OUtil.defineNondefaultProperty(obj, prop, descriptor, ["enumerable"]);
        },
        defineEnumerableProperties: function OSF_OUtil$defineEnumerableProperties(obj, descriptors) {
            return OSF.OUtil.defineNondefaultProperties(obj, descriptors, ["enumerable"]);
        },
        defineMutableProperty: function OSF_OUtil$defineMutableProperty(obj, prop, descriptor) {
            return OSF.OUtil.defineNondefaultProperty(obj, prop, descriptor, ["writable", "enumerable", "configurable"]);
        },
        defineMutableProperties: function OSF_OUtil$defineMutableProperties(obj, descriptors) {
            return OSF.OUtil.defineNondefaultProperties(obj, descriptors, ["writable", "enumerable", "configurable"]);
        },
        finalizeProperties: function OSF_OUtil$finalizeProperties(obj, descriptor) {
            descriptor = descriptor || {};
            var props = Object.getOwnPropertyNames(obj);
            var propsLength = props.length;
            for (var i = 0; i < propsLength; i++) {
                var prop = props[i];
                var desc = Object.getOwnPropertyDescriptor(obj, prop);
                if (!desc.get && !desc.set) {
                    desc.writable = descriptor.writable || false;
                }
                desc.configurable = descriptor.configurable || false;
                desc.enumerable = descriptor.enumerable || true;
                Object.defineProperty(obj, prop, desc);
            }
            return obj;
        },
        mapList: function OSF_OUtil$MapList(list, mapFunction) {
            var ret = [];
            if (list) {
                for (var item in list) {
                    ret.push(mapFunction(list[item]));
                }
            }
            return ret;
        },
        listContainsKey: function OSF_OUtil$listContainsKey(list, key) {
            for (var item in list) {
                if (key == item) {
                    return true;
                }
            }
            return false;
        },
        listContainsValue: function OSF_OUtil$listContainsElement(list, value) {
            for (var item in list) {
                if (value == list[item]) {
                    return true;
                }
            }
            return false;
        },
        augmentList: function OSF_OUtil$augmentList(list, addenda) {
            var add = list.push ? function (key, value) { list.push(value); } : function (key, value) { list[key] = value; };
            for (var key in addenda) {
                add(key, addenda[key]);
            }
        },
        redefineList: function OSF_Outil$redefineList(oldList, newList) {
            for (var key1 in oldList) {
                delete oldList[key1];
            }
            for (var key2 in newList) {
                oldList[key2] = newList[key2];
            }
        },
        isArray: function OSF_OUtil$isArray(obj) {
            return Object.prototype.toString.apply(obj) === "[object Array]";
        },
        isFunction: function OSF_OUtil$isFunction(obj) {
            return Object.prototype.toString.apply(obj) === "[object Function]";
        },
        isDate: function OSF_OUtil$isDate(obj) {
            return Object.prototype.toString.apply(obj) === "[object Date]";
        },
        addEventListener: function OSF_OUtil$addEventListener(element, eventName, listener) {
            if (element.addEventListener) {
                element.addEventListener(eventName, listener, false);
            }
            else if ((Sys.Browser.agent === Sys.Browser.InternetExplorer) && element.attachEvent) {
                element.attachEvent("on" + eventName, listener);
            }
            else {
                element["on" + eventName] = listener;
            }
        },
        removeEventListener: function OSF_OUtil$removeEventListener(element, eventName, listener) {
            if (element.removeEventListener) {
                element.removeEventListener(eventName, listener, false);
            }
            else if ((Sys.Browser.agent === Sys.Browser.InternetExplorer) && element.detachEvent) {
                element.detachEvent("on" + eventName, listener);
            }
            else {
                element["on" + eventName] = null;
            }
        },
        getCookieValue: function OSF_OUtil$getCookieValue(cookieName) {
            var tmpCookieString = RegExp(cookieName + "[^;]+").exec(document.cookie);
            return tmpCookieString.toString().replace(/^[^=]+./, "");
        },
        xhrGet: function OSF_OUtil$xhrGet(url, onSuccess, onError) {
            var xmlhttp;
            try {
                xmlhttp = new XMLHttpRequest();
                xmlhttp.onreadystatechange = function () {
                    if (xmlhttp.readyState == 4) {
                        if (xmlhttp.status == 200) {
                            onSuccess(xmlhttp.responseText);
                        }
                        else {
                            onError(xmlhttp.status);
                        }
                    }
                };
                xmlhttp.open("GET", url, true);
                xmlhttp.send();
            }
            catch (ex) {
                onError(ex);
            }
        },
        xhrGetFull: function OSF_OUtil$xhrGetFull(url, oneDriveFileName, onSuccess, onError) {
            var xmlhttp;
            var requestedFileName = oneDriveFileName;
            try {
                xmlhttp = new XMLHttpRequest();
                xmlhttp.onreadystatechange = function () {
                    if (xmlhttp.readyState == 4) {
                        if (xmlhttp.status == 200) {
                            onSuccess(xmlhttp, requestedFileName);
                        }
                        else {
                            onError(xmlhttp.status);
                        }
                    }
                };
                xmlhttp.open("GET", url, true);
                xmlhttp.send();
            }
            catch (ex) {
                onError(ex);
            }
        },
        encodeBase64: function OSF_Outil$encodeBase64(input) {
            if (!input)
                return input;
            var codex = "ABCDEFGHIJKLMNOP" + "QRSTUVWXYZabcdef" + "ghijklmnopqrstuv" + "wxyz0123456789+/=";
            var output = [];
            var temp = [];
            var index = 0;
            var c1, c2, c3, a, b, c;
            var i;
            var length = input.length;
            do {
                c1 = input.charCodeAt(index++);
                c2 = input.charCodeAt(index++);
                c3 = input.charCodeAt(index++);
                i = 0;
                a = c1 & 255;
                b = c1 >> 8;
                c = c2 & 255;
                temp[i++] = a >> 2;
                temp[i++] = ((a & 3) << 4) | (b >> 4);
                temp[i++] = ((b & 15) << 2) | (c >> 6);
                temp[i++] = c & 63;
                if (!isNaN(c2)) {
                    a = c2 >> 8;
                    b = c3 & 255;
                    c = c3 >> 8;
                    temp[i++] = a >> 2;
                    temp[i++] = ((a & 3) << 4) | (b >> 4);
                    temp[i++] = ((b & 15) << 2) | (c >> 6);
                    temp[i++] = c & 63;
                }
                if (isNaN(c2)) {
                    temp[i - 1] = 64;
                }
                else if (isNaN(c3)) {
                    temp[i - 2] = 64;
                    temp[i - 1] = 64;
                }
                for (var t = 0; t < i; t++) {
                    output.push(codex.charAt(temp[t]));
                }
            } while (index < length);
            return output.join("");
        },
        getSessionStorage: function OSF_Outil$getSessionStorage() {
            return _getSessionStorage();
        },
        getLocalStorage: function OSF_Outil$getLocalStorage() {
            if (!_safeLocalStorage) {
                try {
                    var localStorage = window.localStorage;
                }
                catch (ex) {
                    localStorage = null;
                }
                _safeLocalStorage = new OfficeExt.SafeStorage(localStorage);
            }
            return _safeLocalStorage;
        },
        convertIntToCssHexColor: function OSF_Outil$convertIntToCssHexColor(val) {
            var hex = "#" + (Number(val) + 0x1000000).toString(16).slice(-6);
            return hex;
        },
        attachClickHandler: function OSF_Outil$attachClickHandler(element, handler) {
            element.onclick = function (e) {
                handler();
            };
            element.ontouchend = function (e) {
                handler();
                e.preventDefault();
            };
        },
        getQueryStringParamValue: function OSF_Outil$getQueryStringParamValue(queryString, paramName) {
            var e = Function._validateParams(arguments, [{ name: "queryString", type: String, mayBeNull: false },
                { name: "paramName", type: String, mayBeNull: false }
            ]);
            if (e) {
                OsfMsAjaxFactory.msAjaxDebug.trace("OSF_Outil_getQueryStringParamValue: Parameters cannot be null.");
                return "";
            }
            var queryExp = new RegExp("[\\?&]" + paramName + "=([^&#]*)", "i");
            if (!queryExp.test(queryString)) {
                OsfMsAjaxFactory.msAjaxDebug.trace("OSF_Outil_getQueryStringParamValue: The parameter is not found.");
                return "";
            }
            return queryExp.exec(queryString)[1];
        },
        isiOS: function OSF_Outil$isiOS() {
            return (window.navigator.userAgent.match(/(iPad|iPhone|iPod)/g) ? true : false);
        },
        shallowCopy: function OSF_Outil$shallowCopy(sourceObj) {
            var copyObj = sourceObj.constructor();
            for (var property in sourceObj) {
                if (sourceObj.hasOwnProperty(property)) {
                    copyObj[property] = sourceObj[property];
                }
            }
            return copyObj;
        },
        serializeOMEXResponseErrorMessage: function OSF_Outil$serializeObjectToString(response) {
            if (typeof (JSON) !== "undefined") {
                try {
                    return JSON.stringify(response);
                }
                catch (ex) {
                }
            }
            return "";
        },
        createObject: function OSF_Outil$createObject(properties) {
            var obj = null;
            if (properties) {
                obj = {};
                var len = properties.length;
                for (var i = 0; i < len; i++) {
                    obj[properties[i].name] = properties[i].value;
                }
            }
            return obj;
        }
    };
})();
OSF.OUtil.Guid = (function () {
    var hexCode = ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "a", "b", "c", "d", "e", "f"];
    return {
        generateNewGuid: function OSF_Outil_Guid$generateNewGuid() {
            var result = "";
            var tick = (new Date()).getTime();
            var index = 0;
            for (; index < 32 && tick > 0; index++) {
                if (index == 8 || index == 12 || index == 16 || index == 20) {
                    result += "-";
                }
                result += hexCode[tick % 16];
                tick = Math.floor(tick / 16);
            }
            for (; index < 32; index++) {
                if (index == 8 || index == 12 || index == 16 || index == 20) {
                    result += "-";
                }
                result += hexCode[Math.floor(Math.random() * 16)];
            }
            return result;
        }
    };
})();
window.OSF = OSF;
var OfficeExt;
(function (OfficeExt) {
    var MsAjaxTypeHelper = (function () {
        function MsAjaxTypeHelper() {
        }
        MsAjaxTypeHelper.isInstanceOfType = function (type, instance) {
            if (typeof (instance) === "undefined" || instance === null)
                return false;
            if (instance instanceof type)
                return true;
            var instanceType = instance.constructor;
            if (!instanceType || (typeof (instanceType) !== "function") || instanceType.__typeName === 'Object') {
                instanceType = Object;
            }
            return !!(instanceType === type) ||
                (instanceType.inheritsFrom && instanceType.inheritsFrom(type)) ||
                (instanceType.implementsInterface && instanceType.implementsInterface(type));
        };
        return MsAjaxTypeHelper;
    })();
    OfficeExt.MsAjaxTypeHelper = MsAjaxTypeHelper;
    var MsAjaxError = (function () {
        function MsAjaxError() {
        }
        MsAjaxError.create = function (message, errorInfo) {
            var err = new Error(message);
            err.message = message;
            if (errorInfo) {
                for (var v in errorInfo) {
                    err[v] = errorInfo[v];
                }
            }
            err.popStackFrame();
            return err;
        };
        MsAjaxError.parameterCount = function (message) {
            var displayMessage = "Sys.ParameterCountException: " + (message ? message : "Parameter count mismatch.");
            var err = MsAjaxError.create(displayMessage, { name: 'Sys.ParameterCountException' });
            err.popStackFrame();
            return err;
        };
        MsAjaxError.argument = function (paramName, message) {
            var displayMessage = "Sys.ArgumentException: " + (message ? message : "Value does not fall within the expected range.");
            if (paramName) {
                displayMessage += "\n" + MsAjaxString.format("Parameter name: {0}", paramName);
            }
            var err = MsAjaxError.create(displayMessage, { name: "Sys.ArgumentException", paramName: paramName });
            err.popStackFrame();
            return err;
        };
        MsAjaxError.argumentNull = function (paramName, message) {
            var displayMessage = "Sys.ArgumentNullException: " + (message ? message : "Value cannot be null.");
            if (paramName) {
                displayMessage += "\n" + MsAjaxString.format("Parameter name: {0}", paramName);
            }
            var err = MsAjaxError.create(displayMessage, { name: "Sys.ArgumentNullException", paramName: paramName });
            err.popStackFrame();
            return err;
        };
        MsAjaxError.argumentOutOfRange = function (paramName, actualValue, message) {
            var displayMessage = "Sys.ArgumentOutOfRangeException: " + (message ? message : "Specified argument was out of the range of valid values.");
            if (paramName) {
                displayMessage += "\n" + MsAjaxString.format("Parameter name: {0}", paramName);
            }
            if (typeof (actualValue) !== "undefined" && actualValue !== null) {
                displayMessage += "\n" + MsAjaxString.format("Actual value was {0}.", actualValue);
            }
            var err = MsAjaxError.create(displayMessage, {
                name: "Sys.ArgumentOutOfRangeException",
                paramName: paramName,
                actualValue: actualValue
            });
            err.popStackFrame();
            return err;
        };
        MsAjaxError.argumentType = function (paramName, actualType, expectedType, message) {
            var displayMessage = "Sys.ArgumentTypeException: ";
            if (message) {
                displayMessage += message;
            }
            else if (actualType && expectedType) {
                displayMessage += MsAjaxString.format("Object of type '{0}' cannot be converted to type '{1}'.", actualType.getName ? actualType.getName() : actualType, expectedType.getName ? expectedType.getName() : expectedType);
            }
            else {
                displayMessage += "Object cannot be converted to the required type.";
            }
            if (paramName) {
                displayMessage += "\n" + MsAjaxString.format("Parameter name: {0}", paramName);
            }
            var err = MsAjaxError.create(displayMessage, {
                name: "Sys.ArgumentTypeException",
                paramName: paramName,
                actualType: actualType,
                expectedType: expectedType
            });
            err.popStackFrame();
            return err;
        };
        MsAjaxError.argumentUndefined = function (paramName, message) {
            var displayMessage = "Sys.ArgumentUndefinedException: " + (message ? message : "Value cannot be undefined.");
            if (paramName) {
                displayMessage += "\n" + MsAjaxString.format("Parameter name: {0}", paramName);
            }
            var err = MsAjaxError.create(displayMessage, { name: "Sys.ArgumentUndefinedException", paramName: paramName });
            err.popStackFrame();
            return err;
        };
        MsAjaxError.invalidOperation = function (message) {
            var displayMessage = "Sys.InvalidOperationException: " + (message ? message : "Operation is not valid due to the current state of the object.");
            var err = MsAjaxError.create(displayMessage, { name: 'Sys.InvalidOperationException' });
            err.popStackFrame();
            return err;
        };
        return MsAjaxError;
    })();
    OfficeExt.MsAjaxError = MsAjaxError;
    var MsAjaxString = (function () {
        function MsAjaxString() {
        }
        MsAjaxString.format = function (format) {
            var args = [];
            for (var _i = 1; _i < arguments.length; _i++) {
                args[_i - 1] = arguments[_i];
            }
            var source = format;
            return source.replace(/{(\d+)}/gm, function (match, number) {
                var index = parseInt(number, 10);
                return args[index] === undefined ? '{' + number + '}' : args[index];
            });
        };
        MsAjaxString.startsWith = function (str, prefix) {
            return (str.substr(0, prefix.length) === prefix);
        };
        return MsAjaxString;
    })();
    OfficeExt.MsAjaxString = MsAjaxString;
    var MsAjaxDebug = (function () {
        function MsAjaxDebug() {
        }
        MsAjaxDebug.trace = function (text) {
            if (typeof Debug !== "undefined" && Debug.writeln)
                Debug.writeln(text);
            if (window.console && window.console.log)
                window.console.log(text);
            if (window.opera && window.opera.postError)
                window.opera.postError(text);
            if (window.debugService && window.debugService.trace)
                window.debugService.trace(text);
            var a = document.getElementById("TraceConsole");
            if (a && a.tagName.toUpperCase() === "TEXTAREA") {
                a.innerHTML += text + "\n";
            }
        };
        return MsAjaxDebug;
    })();
    OfficeExt.MsAjaxDebug = MsAjaxDebug;
    if (!OsfMsAjaxFactory.isMsAjaxLoaded()) {
        if (!Function.createCallback) {
            Function.createCallback = function Function$createCallback(method, context) {
                var e = Function._validateParams(arguments, [
                    { name: "method", type: Function },
                    { name: "context", mayBeNull: true }
                ]);
                if (e)
                    throw e;
                return function () {
                    var l = arguments.length;
                    if (l > 0) {
                        var args = [];
                        for (var i = 0; i < l; i++) {
                            args[i] = arguments[i];
                        }
                        args[l] = context;
                        return method.apply(this, args);
                    }
                    return method.call(this, context);
                };
            };
        }
        if (!Function.createDelegate) {
            Function.createDelegate = function Function$createDelegate(instance, method) {
                var e = Function._validateParams(arguments, [
                    { name: "instance", mayBeNull: true },
                    { name: "method", type: Function }
                ]);
                if (e)
                    throw e;
                return function () {
                    return method.apply(instance, arguments);
                };
            };
        }
        if (!Function._validateParams) {
            Function._validateParams = function (params, expectedParams, validateParameterCount) {
                var e, expectedLength = expectedParams.length;
                validateParameterCount = validateParameterCount || (typeof (validateParameterCount) === "undefined");
                e = Function._validateParameterCount(params, expectedParams, validateParameterCount);
                if (e) {
                    e.popStackFrame();
                    return e;
                }
                for (var i = 0, l = params.length; i < l; i++) {
                    var expectedParam = expectedParams[Math.min(i, expectedLength - 1)], paramName = expectedParam.name;
                    if (expectedParam.parameterArray) {
                        paramName += "[" + (i - expectedLength + 1) + "]";
                    }
                    else if (!validateParameterCount && (i >= expectedLength)) {
                        break;
                    }
                    e = Function._validateParameter(params[i], expectedParam, paramName);
                    if (e) {
                        e.popStackFrame();
                        return e;
                    }
                }
                return null;
            };
        }
        if (!Function._validateParameterCount) {
            Function._validateParameterCount = function (params, expectedParams, validateParameterCount) {
                var i, error, expectedLen = expectedParams.length, actualLen = params.length;
                if (actualLen < expectedLen) {
                    var minParams = expectedLen;
                    for (i = 0; i < expectedLen; i++) {
                        var param = expectedParams[i];
                        if (param.optional || param.parameterArray) {
                            minParams--;
                        }
                    }
                    if (actualLen < minParams) {
                        error = true;
                    }
                }
                else if (validateParameterCount && (actualLen > expectedLen)) {
                    error = true;
                    for (i = 0; i < expectedLen; i++) {
                        if (expectedParams[i].parameterArray) {
                            error = false;
                            break;
                        }
                    }
                }
                if (error) {
                    var e = MsAjaxError.parameterCount();
                    e.popStackFrame();
                    return e;
                }
                return null;
            };
        }
        if (!Function._validateParameter) {
            Function._validateParameter = function (param, expectedParam, paramName) {
                var e, expectedType = expectedParam.type, expectedInteger = !!expectedParam.integer, expectedDomElement = !!expectedParam.domElement, mayBeNull = !!expectedParam.mayBeNull;
                e = Function._validateParameterType(param, expectedType, expectedInteger, expectedDomElement, mayBeNull, paramName);
                if (e) {
                    e.popStackFrame();
                    return e;
                }
                var expectedElementType = expectedParam.elementType, elementMayBeNull = !!expectedParam.elementMayBeNull;
                if (expectedType === Array && typeof (param) !== "undefined" && param !== null &&
                    (expectedElementType || !elementMayBeNull)) {
                    var expectedElementInteger = !!expectedParam.elementInteger, expectedElementDomElement = !!expectedParam.elementDomElement;
                    for (var i = 0; i < param.length; i++) {
                        var elem = param[i];
                        e = Function._validateParameterType(elem, expectedElementType, expectedElementInteger, expectedElementDomElement, elementMayBeNull, paramName + "[" + i + "]");
                        if (e) {
                            e.popStackFrame();
                            return e;
                        }
                    }
                }
                return null;
            };
        }
        if (!Function._validateParameterType) {
            Function._validateParameterType = function (param, expectedType, expectedInteger, expectedDomElement, mayBeNull, paramName) {
                var e, i;
                if (typeof (param) === "undefined") {
                    if (mayBeNull) {
                        return null;
                    }
                    else {
                        e = OfficeExt.MsAjaxError.argumentUndefined(paramName);
                        e.popStackFrame();
                        return e;
                    }
                }
                if (param === null) {
                    if (mayBeNull) {
                        return null;
                    }
                    else {
                        e = OfficeExt.MsAjaxError.argumentNull(paramName);
                        e.popStackFrame();
                        return e;
                    }
                }
                if (expectedType && !OfficeExt.MsAjaxTypeHelper.isInstanceOfType(expectedType, param)) {
                    e = OfficeExt.MsAjaxError.argumentType(paramName, typeof (param), expectedType);
                    e.popStackFrame();
                    return e;
                }
                return null;
            };
        }
        if (!window.Type) {
            window.Type = Function;
        }
        if (!Type.registerNamespace) {
            Type.registerNamespace = function (ns) {
                var namespaceParts = ns.split('.');
                var currentNamespace = window;
                for (var i = 0; i < namespaceParts.length; i++) {
                    currentNamespace[namespaceParts[i]] = currentNamespace[namespaceParts[i]] || {};
                    currentNamespace = currentNamespace[namespaceParts[i]];
                }
            };
        }
        if (!Type.prototype.registerClass) {
            Type.prototype.registerClass = function (cls) { cls = {}; };
        }
        if (typeof (Sys) === "undefined") {
            Type.registerNamespace('Sys');
        }
        if (!Error.prototype.popStackFrame) {
            Error.prototype.popStackFrame = function () {
                if (arguments.length !== 0)
                    throw MsAjaxError.parameterCount();
                if (typeof (this.stack) === "undefined" || this.stack === null ||
                    typeof (this.fileName) === "undefined" || this.fileName === null ||
                    typeof (this.lineNumber) === "undefined" || this.lineNumber === null) {
                    return;
                }
                var stackFrames = this.stack.split("\n");
                var currentFrame = stackFrames[0];
                var pattern = this.fileName + ":" + this.lineNumber;
                while (typeof (currentFrame) !== "undefined" &&
                    currentFrame !== null &&
                    currentFrame.indexOf(pattern) === -1) {
                    stackFrames.shift();
                    currentFrame = stackFrames[0];
                }
                var nextFrame = stackFrames[1];
                if (typeof (nextFrame) === "undefined" || nextFrame === null) {
                    return;
                }
                var nextFrameParts = nextFrame.match(/@(.*):(\d+)$/);
                if (typeof (nextFrameParts) === "undefined" || nextFrameParts === null) {
                    return;
                }
                this.fileName = nextFrameParts[1];
                this.lineNumber = parseInt(nextFrameParts[2]);
                stackFrames.shift();
                this.stack = stackFrames.join("\n");
            };
        }
        OsfMsAjaxFactory.msAjaxError = MsAjaxError;
        OsfMsAjaxFactory.msAjaxString = MsAjaxString;
        OsfMsAjaxFactory.msAjaxDebug = MsAjaxDebug;
    }
})(OfficeExt || (OfficeExt = {}));
var OfficeExt;
(function (OfficeExt) {
    var MsAjaxJavaScriptSerializer = (function () {
        function MsAjaxJavaScriptSerializer() {
        }
        MsAjaxJavaScriptSerializer._init = function () {
            var replaceChars = ['\\u0000', '\\u0001', '\\u0002', '\\u0003', '\\u0004', '\\u0005', '\\u0006', '\\u0007',
                '\\b', '\\t', '\\n', '\\u000b', '\\f', '\\r', '\\u000e', '\\u000f', '\\u0010', '\\u0011',
                '\\u0012', '\\u0013', '\\u0014', '\\u0015', '\\u0016', '\\u0017', '\\u0018', '\\u0019',
                '\\u001a', '\\u001b', '\\u001c', '\\u001d', '\\u001e', '\\u001f'];
            MsAjaxJavaScriptSerializer._charsToEscape[0] = '\\';
            MsAjaxJavaScriptSerializer._charsToEscapeRegExs['\\'] = new RegExp('\\\\', 'g');
            MsAjaxJavaScriptSerializer._escapeChars['\\'] = '\\\\';
            MsAjaxJavaScriptSerializer._charsToEscape[1] = '"';
            MsAjaxJavaScriptSerializer._charsToEscapeRegExs['"'] = new RegExp('"', 'g');
            MsAjaxJavaScriptSerializer._escapeChars['"'] = '\\"';
            for (var i = 0; i < 32; i++) {
                var c = String.fromCharCode(i);
                MsAjaxJavaScriptSerializer._charsToEscape[i + 2] = c;
                MsAjaxJavaScriptSerializer._charsToEscapeRegExs[c] = new RegExp(c, 'g');
                MsAjaxJavaScriptSerializer._escapeChars[c] = replaceChars[i];
            }
        };
        MsAjaxJavaScriptSerializer.serialize = function (object) {
            var stringBuilder = new MsAjaxStringBuilder();
            MsAjaxJavaScriptSerializer.serializeWithBuilder(object, stringBuilder, false);
            return stringBuilder.toString();
        };
        MsAjaxJavaScriptSerializer.deserialize = function (data, secure) {
            if (data.length === 0)
                throw OfficeExt.MsAjaxError.argument('data', "Cannot deserialize empty string.");
            try {
                var exp = data.replace(MsAjaxJavaScriptSerializer._dateRegEx, "$1new Date($2)");
                if (secure && MsAjaxJavaScriptSerializer._jsonRegEx.test(exp.replace(MsAjaxJavaScriptSerializer._jsonStringRegEx, '')))
                    throw null;
                return eval('(' + exp + ')');
            }
            catch (e) {
                throw OfficeExt.MsAjaxError.argument('data', "Cannot deserialize. The data does not correspond to valid JSON.");
            }
        };
        MsAjaxJavaScriptSerializer.serializeBooleanWithBuilder = function (object, stringBuilder) {
            stringBuilder.append(object.toString());
        };
        MsAjaxJavaScriptSerializer.serializeNumberWithBuilder = function (object, stringBuilder) {
            if (isFinite(object)) {
                stringBuilder.append(String(object));
            }
            else {
                throw OfficeExt.MsAjaxError.invalidOperation("Cannot serialize non finite numbers.");
            }
        };
        MsAjaxJavaScriptSerializer.serializeStringWithBuilder = function (str, stringBuilder) {
            stringBuilder.append('"');
            if (MsAjaxJavaScriptSerializer._escapeRegEx.test(str)) {
                if (MsAjaxJavaScriptSerializer._charsToEscape.length === 0) {
                    MsAjaxJavaScriptSerializer._init();
                }
                if (str.length < 128) {
                    str = str.replace(MsAjaxJavaScriptSerializer._escapeRegExGlobal, function (x) { return MsAjaxJavaScriptSerializer._escapeChars[x]; });
                }
                else {
                    for (var i = 0; i < 34; i++) {
                        var c = MsAjaxJavaScriptSerializer._charsToEscape[i];
                        if (str.indexOf(c) !== -1) {
                            if ((navigator.userAgent.indexOf("OPR/") > -1) || (navigator.userAgent.indexOf("Firefox") > -1)) {
                                str = str.split(c).join(MsAjaxJavaScriptSerializer._escapeChars[c]);
                            }
                            else {
                                str = str.replace(MsAjaxJavaScriptSerializer._charsToEscapeRegExs[c], MsAjaxJavaScriptSerializer._escapeChars[c]);
                            }
                        }
                    }
                }
            }
            stringBuilder.append(str);
            stringBuilder.append('"');
        };
        MsAjaxJavaScriptSerializer.serializeWithBuilder = function (object, stringBuilder, sort, prevObjects) {
            var i;
            switch (typeof object) {
                case 'object':
                    if (object) {
                        if (prevObjects) {
                            for (var j = 0; j < prevObjects.length; j++) {
                                if (prevObjects[j] === object) {
                                    throw OfficeExt.MsAjaxError.invalidOperation("Cannot serialize object with cyclic reference within child properties.");
                                }
                            }
                        }
                        else {
                            prevObjects = new Array();
                        }
                        try {
                            OfficeExt.MsAjaxArray.add(prevObjects, object);
                            if (OfficeExt.MsAjaxTypeHelper.isInstanceOfType(Number, object)) {
                                MsAjaxJavaScriptSerializer.serializeNumberWithBuilder(object, stringBuilder);
                            }
                            else if (OfficeExt.MsAjaxTypeHelper.isInstanceOfType(Boolean, object)) {
                                MsAjaxJavaScriptSerializer.serializeBooleanWithBuilder(object, stringBuilder);
                            }
                            else if (OfficeExt.MsAjaxTypeHelper.isInstanceOfType(String, object)) {
                                MsAjaxJavaScriptSerializer.serializeStringWithBuilder(object, stringBuilder);
                            }
                            else if (OfficeExt.MsAjaxTypeHelper.isInstanceOfType(Array, object)) {
                                stringBuilder.append('[');
                                for (i = 0; i < object.length; ++i) {
                                    if (i > 0) {
                                        stringBuilder.append(',');
                                    }
                                    MsAjaxJavaScriptSerializer.serializeWithBuilder(object[i], stringBuilder, false, prevObjects);
                                }
                                stringBuilder.append(']');
                            }
                            else {
                                if (OfficeExt.MsAjaxTypeHelper.isInstanceOfType(Date, object)) {
                                    stringBuilder.append('"\\/Date(');
                                    stringBuilder.append(object.getTime());
                                    stringBuilder.append(')\\/"');
                                    break;
                                }
                                var properties = [];
                                var propertyCount = 0;
                                for (var name in object) {
                                    if (OfficeExt.MsAjaxString.startsWith(name, '$')) {
                                        continue;
                                    }
                                    if (name === MsAjaxJavaScriptSerializer._serverTypeFieldName && propertyCount !== 0) {
                                        properties[propertyCount++] = properties[0];
                                        properties[0] = name;
                                    }
                                    else {
                                        properties[propertyCount++] = name;
                                    }
                                }
                                if (sort)
                                    properties.sort();
                                stringBuilder.append('{');
                                var needComma = false;
                                for (i = 0; i < propertyCount; i++) {
                                    var value = object[properties[i]];
                                    if (typeof value !== 'undefined' && typeof value !== 'function') {
                                        if (needComma) {
                                            stringBuilder.append(',');
                                        }
                                        else {
                                            needComma = true;
                                        }
                                        MsAjaxJavaScriptSerializer.serializeWithBuilder(properties[i], stringBuilder, sort, prevObjects);
                                        stringBuilder.append(':');
                                        MsAjaxJavaScriptSerializer.serializeWithBuilder(value, stringBuilder, sort, prevObjects);
                                    }
                                }
                                stringBuilder.append('}');
                            }
                        }
                        finally {
                            OfficeExt.MsAjaxArray.removeAt(prevObjects, prevObjects.length - 1);
                        }
                    }
                    else {
                        stringBuilder.append('null');
                    }
                    break;
                case 'number':
                    MsAjaxJavaScriptSerializer.serializeNumberWithBuilder(object, stringBuilder);
                    break;
                case 'string':
                    MsAjaxJavaScriptSerializer.serializeStringWithBuilder(object, stringBuilder);
                    break;
                case 'boolean':
                    MsAjaxJavaScriptSerializer.serializeBooleanWithBuilder(object, stringBuilder);
                    break;
                default:
                    stringBuilder.append('null');
                    break;
            }
        };
        MsAjaxJavaScriptSerializer.__patchVersion = 0;
        MsAjaxJavaScriptSerializer._charsToEscapeRegExs = [];
        MsAjaxJavaScriptSerializer._charsToEscape = [];
        MsAjaxJavaScriptSerializer._dateRegEx = new RegExp('(^|[^\\\\])\\"\\\\/Date\\((-?[0-9]+)(?:[a-zA-Z]|(?:\\+|-)[0-9]{4})?\\)\\\\/\\"', 'g');
        MsAjaxJavaScriptSerializer._escapeChars = {};
        MsAjaxJavaScriptSerializer._escapeRegEx = new RegExp('["\\\\\\x00-\\x1F]', 'i');
        MsAjaxJavaScriptSerializer._escapeRegExGlobal = new RegExp('["\\\\\\x00-\\x1F]', 'g');
        MsAjaxJavaScriptSerializer._jsonRegEx = new RegExp('[^,:{}\\[\\]0-9.\\-+Eaeflnr-u \\n\\r\\t]', 'g');
        MsAjaxJavaScriptSerializer._jsonStringRegEx = new RegExp('"(\\\\.|[^"\\\\])*"', 'g');
        MsAjaxJavaScriptSerializer._serverTypeFieldName = '__type';
        return MsAjaxJavaScriptSerializer;
    })();
    OfficeExt.MsAjaxJavaScriptSerializer = MsAjaxJavaScriptSerializer;
    var MsAjaxArray = (function () {
        function MsAjaxArray() {
        }
        MsAjaxArray.add = function (array, item) {
            array[array.length] = item;
        };
        MsAjaxArray.removeAt = function (array, index) {
            array.splice(index, 1);
        };
        MsAjaxArray.clone = function (array) {
            if (array.length === 1) {
                return [array[0]];
            }
            else {
                return Array.apply(null, array);
            }
        };
        MsAjaxArray.remove = function (array, item) {
            var index = MsAjaxArray.indexOf(array, item);
            if (index >= 0) {
                array.splice(index, 1);
            }
            return (index >= 0);
        };
        MsAjaxArray.indexOf = function (array, item, start) {
            if (typeof (item) === "undefined")
                return -1;
            var length = array.length;
            if (length !== 0) {
                start = start - 0;
                if (isNaN(start)) {
                    start = 0;
                }
                else {
                    if (isFinite(start)) {
                        start = start - (start % 1);
                    }
                    if (start < 0) {
                        start = Math.max(0, length + start);
                    }
                }
                for (var i = start; i < length; i++) {
                    if ((typeof (array[i]) !== "undefined") && (array[i] === item)) {
                        return i;
                    }
                }
            }
            return -1;
        };
        return MsAjaxArray;
    })();
    OfficeExt.MsAjaxArray = MsAjaxArray;
    var MsAjaxStringBuilder = (function () {
        function MsAjaxStringBuilder(initialText) {
            this._parts = (typeof (initialText) !== 'undefined' && initialText !== null && initialText !== '') ?
                [initialText.toString()] : [];
            this._value = {};
            this._len = 0;
        }
        MsAjaxStringBuilder.prototype.append = function (text) {
            this._parts[this._parts.length] = text;
        };
        MsAjaxStringBuilder.prototype.toString = function (separator) {
            separator = separator || '';
            var parts = this._parts;
            if (this._len !== parts.length) {
                this._value = {};
                this._len = parts.length;
            }
            var val = this._value;
            if (typeof (val[separator]) === 'undefined') {
                if (separator !== '') {
                    for (var i = 0; i < parts.length;) {
                        if ((typeof (parts[i]) === 'undefined') || (parts[i] === '') || (parts[i] === null)) {
                            parts.splice(i, 1);
                        }
                        else {
                            i++;
                        }
                    }
                }
                val[separator] = this._parts.join(separator);
            }
            return val[separator];
        };
        return MsAjaxStringBuilder;
    })();
    OfficeExt.MsAjaxStringBuilder = MsAjaxStringBuilder;
    if (!OsfMsAjaxFactory.isMsAjaxLoaded()) {
        OsfMsAjaxFactory.msAjaxSerializer = MsAjaxJavaScriptSerializer;
    }
})(OfficeExt || (OfficeExt = {}));
OSF.OUtil.setNamespace("Microsoft", window);
OSF.OUtil.setNamespace("Office", Microsoft);
OSF.OUtil.setNamespace("Common", Microsoft.Office);
OSF.SerializerVersion = {
    MsAjax: 0,
    Browser: 1
};
(function (window) {
    "use strict";
    var stringRegEx = new RegExp('"(\\\\.|[^"\\\\])*"', 'g'), trueFalseNullRegEx = new RegExp('\\b(true|false|null)\\b', 'g'), numbersRegEx = new RegExp('-?(0|([1-9]\\d*))(\\.\\d+)?([eE][+-]?\\d+)?', 'g'), badBracketsRegEx = new RegExp('[^{:,\\[\\s](?=\\s*\\[)'), badRemainderRegEx = new RegExp('[^\\s\\[\\]{}:,]'), jsonErrorMsg = "Cannot deserialize. The data does not correspond to valid JSON.";
    function addHandler(element, eventName, handler) {
        if (element.addEventListener) {
            element.addEventListener(eventName, handler, false);
        }
        else if (element.attachEvent) {
            element.attachEvent("on" + eventName, handler);
        }
    }
    function getAjaxSerializer() {
        if (OsfMsAjaxFactory.msAjaxSerializer) {
            return OsfMsAjaxFactory.msAjaxSerializer;
        }
        return null;
    }
    function deserialize(data, secure, oldDeserialize) {
        var transformed;
        if (!secure) {
            return oldDeserialize(data);
        }
        if (window.JSON && window.JSON.parse) {
            return window.JSON.parse(data);
        }
        transformed = data.replace(stringRegEx, "[]");
        transformed = transformed.replace(trueFalseNullRegEx, "[]");
        transformed = transformed.replace(numbersRegEx, "[]");
        if (badBracketsRegEx.test(transformed)) {
            throw jsonErrorMsg;
        }
        if (badRemainderRegEx.test(transformed)) {
            throw jsonErrorMsg;
        }
        try {
            eval("(" + data + ")");
        }
        catch (e) {
            throw jsonErrorMsg;
        }
    }
    function patchDeserializer() {
        var serializer = getAjaxSerializer(), oldDeserialize;
        if (serializer === null || typeof (serializer.deserialize) !== "function") {
            return false;
        }
        if (serializer.__patchVersion >= 1) {
            return true;
        }
        oldDeserialize = serializer.deserialize;
        serializer.deserialize = function (data, secure) {
            return deserialize(data, true, oldDeserialize);
        };
        serializer.__patchVersion = 1;
        return true;
    }
    if (!patchDeserializer()) {
        addHandler(window, "load", function () {
            patchDeserializer();
        });
    }
}(window));
Microsoft.Office.Common.InvokeType = { "async": 0,
    "sync": 1,
    "asyncRegisterEvent": 2,
    "asyncUnregisterEvent": 3,
    "syncRegisterEvent": 4,
    "syncUnregisterEvent": 5
};
Microsoft.Office.Common.InvokeResultCode = {
    "noError": 0,
    "errorInRequest": -1,
    "errorHandlingRequest": -2,
    "errorInResponse": -3,
    "errorHandlingResponse": -4,
    "errorHandlingRequestAccessDenied": -5,
    "errorHandlingMethodCallTimedout": -6
};
Microsoft.Office.Common.MessageType = { "request": 0,
    "response": 1
};
Microsoft.Office.Common.ActionType = { "invoke": 0,
    "registerEvent": 1,
    "unregisterEvent": 2 };
Microsoft.Office.Common.ResponseType = { "forCalling": 0,
    "forEventing": 1
};
Microsoft.Office.Common.MethodObject = function Microsoft_Office_Common_MethodObject(method, invokeType, blockingOthers) {
    this._method = method;
    this._invokeType = invokeType;
    this._blockingOthers = blockingOthers;
};
Microsoft.Office.Common.MethodObject.prototype = {
    getMethod: function Microsoft_Office_Common_MethodObject$getMethod() {
        return this._method;
    },
    getInvokeType: function Microsoft_Office_Common_MethodObject$getInvokeType() {
        return this._invokeType;
    },
    getBlockingFlag: function Microsoft_Office_Common_MethodObject$getBlockingFlag() {
        return this._blockingOthers;
    }
};
Microsoft.Office.Common.EventMethodObject = function Microsoft_Office_Common_EventMethodObject(registerMethodObject, unregisterMethodObject) {
    this._registerMethodObject = registerMethodObject;
    this._unregisterMethodObject = unregisterMethodObject;
};
Microsoft.Office.Common.EventMethodObject.prototype = {
    getRegisterMethodObject: function Microsoft_Office_Common_EventMethodObject$getRegisterMethodObject() {
        return this._registerMethodObject;
    },
    getUnregisterMethodObject: function Microsoft_Office_Common_EventMethodObject$getUnregisterMethodObject() {
        return this._unregisterMethodObject;
    }
};
Microsoft.Office.Common.ServiceEndPoint = function Microsoft_Office_Common_ServiceEndPoint(serviceEndPointId) {
    var e = Function._validateParams(arguments, [
        { name: "serviceEndPointId", type: String, mayBeNull: false }
    ]);
    if (e)
        throw e;
    this._methodObjectList = {};
    this._eventHandlerProxyList = {};
    this._Id = serviceEndPointId;
    this._conversations = {};
    this._policyManager = null;
    this._appDomains = {};
    this._onHandleRequestError = null;
};
Microsoft.Office.Common.ServiceEndPoint.prototype = {
    registerMethod: function Microsoft_Office_Common_ServiceEndPoint$registerMethod(methodName, method, invokeType, blockingOthers) {
        var e = Function._validateParams(arguments, [{ name: "methodName", type: String, mayBeNull: false },
            { name: "method", type: Function, mayBeNull: false },
            { name: "invokeType", type: Number, mayBeNull: false },
            { name: "blockingOthers", type: Boolean, mayBeNull: false }
        ]);
        if (e)
            throw e;
        if (invokeType !== Microsoft.Office.Common.InvokeType.async
            && invokeType !== Microsoft.Office.Common.InvokeType.sync) {
            throw OsfMsAjaxFactory.msAjaxError.argument("invokeType");
        }
        var methodObject = new Microsoft.Office.Common.MethodObject(method, invokeType, blockingOthers);
        this._methodObjectList[methodName] = methodObject;
    },
    unregisterMethod: function Microsoft_Office_Common_ServiceEndPoint$unregisterMethod(methodName) {
        var e = Function._validateParams(arguments, [
            { name: "methodName", type: String, mayBeNull: false }
        ]);
        if (e)
            throw e;
        delete this._methodObjectList[methodName];
    },
    registerEvent: function Microsoft_Office_Common_ServiceEndPoint$registerEvent(eventName, registerMethod, unregisterMethod) {
        var e = Function._validateParams(arguments, [{ name: "eventName", type: String, mayBeNull: false },
            { name: "registerMethod", type: Function, mayBeNull: false },
            { name: "unregisterMethod", type: Function, mayBeNull: false }
        ]);
        if (e)
            throw e;
        var methodObject = new Microsoft.Office.Common.EventMethodObject(new Microsoft.Office.Common.MethodObject(registerMethod, Microsoft.Office.Common.InvokeType.syncRegisterEvent, false), new Microsoft.Office.Common.MethodObject(unregisterMethod, Microsoft.Office.Common.InvokeType.syncUnregisterEvent, false));
        this._methodObjectList[eventName] = methodObject;
    },
    registerEventEx: function Microsoft_Office_Common_ServiceEndPoint$registerEventEx(eventName, registerMethod, registerMethodInvokeType, unregisterMethod, unregisterMethodInvokeType) {
        var e = Function._validateParams(arguments, [{ name: "eventName", type: String, mayBeNull: false },
            { name: "registerMethod", type: Function, mayBeNull: false },
            { name: "registerMethodInvokeType", type: Number, mayBeNull: false },
            { name: "unregisterMethod", type: Function, mayBeNull: false },
            { name: "unregisterMethodInvokeType", type: Number, mayBeNull: false }
        ]);
        if (e)
            throw e;
        var methodObject = new Microsoft.Office.Common.EventMethodObject(new Microsoft.Office.Common.MethodObject(registerMethod, registerMethodInvokeType, false), new Microsoft.Office.Common.MethodObject(unregisterMethod, unregisterMethodInvokeType, false));
        this._methodObjectList[eventName] = methodObject;
    },
    unregisterEvent: function (eventName) {
        var e = Function._validateParams(arguments, [
            { name: "eventName", type: String, mayBeNull: false }
        ]);
        if (e)
            throw e;
        this.unregisterMethod(eventName);
    },
    registerConversation: function Microsoft_Office_Common_ServiceEndPoint$registerConversation(conversationId, conversationUrl, appDomains, serializerVersion) {
        var e = Function._validateParams(arguments, [
            { name: "conversationId", type: String, mayBeNull: false },
            { name: "conversationUrl", type: String, mayBeNull: false, optional: true },
            { name: "appDomains", type: Object, mayBeNull: true, optional: true },
            { name: "serializerVersion", type: Number, mayBeNull: true, optional: true }
        ]);
        if (e)
            throw e;
        ;
        if (appDomains) {
            if (!(appDomains instanceof Array)) {
                throw OsfMsAjaxFactory.msAjaxError.argument("appDomains");
            }
            this._appDomains[conversationId] = appDomains;
        }
        this._conversations[conversationId] = { url: conversationUrl, serializerVersion: serializerVersion };
    },
    unregisterConversation: function Microsoft_Office_Common_ServiceEndPoint$unregisterConversation(conversationId) {
        var e = Function._validateParams(arguments, [
            { name: "conversationId", type: String, mayBeNull: false }
        ]);
        if (e)
            throw e;
        delete this._conversations[conversationId];
    },
    setPolicyManager: function Microsoft_Office_Common_ServiceEndPoint$setPolicyManager(policyManager) {
        var e = Function._validateParams(arguments, [
            { name: "policyManager", type: Object, mayBeNull: false }
        ]);
        if (e)
            throw e;
        if (!policyManager.checkPermission) {
            throw OsfMsAjaxFactory.msAjaxError.argument("policyManager");
        }
        this._policyManager = policyManager;
    },
    getPolicyManager: function Microsoft_Office_Common_ServiceEndPoint$getPolicyManager() {
        return this._policyManager;
    }
};
Microsoft.Office.Common.ClientEndPoint = function Microsoft_Office_Common_ClientEndPoint(conversationId, targetWindow, targetUrl, serializerVersion) {
    var e = Function._validateParams(arguments, [
        { name: "conversationId", type: String, mayBeNull: false },
        { name: "targetWindow", mayBeNull: false },
        { name: "targetUrl", type: String, mayBeNull: false },
        { name: "serializerVersion", type: Number, mayBeNull: true, optional: true }
    ]);
    if (e)
        throw e;
    if (!targetWindow.postMessage) {
        throw OsfMsAjaxFactory.msAjaxError.argument("targetWindow");
    }
    this._conversationId = conversationId;
    this._targetWindow = targetWindow;
    this._targetUrl = targetUrl;
    this._callingIndex = 0;
    this._callbackList = {};
    this._eventHandlerList = {};
    if (serializerVersion != null) {
        this._serializerVersion = serializerVersion;
    }
    else {
        this._serializerVersion = OSF.SerializerVersion.MsAjax;
    }
};
Microsoft.Office.Common.ClientEndPoint.prototype = {
    invoke: function Microsoft_Office_Common_ClientEndPoint$invoke(targetMethodName, callback, param) {
        var e = Function._validateParams(arguments, [{ name: "targetMethodName", type: String, mayBeNull: false },
            { name: "callback", type: Function, mayBeNull: true },
            { name: "param", mayBeNull: true }
        ]);
        if (e)
            throw e;
        var correlationId = this._callingIndex++;
        var now = new Date();
        var callbackEntry = { "callback": callback, "createdOn": now.getTime() };
        if (param && typeof param === "object" && typeof param.__timeout__ === "number") {
            callbackEntry.timeout = param.__timeout__;
            delete param.__timeout__;
        }
        this._callbackList[correlationId] = callbackEntry;
        try {
            var callRequest = new Microsoft.Office.Common.Request(targetMethodName, Microsoft.Office.Common.ActionType.invoke, this._conversationId, correlationId, param);
            var msg = Microsoft.Office.Common.MessagePackager.envelope(callRequest, this._serializerVersion);
            this._targetWindow.postMessage(msg, this._targetUrl);
            Microsoft.Office.Common.XdmCommunicationManager._startMethodTimeoutTimer();
        }
        catch (ex) {
            try {
                if (callback !== null)
                    callback(Microsoft.Office.Common.InvokeResultCode.errorInRequest, ex);
            }
            finally {
                delete this._callbackList[correlationId];
            }
        }
    },
    registerForEvent: function Microsoft_Office_Common_ClientEndPoint$registerForEvent(targetEventName, eventHandler, callback, data) {
        var e = Function._validateParams(arguments, [{ name: "targetEventName", type: String, mayBeNull: false },
            { name: "eventHandler", type: Function, mayBeNull: false },
            { name: "callback", type: Function, mayBeNull: true },
            { name: "data", mayBeNull: true, optional: true }
        ]);
        if (e)
            throw e;
        var correlationId = this._callingIndex++;
        var now = new Date();
        this._callbackList[correlationId] = { "callback": callback, "createdOn": now.getTime() };
        try {
            var callRequest = new Microsoft.Office.Common.Request(targetEventName, Microsoft.Office.Common.ActionType.registerEvent, this._conversationId, correlationId, data);
            var msg = Microsoft.Office.Common.MessagePackager.envelope(callRequest, this._serializerVersion);
            this._targetWindow.postMessage(msg, this._targetUrl);
            Microsoft.Office.Common.XdmCommunicationManager._startMethodTimeoutTimer();
            this._eventHandlerList[targetEventName] = eventHandler;
        }
        catch (ex) {
            try {
                if (callback !== null) {
                    callback(Microsoft.Office.Common.InvokeResultCode.errorInRequest, ex);
                }
            }
            finally {
                delete this._callbackList[correlationId];
            }
        }
    },
    unregisterForEvent: function Microsoft_Office_Common_ClientEndPoint$unregisterForEvent(targetEventName, callback, data) {
        var e = Function._validateParams(arguments, [{ name: "targetEventName", type: String, mayBeNull: false },
            { name: "callback", type: Function, mayBeNull: true },
            { name: "data", mayBeNull: true, optional: true }
        ]);
        if (e)
            throw e;
        var correlationId = this._callingIndex++;
        var now = new Date();
        this._callbackList[correlationId] = { "callback": callback, "createdOn": now.getTime() };
        try {
            var callRequest = new Microsoft.Office.Common.Request(targetEventName, Microsoft.Office.Common.ActionType.unregisterEvent, this._conversationId, correlationId, data);
            var msg = Microsoft.Office.Common.MessagePackager.envelope(callRequest, this._serializerVersion);
            this._targetWindow.postMessage(msg, this._targetUrl);
            Microsoft.Office.Common.XdmCommunicationManager._startMethodTimeoutTimer();
        }
        catch (ex) {
            try {
                if (callback !== null) {
                    callback(Microsoft.Office.Common.InvokeResultCode.errorInRequest, ex);
                }
            }
            finally {
                delete this._callbackList[correlationId];
            }
        }
        finally {
            delete this._eventHandlerList[targetEventName];
        }
    }
};
Microsoft.Office.Common.XdmCommunicationManager = (function () {
    var _invokerQueue = [];
    var _lastMessageProcessTime = null;
    var _messageProcessingTimer = null;
    var _processInterval = 10;
    var _blockingFlag = false;
    var _methodTimeoutTimer = null;
    var _methodTimeoutProcessInterval = 2000;
    var _methodTimeoutDefault = 65000;
    var _methodTimeout = _methodTimeoutDefault;
    var _serviceEndPoints = {};
    var _clientEndPoints = {};
    var _initialized = false;
    function _lookupServiceEndPoint(conversationId) {
        for (var id in _serviceEndPoints) {
            if (_serviceEndPoints[id]._conversations[conversationId]) {
                return _serviceEndPoints[id];
            }
        }
        OsfMsAjaxFactory.msAjaxDebug.trace("Unknown conversation Id.");
        throw OsfMsAjaxFactory.msAjaxError.argument("conversationId");
    }
    ;
    function _lookupClientEndPoint(conversationId) {
        var clientEndPoint = _clientEndPoints[conversationId];
        if (!clientEndPoint) {
            OsfMsAjaxFactory.msAjaxDebug.trace("Unknown conversation Id.");
            throw OsfMsAjaxFactory.msAjaxError.argument("conversationId");
        }
        return clientEndPoint;
    }
    ;
    function _lookupMethodObject(serviceEndPoint, messageObject) {
        var methodOrEventMethodObject = serviceEndPoint._methodObjectList[messageObject._actionName];
        if (!methodOrEventMethodObject) {
            OsfMsAjaxFactory.msAjaxDebug.trace("The specified method is not registered on service endpoint:" + messageObject._actionName);
            throw OsfMsAjaxFactory.msAjaxError.argument("messageObject");
        }
        var methodObject = null;
        if (messageObject._actionType === Microsoft.Office.Common.ActionType.invoke) {
            methodObject = methodOrEventMethodObject;
        }
        else if (messageObject._actionType === Microsoft.Office.Common.ActionType.registerEvent) {
            methodObject = methodOrEventMethodObject.getRegisterMethodObject();
        }
        else {
            methodObject = methodOrEventMethodObject.getUnregisterMethodObject();
        }
        return methodObject;
    }
    ;
    function _enqueInvoker(invoker) {
        _invokerQueue.push(invoker);
    }
    ;
    function _dequeInvoker() {
        if (_messageProcessingTimer !== null) {
            if (!_blockingFlag) {
                if (_invokerQueue.length > 0) {
                    var invoker = _invokerQueue.shift();
                    _executeCommand(invoker);
                }
                else {
                    clearInterval(_messageProcessingTimer);
                    _messageProcessingTimer = null;
                }
            }
        }
        else {
            OsfMsAjaxFactory.msAjaxDebug.trace("channel is not ready.");
        }
    }
    ;
    function _executeCommand(invoker) {
        _blockingFlag = invoker.getInvokeBlockingFlag();
        invoker.invoke();
        _lastMessageProcessTime = (new Date()).getTime();
    }
    ;
    function _checkMethodTimeout() {
        if (_methodTimeoutTimer) {
            var clientEndPoint;
            var methodCallsNotTimedout = 0;
            var now = new Date();
            var timeoutValue;
            for (var conversationId in _clientEndPoints) {
                clientEndPoint = _clientEndPoints[conversationId];
                for (var correlationId in clientEndPoint._callbackList) {
                    var callbackEntry = clientEndPoint._callbackList[correlationId];
                    timeoutValue = callbackEntry.timeout ? callbackEntry.timeout : _methodTimeout;
                    if (timeoutValue >= 0 && Math.abs(now.getTime() - callbackEntry.createdOn) >= timeoutValue) {
                        try {
                            if (callbackEntry.callback) {
                                callbackEntry.callback(Microsoft.Office.Common.InvokeResultCode.errorHandlingMethodCallTimedout, null);
                            }
                        }
                        finally {
                            delete clientEndPoint._callbackList[correlationId];
                        }
                    }
                    else {
                        methodCallsNotTimedout++;
                    }
                    ;
                }
            }
            if (methodCallsNotTimedout === 0) {
                clearInterval(_methodTimeoutTimer);
                _methodTimeoutTimer = null;
            }
        }
        else {
            OsfMsAjaxFactory.msAjaxDebug.trace("channel is not ready.");
        }
    }
    ;
    function _postCallbackHandler() {
        _blockingFlag = false;
    }
    ;
    function _registerListener(listener) {
        if (window.addEventListener) {
            window.addEventListener("message", listener, false);
        }
        else if ((navigator.userAgent.indexOf("MSIE") > -1) && window.attachEvent) {
            window.attachEvent("onmessage", listener);
        }
        else {
            OsfMsAjaxFactory.msAjaxDebug.trace("Browser doesn't support the required API.");
            throw OsfMsAjaxFactory.msAjaxError.argument("Browser");
        }
    }
    ;
    function _checkOrigin(url, origin) {
        var res = false;
        if (url === true) {
            return true;
        }
        if (!url || !origin || !url.length || !origin.length) {
            return res;
        }
        var url_parser, org_parser;
        url_parser = document.createElement('a');
        org_parser = document.createElement('a');
        url_parser.href = url;
        org_parser.href = origin;
        res = _urlCompare(url_parser, org_parser);
        delete url_parser, org_parser;
        return res;
    }
    function _checkOriginWithAppDomains(allowed_domains, origin) {
        var res = false;
        if (!origin || !origin.length || !(allowed_domains) || !(allowed_domains instanceof Array) || !allowed_domains.length) {
            return res;
        }
        var org_parser = document.createElement('a');
        var app_domain_parser = document.createElement('a');
        org_parser.href = origin;
        for (var i = 0; i < allowed_domains.length && !res; i++) {
            if (allowed_domains[i].indexOf("://") !== -1) {
                app_domain_parser.href = allowed_domains[i];
                res = _urlCompare(org_parser, app_domain_parser);
            }
        }
        delete org_parser, app_domain_parser;
        return res;
    }
    function _urlCompare(url_parser1, url_parser2) {
        return ((url_parser1.hostname == url_parser2.hostname) &&
            (url_parser1.protocol == url_parser2.protocol) &&
            (url_parser1.port == url_parser2.port));
    }
    function _receive(e) {
        if (e.data != '') {
            var messageObject;
            var serializerVersion = OSF.SerializerVersion.MsAjax;
            var serializedMessage = e.data;
            try {
                messageObject = Microsoft.Office.Common.MessagePackager.unenvelope(serializedMessage, OSF.SerializerVersion.Browser);
                serializerVersion = messageObject._serializerVersion != null ? messageObject._serializerVersion : serializerVersion;
            }
            catch (ex) {
            }
            if (serializerVersion != OSF.SerializerVersion.Browser) {
                try {
                    messageObject = Microsoft.Office.Common.MessagePackager.unenvelope(serializedMessage, serializerVersion);
                }
                catch (ex) {
                    return;
                }
            }
            if (typeof (messageObject._messageType) == 'undefined') {
                return;
            }
            if (messageObject._messageType === Microsoft.Office.Common.MessageType.request) {
                var requesterUrl = (e.origin == null || e.origin == "null") ? messageObject._origin : e.origin;
                try {
                    var serviceEndPoint = _lookupServiceEndPoint(messageObject._conversationId);
                    ;
                    var conversation = serviceEndPoint._conversations[messageObject._conversationId];
                    serializerVersion = conversation.serializerVersion != null ? conversation.serializerVersion : serializerVersion;
                    ;
                    if (!_checkOrigin(conversation.url, e.origin) && !_checkOriginWithAppDomains(serviceEndPoint._appDomains[messageObject._conversationId], e.origin)) {
                        throw "Failed origin check";
                    }
                    var policyManager = serviceEndPoint.getPolicyManager();
                    if (policyManager && !policyManager.checkPermission(messageObject._conversationId, messageObject._actionName, messageObject._data)) {
                        throw "Access Denied";
                    }
                    var methodObject = _lookupMethodObject(serviceEndPoint, messageObject);
                    var invokeCompleteCallback = new Microsoft.Office.Common.InvokeCompleteCallback(e.source, requesterUrl, messageObject._actionName, messageObject._conversationId, messageObject._correlationId, _postCallbackHandler, serializerVersion);
                    var invoker = new Microsoft.Office.Common.Invoker(methodObject, messageObject._data, invokeCompleteCallback, serviceEndPoint._eventHandlerProxyList, messageObject._conversationId, messageObject._actionName, serializerVersion);
                    var shouldEnque = true;
                    if (_messageProcessingTimer == null) {
                        if ((_lastMessageProcessTime == null || (new Date()).getTime() - _lastMessageProcessTime > _processInterval) && !_blockingFlag) {
                            _executeCommand(invoker);
                            shouldEnque = false;
                        }
                        else {
                            _messageProcessingTimer = setInterval(_dequeInvoker, _processInterval);
                        }
                    }
                    if (shouldEnque) {
                        _enqueInvoker(invoker);
                    }
                }
                catch (ex) {
                    if (serviceEndPoint && serviceEndPoint._onHandleRequestError) {
                        serviceEndPoint._onHandleRequestError(messageObject, ex);
                    }
                    var errorCode = Microsoft.Office.Common.InvokeResultCode.errorHandlingRequest;
                    if (ex == "Access Denied") {
                        errorCode = Microsoft.Office.Common.InvokeResultCode.errorHandlingRequestAccessDenied;
                    }
                    var callResponse = new Microsoft.Office.Common.Response(messageObject._actionName, messageObject._conversationId, messageObject._correlationId, errorCode, Microsoft.Office.Common.ResponseType.forCalling, ex);
                    var envelopedResult = Microsoft.Office.Common.MessagePackager.envelope(callResponse, serializerVersion);
                    if (e.source && e.source.postMessage) {
                        e.source.postMessage(envelopedResult, requesterUrl);
                    }
                }
            }
            else if (messageObject._messageType === Microsoft.Office.Common.MessageType.response) {
                var clientEndPoint = _lookupClientEndPoint(messageObject._conversationId);
                clientEndPoint._serializerVersion = serializerVersion;
                ;
                if (!_checkOrigin(clientEndPoint._targetUrl, e.origin)) {
                    throw "Failed orgin check";
                }
                if (messageObject._responseType === Microsoft.Office.Common.ResponseType.forCalling) {
                    var callbackEntry = clientEndPoint._callbackList[messageObject._correlationId];
                    if (callbackEntry) {
                        try {
                            if (callbackEntry.callback)
                                callbackEntry.callback(messageObject._errorCode, messageObject._data);
                        }
                        finally {
                            delete clientEndPoint._callbackList[messageObject._correlationId];
                        }
                    }
                }
                else {
                    var eventhandler = clientEndPoint._eventHandlerList[messageObject._actionName];
                    if (eventhandler !== undefined && eventhandler !== null) {
                        eventhandler(messageObject._data);
                    }
                }
            }
            else {
                return;
            }
        }
    }
    ;
    function _initialize() {
        if (!_initialized) {
            _registerListener(_receive);
            _initialized = true;
        }
    }
    ;
    return {
        connect: function Microsoft_Office_Common_XdmCommunicationManager$connect(conversationId, targetWindow, targetUrl, serializerVersion) {
            var clientEndPoint = _clientEndPoints[conversationId];
            if (!clientEndPoint) {
                _initialize();
                clientEndPoint = new Microsoft.Office.Common.ClientEndPoint(conversationId, targetWindow, targetUrl, serializerVersion);
                _clientEndPoints[conversationId] = clientEndPoint;
            }
            return clientEndPoint;
        },
        getClientEndPoint: function Microsoft_Office_Common_XdmCommunicationManager$getClientEndPoint(conversationId) {
            var e = Function._validateParams(arguments, [
                { name: "conversationId", type: String, mayBeNull: false }
            ]);
            if (e)
                throw e;
            return _clientEndPoints[conversationId];
        },
        createServiceEndPoint: function Microsoft_Office_Common_XdmCommunicationManager$createServiceEndPoint(serviceEndPointId) {
            _initialize();
            var serviceEndPoint = new Microsoft.Office.Common.ServiceEndPoint(serviceEndPointId);
            _serviceEndPoints[serviceEndPointId] = serviceEndPoint;
            return serviceEndPoint;
        },
        getServiceEndPoint: function Microsoft_Office_Common_XdmCommunicationManager$getServiceEndPoint(serviceEndPointId) {
            var e = Function._validateParams(arguments, [
                { name: "serviceEndPointId", type: String, mayBeNull: false }
            ]);
            if (e)
                throw e;
            return _serviceEndPoints[serviceEndPointId];
        },
        deleteClientEndPoint: function Microsoft_Office_Common_XdmCommunicationManager$deleteClientEndPoint(conversationId) {
            var e = Function._validateParams(arguments, [
                { name: "conversationId", type: String, mayBeNull: false }
            ]);
            if (e)
                throw e;
            delete _clientEndPoints[conversationId];
        },
        _setMethodTimeout: function Microsoft_Office_Common_XdmCommunicationManager$_setMethodTimeout(methodTimeout) {
            var e = Function._validateParams(arguments, [
                { name: "methodTimeout", type: Number, mayBeNull: false }
            ]);
            if (e)
                throw e;
            _methodTimeout = (methodTimeout <= 0) ? _methodTimeoutDefault : methodTimeout;
        },
        _startMethodTimeoutTimer: function Microsoft_Office_Common_XdmCommunicationManager$_startMethodTimeoutTimer() {
            if (!_methodTimeoutTimer) {
                _methodTimeoutTimer = setInterval(_checkMethodTimeout, _methodTimeoutProcessInterval);
            }
        }
    };
})();
Microsoft.Office.Common.Message = function Microsoft_Office_Common_Message(messageType, actionName, conversationId, correlationId, data) {
    var e = Function._validateParams(arguments, [{ name: "messageType", type: Number, mayBeNull: false },
        { name: "actionName", type: String, mayBeNull: false },
        { name: "conversationId", type: String, mayBeNull: false },
        { name: "correlationId", mayBeNull: false },
        { name: "data", mayBeNull: true, optional: true }
    ]);
    if (e)
        throw e;
    this._messageType = messageType;
    this._actionName = actionName;
    this._conversationId = conversationId;
    this._correlationId = correlationId;
    this._origin = window.location.href;
    if (typeof data == "undefined") {
        this._data = null;
    }
    else {
        this._data = data;
    }
};
Microsoft.Office.Common.Message.prototype = {
    getActionName: function Microsoft_Office_Common_Message$getActionName() {
        return this._actionName;
    },
    getConversationId: function Microsoft_Office_Common_Message$getConversationId() {
        return this._conversationId;
    },
    getCorrelationId: function Microsoft_Office_Common_Message$getCorrelationId() {
        return this._correlationId;
    },
    getOrigin: function Microsoft_Office_Common_Message$getOrigin() {
        return this._origin;
    },
    getData: function Microsoft_Office_Common_Message$getData() {
        return this._data;
    },
    getMessageType: function Microsoft_Office_Common_Message$getMessageType() {
        return this._messageType;
    }
};
Microsoft.Office.Common.Request = function Microsoft_Office_Common_Request(actionName, actionType, conversationId, correlationId, data) {
    Microsoft.Office.Common.Request.uber.constructor.call(this, Microsoft.Office.Common.MessageType.request, actionName, conversationId, correlationId, data);
    this._actionType = actionType;
};
OSF.OUtil.extend(Microsoft.Office.Common.Request, Microsoft.Office.Common.Message);
Microsoft.Office.Common.Request.prototype.getActionType = function Microsoft_Office_Common_Request$getActionType() {
    return this._actionType;
};
Microsoft.Office.Common.Response = function Microsoft_Office_Common_Response(actionName, conversationId, correlationId, errorCode, responseType, data) {
    Microsoft.Office.Common.Response.uber.constructor.call(this, Microsoft.Office.Common.MessageType.response, actionName, conversationId, correlationId, data);
    this._errorCode = errorCode;
    this._responseType = responseType;
};
OSF.OUtil.extend(Microsoft.Office.Common.Response, Microsoft.Office.Common.Message);
Microsoft.Office.Common.Response.prototype.getErrorCode = function Microsoft_Office_Common_Response$getErrorCode() {
    return this._errorCode;
};
Microsoft.Office.Common.Response.prototype.getResponseType = function Microsoft_Office_Common_Response$getResponseType() {
    return this._responseType;
};
Microsoft.Office.Common.MessagePackager = {
    envelope: function Microsoft_Office_Common_MessagePackager$envelope(messageObject, serializerVersion) {
        if (serializerVersion == OSF.SerializerVersion.Browser && (typeof (JSON) !== "undefined")) {
            if (typeof (messageObject) === "object") {
                messageObject._serializerVersion = serializerVersion;
            }
            return JSON.stringify(messageObject);
        }
        else {
            if (typeof (messageObject) === "object") {
                messageObject._serializerVersion = OSF.SerializerVersion.MsAjax;
            }
            return OsfMsAjaxFactory.msAjaxSerializer.serialize(messageObject);
        }
    },
    unenvelope: function Microsoft_Office_Common_MessagePackager$unenvelope(messageObject, serializerVersion) {
        if (serializerVersion == OSF.SerializerVersion.Browser && (typeof (JSON) !== "undefined")) {
            return JSON.parse(messageObject);
        }
        else {
            return OsfMsAjaxFactory.msAjaxSerializer.deserialize(messageObject, true);
        }
    }
};
Microsoft.Office.Common.ResponseSender = function Microsoft_Office_Common_ResponseSender(requesterWindow, requesterUrl, actionName, conversationId, correlationId, responseType, serializerVersion) {
    var e = Function._validateParams(arguments, [{ name: "requesterWindow", mayBeNull: false },
        { name: "requesterUrl", type: String, mayBeNull: false },
        { name: "actionName", type: String, mayBeNull: false },
        { name: "conversationId", type: String, mayBeNull: false },
        { name: "correlationId", mayBeNull: false },
        { name: "responsetype", type: Number, maybeNull: false },
        { name: "serializerVersion", type: Number, maybeNull: true, optional: true }
    ]);
    if (e)
        throw e;
    this._requesterWindow = requesterWindow;
    this._requesterUrl = requesterUrl;
    this._actionName = actionName;
    this._conversationId = conversationId;
    this._correlationId = correlationId;
    this._invokeResultCode = Microsoft.Office.Common.InvokeResultCode.noError;
    this._responseType = responseType;
    var me = this;
    this._send = function (result) {
        try {
            var response = new Microsoft.Office.Common.Response(me._actionName, me._conversationId, me._correlationId, me._invokeResultCode, me._responseType, result);
            var envelopedResult = Microsoft.Office.Common.MessagePackager.envelope(response, serializerVersion);
            me._requesterWindow.postMessage(envelopedResult, me._requesterUrl);
            ;
        }
        catch (ex) {
            OsfMsAjaxFactory.msAjaxDebug.trace("ResponseSender._send error:" + ex.message);
        }
    };
};
Microsoft.Office.Common.ResponseSender.prototype = {
    getRequesterWindow: function Microsoft_Office_Common_ResponseSender$getRequesterWindow() {
        return this._requesterWindow;
    },
    getRequesterUrl: function Microsoft_Office_Common_ResponseSender$getRequesterUrl() {
        return this._requesterUrl;
    },
    getActionName: function Microsoft_Office_Common_ResponseSender$getActionName() {
        return this._actionName;
    },
    getConversationId: function Microsoft_Office_Common_ResponseSender$getConversationId() {
        return this._conversationId;
    },
    getCorrelationId: function Microsoft_Office_Common_ResponseSender$getCorrelationId() {
        return this._correlationId;
    },
    getSend: function Microsoft_Office_Common_ResponseSender$getSend() {
        return this._send;
    },
    setResultCode: function Microsoft_Office_Common_ResponseSender$setResultCode(resultCode) {
        this._invokeResultCode = resultCode;
    }
};
Microsoft.Office.Common.InvokeCompleteCallback = function Microsoft_Office_Common_InvokeCompleteCallback(requesterWindow, requesterUrl, actionName, conversationId, correlationId, postCallbackHandler, serializerVersion) {
    Microsoft.Office.Common.InvokeCompleteCallback.uber.constructor.call(this, requesterWindow, requesterUrl, actionName, conversationId, correlationId, Microsoft.Office.Common.ResponseType.forCalling, serializerVersion);
    this._postCallbackHandler = postCallbackHandler;
    var me = this;
    this._send = function (result) {
        try {
            var response = new Microsoft.Office.Common.Response(me._actionName, me._conversationId, me._correlationId, me._invokeResultCode, me._responseType, result);
            var envelopedResult = Microsoft.Office.Common.MessagePackager.envelope(response, serializerVersion);
            me._requesterWindow.postMessage(envelopedResult, me._requesterUrl);
            me._postCallbackHandler();
        }
        catch (ex) {
            OsfMsAjaxFactory.msAjaxDebug.trace("InvokeCompleteCallback._send error:" + ex.message);
        }
    };
};
OSF.OUtil.extend(Microsoft.Office.Common.InvokeCompleteCallback, Microsoft.Office.Common.ResponseSender);
Microsoft.Office.Common.Invoker = function Microsoft_Office_Common_Invoker(methodObject, paramValue, invokeCompleteCallback, eventHandlerProxyList, conversationId, eventName, serializerVersion) {
    var e = Function._validateParams(arguments, [
        { name: "methodObject", mayBeNull: false },
        { name: "paramValue", mayBeNull: true },
        { name: "invokeCompleteCallback", mayBeNull: false },
        { name: "eventHandlerProxyList", mayBeNull: true },
        { name: "conversationId", type: String, mayBeNull: false },
        { name: "eventName", type: String, mayBeNull: false },
        { name: "serializerVersion", type: Number, mayBeNull: true, optional: true }
    ]);
    if (e)
        throw e;
    this._methodObject = methodObject;
    this._param = paramValue;
    this._invokeCompleteCallback = invokeCompleteCallback;
    this._eventHandlerProxyList = eventHandlerProxyList;
    this._conversationId = conversationId;
    this._eventName = eventName;
    this._serializerVersion = serializerVersion;
};
Microsoft.Office.Common.Invoker.prototype = {
    invoke: function Microsoft_Office_Common_Invoker$invoke() {
        try {
            var result;
            switch (this._methodObject.getInvokeType()) {
                case Microsoft.Office.Common.InvokeType.async:
                    this._methodObject.getMethod()(this._param, this._invokeCompleteCallback.getSend());
                    break;
                case Microsoft.Office.Common.InvokeType.sync:
                    result = this._methodObject.getMethod()(this._param);
                    this._invokeCompleteCallback.getSend()(result);
                    break;
                case Microsoft.Office.Common.InvokeType.syncRegisterEvent:
                    var eventHandlerProxy = this._createEventHandlerProxyObject(this._invokeCompleteCallback);
                    result = this._methodObject.getMethod()(eventHandlerProxy.getSend(), this._param);
                    this._eventHandlerProxyList[this._conversationId + this._eventName] = eventHandlerProxy.getSend();
                    this._invokeCompleteCallback.getSend()(result);
                    break;
                case Microsoft.Office.Common.InvokeType.syncUnregisterEvent:
                    var eventHandler = this._eventHandlerProxyList[this._conversationId + this._eventName];
                    result = this._methodObject.getMethod()(eventHandler, this._param);
                    delete this._eventHandlerProxyList[this._conversationId + this._eventName];
                    this._invokeCompleteCallback.getSend()(result);
                    break;
                case Microsoft.Office.Common.InvokeType.asyncRegisterEvent:
                    var eventHandlerProxyAsync = this._createEventHandlerProxyObject(this._invokeCompleteCallback);
                    this._methodObject.getMethod()(eventHandlerProxyAsync.getSend(), this._invokeCompleteCallback.getSend(), this._param);
                    this._eventHandlerProxyList[this._callerId + this._eventName] = eventHandlerProxyAsync.getSend();
                    break;
                case Microsoft.Office.Common.InvokeType.asyncUnregisterEvent:
                    var eventHandlerAsync = this._eventHandlerProxyList[this._callerId + this._eventName];
                    this._methodObject.getMethod()(eventHandlerAsync, this._invokeCompleteCallback.getSend(), this._param);
                    delete this._eventHandlerProxyList[this._callerId + this._eventName];
                    break;
                default:
                    break;
            }
        }
        catch (ex) {
            this._invokeCompleteCallback.setResultCode(Microsoft.Office.Common.InvokeResultCode.errorInResponse);
            this._invokeCompleteCallback.getSend()(ex);
        }
    },
    getInvokeBlockingFlag: function Microsoft_Office_Common_Invoker$getInvokeBlockingFlag() {
        return this._methodObject.getBlockingFlag();
    },
    _createEventHandlerProxyObject: function Microsoft_Office_Common_Invoker$_createEventHandlerProxyObject(invokeCompleteObject) {
        return new Microsoft.Office.Common.ResponseSender(invokeCompleteObject.getRequesterWindow(), invokeCompleteObject.getRequesterUrl(), invokeCompleteObject.getActionName(), invokeCompleteObject.getConversationId(), invokeCompleteObject.getCorrelationId(), Microsoft.Office.Common.ResponseType.forEventing, this._serializerVersion);
    }
};
OSF.OUtil.setNamespace("OSF", window);
OSF.AppName = {
    Unsupported: 0,
    Excel: 1,
    Word: 2,
    PowerPoint: 4,
    Outlook: 8,
    ExcelWebApp: 16,
    WordWebApp: 32,
    OutlookWebApp: 64,
    Project: 128,
    AccessWebApp: 256,
    PowerpointWebApp: 512,
    ExcelIOS: 1024,
    Sway: 2048,
    WordIOS: 4096,
    PowerPointIOS: 8192,
    Access: 16384,
    Lync: 32768,
    OutlookIOS: 65536,
    OneNoteWebApp: 131072,
    OneNote: 262144
};
OSF.InternalPerfMarker = {
    DataCoercionBegin: "Agave.HostCall.CoerceDataStart",
    DataCoercionEnd: "Agave.HostCall.CoerceDataEnd"
};
OSF.HostCallPerfMarker = {
    IssueCall: "Agave.HostCall.IssueCall",
    ReceiveResponse: "Agave.HostCall.ReceiveResponse",
    RuntimeExceptionRaised: "Agave.HostCall.RuntimeExecptionRaised"
};
OSF.AgaveHostAction = {
    "Select": 0,
    "UnSelect": 1,
    "CancelDialog": 2,
    "InsertAgave": 3,
    "CtrlF6In": 4,
    "CtrlF6Exit": 5,
    "CtrlF6ExitShift": 6,
    "SelectWithError": 7,
    "NotifyHostError": 8,
    "RefreshAddinCommands": 9
};
OSF.SharedConstants = {
    "NotificationConversationIdSuffix": '_ntf'
};
OSF.DialogMessageType = {
    DialogMessageReceived: 0,
    DialogClosed: 1,
    NavigationFailed: 2,
    InvalidSchema: 3
};
OSF.OfficeAppContext = function OSF_OfficeAppContext(id, appName, appVersion, appUILocale, dataLocale, docUrl, clientMode, settings, reason, osfControlType, eToken, correlationId, appInstanceId, touchEnabled, commerceAllowed, appMinorVersion, requirementMatrix) {
    this._id = id;
    this._appName = appName;
    this._appVersion = appVersion;
    this._appUILocale = appUILocale;
    this._dataLocale = dataLocale;
    this._docUrl = docUrl;
    this._clientMode = clientMode;
    this._settings = settings;
    this._reason = reason;
    this._osfControlType = osfControlType;
    this._eToken = eToken;
    this._correlationId = correlationId;
    this._appInstanceId = appInstanceId;
    this._touchEnabled = touchEnabled;
    this._commerceAllowed = commerceAllowed;
    this._appMinorVersion = appMinorVersion;
    this._requirementMatrix = requirementMatrix;
    this._isDialog = false;
    this.get_id = function get_id() { return this._id; };
    this.get_appName = function get_appName() { return this._appName; };
    this.get_appVersion = function get_appVersion() { return this._appVersion; };
    this.get_appUILocale = function get_appUILocale() { return this._appUILocale; };
    this.get_dataLocale = function get_dataLocale() { return this._dataLocale; };
    this.get_docUrl = function get_docUrl() { return this._docUrl; };
    this.get_clientMode = function get_clientMode() { return this._clientMode; };
    this.get_bindings = function get_bindings() { return this._bindings; };
    this.get_settings = function get_settings() { return this._settings; };
    this.get_reason = function get_reason() { return this._reason; };
    this.get_osfControlType = function get_osfControlType() { return this._osfControlType; };
    this.get_eToken = function get_eToken() { return this._eToken; };
    this.get_correlationId = function get_correlationId() { return this._correlationId; };
    this.get_appInstanceId = function get_appInstanceId() { return this._appInstanceId; };
    this.get_touchEnabled = function get_touchEnabled() { return this._touchEnabled; };
    this.get_commerceAllowed = function get_commerceAllowed() { return this._commerceAllowed; };
    this.get_appMinorVersion = function get_appMinorVersion() { return this._appMinorVersion; };
    this.get_requirementMatrix = function get_requirementMatrix() { return this._requirementMatrix; };
    this.get_isDialog = function get_isDialog() { return this._isDialog; };
};
OSF.OsfControlType = {
    DocumentLevel: 0,
    ContainerLevel: 1
};
OSF.ClientMode = {
    ReadOnly: 0,
    ReadWrite: 1
};
OSF.OUtil.setNamespace("Microsoft", window);
OSF.OUtil.setNamespace("Office", Microsoft);
OSF.OUtil.setNamespace("Client", Microsoft.Office);
OSF.OUtil.setNamespace("WebExtension", Microsoft.Office);
Microsoft.Office.WebExtension.InitializationReason = {
    Inserted: "inserted",
    DocumentOpened: "documentOpened"
};
Microsoft.Office.WebExtension.ValueFormat = {
    Unformatted: "unformatted",
    Formatted: "formatted"
};
Microsoft.Office.WebExtension.FilterType = {
    All: "all"
};
Microsoft.Office.WebExtension.Parameters = {
    BindingType: "bindingType",
    CoercionType: "coercionType",
    ValueFormat: "valueFormat",
    FilterType: "filterType",
    Columns: "columns",
    SampleData: "sampleData",
    GoToType: "goToType",
    SelectionMode: "selectionMode",
    Id: "id",
    PromptText: "promptText",
    ItemName: "itemName",
    FailOnCollision: "failOnCollision",
    StartRow: "startRow",
    StartColumn: "startColumn",
    RowCount: "rowCount",
    ColumnCount: "columnCount",
    Callback: "callback",
    AsyncContext: "asyncContext",
    Data: "data",
    Rows: "rows",
    OverwriteIfStale: "overwriteIfStale",
    FileType: "fileType",
    EventType: "eventType",
    Handler: "handler",
    SliceSize: "sliceSize",
    SliceIndex: "sliceIndex",
    ActiveView: "activeView",
    Status: "status",
    Xml: "xml",
    Namespace: "namespace",
    Prefix: "prefix",
    XPath: "xPath",
    Text: "text",
    ImageLeft: "imageLeft",
    ImageTop: "imageTop",
    ImageWidth: "imageWidth",
    ImageHeight: "imageHeight",
    TaskId: "taskId",
    FieldId: "fieldId",
    FieldValue: "fieldValue",
    ServerUrl: "serverUrl",
    ListName: "listName",
    ResourceId: "resourceId",
    ViewType: "viewType",
    ViewName: "viewName",
    GetRawValue: "getRawValue",
    CellFormat: "cellFormat",
    TableOptions: "tableOptions",
    TaskIndex: "taskIndex",
    ResourceIndex: "resourceIndex",
    Url: "url",
    MessageHandler: "messageHandler",
    Width: "width",
    Height: "height",
    RequireHTTPs: "requireHTTPS",
    MessageToParent: "messageToParent",
    XFrameDenySafe: "xFrameDenySafe"
};
OSF.OUtil.setNamespace("DDA", OSF);
OSF.DDA.DocumentMode = {
    ReadOnly: 1,
    ReadWrite: 0
};
OSF.DDA.PropertyDescriptors = {
    AsyncResultStatus: "AsyncResultStatus"
};
OSF.DDA.EventDescriptors = {};
OSF.DDA.ListDescriptors = {};
OSF.DDA.UI = {};
OSF.DDA.getXdmEventName = function OSF_DDA$GetXdmEventName(bindingId, eventType) {
    if (eventType == Microsoft.Office.WebExtension.EventType.BindingSelectionChanged || eventType == Microsoft.Office.WebExtension.EventType.BindingDataChanged) {
        return bindingId + "_" + eventType;
    }
    else {
        return eventType;
    }
};
OSF.DDA.MethodDispId = {
    dispidMethodMin: 64,
    dispidGetSelectedDataMethod: 64,
    dispidSetSelectedDataMethod: 65,
    dispidAddBindingFromSelectionMethod: 66,
    dispidAddBindingFromPromptMethod: 67,
    dispidGetBindingMethod: 68,
    dispidReleaseBindingMethod: 69,
    dispidGetBindingDataMethod: 70,
    dispidSetBindingDataMethod: 71,
    dispidAddRowsMethod: 72,
    dispidClearAllRowsMethod: 73,
    dispidGetAllBindingsMethod: 74,
    dispidLoadSettingsMethod: 75,
    dispidSaveSettingsMethod: 76,
    dispidGetDocumentCopyMethod: 77,
    dispidAddBindingFromNamedItemMethod: 78,
    dispidAddColumnsMethod: 79,
    dispidGetDocumentCopyChunkMethod: 80,
    dispidReleaseDocumentCopyMethod: 81,
    dispidNavigateToMethod: 82,
    dispidGetActiveViewMethod: 83,
    dispidGetDocumentThemeMethod: 84,
    dispidGetOfficeThemeMethod: 85,
    dispidGetFilePropertiesMethod: 86,
    dispidClearFormatsMethod: 87,
    dispidSetTableOptionsMethod: 88,
    dispidSetFormatsMethod: 89,
    dispidExecuteRichApiRequestMethod: 93,
    dispidAppCommandInvocationCompletedMethod: 94,
    dispidAddDataPartMethod: 128,
    dispidGetDataPartByIdMethod: 129,
    dispidGetDataPartsByNamespaceMethod: 130,
    dispidGetDataPartXmlMethod: 131,
    dispidGetDataPartNodesMethod: 132,
    dispidDeleteDataPartMethod: 133,
    dispidGetDataNodeValueMethod: 134,
    dispidGetDataNodeXmlMethod: 135,
    dispidGetDataNodesMethod: 136,
    dispidSetDataNodeValueMethod: 137,
    dispidSetDataNodeXmlMethod: 138,
    dispidAddDataNamespaceMethod: 139,
    dispidGetDataUriByPrefixMethod: 140,
    dispidGetDataPrefixByUriMethod: 141,
    dispidGetDataNodeTextMethod: 142,
    dispidSetDataNodeTextMethod: 143,
    dispidMethodMax: 143,
    dispidGetSelectedTaskMethod: 110,
    dispidGetSelectedResourceMethod: 111,
    dispidGetTaskMethod: 112,
    dispidGetResourceFieldMethod: 113,
    dispidGetWSSUrlMethod: 114,
    dispidGetTaskFieldMethod: 115,
    dispidGetProjectFieldMethod: 116,
    dispidGetSelectedViewMethod: 117,
    dispidGetTaskByIndexMethod: 118,
    dispidGetResourceByIndexMethod: 119,
    dispidSetTaskFieldMethod: 120,
    dispidSetResourceFieldMethod: 121,
    dispidGetMaxTaskIndexMethod: 122,
    dispidGetMaxResourceIndexMethod: 123,
    dispidMessageParentMethod: 900
};
OSF.DDA.EventDispId = {
    dispidEventMin: 0,
    dispidInitializeEvent: 0,
    dispidSettingsChangedEvent: 1,
    dispidDocumentSelectionChangedEvent: 2,
    dispidBindingSelectionChangedEvent: 3,
    dispidBindingDataChangedEvent: 4,
    dispidDocumentOpenEvent: 5,
    dispidDocumentCloseEvent: 6,
    dispidActiveViewChangedEvent: 7,
    dispidDocumentThemeChangedEvent: 8,
    dispidOfficeThemeChangedEvent: 9,
    dispidDialogMessageReceivedEvent: 10,
    dispidActivationStatusChangedEvent: 32,
    dispidAppCommandInvokedEvent: 39,
    dispidTaskSelectionChangedEvent: 56,
    dispidResourceSelectionChangedEvent: 57,
    dispidViewSelectionChangedEvent: 58,
    dispidDataNodeAddedEvent: 60,
    dispidDataNodeReplacedEvent: 61,
    dispidDataNodeDeletedEvent: 62,
    dispidEventMax: 63
};
OSF.XmlConstants = {
    MaxXmlSize: 1048576,
    MaxElementDepth: 64
};
OSF.Xpath3Provider = function OSF_Xpath3Provider(xml, xmlNamespaces) {
    this._xmldoc = new DOMParser().parseFromString(xml, "text/xml");
    this._evaluator = new XPathEvaluator();
    this._namespaceMapping = {};
    this._defaultNamespace = null;
    var namespaces = xmlNamespaces.split(' ');
    var matches;
    for (var i = 0; i < namespaces.length; ++i) {
        matches = /xmlns="([^"]*)"/g.exec(namespaces[i]);
        if (matches) {
            this._defaultNamespace = matches[1];
            continue;
        }
        matches = /xmlns:([^=]*)="([^"]*)"/g.exec(namespaces[i]);
        if (matches) {
            this._namespaceMapping[matches[1]] = matches[2];
            continue;
        }
    }
    this._resolver = this;
};
OSF.Xpath3Provider.prototype = {
    lookupNamespaceURI: function OSF_Xpath3Provider$lookupNamespaceURI(prefix) {
        var ns = this._namespaceMapping[prefix];
        return ns || this._defaultNamespace;
    },
    selectSingleNode: function OSF_Xpath3Provider$selectSingleNode(name, contextNode) {
        var xpath = (contextNode ? "./" : "/") + name;
        contextNode = contextNode || this.getDocumentElement();
        var result = this._evaluator.evaluate(xpath, contextNode, this._resolver, XPathResult.FIRST_ORDERED_NODE_TYPE, null);
        if (result) {
            return result.singleNodeValue;
        }
        else {
            return null;
        }
    },
    selectNodes: function OSF_Xpath3Provider$selectNodes(name, contextNode) {
        var xpath = (contextNode ? "./" : "/") + name;
        contextNode = contextNode || this.getDocumentElement();
        var result = this._evaluator.evaluate(xpath, contextNode, this._resolver, XPathResult.ORDERED_NODE_ITERATOR_TYPE, null);
        var nodes = [];
        if (result) {
            var node = result.iterateNext();
            while (node) {
                nodes.push(node);
                node = result.iterateNext();
            }
        }
        return nodes;
    },
    getDocumentElement: function OSF_Xpath3Provider$getDocumentElement() {
        return this._xmldoc.documentElement;
    }
};
OSF.IEXpathProvider = function OSF_IEXpathProvider(xml, xmlNamespaces) {
    var xmldoc = null;
    var msxmlVersions = ['MSXML2.DOMDocument.6.0'];
    for (var i = 0; i < msxmlVersions.length; i++) {
        try {
            xmldoc = new ActiveXObject(msxmlVersions[i]);
            xmldoc.setProperty('ResolveExternals', false);
            xmldoc.setProperty('ValidateOnParse', false);
            xmldoc.setProperty('ProhibitDTD', true);
            xmldoc.setProperty('MaxXMLSize', OSF.XmlConstants.MaxXmlSize);
            xmldoc.setProperty('MaxElementDepth', OSF.XmlConstants.MaxElementDepth);
            xmldoc.async = false;
            xmldoc.loadXML(xml);
            xmldoc.setProperty("SelectionLanguage", "XPath");
            xmldoc.setProperty("SelectionNamespaces", xmlNamespaces);
            break;
        }
        catch (ex) {
            OsfMsAjaxFactory.msAjaxDebug.trace("xml doc creating error:" + ex);
        }
    }
    this._xmldoc = xmldoc;
};
OSF.IEXpathProvider.prototype = {
    selectSingleNode: function OSF_IEXpathProvider$selectSingleNode(name, contextNode) {
        var xpath = (contextNode ? "./" : "/") + name;
        contextNode = contextNode || this.getDocumentElement();
        return contextNode.selectSingleNode(xpath);
    },
    selectNodes: function OSF_IEXpathProvider$selectNodes(name, contextNode) {
        var xpath = (contextNode ? "./" : "/") + name;
        contextNode = contextNode || this.getDocumentElement();
        return contextNode.selectNodes(xpath);
    },
    getDocumentElement: function OSF_IEXpathProvider$getDocumentElement() {
        return this._xmldoc.documentElement;
    },
    getActiveXObject: function OSF_IEXpathProvider$getActiveXObject() {
        return this._xmldoc;
    }
};
OSF.DomParserProvider = function OSF_DomParserProvider(xml, xmlNamespaces) {
    try {
        this._xmldoc = new DOMParser().parseFromString(xml, "text/xml");
    }
    catch (ex) {
        Sys.Debug.trace("xml doc creating error:" + ex);
    }
    this._namespaceMapping = {};
    this._defaultNamespace = null;
    var namespaces = xmlNamespaces.split(' ');
    var matches;
    for (var i = 0; i < namespaces.length; ++i) {
        matches = /xmlns="([^"]*)"/g.exec(namespaces[i]);
        if (matches) {
            this._defaultNamespace = matches[1];
            continue;
        }
        matches = /xmlns:([^=]*)="([^"]*)"/g.exec(namespaces[i]);
        if (matches) {
            this._namespaceMapping[matches[1]] = matches[2];
            continue;
        }
    }
};
OSF.DomParserProvider.prototype = {
    selectSingleNode: function OSF_DomParserProvider$selectSingleNode(name, contextNode) {
        var selectedNode = contextNode || this._xmldoc;
        var nodes = this._selectNodes(name, selectedNode);
        if (nodes.length === 0)
            return null;
        return nodes[0];
    },
    selectNodes: function OSF_DomParserProvider$selectNodes(name, contextNode) {
        var selectedNode = contextNode || this._xmldoc;
        return this._selectNodes(name, selectedNode);
    },
    _selectNodes: function OSF_DomParserProvider$_selectNodes(name, contextNode) {
        var nodes = [];
        if (!name)
            return nodes;
        var nameInfo = name.split(":");
        var ns, nodeName;
        if (nameInfo.length === 1) {
            ns = null;
            nodeName = nameInfo[0];
        }
        else if (nameInfo.length === 2) {
            ns = this._namespaceMapping[nameInfo[0]];
            nodeName = nameInfo[1];
        }
        else {
            throw OsfMsAjaxFactory.msAjaxError.argument("name");
        }
        if (!contextNode.hasChildNodes)
            return nodes;
        var childs = contextNode.childNodes;
        for (var i = 0; i < childs.length; i++) {
            if (nodeName === this._removeNodePrefix(childs[i].nodeName) && (ns === childs[i].namespaceURI)) {
                nodes.push(childs[i]);
            }
        }
        return nodes;
    },
    _removeNodePrefix: function OSF_DomParserProvider$_removeNodePrefix(nodeName) {
        var nodeInfo = nodeName.split(':');
        if (nodeInfo.length === 1) {
            return nodeName;
        }
        else {
            return nodeInfo[1];
        }
    },
    getDocumentElement: function OSF_DomParserProvider$getDocumentElement() {
        return this._xmldoc.documentElement;
    }
};
OSF.XmlProcessor = function OSF_XmlProcessor(xml, xmlNamespaces) {
    var e = Function._validateParams(arguments, [
        { name: "xml", type: String, mayBeNull: false },
        { name: "xmlNamespaces", type: String, mayBeNull: false }
    ]);
    if (e)
        throw e;
    if (document.implementation && document.implementation.hasFeature("XPath", "3.0")) {
        this._provider = new OSF.Xpath3Provider(xml, xmlNamespaces);
    }
    else {
        this._provider = new OSF.IEXpathProvider(xml, xmlNamespaces);
        if (!this._provider.getActiveXObject()) {
            this._provider = new OSF.DomParserProvider(xml, xmlNamespaces);
        }
    }
};
OSF.XmlProcessor.prototype = {
    selectSingleNode: function OSF_XmlProcessor$selectSingleNode(name, contextNode) {
        var e = Function._validateParams(arguments, [{ name: "name", type: String, mayBeNull: false },
            { name: "contextNode", mayBeNull: true, optional: true }
        ]);
        if (e)
            throw e;
        return this._provider.selectSingleNode(name, contextNode);
    },
    selectNodes: function OSF_XmlProcessor$selectNodes(name, contextNode) {
        var e = Function._validateParams(arguments, [{ name: "name", type: String, mayBeNull: false },
            { name: "contextNode", mayBeNull: true, optional: true }
        ]);
        if (e)
            throw e;
        return this._provider.selectNodes(name, contextNode);
    },
    getDocumentElement: function OSF_XmlProcessor$getDocumentElement() {
        return this._provider.getDocumentElement();
    },
    getNodeValue: function OSF_XmlProcessor$getNodeValue(node) {
        var e = Function._validateParams(arguments, [
            { name: "node", type: Object, mayBeNull: false }
        ]);
        if (e)
            throw e;
        var nodeValue;
        if (node.text) {
            nodeValue = node.text;
        }
        else {
            nodeValue = node.textContent;
        }
        return nodeValue;
    },
    getNodeXml: function OSF_XmlProcessor$getNodeXml(node) {
        var e = Function._validateParams(arguments, [
            { name: "node", type: Object, mayBeNull: false }
        ]);
        if (e)
            throw e;
        var nodeXml;
        if (node.xml) {
            nodeXml = node.xml;
        }
        else {
            nodeXml = new XMLSerializer().serializeToString(node);
        }
        return nodeXml;
    },
    getNodeNamespaceURI: function OSF_XmlProcessor$getNodeNamespaceURI(node) {
        var e = Function._validateParams(arguments, [
            { name: "node", type: Object, mayBeNull: false }
        ]);
        if (e)
            throw e;
        return node.namespaceURI;
    },
    getNodePrefix: function OSF_XmlProcessor$getNodePrefix(node) {
        var e = Function._validateParams(arguments, [
            { name: "node", type: Object, mayBeNull: false }
        ]);
        if (e)
            throw e;
        return node.prefix;
    },
    getNodeBaseName: function OSF_XmlProcessor$getNodeBaseName(node) {
        var e = Function._validateParams(arguments, [
            { name: "node", type: Object, mayBeNull: false }
        ]);
        if (e)
            throw e;
        var nodeBaseName;
        if (node.baseName) {
            nodeBaseName = node.baseName;
        }
        else {
            nodeBaseName = node.localName;
        }
        return nodeBaseName;
    },
    getNodeType: function OSF_XmlProcessor$getNodeType(node) {
        var e = Function._validateParams(arguments, [
            { name: "node", type: Object, mayBeNull: false }
        ]);
        if (e)
            throw e;
        return node.nodeType;
    },
    _getAttributeLocalName: function OSF_XmlProcessor$_getAttributeLocalName(attribute) {
        var localName;
        if (attribute.localName) {
            localName = attribute.localName;
        }
        else {
            localName = attribute.baseName;
        }
        return localName;
    },
    readAttributes: function OSF_XmlProcessor$readAttributes(node, attributesToRead, objectToFill) {
        var e = Function._validateParams(arguments, [
            { name: "node", type: Object, mayBeNull: false },
            { name: "attributesToRead", type: Object, mayBeNull: false },
            { name: "objectToFill", type: Object, mayBeNull: false }
        ]);
        if (e)
            throw e;
        var attribute;
        var localName;
        for (var i = 0; i < node.attributes.length; i++) {
            attribute = node.attributes[i];
            localName = this._getAttributeLocalName(attribute);
            for (var p in attributesToRead) {
                if (localName === p) {
                    objectToFill[attributesToRead[p]] = attribute.value;
                }
            }
        }
    }
};
var OfficeExt;
(function (OfficeExt) {
    var SafeSerializer = (function () {
        function SafeSerializer() {
        }
        SafeSerializer.prototype.Serialize = function (value) {
            try {
                if (typeof (JSON) !== "undefined") {
                    return JSON.stringify(value);
                }
                else {
                    return OsfMsAjaxFactory.msAjaxSerializer.serialize(value);
                }
            }
            catch (e) {
                return null;
            }
        };
        SafeSerializer.prototype.Deserialize = function (value) {
            try {
                if (typeof (JSON) !== "undefined") {
                    return JSON.parse(value);
                }
                else {
                    return OsfMsAjaxFactory.msAjaxSerializer.deserialize(value, true);
                }
            }
            catch (e) {
                return null;
            }
        };
        return SafeSerializer;
    })();
    OfficeExt.SafeSerializer = SafeSerializer;
    var AppsDataCacheManager = (function () {
        function AppsDataCacheManager(localStorage, serializer) {
            this._localStorage = localStorage;
            this._serializer = serializer;
        }
        AppsDataCacheManager.prototype.GetCacheItem = function (key, checkRefreshRate, errors) {
            if (checkRefreshRate === void 0) { checkRefreshRate = true; }
            this.ValidateCurrentCache();
            var value = this._localStorage.getItem(key);
            if (value) {
                var cacheEntry = this._serializer.Deserialize(value);
                if (checkRefreshRate) {
                    var now = new Date();
                    if (Math.abs(now.getTime() - cacheEntry.createdOn) < AppsDataCacheManager.msPerDay * cacheEntry.refreshRate) {
                        return cacheEntry.data;
                    }
                    else {
                        this._localStorage.removeItem(key);
                        if (errors) {
                            errors['cacheExpired'] = true;
                        }
                    }
                }
                else {
                    return cacheEntry.data;
                }
            }
        };
        AppsDataCacheManager.prototype.SetCacheItem = function (key, value, refreshRateInDays) {
            refreshRateInDays = refreshRateInDays || AppsDataCacheManager.defaultRefreshRateInDays;
            var now = new Date();
            var cacheEntry = { 'data': value, 'createdOn': now.getTime(), 'refreshRate': refreshRateInDays };
            this._localStorage.setItem(key, this._serializer.Serialize(cacheEntry));
        };
        AppsDataCacheManager.prototype.RemoveCacheItem = function (key) {
            this._localStorage.removeItem(key);
        };
        AppsDataCacheManager.prototype.RemoveAll = function (keyPrefix) {
            var keysToRemove = this._localStorage.getKeysWithPrefix(keyPrefix);
            for (var i = 0, len = keysToRemove.length; i < len; i++) {
                this._localStorage.removeItem(keysToRemove[i]);
            }
        };
        AppsDataCacheManager.prototype.RemoveMatches = function (regexPatterns) {
            var keys = this._localStorage.getKeysWithPrefix("");
            for (var i = 0, len = keys.length; i < len; i++) {
                var key = keys[i];
                for (var j = 0, lenRegex = regexPatterns.length; j < lenRegex; j++) {
                    if (regexPatterns[j].test(key)) {
                        this._localStorage.removeItem(key);
                        break;
                    }
                }
            }
        };
        AppsDataCacheManager.prototype.ValidateCurrentCache = function () {
            var cacheVersion = this._localStorage.getItem(AppsDataCacheManager.cacheVersionKey);
            if (cacheVersion != AppsDataCacheManager.currentSchemaVersion) {
                this.RemoveMatches([new RegExp("__OSF_(?!.*activated).*$", "i")]);
                this._localStorage.setItem(AppsDataCacheManager.cacheVersionKey, AppsDataCacheManager.currentSchemaVersion);
            }
        };
        AppsDataCacheManager.defaultRefreshRateInDays = 3;
        AppsDataCacheManager.msPerDay = 86400000;
        AppsDataCacheManager.currentSchemaVersion = "1";
        AppsDataCacheManager.checkedCache = false;
        AppsDataCacheManager.cacheVersionKey = "osfCacheVersion";
        return AppsDataCacheManager;
    })();
    OfficeExt.AppsDataCacheManager = AppsDataCacheManager;
})(OfficeExt || (OfficeExt = {}));
var OmexDataProvider = (function () {
    function OmexDataProvider(cacheManager) {
        this._cacheManager = cacheManager;
        this._cid = "0";
    }
    OmexDataProvider.GetInstance = function (appsDataCacheManager) {
        if (OmexDataProvider.instance == null) {
            OmexDataProvider.instance = new OmexDataProvider(appsDataCacheManager);
        }
        return OmexDataProvider.instance;
    };
    OmexDataProvider.prototype.validateAndCacheNewManifest = function (contextActivationMgr, assetId, appVersion, manifestStr, manifest) {
        return false;
    };
    OmexDataProvider.prototype.localCacheAddinCommands = function (contextActivationMgr, assetId, version, storeId) {
    };
    OmexDataProvider.prototype.getEntitlementsAsync = function (contextActivationMgr, forAddinCommands, onGetEntitlementCompleted) {
        var myAddins;
        var onGetOmexEntitlementsCompleted = function (asyncResult) {
            if (asyncResult != null && asyncResult.errorCode === OSF.InvokeResultCode.S_OK && asyncResult.value) {
                var entitlementsResult = asyncResult.value;
                myAddins = new Array();
                for (var i = 0; i < entitlementsResult.length; i++) {
                    var item = entitlementsResult[i];
                    var entitlement = {
                        assetId: item[1],
                        appVersion: item[4],
                        storeId: item[11],
                        storeType: OSF.StoreType.OMEX,
                        targetType: item[3]
                    };
                    myAddins.push(entitlement);
                }
            }
            onGetEntitlementCompleted(myAddins);
        };
        var params = { "storeType": OSF.StoreTypeEnum.MarketPlace, "refresh": false, "fromInsertDialog": false, "callbackFunc": onGetOmexEntitlementsCompleted };
        contextActivationMgr.queryEntitlementDetails(params);
    };
    OmexDataProvider.prototype.getManifestAsync = function (context, onGetManifestCompleted) {
        OSF.OUtil.validateParamObject(context, {
            "contextActivationMgr": { type: Object, mayBeNull: false },
            "referenceInUse": { type: Object, mayBeNull: false }
        }, null);
        var me = this;
        var onGetManifestCompletedByAnonymousProxy = function (asyncResult) {
            if (asyncResult != null && asyncResult.statusCode === OSF.ProxyCallStatusCode.Succeeded && asyncResult.value) {
                var manifestAndEToken = asyncResult.value;
                var manifest = new OSF.Manifest.Manifest(manifestAndEToken.manifest, context.contextActivationMgr.getAppUILocale());
                if (!asyncResult.cached) {
                    me.SetManifestAndETokenCache(context, manifestAndEToken);
                }
                OSF.OsfManifestManager.cacheManifest(context.referenceInUse.id, context.referenceInUse.version, manifest);
                onGetManifestCompleted({ "statusCode": OSF.ProxyCallStatusCode.Succeeded, "value": manifest, "context": context });
            }
            else {
                onGetManifestCompleted({ "statusCode": OSF.ProxyCallStatusCode.Failed, "value": null });
            }
        };
        var createAnonymousOmexProxyCompleted = function (asyncResult) {
            if (context.contextActivationMgr._omexAnonymousWSProxy.isReady && context.contextActivationMgr._omexAnonymousWSProxy.clientEndPoint != null) {
                context.anonymous = true;
                context.clientEndPoint = context.contextActivationMgr._omexAnonymousWSProxy.clientEndPoint;
                context.manifestManager = OSF.OsfManifestManager;
                var params = {
                    "applicationName": context.contextActivationMgr._hostType,
                    "assetID": context.referenceInUse.id,
                    "build": OSF.Constants.FileVersion,
                    "clientName": context.contextActivationMgr._getClientNameForOmex(),
                    "clientVersion": context.contextActivationMgr._getClientVersionForOmex(),
                    "clientAppInfoReturnType": OSF.ClientAppInfoReturnType.urlOnly,
                    "errors": {},
                    "correlationId": null,
                    "contentMarket": context.referenceInUse.storeLocator,
                    "clearManifest": false
                };
                me.GetManifestAndEToken(context, params, onGetManifestCompletedByAnonymousProxy);
            }
            else {
                onGetManifestCompleted({ "statusCode": OSF.ProxyCallStatusCode.Failed, "value": null });
            }
        };
        var cachedManifest = OSF.OsfManifestManager.getCachedManifest(context.referenceInUse.id, context.referenceInUse.version);
        if (cachedManifest == null) {
            context.contextActivationMgr._createOmexProxy(context.contextActivationMgr._omexAnonymousWSProxy, createAnonymousOmexProxyCompleted);
        }
        else {
            onGetManifestCompleted({ "statusCode": OSF.ProxyCallStatusCode.Succeeded, "value": cachedManifest, "context": context });
        }
    };
    OmexDataProvider.prototype.GetCacheKeyPrefix = function (context) {
        OSF.OUtil.validateParamObject(context, {
            "anonymous": { type: Boolean, mayBeNull: true }
        }, null);
        if (context.anonymous == null || context.anonymous) {
            return OmexDataProvider.anonymousCacheKeyPrefix;
        }
        return OmexDataProvider.gatedCacheKeyPrefix;
    };
    OmexDataProvider.prototype.GetCustomerId = function () {
        return this._cid;
    };
    OmexDataProvider.prototype.SetCustomerId = function (cid) {
        this._cid = cid;
    };
    OmexDataProvider.prototype.AllCached = function (context, params) {
        if (context && !context.clearCache && !context.clearToken && !context.clearEntitlement && !context.clearAppState && !context.clearManifest) {
            if (this.KilledAppsCached(context) && this.AppStateCached(context, params) && this.ManifestAndETokenCached(context, params)) {
                return true;
            }
        }
        return false;
    };
    OmexDataProvider.prototype.KilledAppsCached = function (context) {
        var cacheKeyPrefix = this.GetCacheKeyPrefix(context);
        var cacheKey = OSF.OUtil.formatString(OmexDataProvider.killedAppsWS.cacheKey, cacheKeyPrefix);
        var value = this._cacheManager.GetCacheItem(cacheKey, false);
        return (value != null);
    };
    OmexDataProvider.prototype.GetKilledApps = function (context, params, callback) {
        OSF.OUtil.validateParamObject(params, {
            "clientName": { type: String, mayBeNull: true },
            "clientVersion": { type: String, mayBeNull: true }
        }, callback);
        params.clearKilledApps = params.clearKilledApps || false;
        var response = { "context": context, "statusCode": OSF.ProxyCallStatusCode.Succeeded, "value": null, "cached": false };
        var cacheKey = OSF.OUtil.formatString(OmexDataProvider.killedAppsWS.cacheKey, this.GetCacheKeyPrefix(context));
        if (params.clearKilledApps) {
            this._cacheManager.RemoveCacheItem(cacheKey);
        }
        else {
            var value = this._cacheManager.GetCacheItem(cacheKey, true, params.errors);
            if (value) {
                response.value = value;
                response.cached = true;
                callback(response);
                return;
            }
        }
        context.manifestManager._invokeProxyMethodAsync(context, "OMEX_getKilledAppsAsync", callback, params);
    };
    OmexDataProvider.prototype.SetKilledAppsCache = function (context, killedAppsInfo) {
        OSF.OUtil.validateParamObject(killedAppsInfo, {
            "refreshRate": { type: String, mayBeNull: false }
        }, null);
        var cacheKey = OSF.OUtil.formatString(OmexDataProvider.killedAppsWS.cacheKey, this.GetCacheKeyPrefix(context));
        this._cacheManager.SetCacheItem(cacheKey, killedAppsInfo, killedAppsInfo.refreshRate / OmexDataProvider.hourToDayConversionFactor);
    };
    OmexDataProvider.prototype.AppStateCached = function (context, params) {
        OSF.OUtil.validateParamObject(params, {
            "assetID": { type: String, mayBeNull: false },
            "contentMarket": { type: String, mayBeNull: false }
        }, null);
        var cacheKey = OSF.OUtil.formatString(OmexDataProvider.appStateWS.cacheKey, this.GetCacheKeyPrefix(context), params.contentMarket, params.assetID);
        var value = this._cacheManager.GetCacheItem(cacheKey, false);
        return (value != null);
    };
    OmexDataProvider.prototype.GetAppState = function (context, params, callback) {
        OSF.OUtil.validateParamObject(params, {
            "assetID": { type: String, mayBeNull: false },
            "contentMarket": { type: String, mayBeNull: false },
            "clientName": { type: String, mayBeNull: true },
            "clientVersion": { type: String, mayBeNull: true }
        }, callback);
        params.clearAppState = params.clearAppState || false;
        var response = { "context": context, "statusCode": OSF.ProxyCallStatusCode.Succeeded, "value": null, "cached": false };
        var cacheKey = OSF.OUtil.formatString(OmexDataProvider.appStateWS.cacheKey, this.GetCacheKeyPrefix(context), params.contentMarket, params.assetID);
        if (params.clearAppState) {
            this._cacheManager.RemoveCacheItem(cacheKey);
        }
        else {
            var value = this._cacheManager.GetCacheItem(cacheKey, true, params.errors);
            if (value) {
                response.value = value;
                response.cached = true;
                callback(response);
                return;
            }
        }
        context.manifestManager._invokeProxyMethodAsync(context, "OMEX_getAppStateAsync", callback, params);
    };
    OmexDataProvider.prototype.SetAppStateCache = function (context, appState) {
        OSF.OUtil.validateParamObject(appState, {
            "refreshRate": { type: String, mayBeNull: false }
        }, null);
        var reference = context.referenceInUse;
        var cacheKey = OSF.OUtil.formatString(OmexDataProvider.appStateWS.cacheKey, this.GetCacheKeyPrefix(context), reference.storeLocator, reference.id);
        this._cacheManager.SetCacheItem(cacheKey, appState, appState.refreshRate / OmexDataProvider.hourToDayConversionFactor);
    };
    OmexDataProvider.prototype.ManifestAndETokenCached = function (context, params) {
        OSF.OUtil.validateParamObject(params, {
            "assetID": { type: String, mayBeNull: false },
            "contentMarket": { type: String, mayBeNull: false }
        }, null);
        var cacheKey;
        if (context.anonymous == null || context.anonymous) {
            cacheKey = OSF.OUtil.formatString(OmexDataProvider.anonymousAppInstallInfoWS.cacheKey, params.assetID, params.contentMarket);
        }
        else {
            cacheKey = OSF.OUtil.formatString(OmexDataProvider.authenticatedAppInstallInfoWS.cacheKey, this._cid, params.assetID, params.userContentMarket, params.assetContentMarket);
        }
        var value = this._cacheManager.GetCacheItem(cacheKey, false);
        return (value != null);
    };
    OmexDataProvider.prototype.GetManifestAndEToken = function (context, params, callback) {
        OSF.OUtil.validateParamObject(params, {
            "assetID": { type: String, mayBeNull: false },
            "applicationName": { type: String, mayBeNull: false },
            "contentMarket": { type: String, mayBeNull: true },
            "userContentMarket": { type: String, mayBeNull: true },
            "assetContentMarket": { type: String, mayBeNull: true },
            "clientName": { type: String, mayBeNull: true },
            "clientVersion": { type: String, mayBeNull: true }
        }, callback);
        params.clearToken = params.clearToken || false;
        params.clearManifest = params.clearManifest || false;
        var response = { "context": context, "statusCode": OSF.ProxyCallStatusCode.Succeeded, "value": null, "cached": false };
        var cacheKey;
        if (context.anonymous == null || context.anonymous) {
            cacheKey = OSF.OUtil.formatString(OmexDataProvider.anonymousAppInstallInfoWS.cacheKey, params.assetID, params.contentMarket);
        }
        else {
            cacheKey = OSF.OUtil.formatString(OmexDataProvider.authenticatedAppInstallInfoWS.cacheKey, this._cid, params.assetID, params.userContentMarket, params.assetContentMarket);
        }
        if (params.clearManifest) {
            this._cacheManager.RemoveCacheItem(cacheKey);
        }
        else {
            var value = this._cacheManager.GetCacheItem(cacheKey, true, params.errors);
            if (value) {
                if (params.clearToken || (value.tokenExpirationDate && new Date(value.tokenExpirationDate) <= new Date())) {
                    delete value.etoken;
                    delete value.tokenExpirationDate;
                    delete value.entitlementType;
                }
                if ((context.anonymous == null || context.anonymous || value.etoken) && value.manifest) {
                    response.value = value;
                    response.cached = true;
                    callback(response);
                    return;
                }
            }
            if (!context.anonymous) {
                cacheKey = OSF.OUtil.formatString(OmexDataProvider.anonymousAppInstallInfoWS.cacheKey, params.assetID, params.userContentMarket);
                var unauthenticated = this._cacheManager.GetCacheItem(cacheKey, true, params.errors);
                if (unauthenticated && unauthenticated.manifest) {
                    var onGetOmexETokenCompleted = function (asyncResult) {
                        var context = asyncResult.context;
                        if (asyncResult.statusCode === OSF.ProxyCallStatusCode.Succeeded && asyncResult.value) {
                            var manifestAndEToken = asyncResult.value;
                            var clientAppStatus = parseInt(manifestAndEToken.status);
                            if (clientAppStatus === OSF.OmexClientAppStatus.OK) {
                                manifestAndEToken.manifest = unauthenticated.manifest;
                            }
                            response.context = context;
                            response.value = manifestAndEToken;
                            callback(response);
                        }
                        else {
                            callback({ "context": context, "statusCode": OSF.ProxyCallStatusCode.Failed, "value": null, "cached": false });
                        }
                    };
                    if (!value || !value.etoken) {
                        params.clientAppInfoReturnType = OSF.ClientAppInfoReturnType.etokenOnly;
                        context.manifestManager._invokeProxyMethodAsync(context, "OMEX_getManifestAndETokenAsync", onGetOmexETokenCompleted, params);
                        return;
                    }
                }
            }
        }
        context.manifestManager._invokeProxyMethodAsync(context, "OMEX_getManifestAndETokenAsync", callback, params);
    };
    OmexDataProvider.prototype.SetManifestAndETokenCache = function (context, manifestAndEToken) {
        var reference = context.referenceInUse;
        var cacheKey;
        if (context.anonymous == null || context.anonymous) {
            cacheKey = OSF.OUtil.formatString(OmexDataProvider.anonymousAppInstallInfoWS.cacheKey, reference.id, reference.storeLocator);
        }
        else {
            cacheKey = OSF.OUtil.formatString(OmexDataProvider.authenticatedAppInstallInfoWS.cacheKey, this._cid, reference.id, reference.storeLocator, context.osfControl._omexEntitlement.contentMarket);
        }
        this._cacheManager.SetCacheItem(cacheKey, manifestAndEToken, OmexDataProvider.manifestRefreshRate);
    };
    OmexDataProvider.prototype.RemoveManifestAndEToken = function (assetId) {
        var REGEX_ANY_CHARACTERS = ".*";
        var cacheKeyPatterns = [];
        cacheKeyPatterns.push(new RegExp(OSF.OUtil.formatString(OmexDataProvider.authenticatedAppInstallInfoWS.cacheKey, REGEX_ANY_CHARACTERS, assetId, REGEX_ANY_CHARACTERS, REGEX_ANY_CHARACTERS), "i"));
        this._cacheManager.RemoveMatches(cacheKeyPatterns);
    };
    OmexDataProvider.prototype.AppDetailsCached = function (context, params) {
        var ids = [];
        for (var cm in params) {
            ids = params[cm].split(",");
            break;
        }
        var notCachedIds = [];
        for (var i = 0; i < ids.length; ++i) {
            var cacheKey = OSF.OUtil.formatString(OmexDataProvider.appDetailKey, ids[i]);
            var value = this._cacheManager.GetCacheItem(cacheKey, false);
            if (!value) {
                return false;
            }
        }
        return true;
    };
    OmexDataProvider.prototype.GetAppDetails = function (context, params, callback) {
        OSF.OUtil.validateParamObject(params, {
            "assetID": { type: String, mayBeNull: false },
            "contentMarket": { type: String, mayBeNull: false },
            "clientName": { type: String, mayBeNull: true },
            "clientVersion": { type: String, mayBeNull: true }
        }, callback);
        params.clearCache = params.clearCache || false;
        params.appDetails = [];
        if (params.clearCache) {
            this._cacheManager.RemoveAll(OmexDataProvider.ungatedCacheKeyPrefix);
        }
        else {
            var ids = params.assetID.split(",");
            var notCachedIds = [];
            for (var i = 0; i < ids.length; ++i) {
                var cacheKey = OSF.OUtil.formatString(OmexDataProvider.appDetailKey, ids[i]);
                var value = this._cacheManager.GetCacheItem(cacheKey, true, params.errors);
                if (value) {
                    params.appDetails.push(value);
                }
                else {
                    notCachedIds.push(ids[i]);
                }
            }
            if (notCachedIds.length === 0) {
                var response = { "context": context, "statusCode": OSF.ProxyCallStatusCode.Succeeded, "value": null, "cached": false };
                response.value = params.appDetails;
                response.cached = true;
                callback(response);
                return;
            }
            params.assetID = notCachedIds.join(",");
        }
        context.appDetails = params.appDetails;
        context.manifestManager._invokeProxyMethodAsync(context, "OMEX_getAppDetailsAsync", callback, params);
    };
    OmexDataProvider.prototype.SetAppDetailsCache = function (context, appDetails) {
        var appDetail;
        var cacheKey;
        for (var i = 0; i < appDetails.length; ++i) {
            appDetail = appDetails[i];
            cacheKey = OSF.OUtil.formatString(OmexDataProvider.appDetailKey, appDetail.assetId);
            this._cacheManager.SetCacheItem(cacheKey, appDetail);
        }
    };
    OmexDataProvider.anonymousCacheKeyPrefix = "__OSF_ANONYMOUS_OMEX.";
    OmexDataProvider.gatedCacheKeyPrefix = "__OSF_GATED_OMEX.";
    OmexDataProvider.ungatedCacheKeyPrefix = "__OSF_OMEX.";
    OmexDataProvider.manifestRefreshRate = 5 * 365;
    OmexDataProvider.hourToDayConversionFactor = 24;
    OmexDataProvider.anonymousAppInstallInfoWS = {
        url: "/appinstall/unauthenticated?cmu={0}&assetid={1}&ret=0",
        cacheKey: OmexDataProvider.anonymousCacheKeyPrefix + "appInstallInfo.{0}.{1}"
    };
    OmexDataProvider.authenticatedAppInstallInfoWS = {
        url: "/appinstall/authenticated?cmu={0}&cmf={1}&assetid={2}&ret={3}&rt=xml",
        cacheKey: OmexDataProvider.gatedCacheKeyPrefix + "appinstall_authenticated.{0}.{1}.{2}.{3}"
    };
    OmexDataProvider.killedAppsWS = {
        url: "/appinfo/query?rt=xml",
        cacheKey: "{0}killedApps"
    };
    OmexDataProvider.appStateWS = {
        url: "/appstate/query?ma={0}:{1}",
        cacheKey: "{0}appState.{1}.{2}"
    };
    OmexDataProvider.appDetailKey = "__OSF_OMEX.appDetails.{0}";
    return OmexDataProvider;
})();
var OSFLog;
(function (OSFLog) {
    var BaseUsageData = (function () {
        function BaseUsageData(table) {
            this._table = table;
            this._fields = {};
        }
        Object.defineProperty(BaseUsageData.prototype, "Fields", {
            get: function () {
                return this._fields;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(BaseUsageData.prototype, "Table", {
            get: function () {
                return this._table;
            },
            enumerable: true,
            configurable: true
        });
        BaseUsageData.prototype.SerializeFields = function () {
        };
        BaseUsageData.prototype.SetSerializedField = function (key, value) {
            if (typeof (value) !== "undefined" && value !== null) {
                this._serializedFields[key] = value.toString();
            }
        };
        BaseUsageData.prototype.SerializeRow = function () {
            this._serializedFields = {};
            this.SetSerializedField("Table", this._table);
            this.SerializeFields();
            return JSON.stringify(this._serializedFields);
        };
        return BaseUsageData;
    })();
    OSFLog.BaseUsageData = BaseUsageData;
    var AppLoadTimeUsageData = (function (_super) {
        __extends(AppLoadTimeUsageData, _super);
        function AppLoadTimeUsageData() {
            _super.call(this, "AppLoadTime");
        }
        Object.defineProperty(AppLoadTimeUsageData.prototype, "CorrelationId", {
            get: function () { return this.Fields["CorrelationId"]; },
            set: function (value) { this.Fields["CorrelationId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppLoadTimeUsageData.prototype, "AppInfo", {
            get: function () { return this.Fields["AppInfo"]; },
            set: function (value) { this.Fields["AppInfo"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppLoadTimeUsageData.prototype, "ActivationInfo", {
            get: function () { return this.Fields["ActivationInfo"]; },
            set: function (value) { this.Fields["ActivationInfo"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppLoadTimeUsageData.prototype, "InstanceId", {
            get: function () { return this.Fields["InstanceId"]; },
            set: function (value) { this.Fields["InstanceId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppLoadTimeUsageData.prototype, "AssetId", {
            get: function () { return this.Fields["AssetId"]; },
            set: function (value) { this.Fields["AssetId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppLoadTimeUsageData.prototype, "Stage1Time", {
            get: function () { return this.Fields["Stage1Time"]; },
            set: function (value) { this.Fields["Stage1Time"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppLoadTimeUsageData.prototype, "Stage2Time", {
            get: function () { return this.Fields["Stage2Time"]; },
            set: function (value) { this.Fields["Stage2Time"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppLoadTimeUsageData.prototype, "Stage3Time", {
            get: function () { return this.Fields["Stage3Time"]; },
            set: function (value) { this.Fields["Stage3Time"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppLoadTimeUsageData.prototype, "Stage4Time", {
            get: function () { return this.Fields["Stage4Time"]; },
            set: function (value) { this.Fields["Stage4Time"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppLoadTimeUsageData.prototype, "Stage5Time", {
            get: function () { return this.Fields["Stage5Time"]; },
            set: function (value) { this.Fields["Stage5Time"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppLoadTimeUsageData.prototype, "Stage6Time", {
            get: function () { return this.Fields["Stage6Time"]; },
            set: function (value) { this.Fields["Stage6Time"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppLoadTimeUsageData.prototype, "Stage7Time", {
            get: function () { return this.Fields["Stage7Time"]; },
            set: function (value) { this.Fields["Stage7Time"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppLoadTimeUsageData.prototype, "Stage8Time", {
            get: function () { return this.Fields["Stage8Time"]; },
            set: function (value) { this.Fields["Stage8Time"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppLoadTimeUsageData.prototype, "Stage9Time", {
            get: function () { return this.Fields["Stage9Time"]; },
            set: function (value) { this.Fields["Stage9Time"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppLoadTimeUsageData.prototype, "Stage10Time", {
            get: function () { return this.Fields["Stage10Time"]; },
            set: function (value) { this.Fields["Stage10Time"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppLoadTimeUsageData.prototype, "Stage11Time", {
            get: function () { return this.Fields["Stage11Time"]; },
            set: function (value) { this.Fields["Stage11Time"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppLoadTimeUsageData.prototype, "ErrorResult", {
            get: function () { return this.Fields["ErrorResult"]; },
            set: function (value) { this.Fields["ErrorResult"] = value; },
            enumerable: true,
            configurable: true
        });
        AppLoadTimeUsageData.prototype.SerializeFields = function () {
            this.SetSerializedField("CorrelationId", this.CorrelationId);
            this.SetSerializedField("AppInfo", this.AppInfo);
            this.SetSerializedField("ActivationInfo", this.ActivationInfo);
            this.SetSerializedField("InstanceId", this.InstanceId);
            this.SetSerializedField("AssetId", this.AssetId);
            this.SetSerializedField("Stage1Time", this.Stage1Time);
            this.SetSerializedField("Stage2Time", this.Stage2Time);
            this.SetSerializedField("Stage3Time", this.Stage3Time);
            this.SetSerializedField("Stage4Time", this.Stage4Time);
            this.SetSerializedField("Stage5Time", this.Stage5Time);
            this.SetSerializedField("Stage6Time", this.Stage6Time);
            this.SetSerializedField("Stage7Time", this.Stage7Time);
            this.SetSerializedField("Stage8Time", this.Stage8Time);
            this.SetSerializedField("Stage9Time", this.Stage9Time);
            this.SetSerializedField("Stage10Time", this.Stage10Time);
            this.SetSerializedField("Stage11Time", this.Stage11Time);
            this.SetSerializedField("ErrorResult", this.ErrorResult);
        };
        return AppLoadTimeUsageData;
    })(BaseUsageData);
    OSFLog.AppLoadTimeUsageData = AppLoadTimeUsageData;
    var AppNotificationUsageData = (function (_super) {
        __extends(AppNotificationUsageData, _super);
        function AppNotificationUsageData() {
            _super.call(this, "AppNotification");
        }
        Object.defineProperty(AppNotificationUsageData.prototype, "CorrelationId", {
            get: function () { return this.Fields["CorrelationId"]; },
            set: function (value) { this.Fields["CorrelationId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppNotificationUsageData.prototype, "ErrorResult", {
            get: function () { return this.Fields["ErrorResult"]; },
            set: function (value) { this.Fields["ErrorResult"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppNotificationUsageData.prototype, "NotificationClickInfo", {
            get: function () { return this.Fields["NotificationClickInfo"]; },
            set: function (value) { this.Fields["NotificationClickInfo"] = value; },
            enumerable: true,
            configurable: true
        });
        AppNotificationUsageData.prototype.SerializeFields = function () {
            this.SetSerializedField("CorrelationId", this.CorrelationId);
            this.SetSerializedField("ErrorResult", this.ErrorResult);
            this.SetSerializedField("NotificationClickInfo", this.NotificationClickInfo);
        };
        return AppNotificationUsageData;
    })(BaseUsageData);
    OSFLog.AppNotificationUsageData = AppNotificationUsageData;
    var AppManagementMenuUsageData = (function (_super) {
        __extends(AppManagementMenuUsageData, _super);
        function AppManagementMenuUsageData() {
            _super.call(this, "AppManagementMenu");
        }
        Object.defineProperty(AppManagementMenuUsageData.prototype, "AssetId", {
            get: function () { return this.Fields["AssetId"]; },
            set: function (value) { this.Fields["AssetId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppManagementMenuUsageData.prototype, "OperationMetadata", {
            get: function () { return this.Fields["OperationMetadata"]; },
            set: function (value) { this.Fields["OperationMetadata"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppManagementMenuUsageData.prototype, "ErrorResult", {
            get: function () { return this.Fields["ErrorResult"]; },
            set: function (value) { this.Fields["ErrorResult"] = value; },
            enumerable: true,
            configurable: true
        });
        AppManagementMenuUsageData.prototype.SerializeFields = function () {
            this.SetSerializedField("AssetId", this.AssetId);
            this.SetSerializedField("OperationMetadata", this.OperationMetadata);
            this.SetSerializedField("ErrorResult", this.ErrorResult);
        };
        return AppManagementMenuUsageData;
    })(BaseUsageData);
    OSFLog.AppManagementMenuUsageData = AppManagementMenuUsageData;
    var InsertionDialogSessionUsageData = (function (_super) {
        __extends(InsertionDialogSessionUsageData, _super);
        function InsertionDialogSessionUsageData() {
            _super.call(this, "InsertionDialogSession");
        }
        Object.defineProperty(InsertionDialogSessionUsageData.prototype, "AssetId", {
            get: function () { return this.Fields["AssetId"]; },
            set: function (value) { this.Fields["AssetId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(InsertionDialogSessionUsageData.prototype, "TotalSessionTime", {
            get: function () { return this.Fields["TotalSessionTime"]; },
            set: function (value) { this.Fields["TotalSessionTime"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(InsertionDialogSessionUsageData.prototype, "TrustPageSessionTime", {
            get: function () { return this.Fields["TrustPageSessionTime"]; },
            set: function (value) { this.Fields["TrustPageSessionTime"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(InsertionDialogSessionUsageData.prototype, "DialogState", {
            get: function () { return this.Fields["DialogState"]; },
            set: function (value) { this.Fields["DialogState"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(InsertionDialogSessionUsageData.prototype, "LastActiveTab", {
            get: function () { return this.Fields["LastActiveTab"]; },
            set: function (value) { this.Fields["LastActiveTab"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(InsertionDialogSessionUsageData.prototype, "LastActiveTabCount", {
            get: function () { return this.Fields["LastActiveTabCount"]; },
            set: function (value) { this.Fields["LastActiveTabCount"] = value; },
            enumerable: true,
            configurable: true
        });
        InsertionDialogSessionUsageData.prototype.SerializeFields = function () {
            this.SetSerializedField("AssetId", this.AssetId);
            this.SetSerializedField("TotalSessionTime", this.TotalSessionTime);
            this.SetSerializedField("TrustPageSessionTime", this.TrustPageSessionTime);
            this.SetSerializedField("DialogState", this.DialogState);
            this.SetSerializedField("LastActiveTab", this.LastActiveTab);
            this.SetSerializedField("LastActiveTabCount", this.LastActiveTabCount);
        };
        return InsertionDialogSessionUsageData;
    })(BaseUsageData);
    OSFLog.InsertionDialogSessionUsageData = InsertionDialogSessionUsageData;
    var UploadFileDevCatelogUsageData = (function (_super) {
        __extends(UploadFileDevCatelogUsageData, _super);
        function UploadFileDevCatelogUsageData() {
            _super.call(this, "UploadFileDevCatelog");
        }
        Object.defineProperty(UploadFileDevCatelogUsageData.prototype, "CorrelationId", {
            get: function () { return this.Fields["CorrelationId"]; },
            set: function (value) { this.Fields["CorrelationId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(UploadFileDevCatelogUsageData.prototype, "OperationMetadata", {
            get: function () { return this.Fields["OperationMetadata"]; },
            set: function (value) { this.Fields["OperationMetadata"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(UploadFileDevCatelogUsageData.prototype, "ErrorResult", {
            get: function () { return this.Fields["ErrorResult"]; },
            set: function (value) { this.Fields["ErrorResult"] = value; },
            enumerable: true,
            configurable: true
        });
        UploadFileDevCatelogUsageData.prototype.SerializeFields = function () {
            this.SetSerializedField("CorrelationId", this.CorrelationId);
            this.SetSerializedField("OperationMetadata", this.OperationMetadata);
            this.SetSerializedField("ErrorResult", this.ErrorResult);
        };
        return UploadFileDevCatelogUsageData;
    })(BaseUsageData);
    OSFLog.UploadFileDevCatelogUsageData = UploadFileDevCatelogUsageData;
    var UploadFileDevCatalogUsageUsageData = (function (_super) {
        __extends(UploadFileDevCatalogUsageUsageData, _super);
        function UploadFileDevCatalogUsageUsageData() {
            _super.call(this, "UploadFileDevCatalogUsage");
        }
        Object.defineProperty(UploadFileDevCatalogUsageUsageData.prototype, "CorrelationId", {
            get: function () { return this.Fields["CorrelationId"]; },
            set: function (value) { this.Fields["CorrelationId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(UploadFileDevCatalogUsageUsageData.prototype, "StoreType", {
            get: function () { return this.Fields["StoreType"]; },
            set: function (value) { this.Fields["StoreType"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(UploadFileDevCatalogUsageUsageData.prototype, "AppId", {
            get: function () { return this.Fields["AppId"]; },
            set: function (value) { this.Fields["AppId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(UploadFileDevCatalogUsageUsageData.prototype, "AppVersion", {
            get: function () { return this.Fields["AppVersion"]; },
            set: function (value) { this.Fields["AppVersion"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(UploadFileDevCatalogUsageUsageData.prototype, "AppTargetType", {
            get: function () { return this.Fields["AppTargetType"]; },
            set: function (value) { this.Fields["AppTargetType"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(UploadFileDevCatalogUsageUsageData.prototype, "IsAppCommand", {
            get: function () { return this.Fields["IsAppCommand"]; },
            set: function (value) { this.Fields["IsAppCommand"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(UploadFileDevCatalogUsageUsageData.prototype, "AppSizeWidth", {
            get: function () { return this.Fields["AppSizeWidth"]; },
            set: function (value) { this.Fields["AppSizeWidth"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(UploadFileDevCatalogUsageUsageData.prototype, "AppSizeHeight", {
            get: function () { return this.Fields["AppSizeHeight"]; },
            set: function (value) { this.Fields["AppSizeHeight"] = value; },
            enumerable: true,
            configurable: true
        });
        UploadFileDevCatalogUsageUsageData.prototype.SerializeFields = function () {
            this.SetSerializedField("CorrelationId", this.CorrelationId);
            this.SetSerializedField("StoreType", this.StoreType);
            this.SetSerializedField("AppId", this.AppId);
            this.SetSerializedField("AppVersion", this.AppVersion);
            this.SetSerializedField("AppTargetType", this.AppTargetType);
            this.SetSerializedField("IsAppCommand", this.IsAppCommand);
            this.SetSerializedField("AppSizeWidth", this.AppSizeWidth);
            this.SetSerializedField("AppSizeHeight", this.AppSizeHeight);
        };
        return UploadFileDevCatalogUsageUsageData;
    })(BaseUsageData);
    OSFLog.UploadFileDevCatalogUsageUsageData = UploadFileDevCatalogUsageUsageData;
    var OneDriveCatalogGetAccessTokenUsageData = (function (_super) {
        __extends(OneDriveCatalogGetAccessTokenUsageData, _super);
        function OneDriveCatalogGetAccessTokenUsageData() {
            _super.call(this, "OneDriveCatalogGetAccessToken");
        }
        Object.defineProperty(OneDriveCatalogGetAccessTokenUsageData.prototype, "Result", {
            get: function () { return this.Fields["Result"]; },
            set: function (value) { this.Fields["Result"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(OneDriveCatalogGetAccessTokenUsageData.prototype, "ResponseTime", {
            get: function () { return this.Fields["ResponseTime"]; },
            set: function (value) { this.Fields["ResponseTime"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(OneDriveCatalogGetAccessTokenUsageData.prototype, "Error", {
            get: function () { return this.Fields["Error"]; },
            set: function (value) { this.Fields["Error"] = value; },
            enumerable: true,
            configurable: true
        });
        OneDriveCatalogGetAccessTokenUsageData.prototype.SerializeFields = function () {
            this.SetSerializedField("Result", this.Result);
            this.SetSerializedField("ResponseTime", this.ResponseTime);
            this.SetSerializedField("Error", this.Error);
        };
        return OneDriveCatalogGetAccessTokenUsageData;
    })(BaseUsageData);
    OSFLog.OneDriveCatalogGetAccessTokenUsageData = OneDriveCatalogGetAccessTokenUsageData;
    var OneDriveCatalogGetManifestsUsageData = (function (_super) {
        __extends(OneDriveCatalogGetManifestsUsageData, _super);
        function OneDriveCatalogGetManifestsUsageData() {
            _super.call(this, "OneDriveCatalogGetManifests");
        }
        Object.defineProperty(OneDriveCatalogGetManifestsUsageData.prototype, "Result", {
            get: function () { return this.Fields["Result"]; },
            set: function (value) { this.Fields["Result"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(OneDriveCatalogGetManifestsUsageData.prototype, "ResponseTime", {
            get: function () { return this.Fields["ResponseTime"]; },
            set: function (value) { this.Fields["ResponseTime"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(OneDriveCatalogGetManifestsUsageData.prototype, "Error", {
            get: function () { return this.Fields["Error"]; },
            set: function (value) { this.Fields["Error"] = value; },
            enumerable: true,
            configurable: true
        });
        OneDriveCatalogGetManifestsUsageData.prototype.SerializeFields = function () {
            this.SetSerializedField("Result", this.Result);
            this.SetSerializedField("ResponseTime", this.ResponseTime);
            this.SetSerializedField("Error", this.Error);
        };
        return OneDriveCatalogGetManifestsUsageData;
    })(BaseUsageData);
    OSFLog.OneDriveCatalogGetManifestsUsageData = OneDriveCatalogGetManifestsUsageData;
    var OneDriveCatalogGetManifestUsageData = (function (_super) {
        __extends(OneDriveCatalogGetManifestUsageData, _super);
        function OneDriveCatalogGetManifestUsageData() {
            _super.call(this, "OneDriveCatalogGetManifest");
        }
        Object.defineProperty(OneDriveCatalogGetManifestUsageData.prototype, "id", {
            get: function () { return this.Fields["id"]; },
            set: function (value) { this.Fields["id"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(OneDriveCatalogGetManifestUsageData.prototype, "Result", {
            get: function () { return this.Fields["Result"]; },
            set: function (value) { this.Fields["Result"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(OneDriveCatalogGetManifestUsageData.prototype, "ResponseTime", {
            get: function () { return this.Fields["ResponseTime"]; },
            set: function (value) { this.Fields["ResponseTime"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(OneDriveCatalogGetManifestUsageData.prototype, "Error", {
            get: function () { return this.Fields["Error"]; },
            set: function (value) { this.Fields["Error"] = value; },
            enumerable: true,
            configurable: true
        });
        OneDriveCatalogGetManifestUsageData.prototype.SerializeFields = function () {
            this.SetSerializedField("id", this.id);
            this.SetSerializedField("Result", this.Result);
            this.SetSerializedField("ResponseTime", this.ResponseTime);
            this.SetSerializedField("Error", this.Error);
        };
        return OneDriveCatalogGetManifestUsageData;
    })(BaseUsageData);
    OSFLog.OneDriveCatalogGetManifestUsageData = OneDriveCatalogGetManifestUsageData;
    var OneDriveCatalogInsertManifestUsageData = (function (_super) {
        __extends(OneDriveCatalogInsertManifestUsageData, _super);
        function OneDriveCatalogInsertManifestUsageData() {
            _super.call(this, "OneDriveCatalogInsertManifest");
        }
        Object.defineProperty(OneDriveCatalogInsertManifestUsageData.prototype, "Result", {
            get: function () { return this.Fields["Result"]; },
            set: function (value) { this.Fields["Result"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(OneDriveCatalogInsertManifestUsageData.prototype, "ResponseTime", {
            get: function () { return this.Fields["ResponseTime"]; },
            set: function (value) { this.Fields["ResponseTime"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(OneDriveCatalogInsertManifestUsageData.prototype, "Error", {
            get: function () { return this.Fields["Error"]; },
            set: function (value) { this.Fields["Error"] = value; },
            enumerable: true,
            configurable: true
        });
        OneDriveCatalogInsertManifestUsageData.prototype.SerializeFields = function () {
            this.SetSerializedField("Result", this.Result);
            this.SetSerializedField("ResponseTime", this.ResponseTime);
            this.SetSerializedField("Error", this.Error);
        };
        return OneDriveCatalogInsertManifestUsageData;
    })(BaseUsageData);
    OSFLog.OneDriveCatalogInsertManifestUsageData = OneDriveCatalogInsertManifestUsageData;
})(OSFLog || (OSFLog = {}));
var Telemetry;
(function (Telemetry) {
    "use strict";
    (function (ULSTraceLevel) {
        ULSTraceLevel[ULSTraceLevel["unexpected"] = 10] = "unexpected";
        ULSTraceLevel[ULSTraceLevel["warning"] = 15] = "warning";
        ULSTraceLevel[ULSTraceLevel["info"] = 50] = "info";
        ULSTraceLevel[ULSTraceLevel["verbose"] = 100] = "verbose";
        ULSTraceLevel[ULSTraceLevel["verboseEx"] = 200] = "verboseEx";
    })(Telemetry.ULSTraceLevel || (Telemetry.ULSTraceLevel = {}));
    var ULSTraceLevel = Telemetry.ULSTraceLevel;
    (function (ULSCat) {
        ULSCat[ULSCat["msoulscat_Osf_Latency"] = 1401] = "msoulscat_Osf_Latency";
        ULSCat[ULSCat["msoulscat_Osf_Notification"] = 1402] = "msoulscat_Osf_Notification";
        ULSCat[ULSCat["msoulscat_Osf_Runtime"] = 1403] = "msoulscat_Osf_Runtime";
        ULSCat[ULSCat["msoulscat_Osf_AppManagementMenu"] = 1404] = "msoulscat_Osf_AppManagementMenu";
        ULSCat[ULSCat["msoulscat_Osf_InsertionDialogSession"] = 1405] = "msoulscat_Osf_InsertionDialogSession";
        ULSCat[ULSCat["msoulscat_Osf_UploadFileDevCatelog"] = 1406] = "msoulscat_Osf_UploadFileDevCatelog";
        ULSCat[ULSCat["msoulscat_Osf_UploadFileDevCatalogUsage"] = 1411] = "msoulscat_Osf_UploadFileDevCatalogUsage";
    })(Telemetry.ULSCat || (Telemetry.ULSCat = {}));
    var ULSCat = Telemetry.ULSCat;
    var AppManagementMenuFlags;
    (function (AppManagementMenuFlags) {
        AppManagementMenuFlags[AppManagementMenuFlags["ConfirmationDialogCancel"] = 256] = "ConfirmationDialogCancel";
        AppManagementMenuFlags[AppManagementMenuFlags["InsertionDialogClosed"] = 512] = "InsertionDialogClosed";
        AppManagementMenuFlags[AppManagementMenuFlags["IsAnonymous"] = 1024] = "IsAnonymous";
    })(AppManagementMenuFlags || (AppManagementMenuFlags = {}));
    var InsertionDialogStateFlags;
    (function (InsertionDialogStateFlags) {
        InsertionDialogStateFlags[InsertionDialogStateFlags["Undefined"] = 0] = "Undefined";
        InsertionDialogStateFlags[InsertionDialogStateFlags["Inserted"] = 1] = "Inserted";
        InsertionDialogStateFlags[InsertionDialogStateFlags["Canceled"] = 2] = "Canceled";
        InsertionDialogStateFlags[InsertionDialogStateFlags["Closed"] = 3] = "Closed";
        InsertionDialogStateFlags[InsertionDialogStateFlags["TrustPageVisited"] = 8] = "TrustPageVisited";
    })(InsertionDialogStateFlags || (InsertionDialogStateFlags = {}));
    var LatencyStopwatch = (function () {
        function LatencyStopwatch() {
            this.timeValue = 0;
        }
        LatencyStopwatch.prototype.Start = function () {
            this.timeValue = -(new Date().getTime());
            this.finishedMeasurement = false;
        };
        LatencyStopwatch.prototype.Stop = function () {
            if (this.timeValue < 0) {
                this.timeValue += (new Date().getTime());
                this.finishedMeasurement = true;
            }
        };
        Object.defineProperty(LatencyStopwatch.prototype, "Finished", {
            get: function () {
                return this.finishedMeasurement;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(LatencyStopwatch.prototype, "ElapsedTime", {
            get: function () {
                var elapsedTime = this.timeValue;
                if (!this.Finished && elapsedTime < 0) {
                    elapsedTime = Math.abs(elapsedTime) - (new Date().getTime());
                }
                return elapsedTime;
            },
            enumerable: true,
            configurable: true
        });
        return LatencyStopwatch;
    })();
    Telemetry.LatencyStopwatch = LatencyStopwatch;
    var Context = (function () {
        function Context() {
        }
        return Context;
    })();
    Telemetry.Context = Context;
    var Logger = (function () {
        function Logger() {
        }
        Logger.SendULSTraceTag = function (category, level, data, tagId) {
            if (!Microsoft.Office.WebExtension.FULSSupported) {
                return;
            }
            Diag.UULS.trace(tagId, category, level, data);
        };
        return Logger;
    })();
    var NotificationLogger = (function () {
        function NotificationLogger() {
        }
        NotificationLogger.Instance = function () {
            if (!NotificationLogger.instance) {
                NotificationLogger.instance = new NotificationLogger();
            }
            return NotificationLogger.instance;
        };
        NotificationLogger.prototype.LogData = function (data) {
            Logger.SendULSTraceTag(NotificationLogger.category, NotificationLogger.level, data.SerializeRow(), 0x005c815f);
        };
        NotificationLogger.category = ULSCat.msoulscat_Osf_Notification;
        NotificationLogger.level = ULSTraceLevel.info;
        return NotificationLogger;
    })();
    Telemetry.NotificationLogger = NotificationLogger;
    var AppManagementMenuLogger = (function () {
        function AppManagementMenuLogger() {
        }
        AppManagementMenuLogger.Instance = function () {
            if (!AppManagementMenuLogger.instance) {
                AppManagementMenuLogger.instance = new AppManagementMenuLogger();
            }
            return AppManagementMenuLogger.instance;
        };
        AppManagementMenuLogger.prototype.LogData = function (data) {
            Logger.SendULSTraceTag(AppManagementMenuLogger.category, AppManagementMenuLogger.level, data.SerializeRow(), 0x0104b605);
        };
        AppManagementMenuLogger.category = ULSCat.msoulscat_Osf_AppManagementMenu;
        AppManagementMenuLogger.level = ULSTraceLevel.info;
        return AppManagementMenuLogger;
    })();
    Telemetry.AppManagementMenuLogger = AppManagementMenuLogger;
    var UploadFileDevCatelogLogger = (function () {
        function UploadFileDevCatelogLogger() {
        }
        UploadFileDevCatelogLogger.Instance = function () {
            if (!UploadFileDevCatelogLogger.instance) {
                UploadFileDevCatelogLogger.instance = new UploadFileDevCatelogLogger();
            }
            return UploadFileDevCatelogLogger.instance;
        };
        UploadFileDevCatelogLogger.prototype.LogData = function (data) {
            Logger.SendULSTraceTag(UploadFileDevCatelogLogger.category, UploadFileDevCatelogLogger.level, data.SerializeRow(), 0x0104b606);
        };
        UploadFileDevCatelogLogger.category = ULSCat.msoulscat_Osf_UploadFileDevCatelog;
        UploadFileDevCatelogLogger.level = ULSTraceLevel.info;
        return UploadFileDevCatelogLogger;
    })();
    Telemetry.UploadFileDevCatelogLogger = UploadFileDevCatelogLogger;
    var UploadFileDevCatalogUsageLogger = (function () {
        function UploadFileDevCatalogUsageLogger() {
        }
        UploadFileDevCatalogUsageLogger.Instance = function () {
            if (!UploadFileDevCatalogUsageLogger.instance) {
                UploadFileDevCatalogUsageLogger.instance = new UploadFileDevCatalogUsageLogger();
            }
            return UploadFileDevCatalogUsageLogger.instance;
        };
        UploadFileDevCatalogUsageLogger.prototype.LogData = function (data) {
            Logger.SendULSTraceTag(UploadFileDevCatalogUsageLogger.category, UploadFileDevCatalogUsageLogger.level, data.SerializeRow(), 0x0110229d);
        };
        UploadFileDevCatalogUsageLogger.category = ULSCat.msoulscat_Osf_UploadFileDevCatalogUsage;
        UploadFileDevCatalogUsageLogger.level = ULSTraceLevel.info;
        return UploadFileDevCatalogUsageLogger;
    })();
    Telemetry.UploadFileDevCatalogUsageLogger = UploadFileDevCatalogUsageLogger;
    var LatencyLogger = (function () {
        function LatencyLogger() {
        }
        LatencyLogger.Instance = function () {
            if (!LatencyLogger.instance) {
                LatencyLogger.instance = new LatencyLogger();
            }
            return LatencyLogger.instance;
        };
        LatencyLogger.prototype.LogData = function (data) {
            Logger.SendULSTraceTag(LatencyLogger.category, LatencyLogger.level, data.SerializeRow(), 0x00487317);
        };
        LatencyLogger.category = ULSCat.msoulscat_Osf_Latency;
        LatencyLogger.level = ULSTraceLevel.info;
        return LatencyLogger;
    })();
    Telemetry.LatencyLogger = LatencyLogger;
    var InsertionDialogSessionLogger = (function () {
        function InsertionDialogSessionLogger() {
        }
        InsertionDialogSessionLogger.Instance = function () {
            if (!InsertionDialogSessionLogger.instance) {
                InsertionDialogSessionLogger.instance = new InsertionDialogSessionLogger();
            }
            return InsertionDialogSessionLogger.instance;
        };
        InsertionDialogSessionLogger.prototype.LogData = function (data) {
            Logger.SendULSTraceTag(InsertionDialogSessionLogger.category, InsertionDialogSessionLogger.level, data.SerializeRow(), 0x0104b607);
        };
        InsertionDialogSessionLogger.category = ULSCat.msoulscat_Osf_InsertionDialogSession;
        InsertionDialogSessionLogger.level = ULSTraceLevel.info;
        return InsertionDialogSessionLogger;
    })();
    Telemetry.InsertionDialogSessionLogger = InsertionDialogSessionLogger;
    var AppNotificationHelper = (function () {
        function AppNotificationHelper() {
        }
        AppNotificationHelper.LogNotification = function (correlationId, errorResult, notificationClickInfo) {
            var notificationData = new OSFLog.AppNotificationUsageData();
            notificationData.CorrelationId = correlationId;
            notificationData.ErrorResult = errorResult;
            notificationData.NotificationClickInfo = notificationClickInfo;
            NotificationLogger.Instance().LogData(notificationData);
        };
        return AppNotificationHelper;
    })();
    Telemetry.AppNotificationHelper = AppNotificationHelper;
    var AppManagementMenuHelper = (function () {
        function AppManagementMenuHelper() {
        }
        AppManagementMenuHelper.LogAppManagementMenuAction = function (assetId, operationMetadata, untrustedCount, isDialogClosed, isAnonymous, hrStatus) {
            var appManagementMenuData = new OSFLog.AppManagementMenuUsageData();
            var assetIdNumber = assetId.toLowerCase().indexOf("wa") === 0 ? parseInt(assetId.substring(2), 10) : parseInt(assetId, 10);
            if (isDialogClosed) {
                operationMetadata |= AppManagementMenuFlags.InsertionDialogClosed;
            }
            if (isAnonymous) {
                operationMetadata |= AppManagementMenuFlags.IsAnonymous;
            }
            appManagementMenuData.AssetId = assetIdNumber;
            appManagementMenuData.OperationMetadata = operationMetadata;
            appManagementMenuData.ErrorResult = hrStatus;
            AppManagementMenuLogger.Instance().LogData(appManagementMenuData);
        };
        return AppManagementMenuHelper;
    })();
    Telemetry.AppManagementMenuHelper = AppManagementMenuHelper;
    var UploadFileDevCatelogHelper = (function () {
        function UploadFileDevCatelogHelper() {
        }
        UploadFileDevCatelogHelper.LogUploadFileDevCatelogAction = function (correlationId, operationMetadata, untrustedCount, isDialogClosed, isAnonymous, hrStatus) {
            var uploadFileDevCatelogData = new OSFLog.UploadFileDevCatelogUsageData();
            uploadFileDevCatelogData.CorrelationId = correlationId;
            uploadFileDevCatelogData.OperationMetadata = operationMetadata;
            uploadFileDevCatelogData.ErrorResult = hrStatus;
            UploadFileDevCatelogLogger.Instance().LogData(uploadFileDevCatelogData);
        };
        return UploadFileDevCatelogHelper;
    })();
    Telemetry.UploadFileDevCatelogHelper = UploadFileDevCatelogHelper;
    var UploadFileDevCatalogUsageHelper = (function () {
        function UploadFileDevCatalogUsageHelper() {
        }
        UploadFileDevCatalogUsageHelper.LogUploadFileDevCatalogUsageAction = function (correlationId, storeType, id, appVersion, appTargetType, isAppCommand, appSizeWidth, appSizeHeight) {
            var uploadFileDevCatalogUsageData = new OSFLog.UploadFileDevCatalogUsageUsageData();
            uploadFileDevCatalogUsageData.CorrelationId = correlationId;
            uploadFileDevCatalogUsageData.StoreType = storeType;
            uploadFileDevCatalogUsageData.AppId = id;
            uploadFileDevCatalogUsageData.AppVersion = appVersion;
            uploadFileDevCatalogUsageData.AppTargetType = appTargetType;
            uploadFileDevCatalogUsageData.IsAppCommand = isAppCommand;
            uploadFileDevCatalogUsageData.AppSizeWidth = appSizeWidth;
            uploadFileDevCatalogUsageData.AppSizeHeight = appSizeHeight;
            UploadFileDevCatalogUsageLogger.Instance().LogData(uploadFileDevCatalogUsageData);
        };
        return UploadFileDevCatalogUsageHelper;
    })();
    Telemetry.UploadFileDevCatalogUsageHelper = UploadFileDevCatalogUsageHelper;
    var AppLoadTimeHelper = (function () {
        function AppLoadTimeHelper() {
        }
        AppLoadTimeHelper.ActivationStart = function (context, appInfo, assetId, correlationId, instanceId) {
            AppLoadTimeHelper.activatingNumber++;
            context.LoadTime = new OSFLog.AppLoadTimeUsageData();
            context.Timers = {};
            context.LoadTime.CorrelationId = correlationId;
            context.LoadTime.AppInfo = appInfo;
            context.LoadTime.ActivationInfo = 0;
            context.LoadTime.InstanceId = instanceId;
            context.LoadTime.AssetId = assetId;
            context.LoadTime.Stage1Time = 0;
            context.Timers["Stage1Time"] = new LatencyStopwatch();
            context.LoadTime.Stage2Time = 0;
            context.Timers["Stage2Time"] = new LatencyStopwatch();
            context.LoadTime.Stage3Time = 0;
            context.LoadTime.Stage4Time = 0;
            context.Timers["Stage4Time"] = new LatencyStopwatch();
            context.LoadTime.Stage5Time = 0;
            context.Timers["Stage5Time"] = new LatencyStopwatch();
            context.LoadTime.Stage6Time = AppLoadTimeHelper.activatingNumber;
            context.LoadTime.Stage7Time = 0;
            context.Timers["Stage7Time"] = new LatencyStopwatch();
            context.LoadTime.Stage8Time = 0;
            context.Timers["Stage8Time"] = new LatencyStopwatch();
            context.LoadTime.Stage9Time = 0;
            context.Timers["Stage9Time"] = new LatencyStopwatch();
            context.LoadTime.Stage10Time = 0;
            context.Timers["Stage10Time"] = new LatencyStopwatch();
            context.LoadTime.Stage11Time = 0;
            context.Timers["Stage11Time"] = new LatencyStopwatch();
            context.LoadTime.ErrorResult = 0;
            AppLoadTimeHelper.StartStopwatch(context, "Stage1Time");
        };
        AppLoadTimeHelper.ActivationEnd = function (context) {
            AppLoadTimeHelper.ActivateEndInternal(context);
        };
        AppLoadTimeHelper.PageStart = function (context) {
            AppLoadTimeHelper.StartStopwatch(context, "Stage2Time");
        };
        AppLoadTimeHelper.PageLoaded = function (context) {
            AppLoadTimeHelper.StopStopwatch(context, "Stage2Time");
        };
        AppLoadTimeHelper.ServerCallStart = function (context) {
            AppLoadTimeHelper.StartStopwatch(context, "Stage4Time");
        };
        AppLoadTimeHelper.ServerCallEnd = function (context) {
            AppLoadTimeHelper.StopStopwatch(context, "Stage4Time");
        };
        AppLoadTimeHelper.AuthenticationStart = function (context) {
            AppLoadTimeHelper.StartStopwatch(context, "Stage5Time");
        };
        AppLoadTimeHelper.AuthenticationEnd = function (context) {
            AppLoadTimeHelper.StopStopwatch(context, "Stage5Time");
        };
        AppLoadTimeHelper.EntitlementCheckStart = function (context) {
            AppLoadTimeHelper.StartStopwatch(context, "Stage7Time");
        };
        AppLoadTimeHelper.EntitlementCheckEnd = function (context) {
            AppLoadTimeHelper.StopStopwatch(context, "Stage7Time");
        };
        AppLoadTimeHelper.KilledAppsCheckStart = function (context) {
            AppLoadTimeHelper.StartStopwatch(context, "Stage8Time");
        };
        AppLoadTimeHelper.KilledAppsCheckEnd = function (context) {
            AppLoadTimeHelper.StopStopwatch(context, "Stage8Time");
        };
        AppLoadTimeHelper.AppStateCheckStart = function (context) {
            AppLoadTimeHelper.StartStopwatch(context, "Stage9Time");
        };
        AppLoadTimeHelper.AppStateCheckEnd = function (context) {
            AppLoadTimeHelper.StopStopwatch(context, "Stage9Time");
        };
        AppLoadTimeHelper.ManifestRequestStart = function (context) {
            AppLoadTimeHelper.StartStopwatch(context, "Stage10Time");
        };
        AppLoadTimeHelper.ManifestRequestEnd = function (context) {
            AppLoadTimeHelper.StopStopwatch(context, "Stage10Time");
        };
        AppLoadTimeHelper.OfficeJSStartToLoad = function (context) {
            AppLoadTimeHelper.StartStopwatch(context, "Stage11Time");
        };
        AppLoadTimeHelper.OfficeJSLoaded = function (context) {
            AppLoadTimeHelper.StopStopwatch(context, "Stage11Time");
        };
        AppLoadTimeHelper.SetAnonymousFlag = function (context, anonymousFlag) {
            AppLoadTimeHelper.SetActivationInfoField(context, AppLoadTimeHelper.ConvertFlagToBit(anonymousFlag), 2, 0);
        };
        AppLoadTimeHelper.SetRetryCount = function (context, retryCount) {
            AppLoadTimeHelper.SetActivationInfoField(context, retryCount, 3, 2);
        };
        AppLoadTimeHelper.SetManifestTrustCachedFlag = function (context, manifestTrustCachedFlag) {
            AppLoadTimeHelper.SetActivationInfoField(context, AppLoadTimeHelper.ConvertFlagToBit(manifestTrustCachedFlag), 2, 5);
        };
        AppLoadTimeHelper.SetManifestDataCachedFlag = function (context, manifestDataCachedFlag) {
            AppLoadTimeHelper.SetActivationInfoField(context, AppLoadTimeHelper.ConvertFlagToBit(manifestDataCachedFlag), 2, 7);
        };
        AppLoadTimeHelper.SetOmexHasEntitlementFlag = function (context, omexHasEntitlementFlag) {
            AppLoadTimeHelper.SetActivationInfoField(context, AppLoadTimeHelper.ConvertFlagToBit(omexHasEntitlementFlag), 2, 9);
        };
        AppLoadTimeHelper.SetManifestDataInvalidFlag = function (context, manifestDataInvalidFlag) {
            AppLoadTimeHelper.SetActivationInfoField(context, AppLoadTimeHelper.ConvertFlagToBit(manifestDataInvalidFlag), 2, 11);
        };
        AppLoadTimeHelper.SetAppStateDataCachedFlag = function (context, appStateDataCachedFlag) {
            AppLoadTimeHelper.SetActivationInfoField(context, AppLoadTimeHelper.ConvertFlagToBit(appStateDataCachedFlag), 2, 13);
        };
        AppLoadTimeHelper.SetAppStateDataInvalidFlag = function (context, appStateDataInvalidFlag) {
            AppLoadTimeHelper.SetActivationInfoField(context, AppLoadTimeHelper.ConvertFlagToBit(appStateDataInvalidFlag), 2, 15);
        };
        AppLoadTimeHelper.SetErrorResult = function (context, result) {
            if (context.LoadTime) {
                context.LoadTime.ErrorResult = result;
                AppLoadTimeHelper.ActivateEndInternal(context);
            }
        };
        AppLoadTimeHelper.ConvertFlagToBit = function (flag) {
            if (flag) {
                return 2;
            }
            else {
                return 1;
            }
        };
        AppLoadTimeHelper.SetActivationInfoField = function (context, value, length, offset) {
            if (context.LoadTime) {
                AppLoadTimeHelper.UpdateActivatingNumber(context);
                context.LoadTime.ActivationInfo = AppLoadTimeHelper.SetBitField(context.LoadTime.ActivationInfo, value, length, offset);
            }
        };
        AppLoadTimeHelper.SetBitField = function (field, value, length, offset) {
            var mask = (Math.pow(2, length) - 1) << offset;
            var cleanField = field & ~mask;
            return cleanField | (value << offset);
        };
        AppLoadTimeHelper.StopStopwatch = function (context, name) {
            if (context.LoadTime && context.Timers && context.Timers[name]) {
                context.Timers[name].Stop();
                AppLoadTimeHelper.UpdateActivatingNumber(context);
            }
        };
        AppLoadTimeHelper.StartStopwatch = function (context, name) {
            if (context.LoadTime && context.Timers && context.Timers[name]) {
                context.Timers[name].Start();
                AppLoadTimeHelper.UpdateActivatingNumber(context);
            }
        };
        AppLoadTimeHelper.UpdateActivatingNumber = function (context) {
            if (context.LoadTime) {
                context.LoadTime.Stage6Time = (context.LoadTime.Stage6Time > AppLoadTimeHelper.activatingNumber) ? context.LoadTime.Stage6Time : AppLoadTimeHelper.activatingNumber;
            }
        };
        AppLoadTimeHelper.ActivateEndInternal = function (context) {
            if (context.LoadTime) {
                AppLoadTimeHelper.StopStopwatch(context, "Stage1Time");
                if (context.Timers) {
                    for (var key in context.Timers) {
                        if (context.Timers[key].ElapsedTime != null) {
                            context.LoadTime[key] = context.Timers[key].ElapsedTime;
                        }
                    }
                }
                LatencyLogger.Instance().LogData(context.LoadTime);
                context.LoadTime = null;
                AppLoadTimeHelper.activatingNumber--;
            }
        };
        AppLoadTimeHelper.activatingNumber = 0;
        return AppLoadTimeHelper;
    })();
    Telemetry.AppLoadTimeHelper = AppLoadTimeHelper;
    var RuntimeTelemetryHelper = (function () {
        function RuntimeTelemetryHelper() {
        }
        RuntimeTelemetryHelper.LogProxyFailure = function (appCorrelationId, methodName, errorInfo) {
            var constructedMessage;
            if (appCorrelationId == null) {
                appCorrelationId = "";
            }
            constructedMessage = OSF.OUtil.formatString("appCorrelationId:{0}, methodName:{1}", appCorrelationId, methodName);
            Object.keys(errorInfo).forEach(function (key) {
                var value = errorInfo[key];
                if (value != null) {
                    value = value.toString();
                }
                constructedMessage += ", " + key + ":" + value;
            });
            Logger.SendULSTraceTag(RuntimeTelemetryHelper.category, ULSTraceLevel.warning, constructedMessage, 0x005c8160);
        };
        RuntimeTelemetryHelper.LogExceptionTag = function (message, exception, appCorrelationId, tagId) {
            var constructedMessage = message;
            if (exception) {
                if (exception.name) {
                    constructedMessage += " Exception name:" + exception.name + ".";
                }
                if (exception.paramName) {
                    constructedMessage += " Param name:" + exception.paramName + ".";
                }
            }
            if (appCorrelationId != null) {
                constructedMessage += " AppCorrelationId:" + appCorrelationId + ".";
            }
            constructedMessage += OSF.OUtil.formatString(" SourceTag: {0}.", tagId);
            Logger.SendULSTraceTag(RuntimeTelemetryHelper.category, ULSTraceLevel.warning, constructedMessage, 0x005c8161);
        };
        RuntimeTelemetryHelper.LogCommonMessage = function (message, tagId) {
            var constructedMessage = message;
            constructedMessage += OSF.OUtil.formatString(" SourceTag: {0}.", tagId);
            Logger.SendULSTraceTag(RuntimeTelemetryHelper.category, ULSTraceLevel.info, constructedMessage, 0x0109b64a);
        };
        RuntimeTelemetryHelper.category = ULSCat.msoulscat_Osf_Runtime;
        return RuntimeTelemetryHelper;
    })();
    Telemetry.RuntimeTelemetryHelper = RuntimeTelemetryHelper;
    var InsertionDialogSessionHelper = (function () {
        function InsertionDialogSessionHelper() {
        }
        InsertionDialogSessionHelper.LogInsertionDialogSession = function (assetId, totalSessionTime, trustPageSessionTime, appInserted, lastActiveTab, lastActiveTabCount) {
            var insertionDialogSessionData = new OSFLog.InsertionDialogSessionUsageData();
            var assetIdNumber = assetId.toLowerCase().indexOf("wa") === 0 ? parseInt(assetId.substring(2), 10) : parseInt(assetId, 10);
            var dialogState = InsertionDialogStateFlags.Undefined;
            if (appInserted) {
                dialogState |= InsertionDialogStateFlags.Inserted;
            }
            else {
                dialogState |= InsertionDialogStateFlags.Canceled;
            }
            if (trustPageSessionTime > 0) {
                dialogState |= InsertionDialogStateFlags.TrustPageVisited;
            }
            insertionDialogSessionData.AssetId = assetIdNumber;
            insertionDialogSessionData.TotalSessionTime = totalSessionTime;
            insertionDialogSessionData.TrustPageSessionTime = trustPageSessionTime;
            insertionDialogSessionData.DialogState = dialogState;
            insertionDialogSessionData.LastActiveTab = lastActiveTab;
            insertionDialogSessionData.LastActiveTabCount = lastActiveTabCount;
            InsertionDialogSessionLogger.Instance().LogData(insertionDialogSessionData);
        };
        return InsertionDialogSessionHelper;
    })();
    Telemetry.InsertionDialogSessionHelper = InsertionDialogSessionHelper;
})(Telemetry || (Telemetry = {}));
var OfficeExt;
(function (OfficeExt) {
    var Parser;
    (function (Parser) {
        var AddInManifestException = (function (_super) {
            __extends(AddInManifestException, _super);
            function AddInManifestException(message) {
                _super.call(this, message);
                this.name = 'AddinManifestError';
                this.message = this.name + ": " + message;
            }
            return AddInManifestException;
        })(Error);
        Parser.AddInManifestException = AddInManifestException;
        var AddInInternalException = (function (_super) {
            __extends(AddInInternalException, _super);
            function AddInInternalException(message) {
                _super.call(this, message);
                this.name = 'AddinInternalError';
                this.message = message;
            }
            return AddInInternalException;
        })(Error);
        Parser.AddInInternalException = AddInInternalException;
        var ManifestResourceManager = (function () {
            function ManifestResourceManager(context) {
                this.Images = {};
                this.Urls = {};
                this.LongStrings = {};
                this.ShortStrings = {};
                var officeAppNode = context.manifest._xmlProcessor.selectSingleNode("o:OfficeApp");
                var resourcesNode = context.selectSingleNode(officeAppNode, ["ov:VersionOverrides", "ov:Resources"]);
                this.parseCollection(context, context.selectChildNodes("bt:Image", resourcesNode, ["bt:Images"]), this.Images);
                this.parseCollection(context, context.selectChildNodes("bt:Url", resourcesNode, ["bt:Urls"]), this.Urls);
                this.parseCollection(context, context.selectChildNodes("bt:String", resourcesNode, ["bt:LongStrings"]), this.LongStrings);
                this.parseCollection(context, context.selectChildNodes("bt:String", resourcesNode, ["bt:ShortStrings"]), this.ShortStrings);
            }
            ManifestResourceManager.prototype.parseCollection = function (context, nodes, map) {
                var len = nodes.length;
                for (var i = 0; i < len; i++) {
                    var node = nodes[i];
                    var id = node.getAttribute("id");
                    map[id] = context.parseLocaleAwareSettingsAndGetValue(node);
                }
            };
            return ManifestResourceManager;
        })();
        Parser.ManifestResourceManager = ManifestResourceManager;
        var ParsingContext = (function () {
            function ParsingContext(hostType, formFactor, entitlement, manifest) {
                this.hostType = hostType;
                this.formFactor = formFactor;
                this.entitlement = entitlement;
                this.manifest = manifest;
                this.resources = new ManifestResourceManager(this);
            }
            ParsingContext.prototype.parseIdRequired = function (node) {
                var id = node.getAttribute("id");
                if (id == null) {
                    throw new AddInManifestException("id requird");
                }
                return id;
            };
            ParsingContext.prototype.parseLabel = function (node) {
                var child = this.manifest._xmlProcessor.selectSingleNode("ov:Label", node);
                if (child != null) {
                    return this.getShortString(child);
                }
                return null;
            };
            ParsingContext.prototype.parseLabelRequired = function (node) {
                var label = this.parseLabel(node);
                if (label == null) {
                    throw new AddInManifestException("label required");
                }
                return label;
            };
            ParsingContext.prototype.parseRequiredSuperTip = function (node) {
                var superTip = this.parseSuperTip(node);
                if (superTip == null) {
                    throw new AddInManifestException("SuperTip required");
                }
                return superTip;
            };
            ParsingContext.prototype.parseSuperTip = function (node) {
                var superTip = null;
                var child = this.manifest._xmlProcessor.selectSingleNode("ov:Supertip", node);
                if (child != null) {
                    var tipNode = child;
                    var title, description;
                    child = this.manifest._xmlProcessor.selectSingleNode("ov:Title", tipNode);
                    title = this.getShortString(child);
                    child = this.manifest._xmlProcessor.selectSingleNode("ov:Description", tipNode);
                    description = this.getLongString(child);
                    superTip = {
                        title: title,
                        description: description
                    };
                }
                return superTip;
            };
            ParsingContext.prototype.parseRequiredIcon = function (node) {
                var icon = this.parseIcon(node);
                if (icon == null) {
                    throw new AddInManifestException("Icon required");
                }
                return icon;
            };
            ParsingContext.prototype.parseIcon = function (node) {
                var iconNodes = this.selectChildNodes("bt:Image", node, ["ov:Icon"]);
                var len = iconNodes.length;
                if (len == 0) {
                    return null;
                }
                var icon = {};
                for (var i = 0; i < len; i++) {
                    var iconNode = iconNodes[i];
                    var size = iconNode.getAttribute("size");
                    icon[size] = this.getImageResource(iconNode);
                }
                return icon;
            };
            ParsingContext.prototype.parseChildControls = function (childNodeName, node, parser) {
                var controls = this.manifest._xmlProcessor.selectNodes(childNodeName, node);
                return this.parseControls(controls, parser);
            };
            ParsingContext.prototype.parseControls = function (nodes, parser) {
                var controls = [];
                var len = nodes.length;
                for (var i = 0; i < len; i++) {
                    var e = nodes[i];
                    var control = parser(this, e);
                    if (control == null) {
                        throw new AddInManifestException("parser must return a control.");
                    }
                    controls.push(control);
                }
                return controls;
            };
            ParsingContext.prototype.parseControlInGroup = function (node) {
                var controlType = node.getAttribute(ParsingContext.typeAttributeName);
                var control;
                switch (controlType) {
                    case "Menu":
                        control = new MenuControl();
                        break;
                    case "Button":
                        control = new Button(controlType);
                        break;
                    default:
                        throw new AddInManifestException("Unsupported control type.");
                }
                control.parse(this, node);
                return control;
            };
            ParsingContext.prototype.parseMenuItem = function (node) {
                var item = new MenuItem();
                item.parse(this, node);
                return item;
            };
            ParsingContext.prototype.parseLocaleAwareSettingsAndGetValue = function (node) {
                var values = {};
                var defaultValue = node.getAttribute("DefaultValue");
                values[this.manifest._defaultLocale] = defaultValue;
                var overrideNodes = this.manifest._xmlProcessor.selectNodes("bt:Override", node);
                if (overrideNodes) {
                    var len = overrideNodes.length;
                    for (var i = 0; i < len; i++) {
                        var node = overrideNodes[i];
                        var locale = node.getAttribute("Locale");
                        var value = node.getAttribute("Value");
                        values[locale] = value;
                    }
                }
                return this.manifest._getDefaultValue(values);
            };
            ParsingContext.prototype.selectSingleNode = function (node, path) {
                if (path == null)
                    return node;
                for (var i = 0; i < path.length; i++) {
                    node = this.manifest._xmlProcessor.selectSingleNode(path[i], node);
                    if (node == null) {
                        break;
                    }
                }
                return node;
            };
            ParsingContext.prototype.selectChildNodes = function (name, node, path) {
                node = this.selectSingleNode(node, path);
                if (node == null) {
                    return [];
                }
                return this.manifest._xmlProcessor.selectNodes(name, node);
            };
            ParsingContext.prototype.getResourceByNode = function (resources, resourceNode) {
                var resid = resourceNode.getAttribute("resid");
                var res = resources[resid];
                if (res == null) {
                    throw new AddInManifestException("resid: " + resid + " not found");
                }
                return res;
            };
            ParsingContext.prototype.getImageResource = function (resourceNode) {
                return this.getResourceByNode(this.resources.Images, resourceNode);
            };
            ParsingContext.prototype.getUrlResource = function (resourceNode) {
                return this.getResourceByNode(this.resources.Urls, resourceNode);
            };
            ParsingContext.prototype.getLongString = function (resourceNode) {
                return this.getResourceByNode(this.resources.LongStrings, resourceNode);
            };
            ParsingContext.prototype.getShortString = function (resourceNode) {
                return this.getResourceByNode(this.resources.ShortStrings, resourceNode);
            };
            ParsingContext.typeAttributeName = "xsi:type";
            return ParsingContext;
        })();
        Parser.ParsingContext = ParsingContext;
        var BuildHelpers = (function () {
            function BuildHelpers() {
            }
            BuildHelpers.buildControls = function (context, controls) {
                var len = controls.length;
                for (var i = 0; i < len; i++) {
                    var child = controls[i];
                    child.apply(context);
                }
            };
            return BuildHelpers;
        })();
        var AddinBuildingContext = (function () {
            function AddinBuildingContext(functionFile, builder) {
                this.functionFile = functionFile;
                this.builder = builder;
            }
            return AddinBuildingContext;
        })();
        var ActionBase = (function () {
            function ActionBase(type) {
                this.type = type;
            }
            ActionBase.prototype.buildAction = function (context) {
            };
            ActionBase.prototype.parse = function (context, node) {
            };
            return ActionBase;
        })();
        var ShowUIAction = (function (_super) {
            __extends(ShowUIAction, _super);
            function ShowUIAction() {
                _super.apply(this, arguments);
            }
            ShowUIAction.prototype.parse = function (context, node) {
                var child = context.manifest._xmlProcessor.selectSingleNode("ov:SourceLocation", node);
                var url = context.getUrlResource(child);
                this.sourceLocation = url;
            };
            return ShowUIAction;
        })(ActionBase);
        var ShowTaskPaneAction = (function (_super) {
            __extends(ShowTaskPaneAction, _super);
            function ShowTaskPaneAction() {
                _super.apply(this, arguments);
            }
            ShowTaskPaneAction.prototype.buildAction = function (context) {
                return context.actionBuilder.buildShowTaskpane(this.sourceLocation);
            };
            return ShowTaskPaneAction;
        })(ShowUIAction);
        var ShowDialogAction = (function (_super) {
            __extends(ShowDialogAction, _super);
            function ShowDialogAction() {
                _super.apply(this, arguments);
            }
            ShowDialogAction.prototype.buildAction = function (context) {
                return context.actionBuilder.buildShowDialog(this.sourceLocation);
            };
            return ShowDialogAction;
        })(ShowUIAction);
        var ExecuteFunctionAction = (function (_super) {
            __extends(ExecuteFunctionAction, _super);
            function ExecuteFunctionAction() {
                _super.apply(this, arguments);
            }
            ExecuteFunctionAction.prototype.buildAction = function (context) {
                return context.actionBuilder.buildCallFunction(context.functionFile, this.functionName);
            };
            ExecuteFunctionAction.prototype.parse = function (context, node) {
                var child = context.manifest._xmlProcessor.selectSingleNode("ov:FunctionName", node);
                this.functionName = context.manifest._xmlProcessor.getNodeValue(child);
            };
            return ExecuteFunctionAction;
        })(ActionBase);
        var UIEntityBase = (function () {
            function UIEntityBase() {
            }
            UIEntityBase.prototype.parse = function (context, node) {
                this.id = context.parseIdRequired(node);
                this.label = context.parseLabelRequired(node);
            };
            return UIEntityBase;
        })();
        var MenuControl = (function (_super) {
            __extends(MenuControl, _super);
            function MenuControl() {
                _super.apply(this, arguments);
            }
            MenuControl.prototype.apply = function (context) {
                context.controlBuilder.startBuildMenu(this);
                BuildHelpers.buildControls(context, this.children);
                context.controlBuilder.endBuildMenu();
            };
            MenuControl.prototype.parse = function (context, node) {
                _super.prototype.parse.call(this, context, node);
                this.icon = context.parseRequiredIcon(node);
                this.superTip = context.parseRequiredSuperTip(node);
                var controls = context.selectChildNodes("ov:Item", node, ["ov:Items"]);
                this.children = context.parseControls(controls, function (context, node) {
                    return context.parseMenuItem(node);
                });
            };
            return MenuControl;
        })(UIEntityBase);
        var ItemControl = (function (_super) {
            __extends(ItemControl, _super);
            function ItemControl(controlType) {
                _super.call(this);
                this.controlType = controlType;
            }
            ItemControl.prototype.apply = function (context) {
                var action = this.action.buildAction(context);
                var control = {
                    id: this.id,
                    label: this.label,
                    icon: this.icon,
                    superTip: this.superTip,
                    actionType: this.actionType,
                    action: action
                };
                this.applyControl(context, control);
            };
            ItemControl.prototype.parse = function (context, node) {
                _super.prototype.parse.call(this, context, node);
                this.icon = this.parseIcon(context, node);
                this.superTip = context.parseRequiredSuperTip(node);
                var child = context.manifest._xmlProcessor.selectSingleNode("ov:Action", node);
                var actionType = child.getAttribute(ParsingContext.typeAttributeName);
                var action;
                switch (actionType) {
                    case "ShowTaskpane":
                        action = new ShowTaskPaneAction(actionType);
                        break;
                    case "ShowDialog":
                        action = new ShowDialogAction(actionType);
                        break;
                    case "ExecuteFunction":
                        action = new ExecuteFunctionAction(actionType);
                        break;
                    default:
                        throw new AddInManifestException("Unsupported action type.");
                }
                action.parse(context, child);
                this.actionType = actionType;
                this.action = action;
            };
            ItemControl.prototype.applyControl = function (context, control) {
                throw new AddInInternalException("Not implemented method applyControl for ItemControl.");
            };
            ItemControl.prototype.parseIcon = function (context, node) {
                return context.parseIcon(node);
            };
            return ItemControl;
        })(UIEntityBase);
        var Button = (function (_super) {
            __extends(Button, _super);
            function Button() {
                _super.apply(this, arguments);
            }
            Button.prototype.applyControl = function (context, control) {
                context.controlBuilder.addButton(control);
            };
            Button.prototype.parseIcon = function (context, node) {
                return context.parseRequiredIcon(node);
            };
            return Button;
        })(ItemControl);
        var MenuItem = (function (_super) {
            __extends(MenuItem, _super);
            function MenuItem() {
                _super.call(this, "MenuItem");
            }
            MenuItem.prototype.applyControl = function (context, control) {
                context.controlBuilder.addMenuItem(control);
            };
            return MenuItem;
        })(ItemControl);
        var RibbonTab = (function () {
            function RibbonTab(isOfficeTab) {
                this.label = null;
                this.isOfficeTab = false;
                this.children = [];
                this.isOfficeTab = isOfficeTab;
            }
            RibbonTab.prototype.apply = function (context) {
                context.builder.ribbonBuilder.startBuildTab(this);
                var len = this.children.length;
                for (var i = 0; i < len; i++) {
                    var child = this.children[i];
                    child.apply(context);
                }
                context.builder.ribbonBuilder.endBuildTab();
            };
            RibbonTab.prototype.parse = function (context, node) {
                this.id = context.parseIdRequired(node);
                if (!this.isOfficeTab) {
                    this.label = context.parseLabelRequired(node);
                }
                this.children = context.parseChildControls("ov:Group", node, function (context, node) {
                    var group = new RibbonGroup();
                    group.parse(context, node);
                    return group;
                });
            };
            return RibbonTab;
        })();
        var RibbonGroup = (function (_super) {
            __extends(RibbonGroup, _super);
            function RibbonGroup() {
                _super.apply(this, arguments);
            }
            RibbonGroup.prototype.parse = function (context, node) {
                _super.prototype.parse.call(this, context, node);
                this.icon = context.parseRequiredIcon(node);
                this.children = context.parseChildControls("ov:Control", node, function (context, node) {
                    return context.parseControlInGroup(node);
                });
            };
            RibbonGroup.prototype.apply = function (context) {
                context.builder.ribbonBuilder.startBuildGroup(this);
                BuildHelpers.buildControls({
                    functionFile: context.functionFile,
                    actionBuilder: context.builder.actionBuilder,
                    controlBuilder: context.builder.ribbonBuilder
                }, this.children);
                context.builder.ribbonBuilder.endBuildGroup();
            };
            return RibbonGroup;
        })(UIEntityBase);
        var RibbonExtensionPoint = (function () {
            function RibbonExtensionPoint(type) {
                this.type = type;
                this.tabs = [];
            }
            RibbonExtensionPoint.prototype.apply = function (context) {
                if (this.type == context.builder.commandSurface) {
                    var len = this.tabs.length;
                    for (var i = 0; i < len; i++) {
                        var tab = this.tabs[i];
                        tab.apply(context);
                    }
                }
            };
            RibbonExtensionPoint.prototype.parse = function (context, node) {
                var tabNode = context.manifest._xmlProcessor.selectSingleNode("ov:OfficeTab", node);
                if (tabNode != null) {
                    var tab = new RibbonTab(true);
                    tab.parse(context, tabNode);
                    this.tabs.push(tab);
                }
                tabNode = context.manifest._xmlProcessor.selectSingleNode("ov:CustomTab", node);
                if (tabNode != null) {
                    var tab = new RibbonTab(false);
                    tab.parse(context, tabNode);
                    this.tabs.push(tab);
                }
            };
            return RibbonExtensionPoint;
        })();
        var OfficeMenuNode = (function () {
            function OfficeMenuNode() {
            }
            OfficeMenuNode.prototype.parse = function (context, node) {
                this.id = context.parseIdRequired(node);
                this.children = context.parseChildControls("ov:Control", node, function (context, node) {
                    return context.parseControlInGroup(node);
                });
            };
            OfficeMenuNode.prototype.apply = function (context) {
                context.controlBuilder.startBuildMenu({
                    id: this.id,
                    label: null,
                    superTip: null
                });
                BuildHelpers.buildControls(context, this.children);
                context.controlBuilder.endBuildMenu();
            };
            return OfficeMenuNode;
        })();
        var ContextMenuExtensionPoint = (function () {
            function ContextMenuExtensionPoint(type) {
                this.type = type;
            }
            ContextMenuExtensionPoint.prototype.apply = function (context) {
                if (this.type == context.builder.commandSurface) {
                    BuildHelpers.buildControls({
                        functionFile: context.functionFile,
                        actionBuilder: context.builder.actionBuilder,
                        controlBuilder: context.builder.contextMenuBuilder
                    }, this.children);
                }
            };
            ContextMenuExtensionPoint.prototype.parse = function (context, node) {
                this.children = context.parseChildControls("ov:OfficeMenu", node, function (context, node) {
                    var menu = new OfficeMenuNode();
                    menu.parse(context, node);
                    return menu;
                });
            };
            return ContextMenuExtensionPoint;
        })();
        var VersionOverrides = (function () {
            function VersionOverrides() {
                this.cacheableUrls = [];
                this.ExtensionPoints = [];
            }
            VersionOverrides.prototype.cacheableResources = function () {
                return this.cacheableUrls;
            };
            VersionOverrides.prototype.apply = function (builder) {
                builder.startApplyAddin({
                    entitlement: this.extensionEntitlement,
                    manifest: this.extensionManifest,
                    overrides: {
                        type: this.type,
                        description: this.description
                    }
                });
                var context = new AddinBuildingContext(this.functionFile, builder);
                var len = this.ExtensionPoints.length;
                for (var i = 0; i < len; i++) {
                    var ext = this.ExtensionPoints[i];
                    ext.apply(context);
                }
                builder.endApplyAddin();
            };
            VersionOverrides.prototype.parse = function (context, overridesNode) {
                this.type = overridesNode.getAttribute(ParsingContext.typeAttributeName);
                var node = context.manifest._xmlProcessor.selectSingleNode("ov:Description", overridesNode);
                if (node != null) {
                    this.description = context.getLongString(node);
                }
                var hostNodes = context.selectChildNodes("ov:Host", overridesNode, ["ov:Hosts"]);
                var formFactorNode = null;
                var len = hostNodes.length;
                for (var i = 0; i < len; i++) {
                    var hostNode = hostNodes[i];
                    var hostType = hostNode.getAttribute(ParsingContext.typeAttributeName);
                    if (hostType == context.hostType) {
                        formFactorNode = context.manifest._xmlProcessor.
                            selectSingleNode("ov:" + context.formFactor, hostNode);
                        break;
                    }
                }
                node = context.manifest._xmlProcessor.selectSingleNode("ov:FunctionFile", formFactorNode);
                this.functionFile = context.getUrlResource(node);
                var extNodes = context.manifest._xmlProcessor.selectNodes("ov:ExtensionPoint", formFactorNode);
                len = extNodes.length;
                for (var i = 0; i < len; i++) {
                    var extNode = extNodes[i];
                    var extType = extNode.getAttribute(ParsingContext.typeAttributeName);
                    var ext;
                    switch (extType) {
                        case "Events":
                            throw "Not Implemented";
                        case "PrimaryCommandSurface":
                            ext = new RibbonExtensionPoint(extType);
                            break;
                        case "ContextMenu":
                            ext = new ContextMenuExtensionPoint(extType);
                            break;
                        default:
                            throw new AddInManifestException("Unsupported extension type.");
                    }
                    ext.parse(context, extNode);
                    this.ExtensionPoints.push(ext);
                }
                var images = context.resources.Images;
                for (var e in images) {
                    this.cacheableUrls.push(images[e]);
                }
                this.extensionEntitlement = context.entitlement;
                this.extensionManifest = context.manifest;
            };
            return VersionOverrides;
        })();
        var emptyAddin = (function () {
            function emptyAddin() {
            }
            emptyAddin.prototype.cacheableResources = function () {
                return [];
            };
            emptyAddin.prototype.apply = function (builder) {
            };
            return emptyAddin;
        })();
        var AddinCommandsManifestParser = (function () {
            function AddinCommandsManifestParser(hostType, formFactor) {
                this.hostType = hostType;
                this.formFactor = formFactor;
            }
            AddinCommandsManifestParser.prototype.parseExtensions = function (entitlement, manifest, errorManager) {
                try {
                    var officeAppNode = manifest._xmlProcessor.selectSingleNode("o:OfficeApp");
                    var overridesNode = manifest._xmlProcessor.selectSingleNode("ov:VersionOverrides", officeAppNode);
                    if (overridesNode == null) {
                        return new emptyAddin();
                    }
                    var versionOverrides = new VersionOverrides();
                    versionOverrides.parse(new ParsingContext(this.hostType, this.formFactor, entitlement, manifest), overridesNode);
                    return versionOverrides;
                }
                catch (ex) {
                    if (ex instanceof AddInManifestException) {
                        errorManager.setErrorMessageForAddin(entitlement.assetId, Strings.OsfRuntime.L_AddinCommands_AddinNotSupported_Message);
                    }
                    throw ex;
                }
            };
            return AddinCommandsManifestParser;
        })();
        Parser.AddinCommandsManifestParser = AddinCommandsManifestParser;
    })(Parser = OfficeExt.Parser || (OfficeExt.Parser = {}));
})(OfficeExt || (OfficeExt = {}));
var OfficeExt;
(function (OfficeExt) {
    var AddinCommandsManifestManagerImpl = (function () {
        function AddinCommandsManifestManagerImpl() {
        }
        AddinCommandsManifestManagerImpl.createManifestForAddinAction = function (addinManifest, sourceLocation) {
            var settings = new OSF.Manifest.ExtensionSettings();
            var formFactor = OSF.FormFactor.Default;
            settings._defaultHeight = null;
            settings._defaultWidth = null;
            settings._sourceLocations = {};
            settings._sourceLocations[addinManifest._UILocale] = sourceLocation;
            var template = addinManifest;
            var manifest = new OSF.Manifest.Manifest(function (manifest) {
                manifest._xmlProcessor = template._xmlProcessor;
                manifest._displayNames = template._displayNames;
                manifest._iconUrls = template._iconUrls;
                manifest._extensionSettings = { formFactor: settings };
                manifest._highResolutionIconUrls = template._highResolutionIconUrls;
                manifest._target = template._target;
                manifest._id = template._id;
                manifest._version = template._version;
                manifest._providerName = template._providerName;
                manifest._idIssuer = template._idIssuer;
                manifest._alternateId = template._alternateId;
                manifest._defaultLocale = template._defaultLocale;
                manifest._signature = template._signature;
                manifest._capabilities = template._capabilities;
                manifest._hosts = template._hosts;
                manifest._descriptions = template._descriptions;
                manifest._appDomains = template._appDomains;
                manifest._permissions = template._permissions;
                manifest._requirements = template._requirements;
            }, template._UILocale);
            return manifest;
        };
        AddinCommandsManifestManagerImpl.cacheManifestForAction = function (manifest, assetId, appVersion) {
            OSF.OsfManifestManager.cacheManifest(assetId, appVersion, manifest);
        };
        AddinCommandsManifestManagerImpl.purgeManifestForAction = function (assetId, appVersion) {
            OSF.OsfManifestManager.purgeManifest(assetId, appVersion);
        };
        return AddinCommandsManifestManagerImpl;
    })();
    OfficeExt.AddinCommandsManifestManager = AddinCommandsManifestManagerImpl;
})(OfficeExt || (OfficeExt = {}));
var OfficeExt;
(function (OfficeExt) {
    OfficeExt.AddinActionContextMap = {};
    var ErrorCodes;
    (function (ErrorCodes) {
        ErrorCodes[ErrorCodes["ooeTimeout"] = -1] = "ooeTimeout";
        ErrorCodes[ErrorCodes["ooeSuccess"] = 0] = "ooeSuccess";
        ErrorCodes[ErrorCodes["ooeOperationNotSupported"] = 5000] = "ooeOperationNotSupported";
    })(ErrorCodes || (ErrorCodes = {}));
    var ControlBoundActionBuilder = (function () {
        function ControlBoundActionBuilder(contextProvider, entitlement, addinManifest) {
            this.contextProvider = contextProvider;
            this.manifest = addinManifest;
            this.entitlement = entitlement;
        }
        ControlBoundActionBuilder.prototype.buildShowTaskpane = function (sourceLocation) {
            return this.CreateManifestAndActivationContextForTaskpaneCommand(sourceLocation);
        };
        ControlBoundActionBuilder.prototype.buildShowDialog = function (sourceLocation) {
            return this.CreateManifestAndActivationContextForDialogCommand(sourceLocation);
        };
        ControlBoundActionBuilder.prototype.buildCallFunction = function (functionFile, functionName) {
            return this.CreateManifestAndActivationContextForUILessCommand(functionFile, functionName);
        };
        ControlBoundActionBuilder.prototype.CreateManifestAndActivationContextForTaskpaneCommand = function (sourceLocation) {
            var onHostReady = function (osfControlId, result) {
                if (result != ErrorCodes.ooeSuccess) {
                    return;
                }
            };
            return this.CreateManifestAndActivationContext(sourceLocation, function (host, entitlement, actionManifest) {
                host.showTaskpane(entitlement, actionManifest, onHostReady);
            });
        };
        ControlBoundActionBuilder.prototype.CreateManifestAndActivationContextForDialogCommand = function (sourceLocation) {
            var onHostReady = function (osfControlId, result) {
                if (result != ErrorCodes.ooeSuccess) {
                    return;
                }
            };
            return this.CreateManifestAndActivationContext(sourceLocation, function (host, entitlement, actionManifest) {
                host.showDialog(entitlement, actionManifest, onHostReady);
            });
        };
        ControlBoundActionBuilder.prototype.CreateManifestAndActivationContext = function (sourceLocation, callback) {
            var _this = this;
            var actionManifest = OfficeExt.AddinCommandsManifestManager.createManifestForAddinAction(this.manifest, sourceLocation);
            return function (controlID) {
                var context = _this.contextProvider.createActionContext(controlID, function (host) {
                    var entitlement = {
                        assetId: _this.entitlement.assetId + controlID,
                        appVersion: _this.entitlement.appVersion,
                        storeId: _this.entitlement.storeId,
                        storeType: OSF.StoreType.InMemory,
                        targetType: OSF.OsfControlTarget.TaskPane
                    };
                    OfficeExt.AddinCommandsManifestManager.cacheManifestForAction(actionManifest, entitlement.assetId, entitlement.appVersion);
                    callback(host, entitlement, actionManifest);
                });
                OfficeExt.AddinActionContextMap[controlID] = context;
            };
        };
        ControlBoundActionBuilder.prototype.CreateManifestAndActivationContextForUILessCommand = function (functionFile, functionName) {
            var _this = this;
            var onHostReady = function (osfControlId, result) {
                if (result != ErrorCodes.ooeSuccess) {
                    return;
                }
                AddinCommandsRuntimeManager.invokeAppCommand(osfControlId, functionName, null, function (status, data) {
                });
            };
            var actionManifest = OfficeExt.AddinCommandsManifestManager.createManifestForAddinAction(this.manifest, functionFile);
            return function (controlID) {
                var context = _this.contextProvider.createActionContext(controlID, function (host) {
                    var entitlement = {
                        assetId: _this.entitlement.assetId + "##UILessContainer##",
                        appVersion: _this.entitlement.appVersion,
                        storeId: _this.entitlement.storeId,
                        storeType: OSF.StoreType.InMemory,
                        targetType: OSF.OsfControlTarget.TaskPane
                    };
                    OfficeExt.AddinCommandsManifestManager.cacheManifestForAction(actionManifest, entitlement.assetId, entitlement.appVersion);
                    host.createUILessSandbox(entitlement, actionManifest, function (osfControlId, result) {
                        if (result == ErrorCodes.ooeSuccess) {
                            var entry = AddinCommandsRuntimeManager.ensureOsfHostEntry(osfControlId, host);
                        }
                        onHostReady(osfControlId, result);
                    });
                });
                OfficeExt.AddinActionContextMap[controlID] = context;
            };
        };
        return ControlBoundActionBuilder;
    })();
    var ControlDeleterActionBuilder = (function () {
        function ControlDeleterActionBuilder(contextProvider, entitlement) {
            this.contextProvider = contextProvider;
            this.entitlement = entitlement;
        }
        ControlDeleterActionBuilder.prototype.buildShowTaskpane = function (sourceLocation) {
            var _this = this;
            return function (controlID) {
                delete OfficeExt.AddinActionContextMap[controlID];
                OfficeExt.AddinCommandsManifestManager.purgeManifestForAction(_this.entitlement.assetId + controlID, _this.entitlement.appVersion);
            };
        };
        ControlDeleterActionBuilder.prototype.buildShowDialog = function (sourceLocation) {
            var _this = this;
            return function (controlID) {
                delete OfficeExt.AddinActionContextMap[controlID];
                OfficeExt.AddinCommandsManifestManager.purgeManifestForAction(_this.entitlement.assetId + controlID, _this.entitlement.appVersion);
            };
        };
        ControlDeleterActionBuilder.prototype.buildCallFunction = function (functionFile, functionName) {
            var _this = this;
            return function (controlID) {
                delete OfficeExt.AddinActionContextMap[controlID];
                OfficeExt.AddinCommandsManifestManager.purgeManifestForAction(_this.entitlement.assetId + "##UILessContainer##", _this.entitlement.appVersion);
            };
        };
        return ControlDeleterActionBuilder;
    })();
    var DefaultControlActionBinder = (function () {
        function DefaultControlActionBinder() {
        }
        DefaultControlActionBinder.createActionBuilder = function (contextProvider, entitlement, addinManifest) {
            return new ControlBoundActionBuilder(contextProvider, entitlement, addinManifest);
        };
        DefaultControlActionBinder.createDeleterActionBuilder = function (contextProvider, entitlement, addinManifest) {
            return new ControlDeleterActionBuilder(contextProvider, entitlement);
        };
        DefaultControlActionBinder.bindAction = function (controlID, action) {
            var callBack = action;
            callBack(controlID);
        };
        DefaultControlActionBinder.retrieveActionContext = function (controlID) {
            return OfficeExt.AddinActionContextMap[controlID];
        };
        return DefaultControlActionBinder;
    })();
    OfficeExt.DefaultControlActionBinder = DefaultControlActionBinder;
    var OsfHostEntry = (function () {
        function OsfHostEntry() {
            this.invocationQueue = {};
        }
        OsfHostEntry.prototype.invokeAppCommand = function (appCommandId, callbackName, eventObjStr, onComplete, timeout) {
            var args = {
                dispid: OSF.DDA.EventDispId.dispidAppCommandInvokedEvent,
                controlId: this.osfControlId
            };
            args[0] = appCommandId;
            args[1] = callbackName;
            args[2] = eventObjStr;
            timeout = timeout ? timeout : OsfHostEntry.defaultTimeout;
            var e = {
                args: args,
                onComplete: onComplete,
                timeout: timeout,
                timeoutTimer: null
            };
            var queue = this.invocationQueue[callbackName];
            if (queue == null) {
                queue = [];
                this.invocationQueue[callbackName] = queue;
            }
            queue.push(e);
            if (queue.length == 1) {
                this.trySendInvocation(e);
            }
        };
        OsfHostEntry.prototype.invocationCompleted = function (appCommandId, status, data) {
            var _this = this;
            for (var callbackName in this.invocationQueue) {
                var queue = this.invocationQueue[callbackName];
                var e = queue[0];
                if (appCommandId == e.args[0]) {
                    if (queue.length > 1) {
                        queue.splice(0, 1);
                        var next = queue[0];
                        window.setTimeout(function () {
                            _this.trySendInvocation(next);
                        }, 0);
                    }
                    else {
                        delete this.invocationQueue[callbackName];
                    }
                    if (e.timeoutTimer != null) {
                        window.clearTimeout(e.timeoutTimer);
                        e.timeoutTimer = null;
                    }
                    e.onComplete(status, data);
                    return;
                }
            }
        };
        OsfHostEntry.prototype.sendInvocations = function (handler) {
            this.eventHandler = handler;
            if (handler == null)
                return;
            for (var callbackName in this.invocationQueue) {
                var queue = this.invocationQueue[callbackName];
                var e = queue[0];
                this.sendInvocation(handler, e);
            }
        };
        OsfHostEntry.prototype.trySendInvocation = function (e) {
            var handler = this.eventHandler;
            if (handler != null) {
                this.sendInvocation(handler, e);
            }
        };
        OsfHostEntry.prototype.sendInvocation = function (handler, e) {
            var _this = this;
            var args = e.args;
            handler(args);
            e.timeoutTimer = window.setTimeout(function () {
                _this.invocationTimeout(args[0]);
            }, e.timeout);
        };
        OsfHostEntry.prototype.invocationTimeout = function (appCommandId) {
            this.invocationCompleted(appCommandId, ErrorCodes.ooeTimeout, null);
        };
        OsfHostEntry.prototype.disposeHost = function () {
            for (var callbackName in this.invocationQueue) {
                return;
            }
            this.host.disposeHost(this.osfControlId);
            if (this.onDisposed != null) {
                this.onDisposed(this);
            }
        };
        OsfHostEntry.defaultTimeout = 300 * 1000;
        return OsfHostEntry;
    })();
    var AddinCommandsRuntimeManager = (function () {
        function AddinCommandsRuntimeManager() {
        }
        AddinCommandsRuntimeManager.registerEvent = function (eventHandler, callback, params) {
            AddinCommandsRuntimeManager.registerEventInternal(params["eventDispId"], params["controlId"], params["targetId"], eventHandler, function (errorCode) {
                if (callback) {
                    callback(errorCode == 0);
                }
            });
        };
        AddinCommandsRuntimeManager.unregisterEvent = function (eventHandler, callback, params) {
            AddinCommandsRuntimeManager.unregisterEventInternal(params["eventDispId"], params["controlId"], params["targetId"], eventHandler, function (errorCode) {
                if (callback) {
                    callback(errorCode == 0);
                }
            });
        };
        AddinCommandsRuntimeManager.invokeAppCommand = function (osfControlId, callbackName, eventObj, onComplete) {
            var entry = AddinCommandsRuntimeManager.getOrCreateOsfHostEntry(osfControlId);
            var appCommandId = AddinCommandsRuntimeManager.getNextAppCommandId();
            var eventObjStr = JSON.stringify(eventObj);
            entry.invokeAppCommand(appCommandId, callbackName, eventObjStr, onComplete);
        };
        AddinCommandsRuntimeManager.invocationCompleted = function (osfControlId, args) {
            var entry = AddinCommandsRuntimeManager.getOrCreateOsfHostEntry(osfControlId);
            var appCommandId = args[0];
            var status = args[1];
            var data = JSON.parse(args[2]);
            entry.invocationCompleted(appCommandId, status, data);
        };
        AddinCommandsRuntimeManager.ensureOsfHostEntry = function (osfControlId, host) {
            var entry = AddinCommandsRuntimeManager.getOrCreateOsfHostEntry(osfControlId);
            if (entry.host == null) {
                entry.host = host;
            }
            return entry;
        };
        AddinCommandsRuntimeManager.getNextAppCommandId = function () {
            return "AppCmd" + AddinCommandsRuntimeManager._nextAppCommandId++;
        };
        AddinCommandsRuntimeManager.getOrCreateOsfHostEntry = function (osfControlId) {
            var entry = AddinCommandsRuntimeManager._invocationQueues[osfControlId];
            if (entry == null) {
                entry = new OsfHostEntry();
                entry.osfControlId = osfControlId;
                entry.onDisposed = function (e) {
                    delete AddinCommandsRuntimeManager._invocationQueues[e.osfControlId];
                };
                AddinCommandsRuntimeManager._invocationQueues[osfControlId] = entry;
            }
            return entry;
        };
        AddinCommandsRuntimeManager.registerEventInternal = function (eventDispId, controlId, targetId, handler, callback) {
            switch (eventDispId) {
                case OSF.DDA.EventDispId.dispidAppCommandInvokedEvent:
                    {
                        var entry = AddinCommandsRuntimeManager.getOrCreateOsfHostEntry(controlId);
                        callback(ErrorCodes.ooeSuccess);
                        entry.sendInvocations(handler);
                    }
                    return;
                default:
                    break;
            }
            callback(ErrorCodes.ooeOperationNotSupported);
        };
        AddinCommandsRuntimeManager.unregisterEventInternal = function (eventDispId, controlId, targetId, handler, callback) {
            switch (eventDispId) {
                case OSF.DDA.EventDispId.dispidAppCommandInvokedEvent:
                    {
                        var entry = AddinCommandsRuntimeManager.getOrCreateOsfHostEntry(controlId);
                        if (entry.eventHandler == handler) {
                            entry.eventHandler = null;
                        }
                        callback(ErrorCodes.ooeSuccess);
                    }
                    return;
                default:
                    break;
            }
            callback(ErrorCodes.ooeOperationNotSupported);
        };
        AddinCommandsRuntimeManager._nextAppCommandId = 0;
        AddinCommandsRuntimeManager._invocationQueues = {};
        return AddinCommandsRuntimeManager;
    })();
    OfficeExt.AddinCommandsRuntimeManager = AddinCommandsRuntimeManager;
})(OfficeExt || (OfficeExt = {}));
var SPCatalogDataProvider = (function () {
    function SPCatalogDataProvider() {
    }
    SPCatalogDataProvider.prototype.validateAndCacheNewManifest = function (contextActivationMgr, assetId, appVersion, manifestStr, manifest) {
        return false;
    };
    SPCatalogDataProvider.prototype.localCacheAddinCommands = function (contextActivationMgr, assetId, version, storeId) {
    };
    SPCatalogDataProvider.prototype.getEntitlementsAsync = function (contextActivationMgr, forAddinCommands, onGetEntitlementCompleted) {
        if (!contextActivationMgr || !contextActivationMgr._enableMyOrg) {
            onGetEntitlementCompleted({ "statusCode": OSF.InvokeResultCode.E_CATALOG_NO_APPS, "value": null });
            return;
        }
        var myAddins;
        var onGetSPCatalogEntitlementsCompleted = function (asyncResult) {
            if (asyncResult != null && asyncResult.errorCode === OSF.InvokeResultCode.S_OK && asyncResult.value) {
                var entitlementsResult = asyncResult.value;
                myAddins = new Array();
                for (var i = 0; i < entitlementsResult.length; i++) {
                    var item = entitlementsResult[i];
                    var entitlement = {
                        assetId: item[1],
                        appVersion: item[4],
                        storeId: item[11],
                        storeType: OSF.StoreType.SPCatalog,
                        targetType: item[3]
                    };
                    myAddins.push(entitlement);
                }
            }
            onGetEntitlementCompleted(myAddins);
        };
        var onGetCorporateCatalogUrlCompleted = function (asyncResult) {
            if (asyncResult != null && asyncResult.value) {
                contextActivationMgr._myOrgCatalogUrl = asyncResult.value;
                var params = { "storeType": OSF.StoreTypeEnum.Catalog, "refresh": false, "fromInsertDialog": false, "callbackFunc": onGetSPCatalogEntitlementsCompleted };
                contextActivationMgr.queryEntitlementDetails(params);
            }
            else {
                onGetSPCatalogEntitlementsCompleted({ "statusCode": OSF.InvokeResultCode.E_GENERIC_ERROR, "value": null });
            }
        };
        var retries = 0;
        var onCreateSharePointProxyCompleted = function (clientEndPoint) {
            if (clientEndPoint) {
                var context = { "webUrl": contextActivationMgr._spBaseUrl, "clientEndPoint": clientEndPoint };
                var params = { "webUrl": contextActivationMgr._spBaseUrl };
                OSF.OsfManifestManager._invokeProxyMethodAsync(context, "OEM_getSPCatalogUrlAsync", onGetCorporateCatalogUrlCompleted, params);
            }
            else {
                if (retries++ < OSF.Constants.AuthenticatedConnectMaxTries) {
                    setTimeout(function () {
                        contextActivationMgr._createSharePointIFrameProxy(contextActivationMgr._spBaseUrl, onCreateSharePointProxyCompleted);
                    }, 500);
                }
                else {
                    onGetSPCatalogEntitlementsCompleted({ "statusCode": OSF.InvokeResultCode.E_GENERIC_ERROR, "value": null });
                }
            }
        };
        if (contextActivationMgr._spBaseUrl && contextActivationMgr._iframeProxies && contextActivationMgr._iframeProxies[contextActivationMgr._spBaseUrl] && contextActivationMgr._iframeProxies[contextActivationMgr._spBaseUrl].clientEndPoint) {
            onCreateSharePointProxyCompleted(contextActivationMgr._iframeProxies[contextActivationMgr._spBaseUrl].clientEndPoint);
        }
        else {
            contextActivationMgr._createSharePointIFrameProxy(contextActivationMgr._spBaseUrl, onCreateSharePointProxyCompleted);
        }
    };
    SPCatalogDataProvider.prototype.getManifestAsync = function (context, onGetManifestCompleted) {
        OSF.OUtil.validateParamObject(context, {
            "contextActivationMgr": { type: Object, mayBeNull: false },
            "referenceInUse": { type: Object, mayBeNull: false }
        }, null);
        var onGetEntitlementsCompleted = function (asyncResult) {
            OSF.OUtil.writeProfilerMark(OSF.OsfOfficeExtensionManagerPerfMarker.GetEntitlementEnd);
            if (asyncResult.statusCode === OSF.ProxyCallStatusCode.Succeeded && asyncResult.value) {
                var reference = asyncResult.context.referenceInUse;
                var entitlements = asyncResult.value.entitlements;
                var entitlementCount = entitlements.length;
                var newestEntitlement = null;
                for (var i = 0; i < entitlementCount; i++) {
                    var entitlement = entitlements[i];
                    if (entitlement.OfficeExtensionID && reference.id && entitlement.OfficeExtensionID.toLowerCase() === reference.id.toLowerCase()) {
                        if (!newestEntitlement || OSF.OsfManifestManager.versionLessThan(newestEntitlement.OfficeExtensionVersion, entitlement.OfficeExtensionVersion)) {
                            newestEntitlement = entitlement;
                        }
                    }
                }
                if (newestEntitlement != null) {
                    asyncResult.context.manifestUrl = newestEntitlement.EncodedAbsUrl;
                    asyncResult.context.appInstanceId = newestEntitlement.AppInstanceID;
                    asyncResult.context.productId = newestEntitlement.ProductID;
                    OSF.OsfManifestManager.getManifestAsync(asyncResult.context, Function.createDelegate(this, onGetManifestCompleted));
                }
                else {
                    onGetManifestCompleted({ "statusCode": OSF.ProxyCallStatusCode.Failed, "value": null });
                }
            }
            else {
                onGetManifestCompleted({ "statusCode": OSF.ProxyCallStatusCode.Failed, "value": null });
            }
        };
        var cachedManifest = OSF.OsfManifestManager.getCachedManifest(context.referenceInUse.id, context.referenceInUse.version);
        if (cachedManifest == null) {
            context.webUrl = context.referenceInUse.storeLocator;
            context.osfControl = { "_contextActivationMgr": context.contextActivationMgr, "_telemetryContext": {} };
            context.noTargetType = true;
            context.hostType = context.contextActivationMgr._hostType;
            OSF.OUtil.writeProfilerMark(OSF.OsfOfficeExtensionManagerPerfMarker.GetEntitlementStart);
            OSF.OsfManifestManager.getCorporateCatalogEntitlementsAsync(context, Function.createDelegate(this, onGetEntitlementsCompleted));
        }
        else {
            onGetManifestCompleted({ "statusCode": OSF.ProxyCallStatusCode.Succeeded, "value": cachedManifest, "context": context });
        }
    };
    return SPCatalogDataProvider;
})();
var OfficeExt;
(function (OfficeExt) {
    var UploadFileCatalogDataProvider = (function () {
        function UploadFileCatalogDataProvider(cacheManager) {
            this._cacheManager = cacheManager;
        }
        UploadFileCatalogDataProvider.prototype.validateAndCacheNewManifest = function (contextActivationMgr, assetId, appVersion, manifestStr, manifest) {
            if (this.validateAddinIdAndVersion(assetId, appVersion)) {
                this.cacheManifest(contextActivationMgr, assetId, appVersion, manifestStr, manifest);
                return true;
            }
            return false;
        };
        UploadFileCatalogDataProvider.prototype.localCacheAddinCommands = function (contextActivationMgr, assetId, version, storeId) {
            this.cacheMyAddins(this.getUploadFileDevCatalogAddinCommandsMyAddinsCacheKey(contextActivationMgr), assetId);
        };
        UploadFileCatalogDataProvider.prototype.getEntitlementsAsync = function (contextActivationMgr, forAddinCommands, onGetEntitlementCompleted) {
            var _this = this;
            var assetidList = new Array();
            var myAddins = new Array();
            var cacheKey = forAddinCommands ? this.getUploadFileDevCatalogAddinCommandsMyAddinsCacheKey(contextActivationMgr) : this.getUploadFileDevCatalogAddinsCacheKey(contextActivationMgr);
            var cachedMyAddins = this._cacheManager.GetCacheItem(cacheKey);
            if (cachedMyAddins != null) {
                cachedMyAddins.forEach(function (cachedAssetId, index, array) {
                    var cachedManifestStr = _this._cacheManager.GetCacheItem(_this.getUploadFileDevCatalogManifestCacheKey(contextActivationMgr, cachedAssetId));
                    if (cachedManifestStr != null) {
                        var cachedManifest = new OSF.Manifest.Manifest(cachedManifestStr, contextActivationMgr.getAppUILocale());
                        var entitlement = {
                            assetId: cachedAssetId,
                            appVersion: cachedManifest.getMarketplaceVersion(),
                            storeId: "developer",
                            storeType: OSF.StoreType.UploadFileDevCatalog,
                            targetType: cachedManifest.getTarget()
                        };
                        myAddins.push(entitlement);
                        assetidList.push(cachedAssetId);
                    }
                });
                if (assetidList.length != cachedMyAddins.length) {
                    if (assetidList.length > 0) {
                        this._cacheManager.SetCacheItem(cacheKey, assetidList);
                    }
                    else {
                        this._cacheManager.RemoveCacheItem(cacheKey);
                    }
                }
            }
            onGetEntitlementCompleted(myAddins);
        };
        UploadFileCatalogDataProvider.prototype.getManifestAsync = function (context, onGetManifestCompleted) {
            OSF.OUtil.validateParamObject(context, {
                "contextActivationMgr": { type: Object, mayBeNull: false },
                "referenceInUse": { type: Object, mayBeNull: false }
            }, null);
            var cachedManifest = OSF.OsfManifestManager.getCachedManifest(context.referenceInUse.id, context.referenceInUse.version);
            if (cachedManifest == null) {
                var manifestStr = this._cacheManager.GetCacheItem(this.getUploadFileDevCatalogManifestCacheKey(context.contextActivationMgr, context.referenceInUse.id));
                if (manifestStr != null) {
                    cachedManifest = new OSF.Manifest.Manifest(manifestStr, context.contextActivationMgr.getAppUILocale());
                }
            }
            onGetManifestCompleted(cachedManifest != null ? { "statusCode": OSF.ProxyCallStatusCode.Succeeded, "value": cachedManifest, "context": context } : { "statusCode": OSF.ProxyCallStatusCode.Failed, "value": null });
        };
        UploadFileCatalogDataProvider.prototype.cleanMyAddins = function (contextActivationMgr) {
            var cacheKey = this.getUploadFileDevCatalogAddinsCacheKey(contextActivationMgr);
            var myAddins = this._cacheManager.GetCacheItem(cacheKey);
            if (myAddins != null) {
                this._cacheManager.RemoveCacheItem(cacheKey);
            }
        };
        UploadFileCatalogDataProvider.prototype.validateAddinIdAndVersion = function (marketplaceId, marketplaceVersion) {
            return UploadFileCatalogDataProvider.idRegExp.test(marketplaceId) && UploadFileCatalogDataProvider.versionRegExp.test(marketplaceVersion);
        };
        UploadFileCatalogDataProvider.prototype.cacheManifest = function (contextActivationMgr, assetId, appVersion, manifestStr, manifest) {
            OSF.OsfManifestManager.cacheManifest(assetId, appVersion, manifest);
            this.cacheMyAddins(this.getUploadFileDevCatalogAddinsCacheKey(contextActivationMgr), assetId);
            this._cacheManager.SetCacheItem(this.getUploadFileDevCatalogManifestCacheKey(contextActivationMgr, assetId), manifestStr);
        };
        UploadFileCatalogDataProvider.prototype.cacheMyAddins = function (myAddinsCacheKey, assetId) {
            var myAddins = this._cacheManager.GetCacheItem(myAddinsCacheKey);
            if (myAddins == null) {
                myAddins = new Array();
            }
            if (myAddins.lastIndexOf(assetId) < 0) {
                myAddins.push(assetId);
                this._cacheManager.SetCacheItem(myAddinsCacheKey, myAddins);
            }
        };
        UploadFileCatalogDataProvider.prototype.getUploadFileDevCatalogAddinsCacheKey = function (contextActivationMgr) {
            return UploadFileCatalogDataProvider.UploadFileCatalogCacheKey.myAddins + "." + contextActivationMgr.getAppName() + "." + contextActivationMgr.getUserNameHashCode();
        };
        UploadFileCatalogDataProvider.prototype.getUploadFileDevCatalogAddinCommandsMyAddinsCacheKey = function (contextActivationMgr) {
            return UploadFileCatalogDataProvider.UploadFileCatalogCacheKey.addinCommandsMyAddins + "." + contextActivationMgr.getAppName() + "." + contextActivationMgr.getUserNameHashCode();
        };
        UploadFileCatalogDataProvider.prototype.getUploadFileDevCatalogManifestCacheKey = function (contextActivationMgr, assetId) {
            return UploadFileCatalogDataProvider.UploadFileCatalogCacheKey.manifest + "." + contextActivationMgr.getAppName() + "." + assetId;
        };
        UploadFileCatalogDataProvider.UploadFileCatalogCacheKeyPrefix = "__OSF_UPLOADFILE.";
        UploadFileCatalogDataProvider.UploadFileCatalogCacheKey = {
            myAddins: UploadFileCatalogDataProvider.UploadFileCatalogCacheKeyPrefix + "MyAddins",
            addinCommandsMyAddins: UploadFileCatalogDataProvider.UploadFileCatalogCacheKeyPrefix + "AddinCommandsMyAddins",
            manifest: UploadFileCatalogDataProvider.UploadFileCatalogCacheKeyPrefix + "Manifest"
        };
        UploadFileCatalogDataProvider.idRegExp = new RegExp("[\\w-]+");
        UploadFileCatalogDataProvider.versionRegExp = new RegExp("(\\d{1,2}\.)*\\d{1,2}");
        return UploadFileCatalogDataProvider;
    })();
    var AddinDataProviderFactoryImpl = (function () {
        function AddinDataProviderFactoryImpl() {
        }
        AddinDataProviderFactoryImpl.createAddinDataProvider = function (providerType, contextActivationMgr) {
            var dataProvider = null;
            switch (providerType) {
                case OSF.StoreType.UploadFileDevCatalog:
                    dataProvider = new UploadFileCatalogDataProvider(new OfficeExt.AppsDataCacheManager(OSF.OUtil.getLocalStorage(), new OfficeExt.SafeSerializer()));
                    break;
                case OSF.StoreType.OMEX:
                    dataProvider = OmexDataProvider.GetInstance(new OfficeExt.AppsDataCacheManager(OSF.OUtil.getLocalStorage(), new OfficeExt.SafeSerializer()));
                    break;
                case OSF.StoreType.SPCatalog:
                    dataProvider = new SPCatalogDataProvider();
                    break;
                default:
                    break;
            }
            return dataProvider;
        };
        return AddinDataProviderFactoryImpl;
    })();
    OfficeExt.AddinDataProviderFactory = AddinDataProviderFactoryImpl;
})(OfficeExt || (OfficeExt = {}));
if (!isDefinePropertySupported) {
    if (!OsfMsAjaxFactory.msAjaxError) {
        OsfMsAjaxFactory.msAjaxError = Error;
    }
    if (!OsfMsAjaxFactory.msAjaxSerializer) {
        OsfMsAjaxFactory.msAjaxSerializer = Sys.Serialization.JavaScriptSerializer;
    }
    if (!OsfMsAjaxFactory.msAjaxString) {
        OsfMsAjaxFactory.msAjaxString = String;
    }
    if (!OsfMsAjaxFactory.msAjaxDebug) {
        OsfMsAjaxFactory.msAjaxDebug = Sys.Debug;
    }
    Object.defineProperty = definePropertySave;
}
OSF.HostType = {
    Excel: "Excel",
    Outlook: "Outlook",
    Access: "Access",
    PowerPoint: "PowerPoint",
    Word: "Word",
    Sway: "Sway",
    OneNote: "OneNote"
};
OSF.ManifestHostType = {
    Workbook: "Workbook",
    Document: "Document",
    Presentation: "Presentation",
    Notebook: "Notebook"
};
OSF.HostPlatform = {
    Web: "Web"
};
OSF.HostSpecificFileVersionDefault = "16.00";
OSF.HostSpecificFileVersionMap = {
    "access": {
        "web": "16.00"
    },
    "agavito": {
        "winrt": "16.00"
    },
    "excel": {
        "ios": "16.00",
        "mac": "16.00",
        "web": "16.00",
        "win32": "16.01"
    },
    "onenote": {
        "web": "16.00",
        "win32": "16.00"
    },
    "outlook": {
        "ios": "16.00",
        "mac": "16.00",
        "web": "16.01",
        "win32": "16.00"
    },
    "powerpoint": {
        "ios": "16.00",
        "mac": "16.00",
        "web": "16.00",
        "win32": "16.01"
    },
    "project": {
        "win32": "16.00"
    },
    "sway": {
        "web": "16.00"
    },
    "word": {
        "ios": "16.00",
        "mac": "16.00",
        "web": "16.00",
        "win32": "16.01"
    }
};
OSF.getAppVerCode = function OSF$getAppVerCode(appName) {
    var appVerCode;
    switch (appName) {
        case OSF.AppName.ExcelWebApp:
            appVerCode = "excel.exe";
            break;
        case OSF.AppName.AccessWebApp:
            appVerCode = "ZAC";
            break;
        case OSF.AppName.PowerpointWebApp:
            appVerCode = "WAP";
            break;
        case OSF.AppName.WordWebApp:
            appVerCode = "WAW";
            break;
        case OSF.AppName.OneNoteWebApp:
            appVerCode = "WAO";
            break;
        default:
            OsfMsAjaxFactory.msAjaxDebug.trace("Invalid appName.");
            throw "Invalid appName.";
    }
    return appVerCode;
};
OSF.getManifestHostType = function OSF$getManifestHostType(hostType) {
    var manifestHostType;
    switch (hostType) {
        case OSF.HostType.Excel:
            manifestHostType = OSF.ManifestHostType.Workbook;
            break;
        case OSF.HostType.Word:
            manifestHostType = OSF.ManifestHostType.Document;
            break;
        case OSF.HostType.PowerPoint:
            manifestHostType = OSF.ManifestHostType.Presentation;
            break;
        case OSF.HostType.OneNote:
            manifestHostType = OSF.ManifestHostType.Notebook;
            break;
        default:
            OsfMsAjaxFactory.msAjaxDebug.trace("Invalid host type.");
            throw "Invalid hostType.";
    }
    return manifestHostType;
};
OSF.Capability = {
    Mailbox: "Mailbox",
    Document: "Document",
    Workbook: "Workbook",
    Project: "Project",
    Presentation: "Presentation",
    Database: "Database",
    Sway: "Sway",
    Notebook: "Notebook"
};
OSF.HostCapability = {
    Excel: OSF.Capability.Workbook,
    Outlook: OSF.Capability.Mailbox,
    Access: OSF.Capability.Database,
    PowerPoint: OSF.Capability.Presentation,
    Word: OSF.Capability.Document,
    Sway: OSF.Capability.Sway,
    OneNote: OSF.Capability.Notebook
};
OSF.OsfControlTarget = {
    InContent: 0,
    TaskPane: 1,
    Contextual: 2
};
OSF.OsfControlPermission = {
    Restricted: 1,
    ReadDocument: 2,
    WriteDocument: 4,
    ReadWriteDocument: 6,
    ReadItem: 32,
    ReadWriteMailbox: 64,
    ReadAllDocument: 131
};
OSF.OsfControlStatus = {
    NotActivated: 1,
    Activated: 2,
    AppStoreNotReachable: 3,
    InvalidOsfControl: 4,
    UnsupportedStore: 5,
    UnknownStore: 6,
    ActivationFailed: 7,
    NotSandBoxSupported: 8
};
OSF.OsfControlPageStatus = {
    NotStarted: 1,
    Loading: 2,
    Ready: 3,
    Failed: 4,
    FailedOriginCheck: 5,
    FailedPermissionCheck: 6
};
OSF.StoreType = {
    OMEX: "omex",
    SPCatalog: "spcatalog",
    SPApp: "spapp",
    FileSystem: "filesystem",
    Exchange: "exchange",
    OneDrive: "onedrive",
    Registry: "registry",
    InMemory: "inmemory",
    UploadFileDevCatalog: "uploadfiledevcatalog"
};
OSF.ManifestIdIssuer = {
    Microsoft: "Microsoft",
    Custom: "Custom"
};
OSF.OmexClientAppStatus = {
    OK: 1,
    UnknownAssetId: 2,
    KilledAsset: 3,
    NoEntitlement: 4,
    DownloadsExceeded: 5,
    Expired: 6,
    Invalid: 7,
    Revoked: 8,
    ServerError: 9,
    BadRequest: 10,
    LimitedTrial: 11,
    TrialNotSupported: 12,
    EntitlementDeactivated: 13,
    VersionMismatch: 14,
    VersionNotSupported: 15
};
OSF.OmexState = {
    Killed: 0,
    OK: 1,
    Withdrawn: 2,
    Flagged: 3,
    DeveloperWithdrawn: 4
};
OSF.OmexTrialType = {
    None: 0,
    Office: 1,
    External: 2
};
OSF.OmexEntitlementType = {
    Free: "free",
    Trial: "trial",
    Paid: "paid"
};
OSF.OmexAuthNStatus = {
    NotAttempted: -1,
    CheckFailed: 0,
    Authenticated: 1,
    Anonymous: 2,
    Unknown: 3
};
OSF.OmexRemoveAppStatus = {
    Failed: 0,
    Success: 1
};
OSF.OfficeAppType = {
    ContentApp: OSF.OsfControlTarget.InContent,
    TaskPaneApp: OSF.OsfControlTarget.TaskPane,
    MailApp: OSF.OsfControlTarget.Contextual
};
OSF.FormFactor = {
    Default: "DefaultSettings",
    Desktop: "DesktopSettings",
    Tablet: "TabletSettings",
    Phone: "PhoneSettings"
};
OSF.OsfOfficeExtensionManagerPerfMarker = {
    GetEntitlementStart: "Agave.OfficeExtensionManager.GetEntitlementStart",
    GetEntitlementEnd: "Agave.OfficeExtensionManager.GetEntitlementEnd"
};
OSF.OsfControlActivationPerfMarker = {
    ActivationStart: "Agave.AgaveActivationStart",
    ActivationEnd: "Agave.AgaveActivationEnd",
    DeactivationStart: "Agave.AgaveDeactivationStart",
    DeactivationEnd: "Agave.AgaveDeactivationEnd",
    SelectionTimeout: "Agave.AgaveSelectionTimeout"
};
OSF.NotificationUxPerfMarker = {
    RenderLoadingAnimationStart: "Agave.NotificationUx.RenderLoadingAnimationStart",
    RenderLoadingAnimationEnd: "Agave.NotificationUx.RenderLoadingAnimationEnd",
    RemoveLoadingAnimationStart: "Agave.NotificationUx.RemoveLoadingAnimationStart",
    RemoveLoadingAnimationEnd: "Agave.NotificationUx.RemoveLoadingAnimationEnd",
    RenderStage1Start: "Agave.NotificationUx.RenderStage1Start",
    RenderStage1End: "Agave.NotificationUx.RenderStage1End",
    RemoveStage1Start: "Agave.NotificationUx.RemoveStage1Start",
    RemoveStage1End: "Agave.NotificationUx.RemoveStage1End",
    RenderStage2Start: "Agave.NotificationUx.RenderStage2Start",
    RenderStage2End: "Agave.NotificationUx.RenderStage2End",
    RemoveStage2Start: "Agave.NotificationUx.RemoveStage2Start",
    RemoveStage2End: "Agave.NotificationUx.RemoveStage2End"
};
OSF.ProxyCallStatusCode = {
    Succeeded: 1,
    Failed: 0,
    ProxyNotReady: -1
};
OSF.Constants = {
    FileVersion: "16.0.6723.3002",
    ThreePartsFileVersion: "16.0.6723",
    OmexAnonymousServiceExtension: "/anonymousserviceextension.aspx",
    OmexGatedServiceExtension: "/gatedserviceextension.aspx",
    OmexUnGatedServiceExtension: "/ungatedserviceextension.aspx",
    Http: "http",
    Https: "https",
    ProtocolSeparator: "://",
    SignInRedirectUrl: "/logontoliveforwac.aspx?returnurl=",
    ETokenParameterName: "et",
    ActivatedCacheKey: "__OSF_RUNTIME_.Activated.{0}.{1}.{2}",
    AuthenticatedConnectMaxTries: 3,
    IgnoreSandBoxSupport: "Ignore_SandBox_Support",
    IEUpgradeUrl: "http://office.microsoft.com/redir/HA102789344.aspx",
    OmexForceAnonymousParamName: "SKAV",
    OmexForceAnonymousParamValue: "274AE4CD-E50B-4342-970E-1E7F36C70037",
    EndPointInternalSuffix: "_internal",
    PreloadOfficeJsId: "OFFICEJSPRELOAD",
    PreloadOfficeJsUrl: "//appsforoffice.microsoft.com/preloading/preloadoffice.js"
};
var _omexXmlNamespaces = 'xmlns="urn:schemas-microsoft-com:office:office" xmlns:o="urn:schemas-microsoft-com:office:office"';
OSF.AppVersion = {
    access: "ZAC150",
    excel: "ZXL150",
    excelwebapp: "WAE160",
    outlook: "ZOL150",
    outlookwebapp: "MOW150",
    powerpoint: "ZPP151",
    powerpointwebapp: "WAP160",
    project: "ZPJ150",
    word: "ZWD150",
    wordwebapp: "WAW160",
    onenotewebapp: "WAO160"
};
OSF.AppSubType = {
    Taskpane: 1,
    Content: 2,
    Contextual: 3,
    Dictionary: 4
};
OSF.ClientAppInfoReturnType = {
    urlOnly: 0,
    etokenOnly: 1,
    both: 2
};
function _getAppSubType(officeExtentionTarget) {
    var appSubType;
    if (officeExtentionTarget === 0) {
        appSubType = OSF.AppSubType.Content;
    }
    else if (officeExtentionTarget === 1) {
        appSubType = OSF.AppSubType.Taskpane;
    }
    else {
        throw OsfMsAjaxFactory.msAjaxError.argument("officeExtentionTarget");
    }
    return appSubType;
}
;
function _getAppVersion(applicationName) {
    var appVersion = OSF.AppVersion[applicationName.toLowerCase()];
    if (typeof appVersion == "undefined") {
        throw OsfMsAjaxFactory.msAjaxError.argument("applicationName");
    }
    return appVersion;
}
;
function _invokeCallbackTag(callback, status, result, errorMessage, executor, tagId) {
    var constructedMessage = errorMessage;
    if (callback) {
        try {
            var response = { "status": status, "result": result, "failureInfo": null };
            var setFailureInfoProperty = function _invokeCallbackTag$setFailureInfoProperty(response, name, value) {
                if (response.failureInfo === null) {
                    response.failureInfo = {};
                }
                response.failureInfo[name] = value;
            };
            if (executor) {
                var httpStatusCode = -1;
                if (executor.get_statusCode) {
                    httpStatusCode = executor.get_statusCode();
                }
                if (!constructedMessage) {
                    if (executor.get_timedOut && executor.get_timedOut()) {
                        constructedMessage = "Request timed out.";
                    }
                    else if (executor.get_aborted && executor.get_aborted()) {
                        constructedMessage = "Request aborted.";
                    }
                }
                if (httpStatusCode >= 400 || status === statusCode.Failed || constructedMessage) {
                    setFailureInfoProperty(response, "statusCode", httpStatusCode);
                    setFailureInfoProperty(response, "tagId", tagId);
                    var webRequest = executor.get_webRequest();
                    if (webRequest) {
                        if (webRequest.getResolvedUrl) {
                            setFailureInfoProperty(response, "url", webRequest.getResolvedUrl());
                        }
                        if (executor.getResponseHeader && webRequest.get_userContext && webRequest.get_userContext() && !webRequest.get_userContext().correlationId) {
                            var correlationId = executor.getResponseHeader("X-CorrelationId");
                            setFailureInfoProperty(response, "correlationId", correlationId);
                        }
                    }
                }
            }
            if (constructedMessage) {
                setFailureInfoProperty(response, "message", constructedMessage);
                OsfMsAjaxFactory.msAjaxDebug.trace(constructedMessage);
            }
        }
        catch (ex) {
            OsfMsAjaxFactory.msAjaxDebug.trace("Encountered exception with logging: " + ex);
        }
        callback(response);
    }
}
;
var _serviceEndPoint = null;
var _defaultRefreshRate = 3;
var _msPerDay = 86400000;
var _defaultTimeout = 60000;
var _officeVersionHeader = "X-Office-Version";
var _hourToDayConversionFactor = 24;
var _buildParameter = "&build=";
var _expectedVersionParameter = "&expver=";
var _queryStringParameters = {
    clientName: "client",
    clientVersion: "cv"
};
var statusCode = {
    Succeeded: 1,
    Failed: 0
};
function _sendWebRequest(url, verb, headers, onCompleted, context, body) {
    context = context || {};
    var webRequest = new Sys.Net.WebRequest();
    for (var p in headers) {
        webRequest.get_headers()[p] = headers[p];
    }
    if (context) {
        if (context.officeVersion) {
            webRequest.get_headers()[_officeVersionHeader] = context.officeVersion;
        }
        if (context.correlationId) {
            url += "&corr=" + context.correlationId;
        }
    }
    if (body) {
        webRequest.set_body(body);
    }
    webRequest.set_url(url);
    webRequest.set_httpVerb(verb);
    webRequest.set_timeout(_defaultTimeout);
    webRequest.set_userContext(context);
    webRequest.add_completed(onCompleted);
    webRequest.invoke();
}
;
function _onCompleted(executor, eventArgs) {
    var context = executor.get_webRequest().get_userContext();
    var url = executor.get_webRequest().get_url();
    if (executor.get_timedOut()) {
        OsfMsAjaxFactory.msAjaxDebug.trace("Request timed out: " + url);
        _invokeCallbackTag(context.callback, statusCode.Failed, null, null, executor, 0x0085a2c3);
    }
    else if (executor.get_aborted()) {
        OsfMsAjaxFactory.msAjaxDebug.trace("Request aborted: " + url);
        _invokeCallbackTag(context.callback, statusCode.Failed, null, null, executor, 0x0085a2c4);
    }
    else if (executor.get_responseAvailable()) {
        if (executor.get_statusCode() == 200) {
            try {
                context._onCompleteHandler(executor, eventArgs);
            }
            catch (ex) {
                OsfMsAjaxFactory.msAjaxDebug.trace("Request failed with exception " + ex + ": " + url);
                _invokeCallbackTag(context.callback, statusCode.Failed, ex, null, executor, 0x0085a2c5);
            }
        }
        else {
            var statusText = executor.get_statusText();
            OsfMsAjaxFactory.msAjaxDebug.trace("Request failed with status code " + statusText + ": " + url);
            _invokeCallbackTag(context.callback, statusCode.Failed, statusText, null, executor, 0x0085a2c6);
        }
    }
    else {
        OsfMsAjaxFactory.msAjaxDebug.trace("Request failed: " + url);
        _invokeCallbackTag(context.callback, statusCode.Failed, statusText, null, executor, 0x0085a2c7);
    }
}
;
function _isProxyReady(params, callback) {
    if (callback) {
        callback({ "status": statusCode.Succeeded, "result": true });
    }
}
;
function _createQueryStringFragment(paramDictionary) {
    var queryString = "";
    for (var param in paramDictionary) {
        var value = paramDictionary[param];
        if (value === null || value === undefined || value === "") {
            continue;
        }
        queryString += '&' + encodeURIComponent(param) + '=' + encodeURIComponent(value);
    }
    return queryString;
}
;
OSF.SQMDataPoints = {
    DATAID_APPSFOROFFICEUSAGE: 10595,
    DATAID_APPSFOROFFICENOTIFICATIONS: 10942
};
OSF.BWsaStreamTypes = {
    Static: 1
};
OSF.BWsaConfig = {
    defaultMaxStreamRows: 1000
};
OSF.ErrorStatusCodes = {
    E_OEM_EXTENSION_NOT_ENTITLED: 2147758662,
    E_MANIFEST_SERVER_UNAVAILABLE: 2147758672,
    E_USER_NOT_SIGNED_IN: 2147758677,
    E_MANIFEST_DOES_NOT_EXIST: 2147758673,
    E_OEM_EXTENSION_KILLED: 2147758660,
    E_OEM_OMEX_EXTENSION_KILLED: 2147758686,
    E_MANIFEST_UPDATE_AVAILABLE: 2147758681,
    S_OEM_EXTENSION_TRIAL_MODE: 275013,
    E_OEM_EXTENSION_WITHDRAWN_FROM_SALE: 2147758690,
    E_TOKEN_EXPIRED: 2147758675,
    E_TRUSTCENTER_CATALOG_UNTRUSTED_ADMIN_CONTROLLED: 2147757992,
    E_MANIFEST_REFERENCE_INVALID: 2147758678,
    S_USER_CLICKED_BUY: 274205,
    E_MANIFEST_INVALID_VALUE_FORMAT: 2147758654,
    E_TRUSTCENTER_MOE_UNACTIVATED: 2147757996,
    S_OEM_EXTENSION_FLAGGED: 275040,
    S_OEM_EXTENSION_DEVELOPER_WITHDRAWN_FROM_SALE: 275041,
    E_BROWSER_VERSION: 2147758194,
    WAC_AgaveUnsupportedStoreType: 1041,
    WAC_AgaveActivationError: 1042,
    WAC_ActivateAttempLoading: 1043,
    WAC_HTML5IframeSandboxNotSupport: 1044,
    WAC_AgaveRequirementsErrorOmex: 1045,
    WAC_AgaveRequirementsError: 1046,
    WAC_AgaveOriginCheckError: 1047,
    WAC_AgavePermissionCheckError: 1048,
    WAC_AgaveHostHandleRequestError: 1049,
    WAC_AgaveUnknownClientAppStatus: 1050,
    WAC_AgaveAnonymousProxyCreationError: 1051,
    WAC_AgaveOsfControlActivationError: 1052
};
OSF.InvokeResultCode = {
    "S_OK": 0,
    "E_REQUEST_TIME_OUT": -2147471590,
    "E_USER_NOT_SIGNED_IN": -2147208619,
    "E_CATALOG_ACCESS_DENIED": -2147471591,
    "E_CATALOG_REQUEST_FAILED": -2147471589,
    "E_OEM_NO_NETWORK_CONNECTION": -2147208640,
    "E_PROVIDER_NOT_REGISTERED": -2147208617,
    "E_OEM_CACHE_SHUTDOWN": -2147208637,
    "E_OEM_REMOVED_FAILED": -2147209421,
    "E_CATALOG_NO_APPS": -1,
    "E_GENERIC_ERROR": -1000,
    "S_HIDE_PROVIDER": 10
};
OSF.OmexClientNames = (function OSF_OmexClientNames() {
    var nameMap = {}, appNameList = OSF.AppName;
    nameMap[appNameList.ExcelWebApp] = "WAC_Excel";
    nameMap[appNameList.WordWebApp] = "WAC_Word";
    nameMap[appNameList.OutlookWebApp] = "WAC_Outlook";
    nameMap[appNameList.AccessWebApp] = "WAC_Access";
    nameMap[appNameList.PowerpointWebApp] = "WAC_Powerpoint";
    nameMap[appNameList.OneNoteWebApp] = "WAC_OneNote";
    return nameMap;
})();
OSF.OmexAppVersions = (function OSF_OmexAppVersions() {
    var nameMap = {}, appNameList = OSF.AppName;
    nameMap[appNameList.ExcelWebApp] = OSF.AppVersion.excelwebapp;
    nameMap[appNameList.WordWebApp] = OSF.AppVersion.wordwebapp;
    nameMap[appNameList.OutlookWebApp] = OSF.AppVersion.outlookwebapp;
    nameMap[appNameList.AccessWebApp] = OSF.AppVersion.access;
    nameMap[appNameList.PowerpointWebApp] = OSF.AppVersion.powerpointwebapp;
    nameMap[appNameList.OneNoteWebApp] = OSF.AppVersion.OneNoteWebApp;
    return nameMap;
})();
OSF.ManifestSchemaVersion = {
    "1.0": "1.0",
    "1.1": "1.1"
};
OSF.ManifestNamespaces = {
    "1.0": 'xmlns="http://schemas.microsoft.com/office/appforoffice/1.0" xmlns:o="http://schemas.microsoft.com/office/appforoffice/1.0" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"',
    "1.1": 'xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:o="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"'
};
OSF.OsfNavigationMode = {
    DefaultMode: 0,
    CategoryMode: 1,
    TrustPageMode: 2,
    QueryResultMode: 3
};
OSF.StoreTypeEnum = {
    MarketPlace: 0,
    Catalog: 1,
    OneDrive: 9
};
OSF.RequirementsChecker = function OSF_RequirementsChecker(supportedCapabilities, supportedHosts, supportedRequirements, supportedControlTargets, supportedOmexAppVersions) {
    this.setCapabilities(supportedCapabilities);
    this.setHosts(supportedHosts);
    this.setRequirements(supportedRequirements);
    this.setSupportedControlTargets(supportedControlTargets);
    this.setSupportedOmexAppVersions(supportedOmexAppVersions);
    this.setFilteringEnabled(false);
};
OSF.RequirementsChecker.prototype = {
    defaultMinMaxVersion: "1.1",
    isManifestSupported: function OSF_RequirementsChecker$isManifestSupported(manifest) {
        if (!this.isFilteringEnabled()) {
            return true;
        }
        if (!manifest) {
            return false;
        }
        var manifestSchemaVersion = manifest.getManifestSchemaVersion() || OSF.ManifestSchemaVersion["1.0"];
        switch (manifestSchemaVersion) {
            case OSF.ManifestSchemaVersion["1.0"]:
                return this._checkManifest1_0(manifest);
            case OSF.ManifestSchemaVersion["1.1"]:
                return this._checkManifest1_1(manifest);
            default:
                return false;
        }
    },
    isEntitlementFromOmexSupported: function OSF_RequirementsChecker$isEntitlementFromOmexSupported(entitlement) {
        if (!this.isFilteringEnabled()) {
            return true;
        }
        if (!entitlement) {
            return false;
        }
        var targetType;
        switch (entitlement.appSubType) {
            case "1":
                targetType = OSF.OsfControlTarget.TaskPane;
                break;
            case "2":
                targetType = OSF.OsfControlTarget.InContent;
                break;
            case "3":
                targetType = OSF.OsfControlTarget.Contextual;
                break;
            case "4":
                targetType = OSF.OsfControlTarget.TaskPane;
                break;
            default:
                return false;
        }
        if (!this._checkControlTarget(targetType)) {
            return false;
        }
        if (!entitlement.requirements && !entitlement.hosts) {
            if (!entitlement.hasOwnProperty("appVersions")) {
                return true;
            }
            return this._checkOmexAppVersions(entitlement.appVersions);
        }
        var pseudoParser = new OSF.Manifest.Manifest(function () { });
        var requirements, requirementsNode, hosts, hostsNode;
        if (entitlement.requirements) {
            pseudoParser._xmlProcessor = new OSF.XmlProcessor(entitlement.requirements, OSF.ManifestNamespaces["1.1"]);
            requirementsNode = pseudoParser._xmlProcessor.getDocumentElement();
        }
        requirements = pseudoParser._parseRequirements(requirementsNode);
        if (entitlement.hosts) {
            pseudoParser._xmlProcessor = new OSF.XmlProcessor(entitlement.hosts, OSF.ManifestNamespaces["1.1"]);
            hostsNode = pseudoParser._xmlProcessor.getDocumentElement();
        }
        hosts = pseudoParser._parseHosts(hostsNode);
        return this._checkHosts(hosts) &&
            this._checkSets(requirements.sets) &&
            this._checkMethods(requirements.methods);
    },
    isEntitlementFromCorpCatalogSupported: function OSF_RequirementsChecker$isEntitlementFromCorpCatalogSupported(entitlement) {
        if (!this.isFilteringEnabled()) {
            return true;
        }
        if (!entitlement) {
            return false;
        }
        var targetType = OSF.OfficeAppType[entitlement.OEType];
        if (!this._checkControlTarget(targetType)) {
            return false;
        }
        var pseudoParser = new OSF.Manifest.Manifest(function () {
        });
        var hosts, sets, methods;
        if (entitlement.OfficeExtensionCapabilitiesXML) {
            pseudoParser._xmlProcessor = new OSF.XmlProcessor(entitlement.OfficeExtensionCapabilitiesXML, OSF.ManifestNamespaces["1.1"]);
            var xmlNode, requirements;
            xmlNode = pseudoParser._xmlProcessor.getDocumentElement();
            requirements = pseudoParser._parseRequirements(xmlNode);
            sets = requirements.sets;
            methods = requirements.methods;
            hosts = pseudoParser._parseHosts(xmlNode);
        }
        return this._checkHosts(hosts) && this._checkSets(sets) && this._checkMethods(methods);
    },
    setCapabilities: function OSF_RequirementsChecker$setCapabilities(capabilities) {
        this._supportedCapabilities = this._scalarArrayToObject(capabilities);
    },
    setHosts: function OSF_RequirementsChecker$setHosts(hosts) {
        this._supportedHosts = this._scalarArrayToObject(hosts);
    },
    setRequirements: function OSF_RequirementsChecker$setRequirements(requirements) {
        this._supportedSets = requirements && this._arrayToSetsObject(requirements.sets) || {};
        this._supportedMethods = requirements && this._scalarArrayToObject(requirements.methods) || {};
    },
    setSupportedControlTargets: function OSF_RequirementsChecker$setSupportedControlTargets(controlTargets) {
        this._supportedControlTargets = this._scalarArrayToObject(controlTargets);
    },
    setSupportedOmexAppVersions: function OSF_RequirementsChecker$setSupportedOmexAppVersions(appVersions) {
        this._supportedOmexAppVersions = appVersions && appVersions.slice ? appVersions.slice(0) : [];
    },
    setFilteringEnabled: function OSF_RequirementsChecker$setFilteringEnabled(filteringEnabled) {
        this._filteringEnabled = filteringEnabled ? true : false;
    },
    isFilteringEnabled: function OSF_RequirementsChecker$isFilteringEnabled() {
        return this._filteringEnabled;
    },
    _checkManifest1_0: function OSF_RequirementsChecker$_checkManifest1_0(manifest) {
        return this._checkCapabilities(manifest.getCapabilities());
    },
    _checkCapabilities: function OSF_RequirementsChecker$_checkCapabilities(askedCapabilities) {
        if (!askedCapabilities || askedCapabilities.length === 0) {
            return true;
        }
        for (var i = 0; i < askedCapabilities.length; i++) {
            if (this._supportedCapabilities[askedCapabilities[i]]) {
                return true;
            }
        }
        return false;
    },
    _checkManifest1_1: function OSF_RequirementsChecker$_checkManifest1_1(manifest) {
        var askedRequirements = manifest.getRequirements() || {};
        return this._checkHosts(manifest.getHosts()) &&
            this._checkSets(askedRequirements.sets) &&
            this._checkMethods(askedRequirements.methods);
    },
    _checkHosts: function OSF_RequirementsChecker$_checkHosts(askedHosts) {
        if (!askedHosts || askedHosts.length === 0) {
            return true;
        }
        for (var i = 0; i < askedHosts.length; i++) {
            if (this._supportedHosts[askedHosts[i]]) {
                return true;
            }
        }
        return false;
    },
    _checkSets: function OSF_RequirementsChecker$_checkSets(askedSets) {
        if (!askedSets || askedSets.length === 0) {
            return true;
        }
        for (var i = 0; i < askedSets.length; i++) {
            var askedSet = askedSets[i];
            var supportedSet = this._supportedSets[askedSet.name];
            if (!supportedSet) {
                return false;
            }
            if (askedSet.version) {
                if (this._compareVersionStrings(supportedSet.minVersion || this.defaultMinMaxVersion, askedSet.version) > 0 ||
                    this._compareVersionStrings(supportedSet.maxVersion || this.defaultMinMaxVersion, askedSet.version) < 0) {
                    return false;
                }
            }
        }
        return true;
    },
    _checkMethods: function OSF_RequirementsChecker$_checkMethods(askedMethods) {
        if (!askedMethods || askedMethods.length === 0) {
            return true;
        }
        for (var i = 0; i < askedMethods.length; i++) {
            if (!this._supportedMethods[askedMethods[i]]) {
                return false;
            }
        }
        return true;
    },
    _checkControlTarget: function OSF_RequirementsChecker$_checkControlTarget(askedControlTarget) {
        return askedControlTarget != undefined && this._supportedControlTargets[askedControlTarget];
    },
    _checkOmexAppVersions: function OSF_RequirementsChecker$_checkOmexAppVersions(askedAppVersions) {
        if (!askedAppVersions) {
            return false;
        }
        for (var i = 0; i < this._supportedOmexAppVersions.length; i++) {
            if (askedAppVersions.indexOf(this._supportedOmexAppVersions[i]) >= 0) {
                return true;
            }
        }
        return false;
    },
    _scalarArrayToObject: function OSF_RequirementsChecker$_scalarArrayToObject(array) {
        var obj = {};
        if (array && array.length) {
            for (var i = 0; i < array.length; i++) {
                if (array[i] != undefined) {
                    obj[array[i]] = true;
                }
            }
        }
        return obj;
    },
    _arrayToSetsObject: function OSF_RequirementsChecker$_arrayToSetsObject(array) {
        var obj = {};
        if (array && array.length) {
            for (var i = 0; i < array.length; i++) {
                var set = array[i];
                if (set && set.name != undefined) {
                    obj[set.name] = set;
                }
            }
        }
        return obj;
    },
    _getSupportedSet: function OSF_RequirementsChecker$_getSupportedSet() {
        var obj = {};
        var supportSetNames = Object.getOwnPropertyNames(this._supportedSets);
        for (var i = 0; i < supportSetNames.length; i++) {
            var supportedSet = this._supportedSets[supportSetNames[i]];
            obj[supportedSet.name.toLowerCase()] = supportedSet.maxVersion || this.defaultMinMaxVersion;
        }
        if (typeof (JSON) !== "undefined") {
            return JSON.stringify(obj);
        }
        return obj;
    },
    _compareVersionStrings: function OSF_RequirementsChecker$_compareVersionStrings(leftVersion, rightVersion) {
        leftVersion = leftVersion.split('.');
        rightVersion = rightVersion.split('.');
        var maxComponentCount = Math.max(leftVersion.length, rightVersion.length);
        for (var i = 0; i < maxComponentCount; i++) {
            var leftInt = parseInt(leftVersion[i], 10) || 0, rightInt = parseInt(rightVersion[i], 10) || 0;
            if (leftInt === rightInt) {
                continue;
            }
            return leftInt - rightInt;
        }
        return 0;
    }
};
var _omexDataProvider = OmexDataProvider.GetInstance(new OfficeExt.AppsDataCacheManager(OSF.OUtil.getLocalStorage(), new OfficeExt.SafeSerializer()));
OSF.OsfControl = function OSF_OsfControl(params) {
    OSF.OUtil.validateParamObject(params, {
        "div": { type: Object, mayBeNull: false },
        "contextActivationMgr": { type: Object, mayBeNull: false },
        "id": { type: String, mayBeNull: false },
        "marketplaceID": { type: String, mayBeNull: false },
        "marketplaceVersion": { type: String, mayBeNull: false },
        "store": { type: String, mayBeNull: false },
        "storeType": { type: String, mayBeNull: false },
        "alternateReference": { type: Object, mayBeNull: true },
        "settings": { type: Object, mayBeNull: true },
        "reason": { type: String, mayBeNull: true },
        "osfControlType": { type: Number, mayBeNull: true },
        "snapshotUrl": { type: String, mayBeNull: true },
        "preactivationCallback": { type: Object, mayBeNull: true },
        "virtualOsfControlActivationCallback": { type: Object, mayBeNull: true },
        "isvirtualOsfControl": { type: Boolean, mayBeNull: true }
    }, null);
    this._div = params.div;
    this._contextActivationMgr = params.contextActivationMgr;
    this._id = params.id;
    this._storeType = params.storeType.toLowerCase();
    this._storeLocator = params.store;
    this._marketplaceID = params.marketplaceID;
    this._marketplaceVersion = params.marketplaceVersion;
    this._alternateReference = params.alternateReference;
    this._settings = params.settings || {};
    this._reason = params.reason == undefined ? Microsoft.Office.WebExtension.InitializationReason.DocumentOpened : params.reason;
    this._osfControlType = params.osfControlType == undefined ? OSF.OsfControlType.DocumentLevel : params.osfControlType;
    this._snapshotUrl = params.snapshotUrl;
    this._status = OSF.OsfControlStatus.NotActivated;
    this._errorcode = 0;
    this._notificationParams = null;
    this._iframeUrl = null;
    this._permission = null;
    this._conversationId = null;
    this._manifestUrl = null;
    this._pageStatus = OSF.OsfControlPageStatus.NotStarted;
    this._pageIsReadyTimerExpired = false;
    this._timer = null;
    this._retryLoadingNum = 2;
    this._frame = null;
    this._agaveEndPoint = null;
    this._etoken = "";
    this._sqmDWords = [0, 0];
    this._preactivationCallback = params.preactivationCallback;
    this._telemetryContext = {};
    this._controlFocus = false;
    this._dataProvider = OfficeExt.AddinDataProviderFactory.createAddinDataProvider(this._storeType, this._contextActivationMgr);
    this._manifest = null;
    this._virtualOsfControlActivationCallback = params.virtualOsfControlActivationCallback;
    this._isvirtualOsfControl = params.isvirtualOsfControl == undefined ? false : params.isvirtualOsfControl;
    this._isvirtualOsfControlCallbackInvoked = false;
    if (OSF.OUtil.isiOS()) {
        this._div.style.webkitOverflowScrolling = "touch";
        this._div.style.overflow = "auto";
    }
    this._appCorrelationId = OSF.OUtil.Guid.generateNewGuid();
    this._iframeOnLoadDelegate = Function.createDelegate(this, this._iframeOnLoad);
    this._retryActivate = null;
};
OSF.OsfControl.prototype = {
    activate: function OSF_OsfControl$activate(context) {
        try {
            Telemetry.AppLoadTimeHelper.ActivationStart(this._telemetryContext, this._sqmDWords[0], this._sqmDWords[1], this.getCorrelationId(), this._id);
            this._controlFocus = false;
            this._pageStatus = OSF.OsfControlPageStatus.Loading;
            OSF.OUtil.writeProfilerMark(OSF.OsfControlActivationPerfMarker.ActivationStart);
            if (this._status !== OSF.OsfControlStatus.Activated) {
                if (this._frame) {
                    OSF.OUtil.removeEventListener(this._frame, "load", this._iframeOnLoadDelegate);
                    this._frame = null;
                }
                while (this._div.childNodes.length > 0) {
                    this._div.removeChild(this._div.childNodes.item(0));
                }
                this._contextActivationMgr._ErrorUXHelper.showProgress(this._div, this._id);
                this._contextActivationMgr.registerOsfControl(this);
                context = context || {};
                context.hostType = this._contextActivationMgr._hostType;
                context.osfControl = context.osfControl || this;
                context.referenceInUse = context.referenceInUse || { id: this._marketplaceID, version: this._marketplaceVersion, storeType: this._storeType, storeLocator: this._storeLocator };
                context.correlationId = context.osfControl._appCorrelationId;
                if (!this._doesBrowserSupportRequiredFeatures()) {
                    this._showActivationError(OSF.OsfControlStatus.ActivationFailed, Strings.OsfRuntime.L_AgaveUnsupportedBroswer_ERR, Strings.OsfRuntime.L_RetryButton_TXT, Function.createDelegate(this, this._refresh), null, null, null, true, OSF.ErrorStatusCodes.E_BROWSER_VERSION);
                    this.invokePreactivationCompletedCallback();
                    return;
                }
                var frame = document.createElement("iframe");
                var sandboxSupported = "sandbox" in frame;
                frame = null;
                var ignoreSandBox = false;
                var osfLocalStorage = OSF.OUtil.getLocalStorage();
                if (osfLocalStorage) {
                    ignoreSandBox = osfLocalStorage.getItem(OSF.Constants.IgnoreSandBoxSupport);
                }
                if (this._contextActivationMgr._autoTrusted && !sandboxSupported && !ignoreSandBox) {
                    this._contextActivationMgr._ErrorUXHelper.removeProgressDiv(this._div, this._id);
                    this._status = OSF.OsfControlStatus.NotSandBoxSupported;
                    this._errorcode = OSF.ErrorStatusCodes.WAC_HTML5IframeSandboxNotSupport;
                    this._contextActivationMgr.displayNotification({
                        "id": this._id,
                        "infoType": OSF.InfoType.Warning,
                        "title": Strings.OsfRuntime.L_AppsDisabled_WRN,
                        "description": Strings.OsfRuntime.L_NotSandBoxSupported_ERR,
                        "buttonTxt": Strings.OsfRuntime.L_EnableAppsButton_TXT,
                        "buttonCallback": Function.createCallback(this._activateAgavesBlockedBySandboxNotSupport, this._contextActivationMgr),
                        "url": OSF.Constants.IEUpgradeUrl,
                        "urlButtonTxt": Strings.OsfRuntime.L_UpgradeBrowserButton_TXT,
                        "dismissCallback": null,
                        "reDisplay": true,
                        "displayDeactive": true,
                        "errorCode": OSF.ErrorStatusCodes.WAC_HTML5IframeSandboxNotSupport,
                        "highPriority": true
                    });
                    this.invokePreactivationCompletedCallback();
                    return;
                }
                var me = this;
                Telemetry.AppLoadTimeHelper.ServerCallStart(this._telemetryContext);
                var reference = context.referenceInUse;
                if (reference.storeType === OSF.StoreType.UploadFileDevCatalog) {
                    this._dataProvider.getManifestAsync({ "referenceInUse": context.referenceInUse, "contextActivationMgr": this._contextActivationMgr }, Function.createDelegate(this, this._onGetManifestCompleted));
                }
                else if (reference.storeType === OSF.StoreType.Exchange || reference.storeType === OSF.StoreType.InMemory) {
                    OSF.OsfManifestManager.getManifestAsync(context, Function.createDelegate(this, this._onGetManifestCompleted));
                }
                else if (reference.storeType === OSF.StoreType.Registry && this._contextActivationMgr._enableDevCatalog) {
                    if (context.osfControl.getReason() == Microsoft.Office.WebExtension.InitializationReason.DocumentOpened) {
                        var procManifestFile = function OSF_OsfControl_activate$procManifestFile(manifestString) {
                            var parsedManifest = new OSF.Manifest.Manifest(manifestString, me._contextActivationMgr.getAppUILocale());
                            if (!OSF.OsfManifestManager.hasManifest(parsedManifest.getMarketplaceID(), parsedManifest.getMarketplaceVersion())) {
                                OSF.OsfManifestManager.cacheManifest(parsedManifest.getMarketplaceID(), parsedManifest.getMarketplaceVersion(), parsedManifest);
                            }
                            OSF.OsfManifestManager.getManifestAsync(context, Function.createDelegate(me, me._onGetManifestCompleted));
                        };
                        var onGetManifestError = function OSF_OsfControl_activate$onGetManifestError(errorString) {
                            alert("Error when requsting manifest file: " + errorString);
                        };
                        OSF.OUtil.xhrGet(this._contextActivationMgr._devCatalogUrl + "/" + reference.id + ".xml", procManifestFile, onGetManifestError);
                    }
                    else {
                        OSF.OsfManifestManager.getManifestAsync(context, Function.createDelegate(this, this._onGetManifestCompleted));
                    }
                }
                else if (reference.storeType === OSF.StoreType.OneDrive && this._contextActivationMgr._enableOneDriveCatalog) {
                    var processManifestFileOneDrive = function (manifestString) {
                        var parsedManifest = new OSF.Manifest.Manifest(manifestString, me._contextActivationMgr.getAppUILocale());
                        OSF.OsfManifestManager.cacheManifest(reference.id, parsedManifest.getMarketplaceVersion(), parsedManifest);
                        OSF.OsfManifestManager.getManifestAsync(context, Function.createDelegate(me, me._onGetManifestCompleted));
                    };
                    var onGetManifestErrorOneDrive = function (errorString) {
                        me._showActivationError(OSF.OsfControlStatus.ActivationFailed, Strings.OsfRuntime.L_AgaveServerConnectionFailed_ERR, Strings.OsfRuntime.L_RetryButton_TXT, Function.createDelegate(me, me._refresh), null, null, null, null, OSF.ErrorStatusCodes.E_MANIFEST_SERVER_UNAVAILABLE);
                    };
                    var hostCallBackUri = window.location.protocol + "//" + window.location.host;
                    OSF.OneDriveOAuth.setHostCallbackUri(hostCallBackUri);
                    var manifestFullUrl = this._contextActivationMgr._oneDriveCatalogBaseApiUrl + "/root:/Manifests/" + reference.id + ":/content?access_token=";
                    var onAccessTokenSuccess = function (accessToken) {
                        manifestFullUrl = manifestFullUrl + accessToken;
                        OSF.OUtil.xhrGet(manifestFullUrl, processManifestFileOneDrive, onGetManifestErrorOneDrive);
                    };
                    var onAccessTokenError = function () {
                        me._showActivationError(OSF.OsfControlStatus.ActivationFailed, Strings.OsfRuntime.L_AgaveServerConnectionFailed_ERR, Strings.OsfRuntime.L_RetryButton_TXT, Function.createDelegate(me, me._refresh), null, null, null, null, OSF.ErrorStatusCodes.E_MANIFEST_SERVER_UNAVAILABLE);
                    };
                    OSF.OneDriveOAuth.getAccessToken(onAccessTokenSuccess, onAccessTokenError);
                }
                else if (reference.storeType === OSF.StoreType.SPCatalog) {
                    if (this._contextActivationMgr._doesUrlHaveSupportedProtocol(reference.storeLocator)) {
                        Telemetry.AppLoadTimeHelper.EntitlementCheckStart(this._telemetryContext);
                        context.webUrl = reference.storeLocator;
                        OSF.OUtil.writeProfilerMark(OSF.OsfOfficeExtensionManagerPerfMarker.GetEntitlementStart);
                        OSF.OsfManifestManager.getCorporateCatalogEntitlementsAsync(context, Function.createDelegate(this, this._onGetEntitlementsCompleted));
                    }
                    else {
                        this._showActivationError(OSF.OsfControlStatus.ActivationFailed, Strings.OsfRuntime.L_AgaveUnknownStoreType_ERR, null, null, null, null, null, null, OSF.ErrorStatusCodes.E_MANIFEST_REFERENCE_INVALID);
                    }
                }
                else if (reference.storeType === OSF.StoreType.SPApp) {
                    Telemetry.AppLoadTimeHelper.EntitlementCheckStart(this._telemetryContext);
                    context.baseUrl = this._contextActivationMgr.getPageBaseUrl();
                    context.pageUrl = this._contextActivationMgr._docUrl;
                    context.webUrl = this._contextActivationMgr._webUrl;
                    context.appWebUrl = this._contextActivationMgr._webUrl;
                    OSF.OsfManifestManager.getSPAppEntitlementsAsync(context, Function.createDelegate(this, this._onGetEntitlementsCompleted));
                }
                else if (reference.storeType === OSF.StoreType.OMEX) {
                    if (this._contextActivationMgr.isExternalMarketplaceAllowed() && this._contextActivationMgr._doesUrlHaveSupportedProtocol(this._contextActivationMgr._osfOmexBaseUrl)) {
                        if (!this._omexEntitlement) {
                            this._omexEntitlement = {
                                "contentMarket": reference.storeLocator,
                                "version": reference.version,
                                "assetId": reference.id,
                                "etoken": "",
                                "hasEntitlement": false
                            };
                            this._omexEntitlement.endPointUrl = this._contextActivationMgr._getOmexEndPointPageUrl(reference.id, reference.storeLocator);
                        }
                        context.clientVersion = this._contextActivationMgr._getClientVersionForOmex();
                        if (context.clientVersion) {
                            context.clientName = this._contextActivationMgr._getClientNameForOmex();
                            context.appVersion = this._contextActivationMgr._getAppVersionForOmex();
                        }
                        var onGetAuthNStatusCompleted = function (asyncResult) {
                            if (asyncResult.statusCode === OSF.ProxyCallStatusCode.Succeeded && asyncResult.value) {
                                var authNStatus = parseInt(asyncResult.value);
                                if (authNStatus == OSF.OmexAuthNStatus.Authenticated) {
                                    me._contextActivationMgr._omexAuthNStatus = OSF.OmexAuthNStatus.Authenticated;
                                    me._contextActivationMgr._createOmexProxy(me._contextActivationMgr._omexGatedWSProxy, createAuthenticatedOmexProxyCompleted);
                                }
                                else if (authNStatus == OSF.OmexAuthNStatus.Anonymous || authNStatus == OSF.OmexAuthNStatus.Unknown) {
                                    me._contextActivationMgr._omexAuthNStatus = OSF.OmexAuthNStatus.Anonymous;
                                    me._contextActivationMgr._createOmexProxy(me._contextActivationMgr._omexAnonymousWSProxy, createAnonymousOmexProxyCompleted);
                                }
                            }
                            else {
                                Telemetry.RuntimeTelemetryHelper.LogExceptionTag("OMEX_getAuthNStatus failed:" + OSF.OUtil.serializeOMEXResponseErrorMessage(asyncResult), null, null, 0x010c5398);
                                me._contextActivationMgr._omexAuthNStatus = OSF.OmexAuthNStatus.CheckFailed;
                                if (me._contextActivationMgr._omexForceAnonymous) {
                                    me._contextActivationMgr._createOmexProxy(me._contextActivationMgr._omexAnonymousWSProxy, createAnonymousOmexProxyCompleted);
                                }
                                else {
                                    me._contextActivationMgr._createOmexProxy(me._contextActivationMgr._omexGatedWSProxy, createAuthenticatedOmexProxyCompleted);
                                }
                            }
                        };
                        var omexAuthenticatedConnectTries = 1;
                        var createAnonymousOmexProxyCompleted = function (clientEndPoint) {
                            if (clientEndPoint) {
                                if (me._contextActivationMgr._omexAuthNStatus == OSF.OmexAuthNStatus.NotAttempted) {
                                    var params = { "clientEndPoint": null };
                                    params.clientEndPoint = clientEndPoint;
                                    OSF.OsfManifestManager._invokeProxyMethodAsync(context, "OMEX_getAuthNStatus", onGetAuthNStatusCompleted, params);
                                }
                                else {
                                    Telemetry.AppLoadTimeHelper.AuthenticationEnd(me._telemetryContext);
                                    Telemetry.AppLoadTimeHelper.KilledAppsCheckStart(me._telemetryContext);
                                    if (me._contextActivationMgr._omexAuthNStatus == OSF.OmexAuthNStatus.CheckFailed) {
                                        me._contextActivationMgr._omexForceAnonymous = true;
                                    }
                                    context.anonymous = true;
                                    Telemetry.AppLoadTimeHelper.SetAnonymousFlag(me._telemetryContext, context.anonymous);
                                    context.clientEndPoint = clientEndPoint;
                                    OSF.OsfManifestManager.getOmexKilledAppsAsync(context, Function.createDelegate(me, me._onGetOmexKilledAppsCompleted));
                                    Telemetry.AppLoadTimeHelper.AppStateCheckStart(me._telemetryContext);
                                    OSF.OsfManifestManager.getOmexAppStateAsync(context, Function.createDelegate(me, me._onGetOmexAppStateCompleted));
                                }
                            }
                            else {
                                Telemetry.RuntimeTelemetryHelper.LogExceptionTag("ClientEndPoint is null after creating anonymouseOmexProxy.", null, null, 0x0108d740);
                                me._showActivationError(OSF.OsfControlStatus.ActivationFailed, Strings.OsfRuntime.L_AgaveServerConnectionFailed_ERR, Strings.OsfRuntime.L_RetryButton_TXT, Function.createDelegate(me, me._refresh), null, null, null, null, OSF.ErrorStatusCodes.WAC_AgaveAnonymousProxyCreationError);
                            }
                        };
                        Telemetry.AppLoadTimeHelper.AuthenticationStart(this._telemetryContext);
                        var createAuthenticatedOmexProxyCompleted = function (clientEndPoint) {
                            Telemetry.AppLoadTimeHelper.SetRetryCount(me._telemetryContext, omexAuthenticatedConnectTries);
                            if (clientEndPoint) {
                                Telemetry.AppLoadTimeHelper.AuthenticationEnd(me._telemetryContext);
                                Telemetry.AppLoadTimeHelper.EntitlementCheckStart(me._telemetryContext);
                                context.anonymous = false;
                                Telemetry.AppLoadTimeHelper.SetAnonymousFlag(me._telemetryContext, context.anonymous);
                                context.clientEndPoint = clientEndPoint;
                                OSF.OsfManifestManager.getOmexEntitlementsAsync(context, Function.createDelegate(me, me._onGetOmexEntitlementsCompleted));
                                Telemetry.AppLoadTimeHelper.KilledAppsCheckStart(me._telemetryContext);
                                OSF.OsfManifestManager.getOmexKilledAppsAsync(context, Function.createDelegate(me, me._onGetOmexKilledAppsCompleted));
                            }
                            else {
                                if (omexAuthenticatedConnectTries < OSF.Constants.AuthenticatedConnectMaxTries) {
                                    omexAuthenticatedConnectTries++;
                                    if (me._contextActivationMgr._omexAuthNStatus == OSF.OmexAuthNStatus.CheckFailed) {
                                        me._contextActivationMgr._createOmexProxy(me._contextActivationMgr._omexGatedWSProxy, createAuthenticatedOmexProxyCompleted);
                                    }
                                    else {
                                        me._contextActivationMgr._createOmexProxy(me._contextActivationMgr._omexAnonymousWSProxy, createAnonymousOmexProxyCompleted);
                                    }
                                }
                                else {
                                    me._contextActivationMgr._createOmexProxy(me._contextActivationMgr._omexAnonymousWSProxy, createAnonymousOmexProxyCompleted);
                                }
                            }
                        };
                        var proxyRequired = true;
                        var params = { "contentMarket": reference.storeLocator, "assetID": reference.id };
                        if (_omexDataProvider.AllCached(context, params)) {
                            try {
                                Telemetry.AppLoadTimeHelper.AuthenticationEnd(me._telemetryContext);
                                context.anonymous = true;
                                context.clientEndPoint = {};
                                Telemetry.AppLoadTimeHelper.KilledAppsCheckStart(me._telemetryContext);
                                OSF.OsfManifestManager.getOmexKilledAppsAsync(context, Function.createDelegate(me, me._onGetOmexKilledAppsCompleted));
                                Telemetry.AppLoadTimeHelper.AppStateCheckStart(me._telemetryContext);
                                OSF.OsfManifestManager.getOmexAppStateAsync(context, Function.createDelegate(me, me._onGetOmexAppStateCompleted));
                                proxyRequired = false;
                            }
                            catch (e) {
                                Telemetry.RuntimeTelemetryHelper.LogExceptionTag("omexDataProvider cache check failed.", e, null, 0x010c5399);
                            }
                        }
                        if (proxyRequired) {
                            if (this._contextActivationMgr._omexAuthNStatus == OSF.OmexAuthNStatus.CheckFailed) {
                                if (this._contextActivationMgr._omexForceAnonymous) {
                                    this._contextActivationMgr._createOmexProxy(this._contextActivationMgr._omexAnonymousWSProxy, createAnonymousOmexProxyCompleted);
                                }
                                else {
                                    this._contextActivationMgr._createOmexProxy(this._contextActivationMgr._omexGatedWSProxy, createAuthenticatedOmexProxyCompleted);
                                }
                            }
                            else {
                                if (this._contextActivationMgr._omexForceAnonymous) {
                                    this._contextActivationMgr._createOmexProxy(this._contextActivationMgr._omexAnonymousWSProxy, createAnonymousOmexProxyCompleted);
                                }
                                else {
                                    this._contextActivationMgr._omexAuthNStatus = OSF.OmexAuthNStatus.NotAttempted;
                                    this._contextActivationMgr._createOmexProxy(this._contextActivationMgr._omexGatedWSProxy, createAuthenticatedOmexProxyCompleted);
                                }
                            }
                        }
                        this._contextActivationMgr._preloadOfficeJs();
                    }
                    else {
                        this._showActivationError(OSF.OsfControlStatus.ActivationFailed, Strings.OsfRuntime.L_AgaveOmexNotConfigured_ERR, null, null, null, null, null, null, OSF.ErrorStatusCodes.E_TRUSTCENTER_CATALOG_UNTRUSTED_ADMIN_CONTROLLED);
                    }
                }
                else if (reference.storeType === OSF.StoreType.FileSystem || reference.storeType === OSF.StoreType.Registry) {
                    this._showActivationWarning(OSF.OsfControlStatus.UnsupportedStore, Strings.OsfRuntime.L_AgaveUnsupportedStoreType_ERR, null, null, null, null, OSF.ErrorStatusCodes.WAC_AgaveUnsupportedStoreType);
                }
                else if (reference.storeType === OSF.StoreType.OneDrive) {
                    this._showActivationWarning(OSF.OsfControlStatus.UnsupportedStore, Strings.OsfRuntime.L_AgaveUnsupportedStoreType_ERR, null, null, null, null, OSF.ErrorStatusCodes.WAC_AgaveUnsupportedStoreType);
                }
                else {
                    this._showActivationError(OSF.OsfControlStatus.UnknownStore, Strings.OsfRuntime.L_AgaveUnknownStoreType_ERR, null, null, null, null, null, null, OSF.ErrorStatusCodes.E_MANIFEST_REFERENCE_INVALID);
                }
            }
        }
        catch (ex) {
            OsfMsAjaxFactory.msAjaxDebug.trace("Error getting app data: " + ex);
            Telemetry.RuntimeTelemetryHelper.LogExceptionTag("Error getting app data.", ex, this.getCorrelationId(), 0x007d4263);
            this._showActivationError(OSF.OsfControlStatus.ActivationFailed, Strings.OsfRuntime.L_AgaveServerConnectionFailed_ERR, Strings.OsfRuntime.L_RetryButton_TXT, Function.createDelegate(this, this._refresh), null, null, null, null, OSF.ErrorStatusCodes.WAC_AgaveOsfControlActivationError);
        }
        this.invokePreactivationCompletedCallback();
    },
    deActivate: function OSF_OsfControl$deActivate() {
        try {
            OSF.OUtil.writeProfilerMark(OSF.OsfControlActivationPerfMarker.DeactivationStart);
            this._contextActivationMgr.dismissMessages(this._id);
            this._retryActivate = null;
            if (this._status !== OSF.OsfControlStatus.NotActivated) {
                if (this._agaveEndPoint) {
                    Microsoft.Office.Common.XdmCommunicationManager.deleteClientEndPoint(this._agaveEndPoint._conversationId);
                    this._agaveEndPoint = null;
                }
                if (this._frame) {
                    OSF.OUtil.removeEventListener(this._frame, "load", this._iframeOnLoadDelegate);
                    this._frame = null;
                }
                if (this._timer) {
                    window.clearTimeout(this._timer);
                    this._timer = null;
                }
                while (this._div.childNodes.length > 0) {
                    this._div.removeChild(this._div.childNodes.item(0));
                }
                this._status = OSF.OsfControlStatus.NotActivated;
                if (this._conversationId) {
                    this._contextActivationMgr._getServiceEndPoint().unregisterConversation(this._conversationId);
                }
                this._contextActivationMgr.raiseOsfControlStatusChange(this);
            }
            this._controlFocus = false;
            OSF.OUtil.writeProfilerMark(OSF.OsfControlActivationPerfMarker.DeactivationEnd);
        }
        catch (ex) {
            OsfMsAjaxFactory.msAjaxDebug.trace("Deactivate failed: " + ex);
            Telemetry.RuntimeTelemetryHelper.LogExceptionTag("Deactivate failed.", ex, this.getCorrelationId(), 0x007d4280);
        }
    },
    purge: function OSF_OsfControl$purge(purgeManifest) {
        var e = Function._validateParams(arguments, [
            { name: "purgeManifest", type: Boolean, mayBeNull: false }
        ]);
        if (e)
            throw e;
        try {
            this._contextActivationMgr._ErrorUXHelper.purgeOsfControlNotification(this._id);
            this.deActivate();
            if (purgeManifest)
                OSF.OsfManifestManager.purgeManifest(this._marketplaceID, this._marketplaceVersion);
            this._contextActivationMgr.unregisterOsfControl(this);
        }
        catch (ex) {
            OsfMsAjaxFactory.msAjaxDebug.trace("Purge failed: " + ex);
            Telemetry.RuntimeTelemetryHelper.LogExceptionTag("Purge failed.", ex, this.getCorrelationId(), 0x007d4281);
        }
    },
    invokeVirtualOsfControlActivationCallback: function OSF_OsfControl$invokeVirtualOsfControlActivationCallback(manifest) {
        if (!this._isvirtualOsfControlCallbackInvoked && this._virtualOsfControlActivationCallback) {
            this._virtualOsfControlActivationCallback(this, manifest);
            this._isvirtualOsfControlCallbackInvoked = true;
        }
    },
    onForceGetVirtualOsfControlManifestCompleted: function OSF_OsfControl$onForceGetVirtualOsfControlManifestCompleted(asyncResult) {
        var manifest = null;
        if (asyncResult.statusCode === OSF.ProxyCallStatusCode.Succeeded && asyncResult.value) {
            manifest = asyncResult.value;
        }
        else {
            this._status = OSF.OsfControlStatus.ActivationFailed;
            this._errorcode = OSF.ErrorStatusCodes.E_MANIFEST_SERVER_UNAVAILABLE;
        }
        this.invokeVirtualOsfControlActivationCallback(manifest);
    },
    forceGetVirtualOsfControlManifest: function OSF_OsfControl$forceGetVirtualOsfControlManifest() {
        var context = {
            "contextActivationMgr": null,
            "referenceInUse": null
        };
        context.referenceInUse = { id: this._marketplaceID, version: this._marketplaceVersion, storeType: this._storeType, storeLocator: this._storeLocator };
        context.contextActivationMgr = this._contextActivationMgr;
        if (this._dataProvider != null) {
            this._dataProvider.getManifestAsync(context, this.onForceGetVirtualOsfControlManifestCompleted);
        }
        else {
            alert("AddinCommands is not supported in " + context.referenceInUse.storeType + " store.");
            this.onForceGetVirtualOsfControlManifestCompleted({ "statusCode": OSF.ProxyCallStatusCode.Failed, "value": null });
        }
    },
    invokePreactivationCompletedCallback: function OSF_OsfControl$invokePreactivationCompletedCallback() {
        if (this._preactivationCallback) {
            this._preactivationCallback();
        }
    },
    getMarketplaceID: function OSF_OsfControl$getMarketplaceID() {
        return this._marketplaceID;
    },
    getMarketplaceVersion: function OSF_OsfControl$getMarketplaceVersion() {
        return this._marketplaceVersion;
    },
    getContainingDiv: function OSF_OsfControl$getContainingDiv() {
        return this._div;
    },
    getID: function OSF_OsfControl$getID() {
        return this._id;
    },
    getSettings: function OSF_OsfControl$getSettings() {
        return this._settings;
    },
    setSettings: function OSF_OsfControl$setSettings(settings) {
        this._settings = settings;
    },
    getReason: function OSF_OsfControl$getReason() {
        return this._reason;
    },
    getOsfControlType: function OSF_OsfControl$getOsfControlType() {
        return this._osfControlType;
    },
    getSnapshotUrl: function OSF_OsfControl$getSnapshotUrl() {
        return this._snapshotUrl;
    },
    getStoreType: function OSF_OsfControl$getStoreType() {
        return this._storeType;
    },
    getStoreLocator: function OSF_OsfControl$getStoreLocator() {
        return this._storeLocator;
    },
    getProperty: function OSF_OsfControl$getProperty(name) {
        var e = Function._validateParams(arguments, [
            { name: "name", type: String, mayBeNull: false }
        ]);
        if (e)
            throw e;
        return this._settings[name];
    },
    addProperty: function OSF_OsfControl$addProperty(name, value) {
        var e = Function._validateParams(arguments, [
            { name: "name", type: String, mayBeNull: false },
            { name: "value", type: String, mayBeNull: false }
        ]);
        if (e)
            throw e;
        this._settings[name] = value;
    },
    removeProperty: function OSF_OsfControl$removeProperty(name) {
        var e = Function._validateParams(arguments, [
            { name: "name", type: String, mayBeNull: false }
        ]);
        if (e)
            throw e;
        delete this._settings[name];
    },
    getStatus: function OSF_OsfControl$getStatus() {
        return this._status;
    },
    getErrorcode: function OSF_OsfControl$getErrorcode() {
        return this._errorcode;
    },
    getNotificationParams: function OSF_OsfControl$getNotificationParams() {
        return this._notificationParams;
    },
    getIframeUrl: function OSF_OsfControl$getIframeUrl() {
        return this._iframeUrl;
    },
    getPermission: function OSF_OsfControl$getPermission() {
        return this._permission;
    },
    getTrustNoPrompt: function OSF_OsfControl$getTrustNoPrompt() {
        return false;
    },
    getEToken: function OSF_OsfControl$getEToken() {
        return this._omexEntitlement ? this._omexEntitlement.etoken : this._etoken;
    },
    getCorrelationId: function OSF_OsfControl$getCorrelationId() {
        return this._appCorrelationId;
    },
    notifyAgave: function OSF_OsfControl$notifyAgave(actionId) {
        if (this._agaveEndPoint) {
            this._agaveEndPoint.invoke("Office_notifyAgave", null, actionId);
        }
    },
    _onGetEntitlementsCompleted: function OSF_OsfControl$_onGetEntitlementsCompleted(asyncResult) {
        if (asyncResult.context && asyncResult.context.referenceInUse.storeType === OSF.StoreType.SPCatalog) {
            OSF.OUtil.writeProfilerMark(OSF.OsfOfficeExtensionManagerPerfMarker.GetEntitlementEnd);
        }
        if (asyncResult.statusCode === OSF.ProxyCallStatusCode.Succeeded && asyncResult.value) {
            Telemetry.AppLoadTimeHelper.EntitlementCheckEnd(this._telemetryContext);
            var reference = asyncResult.context.referenceInUse;
            var entitlements = asyncResult.value.entitlements;
            var entitlementCount = entitlements.length;
            var entitlement;
            var newestEntitlement = null;
            for (var i = 0; i < entitlementCount; i++) {
                entitlement = entitlements[i];
                if (entitlement.OfficeExtensionID && reference.id && entitlement.OfficeExtensionID.toLowerCase() === reference.id.toLowerCase()) {
                    if (!newestEntitlement || this._lessThan(newestEntitlement.OfficeExtensionVersion, entitlement.OfficeExtensionVersion)) {
                        newestEntitlement = entitlement;
                    }
                }
            }
            entitlement = newestEntitlement;
            if (!entitlement) {
                this._showActivationError(OSF.OsfControlStatus.ActivationFailed, Strings.OsfRuntime.L_AgaveNotExist_ERR, Strings.OsfRuntime.L_RetryButton_TXT, Function.createCallback(function (context) {
                    context.osfControl._refresh(context);
                }, { "clearCache": true, "referenceInUse": asyncResult.context.referenceInUse, "osfControl": asyncResult.context.osfControl }), null, null, null, null, OSF.ErrorStatusCodes.E_MANIFEST_DOES_NOT_EXIST);
            }
            else if (entitlement.OfficeExtensionKillbit) {
                this._showActivationError(OSF.OsfControlStatus.ActivationFailed, Strings.OsfRuntime.L_AgaveDisabledByAdmin_ERR, null, null, null, null, null, null, OSF.ErrorStatusCodes.E_OEM_EXTENSION_KILLED);
            }
            else {
                asyncResult.context.manifestUrl = entitlement.EncodedAbsUrl;
                asyncResult.context.appInstanceId = entitlement.AppInstanceID;
                asyncResult.context.productId = entitlement.ProductID;
                if (asyncResult.context.appInstanceId) {
                    OSF.OsfManifestManager.getAppInstanceInfoByIdAsync(asyncResult.context, Function.createDelegate(this, this._onGetAppInstanceInfoByIdCompleted));
                }
                else {
                    OSF.OsfManifestManager.getManifestAsync(asyncResult.context, Function.createDelegate(this, this._onGetManifestCompleted));
                }
            }
        }
        else {
            this._showActivationError(OSF.OsfControlStatus.ActivationFailed, Strings.OsfRuntime.L_AgaveServerConnectionFailed_ERR, Strings.OsfRuntime.L_RetryButton_TXT, Function.createDelegate(this, this._refresh), null, null, null, null, OSF.ErrorStatusCodes.E_MANIFEST_SERVER_UNAVAILABLE);
        }
    },
    _onGetAppInstanceInfoByIdCompleted: function OSF_OsfControl$_onGetAppInstanceInfoByIdCompleted(asyncResult) {
        if (asyncResult.statusCode === OSF.ProxyCallStatusCode.Succeeded && asyncResult.value) {
            var appInstanceInfo = asyncResult.value;
            var context = asyncResult.context;
            if (appInstanceInfo.AppWebFullUrl) {
                context.appWebUrl = appInstanceInfo.AppWebFullUrl;
            }
            context.clientId = appInstanceInfo.AppPrincipalId;
            context.remoteAppUrl = appInstanceInfo.RemoteAppUrl;
            if (context.appWebUrl && context.productId) {
                OSF.OsfManifestManager.getSPTokenByProductIdAsync(context, Function.createDelegate(this, this._onGetSPTokenByProductIdCompleted));
            }
            else {
                OSF.OsfManifestManager.getManifestAsync(context, Function.createDelegate(this, this._onGetManifestCompleted));
            }
        }
        else {
            this._showActivationError(OSF.OsfControlStatus.ActivationFailed, Strings.OsfRuntime.L_AgaveServerConnectionFailed_ERR, Strings.OsfRuntime.L_RetryButton_TXT, Function.createDelegate(this, this._refresh), null, null, null, null, OSF.ErrorStatusCodes.E_MANIFEST_SERVER_UNAVAILABLE);
        }
    },
    _onGetSPTokenByProductIdCompleted: function OSF_OsfControl$_onGetSPTokenByProductIdCompleted(asyncResult) {
        if (asyncResult.statusCode === OSF.ProxyCallStatusCode.Succeeded && asyncResult.value) {
            this._etoken = asyncResult.value;
        }
        OSF.OsfManifestManager.getManifestAsync(asyncResult.context, Function.createDelegate(this, this._onGetManifestCompleted));
    },
    _onGetOmexEntitlementsCompleted: function OSF_OsfControl$_onGetOmexEntitlementsCompleted(asyncResult) {
        if (asyncResult.statusCode === OSF.ProxyCallStatusCode.Succeeded && asyncResult.value) {
            Telemetry.AppLoadTimeHelper.EntitlementCheckEnd(this._telemetryContext);
            var reference = asyncResult.context.referenceInUse;
            var entitlements = asyncResult.value.entitlements;
            var entitlementCount = entitlements.length;
            var entitlement;
            var found = false;
            this._contextActivationMgr._omexBillingMarket = asyncResult.value.billingMarket;
            _omexDataProvider.SetCustomerId(asyncResult.value.cid);
            for (var i = 0; i < entitlementCount; i++) {
                entitlement = entitlements[i];
                if (entitlement.assetId && reference.id && entitlement.assetId.toLowerCase() === reference.id.toLowerCase()) {
                    found = true;
                    break;
                }
            }
            if (found) {
                this._omexEntitlement.hasEntitlement = true;
                this._omexEntitlement.productId = entitlement.productId;
                this._omexEntitlement.version = entitlement.version;
                this._omexEntitlement.contentMarket = entitlement.contentMarket;
                this._omexEntitlement.licenseType = entitlement.licenseType;
                this._omexEntitlement.endPointUrl = this._contextActivationMgr._getOmexEndPointPageUrl(this._omexEntitlement.assetId, this._omexEntitlement.contentMarket);
            }
            Telemetry.AppLoadTimeHelper.SetOmexHasEntitlementFlag(this._telemetryContext, this._omexEntitlement.hasEntitlement);
            Telemetry.AppLoadTimeHelper.AppStateCheckStart(this._telemetryContext);
            OSF.OsfManifestManager.getOmexAppStateAsync(asyncResult.context, Function.createDelegate(this, this._onGetOmexAppStateCompleted));
        }
        else {
            Telemetry.RuntimeTelemetryHelper.LogExceptionTag("getOmexEntilements failed:" + OSF.OUtil.serializeOMEXResponseErrorMessage(asyncResult), null, null, 0x010c539a);
            this._showActivationError(OSF.OsfControlStatus.ActivationFailed, Strings.OsfRuntime.L_AgaveServerConnectionFailed_ERR, Strings.OsfRuntime.L_RetryButton_TXT, Function.createDelegate(this, this._refresh), null, null, null, null, OSF.ErrorStatusCodes.E_MANIFEST_SERVER_UNAVAILABLE);
        }
    },
    _onGetOmexKilledAppsCompleted: function OSF_OsfControl$_onGetOmexKilledAppsCompleted(asyncResult) {
        if (asyncResult.statusCode === OSF.ProxyCallStatusCode.Succeeded && asyncResult.value) {
            Telemetry.AppLoadTimeHelper.KilledAppsCheckEnd(this._telemetryContext);
            if (!asyncResult.cached) {
                _omexDataProvider.SetKilledAppsCache(asyncResult.context, asyncResult.value);
            }
            var killedApps = asyncResult.value.killedApps;
            var len = killedApps.length;
            var found = false;
            for (var i = 0; i < len; i++) {
                if (killedApps[i].assetId === this._marketplaceID) {
                    found = true;
                    break;
                }
            }
            if (found) {
                this._showActivationError(OSF.OsfControlStatus.ActivationFailed, Strings.OsfRuntime.L_AgaveDisabledByOmex_ERR, null, null, this._omexEntitlement.endPointUrl, null, null, null, OSF.ErrorStatusCodes.E_OEM_OMEX_EXTENSION_KILLED);
                return;
            }
        }
        else {
            Telemetry.RuntimeTelemetryHelper.LogExceptionTag("getOmexKilledApps failed:" + OSF.OUtil.serializeOMEXResponseErrorMessage(asyncResult), null, null, 0x010c539b);
        }
    },
    _onGetOmexAppStateCompleted: function OSF_OsfControl$_onGetOmexAppStateCompleted(asyncResult) {
        if (asyncResult.statusCode === OSF.ProxyCallStatusCode.Succeeded && asyncResult.value) {
            Telemetry.AppLoadTimeHelper.AppStateCheckEnd(this._telemetryContext);
            Telemetry.AppLoadTimeHelper.SetAppStateDataCachedFlag(this._telemetryContext, asyncResult.cached);
            var context = asyncResult.context;
            var appState = asyncResult.value;
            var state = parseInt(appState.state);
            if (state === OSF.OmexState.Killed) {
                this._showActivationError(OSF.OsfControlStatus.ActivationFailed, Strings.OsfRuntime.L_AgaveDisabledByOmex_ERR, null, null, this._omexEntitlement.endPointUrl, null, null, null, OSF.ErrorStatusCodes.E_OEM_OMEX_EXTENSION_KILLED);
                return;
            }
            else if (state === OSF.OmexState.Flagged) {
                context.showSoftKilled = true;
            }
            else if (state === OSF.OmexState.DeveloperWithdrawn) {
                context.showDeveloperWithDrawWarning = true;
            }
            if (!this._omexEntitlement.hasEntitlement && (context.showDeveloperWithDrawWarning || context.showSoftKilled)) {
                this._showActivationError(OSF.OsfControlStatus.ActivationFailed, Strings.OsfRuntime.L_AgaveRetired_ERR, null, null, this._omexEntitlement.endPointUrl, null, null, null, OSF.ErrorStatusCodes.E_OEM_EXTENSION_WITHDRAWN_FROM_SALE);
                if (context.showSoftKilled) {
                    return;
                }
            }
            this._omexEntitlement.latestVersion = appState.version;
            if (this._omexEntitlement.hasEntitlement && this._lessThan(this._omexEntitlement.version, this._omexEntitlement.latestVersion)) {
                context.showNewerVersion = true;
                context.expectedVersion = this._omexEntitlement.latestVersion;
            }
            if (!asyncResult.cached) {
                _omexDataProvider.SetAppStateCache(context, appState);
            }
            if (!context.anonymous && this._reason != Microsoft.Office.WebExtension.InitializationReason.Inserted && !this._omexEntitlement.hasEntitlement) {
                var me = this;
                var createAnonymousOmexProxyCompleted = function (asyncResult) {
                    context.anonymous = true;
                    context.clientEndPoint = me._contextActivationMgr._omexAnonymousWSProxy.clientEndPoint;
                    Telemetry.AppLoadTimeHelper.ManifestRequestStart(me._telemetryContext);
                    OSF.OsfManifestManager.getOmexManifestAndETokenAsync(context, Function.createDelegate(me, me._onGetOmexManifestAndETokenCompleted));
                };
                me._contextActivationMgr._createOmexProxy(me._contextActivationMgr._omexAnonymousWSProxy, createAnonymousOmexProxyCompleted);
                return;
            }
        }
        else {
            Telemetry.RuntimeTelemetryHelper.LogExceptionTag("getOmexAppState failed:" + OSF.OUtil.serializeOMEXResponseErrorMessage(asyncResult), null, null, 0x010c539c);
        }
        Telemetry.AppLoadTimeHelper.ManifestRequestStart(this._telemetryContext);
        OSF.OsfManifestManager.getOmexManifestAndETokenAsync(asyncResult.context, Function.createDelegate(this, this._onGetOmexManifestAndETokenCompleted));
    },
    _onGetOmexManifestAndETokenCompleted: function OSF_OsfControl$_onGetOmexManifestAndETokenCompleted(asyncResult) {
        if (asyncResult.statusCode === OSF.ProxyCallStatusCode.Succeeded && asyncResult.value) {
            var manifestAndEToken = asyncResult.value;
            var clientAppStatus = parseInt(manifestAndEToken.status);
            var context = asyncResult.context;
            var reference = context.referenceInUse;
            Telemetry.AppLoadTimeHelper.SetManifestDataCachedFlag(this._telemetryContext, asyncResult.cached);
            if (clientAppStatus === OSF.OmexClientAppStatus.OK) {
                if (context.acceptedUpgrade) {
                    this._refresh({ "clearEntitlement": true, "clearAppState": true, "referenceInUse": reference, "osfControl": this });
                    return;
                }
                if (manifestAndEToken.tokenExpirationDate && new Date(manifestAndEToken.tokenExpirationDate) <= new Date()) {
                    this._showActivationWarning(OSF.OsfControlStatus.ActivationFailed, Strings.OsfRuntime.L_AgaveLicenseExpired_ERR, Strings.OsfRuntime.L_RetryButton_TXT, Function.createDelegate(this, this._refresh), null, null, OSF.ErrorStatusCodes.E_TOKEN_EXPIRED);
                    return;
                }
                this._omexEntitlement.etoken = manifestAndEToken.etoken;
                this._omexEntitlement.entitlementType = manifestAndEToken.entitlementType;
                if (!context.anonymous && this._omexEntitlement.entitlementType && (this._omexEntitlement.entitlementType.toLowerCase() === OSF.OmexEntitlementType.Trial)) {
                    context.showTrialInfo = true;
                }
                try {
                    if (!asyncResult.cached) {
                        _omexDataProvider.SetManifestAndETokenCache(context, manifestAndEToken);
                    }
                    var manifest = new OSF.Manifest.Manifest(manifestAndEToken.manifest, this._contextActivationMgr.getAppUILocale());
                    context.manifestCached = manifestAndEToken.cached;
                    OSF.OsfManifestManager.cacheManifest(reference.id, reference.version, manifest);
                    if (this._omexEntitlement.latestVersion && this._lessThan(manifest.getMarketplaceVersion(), this._omexEntitlement.latestVersion)) {
                        context.showNewerVersion = true;
                        context.expectedVersion = this._omexEntitlement.latestVersion;
                    }
                    this._onGetManifestCompleted({ "statusCode": OSF.ProxyCallStatusCode.Succeeded, "value": manifest, "context": context });
                }
                catch (ex) {
                    OsfMsAjaxFactory.msAjaxDebug.trace("Invalid manifest from marketplace: " + ex);
                    Telemetry.RuntimeTelemetryHelper.LogExceptionTag("Invalid manifest from marketplace.", ex, this.getCorrelationId(), 0x007d4282);
                    this._showActivationError(OSF.OsfControlStatus.ActivationFailed, Strings.OsfRuntime.L_AgaveManifestRetrieve_ERR, Strings.OsfRuntime.L_RetryButton_TXT, Function.createDelegate(this, this._refresh), null, null, null, null, OSF.ErrorStatusCodes.E_MANIFEST_INVALID_VALUE_FORMAT);
                    return;
                }
            }
            else if (clientAppStatus === OSF.OmexClientAppStatus.KilledAsset) {
                this._showActivationError(OSF.OsfControlStatus.ActivationFailed, Strings.OsfRuntime.L_AgaveDisabledByOmex_ERR, null, null, this._omexEntitlement.endPointUrl, null, null, null, OSF.ErrorStatusCodes.E_OEM_OMEX_EXTENSION_KILLED);
            }
            else if (clientAppStatus === OSF.OmexClientAppStatus.NoEntitlement || clientAppStatus === OSF.OmexClientAppStatus.TrialNotSupported ||
                clientAppStatus === OSF.OmexClientAppStatus.LimitedTrial || clientAppStatus === OSF.OmexClientAppStatus.EntitlementDeactivated) {
                if (!context.anonymous) {
                    var refreshOmexEntitlementAndToken = function (context) {
                        var thisOsfControl = context.osfControl;
                        thisOsfControl._refresh({ "clearCache": true, "referenceInUse": context.referenceInUse, "osfControl": thisOsfControl });
                    };
                    var buyPaidVersion = function (context) {
                        var thisOsfControl = context.osfControl;
                        window.open(thisOsfControl._omexEntitlement.endPointUrl);
                        thisOsfControl._contextActivationMgr.displayNotification({
                            "id": thisOsfControl._id,
                            "infoType": OSF.InfoType.Warning,
                            "title": Strings.OsfRuntime.L_AgaveWarningTitle_TXT,
                            "description": Strings.OsfRuntime.L_AgaveLicenseNotAquiredRefresh_ERR,
                            "buttonTxt": Strings.OsfRuntime.L_RefreshButton_TXT,
                            "buttonCallback": Function.createCallback(refreshOmexEntitlementAndToken, context),
                            "url": null,
                            "dismissCallback": null,
                            "detailView": true,
                            "reDisplay": true,
                            "displayDeactive": true,
                            "retryAll": true,
                            "errorCode": OSF.ErrorStatusCodes.E_OEM_EXTENSION_NOT_ENTITLED
                        });
                    };
                    this._retryActivate = Function.createCallback(refreshOmexEntitlementAndToken, context);
                    this._showActivationWarning(OSF.OsfControlStatus.ActivationFailed, Strings.OsfRuntime.L_AgaveLicenseNotAquired_ERR, Strings.OsfRuntime.L_BuyButton_TXT, Function.createCallback(buyPaidVersion, context), null, null, OSF.ErrorStatusCodes.E_OEM_EXTENSION_NOT_ENTITLED);
                }
                else {
                    var signInRedirect = function (context) {
                        var thisOsfControl = context.osfControl;
                        var currentUrl = window.location.href;
                        var signInRedirectUrl = thisOsfControl._contextActivationMgr._osfOmexBaseUrl + OSF.Constants.SignInRedirectUrl + encodeURIComponent(currentUrl);
                        window.open(signInRedirectUrl);
                    };
                    this._showActivationError(OSF.OsfControlStatus.ActivationFailed, Strings.OsfRuntime.L_AgaveNotViewableAnonymous_ERR, Strings.OsfRuntime.L_SignInButton_TXT, Function.createCallback(signInRedirect, context), null, null, Strings.OsfRuntime.L_AgaveSigninRequiredTitle_TXT, null, OSF.ErrorStatusCodes.E_USER_NOT_SIGNED_IN);
                }
            }
            else if (clientAppStatus === OSF.OmexClientAppStatus.UnknownAssetId) {
                this._showActivationError(OSF.OsfControlStatus.ActivationFailed, Strings.OsfRuntime.L_AgaveNotExist_ERR, null, null, null, null, null, null, OSF.ErrorStatusCodes.E_MANIFEST_DOES_NOT_EXIST);
            }
            else if (clientAppStatus === OSF.OmexClientAppStatus.Expired || clientAppStatus === OSF.OmexClientAppStatus.Invalid) {
                this._showActivationWarning(OSF.OsfControlStatus.ActivationFailed, Strings.OsfRuntime.L_AgaveLicenseExpired_ERR, Strings.OsfRuntime.L_RetryButton_TXT, Function.createDelegate(this, this._refresh), null, null, OSF.ErrorStatusCodes.E_TOKEN_EXPIRED);
            }
            else if (clientAppStatus === OSF.OmexClientAppStatus.Revoked) {
                this._showActivationError(OSF.OsfControlStatus.ActivationFailed, Strings.OsfRuntime.L_AgaveRetired_ERR, null, null, this._omexEntitlement.endPointUrl, null, null, null, OSF.ErrorStatusCodes.E_OEM_EXTENSION_WITHDRAWN_FROM_SALE);
            }
            else if (clientAppStatus === OSF.OmexClientAppStatus.VersionMismatch) {
                var refreshManifest = function (context) {
                    var thisOsfControl = context.osfControl;
                    thisOsfControl._refresh({ "clearManifest": true, "clearToken": true, "referenceInUse": context.referenceInUse, "osfControl": thisOsfControl, "acceptedUpgrade": true, "expectedVersion": manifestAndEToken.version });
                };
                this._contextActivationMgr.displayNotification({
                    "id": this._id,
                    "infoType": OSF.InfoType.Warning,
                    "title": Strings.OsfRuntime.L_AgaveWarningTitle_TXT,
                    "description": Strings.OsfRuntime.L_AgaveNewerVersion_ERR,
                    "buttonTxt": Strings.OsfRuntime.L_UpdateButton_TXT,
                    "buttonCallback": Function.createCallback(refreshManifest, context),
                    "url": this._omexEntitlement.endPointUrl,
                    "dismissCallback": null,
                    "detailView": false,
                    "reDisplay": true,
                    "highPriority": true,
                    "displayDeactive": true,
                    "logAsError": true,
                    "errorCode": OSF.ErrorStatusCodes.E_MANIFEST_UPDATE_AVAILABLE
                });
            }
            else if (clientAppStatus === OSF.OmexClientAppStatus.VersionNotSupported) {
                this._showActivationWarning(OSF.OsfControlStatus.UnsupportedStore, Strings.OsfRuntime.L_AgaveUnsupportedStoreType_ERR, null, null, null, null, OSF.ErrorStatusCodes.WAC_AgaveUnsupportedStoreType);
            }
            else {
                this._showActivationError(OSF.OsfControlStatus.ActivationFailed, Strings.OsfRuntime.L_AgaveServerConnectionFailed_ERR, Strings.OsfRuntime.L_RetryButton_TXT, Function.createDelegate(this, this._refresh), null, null, null, null, OSF.ErrorStatusCodes.WAC_AgaveUnknownClientAppStatus);
            }
        }
        else {
            Telemetry.RuntimeTelemetryHelper.LogExceptionTag("GetOmexManifestAndEToken failed:" + OSF.OUtil.serializeOMEXResponseErrorMessage(asyncResult), null, null, 0x010c539d);
            this._showActivationError(OSF.OsfControlStatus.ActivationFailed, Strings.OsfRuntime.L_AgaveServerConnectionFailed_ERR, Strings.OsfRuntime.L_RetryButton_TXT, Function.createDelegate(this, this._refresh), null, null, null, null, OSF.ErrorStatusCodes.E_MANIFEST_SERVER_UNAVAILABLE);
        }
    },
    _retryActivation: function OSF_OsfControl$_retryActivation() {
        this._retryLoadingNum--;
        if (this._pageStatus === OSF.OsfControlPageStatus.Ready || this._retryLoadingNum <= 0) {
            this._contextActivationMgr._ErrorUXHelper.removeProgressDiv(this._div, this._id);
            this._retryLoadingNum = 2;
        }
        else {
            this._refresh();
        }
    },
    _iframeOnLoad: function OSF_OsfControl$__iframeOnLoad() {
        var osfControl = this;
        Telemetry.AppLoadTimeHelper.PageLoaded(this._telemetryContext);
        var onTimeOut = function OSF_OsfControl$__onTimeOut(osfControl) {
            if (osfControl) {
                if (osfControl._contextActivationMgr._ErrorUXHelper) {
                    OSF.OUtil.writeProfilerMark(OSF.OsfControlActivationPerfMarker.SelectionTimeout);
                    osfControl._contextActivationMgr._ErrorUXHelper.removeProgressDiv(osfControl._div, osfControl._id);
                }
                if (osfControl._pageStatus !== OSF.OsfControlPageStatus.Ready) {
                    if (osfControl._retryLoadingNum === 2) {
                        var errorCode = OSF.ErrorStatusCodes.WAC_AgaveActivationError;
                        switch (osfControl._pageStatus) {
                            case OSF.OsfControlPageStatus.FailedOriginCheck:
                                errorCode = OSF.ErrorStatusCodes.WAC_AgaveOriginCheckError;
                                break;
                            case OSF.OsfControlPageStatus.FailedPermissionCheck:
                                errorCode = OSF.ErrorStatusCodes.WAC_AgavePermissionCheckError;
                                break;
                            case OSF.OsfControlPageStatus.Failed:
                                errorCode = OSF.ErrorStatusCodes.WAC_AgaveHostHandleRequestError;
                                break;
                        }
                        osfControl._showActivationError(OSF.OsfControlStatus.ActivationFailed, Strings.OsfRuntime.L_AgaveActivationError_ERR, Strings.OsfRuntime.L_RetryButton_TXT, Function.createDelegate(osfControl, osfControl._retryActivation), null, null, null, null, errorCode);
                    }
                    else if (osfControl._retryLoadingNum === 1) {
                        osfControl._showActivationError(OSF.OsfControlStatus.ActivationFailed, Strings.OsfRuntime.L_ActivateAttempLoading_ERR, Strings.OsfRuntime.L_ActivateButton_TXT, Function.createDelegate(osfControl, osfControl._retryActivation), null, null, null, null, OSF.ErrorStatusCodes.WAC_ActivateAttempLoading);
                    }
                }
                if (osfControl._timer) {
                    window.clearTimeout(osfControl._timer);
                    osfControl._timer = null;
                }
                osfControl._pageIsReadyTimerExpired = true;
            }
        };
        if (osfControl._pageStatus !== OSF.OsfControlPageStatus.Ready) {
            osfControl._timer = window.setTimeout(function () { onTimeOut(osfControl); }, 5000 * osfControl._retryLoadingNum + 1);
        }
    },
    _isOsfControlInEmbeddingMode: function OSF_OsfControl$_isOsfControlInEmbeddingMode(osfControl) {
        var embedded = false;
        try {
            var webExtensionDiv = osfControl.getContainingDiv().parentNode.parentNode;
            var webExtensionDivId = webExtensionDiv.id;
            var matches = webExtensionDivId.match(/^(m_excelEmbedRenderer_|ewaSynd).+$/ig);
            embedded = (matches != null);
        }
        catch (ex) {
            OsfMsAjaxFactory.msAjaxDebug.trace("_isOsfControlInEmbeddingMode error: " + ex);
        }
        return embedded;
    },
    _createIframeAndActivateOsfControl: function OSF_OsfControl$_createIframeAndActivateOsfControl(defaultDisplayName) {
        var frame = document.createElement("iframe");
        frame.setAttribute("id", this._id);
        frame.setAttribute("width", "100%");
        frame.setAttribute("height", "100%");
        frame.setAttribute("frameborder", "0");
        var iframeTitle = defaultDisplayName ? defaultDisplayName : Strings.OsfRuntime.L_IframeTitle_TXT;
        frame.setAttribute("title", iframeTitle);
        frame.style.msUserSelect = "element";
        frame.setAttribute("sandbox", "allow-scripts allow-forms allow-same-origin ms-allow-popups allow-popups");
        for (var name in this._contextActivationMgr._iframeAttributeBag) {
            frame.setAttribute(name, this._contextActivationMgr._iframeAttributeBag[name]);
        }
        this._activate(frame, this._iframeUrl);
        this._frame = frame;
    },
    _onGetManifestCompleted: function OSF_OsfControl$_onGetManifestCompleted(asyncResult) {
        if (asyncResult.statusCode === OSF.ProxyCallStatusCode.Succeeded && asyncResult.value) {
            Telemetry.AppLoadTimeHelper.ManifestRequestEnd(this._telemetryContext);
            Telemetry.AppLoadTimeHelper.ServerCallEnd(this._telemetryContext);
            var manifest = asyncResult.value;
            var context = asyncResult.context;
            var reference = context.referenceInUse;
            var currentStoreType = reference.storeType;
            var defaultDisplayName = manifest.getDefaultDisplayName();
            this._manifest = manifest;
            if (context.manifestCached && !context.retried) {
                var manifestVersion = manifest.getMarketplaceVersion();
                if (this._lessThan(manifestVersion, reference.version)) {
                    if (currentStoreType === OSF.StoreType.SPApp || currentStoreType === OSF.StoreType.SPCatalog) {
                        this._refresh({ "clearCache": true, "referenceInUse": reference, "osfControl": context.osfControl, "retried": true });
                        return;
                    }
                    else if (currentStoreType === OSF.StoreType.OMEX && !context.showNewerVersion) {
                        context.showNewerVersion = true;
                        context.expectedVersion = reference.version;
                    }
                }
            }
            if (manifest.requirementsSupported === false || manifest.requirementsSupported === undefined && !this._contextActivationMgr.getRequirementsChecker().isManifestSupported(manifest)) {
                manifest.requirementsSupported = false;
                var message, errorCode, url = null;
                if (currentStoreType === OSF.StoreType.OMEX) {
                    message = Strings.OsfRuntime.L_AgaveManifestRequirementsErrorOmex_ERR ||
                        Strings.OsfRuntime.L_AgaveManifestError_ERR;
                    errorCode = OSF.ErrorStatusCodes.WAC_AgaveRequirementsErrorOmex;
                    url = this._omexEntitlement.endPointUrl;
                }
                else {
                    message = Strings.OsfRuntime.L_AgaveManifestRequirementsError_ERR ||
                        Strings.OsfRuntime.L_AgaveManifestError_ERR;
                    errorCode = OSF.ErrorStatusCodes.WAC_AgaveRequirementsError;
                }
                this._showActivationError(OSF.OsfControlStatus.ActivationFailed, message, null, null, url, null, null, true, errorCode);
                return;
            }
            manifest.requirementsSupported = true;
            this._iframeUrl = manifest.getDefaultSourceLocation(this._contextActivationMgr.getFormFactor());
            this._permission = manifest.getPermission();
            this._appDomains = manifest.getAppDomains();
            if ((currentStoreType === OSF.StoreType.SPApp || currentStoreType === OSF.StoreType.SPCatalog) && this._iframeUrl) {
                if (context.clientId) {
                    this._iframeUrl = this._iframeUrl.replace(/~clientid/ig, context.clientId);
                }
                if (context.appWebUrl) {
                    this._iframeUrl = this._iframeUrl.replace(/~appweburl/ig, context.appWebUrl);
                }
                if (context.remoteAppUrl) {
                    this._iframeUrl = this._iframeUrl.replace(/~remoteappurl/ig, context.remoteAppUrl);
                }
            }
            if (!this._contextActivationMgr._doesUrlHaveSupportedProtocol(this._iframeUrl)) {
                this._showActivationError(OSF.OsfControlStatus.ActivationFailed, Strings.OsfRuntime.L_AgaveManifestError_ERR, null, null, null, null, null, null, OSF.ErrorStatusCodes.E_MANIFEST_INVALID_VALUE_FORMAT);
                return;
            }
            if ((!context.anonymous || !this._isOsfControlInEmbeddingMode(context.osfControl))
                && !this._isvirtualOsfControl
                && (currentStoreType != OSF.StoreType.SPApp)
                && (currentStoreType != OSF.StoreType.Exchange)
                && (currentStoreType != OSF.StoreType.InMemory)
                && (currentStoreType != OSF.StoreType.Registry)
                && (currentStoreType != OSF.StoreType.UploadFileDevCatalog)) {
                var hasAgaveBeenActivatedBefore = false;
                var autoTrusted = (this._reason == Microsoft.Office.WebExtension.InitializationReason.Inserted) || this._contextActivationMgr._autoTrusted || false;
                context.cacheKey = OSF.OUtil.formatString(OSF.Constants.ActivatedCacheKey, context.referenceInUse.id.toLowerCase(), context.referenceInUse.storeType, context.referenceInUse.storeLocator);
                hasAgaveBeenActivatedBefore = this._getCachedFlag(context.cacheKey);
                if ((currentStoreType == OSF.StoreType.OMEX) && (_omexDataProvider.GetCustomerId() != "0")) {
                    hasAgaveBeenActivatedBefore = this._omexEntitlement && this._omexEntitlement.hasEntitlement;
                }
                Telemetry.AppLoadTimeHelper.SetManifestTrustCachedFlag(this._telemetryContext, hasAgaveBeenActivatedBefore);
                if (!hasAgaveBeenActivatedBefore && !autoTrusted && !this.getTrustNoPrompt()) {
                    OSF.OUtil.writeProfilerMark(OSF.OsfControlActivationPerfMarker.ActivationEnd);
                    var manualActivate = function (context) {
                        var thisOsfControl = context.osfControl;
                        thisOsfControl._setCachedFlag(context.cacheKey);
                        Telemetry.AppLoadTimeHelper.ActivationStart(thisOsfControl._telemetryContext, thisOsfControl._sqmDWords[0], thisOsfControl._sqmDWords[1], thisOsfControl._appCorrelationId, thisOsfControl._id);
                        this._pageStatus = OSF.OsfControlPageStatus.Loading;
                        if (currentStoreType == OSF.StoreType.OMEX && _omexDataProvider.GetCustomerId() != "0") {
                            context.anonymous = false;
                            context.clientEndPoint = thisOsfControl._contextActivationMgr._omexGatedWSProxy.clientEndPoint;
                            var onGetOmexManifestAndETokenCompleted = function (asyncResult) {
                                if (asyncResult.statusCode === OSF.ProxyCallStatusCode.Succeeded && asyncResult.value) {
                                    var manifestAndEToken = asyncResult.value;
                                    var clientAppStatus = parseInt(manifestAndEToken.status);
                                    var context = asyncResult.context;
                                    if (clientAppStatus === OSF.OmexClientAppStatus.OK) {
                                        if (!asyncResult.cached) {
                                            _omexDataProvider.SetManifestAndETokenCache(context, manifestAndEToken);
                                        }
                                        thisOsfControl._refresh({ "clearEntitlement": true, "osfControl": thisOsfControl });
                                        return;
                                    }
                                }
                                else {
                                    Telemetry.RuntimeTelemetryHelper.LogExceptionTag("Manual active getOmexManifestAndEToken failed:" + OSF.OUtil.serializeOMEXResponseErrorMessage(asyncResult), null, null, 0x010c539e);
                                }
                                thisOsfControl._showActivationError(OSF.OsfControlStatus.ActivationFailed, Strings.OsfRuntime.L_AgaveServerConnectionFailed_ERR, Strings.OsfRuntime.L_RetryButton_TXT, Function.createDelegate(thisOsfControl, thisOsfControl._refresh), null, null, null, null, OSF.ErrorStatusCodes.E_MANIFEST_SERVER_UNAVAILABLE);
                            };
                            OSF.OsfManifestManager.getOmexManifestAndETokenAsync(context, Function.createDelegate(thisOsfControl, onGetOmexManifestAndETokenCompleted));
                        }
                        else {
                            Telemetry.AppLoadTimeHelper.SetAnonymousFlag(thisOsfControl._telemetryContext, context.anonymous);
                            thisOsfControl._createIframeAndActivateOsfControl(defaultDisplayName);
                        }
                    };
                    this._showTrustError(defaultDisplayName, manifest.getProviderName(), currentStoreType, Function.createCallback(manualActivate, context));
                    return;
                }
                if (this._reason == Microsoft.Office.WebExtension.InitializationReason.Inserted) {
                    this._contextActivationMgr._setCachedFlag(context.cacheKey);
                    context.osfControl._retryActivate = null;
                    this._contextActivationMgr.retryAll(context.osfControl._marketplaceID);
                }
            }
            if (context.showNewerVersion) {
                var refreshManifest = function (context) {
                    var thisOsfControl = context.osfControl;
                    thisOsfControl._refresh({ "clearManifest": true, "clearToken": true, "referenceInUse": context.referenceInUse, "osfControl": thisOsfControl, "acceptedUpgrade": true, "expectedVersion": context.expectedVersion, "retried": true });
                };
                this._contextActivationMgr.displayNotification({
                    "id": this._id,
                    "infoType": OSF.InfoType.Warning,
                    "title": Strings.OsfRuntime.L_AgaveWarningTitle_TXT,
                    "description": Strings.OsfRuntime.L_AgaveNewerVersion_ERR,
                    "buttonTxt": Strings.OsfRuntime.L_UpdateButton_TXT,
                    "buttonCallback": Function.createCallback(refreshManifest, context),
                    "url": this._omexEntitlement.endPointUrl,
                    "dismissCallback": null,
                    "detailView": false,
                    "reDisplay": true,
                    "highPriority": true,
                    "displayDeactive": false,
                    "errorCode": OSF.ErrorStatusCodes.E_MANIFEST_UPDATE_AVAILABLE
                });
            }
            else if (context.showDeveloperWithDrawWarning) {
                this._contextActivationMgr.displayNotification({
                    "id": this._id,
                    "infoType": OSF.InfoType.Warning,
                    "title": Strings.OsfRuntime.L_AgaveWarningTitle_TXT,
                    "description": Strings.OsfRuntime.L_AgaveRetiring_ERR,
                    "buttonTxt": Strings.OsfRuntime.L_OkButton_TXT,
                    "buttonCallback": null,
                    "url": this._omexEntitlement.endPointUrl,
                    "dismissCallback": null,
                    "detailView": false,
                    "reDisplay": true,
                    "displayDeactive": false,
                    "errorCode": OSF.ErrorStatusCodes.S_OEM_EXTENSION_DEVELOPER_WITHDRAWN_FROM_SALE
                });
            }
            else if (context.showSoftKilled) {
                this._contextActivationMgr.displayNotification({
                    "id": this._id,
                    "infoType": OSF.InfoType.Warning,
                    "title": Strings.OsfRuntime.L_AgaveWarningTitle_TXT,
                    "description": Strings.OsfRuntime.L_AgaveSoftKilled_ERR,
                    "buttonTxt": Strings.OsfRuntime.L_OkButton_TXT,
                    "buttonCallback": null,
                    "url": this._omexEntitlement.endPointUrl,
                    "dismissCallback": null,
                    "detailView": false,
                    "reDisplay": true,
                    "displayDeactive": false,
                    "errorCode": OSF.ErrorStatusCodes.S_OEM_EXTENSION_FLAGGED
                });
            }
            else if (context.showTrialInfo) {
                var refreshOmexEntitlementAndToken = function (context) {
                    var thisOsfControl = context.osfControl;
                    thisOsfControl._refresh({ "clearCache": true, "referenceInUse": context.referenceInUse, "osfControl": thisOsfControl });
                };
                var buyTrialVersion = function (context) {
                    var thisOsfControl = context.osfControl;
                    window.open(thisOsfControl._omexEntitlement.endPointUrl);
                };
                this._contextActivationMgr.displayNotification({
                    "id": this._id,
                    "infoType": OSF.InfoType.Information,
                    "title": Strings.OsfRuntime.L_AgaveInformationTitle_TXT,
                    "description": Strings.OsfRuntime.L_AgaveTrial_ERR,
                    "buttonTxt": Strings.OsfRuntime.L_BuyButton_TXT,
                    "buttonCallback": Function.createCallback(buyTrialVersion, context),
                    "url": null,
                    "dismissCallback": null,
                    "reDisplay": true,
                    "highPriority": true,
                    "displayDeactive": false,
                    "errorCode": OSF.ErrorStatusCodes.S_OEM_EXTENSION_TRIAL_MODE
                });
                this._contextActivationMgr.displayNotification({
                    "id": this._id,
                    "infoType": OSF.InfoType.Information,
                    "title": Strings.OsfRuntime.L_AgaveInformationTitle_TXT,
                    "description": Strings.OsfRuntime.L_AgaveTrialRefresh_ERR,
                    "buttonTxt": Strings.OsfRuntime.L_RefreshButton_TXT,
                    "buttonCallback": Function.createCallback(refreshOmexEntitlementAndToken, context),
                    "url": null,
                    "dismissCallback": null,
                    "detailView": true,
                    "reDisplay": true,
                    "highPriority": true,
                    "displayDeactive": false,
                    "errorCode": OSF.ErrorStatusCodes.S_USER_CLICKED_BUY
                });
            }
            if (!this._isvirtualOsfControl) {
                this._createIframeAndActivateOsfControl(defaultDisplayName);
            }
            else if (!context.showNewerVersion) {
                this.invokeVirtualOsfControlActivationCallback(manifest);
            }
        }
        else {
            Telemetry.RuntimeTelemetryHelper.LogExceptionTag("GetManifest failed:" + OSF.OUtil.serializeOMEXResponseErrorMessage(asyncResult), null, null, 0x010c539f);
            this._showActivationError(OSF.OsfControlStatus.ActivationFailed, Strings.OsfRuntime.L_AgaveManifestRetrieve_ERR, Strings.OsfRuntime.L_RetryButton_TXT, Function.createDelegate(this, this._refresh), null, null, null, null, OSF.ErrorStatusCodes.E_MANIFEST_SERVER_UNAVAILABLE);
        }
    },
    _lessThan: function OSF_OsfControl$_lessThan(version1, version2) {
        return OSF.OsfManifestManager.versionLessThan(version1, version2);
    },
    _showTrustError: function OSF_OsfControl$_showTrustError(displayName, providerName, storeType, onManualActivate) {
        var agaveName = OSF.OUtil.formatString(Strings.OsfRuntime.L_AgaveName_INFO, displayName ? displayName : "");
        agaveName = this._contextActivationMgr._ErrorUXHelper.getHTMLEncodedString(agaveName);
        var agaveProvider = OSF.OUtil.formatString(Strings.OsfRuntime.L_AgaveProvider_INFO, providerName ? providerName : "");
        agaveProvider = this._contextActivationMgr._ErrorUXHelper.getHTMLEncodedString(agaveProvider);
        var messageToDisplay = OSF.OUtil.formatString(Strings.OsfRuntime.L_AgaveUntrusted_INFO, agaveName, agaveProvider);
        this._retryActivate = onManualActivate;
        this._errorcode = OSF.ErrorStatusCodes.E_TRUSTCENTER_MOE_UNACTIVATED;
        this._contextActivationMgr.displayNotification({
            "id": this._id,
            "infoType": OSF.InfoType.SecurityInfo,
            "title": Strings.OsfRuntime.L_AgaveNewAppTitle_TXT,
            "description": messageToDisplay,
            "buttonTxt": Strings.OsfRuntime.L_ActivateButton_TXT,
            "buttonCallback": onManualActivate,
            "url": storeType === OSF.StoreType.OMEX ? this._omexEntitlement.endPointUrl : null,
            "dismissCallback": null,
            "reDisplay": true,
            "displayDeactive": true,
            "logAsError": true,
            "retryAll": true,
            "errorCode": OSF.ErrorStatusCodes.E_TRUSTCENTER_MOE_UNACTIVATED
        });
    },
    _showActivationError: function OSF_OsfControl$_showActivationError(status, msg, buttonTxt, buttonCallback, url, dismissCallback, titleOverride, detailView, errorCode) {
        this._status = status;
        this._errorcode = errorCode;
        var params = {
            "id": this._id,
            "infoType": OSF.InfoType.Error,
            "status": status,
            "title": titleOverride || Strings.OsfRuntime.L_AgaveErrorTile_TXT,
            "description": msg,
            "buttonTxt": buttonTxt || Strings.OsfRuntime.L_OkButton_TXT,
            "buttonCallback": buttonCallback || null,
            "url": url || null,
            "dismissCallback": dismissCallback || null,
            "reDisplay": !dismissCallback ? true : false,
            "displayDeactive": true,
            "detailView": detailView ? true : false,
            "logAsError": true,
            "errorCode": errorCode ? errorCode : 0
        };
        this._contextActivationMgr.displayNotification(params);
        this._contextActivationMgr.raiseOsfControlStatusChange(this);
    },
    _showActivationWarning: function OSF_OsfControl$_showActivationWarning(status, msg, buttonTxt, buttonCallback, url, dismissCallback, errorCode) {
        this._status = status;
        this._errorcode = errorCode;
        var params = {
            "id": this._id,
            "infoType": OSF.InfoType.Warning,
            "status": status,
            "title": Strings.OsfRuntime.L_AgaveWarningTitle_TXT,
            "description": msg,
            "buttonTxt": buttonTxt || Strings.OsfRuntime.L_OkButton_TXT,
            "buttonCallback": buttonCallback || null,
            "url": url || null,
            "dismissCallback": dismissCallback || null,
            "reDisplay": !dismissCallback ? true : false,
            "displayDeactive": true,
            "logAsError": true,
            "errorCode": errorCode ? errorCode : 0
        };
        this._contextActivationMgr.displayNotification(params);
        this._contextActivationMgr.raiseOsfControlStatusChange(this);
    },
    _refresh: function OSF_OsfControl$_refresh(context) {
        this.deActivate();
        this.activate(context);
    },
    _activateAgavesBlockedBySandboxNotSupport: function OSF_OsfControl$_activateAgavesBlockedBySandboxNotSupport(contextActivationMgr) {
        contextActivationMgr.activateAgavesBlockedBySandboxNotSupport();
        var osfLocalStorage = OSF.OUtil.getLocalStorage();
        if (osfLocalStorage)
            osfLocalStorage.setItem(OSF.Constants.IgnoreSandBoxSupport, "true");
    },
    _setCachedFlag: function OSF_OsfControl$_setCachedFlag(cacheKey) {
        this._contextActivationMgr._setCachedFlag(cacheKey);
    },
    _getCachedFlag: function OSF_OsfControl$_getCachedFlag(cacheKey) {
        return this._contextActivationMgr._getCachedFlag(cacheKey);
    },
    _deleteCachedFlag: function OSF_OsfControl$_deleteCachedFlag(cacheKey) {
        this._contextActivationMgr._deleteCachedFlag(cacheKey);
    },
    _addETokenAsQueryParameter: function OSF_OsfControl$_addETokenAsQueryParameter(iframeUrl) {
        var aElement = document.createElement('a');
        aElement.href = iframeUrl;
        var etoken = this.getEToken();
        var etokenQueryString = OSF.Constants.ETokenParameterName + "=" + encodeURIComponent(OSF.OUtil.encodeBase64(etoken));
        var queryString = aElement.search.length > 1 ? aElement.search.substr(1) + "&" : "";
        aElement.search = queryString + etokenQueryString;
        var modifiedUrl = aElement.href;
        aElement = null;
        return modifiedUrl;
    },
    _activate: function OSF_OsfControl$_activate(frame, iframeUrl) {
        Telemetry.AppLoadTimeHelper.PageStart(this._telemetryContext);
        iframeUrl = this._addETokenAsQueryParameter(iframeUrl);
        var cacheKey = this._contextActivationMgr.getClientId() + "_" + this._contextActivationMgr.getDocUrl() + "_" + this._id;
        this._conversationId = OSF.OUtil.getFrameNameAndConversationId(cacheKey, frame);
        var addHostInfoAsQueryParam = function OSF_OsfControl__activate$addHostInfoAsQueryParam(url, hostInfoValue) {
            url = url.trim() || '';
            var questionMark = "?";
            var hostInfo = "_host_Info=";
            var ampHostInfo = "&_host_Info=";
            var fragmentSeparator = "#";
            var urlParts = url.split(fragmentSeparator);
            var urlWithoutFragment = urlParts.shift();
            var fragment = urlParts.join(fragmentSeparator);
            var querySplits = urlWithoutFragment.split(questionMark);
            var urlWithoutFragmentWithHostInfo;
            if (querySplits.length > 1) {
                urlWithoutFragmentWithHostInfo = urlWithoutFragment + ampHostInfo + hostInfoValue;
            }
            else if (querySplits.le > 0) {
                urlWithoutFragmentWithHostInfo = urlWithoutFragment + questionMark + ampHostInfo + hostInfoValue;
            }
            return [urlWithoutFragmentWithHostInfo, fragmentSeparator, fragment].join('');
        };
        var hostInfoVals = [
            this._contextActivationMgr._hostType,
            this._contextActivationMgr._hostPlatform,
            this._contextActivationMgr._hostSpecificFileVersion,
            this._contextActivationMgr._appUILocale,
            this.getCorrelationId()
        ];
        var hostInfo = hostInfoVals.join("|");
        var newUrl = addHostInfoAsQueryParam(iframeUrl, hostInfo);
        var hrefUrl = window.location.href;
        newUrl = OSF.OUtil.addXdmInfoAsHash(newUrl, this._conversationId + "|" + this._id + "|" + (hrefUrl.indexOf("$") >= 0 ? encodeURIComponent(hrefUrl) : hrefUrl));
        newUrl = OSF.OUtil.addSerializerVersionAsHash(newUrl, OSF.SerializerVersion.Browser);
        this._contextActivationMgr._getServiceEndPoint().registerConversation(this._conversationId, newUrl, this._appDomains);
        this._pageIsReadyTimerExpired = false;
        OSF.OUtil.addEventListener(frame, "load", this._iframeOnLoadDelegate);
        frame.setAttribute("src", newUrl);
        this._div.appendChild(frame);
        this._status = OSF.OsfControlStatus.Activated;
        OSF.OUtil.writeProfilerMark(OSF.OsfControlActivationPerfMarker.ActivationEnd);
        this._contextActivationMgr.raiseOsfControlStatusChange(this);
        Telemetry.AppLoadTimeHelper.OfficeJSStartToLoad(this._telemetryContext);
    },
    _getConversationId: function OSF_OsfControl$_getConversationId() {
        return this._conversationId;
    },
    _doesBrowserSupportRequiredFeatures: function OSF_OsfControl$_doesBrowserSupportRequiredFeatures() {
        var isRequiredFeaturesSupported = false;
        if (Object.defineProperty) {
            try {
                Object.defineProperty({}, "myTestProperty", {
                    get: function () {
                        return this.desc;
                    },
                    set: function (val) {
                        this.desc = val;
                    }
                });
                isRequiredFeaturesSupported = true;
            }
            catch (ex) {
                ;
            }
        }
        return isRequiredFeaturesSupported;
    }
};
OSF.OUtil.setNamespace("Manifest", OSF);
OSF.Manifest.HostApp = function OSF_Manifest_HostApp(appName) {
    this._appName = appName;
    this._minVersion = null;
};
OSF.Manifest.HostApp.prototype = {
    getAppName: function OSF_Manifest_HostApp$getAppName() {
        return this._appName;
    },
    getMinVersion: function OSF_Manifest_HostApp$getMinVersion() {
        return this._minVersion;
    },
    _setMinVersion: function OSF_Manifest_HostApp$_setMinVersion(minVersion) {
        this._minVersion = minVersion;
    }
};
OSF.Manifest.ExtensionSettings = function OSF_Manifest_ExtensionSettings() {
    this._sourceLocations = {};
    this._defaultHeight = null;
    this._defaultWidth = null;
};
OSF.Manifest.ExtensionSettings.prototype = {
    getDefaultHeight: function OSF_Manifest_ExtensionSettings$getDefaultHeight() {
        return this._defaultHeight;
    },
    getDefaultWidth: function OSF_Manifest_ExtensionSettings$getDefaultWidth() {
        return this._defaultWidth;
    },
    getSourceLocations: function OSF_Manifest_ExtensionSettings$getSourceLocations() {
        return this._sourceLocations;
    },
    _addSourceLocation: function OSF_Manifest_ExtensionSettings$_addSourceLocation(locale, value) {
        this._sourceLocations[locale] = value;
    },
    _setDefaultWidth: function OSF_Manifest_ExtensionSettings$_setDefaultWidth(defaultWidth) {
        this._defaultWidth = defaultWidth;
    },
    _setDefaultHeight: function OSF_Manifest_ExtensionSettings$_setDefaultHeight(defaultHeight) {
        this._defaultHeight = defaultHeight;
    }
};
OSF.Manifest.Manifest = function OSF_Manifest_Manifest(para, uiLocale) {
    this._UILocale = uiLocale || "en-us";
    if (typeof (para) !== 'string') {
        para(this);
        return;
    }
    this._displayNames = {};
    this._descriptions = {};
    this._iconUrls = {};
    this._extensionSettings = {};
    this._highResolutionIconUrls = {};
    var versionSpecificDelegate;
    this._xmlProcessor = new OSF.XmlProcessor(para, OSF.ManifestNamespaces["1.1"]);
    if (this._xmlProcessor.selectSingleNode("o:OfficeApp")) {
        versionSpecificDelegate = OSF_Manifest_Manifest_Manifest1_1;
        this._manifestSchemaVersion = OSF.ManifestSchemaVersion["1.1"];
    }
    else {
        this._xmlProcessor = new OSF.XmlProcessor(para, OSF.ManifestNamespaces["1.0"]);
        versionSpecificDelegate = OSF_Manifest_Manifest_Manifest1_0;
        this._manifestSchemaVersion = OSF.ManifestSchemaVersion["1.0"];
    }
    var node = this._xmlProcessor.getDocumentElement();
    this._target = OSF.OUtil.parseEnum(node.getAttribute("xsi:type"), OSF.OfficeAppType);
    var officeAppNode = this._xmlProcessor.selectSingleNode("o:OfficeApp");
    node = this._xmlProcessor.selectSingleNode("o:Id", officeAppNode);
    this._id = this._xmlProcessor.getNodeValue(node);
    node = this._xmlProcessor.selectSingleNode("o:Version", officeAppNode);
    this._version = this._xmlProcessor.getNodeValue(node);
    node = this._xmlProcessor.selectSingleNode("o:ProviderName", officeAppNode);
    this._providerName = this._xmlProcessor.getNodeValue(node);
    node = this._xmlProcessor.selectSingleNode("o:IdIssuer", officeAppNode);
    this._idIssuer = this._parseIdIssuer(node);
    node = this._xmlProcessor.selectSingleNode("o:AlternateId", officeAppNode);
    if (node) {
        this._alternateId = this._xmlProcessor.getNodeValue(node);
    }
    node = this._xmlProcessor.selectSingleNode("o:DefaultLocale", officeAppNode);
    this._defaultLocale = this._xmlProcessor.getNodeValue(node);
    node = this._xmlProcessor.selectSingleNode("o:DisplayName", officeAppNode);
    this._parseLocaleAwareSettings(node, Function.createDelegate(this, this._addDisplayName));
    node = this._xmlProcessor.selectSingleNode("o:Description", officeAppNode);
    this._parseLocaleAwareSettings(node, Function.createDelegate(this, this._addDescription));
    node = this._xmlProcessor.selectSingleNode("o:AppDomains", officeAppNode);
    this._appDomains = this._parseAppDomains(node);
    node = this._xmlProcessor.selectSingleNode("o:IconUrl", officeAppNode);
    if (node) {
        this._parseLocaleAwareSettings(node, Function.createDelegate(this, this._addIconUrl));
    }
    node = this._xmlProcessor.selectSingleNode("o:Signature", officeAppNode);
    if (node) {
        this._signature = this._xmlProcessor.getNodeValue(node);
    }
    this._parseExtensionSettings();
    node = this._xmlProcessor.selectSingleNode("o:Permissions", officeAppNode);
    this._permissions = this._parsePermission(node);
    versionSpecificDelegate.apply(this);
    function OSF_Manifest_Manifest_Manifest1_0() {
        var node = this._xmlProcessor.selectSingleNode("o:Capabilities", officeAppNode);
        var nodes = this._xmlProcessor.selectNodes("o:Capability", node);
        this._capabilities = this._parseCapabilities(nodes);
    }
    function OSF_Manifest_Manifest_Manifest1_1() {
        var node = this._xmlProcessor.selectSingleNode("o:Hosts", officeAppNode);
        this._hosts = this._parseHosts(node);
        node = this._xmlProcessor.selectSingleNode("o:Requirements", officeAppNode);
        this._requirements = this._parseRequirements(node);
        node = this._xmlProcessor.selectSingleNode("o:HighResolutionIconUrl", officeAppNode);
        if (node) {
            this._parseLocaleAwareSettings(node, Function.createDelegate(this, this._addHighResolutionIconUrl));
        }
    }
};
OSF.Manifest.Manifest.prototype = {
    getManifestSchemaVersion: function OSF_Manifest_Manifest$getManifestSchemaVersion() {
        return this._manifestSchemaVersion;
    },
    getMarketplaceID: function OSF_Manifest_Manifest$getMarketplaceID() {
        return this._id;
    },
    getMarketplaceVersion: function OSF_Manifest_Manifest$getMarketplaceVersion() {
        return this._version;
    },
    getDefaultLocale: function OSF_Manifest_Manifest$getDefaultLocale() {
        return this._defaultLocale;
    },
    getProviderName: function OSF_Manifest_Manifest$getProviderName() {
        return this._providerName;
    },
    getIdIssuer: function OSF_Manifest_Manifest$getIdIssuer() {
        return this._idIssuer;
    },
    getAlternateId: function OSF_Manifest_Manifest$getAlternateId() {
        return this._alternateId;
    },
    getSignature: function OSF_Manifest_Manifest$getSignature() {
        return this._signature;
    },
    getCapabilities: function OSF_Manifest_Manifest$getCapabilities() {
        return this._capabilities;
    },
    getDisplayName: function OSF_Manifest_Manifest$getDisplayName(locale) {
        return this._displayNames[locale];
    },
    getDefaultDisplayName: function OSF_Manifest_Manifest$getDefaultDisplayName() {
        return this._getDefaultValue(this._displayNames);
    },
    getDescription: function OSF_Manifest_Manifest$getDescription(locale) {
        return this._descriptions[locale];
    },
    getDefaultDescription: function OSF_Manifest_Manifest$getDefaultDescription() {
        return this._getDefaultValue(this._descriptions);
    },
    getIconUrl: function OSF_Manifest_Manifest$getIconUrl(locale) {
        return this._iconUrls[locale];
    },
    getDefaultIconUrl: function OSF_Manifest_Manifest$getDefaultIconUrl() {
        return this._getDefaultValue(this._iconUrls);
    },
    getSourceLocation: function OSF_Manifest_Manifest$getSourceLocation(locale, formFactor) {
        var extensionSetting = this._getExtensionSetting(formFactor);
        var sourceLocations = extensionSetting.getSourceLocations();
        return sourceLocations[locale];
    },
    getDefaultSourceLocation: function OSF_Manifest_Manifest$getDefaultSourceLocation(formFactor) {
        var extensionSetting = this._getExtensionSetting(formFactor);
        var sourceLocations = extensionSetting.getSourceLocations();
        return this._getDefaultValue(sourceLocations);
    },
    getDefaultWidth: function OSF_Manifest_Manifest$getDefaultWidth(formFactor) {
        var extensionSetting = this._getExtensionSetting(formFactor);
        return extensionSetting.getDefaultWidth();
    },
    getDefaultHeight: function OSF_Manifest_Manifest$getDefaultHeight(formFactor) {
        var extensionSetting = this._getExtensionSetting(formFactor);
        return extensionSetting.getDefaultHeight();
    },
    getTarget: function OSF_Manifest_Manifest$getTarget() {
        return this._target;
    },
    getPermission: function OSF_Manifest_Manifest$getPermission() {
        return this._permissions;
    },
    hasPermission: function OSF_Manifest_Manifest$hasPermission(permissionNeeded) {
        return (this._permissions & permissionNeeded) === permissionNeeded;
    },
    getHosts: function OSF_Manifest_Manifest$getHosts() {
        return this._hosts;
    },
    getRequirements: function OSF_Manifest_Manifest$getRequirements() {
        return this._requirements;
    },
    getHighResolutionIconUrl: function OSF_Manifest_Manifest$getHighResolutionIconUrl(locale) {
        return this._highResolutionIconUrls[locale];
    },
    getAppDomains: function OSF_Manifest_Manifest$getAppDomains() {
        return this._appDomains;
    },
    _getDefaultValue: function OSF_Manifest_Manifest$_getDefaultValue(obj) {
        var locale;
        if (typeof obj[this._UILocale] == "undefined") {
            if (typeof obj[this._defaultLocale] == "undefined") {
                for (var p in obj) {
                    locale = p;
                    break;
                }
            }
            else {
                locale = this._defaultLocale;
            }
        }
        else {
            locale = this._UILocale;
        }
        return obj[locale];
    },
    _getExtensionSetting: function OSF_Manifest_Manifest$_getExtensionSetting(formFactor) {
        var extensionSetting;
        if (typeof this._extensionSettings[formFactor] != "undefined") {
            extensionSetting = this._extensionSettings[formFactor];
        }
        else {
            for (var p in this._extensionSettings) {
                extensionSetting = this._extensionSettings[p];
                break;
            }
        }
        return extensionSetting;
    },
    _addDisplayName: function OSF_Manifest_Manifest$_addDisplayName(locale, value) {
        this._displayNames[locale] = value;
    },
    _addDescription: function OSF_Manifest_Manifest$_addDescription(locale, value) {
        this._descriptions[locale] = value;
    },
    _addIconUrl: function OSF_Manifest_Manifest$_addIconUrl(locale, value) {
        this._iconUrls[locale] = value;
    },
    _parseLocaleAwareSettings: function OSF_Manifest_Manifest$_parseLocaleAwareSettings(localeAwareNode, addCallback) {
        if (!localeAwareNode) {
            throw OsfMsAjaxFactory.msAjaxError.argument("Manifest");
        }
        var defaultValue = localeAwareNode.getAttribute("DefaultValue");
        addCallback(this._defaultLocale, defaultValue);
        var overrideNodes = this._xmlProcessor.selectNodes("o:Override", localeAwareNode);
        if (overrideNodes) {
            var len = overrideNodes.length;
            for (var i = 0; i < len; i++) {
                var node = overrideNodes[i];
                var locale = node.getAttribute("Locale");
                var value = node.getAttribute("Value");
                addCallback(locale, value);
            }
        }
    },
    _parseBooleanNode: function OSF_Manifest_Manifest$_parseBooleanNode(node) {
        if (!node) {
            return false;
        }
        else {
            var value = this._xmlProcessor.getNodeValue(node).toLowerCase();
            return value === "true" || value === "1";
        }
    },
    _parseIdIssuer: function OSF_Manifest_Manifest$_parseIdIssuer(node) {
        if (!node) {
            return OSF.ManifestIdIssuer.Custom;
        }
        else {
            var value = this._xmlProcessor.getNodeValue(node);
            return OSF.OUtil.parseEnum(value, OSF.ManifestIdIssuer);
        }
    },
    _parseCapabilities: function OSF_Manifest_Manifest$_parseCapabilities(nodes) {
        var capabilities = [];
        var capability;
        for (var i = 0; i < nodes.length; i++) {
            var node = nodes[i];
            capability = node.getAttribute("Name");
            capability = OSF.OUtil.parseEnum(capability, OSF.Capability);
            capabilities.push(capability);
        }
        return capabilities;
    },
    _parsePermission: function OSF_Manifest_Manifest$_parsePermission(capabilityNode) {
        if (!capabilityNode) {
            throw OsfMsAjaxFactory.msAjaxError.argument("Manifest");
        }
        var value = this._xmlProcessor.getNodeValue(capabilityNode);
        return OSF.OUtil.parseEnum(value, OSF.OsfControlPermission);
    },
    _parseExtensionSettings: function OSF_Manifest_Manifest$_parseExtensionSettings() {
        var settings;
        var settingNode;
        var node;
        for (var formFactor in OSF.FormFactor) {
            var officeAppNode = this._xmlProcessor.selectSingleNode("o:OfficeApp");
            settingNode = this._xmlProcessor.selectSingleNode("o:" + OSF.FormFactor[formFactor], officeAppNode);
            if (settingNode) {
                settings = new OSF.Manifest.ExtensionSettings();
                node = this._xmlProcessor.selectSingleNode("o:SourceLocation", settingNode);
                var addSourceLocation = function (locale, value) {
                    settings._addSourceLocation(locale, value);
                };
                this._parseLocaleAwareSettings(node, addSourceLocation);
                node = this._xmlProcessor.selectSingleNode("o:RequestedWidth", settingNode);
                if (node) {
                    settings._setDefaultWidth(this._xmlProcessor.getNodeValue(node));
                }
                node = this._xmlProcessor.selectSingleNode("o:RequestedHeight", settingNode);
                if (node) {
                    settings._setDefaultHeight(this._xmlProcessor.getNodeValue(node));
                }
                this._extensionSettings[formFactor] = settings;
            }
        }
        if (!settings) {
            throw OsfMsAjaxFactory.msAjaxError.argument("Manifest");
        }
    },
    _parseHosts: function OSF_Manifest_Manifest$_parseHosts(hostsNode) {
        var targetHosts = [];
        if (hostsNode) {
            var hostNodes = this._xmlProcessor.selectNodes("o:Host", hostsNode);
            for (var i = 0; i < hostNodes.length; i++) {
                targetHosts.push(hostNodes[i].getAttribute("Name"));
            }
        }
        return targetHosts;
    },
    _parseRequirements: function OSF_Manifest_Manifest$_parseRequirements(requirementsNode) {
        var requirements = {
            sets: [],
            methods: []
        };
        if (requirementsNode) {
            var setsNode = this._xmlProcessor.selectSingleNode("o:Sets", requirementsNode);
            requirements.sets = this._parseSets(setsNode);
            var methodsNode = this._xmlProcessor.selectSingleNode("o:Methods", requirementsNode);
            requirements.methods = this._parseMethods(methodsNode);
        }
        return requirements;
    },
    _parseSets: function OSF_Manifest_Manifest$_parseSets(setsNode) {
        var sets = [];
        if (setsNode) {
            var defaultVersion = setsNode.getAttribute("DefaultMinVersion");
            var setNodes = this._xmlProcessor.selectNodes("o:Set", setsNode);
            for (var i = 0; i < setNodes.length; i++) {
                var setNode = setNodes[i];
                var overrideVersion = setNode.getAttribute("MinVersion");
                sets.push({
                    name: setNode.getAttribute("Name"),
                    version: overrideVersion || defaultVersion
                });
            }
        }
        return sets;
    },
    _parseMethods: function OSF_Manifest_Manifest$_parseMethods(methodsNode) {
        var methods = [];
        if (methodsNode) {
            var methodNodes = this._xmlProcessor.selectNodes("o:Method", methodsNode);
            for (var i = 0; i < methodNodes.length; i++) {
                methods.push(methodNodes[i].getAttribute("Name"));
            }
        }
        return methods;
    },
    _parseAppDomains: function OSF_Manifest_Manifest$_parseAppDomains(appDomainsNode) {
        var appDomains = [];
        if (appDomainsNode) {
            var appDomainNodes = this._xmlProcessor.selectNodes("o:AppDomain", appDomainsNode);
            for (var i = 0; i < appDomainNodes.length; i++) {
                appDomains.push(this._xmlProcessor.getNodeValue(appDomainNodes[i]));
            }
        }
        return appDomains;
    },
    _addHighResolutionIconUrl: function OSF_Manifest_Manifest$_addHighResolutionIconUrl(locale, url) {
        this._highResolutionIconUrls[locale] = url;
    },
    _isAddinCommandsManifest: function OSF_Manifest_Manifest$_isAddinCommandsManifest(currentHostType) {
        var officeAppNode = this._xmlProcessor.selectSingleNode("o:OfficeApp");
        if (officeAppNode !== null && this.getManifestSchemaVersion() === OSF.ManifestSchemaVersion["1.1"]) {
            var overridesNode = this._xmlProcessor.selectSingleNode("ov:VersionOverrides", officeAppNode);
            if (overridesNode !== null) {
                var hostsNode = this._xmlProcessor.selectSingleNode("ov:Hosts", overridesNode);
                if (hostsNode !== null) {
                    var hostNodes = this._xmlProcessor.selectNodes("ov:Host", hostsNode);
                    if (hostNodes !== null) {
                        var len = hostNodes.length;
                        for (var i = 0; i < len; i++) {
                            var hostType = hostNodes[i].getAttribute("xsi:type");
                            if (hostType === currentHostType) {
                                var formFactorNode = this._xmlProcessor.selectSingleNode("ov:DesktopFormFactor", hostNodes[i]);
                                if (formFactorNode != null) {
                                    return true;
                                }
                                else {
                                    return false;
                                }
                            }
                        }
                    }
                }
            }
        }
        return false;
    }
};
OSF.OUtil.setNamespace("AppSpecificSetup", OSF);
OSF.ContextActivationManager = function OSF_ContextActivationManager(params) {
    OSF.OUtil.validateParamObject(params, {
        "appName": { type: Number, mayBeNull: false },
        "appVersion": { type: String, mayBeNull: false },
        "clientMode": { type: Number, mayBeNull: false },
        "appUILocale": { type: String, mayBeNull: false },
        "dataLocale": { type: String, mayBeNull: false },
        "osfOmexBaseUrl": { type: String, mayBeNull: true },
        "devCatalogUrl": { type: String, mayBeNull: true },
        "oneDriveCatalogBaseApiUrl": { type: String, mayBeNull: true },
        "spBaseUrl": { type: String, mayBeNull: true },
        "docUrl": { type: String, mayBeNull: true },
        "hostControl": { type: Object, mayBeNull: true },
        "pageBaseUrl": { type: String, mayBeNull: true },
        "lcid": { type: String, mayBeNull: true },
        "formFactor": { type: String, mayBeNull: true },
        "controlStatusChanged": { type: Object, mayBeNull: true },
        "notifyHost": { type: Object, mayBeNull: true },
        "allowExternalMarketplace": { type: Boolean, mayBeNull: true },
        "localizedScriptsUrl": { type: String, mayBeNull: true },
        "localizedImagesUrl": { type: String, mayBeNull: true },
        "localizedStylesUrl": { type: String, mayBeNull: true },
        "localizedResourcesUrl": { type: String, mayBeNull: true },
        "trustAgaves": { type: Boolean, mayBeNull: true },
        "enableMyOrg": { type: Boolean, mayBeNull: true },
        "enableMyApps": { type: Boolean, mayBeNull: true },
        "enableDevCatalog": { type: Boolean, mayBeNull: true },
        "enableOneDriveCatalog": { type: Boolean, mayBeNull: true },
        "enableUploadFileDevCatalog": { type: Boolean, mayBeNull: true },
        "omexForceAnonymous": { type: Boolean, mayBeNull: true },
        "userNameHashCode": { type: Number, mayBeNull: true }
    }, null);
    this._osfOmexBaseUrl = params.osfOmexBaseUrl;
    this._devCatalogUrl = params.devCatalogUrl;
    this._oneDriveCatalogBaseApiUrl = params.oneDriveCatalogBaseApiUrl;
    this._spBaseUrl = params.spBaseUrl;
    this._myOrgCatalogUrl = null;
    this._enableMyOrg = params.enableMyOrg || false;
    this._enableMyApps = params.enableMyApps || false;
    this._enableDevCatalog = params.enableDevCatalog || false;
    this._enableOneDriveCatalog = params.enableOneDriveCatalog || false;
    this._enableUploadFileDevCatalog = params.enableUploadFileDevCatalog || false;
    this._appName = params.appName;
    this._appVersion = params.appVersion;
    this._clientMode = params.clientMode;
    this._appUILocale = params.appUILocale;
    this._dataLocale = params.dataLocale;
    this._docUrl = params.docUrl;
    this._hostControl = params.hostControl;
    this._formFactor = params.formFactor || OSF.FormFactor.Default;
    this._pageBaseUrl = params.pageBaseUrl;
    this._lcid = params.lcid;
    this._controlStatusChanged = params.controlStatusChanged;
    this._notifyHost = params.notifyHost;
    this._allowExternalMarketplace = params.allowExternalMarketplace;
    this._localizedScriptsUrl = params.localizedScriptsUrl;
    this._localizedImagesUrl = params.localizedImagesUrl;
    this._localizedStylesUrl = params.localizedStylesUrl;
    this._localizedResourcesUrl = params.localizedResourcesUrl;
    this._autoTrusted = params.trustAgaves;
    this._omexForceAnonymous = params.omexForceAnonymous || false;
    this._userNameHashCode = params.userNameHashCode || 0;
    if (this._pageBaseUrl && this._pageBaseUrl.charAt(this._pageBaseUrl.length - 1) !== '/') {
        this._pageBaseUrl = this._pageBaseUrl + '/';
    }
    if (this._localizedResourcesUrl && this._localizedResourcesUrl.charAt(this._localizedResourcesUrl.length - 1) !== '/') {
        this._localizedResourcesUrl = this._localizedResourcesUrl + '/';
    }
    this._clientId = OSF.OUtil.getUniqueId();
    this._cachedOsfControls = {};
    this._iframeAttributeBag = {};
    this._serviceEndPoint = null;
    this._serviceEndPointInternal = null;
    this._internalConversationId = null;
    this._iframeProxies = {};
    this._iframeProxyCount = 0;
    this._iframeNamePrefix = "__officeExtensionProxy";
    this._webUrl = null;
    this._wsa = null;
    this._insertDialogDiv = null;
    this._hasPreloadedOfficeJs = false;
    this._contextMgrCorrelationId = OSF.OUtil.Guid.generateNewGuid();
    Telemetry.RuntimeTelemetryHelper.LogCommonMessage("_contextMgrCorrelationId :" + this._contextMgrCorrelationId, 0);
    this._hostType = null;
    this._hostPlatform = null;
    this._hostSpecificFileVersion = null;
    this._requirementsChecker = new OSF.RequirementsChecker();
    if (this._osfOmexBaseUrl) {
        var baseUrlWithoutProtocol;
        var protocolSeparatorIndex = this._osfOmexBaseUrl.indexOf(OSF.Constants.ProtocolSeparator);
        if (protocolSeparatorIndex >= 0) {
            baseUrlWithoutProtocol = this._osfOmexBaseUrl.substr(protocolSeparatorIndex);
        }
        else {
            baseUrlWithoutProtocol = OSF.Constants.ProtocolSeparator + this._osfOmexBaseUrl;
        }
        var omexGatedBaseUrl = OSF.Constants.Https + baseUrlWithoutProtocol;
        var omexUngatedBaseUrl = OSF.Constants.Https + baseUrlWithoutProtocol;
        if (OSF.OUtil.getQueryStringParamValue(window.location.search, OSF.Constants.OmexForceAnonymousParamName).toLowerCase() == OSF.Constants.OmexForceAnonymousParamValue.toLowerCase()) {
            this._omexAuthNStatus = OSF.OmexAuthNStatus.Anonymous;
            this._omexForceAnonymous = true;
        }
        else {
            this._omexAuthNStatus = OSF.OmexAuthNStatus.NotAttempted;
        }
        this._omexGatedWSProxy = { "proxyUrl": omexGatedBaseUrl + OSF.Constants.OmexGatedServiceExtension, "proxyName": "__omexExtensionGatedProxy", "isReady": false, "clientEndPoint": null, "pendingCallbacks": [] };
        this._omexWSProxy = { "proxyUrl": omexUngatedBaseUrl + OSF.Constants.OmexUnGatedServiceExtension, "proxyName": "__omexExtensionProxy", "isReady": false, "clientEndPoint": null, "pendingCallbacks": [] };
        this._omexAnonymousWSProxy = { "proxyUrl": omexGatedBaseUrl + OSF.Constants.OmexAnonymousServiceExtension, "proxyName": "__omexExtensionAnonymousProxy", "isReady": false, "clientEndPoint": null, "pendingCallbacks": [] };
        this._omexBillingMarket = null;
        this._omexEndPointBaseUrl = omexUngatedBaseUrl;
    }
    OSF.OsfManifestManager._setUILocale(this._appUILocale);
    var me = this;
    var getAppContextAsync = function OSF_ContextActivationManager$getAppContextAsync(contextId, gotAppContext) {
        var e = Function._validateParams(arguments, [{ name: "contextId", type: String, mayBeNull: false },
            { name: "gotAppContext", type: Function, mayBeNull: false }
        ]);
        if (e) {
            Telemetry.RuntimeTelemetryHelper.LogExceptionTag("Parameter validation error in getAppContextAsync.", e, null, 0x008c7450);
            throw e;
        }
        var osfControl = me.getOsfControl(contextId);
        if (!osfControl) {
            OsfMsAjaxFactory.msAjaxDebug.trace("osfControl for the given ID doesn't exist.");
            Telemetry.RuntimeTelemetryHelper.LogExceptionTag("Cannot get osfControl with given ID.", null, null, 0x008c7451);
            throw OsfMsAjaxFactory.msAjaxError.argument("contextId");
        }
        else {
            Telemetry.AppLoadTimeHelper.OfficeJSLoaded(osfControl._telemetryContext);
            var eToken = osfControl.getEToken();
            var minor = 0;
            if (me._hostSpecificFileVersion && me._hostSpecificFileVersion.indexOf(".") != -1) {
                var versions = me._hostSpecificFileVersion.split(".");
                var minorVersionString = versions[1];
                if (!isNaN(minorVersionString)) {
                    minor = parseInt(versions[1]);
                }
            }
            var requirements = me._requirementsChecker._getSupportedSet();
            var appContext = new OSF.OfficeAppContext(contextId, me._appName, me._appVersion, me._appUILocale, me._dataLocale, me._docUrl || window.location.href, me._clientMode, osfControl.getSettings(), osfControl.getReason(), osfControl.getOsfControlType(), eToken, osfControl._appCorrelationId, osfControl._appInstanceId, false, false, minor, requirements);
            gotAppContext(appContext);
            osfControl._pageStatus = OSF.OsfControlPageStatus.Ready;
            if (osfControl._pageIsReadyTimerExpired) {
                Telemetry.RuntimeTelemetryHelper.LogExceptionTag("App attempted to retrieve context after app activation error has occured.", null, osfControl.getCorrelationId(), 0x007d4283);
            }
            if (osfControl._contextActivationMgr._ErrorUXHelper) {
                osfControl._contextActivationMgr._ErrorUXHelper.removeProgressDiv(osfControl._div, osfControl._id);
            }
            if (osfControl._timer) {
                window.clearTimeout(osfControl._timer);
                osfControl._timer = null;
            }
            var notificationConversationId = osfControl._conversationId + OSF.SharedConstants.NotificationConversationIdSuffix;
            osfControl._agaveEndPoint = Microsoft.Office.Common.XdmCommunicationManager.connect(notificationConversationId, osfControl._frame.contentWindow, osfControl._iframeUrl);
            Telemetry.AppLoadTimeHelper.ActivationEnd(osfControl._telemetryContext);
        }
    };
    var notifyHost = function OSF_ContextActivationManager$notifyHost(params) {
        if (!params || params.length != 2) {
            OsfMsAjaxFactory.msAjaxDebug.trace("ContextActivationManager_notifyHost params is wrong.");
        }
        var contextId = params[0];
        var actionId = params[1];
        var osfControl = me.getOsfControl(contextId);
        if (!osfControl) {
            OsfMsAjaxFactory.msAjaxDebug.trace("osfControl for the given ID doesn't exist.");
        }
        else {
            if (osfControl._contextActivationMgr._notifyHost) {
                osfControl._contextActivationMgr._notifyHost(contextId, actionId);
            }
            else {
                OsfMsAjaxFactory.msAjaxDebug.trace("No notifyHost provided by the host.");
            }
        }
    };
    var openWindowInHost = function OSF_ContextActivationManager$openWindowInHost(params) {
        window.open(params.strUrl, params.strWindowName, params.strWindowFeatures);
    };
    var getEntitlementsForInsertDialog = function OSF_ContextActivationManager$getEntitlementsForInsertDialog(params, onGetEntitlements) {
        if (!params.hasOwnProperty("fromInsertDialog") || params.fromInsertDialog) {
            if (me._insertDialogDiv != null && me._insertDialogDiv.childNodes.length == 2) {
                me._insertDialogDiv.removeChild(me._insertDialogDiv.lastChild);
            }
        }
        var referenceInUse;
        if (params.storeType == OSF.StoreTypeEnum.MarketPlace) {
            referenceInUse = {
                "storeType": OSF.StoreType.OMEX,
                "storeLocator": me._osfOmexBaseUrl
            };
        }
        else if (params.storeType == OSF.StoreTypeEnum.Catalog) {
            referenceInUse = {
                "storeType": OSF.StoreType.SPCatalog,
                "storeLocator": me._myOrgCatalogUrl
            };
        }
        else if (params.storeType == OSF.StoreTypeEnum.OneDrive) {
            referenceInUse = {
                "storeType": OSF.StoreType.OneDrive,
                "storeLocator": me._oneDriveCatalogBaseApiUrl
            };
        }
        var context = {
            "assetId": "",
            "contentMarket": "",
            "anonymous": false,
            "clientEndPoint": null,
            "clearCache": false || params.refresh,
            "clearKilledApps": false,
            "referenceInUse": referenceInUse,
            "hostType": me._hostType
        };
        var omexAuthenticatedConnectTries = 1;
        if (params && params.storeType == OSF.StoreTypeEnum.MarketPlace) {
            context.clientVersion = me._getClientVersionForOmex();
            if (context.clientVersion) {
                context.clientName = me._getClientNameForOmex();
                context.appVersion = me._getAppVersionForOmex();
            }
            var onGetOmexEntitlementsCompleted = function OSF_ContextActivationManager_getEntitlementsForInsertDialog$onGetOmexEntitlementsCompleted(asyncResult) {
                if (asyncResult.statusCode === OSF.ProxyCallStatusCode.Succeeded && asyncResult.value) {
                    var reference = asyncResult.context.referenceInUse;
                    var entitlements = asyncResult.value.entitlements;
                    var entitlementCount = entitlements.length;
                    if (entitlementCount === 0) {
                        onGetEntitlements({ "errorCode": OSF.InvokeResultCode.S_OK });
                    }
                    var entitlement;
                    var params = {};
                    var result = [];
                    for (var i = 0; i < entitlementCount; i++) {
                        entitlement = entitlements[i];
                        if (params[entitlement.contentMarket]) {
                            params[entitlement.contentMarket] += "," + entitlement.assetId;
                        }
                        else {
                            params[entitlement.contentMarket] = entitlement.assetId;
                        }
                    }
                    omexAuthenticatedConnectTries = 1;
                    var appCount = 0;
                    var onGetOmexAppDetailsCompleted = function OSF_ContextActivationManager_getEntitlementsForInsertDialog$onGetOmexAppDetailsCompleted(asyncResult) {
                        if (asyncResult.statusCode === OSF.ProxyCallStatusCode.Succeeded && asyncResult.value && asyncResult.value.length && asyncResult.value.length > 0) {
                            var galleryItems = asyncResult.value;
                            if (!asyncResult.cached) {
                                _omexDataProvider.SetAppDetailsCache(context, galleryItems);
                            }
                            if (context.appDetails && context.appDetails.length > 0) {
                                for (var i = 0; i < context.appDetails.length; i++) {
                                    galleryItems.push(context.appDetails[i]);
                                }
                            }
                            var requirementsChecker = me.getRequirementsChecker();
                            for (var k = 0; k < galleryItems.length; k++) {
                                if (requirementsChecker.isEntitlementFromOmexSupported(galleryItems[k])) {
                                    var galleryItem = [];
                                    galleryItem.push(galleryItems[k].name);
                                    galleryItem.push(galleryItems[k].assetId);
                                    galleryItem.push(galleryItems[k].description);
                                    galleryItem.push(OSF.OUtil.getTargetType(galleryItems[k].appSubType));
                                    galleryItem.push(OSF.OUtil.normalizeAppVersion(galleryItems[k].version));
                                    galleryItem.push(galleryItems[k].assetId);
                                    galleryItem.push(OSF.StoreType.OMEX);
                                    galleryItem.push(parseInt(galleryItems[k].defaultWidth));
                                    galleryItem.push(parseInt(galleryItems[k].defaultHeight));
                                    galleryItem.push(galleryItems[k].iconUrl);
                                    galleryItem.push(galleryItems[k].provider);
                                    galleryItem.push(asyncResult.context.contentMarket);
                                    galleryItem.push(OSF.StoreType.OMEX);
                                    result.push(galleryItem);
                                }
                                appCount++;
                            }
                        }
                        if (appCount === entitlementCount) {
                            var response = { "value": result, "errorCode": OSF.InvokeResultCode.S_OK };
                            onGetEntitlements(response);
                        }
                    };
                    var createUngatedOmexProxyCompleted = function OSF_ContextActivationManager_getEntitlementsForInsertDialog$createUngatedOmexProxyCompleted(clientEndPoint) {
                        if (clientEndPoint) {
                            for (var cm in params) {
                                context.anonymous = false;
                                context.clientEndPoint = clientEndPoint;
                                context.assetId = params[cm];
                                context.contentMarket = cm;
                                var contextCopy = OSF.OUtil.shallowCopy(context);
                                OSF.OsfManifestManager.getOmexAppDetailsAsync(contextCopy, Function.createDelegate(me, onGetOmexAppDetailsCompleted));
                            }
                        }
                        else {
                            if (omexAuthenticatedConnectTries < OSF.Constants.AuthenticatedConnectMaxTries) {
                                omexAuthenticatedConnectTries++;
                                me._createOmexProxy(me._omexWSProxy, createUngatedOmexProxyCompleted);
                            }
                            else {
                                var currentUrl = window.location.href;
                                var signInRedirectUrl = me._osfOmexBaseUrl + OSF.Constants.SignInRedirectUrl + encodeURIComponent(currentUrl);
                                window.location.assign(signInRedirectUrl);
                            }
                        }
                    };
                    var proxyRequired = true;
                    if (_omexDataProvider.AppDetailsCached(context, params)) {
                        try {
                            context.clientEndPoint = me._omexWSProxy.clientEndPoint;
                            for (var cm in params) {
                                context.anonymous = false;
                                context.assetId = params[cm];
                                context.contentMarket = cm;
                                OSF.OsfManifestManager.getOmexAppDetailsAsync(context, Function.createDelegate(me, onGetOmexAppDetailsCompleted));
                            }
                            proxyRequired = false;
                        }
                        catch (e) {
                        }
                    }
                    if (proxyRequired) {
                        me._createOmexProxy(me._omexWSProxy, createUngatedOmexProxyCompleted);
                    }
                }
            };
            var checkMyAppsCompleted = function OSF_ContextActivationManager_getEntitlementsForInsertDialog$checkMyAppsCompleted(asyncResult) {
                if (asyncResult.isReady) {
                    context.anonymous = false;
                    context.clientEndPoint = asyncResult.clientEndPoint;
                    OSF.OsfManifestManager.getOmexEntitlementsAsync(context, Function.createDelegate(me, onGetOmexEntitlementsCompleted));
                }
                else {
                    var response = { "value": null, "errorCode": OSF.InvokeResultCode.E_USER_NOT_SIGNED_IN };
                    onGetEntitlements(response);
                }
            };
            me._omexGatedWSProxy.refresh = params.refresh;
            me.isMyAppsReady(checkMyAppsCompleted);
        }
        else if (params && params.storeType == OSF.StoreTypeEnum.Catalog) {
            var getMyOrgEntitmentsDetailCompleted = function OSF_ContextActivationManager_getEntitlementsForInsertDialog$getMyOrgEntitmentsDetailCompleted(asyncResult) {
                if (asyncResult.context && asyncResult.context.referenceInUse.storeType === OSF.StoreType.SPCatalog) {
                    OSF.OUtil.writeProfilerMark(OSF.OsfOfficeExtensionManagerPerfMarker.GetEntitlementEnd);
                }
                var response;
                if (asyncResult.statusCode === OSF.ProxyCallStatusCode.Succeeded && asyncResult.value) {
                    var entitlements = asyncResult.value.entitlements;
                    var entitlementCount = entitlements.length;
                    var entitlement;
                    var result = [];
                    var supportedEntitlementCount = 0;
                    var requirementsChecker = me.getRequirementsChecker();
                    for (var i = 0; i < entitlementCount; i++) {
                        entitlement = entitlements[i];
                        if (requirementsChecker.isEntitlementFromCorpCatalogSupported(entitlement)) {
                            var galleryItem = [];
                            galleryItem.push(entitlement.Title);
                            galleryItem.push(entitlement.OfficeExtensionID);
                            galleryItem.push(entitlement.OfficeExtensionDescription);
                            galleryItem.push(OSF.OfficeAppType[entitlement.OEType]);
                            galleryItem.push(OSF.OUtil.normalizeAppVersion(entitlement.OfficeExtensionVersion));
                            galleryItem.push(entitlement.OfficeExtensionID);
                            galleryItem.push(OSF.StoreType.SPCatalog);
                            galleryItem.push(parseInt(entitlement.OfficeExtensionDefaultWidth.toString()));
                            galleryItem.push(parseInt(entitlement.OfficeExtensionDefaultHeight.toString()));
                            galleryItem.push(entitlement.OfficeExtensionIcon);
                            galleryItem.push(entitlement.OEProviderName);
                            galleryItem.push(referenceInUse.storeLocator);
                            galleryItem.push(OSF.StoreType.SPCatalog);
                            result.push(galleryItem);
                            supportedEntitlementCount++;
                        }
                    }
                    if (supportedEntitlementCount > 0) {
                        response = {
                            "value": result,
                            "errorCode": OSF.InvokeResultCode.S_OK
                        };
                    }
                    else {
                        response = {
                            "value": null,
                            "errorCode": OSF.InvokeResultCode.E_CATALOG_NO_APPS
                        };
                    }
                }
                else if (!referenceInUse.storeLocator) {
                    response = {
                        "value": null,
                        "errorCode": OSF.InvokeResultCode.E_CATALOG_NO_APPS
                    };
                }
                else {
                    response = {
                        "value": null,
                        "errorCode": OSF.InvokeResultCode.E_GENERIC_ERROR
                    };
                }
                onGetEntitlements(response);
            };
            context.webUrl = referenceInUse.storeLocator;
            context.osfControl = { "_contextActivationMgr": me, "_telemetryContext": {} };
            context.noTargetType = true;
            OSF.OsfManifestManager.getCorporateCatalogEntitlementsAsync(context, Function.createDelegate(me, getMyOrgEntitmentsDetailCompleted));
        }
        else if (params && params.storeType == OSF.StoreTypeEnum.OneDrive) {
            var getOneDriveEntitmentsDetailCompleted = function (asyncResult) {
                var response;
                if (asyncResult.statusCode === OSF.ProxyCallStatusCode.Succeeded && asyncResult.value) {
                    var result = asyncResult.value.entitlements;
                    response = {
                        "value": result,
                        "errorCode": OSF.InvokeResultCode.S_OK
                    };
                    onGetEntitlements(response);
                }
            };
            OSF.OsfManifestManager.getOneDriveEntitlementsAsync(context, getOneDriveEntitmentsDetailCompleted);
        }
        else {
            OsfMsAjaxFactory.msAjaxDebug.trace("Unknown store type.");
        }
    };
    var invokeSignIn = function OSF_ContextActivationManager$invokeSignIn(params) {
        var currentUrl = window.location.href;
        var signInRedirectUrl = me._osfOmexBaseUrl + OSF.Constants.SignInRedirectUrl + encodeURIComponent(currentUrl);
        window.location.assign(signInRedirectUrl);
    };
    var invokeWindowOpen = function OSF_ContextActivationManager$invokeWindowOpen(params) {
        window.open(params.pageUrl);
    };
    var onClickInsertOsfControl = function (params, callback) {
        OsfMsAjaxFactory.msAjaxDebug.trace("onClickInsertOsfControl!");
        me._notifyHost("0", OSF.AgaveHostAction.InsertAgave, params);
    };
    var onClickInstallOsfControl = function (params, callback) {
        if (!me._enableUploadFileDevCatalog) {
            return;
        }
        OsfMsAjaxFactory.msAjaxDebug.trace("onClickInstallOsfControl!");
        try {
            OSF.OUtil.validateParamObject(params, {
                "manifest": { type: String, mayBeNull: false }
            }, null);
            var parsedManifest = new OSF.Manifest.Manifest(params.manifest, me.getAppUILocale());
            var assetId = parsedManifest.getMarketplaceID();
            var appVersion = parsedManifest.getMarketplaceVersion();
            var dataProvider = OfficeExt.AddinDataProviderFactory.createAddinDataProvider(OSF.StoreType.UploadFileDevCatalog, null);
            if (dataProvider.validateAndCacheNewManifest(me, assetId, appVersion, params.manifest, parsedManifest)) {
                var installParams = {
                    "id": assetId,
                    "targetType": parsedManifest.getTarget(),
                    "appVersion": appVersion,
                    "currentStoreType": OSF.StoreType.UploadFileDevCatalog,
                    "storeId": "developer",
                    "assetId": assetId,
                    "assetStoreId": OSF.StoreType.UploadFileDevCatalog,
                    "width": parseInt(parsedManifest.getDefaultWidth() || "0"),
                    "height": parseInt(parsedManifest.getDefaultHeight() || "0"),
                    "isAddinCommands": parsedManifest._isAddinCommandsManifest(OSF.getManifestHostType(me._hostType))
                };
                OSF.OUtil.validateParamObject(installParams, {
                    "id": { type: String, mayBeNull: false },
                    "targetType": { type: Number, mayBeNull: false },
                    "appVersion": { type: String, mayBeNull: false },
                    "width": { type: Number, mayBeNull: true },
                    "height": { type: Number, mayBeNull: true }
                }, null);
                Telemetry.UploadFileDevCatalogUsageHelper.LogUploadFileDevCatalogUsageAction(this._appCorrelationId, installParams.currentStoreType, installParams.assetId, installParams.appVersion, installParams.targetType, installParams.isAddinCommands, installParams.width, installParams.height);
                me._notifyHost("0", OSF.AgaveHostAction.InsertAgave, installParams);
            }
            else {
                me._notifyHost("0", OSF.AgaveHostAction.NotifyHostError, null);
            }
        }
        catch (ex) {
            OsfMsAjaxFactory.msAjaxDebug.trace("Invalid manifest from addincommands dev catalog: " + ex);
            Telemetry.RuntimeTelemetryHelper.LogExceptionTag("Invalid manifest from addincommands dev catalog.", ex, this._appCorrelationId, 0x008e23c5);
            me._notifyHost("0", OSF.AgaveHostAction.NotifyHostError, null);
            return;
        }
    };
    var onClickRefreshAddinCommands = function (params, callback) {
        OsfMsAjaxFactory.msAjaxDebug.trace("onClickRefreshfAddinCommands!");
        me._notifyHost("0", OSF.AgaveHostAction.RefreshAddinCommands, params);
    };
    var onClickCancelDialog = function (params, callback) {
        OsfMsAjaxFactory.msAjaxDebug.trace("onClickCancelDialog!");
        me._notifyHost("0", OSF.AgaveHostAction.CancelDialog);
        if (me._internalConversationId) {
            me._serviceEndPointInternal.unregisterConversation(me._internalConversationId);
            me._internalConversationId = null;
        }
    };
    var getOmexData = function (params) {
        var context = { "anonymous": null,
            "clientAppInfoReturnType": OSF.ClientAppInfoReturnType.both,
            "clientEndPoint": null,
            "clientName": null,
            "appVersion": null,
            "clientVersion": null,
            "hostType": null,
            "osfControl": { "_omexEntitlement": null },
            "referenceInUse": null
        };
        var proxyRequired = true;
        context.referenceInUse = {
            "id": params.assetId,
            "storeType": OSF.StoreType.OMEX,
            "storeLocator": params.storeId
        };
        context.clientVersion = me._getClientVersionForOmex();
        if (context.clientVersion) {
            context.clientName = me._getClientNameForOmex();
            context.appVersion = me._getAppVersionForOmex();
        }
        context.hostType = me._hostType;
        params.contentMarket = params.storeId;
        params.assetID = params.assetId;
        if (_omexDataProvider.AllCached(context, params)) {
            try {
                context.anonymous = true;
                context.clientEndPoint = {};
                OSF.OsfManifestManager.getOmexKilledAppsAsync(context, onGetOmexKilledAppsCompleted);
                OSF.OsfManifestManager.getOmexAppStateAsync(context, onGetOmexAppStateCompleted);
                proxyRequired = false;
            }
            catch (e) {
            }
        }
        if (proxyRequired) {
            var onGetOmexEntitlementsCompleted = function (asyncResult) {
                if (asyncResult.statusCode === OSF.ProxyCallStatusCode.Succeeded && asyncResult.value) {
                    var reference = asyncResult.context.referenceInUse;
                    var entitlements = asyncResult.value.entitlements;
                    var entitlementCount = entitlements.length;
                    var entitlement;
                    var found = false;
                    _omexDataProvider.SetCustomerId(asyncResult.value.cid);
                    for (var i = 0; i < entitlementCount; i++) {
                        entitlement = entitlements[i];
                        if (entitlement.assetId && reference.id && entitlement.assetId.toLowerCase() === reference.id.toLowerCase()) {
                            found = true;
                            break;
                        }
                    }
                    var context = asyncResult.context;
                    context.osfControl._omexEntitlement = { "contentMarket": null, "hasEntitlement": false, "version": null };
                    if (found) {
                        context.osfControl._omexEntitlement.contentMarket = entitlement.contentMarket;
                        context.osfControl._omexEntitlement.hasEntitlement = true;
                        context.osfControl._omexEntitlement.version = entitlement.version;
                    }
                    OSF.OsfManifestManager.getOmexAppStateAsync(context, onGetOmexAppStateCompleted);
                }
            };
            var onGetOmexManifestAndETokenCompleted = function (asyncResult) {
                if (asyncResult.statusCode === OSF.ProxyCallStatusCode.Succeeded && asyncResult.value) {
                    var manifestAndEToken = asyncResult.value;
                    var clientAppStatus = parseInt(manifestAndEToken.status);
                    var context = asyncResult.context;
                    if (clientAppStatus === OSF.OmexClientAppStatus.OK) {
                        if (!asyncResult.cached) {
                            _omexDataProvider.SetManifestAndETokenCache(context, manifestAndEToken);
                        }
                    }
                }
            };
            var onGetOmexKilledAppsCompleted = function (asyncResult) {
                if (asyncResult.statusCode === OSF.ProxyCallStatusCode.Succeeded && asyncResult.value) {
                    if (!asyncResult.cached) {
                        _omexDataProvider.SetKilledAppsCache(asyncResult.context, asyncResult.value);
                    }
                }
            };
            var onGetOmexAppStateCompleted = function (asyncResult) {
                if (asyncResult.statusCode === OSF.ProxyCallStatusCode.Succeeded && asyncResult.value) {
                    var appState = asyncResult.value;
                    var context = asyncResult.context;
                    if (!asyncResult.cached) {
                        _omexDataProvider.SetAppStateCache(asyncResult.context, appState);
                    }
                    var createAnonymousOmexProxyCompleted = function (asyncResult) {
                        context.clientEndPoint = me._omexAnonymousWSProxy.clientEndPoint;
                        OSF.OsfManifestManager.getOmexManifestAndETokenAsync(context, onGetOmexManifestAndETokenCompleted);
                    };
                    if (context.osfControl._omexEntitlement && !context.osfControl._omexEntitlement.hasEntitlement) {
                        context.anonymous = true;
                        me._createOmexProxy(me._omexAnonymousWSProxy, createAnonymousOmexProxyCompleted);
                    }
                    else {
                        OSF.OsfManifestManager.getOmexManifestAndETokenAsync(context, onGetOmexManifestAndETokenCompleted);
                    }
                }
            };
            var checkMyAppsCompleted = function (asyncResult) {
                try {
                    context.anonymous = !asyncResult.isReady;
                    context.clientEndPoint = context.anonymous ? me._omexAnonymousWSProxy.clientEndPoint : me._omexGatedWSProxy.clientEndPoint;
                    if (!context.anonymous) {
                        OSF.OsfManifestManager.getOmexEntitlementsAsync(context, onGetOmexEntitlementsCompleted);
                    }
                    OSF.OsfManifestManager.getOmexKilledAppsAsync(context, onGetOmexKilledAppsCompleted);
                    if (context.anonymous) {
                        OSF.OsfManifestManager.getOmexAppStateAsync(context, onGetOmexAppStateCompleted);
                    }
                }
                catch (e) {
                }
            };
            me.isMyAppsReady(checkMyAppsCompleted);
        }
    };
    var removeAppForInsertDialog = function OSF_ContextActivationManager$removeAppForInsertDialog(params, onRemoveComplete) {
        var response = { "errorCode": OSF.InvokeResultCode.E_OEM_REMOVED_FAILED };
        var context = {
            "assetId": params.id,
            "clientVersion": me._getClientVersionForOmex(),
            "clientName": me._getClientNameForOmex(),
            "clientEndPoint": me._omexGatedWSProxy.clientEndPoint
        };
        var currentInsertDialog = me._insertDialogDiv;
        var onRemoveAppCompleted = function OSF_ContextActivationManager_removeAppForInsertDialog$onRemoveAppCompleted(asyncResult) {
            var untrustControlCount = 0;
            if (asyncResult.statusCode === OSF.ProxyCallStatusCode.Succeeded && asyncResult.value && asyncResult.value.removedApps && asyncResult.value.removedApps.length > 0) {
                var removedApps = asyncResult.value.removedApps;
                var invokeResultCode = OSF.InvokeResultCode.E_OEM_REMOVED_FAILED;
                for (var i = 0; i < removedApps.length; i++) {
                    if (removedApps[i].assetId == context.assetId) {
                        if (removedApps[i].result == OSF.OmexRemoveAppStatus.Success) {
                            invokeResultCode = OSF.InvokeResultCode.S_OK;
                        }
                        break;
                    }
                }
                _omexDataProvider.RemoveManifestAndEToken(context.assetId);
                if (invokeResultCode == OSF.InvokeResultCode.S_OK) {
                    untrustControlCount = me.untrustOsfControls(params);
                }
                response.errorCode = invokeResultCode;
            }
            var isDialogClosed = !document.documentElement.contains(currentInsertDialog);
            Telemetry.AppManagementMenuHelper.LogAppManagementMenuAction(context.assetId, OSF.AppManagementAction.Remove, untrustControlCount, isDialogClosed, false, response.errorCode);
            onRemoveComplete(response);
        };
        OSF.OsfManifestManager.removeOmexAppAsync(context, Function.createDelegate(me, onRemoveAppCompleted));
    };
    var logTelemetryDataForInsertDialog = function OSF_ContextActivationManager$logTelemetryDataForInsertDialog(params, onComplete) {
        switch (params.datapointName) {
            case OSF.DataPointNames.AppManagementMenu:
                OSF.OUtil.validateParamObject(params, {
                    "assetId": { type: String, mayBeNull: false },
                    "operationMetadata": { type: Number, mayBeNull: false },
                    "hrStatus": { type: Number, mayBeNull: false }
                }, null);
                Telemetry.AppManagementMenuHelper.LogAppManagementMenuAction(params.assetId, params.operationMetadata, 0, false, false, params.hrStatus);
                break;
            case OSF.DataPointNames.InsertionDialogSession:
                OSF.OUtil.validateParamObject(params, {
                    "assetId": { type: String, mayBeNull: false },
                    "totalSessionTime": { type: Number, mayBeNull: false },
                    "trustPageSessionTime": { type: Number, mayBeNull: false },
                    "appInserted": { type: Boolean, mayBeNull: false },
                    "lastActiveTab": { type: Number, mayBeNull: false },
                    "lastActiveTabCount": { type: Number, mayBeNull: false }
                }, null);
                Telemetry.InsertionDialogSessionHelper.LogInsertionDialogSession(params.assetId, params.totalSessionTime, params.trustPageSessionTime, params.appInserted, params.lastActiveTab, params.lastActiveTabCount);
                break;
            case OSF.DataPointNames.UploadFileDevCatelog:
                OSF.OUtil.validateParamObject(params, {
                    "operationMetadata": { type: Number, mayBeNull: false },
                    "hrStatus": { type: Number, mayBeNull: false }
                }, null);
                Telemetry.UploadFileDevCatelogHelper.LogUploadFileDevCatelogAction(this._appCorrelationId, params.operationMetadata, 0, false, false, params.hrStatus);
                break;
            case OSF.DataPointNames.UploadFileDevCatalogUsage:
                OSF.OUtil.validateParamObject(params, {
                    "currentStoreType": { type: String, mayBeNull: false },
                    "id": { type: String, mayBeNull: true },
                    "appVersion": { type: String, mayBeNull: true },
                    "appTargetType": { type: Number, mayBeNull: false },
                    "isAppCommand": { type: String, mayBeNull: false },
                    "appSizeWidth": { type: Number, mayBeNull: true },
                    "appSizeHeight": { type: Number, mayBeNull: true }
                }, null);
                Telemetry.UploadFileDevCatalogUsageHelper.LogUploadFileDevCatalogUsageAction(this._appCorrelationId, params.currentStoreType, params.Id, params.appVersion, params.appTargetType, params.isAppCommand, params.appSizeWidth, params.appSizeHeight);
                break;
        }
    };
    this._serviceEndPoint = Microsoft.Office.Common.XdmCommunicationManager.createServiceEndPoint(this._clientId);
    this._serviceEndPoint._onHandleRequestError = function (requestMessage, exception) {
        var param;
        var osfControlInstanceId;
        if (requestMessage && (param = requestMessage._data)) {
            osfControlInstanceId = (typeof (param) === "object" ? param[0] : param);
            osfControlInstanceId = (typeof (osfControlInstanceId) === "string" ? osfControlInstanceId : null);
            var osfControl;
            if (osfControlInstanceId && (osfControl = me.getOsfControl(osfControlInstanceId))) {
                if (exception === "Failed origin check") {
                    osfControl._pageStatus = OSF.OsfControlPageStatus.FailedOriginCheck;
                }
                else if (exception === "Access Denied") {
                    osfControl._pageStatus = OSF.OsfControlPageStatus.FailedPermissionCheck;
                }
                else {
                    osfControl._pageStatus = OSF.OsfControlPageStatus.Failed;
                }
                Telemetry.RuntimeTelemetryHelper.LogExceptionTag(exception, null, osfControl.getCorrelationId(), 0x010c53a0);
            }
        }
    };
    this._serviceEndPoint.registerMethod("ContextActivationManager_getAppContextAsync", getAppContextAsync, Microsoft.Office.Common.InvokeType.async, false);
    this._serviceEndPoint.registerMethod("ContextActivationManager_notifyHost", notifyHost, Microsoft.Office.Common.InvokeType.async, false);
    this._serviceEndPoint.registerMethod("ContextActivationManager_openWindowInHost", openWindowInHost, Microsoft.Office.Common.InvokeType.async, false);
    this._serviceEndPointInternal = Microsoft.Office.Common.XdmCommunicationManager.createServiceEndPoint(this._clientId + OSF.Constants.EndPointInternalSuffix);
    this._serviceEndPointInternal.registerMethod("ContextActivationManager_getEntitlementsForInsertDialog", getEntitlementsForInsertDialog, Microsoft.Office.Common.InvokeType.async, false);
    this._serviceEndPointInternal.registerMethod("ContextActivationManager_invokeSignIn", invokeSignIn, Microsoft.Office.Common.InvokeType.async, false);
    this._serviceEndPointInternal.registerMethod("ContextActivationManager_invokeWindowOpen", invokeWindowOpen, Microsoft.Office.Common.InvokeType.async, false);
    this._serviceEndPointInternal.registerMethod("ContextActivationManager_onClickInsertOsfControl", onClickInsertOsfControl, Microsoft.Office.Common.InvokeType.async, false);
    this._serviceEndPointInternal.registerMethod("ContextActivationManager_onClickInstallOsfControl", onClickInstallOsfControl, Microsoft.Office.Common.InvokeType.async, false);
    this._serviceEndPointInternal.registerMethod("ContextActivationManager_onClickRefreshAddinCommands", onClickRefreshAddinCommands, Microsoft.Office.Common.InvokeType.async, false);
    this._serviceEndPointInternal.registerMethod("ContextActivationManager_onClickCancelDialog", onClickCancelDialog, Microsoft.Office.Common.InvokeType.async, false);
    this._serviceEndPointInternal.registerMethod("ContextActivationManager_removeAppForInsertDialog", removeAppForInsertDialog, Microsoft.Office.Common.InvokeType.async, false);
    this._serviceEndPointInternal.registerMethod("ContextActivationManager_logTelemetryDataForInsertDialog", logTelemetryDataForInsertDialog, Microsoft.Office.Common.InvokeType.async, false);
    this._serviceEndPointInternal.registerMethod("ContextActivationManager_getOmexData", getOmexData, Microsoft.Office.Common.InvokeType.async, false);
    OSF.AppSpecificSetup._setupFacade(this._hostControl, this, this._serviceEndPoint, this._serviceEndPointInternal);
    this._localeStringLoadingPendingCallbacks = [];
    if (this._localizedScriptsUrl && this._localizedScriptsUrl != "null/") {
        var localeStringLoaded = function () {
            this._ErrorUXHelper = new OSF._ErrorUXHelper(this);
        };
        this._loadLocaleString(Function.createDelegate(this, localeStringLoaded));
    }
    this.queryEntitlementDetails = function OSF_ContextActivationManager$queryEntitlementDetailsInternal(params) {
        getEntitlementsForInsertDialog(params, params.callbackFunc);
    };
    ;
};
OSF.ContextActivationManager.prototype = {
    insertOsfControl: function OSF_ContextActivationManager$insertOsfControl(params) {
        OSF.OUtil.validateParamObject(params, {
            "div": { type: Object, mayBeNull: false },
            "id": { type: String, mayBeNull: false },
            "marketplaceID": { type: String, mayBeNull: false },
            "marketplaceVersion": { type: String, mayBeNull: false },
            "store": { type: String, mayBeNull: false },
            "storeType": { type: String, mayBeNull: false },
            "alternateReference": { type: Object, mayBeNull: true },
            "settings": { type: Object, mayBeNull: true },
            "reason": { type: String, mayBeNull: true },
            "osfControlType": { type: Number, mayBeNull: true },
            "snapshotUrl": { type: String, mayBeNull: true },
            "preactivationCallback": { type: Object, mayBeNull: true },
            "virtualOsfControlActivationCallback": { type: Object, mayBeNull: true },
            "isvirtualOsfControl": { type: Boolean, mayBeNull: true }
        }, null);
        var osfControlParams = {
            "div": params.div,
            "id": params.id,
            "marketplaceID": params.marketplaceID,
            "marketplaceVersion": params.marketplaceVersion,
            "store": params.store,
            "storeType": params.storeType,
            "alternateReference": params.alternateReference,
            "settings": params.settings,
            "reason": params.reason,
            "osfControlType": params.osfControlType,
            "snapshotUrl": params.snapshotUrl,
            "contextActivationMgr": this,
            "preactivationCallback": params.preactivationCallback,
            "virtualOsfControlActivationCallback": params.virtualOsfControlActivationCallback,
            "isvirtualOsfControl": params.isvirtualOsfControl
        };
        var sqmDWords = this.getSQMAgaveUsage(osfControlParams.storeType, osfControlParams.osfControlType, osfControlParams.reason, params.marketplaceID);
        var osfControl = new OSF.OsfControl(osfControlParams);
        osfControl._sqmDWords[0] = sqmDWords[0];
        osfControl._sqmDWords[1] = sqmDWords[1];
        if (osfControl._contextActivationMgr._ErrorUXHelper) {
            osfControl._contextActivationMgr._ErrorUXHelper.showProgress(osfControl._div, osfControl._id);
        }
        var localeStringLoaded = function () {
            osfControl.activate();
        };
        this._loadLocaleString(Function.createDelegate(this, localeStringLoaded));
        return osfControl;
    },
    setLocalizedUrl: function OSF_ContextActivationManager$setLocalizedUrl(scriptUrl, imageUrl, styleUrl) {
        var e = Function._validateParams(arguments, [
            { name: "scriptUrl", type: String, mayBeNull: false },
            { name: "imageUrl", type: String, mayBeNull: false },
            { name: "styleUrl", type: String, mayBeNull: false }
        ]);
        if (e)
            throw e;
        this._localizedScriptsUrl = scriptUrl;
        this._localizedImagesUrl = imageUrl;
        this._localizedStylesUrl = styleUrl;
        if (this._localizedScriptsUrl && this._localizedScriptsUrl != "null/") {
            var localeStringLoaded = function () {
                if (!this._ErrorUXHelper) {
                    this._ErrorUXHelper = new OSF._ErrorUXHelper(this);
                }
            };
            this._loadLocaleString(Function.createDelegate(this, localeStringLoaded));
        }
    },
    activateOsfControl: function OSF_ContextActivationManager$activateOsfControl(id) {
        var e = Function._validateParams(arguments, [
            { name: "id", type: String, mayBeNull: false }
        ]);
        if (e)
            throw e;
        var osfControl = this._cachedOsfControls[id];
        if (typeof osfControl != "undefined") {
            osfControl.activate();
        }
    },
    deActivateOsfControl: function OSF_ContextActivationManager$deActivateOsfControl(id) {
        var e = Function._validateParams(arguments, [
            { name: "id", type: String, mayBeNull: false }
        ]);
        if (e)
            throw e;
        var osfControl = this._cachedOsfControls[id];
        if (typeof osfControl != "undefined") {
            osfControl.deActivate();
        }
    },
    purgeOsfControl: function OSF_ContextActivationManager$purgeOsfControl(id, purgeManifest) {
        var e = Function._validateParams(arguments, [{ name: "id", type: String, mayBeNull: false },
            { name: "purgeManifest", type: Boolean, mayBeNull: false }
        ]);
        if (e)
            throw e;
        var osfControl = this._cachedOsfControls[id];
        if (typeof osfControl != "undefined") {
            osfControl.purge(purgeManifest);
        }
    },
    purgeOsfControlNotifications: function OSF_ContextActivationManager$purgeOsfControlNotifications() {
        for (var id in this._cachedOsfControls) {
            this._ErrorUXHelper.purgeOsfControlNotification(id);
        }
    },
    untrustOsfControls: function OSF_ContextActivationManager$untrustOsfControl(params) {
        var untrustControlCount = 0;
        var cacheKey = OSF.OUtil.formatString(OSF.Constants.ActivatedCacheKey, params.id.toLowerCase(), params.currentStoreType, params.storeId);
        this._deleteCachedFlag(cacheKey);
        for (var id in this._cachedOsfControls) {
            var osfControl = this._cachedOsfControls[id];
            if (osfControl._marketplaceID.toLowerCase() == params.id.toLowerCase()) {
                this._ErrorUXHelper.purgeOsfControlNotification(osfControl._id);
                osfControl.deActivate();
                osfControl._showTrustError(params.displayName, params.providerName, params.currentStoreType, Function.createDelegate(osfControl, osfControl.activate));
                untrustControlCount++;
            }
        }
        return untrustControlCount;
    },
    retryAll: function OSF_ContextActivationManager$retryAll(solutionId) {
        for (var id in this._cachedOsfControls) {
            var osfControl = this._cachedOsfControls[id];
            if (osfControl._marketplaceID.toLowerCase() == solutionId.toLowerCase()) {
                if (osfControl._retryActivate) {
                    this._ErrorUXHelper.purgeOsfControlNotification(osfControl._id);
                    this._ErrorUXHelper.removeInfoBarDiv(id, false);
                    osfControl._retryActivate();
                    osfControl._retryActivate = null;
                }
            }
        }
    },
    getOsfControl: function OSF_ContextActivationManager$getOsfControl(id) {
        var e = Function._validateParams(arguments, [
            { name: "id", type: String, mayBeNull: false }
        ]);
        if (e)
            throw e;
        return this._cachedOsfControls[id];
    },
    getOsfControls: function OSF_ContextActivationManager$getOsfControls() {
        var osfControls = [];
        for (var id in this._cachedOsfControls) {
            osfControls.push(this._cachedOsfControls[id]);
        }
        return osfControls;
    },
    getOsfOmexBaseUrl: function OSF_ContextActivationManager$getOsfOmexBaseUrl() {
        return this._osfOmexBaseUrl;
    },
    getAppName: function OSF_ContextActivationManager$getAppName() {
        return this._appName;
    },
    getClientMode: function OSF_ContextActivationManager$getClientMode() {
        return this._clientMode;
    },
    getClientId: function OSF_ContextActivationManager$getClientId() {
        return this._clientId;
    },
    getFormFactor: function OSF_ContextActivationManager$getFormFactor() {
        return this._formFactor;
    },
    getDocUrl: function OSF_ContextActivationManager$getDocUrl() {
        return this._docUrl;
    },
    getAppUILocale: function OSF_ContextActivationManager$getAppUILocale() {
        return this._appUILocale;
    },
    getDataLocale: function OSF_ContextActivationManager$getDataLocale() {
        return this._dataLocale;
    },
    getPageBaseUrl: function OSF_ContextActivationManager$getPageBaseUrl() {
        return this._pageBaseUrl;
    },
    getLcid: function OSF_ContextActivationManager$getLcid() {
        return this._lcid;
    },
    isExternalMarketplaceAllowed: function OSF_ContextActivationManager$isExternalMarketplaceAllowed() {
        return this._allowExternalMarketplace;
    },
    getLocalizedScriptsUrl: function OSF_ContextActivationManager$getLocalizedScriptsUrl() {
        return (this._localizedScriptsUrl ? this._localizedScriptsUrl : "");
    },
    getLocalizedImagesUrl: function OSF_ContextActivationManager$getLocalizedImagesUrl() {
        return this._localizedImagesUrl ? this._localizedImagesUrl : "";
    },
    getLocalizedStylesUrl: function OSF_ContextActivationManager$getLocalizedStylesUrl() {
        return this._localizedStylesUrl ? this._localizedStylesUrl : "";
    },
    raiseOsfControlStatusChange: function OSF_ContextActivationManager$raiseOsfControlStatusChange(osfControl) {
        if (this._controlStatusChanged) {
            this._controlStatusChanged(osfControl);
        }
    },
    registerOsfControl: function OSF_ContextActivationManager$registerOsfControl(osfControl) {
        var e = Function._validateParams(arguments, [
            { name: "osfControl", type: Object, mayBeNull: false }
        ]);
        if (e)
            throw e;
        this._cachedOsfControls[osfControl.getID()] = osfControl;
    },
    unregisterOsfControl: function OSF_ContextActivationManager$unregisterOsfControl(osfControl) {
        var e = Function._validateParams(arguments, [
            { name: "osfControl", type: Object, mayBeNull: false }
        ]);
        if (e)
            throw e;
        delete this._cachedOsfControls[osfControl.getID()];
    },
    setIframeAttributeBag: function OSF_ContextActivationManager$setIframeAttributeBag(iframeAttributeBag) { this._iframeAttributeBag = iframeAttributeBag; },
    displayNotification: function OSF_ContextActivationManager$displayNotification(params) {
        OSF.OUtil.validateParamObject(params, {
            "infoType": { type: Number, mayBeNull: false },
            "id": { type: String, mayBeNull: false },
            "title": { type: String, mayBeNull: false },
            "description": { type: String, mayBeNull: false },
            "url": { type: String, mayBeNull: true },
            "buttonTxt": { type: String, mayBeNull: true },
            "buttonCallback": { type: Function, mayBeNull: true },
            "dismissCallback": { type: Function, mayBeNull: true },
            "displayDeactive ": { type: Boolean, mayBeNull: true },
            "highPriority": { type: Boolean, mayBeNull: true },
            "detailView": { type: Boolean, mayBeNull: true },
            "reDisplay": { type: Boolean, mayBeNull: true },
            "logAsError": { type: Boolean, mayBeNull: true },
            "errorCode": { type: Number, mayBeNull: true },
            "retryAll": { type: Boolean, mayBeNull: true }
        }, null);
        if (!params.errorCode) {
            params.errorCode = 0;
        }
        var osfControl = this._cachedOsfControls[params.id];
        if (osfControl) {
            if (osfControl._isvirtualOsfControl) {
                osfControl._notificationParams = params;
                if (!osfControl._isvirtualOsfControlCallbackInvoked) {
                    if (osfControl._manifest != null) {
                        osfControl.invokeVirtualOsfControlActivationCallback(osfControl._manifest);
                    }
                    else {
                        osfControl.forceGetVirtualOsfControlManifest();
                    }
                }
            }
            else {
                if (params.logAsError) {
                    Telemetry.AppLoadTimeHelper.SetErrorResult(osfControl._telemetryContext, params.errorCode);
                    if (this._notifyHostError) {
                        var notifyHostParams = {
                            "errorCode": params.errorCode
                        };
                        this._notifyHostError(params.id, OSF.AgaveHostAction.NotifyHostError, notifyHostParams);
                    }
                }
                if (params.displayDeactive) {
                    params.detailView = true;
                }
                params["div"] = osfControl._div;
                this._ErrorUXHelper.showNotification(params);
            }
        }
    },
    dismissMessages: function OSF_ContextActivationManager$dismissMessages(id) {
        if (this._ErrorUXHelper) {
            this._ErrorUXHelper.dismissMessages(id);
        }
    },
    notifyAgave: function OSF_ContextActivationManager$notifyAgave(id, actionId) {
        var e = Function._validateParams(arguments, [{ name: "id", type: String, mayBeNull: false },
            { name: "actionId", type: Number, mayBeNull: false }
        ]);
        if (e)
            throw e;
        var osfControl = this._cachedOsfControls[id];
        if (typeof osfControl != "undefined") {
            if (osfControl._pageStatus !== OSF.OsfControlPageStatus.Ready && (actionId === OSF.AgaveHostAction.CtrlF6In || actionId === OSF.AgaveHostAction.Select)) {
                this._ErrorUXHelper.focusOnNotificationUx(id);
            }
            else {
                osfControl.notifyAgave(actionId);
            }
        }
    },
    setWSA: function OSF_ContextActivationManager$setWSA(wsa) {
        this._wsa = wsa;
        if (this._wsa) {
            this._wsa.createStreamUnobfuscated(OSF.SQMDataPoints.DATAID_APPSFOROFFICEUSAGE, OSF.BWsaStreamTypes.Static, 2, OSF.BWsaConfig.defaultMaxStreamRows);
            this._wsa.createStreamUnobfuscated(OSF.SQMDataPoints.DATAID_APPSFOROFFICENOTIFICATIONS, OSF.BWsaStreamTypes.Static, 4, OSF.BWsaConfig.defaultMaxStreamRows);
        }
    },
    getWSA: function OSF_ContextActivationManager$getWSA() {
        return this._wsa;
    },
    getSQMAgaveUsage: function OSF_ContextActivationManager$getSQMAgaveUsage(provider, shape, context, assetId) {
        var sqmProvider;
        switch (provider.toLowerCase()) {
            case OSF.StoreType.OMEX:
                sqmProvider = 0;
                break;
            case OSF.StoreType.SPCatalog:
                sqmProvider = 1;
                break;
            case OSF.StoreType.FileSystem:
                sqmProvider = 2;
                break;
            case OSF.StoreType.Registry:
                sqmProvider = 3;
                break;
            case OSF.StoreType.Exchange:
                sqmProvider = 4;
                break;
            case OSF.StoreType.SPApp:
                sqmProvider = 5;
                break;
            default:
                sqmProvider = 15;
                break;
        }
        var sqmShape;
        switch (shape) {
            case OSF.OsfControlType.DocumentLevel:
                sqmShape = 1;
                break;
            case OSF.OsfControlType.ContainerLevel:
                sqmShape = 0;
                break;
            default:
                sqmShape = 2;
                break;
        }
        var sqmContext = 7;
        if (context && context.toLowerCase) {
            switch (context.toLowerCase()) {
                case Microsoft.Office.WebExtension.InitializationReason.Inserted.toLowerCase():
                    sqmContext = 0;
                    break;
                case Microsoft.Office.WebExtension.InitializationReason.DocumentOpened.toLowerCase():
                    sqmContext = 1;
                    break;
                default:
                    break;
            }
        }
        var sqmAssetId = assetId.toLowerCase().indexOf("wa") === 0 ? parseInt(assetId.substring(2), 10) : 0;
        var dWord1 = 0;
        dWord1 = sqmContext << 8 | sqmShape << 4 | sqmProvider;
        return [dWord1, sqmAssetId];
    },
    isMyAppsReady: function OSF_ContextActivationManager$isMyAppsReady(onCompleted) {
        var me = this;
        if (!me._enableMyApps) {
            onCompleted({
                "isReady": false
            });
            return;
        }
        if (me._omexGatedWSProxy && me._omexGatedWSProxy.clientEndPoint) {
            onCompleted({
                "isReady": true,
                "clientEndPoint": me._omexGatedWSProxy.clientEndPoint
            });
            return;
        }
        var omexAuthenticatedConnectTries = 1;
        var onGetAuthNStatusCompleted = function OSF_ContextActivationManager_isMyAppsReady$onGetAuthNStatusCompleted(asyncResult) {
            if (asyncResult.statusCode === OSF.ProxyCallStatusCode.Succeeded && asyncResult.value) {
                var authNStatus = parseInt(asyncResult.value);
                if (authNStatus == OSF.OmexAuthNStatus.Authenticated) {
                    me._omexAuthNStatus = OSF.OmexAuthNStatus.Authenticated;
                }
                else if (authNStatus == OSF.OmexAuthNStatus.Anonymous || authNStatus == OSF.OmexAuthNStatus.Unknown) {
                    onCompleted({
                        "isReady": false
                    });
                    return;
                }
            }
            else {
                me._omexAuthNStatus = OSF.OmexAuthNStatus.CheckFailed;
            }
            me._createOmexProxy(me._omexGatedWSProxy, createAuthenticatedOmexProxyCompleted);
        };
        var createAnonymousOmexProxyCompleted = function OSF_ContextActivationManager_isMyAppsReady$createAnonymousOmexProxyCompleted(clientEndPoint) {
            if (clientEndPoint && (me._omexAuthNStatus == OSF.OmexAuthNStatus.NotAttempted || me._omexAuthNStatus == OSF.OmexAuthNStatus.Anonymous)) {
                if (me._omexAuthNStatus == OSF.OmexAuthNStatus.NotAttempted) {
                    var params = { "clientEndPoint": null };
                    params.clientEndPoint = clientEndPoint;
                    OSF.OsfManifestManager._invokeProxyMethodAsync(params, "OMEX_getAuthNStatus", onGetAuthNStatusCompleted, params);
                }
                else if (me._omexAuthNStatus == OSF.OmexAuthNStatus.Anonymous) {
                    onCompleted({
                        "isReady": false,
                        "clientEndPoint": clientEndPoint
                    });
                }
            }
            else {
                me._createOmexProxy(me._omexGatedWSProxy, createAuthenticatedOmexProxyCompleted);
            }
        };
        var createAuthenticatedOmexProxyCompleted = function OSF_ContextActivationManager_isMyAppsReady$createAuthenticatedOmexProxyCompleted(clientEndPoint) {
            if (clientEndPoint) {
                onCompleted({
                    "isReady": true,
                    "clientEndPoint": clientEndPoint
                });
            }
            else {
                if (omexAuthenticatedConnectTries < OSF.Constants.AuthenticatedConnectMaxTries) {
                    omexAuthenticatedConnectTries++;
                    if (me._omexAuthNStatus == OSF.OmexAuthNStatus.CheckFailed) {
                        setTimeout(function () { me._createOmexProxy(me._omexGatedWSProxy, createAuthenticatedOmexProxyCompleted); }, 500);
                    }
                    else {
                        me._createOmexProxy(me._omexAnonymousWSProxy, createAnonymousOmexProxyCompleted);
                    }
                }
                else {
                    onCompleted({
                        "isReady": false
                    });
                }
            }
        };
        me._createOmexProxy(me._omexGatedWSProxy, createAuthenticatedOmexProxyCompleted);
    },
    isMyOrgReady: function OSF_ContextActivationManager$isMyOrgReady(onCompleted) {
        var me = this;
        if (!me._enableMyOrg) {
            onCompleted({
                "isReady": false
            });
            return;
        }
        if (me._spBaseUrl && me._iframeProxies && me._iframeProxies[me._spBaseUrl] && me._iframeProxies[me._spBaseUrl].clientEndPoint) {
            onCompleted({
                "isReady": true,
                "clientEndPoint": me._iframeProxies[me._spBaseUrl].clientEndPoint
            });
            return;
        }
        var createSharePointProxyCompleted = function OSF_ContextActivationManager_isMyOrgReady$createSharePointProxyCompleted(clientEndPoint) {
            if (clientEndPoint) {
                onCompleted({
                    "isReady": true,
                    "clientEndPoint": clientEndPoint
                });
            }
            else {
                onCompleted({
                    "isReady": false
                });
            }
        };
        me._createSharePointIFrameProxy(me._spBaseUrl, createSharePointProxyCompleted);
    },
    openInputUrlDialog: function OSF_ContextActivationManager$openInputUrlDialog(divContainer) {
        var titleText = document.createElement("p");
        titleText.setAttribute("id", "title-p");
        titleText.textContent = "DevCatalog server Url is configured in settings.xml - Enter manifest file name only (AppId.xml):";
        divContainer.appendChild(titleText);
        var urlInput = document.createElement("input");
        urlInput.setAttribute("type", "url");
        urlInput.setAttribute("id", "url-input");
        urlInput.setAttribute("size", "80");
        divContainer.appendChild(urlInput);
        var me = this;
        var processManifestFile = function OSF_ContextActivationManager_openInputUrlDialog$processManifestFile(manifestString, urlInputElement) {
            var parsedManifest = new OSF.Manifest.Manifest(manifestString, me.getAppUILocale());
            if (!OSF.OsfManifestManager.hasManifest(parsedManifest.getMarketplaceID(), parsedManifest.getMarketplaceVersion())) {
                OSF.OsfManifestManager.cacheManifest(parsedManifest.getMarketplaceID(), parsedManifest.getMarketplaceVersion(), parsedManifest);
            }
            var params = {
                "id": parsedManifest.getMarketplaceID(),
                "targetType": parsedManifest.getTarget(),
                "appVersion": parsedManifest.getMarketplaceVersion(),
                "currentStoreType": OSF.StoreType.Registry,
                "storeId": "developer",
                "assetId": parsedManifest.getMarketplaceID(),
                "assetStoreId": OSF.StoreType.Registry,
                "width": parsedManifest.getDefaultWidth() || 0,
                "height": parsedManifest.getDefaultHeight() || 0,
                "isAddinCommands": parsedManifest._isAddinCommandsManifest(OSF.getManifestHostType(me._hostType))
            };
            Telemetry.UploadFileDevCatalogUsageHelper.LogUploadFileDevCatalogUsageAction(this._appCorrelationId, params.currentStoreType, params.id, params.appVersion, params.targetType, params.isAddinCommands, params.width, params.height);
            me._notifyHost("0", OSF.AgaveHostAction.InsertAgave, params);
        };
        var onGetManifestError = function OSF_ContextActivationManager_openInputUrlDialog$onGetManifestError(errorString) {
            alert("Error when requsting manifest file: " + errorString);
        };
        var onInsertButton = function OSF_ContextActivationManager_openInputUrlDialog$onInsertButton() {
            OSF.OUtil.xhrGet(me._devCatalogUrl + "/" + urlInput.value, processManifestFile, onGetManifestError);
        };
        var insertButton = document.createElement("input");
        insertButton.setAttribute("type", "button");
        insertButton.setAttribute("value", "Insert");
        OSF.OUtil.addEventListener(insertButton, "click", onInsertButton);
        divContainer.appendChild(insertButton);
    },
    launchInsertDialog: function OSF_ContextActivationManager$launchInsertDialog(containerDiv, storeId, navigationParams) {
        if (containerDiv.childNodes.length != 0) {
            containerDiv.removeChild(containerDiv.childNodes.item(0));
        }
        var div;
        if (this._enableDevCatalog) {
            div = document.createElement("div");
            div.style.width = "100%";
            div.style.height = "80%";
            containerDiv.appendChild(div);
            var div1 = document.createElement("div");
            div1.style.width = "100%";
            div1.style.height = "20%";
            containerDiv.appendChild(div1);
            this.openInputUrlDialog(div1);
        }
        else {
            div = containerDiv;
        }
        var me = this;
        me._insertDialogDiv = div;
        var loadingDiv = document.createElement('div');
        loadingDiv.style.width = "100%";
        loadingDiv.style.height = "100%";
        loadingDiv.style.backgroundImage = "url(" + me.getLocalizedImageFilePath("progress.gif") + ")";
        loadingDiv.style.backgroundRepeat = "no-repeat";
        loadingDiv.style.backgroundPosition = "center";
        div.appendChild(loadingDiv);
        var storeIds = {
            MyApp: "0",
            MyOrg: "1",
            Store: "{98143890-AC66-440E-A448-ED8771A02D52}",
            OneDrive: "9"
        };
        var getCorporateCatalogUrlAsync = function OSF_ContextActivationManager_launchInsertDialog$getCorporateCatalogUrlAsync(context, onCompleted) {
            if (!me._enableMyOrg) {
                onCompleted({
                    "statusCode": OSF.InvokeResultCode.E_CATALOG_NO_APPS,
                    "value": null,
                    "context": context
                });
                return;
            }
            OSF.OUtil.validateParamObject(context, {
                "webUrl": {
                    type: String,
                    mayBeNull: false
                }
            }, onCompleted);
            var checkMyOrgCompleted = function OSF_ContextActivationManager_launchInsertDialog$checkMyOrgCompleted(asyncResult) {
                if (asyncResult.isReady) {
                    context.clientEndPoint = asyncResult.clientEndPoint;
                    var params = {
                        "webUrl": context.webUrl
                    };
                    OSF.OsfManifestManager._invokeProxyMethodAsync(context, "OEM_getSPCatalogUrlAsync", onCompleted, params);
                }
                else {
                    onCompleted({
                        "statusCode": OSF.InvokeResultCode.E_GENERIC_ERROR,
                        "value": null,
                        "context": context
                    });
                }
            };
            me.isMyOrgReady(checkMyOrgCompleted);
        };
        var constructInsertDialog = function OSF_ContextActivationManager_launchInsertDialog$constructInsertDialog(asyncResult) {
            var frame = document.createElement("iframe");
            frame.setAttribute("id", "InsertDialog");
            frame.setAttribute("src", "about:blank");
            frame.setAttribute("width", "100%");
            frame.setAttribute("height", "100%");
            frame.setAttribute("marginHeight", "0");
            frame.setAttribute("marginWidth", "0");
            frame.setAttribute("frameBorder", "0");
            frame.setAttribute("sandbox", "allow-scripts allow-forms allow-same-origin ms-allow-popups allow-popups");
            if (typeof Strings != 'undefined' && Strings && Strings.OsfRuntime) {
                frame.setAttribute("title", Strings.OsfRuntime.L_InsertionDialogTile_TXT);
            }
            var myAppProvider = null;
            var providers = null;
            var envSetting = '{ "IsUploadFileDevCatalogEnabled":' + me._enableUploadFileDevCatalog + ' }';
            if (me.isExternalMarketplaceAllowed()) {
                var pHres = me._enableMyApps ? OSF.InvokeResultCode.S_OK.toString() : OSF.InvokeResultCode.S_HIDE_PROVIDER.toString();
                myAppProvider = '{ "provValues":[0,0,0,' + pHres + '], "url":"' + me._osfOmexBaseUrl + '", "client":"' + me._getClientNameForOmex() + '"}';
                providers = '{ "myApp":' + myAppProvider + ' }';
            }
            var myOrgProvider = null;
            if (me._enableMyOrg) {
                me._myOrgCatalogUrl = asyncResult.value;
                var providerHResult = asyncResult.statusCode;
                myOrgProvider = '{ "provValues":[1,1,0,' + providerHResult + '], "url":"' + asyncResult.value + '"}';
                providers = me.isExternalMarketplaceAllowed() ? '{ "myApp":' + myAppProvider + ', "myOrg":' + myOrgProvider + ' }' : '{ "myOrg":' + myOrgProvider + ' }';
            }
            if (me._enableOneDriveCatalog) {
                if (document.cookie.indexOf('OneDriveCatalog=true') != -1) {
                    var myOneDriveProvider = null;
                    me._myOneDriveCatalogUrl = asyncResult.value;
                    myOneDriveProvider = '{ "provValues":[9,9,0,0], "url":"' + asyncResult.value + '"}';
                    if (me._enableMyOrg) {
                        providers = me.isExternalMarketplaceAllowed() ? '{ "myApp":' + myAppProvider + ', "myOrg":' + myOrgProvider + ', "myOneDrive":' + myOneDriveProvider + ' }' : '{ "myOrg":' + myOrgProvider + ',"myOneDrive":' + myOneDriveProvider + ' }';
                    }
                    else {
                        providers = me.isExternalMarketplaceAllowed() ? '{ "myApp":' + myAppProvider + ', "myOneDrive":' + myOneDriveProvider + ' }' : '{"myOneDrive":' + myOneDriveProvider + ' }';
                    }
                }
            }
            if (me._internalConversationId) {
                me._serviceEndPointInternal.unregisterConversation(me._internalConversationId);
            }
            var cacheKey = me.getClientId() + "_" + me.getDocUrl();
            var conversationId = OSF.OUtil.getFrameNameAndConversationId(cacheKey, frame);
            me._internalConversationId = OSF.OUtil.getFrameNameAndConversationId(cacheKey, frame);
            var navigation = null;
            if (navigationParams != null) {
                navigation = '{ "navigationMode":' + navigationParams.navigationMode + ', "navigationModeParameter": "' + navigationParams.assetId + '", "category":"' + navigationParams.category + '"}';
            }
            var urlWithoutHash = "";
            var urlParamArray = [
                conversationId,
                encodeURIComponent(window.location.href),
                me._appVersion,
                OSF.getAppVerCode(me._appName),
                me.getLcid(),
                OSF.Constants.FileVersion,
                encodeURIComponent(providers),
                storeId,
                encodeURIComponent(envSetting),
                encodeURIComponent(navigation),
                me._contextMgrCorrelationId
            ];
            if (me._localizedResourcesUrl) {
                urlWithoutHash = me._localizedResourcesUrl + "WefGallery.htm" + "?b=" + OSF.Constants.FileVersion;
            }
            else {
                urlWithoutHash = me._pageBaseUrl + me.getLcid() + "/WefGallery.htm" + "?b=" + OSF.Constants.FileVersion;
            }
            var newUrl = OSF.OUtil.addXdmInfoAsHash(urlWithoutHash, urlParamArray.join("|"));
            me._serviceEndPointInternal.registerConversation(conversationId, newUrl);
            frame.setAttribute("src", newUrl);
            if (me._insertDialogDiv) {
                if (div.childNodes.length != 0) {
                    div.removeChild(div.childNodes.item(0));
                }
                div.appendChild(frame);
            }
            else {
                div.insertBefore(frame, div.firstChild);
                me._insertDialogDiv = div;
            }
        };
        var context = { "webUrl": me._spBaseUrl };
        getCorporateCatalogUrlAsync(context, constructInsertDialog);
        me._preloadOfficeJs();
    },
    activateAgavesBlockedBySandboxNotSupport: function OSF_ContextActivationManager$activateAgavesBlockedBySandboxNotSupport() {
        for (var id in this._cachedOsfControls) {
            var osfcontrol = this._cachedOsfControls[id];
            if (osfcontrol._status === OSF.OsfControlStatus.NotSandBoxSupported) {
                osfcontrol.activate();
            }
        }
    },
    setRequirementsChecker: function OSF_ContextActivationManager$setRequirementsChecker(requirementsChecker) {
        this._requirementsChecker = requirementsChecker;
    },
    getRequirementsChecker: function OSF_ContextActivationManager$getRequirementsChecker() {
        return this._requirementsChecker;
    },
    appHasNotifications: function OSF_ContextActivationManager$appHasNotifications(id) {
        if (this._ErrorUXHelper) {
            return this._ErrorUXHelper.appHasNotifications(id);
        }
        return false;
    },
    _doesUrlHaveSupportedProtocol: function OSF_ContextActivationManager$_doesUrlHaveSupportedProtocol(url) {
        var isValid = false;
        if (url) {
            var decodedUrl = decodeURIComponent(url);
            var matches = decodedUrl.match(/^https?:\/\/.+$/ig);
            isValid = (matches != null);
        }
        return isValid;
    },
    _loadLocaleString: function OSF_ContextActivationManager$_loadLocaleString(callback) {
        if (typeof Strings == 'undefined' || !Strings || !Strings.OsfRuntime) {
            this._localeStringLoadingPendingCallbacks.push(callback);
            var loadStringPendingCallbacks = this._localeStringLoadingPendingCallbacks;
            if (loadStringPendingCallbacks.length === 1) {
                var loadLocaleStringBatchCallback = function () {
                    var pendingCallbackCount = loadStringPendingCallbacks.length;
                    for (var i = 0; i < pendingCallbackCount; i++) {
                        var currentCallback = loadStringPendingCallbacks.shift();
                        currentCallback();
                    }
                };
                this._loadStringScript(loadLocaleStringBatchCallback);
            }
        }
        else {
            var pendingCallbackCount = this._localeStringLoadingPendingCallbacks.length;
            for (var i = 0; i < pendingCallbackCount; i++) {
                var currentCallback = this._localeStringLoadingPendingCallbacks.shift();
                currentCallback();
            }
            callback();
        }
    },
    _loadStringScript: function OSF_ContextActivationManager$_loadStringScript(callback) {
        var path = this.getLocalizedScriptsUrl();
        path += "osfruntime_strings.debug.js";
        var localeStringFileLoaded = function () {
            if (typeof Strings == 'undefined' || !Strings || !Strings.OsfRuntime) {
                this._localeStringLoadingPendingCallbacks.length = 0;
                throw OSF.OUtil.formatString("The locale, {0}, provided by the host app is not supported.", this.getLcid());
            }
            else {
                callback();
            }
        };
        OSF.OUtil.loadScript(path, Function.createDelegate(this, localeStringFileLoaded));
    },
    _getServiceEndPoint: function OSF_ContextActivationManager$_getServiceEndPoint() {
        return this._serviceEndPoint;
    },
    _getOmexEndPointPageUrl: function OSF_ContextActivationManager$_getOmexEndPointPageUrl(assetId, contentMarketplace) {
        return OSF.OUtil.formatString("{0}/{1}/downloads/{2}.aspx", this._omexEndPointBaseUrl, contentMarketplace, assetId);
    },
    _getManifestAndTargetByConversationId: function OSF_ContextActivationManager$_getManifestAndTargetByConversationId(conversationId) {
        for (var id in this._cachedOsfControls) {
            var osfcontrol = this._cachedOsfControls[id];
            if (conversationId === osfcontrol._getConversationId()) {
                return { "manifest": OSF.OsfManifestManager.getCachedManifest(osfcontrol.getMarketplaceID(), osfcontrol.getMarketplaceVersion()), "target": osfcontrol.getOsfControlType() };
            }
        }
        return null;
    },
    _createOmexProxy: function OSF_ContextActivationManager$_createOmexProxy(omexProxy, callback) {
        var returnClientEndPoint = null;
        var onLoadCallback = function () {
            if (omexProxy.clientEndPoint) {
                omexProxy.clientEndPoint.invoke("OMEX_isProxyReady", onIsProxyReadyCallback, {
                    __timeout__: 500
                });
            }
            else {
                Telemetry.RuntimeTelemetryHelper.LogExceptionTag("Unexpected error, iframe loaded again after failing OMEX_isProxyReady.", null, null, 0x007d4284);
            }
            OSF.OUtil.set_entropy(new Date().getTime());
        };
        if (omexProxy.refresh) {
            omexProxy.isReady = false;
            omexProxy.refresh = false;
            if (omexProxy.clientEndPoint) {
                Microsoft.Office.Common.XdmCommunicationManager.deleteClientEndPoint(omexProxy.clientEndPoint._conversationId);
                omexProxy.clientEndPoint = null;
            }
            if (omexProxy.iframe) {
                OSF.OUtil.removeEventListener(omexProxy.iframe, "load", onLoadCallback);
                omexProxy.iframe.parentNode.removeChild(omexProxy.iframe);
            }
        }
        if (omexProxy.isReady) {
            if (!omexProxy.clientEndPoint) {
                Telemetry.RuntimeTelemetryHelper.LogExceptionTag("OmexProxy is ready while clientEndPoint is null.", null, null, 0x0108d741);
            }
            callback(omexProxy.clientEndPoint);
        }
        else if (!omexProxy.clientEndPoint) {
            if (!this._doesUrlHaveSupportedProtocol(omexProxy.proxyUrl)) {
                Telemetry.RuntimeTelemetryHelper.LogExceptionTag("Protocal of proxyUrl is not supported.", null, null, 0x0108d742);
                callback(null);
                return;
            }
            var conversationId = OSF.OUtil.generateConversationId();
            var iframe = document.createElement("iframe");
            iframe.setAttribute('id', omexProxy.proxyName);
            iframe.setAttribute('name', omexProxy.proxyName);
            var newUrl = omexProxy.proxyUrl + "?" + conversationId + "&ClientSessionId=" + this._contextMgrCorrelationId;
            newUrl = OSF.OUtil.addXdmInfoAsHash(newUrl, conversationId + "|" + omexProxy.proxyName + "|" + window.location.href);
            newUrl = OSF.OUtil.addSerializerVersionAsHash(newUrl, OSF.SerializerVersion.Browser);
            iframe.setAttribute('src', newUrl);
            iframe.setAttribute('scrolling', 'auto');
            iframe.setAttribute('border', '0');
            iframe.setAttribute('width', '0');
            iframe.setAttribute('height', '0');
            iframe.setAttribute('style', "position: absolute; left: -100px; top:0px;");
            var onIsProxyReadyCallback = function (errorCode, response) {
                var pendingCallbackCount = omexProxy.pendingCallbacks.length;
                if (pendingCallbackCount == 0) {
                    return;
                }
                if (errorCode === 0 && response.status) {
                    omexProxy.isReady = true;
                    omexProxy.iframe = iframe;
                    returnClientEndPoint = omexProxy.clientEndPoint;
                }
                else {
                    var errorInfo = OSF.OUtil.formatString("Proxy is not ready, errorCode:{0}, proxyUrl:{1}. {2}", errorCode, omexProxy.proxyUrl, OSF.OUtil.serializeOMEXResponseErrorMessage(response));
                    Telemetry.RuntimeTelemetryHelper.LogExceptionTag(errorInfo, null, null, 0x0109304c);
                    omexProxy.clientEndPoint = null;
                    if (Microsoft.Office.Common.XdmCommunicationManager.getClientEndPoint(conversationId)) {
                        Microsoft.Office.Common.XdmCommunicationManager.deleteClientEndPoint(conversationId);
                        OSF.OUtil.removeEventListener(iframe, "load", onLoadCallback);
                        iframe.parentNode.removeChild(iframe);
                    }
                    else {
                        Telemetry.RuntimeTelemetryHelper.LogExceptionTag("Unexpected error occured with marketplace proxy.", null, null, 0x007d4285);
                    }
                    returnClientEndPoint = null;
                }
                for (var i = 0; i < pendingCallbackCount; i++) {
                    var currentCallback = omexProxy.pendingCallbacks.shift();
                    currentCallback(returnClientEndPoint);
                }
            };
            try {
                document.body.appendChild(iframe);
                OSF.OUtil.addEventListener(iframe, "load", onLoadCallback);
                omexProxy.clientEndPoint = Microsoft.Office.Common.XdmCommunicationManager.connect(conversationId, iframe.contentWindow, omexProxy.proxyUrl);
                omexProxy.pendingCallbacks.push(callback);
            }
            catch (ex) {
                Telemetry.RuntimeTelemetryHelper.LogExceptionTag("Error creating client endpoint with proxyUrl=[" + omexProxy.proxyUrl + "].", ex, null, 0x0108d744);
                throw ex;
            }
        }
        else {
            omexProxy.pendingCallbacks.push(callback);
        }
    },
    _createSharePointIFrameProxy: function OSF_ContextActivationManager$_createSharePointIFrameProxy(url, callback) {
        if (!this._doesUrlHaveSupportedProtocol(url)) {
            callback(null);
            return;
        }
        var urlLength = url.length;
        if (url.charAt(urlLength - 1) === '/') {
            url = url.substr(0, urlLength - 1);
        }
        var proxy = this._iframeProxies[url];
        if (!proxy) {
            var conversationId = OSF.OUtil.generateConversationId();
            var iframe = document.createElement("iframe");
            this._iframeProxyCount = this._iframeProxyCount + 1;
            var frameName = this._iframeNamePrefix + this._iframeProxyCount;
            iframe.setAttribute('id', frameName);
            iframe.setAttribute('name', frameName);
            var newUrl = url + "/_layouts/15/OfficeExtensionManager.aspx?" + conversationId;
            newUrl = OSF.OUtil.addXdmInfoAsHash(newUrl, conversationId + "|" + frameName + "|" + window.location.href);
            newUrl = OSF.OUtil.addSerializerVersionAsHash(newUrl, OSF.SerializerVersion.Browser);
            iframe.setAttribute('scrolling', 'auto');
            iframe.setAttribute('border', '0');
            iframe.setAttribute('width', '0');
            iframe.setAttribute('height', '0');
            iframe.setAttribute('style', "position: absolute; left: -100px; top:0px;");
            var me = this;
            var onIsProxyReadyCallback = function (errorCode, response) {
                var returnClientEndPoint;
                var proxy = me._iframeProxies[url];
                if (proxy && errorCode === 0 && response.status) {
                    proxy.isReady = true;
                    returnClientEndPoint = proxy.clientEndPoint;
                }
                else {
                    delete me._iframeProxies[url];
                    if (Microsoft.Office.Common.XdmCommunicationManager.getClientEndPoint(conversationId)) {
                        Microsoft.Office.Common.XdmCommunicationManager.deleteClientEndPoint(conversationId);
                        OSF.OUtil.removeEventListener(iframe, "load", onLoadCallback);
                        iframe.parentNode.removeChild(iframe);
                    }
                    else {
                        Telemetry.RuntimeTelemetryHelper.LogExceptionTag("Unexpected error occured with SharePoint proxy.", null, null, 0x007d4286);
                    }
                    returnClientEndPoint = null;
                }
                if (proxy) {
                    var pendingCallbackCount = proxy.pendingCallbacks.length;
                    for (var i = 0; i < pendingCallbackCount; i++) {
                        var currentCallback = proxy.pendingCallbacks.shift();
                        currentCallback(returnClientEndPoint);
                    }
                }
            };
            var onLoadCallback = function () {
                if (me._iframeProxies[url]) {
                    me._iframeProxies[url].clientEndPoint = Microsoft.Office.Common.XdmCommunicationManager.connect(conversationId, iframe.contentWindow, url);
                    me._iframeProxies[url].clientEndPoint.invoke("OEM_isProxyReady", onIsProxyReadyCallback, { __timeout__: 2000 });
                }
            };
            document.body.appendChild(iframe);
            OSF.OUtil.addEventListener(iframe, "load", onLoadCallback);
            iframe.setAttribute('src', newUrl);
            this._iframeProxies[url] = { "clientEndPoint": null, "isReady": false, "pendingCallbacks": [callback] };
        }
        else if (proxy.isReady) {
            callback(proxy.clientEndPoint);
        }
        else {
            proxy.pendingCallbacks.push(callback);
        }
    },
    _getClientVersionForOmex: function OSF_ContextActivationManager$_getClientVersionForOmex() {
        if (!this._appVersion) {
            return undefined;
        }
        var appVersion = this._appVersion.split('.');
        var major = parseInt(appVersion[0], 10);
        var minor = parseInt(appVersion[1], 10) || 0;
        if (major <= 15 && minor <= 0) {
            return undefined;
        }
        var fileVersion = OSF.Constants.FileVersion.split(".");
        return major + "." + minor + "." + fileVersion[2] + "." + fileVersion[3];
    },
    _getClientNameForOmex: function OSF_ContextActivationManager$_getClientNameForOmex() {
        return OSF.OmexClientNames[this._appName];
    },
    _getAppVersionForOmex: function OSF_ContextActivationManager$_getAppVersionForOmex() {
        return OSF.OmexAppVersions[this._appName];
    },
    _setCachedFlag: function OSF_ContextActivationManager$_setCachedFlag(cacheKey) {
        var osfLocalStorage = OSF.OUtil.getLocalStorage();
        if (osfLocalStorage) {
            osfLocalStorage.setItem(cacheKey, "true");
        }
    },
    _getCachedFlag: function OSF_ContextActivationManager$_getCachedFlag(cacheKey) {
        var osfLocalStorage = OSF.OUtil.getLocalStorage();
        if (osfLocalStorage) {
            var cacheValue = osfLocalStorage.getItem(cacheKey);
            return cacheValue ? true : false;
        }
    },
    _deleteCachedFlag: function OSF_ContextActivationManager$_deleteCachedFlag(cacheKey) {
        var osfLocalStorage = OSF.OUtil.getLocalStorage();
        if (osfLocalStorage) {
            osfLocalStorage.removeItem(cacheKey);
        }
    },
    _preloadOfficeJs: function OSF_ContextActivationManager$_preloadOfficeJs() {
        if (this._hasPreloadedOfficeJs) {
            return;
        }
        var preloadServiceScript = document.createElement("script");
        preloadServiceScript.src = OSF.OUtil.formatString("{0}?locale={1}&host={2}&version={3}", OSF.Constants.PreloadOfficeJsUrl, this._appUILocale, this._appName, this._hostSpecificFileVersion);
        preloadServiceScript.type = "text/javascript";
        preloadServiceScript.id = OSF.Constants.PreloadOfficeJsId;
        preloadServiceScript.onerror = function () {
            Telemetry.RuntimeTelemetryHelper.LogExceptionTag("Failed to connect to preload service.", null, null, 0x0085a2c2);
        };
        document.getElementsByTagName("head")[0].appendChild(preloadServiceScript);
        this._hasPreloadedOfficeJs = true;
    },
    isAddinCommandsAsync: function OSF_ContextActivationManager$isAddinCommandsAsync(addinCommandsCheckingContext, callback) {
        OSF.OUtil.validateParamObject(addinCommandsCheckingContext, {
            "marketplaceId": { type: String, mayBeNull: false },
            "marketplaceVersion": { type: String, mayBeNull: false },
            "storeType": { type: String, mayBeNull: false },
            "storeId": { type: String, mayBeNull: false },
            "targetData": { type: Object, mayBeNull: false }
        }, null);
        var storeTypeStr = addinCommandsCheckingContext.storeType.toLowerCase();
        var dataProvider;
        var context = {
            "contextActivationMgr": null,
            "referenceInUse": null
        };
        try {
            context.referenceInUse = { id: addinCommandsCheckingContext.marketplaceId, version: addinCommandsCheckingContext.marketplaceVersion, storeType: addinCommandsCheckingContext.storeType, storeLocator: addinCommandsCheckingContext.storeId };
            context.contextActivationMgr = this;
            if (OSF.StoreType.OMEX === storeTypeStr || OSF.StoreType.UploadFileDevCatalog === storeTypeStr || OSF.StoreType.SPCatalog === storeTypeStr) {
                var onGetManifestCompletedForIsAddinCommands = function (asyncResult) {
                    if (asyncResult != null && asyncResult.statusCode == OSF.ProxyCallStatusCode.Succeeded && asyncResult.value != null) {
                        callback(true, asyncResult.value._isAddinCommandsManifest(OSF.getManifestHostType(context.contextActivationMgr._hostType)), addinCommandsCheckingContext);
                    }
                    else {
                        callback(false, false, addinCommandsCheckingContext);
                    }
                };
                dataProvider = OfficeExt.AddinDataProviderFactory.createAddinDataProvider(storeTypeStr, this);
                if (dataProvider != null) {
                    dataProvider.getManifestAsync(context, onGetManifestCompletedForIsAddinCommands);
                }
                else {
                    callback(false, false, addinCommandsCheckingContext);
                }
            }
            else {
                callback(true, false, addinCommandsCheckingContext);
            }
        }
        catch (ex) {
            Telemetry.RuntimeTelemetryHelper.LogExceptionTag("Error for checking is addincommands or not", ex, null, 0x0109e4c4);
            callback(false, false, addinCommandsCheckingContext);
        }
    },
    getManifestInvaildErrorString: function OSF_ContextActivationManager$getManifestInvaildErrorString() {
        return Strings.OsfRuntime.L_AddinCommands_AddinNotSupported_Message;
    },
    getUserNameHashCode: function OSF_ContextActivationManager$getUserNameHashCode() {
        return this._userNameHashCode;
    }
};
OSF.OsfManifestManager = (function () {
    var _cachedManifests = {};
    var _UILocale = "en-us";
    var _pendingRequests = {};
    function _generateKey(marketplaceID, marketplaceVersion) {
        return marketplaceID + "_" + marketplaceVersion;
    }
    return {
        getManifestAsync: function OSF_OsfManifestManager$getManifestAsync(context, onCompleted) {
            OSF.OUtil.validateParamObject(context, {
                "osfControl": { type: Object, mayBeNull: false },
                "referenceInUse": { type: Object, mayBeNull: false }
            }, onCompleted);
            var reference = context.referenceInUse;
            var cacheKey = _generateKey(reference.id, reference.version);
            var manifest = _cachedManifests[cacheKey];
            context.manifestCached = false;
            if (manifest) {
                context.manifestCached = true;
                onCompleted({ "statusCode": OSF.ProxyCallStatusCode.Succeeded, "value": manifest, "context": context });
            }
            else if (context.clientEndPoint && context.manifestUrl) {
                Telemetry.AppLoadTimeHelper.ManifestRequestStart(context.osfControl._telemetryContext);
                var onRetrieveManifestCompleted = function (asyncResult) {
                    if (asyncResult.statusCode === OSF.ProxyCallStatusCode.Succeeded && asyncResult.value) {
                        var osfControl;
                        try {
                            osfControl = asyncResult.context.osfControl;
                            var manifestString;
                            if (typeof (asyncResult.value) === "string") {
                                asyncResult.context.manifestCached = true;
                                manifestString = asyncResult.value;
                            }
                            else {
                                asyncResult.context.manifestCached = asyncResult.value.cached;
                                manifestString = asyncResult.value.manifest;
                            }
                            Telemetry.AppLoadTimeHelper.SetManifestDataCachedFlag(osfControl._telemetryContext, asyncResult.value.cached);
                            asyncResult.value = new OSF.Manifest.Manifest(manifestString, osfControl._contextActivationMgr.getAppUILocale());
                            OSF.OsfManifestManager.cacheManifest(reference.id, reference.version, asyncResult.value);
                        }
                        catch (ex) {
                            asyncResult.value = null;
                            var appCorrelationId;
                            if (osfControl) {
                                appCorrelationId = osfControl._appCorrelationId;
                            }
                            OsfMsAjaxFactory.msAjaxDebug.trace("Invalid manifest in getManifestAsync: " + ex);
                            Telemetry.RuntimeTelemetryHelper.LogExceptionTag("Invalid manifest in getManifestAsync.", ex, appCorrelationId, 0x007d4287);
                        }
                    }
                    else {
                        Telemetry.RuntimeTelemetryHelper.LogExceptionTag("getManifestAsync failed:" + OSF.OUtil.serializeOMEXResponseErrorMessage(asyncResult), null, null, 0x010c53a1);
                    }
                    onCompleted(asyncResult);
                };
                var params = {
                    "manifestUrl": context.manifestUrl,
                    "id": reference.id,
                    "version": reference.version,
                    "clearCache": context.clearCache || false
                };
                this._invokeProxyMethodAsync(context, "OEM_getManifestAsync", onRetrieveManifestCompleted, params);
            }
            else {
                onCompleted({ "statusCode": OSF.ProxyCallStatusCode.Failed, "value": null, "context": context });
            }
        },
        getAppInstanceInfoByIdAsync: function OSF_OsfManifestManager$getAppInstanceInfoByIdAsync(context, onCompleted) {
            OSF.OUtil.validateParamObject(context, {
                "webUrl": { type: String, mayBeNull: false },
                "appInstanceId": { type: String, mayBeNull: false },
                "clientEndPoint": { type: Object, mayBeNull: false }
            }, onCompleted);
            var params = { "webUrl": context.webUrl, "appInstanceId": context.appInstanceId, "clearCache": context.clearCache || false };
            this._invokeProxyMethodAsync(context, "OEM_getSPAppInstanceInfoByIdAsync", onCompleted, params);
        },
        getSPTokenByProductIdAsync: function OSF_OsfManifestManager$getSPTokenByProductIdAsync(context, onCompleted) {
            OSF.OUtil.validateParamObject(context, {
                "appWebUrl": { type: String, mayBeNull: false },
                "productId": { type: String, mayBeNull: false }
            }, onCompleted);
            var me = this;
            var createSharePointProxyCompleted = function (clientEndPoint) {
                if (clientEndPoint) {
                    var params = { "webUrl": context.appWebUrl, "productId": context.productId, "clearCache": context.clearCache || false, "clientEndPoint": clientEndPoint };
                    me._invokeProxyMethodAsync(context, "OEM_getSPTokenByProductIdAsync", onCompleted, params);
                }
                else {
                    onCompleted({ "statusCode": OSF.ProxyCallStatusCode.ProxyNotReady, "value": null, "context": context });
                }
            };
            context.osfControl._contextActivationMgr._createSharePointIFrameProxy(context.appWebUrl, createSharePointProxyCompleted);
        },
        getSPAppEntitlementsAsync: function OSF_OsfManifestManager$getSPAppEntitlementsAsync(context, onCompleted) {
            OSF.OUtil.validateParamObject(context, {
                "osfControl": { type: Object, mayBeNull: false },
                "referenceInUse": { type: Object, mayBeNull: false },
                "baseUrl": { type: String, mayBeNull: false },
                "pageUrl": { type: String, mayBeNull: false },
                "webUrl": { type: String, mayBeNull: true }
            }, onCompleted);
            if (!context.webUrl) {
                var aElement = document.createElement('a');
                aElement.href = context.pageUrl;
                var pathName = aElement.pathname;
                var subPaths = pathName.split("/");
                var subPathCount = subPaths.length - 1;
                var path = aElement.href.substring(0, aElement.href.length - pathName.length);
                if (path && path.charAt(path.length - 1) !== '/') {
                    path += '/';
                }
                var paths = [path];
                for (var i = 0; i < subPathCount; i++) {
                    if (subPaths[i]) {
                        path = path + subPaths[i] + "/";
                        paths.push(path);
                    }
                }
                aElement = null;
                var me = this;
                var contextActivationMgr = context.osfControl._contextActivationMgr;
                var baseUrl = paths.pop();
                var onResolvePageUrlCompleted = function (asyncResult) {
                    if (asyncResult.statusCode === OSF.ProxyCallStatusCode.Succeeded) {
                        var resolvedUrl = asyncResult.value;
                        if (resolvedUrl && resolvedUrl.charAt(resolvedUrl.length - 1) !== '/') {
                            resolvedUrl += '/';
                        }
                        contextActivationMgr._webUrl = resolvedUrl;
                        asyncResult.context.webUrl = resolvedUrl;
                        asyncResult.context.appWebUrl = resolvedUrl;
                        me.getCorporateCatalogEntitlementsAsync(asyncResult.context, onCompleted);
                    }
                    else {
                        onCompleted(asyncResult);
                    }
                };
                var createSPAppProxyCompleted = function (clientEndPoint) {
                    if (clientEndPoint) {
                        context.clientEndPoint = clientEndPoint;
                        var params = { "pageUrl": context.pageUrl, "baseUrl": baseUrl, "clearCache": context.clearCache || false };
                        me._invokeProxyMethodAsync(context, "OEM_getSPAppWebUrlFromPageUrlAsync", onResolvePageUrlCompleted, params);
                    }
                    else if (paths.length > 0) {
                        baseUrl = paths.pop();
                        contextActivationMgr._createSharePointIFrameProxy(baseUrl, createSPAppProxyCompleted);
                    }
                    else {
                        onCompleted({ "statusCode": OSF.ProxyCallStatusCode.Failed, "value": null, "context": context });
                    }
                };
                contextActivationMgr._createSharePointIFrameProxy(baseUrl, createSPAppProxyCompleted);
            }
            else {
                this.getCorporateCatalogEntitlementsAsync(context, onCompleted);
            }
        },
        getOneDriveEntitlementsAsync: function (context, onCompleted) {
            var odDriveUrl = context.referenceInUse.storeLocator;
            var odManifestFolderUrl = context.referenceInUse.storeLocator + '/root:/Manifests';
            var accessToken = OSF.OneDriveOAuth.getAccessTokenFromCookie();
            if (accessToken == "" || accessToken == undefined) {
            }
            var accessTokenFullString = '?access_token=' + accessToken;
            var manifestFolderUrlWithAccessToken = odManifestFolderUrl + accessTokenFullString;
            var requestManifestsFolder = new XMLHttpRequest();
            var requestManifestsFilesList = new XMLHttpRequest();
            requestManifestsFolder.onreadystatechange = function () {
                if (requestManifestsFolder.readyState == 4) {
                    if (requestManifestsFolder.status == 200) {
                        var tmpJson = JSON.parse(requestManifestsFolder.response);
                        var manifestFolderId = tmpJson.id;
                        var manifestFilesListUrlWithAccessToken = context.referenceInUse.storeLocator + '/items/' + manifestFolderId + '/children' + accessTokenFullString;
                        requestManifestsFilesList.open('GET', manifestFilesListUrlWithAccessToken);
                        requestManifestsFilesList.setRequestHeader('Content-Type', 'json');
                        requestManifestsFilesList.send();
                    }
                    else {
                    }
                }
            };
            var oneDriveEntitlements = [];
            var requestCounts = 0;
            var requestCompleted = 0;
            var onCallBackComplete = function onCallBackComplete() {
                if (requestCounts == requestCompleted) {
                    onCompleted({
                        "statusCode": OSF.ProxyCallStatusCode.Succeeded,
                        "value": {
                            "entitlements": oneDriveEntitlements
                        },
                        "context": context
                    });
                }
            };
            var onSuccessPopulateGallery = function onSuccessPopulateGallery(manifestContent, requestedFileName) {
                requestCompleted++;
                var manifestFileName = requestedFileName;
                try {
                    var parsedManifest = new OSF.Manifest.Manifest(manifestContent.responseText);
                    var galleryItem = [];
                    galleryItem.push(parsedManifest.getDefaultDisplayName());
                    galleryItem.push(manifestFileName);
                    galleryItem.push(parsedManifest.getDefaultDescription());
                    galleryItem.push(parsedManifest.getTarget());
                    galleryItem.push(parsedManifest.getMarketplaceVersion());
                    galleryItem.push(manifestFileName);
                    galleryItem.push(OSF.StoreType.OneDrive);
                    galleryItem.push(parsedManifest.getDefaultWidth() || 0);
                    galleryItem.push(parsedManifest.getDefaultHeight() || 0);
                    galleryItem.push(parsedManifest.getDefaultIconUrl());
                    galleryItem.push(parsedManifest.getProviderName());
                    galleryItem.push(parsedManifest.getDefaultLocale());
                    galleryItem.push(OSF.StoreType.OneDrive);
                    oneDriveEntitlements.push(galleryItem);
                }
                catch (ex) {
                }
                onCallBackComplete();
            };
            var onErrorLogInfo = function onErrorLogInfo(errorString) {
                requestCompleted++;
                onCallBackComplete();
            };
            requestManifestsFilesList.onreadystatechange = function () {
                if (requestManifestsFilesList.readyState == 4 && requestManifestsFilesList.status == 200) {
                    var tmpJson = JSON.parse(requestManifestsFilesList.responseText);
                    for (var i = 0; i < tmpJson.value.length; i++) {
                        if (tmpJson.value[i].name.substr(tmpJson.value[i].name.length - 4).toLowerCase() == ".xml") {
                            var manifestFileUrlWithAccessToken = odDriveUrl + "/root:/Manifests/" + tmpJson.value[i].name + ":/content?access_token=" + accessToken;
                            OSF.OUtil.xhrGetFull(manifestFileUrlWithAccessToken, tmpJson.value[i].name, onSuccessPopulateGallery, onErrorLogInfo);
                            requestCounts++;
                        }
                    }
                    onCallBackComplete();
                }
            };
            requestManifestsFolder.open('GET', manifestFolderUrlWithAccessToken);
            requestManifestsFolder.setRequestHeader('Content-Type', 'json');
            requestManifestsFolder.send();
        },
        getCorporateCatalogEntitlementsAsync: function OSF_OsfManifestManager$getCorporateCatalogEntitlementsAsync(context, onCompleted) {
            OSF.OUtil.validateParamObject(context, {
                "osfControl": { type: Object, mayBeNull: false },
                "referenceInUse": { type: Object, mayBeNull: false },
                "webUrl": { type: String, mayBeNull: false }
            }, onCompleted);
            Telemetry.AppLoadTimeHelper.AuthenticationStart(context.osfControl._telemetryContext);
            var me = this;
            var retries = 0;
            var createSharePointProxyCompleted = function (clientEndPoint) {
                if (clientEndPoint) {
                    Telemetry.AppLoadTimeHelper.AuthenticationEnd(context.osfControl._telemetryContext);
                    context.clientEndPoint = clientEndPoint;
                    var params = {
                        "webUrl": context.webUrl,
                        "applicationName": OSF.HostCapability[context.hostType],
                        "officeExtentionTarget": (context.noTargetType || context.osfControl.getOsfControlType() === OSF.OsfControlTarget.TaskPane) ? null : context.osfControl.getOsfControlType(),
                        "clearCache": context.clearCache || false,
                        "supportedManifestVersions": {
                            "1.0": true,
                            "1.1": true
                        }
                    };
                    me._invokeProxyMethodAsync(context, "OEM_getEntitlementSummaryAsync", onCompleted, params);
                }
                else {
                    if (retries < OSF.Constants.AuthenticatedConnectMaxTries) {
                        retries++;
                        setTimeout(function () {
                            context.osfControl._contextActivationMgr._createSharePointIFrameProxy(context.webUrl, createSharePointProxyCompleted);
                        }, 500);
                    }
                    else {
                        onCompleted({
                            "statusCode": OSF.ProxyCallStatusCode.ProxyNotReady,
                            "value": null,
                            "context": context
                        });
                    }
                }
            };
            context.osfControl._contextActivationMgr._createSharePointIFrameProxy(context.webUrl, createSharePointProxyCompleted);
        },
        _invokeProxyMethodAsync: function OSF_OsfManifestManager$_invokeProxyMethodAsync(context, methodName, onCompleted, params) {
            var clientEndPointUrl = params.clientEndPoint ? params.clientEndPoint._targetUrl : context.clientEndPoint._targetUrl;
            var requestKeyParts = [clientEndPointUrl, methodName];
            var runtimeType;
            for (var p in params) {
                runtimeType = typeof params[p];
                if (runtimeType === "string" || runtimeType === "number" || runtimeType === "boolean") {
                    requestKeyParts.push(params[p]);
                }
            }
            var requestKey = requestKeyParts.join(".");
            var myPendingRequests = _pendingRequests;
            var newRequestHandler = { "onCompleted": onCompleted, "context": context, "methodName": methodName };
            var pendingRequestHandlers = myPendingRequests[requestKey];
            if (!pendingRequestHandlers) {
                myPendingRequests[requestKey] = [newRequestHandler];
                var onMethodCallCompleted = function (errorCode, response) {
                    var value = null;
                    var statusCode = OSF.ProxyCallStatusCode.Failed;
                    if (errorCode === 0 && response.status) {
                        value = response.result;
                        statusCode = OSF.ProxyCallStatusCode.Succeeded;
                    }
                    var currentPendingRequests = myPendingRequests[requestKey];
                    delete myPendingRequests[requestKey];
                    var pendingRequestHandlerCount = currentPendingRequests.length;
                    for (var i = 0; i < pendingRequestHandlerCount; i++) {
                        var currentRequestHandler = currentPendingRequests.shift();
                        var appCorrelationId;
                        try {
                            if (currentRequestHandler.context && currentRequestHandler.context.osfControl) {
                                appCorrelationId = currentRequestHandler.context.osfControl._appCorrelationId;
                            }
                            if (response && response.failureInfo) {
                                Telemetry.RuntimeTelemetryHelper.LogProxyFailure(appCorrelationId, currentRequestHandler.methodName, response.failureInfo);
                            }
                            currentRequestHandler.onCompleted({ "statusCode": statusCode, "value": value, "context": currentRequestHandler.context });
                        }
                        catch (ex) {
                            OsfMsAjaxFactory.msAjaxDebug.trace("_invokeProxyMethodAsync failed: " + ex);
                            Telemetry.RuntimeTelemetryHelper.LogExceptionTag("_invokeProxyMethodAsync failed.", ex, appCorrelationId, 0x007d4288);
                        }
                    }
                };
                var clientEndPoint = context.clientEndPoint;
                if (params.clientEndPoint) {
                    clientEndPoint = params.clientEndPoint;
                    delete params.clientEndPoint;
                }
                if (context.referenceInUse && context.referenceInUse.storeType === OSF.StoreType.OMEX) {
                    params.officeVersion = OSF.Constants.ThreePartsFileVersion;
                }
                clientEndPoint.invoke(methodName, onMethodCallCompleted, params);
            }
            else {
                pendingRequestHandlers.push(newRequestHandler);
            }
        },
        getOmexEntitlementsAsync: function OSF_OsfManifestManager$getOmexEntitlementsAsync(context, onCompleted) {
            OSF.OUtil.validateParamObject(context, {
                "referenceInUse": { type: Object, mayBeNull: false },
                "clientEndPoint": { type: Object, mayBeNull: false },
                "anonymous": { type: Boolean, mayBeNull: false }
            }, onCompleted);
            var params = {
                "applicationName": context.hostType,
                "appVersion": context.appVersion,
                "build": OSF.Constants.FileVersion,
                "clearEntitlement": context.clearCache || context.clearEntitlement || false,
                "clientName": context.clientName,
                "clientVersion": context.clientVersion,
                "correlationId": context.correlationId
            };
            this._invokeProxyMethodAsync(context, "OMEX_getEntitlementSummaryAsync", onCompleted, params);
        },
        getOmexAppDetailsAsync: function OSF_OsfManifestManager$getOmexAppDetailsAsync(context, onCompleted) {
            OSF.OUtil.validateParamObject(context, {
                "assetId": {
                    type: String,
                    mayBeNull: false
                },
                "contentMarket": {
                    type: String,
                    mayBeNull: false
                },
                "clientEndPoint": {
                    type: Object,
                    mayBeNull: true
                },
                "anonymous": {
                    type: Boolean,
                    mayBeNull: false
                },
                "clientName": {
                    type: String,
                    mayBeNull: false
                },
                "clientVersion": {
                    type: String,
                    mayBeNull: false
                }
            }, onCompleted);
            var params = {
                "assetid": "",
                "assetID": context.assetId,
                "contentMarket": context.contentMarket,
                "build": OSF.Constants.FileVersion,
                "clearCache": context.clearCache || context.clearEntitlement || false,
                "clientName": context.clientName,
                "clientVersion": context.clientVersion,
                "correlationId": context.correlationId
            };
            context.manifestManager = this;
            _omexDataProvider.GetAppDetails(context, params, onCompleted);
        },
        getOmexKilledAppsAsync: function OSF_OsfManifestManager$getOmexKilledAppsAsync(context, onCompleted) {
            OSF.OUtil.validateParamObject(context, {
                "osfControl": { type: Object, mayBeNull: false },
                "referenceInUse": { type: Object, mayBeNull: false },
                "clientEndPoint": { type: Object, mayBeNull: false },
                "anonymous": { type: Boolean, mayBeNull: false }
            }, onCompleted);
            var params = {
                "clearKilledApps": context.clearCache || context.clearKilledApps || false,
                "clientName": context.clientName,
                "clientVersion": context.clientVersion,
                "correlationId": context.correlationId
            };
            context.manifestManager = this;
            _omexDataProvider.GetKilledApps(context, params, onCompleted);
        },
        getOmexAppStateAsync: function OSF_OsfManifestManager$getOmexAppStateAsync(context, onCompleted) {
            OSF.OUtil.validateParamObject(context, {
                "osfControl": { type: Object, mayBeNull: false },
                "referenceInUse": { type: Object, mayBeNull: false },
                "clientEndPoint": { type: Object, mayBeNull: false },
                "anonymous": { type: Boolean, mayBeNull: true }
            }, onCompleted);
            var reference = context.referenceInUse;
            var params = {
                "assetID": reference.id,
                "contentMarket": reference.storeLocator,
                "clearAppState": context.clearCache || context.clearAppState || false,
                "clientName": context.clientName,
                "errors": {},
                "clientVersion": context.clientVersion,
                "correlationId": context.correlationId
            };
            context.manifestManager = this;
            _omexDataProvider.GetAppState(context, params, onCompleted);
            if (context.osfControl._telemetryContext) {
                Telemetry.AppLoadTimeHelper.SetAppStateDataInvalidFlag(context.osfControl._telemetryContext, params.errors["cacheExpired"]);
            }
        },
        getOmexManifestAndETokenAsync: function OSF_OsfManifestManager$getOmexManifestAndETokenAsync(context, onCompleted) {
            OSF.OUtil.validateParamObject(context, {
                "osfControl": { type: Object, mayBeNull: false },
                "referenceInUse": { type: Object, mayBeNull: false },
                "clientEndPoint": { type: Object, mayBeNull: false },
                "anonymous": { type: Boolean, mayBeNull: false }
            }, onCompleted);
            var reference = context.referenceInUse;
            var params = {
                "applicationName": context.hostType,
                "assetID": reference.id,
                "build": OSF.Constants.FileVersion,
                "clientName": context.clientName,
                "clientVersion": context.clientVersion,
                "clientAppInfoReturnType": context.clientAppInfoReturnType,
                "errors": {},
                "correlationId": context.correlationId
            };
            if (context.anonymous) {
                params.contentMarket = reference.storeLocator;
                params.clearManifest = context.clearCache || context.clearManifest || false;
            }
            else {
                params.assetContentMarket = context.osfControl._omexEntitlement.contentMarket;
                params.userContentMarket = reference.storeLocator;
                params.clearToken = context.clearCache || context.clearToken || false;
                params.clearManifest = context.clearCache || context.clearManifest || false;
                if (context.acceptedUpgrade) {
                    params.expectedVersion = context.expectedVersion;
                }
                else if (context.osfControl._omexEntitlement.hasEntitlement) {
                    params.expectedVersion = context.osfControl._omexEntitlement.version;
                }
            }
            context.manifestManager = this;
            _omexDataProvider.GetManifestAndEToken(context, params, onCompleted);
            if (context.osfControl._telemetryContext) {
                Telemetry.AppLoadTimeHelper.SetManifestDataInvalidFlag(context.osfControl._telemetryContext, params.errors["cacheExpired"]);
            }
        },
        removeOmexAppAsync: function OSF_OsfManifestManager$removeOmexAppAsync(context, onCompleted) {
            OSF.OUtil.validateParamObject(context, {
                "assetId": {
                    type: String,
                    mayBeNull: false
                },
                "clientName": {
                    type: String,
                    mayBeNull: false
                },
                "clientVersion": {
                    type: String,
                    mayBeNull: false
                }
            }, onCompleted);
            var params = {
                "assetID": context.assetId,
                "clientName": context.clientName,
                "clientVersion": context.clientVersion,
                "correlationId": context.correlationId
            };
            this._invokeProxyMethodAsync(context, "OMEX_removeAppAsync", onCompleted, params);
        },
        removeOmexCacheAsync: function OSF_OsfManifestManager$removeOmexCacheAsync(context, onCompleted) {
            OSF.OUtil.validateParamObject(context, {
                "osfControl": { type: Object, mayBeNull: false },
                "referenceInUse": { type: Object, mayBeNull: false },
                "clientEndPoint": { type: Object, mayBeNull: false }
            }, onCompleted);
            var reference = context.referenceInUse;
            var params = {
                "applicationName": context.hostType,
                "assetID": reference.id,
                "officeExtentionTarget": context.osfControl.getOsfControlType(),
                "clearEntitlement": context.clearEntitlement || false,
                "clearToken": context.clearToken || false,
                "clearAppState": context.clearAppState || false,
                "clearManifest": context.clearManifest || false,
                "appVersion": context.appVersion
            };
            if (context.anonymous) {
                params.contentMarket = reference.storeLocator;
            }
            else {
                params.assetContentMarket = context.osfControl._omexEntitlement.contentMarket;
                params.userContentMarket = reference.storeLocator;
            }
            this._invokeProxyMethodAsync(context, "OMEX_removeCacheAsync", onCompleted, params);
        },
        purgeManifest: function OSF_OsfManifestManager$purgeManifest(marketplaceID, marketplaceVersion) {
            var e = Function._validateParams(arguments, [
                { name: "marketplaceID", type: String, mayBeNull: false },
                { name: "marketplaceVersion", type: String, mayBeNull: false }
            ]);
            if (e)
                throw e;
            var cacheKey = _generateKey(marketplaceID, marketplaceVersion);
            if (typeof _cachedManifests[cacheKey] != "undefined") {
                delete _cachedManifests[cacheKey];
            }
        },
        cacheManifest: function OSF_OsfManifestManager$cacheManifest(marketplaceID, marketplaceVersion, manifest) {
            var e = Function._validateParams(arguments, [{ name: "marketplaceID", type: String, mayBeNull: false },
                { name: "marketplaceVersion", type: String, mayBeNull: false },
                { name: "manifest", type: Object, mayBeNull: false }
            ]);
            if (e)
                throw e;
            var cacheKey = _generateKey(marketplaceID, marketplaceVersion);
            manifest._UILocale = _UILocale;
            _cachedManifests[cacheKey] = manifest;
        },
        hasManifest: function OSF_OsfManifestManager$hasManifest(marketplaceID, marketplaceVersion) {
            var e = Function._validateParams(arguments, [
                { name: "marketplaceID", type: String, mayBeNull: false },
                { name: "marketplaceVersion", type: String, mayBeNull: false }
            ]);
            if (e)
                throw e;
            var cacheKey = _generateKey(marketplaceID, marketplaceVersion);
            if (typeof _cachedManifests[cacheKey] != "undefined")
                return true;
            return false;
        },
        getCachedManifest: function OSF_OsfManifestManager$getCachedManifest(marketplaceID, marketplaceVersion) {
            var e = Function._validateParams(arguments, [
                { name: "marketplaceID", type: String, mayBeNull: false },
                { name: "marketplaceVersion", type: String, mayBeNull: false }
            ]);
            if (e)
                throw e;
            var cacheKey = _generateKey(marketplaceID, marketplaceVersion);
            return _cachedManifests[cacheKey];
        },
        versionLessThan: function OSF_OsfManifestManager$versionLessThan(version1, version2) {
            var version1Parts = version1.split(".");
            var version2Parts = version2.split(".");
            var len = Math.min(version1Parts.length, version2Parts.length);
            var version1Part, version2Part, i;
            for (i = 0; i < len; i++) {
                try {
                    version1Part = parseFloat(version1Parts[i]);
                    version2Part = parseFloat(version2Parts[i]);
                    if (version1Part != version2Part) {
                        return version1Part < version2Part;
                    }
                }
                catch (ex) { }
            }
            if (version1Parts.length >= version2Parts.length) {
                return false;
            }
            else {
                len = version2Parts.length;
                var remainingSum = 0;
                for (i = version1Parts.length; i < len; i++) {
                    try {
                        version2Part = parseFloat(version2Parts[i]);
                    }
                    catch (ex) {
                        version2Part = 0;
                    }
                    remainingSum += version2Part;
                }
                return remainingSum > 0;
            }
        },
        _setUILocale: function (UILocale) { _UILocale = UILocale; }
    };
})();
OSF.InfoType = {
    Error: 0,
    Warning: 1,
    Information: 2,
    SecurityInfo: 3
};
OSF._ErrorUXHelper = function OSF__ErrorUXHelper(contextActivationManager) {
    var _contextActivationManager = contextActivationManager;
    OSF.OUtil.loadCSS(_contextActivationManager.getLocalizedCSSFilePath("moeerrorux.css"));
    var loadingImgInit = document.createElement("img");
    loadingImgInit.src = _contextActivationManager.getLocalizedImageFilePath("progress.gif");
    var statusTwoIconsImg = document.createElement("img");
    statusTwoIconsImg.src = _contextActivationManager.getLocalizedImageFilePath("moe_status_icons.png");
    var backgroundImgInit = document.createElement("img");
    backgroundImgInit.src = _contextActivationManager.getLocalizedImageFilePath("agavedefaulticon96x96.png");
    var _notificationQueues = {};
    var _highPriorityCount = 0;
    var _cleanupDiv = function (containerDiv) {
        var nodeCount = containerDiv.childNodes.length;
        var j = 0, node;
        while (j < nodeCount) {
            node = containerDiv.childNodes.item(j);
            if (node.tagName.toLowerCase() === "iframe") {
                j++;
            }
            else {
                containerDiv.removeChild(node);
                nodeCount--;
            }
        }
    };
    var _removeDOMElement = function (id) {
        var elm = document.getElementById(id);
        if (elm) {
            elm.parentNode.removeChild(elm);
        }
    };
    var _removeIConDiv = function (id) {
        OSF.OUtil.writeProfilerMark(OSF.NotificationUxPerfMarker.RemoveStage1Start);
        _removeDOMElement("icon_" + id);
        OSF.OUtil.writeProfilerMark(OSF.NotificationUxPerfMarker.RemoveStage1End);
    };
    var _removeInfoBarDiv = function (id, displayDeactive) {
        OSF.OUtil.writeProfilerMark(OSF.NotificationUxPerfMarker.RemoveStage2Start);
        var targetid;
        if (displayDeactive) {
            targetid = "moe-infobar-body_" + id;
        }
        else {
            targetid = "notificationbackground_" + id;
        }
        var isValidQueue = (_notificationQueues[id] && _notificationQueues[id].length > 0);
        if (isValidQueue && _notificationQueues[id][0].highPriority) {
            _highPriorityCount--;
        }
        _removeDOMElement(targetid);
        if (isValidQueue) {
            _notificationQueues[id].shift();
        }
        OSF.OUtil.writeProfilerMark(OSF.NotificationUxPerfMarker.RemoveStage2End);
    };
    var _showICon = function (params) {
        OSF.OUtil.writeProfilerMark(OSF.NotificationUxPerfMarker.RenderStage1Start);
        _cleanupDiv(params.div);
        var backgroundDiv = document.createElement('div');
        backgroundDiv.setAttribute("class", "moe-background");
        backgroundDiv.setAttribute("id", "icon_" + params.id);
        var statusIconImg = document.createElement("input");
        statusIconImg.setAttribute("id", "iconImg_" + params.id);
        statusIconImg.setAttribute("type", "image");
        statusIconImg.setAttribute("tabindex", "0");
        statusIconImg.src = _contextActivationManager.getLocalizedImageFilePath("moe_status_icons.png");
        var getIntoStage2 = function OSF__ErrorUXHelper_showICon$getIntoStage2(params) {
            params.sqmDWords[1] |= 2;
            _showInfoBar(params);
            _setControlFocusTrue(params.id);
        };
        statusIconImg.setAttribute("onclick", "getIntoStage2(params)");
        OSF.OUtil.attachClickHandler(statusIconImg, function () { getIntoStage2(params); });
        backgroundDiv.appendChild(statusIconImg);
        if (params.displayDeactive) {
            backgroundDiv.style.backgroundImage = "url(" + _contextActivationManager.getLocalizedImageFilePath("agavedefaulticon96x96.png") + ")";
            backgroundDiv.style.backgroundColor = 'white';
            backgroundDiv.style.opacity = '1';
            backgroundDiv.style.filter = 'alpha(opacity=100)';
            backgroundDiv.style.backgroundRepeat = "no-repeat";
            backgroundDiv.style.backgroundPosition = "center";
            backgroundDiv.style.height = '100%';
        }
        var className, id, altText;
        if (params.infoType === OSF.InfoType.Error) {
            className = "moe-status-error-icon";
            id = "iconImg_error_" + params.id;
            altText = Strings.OsfRuntime.L_InfobarIconErrorAccessibleName_TXT;
        }
        else if (params.infoType === OSF.InfoType.Warning) {
            className = "moe-status-warning-icon";
            id = "iconImg_warning_" + params.id;
            altText = Strings.OsfRuntime.L_InfobarIconWarningAccessibleName_TXT;
        }
        else if (params.infoType === OSF.InfoType.Information) {
            className = "moe-status-info-icon";
            id = "iconImg_info_" + params.id;
            altText = Strings.OsfRuntime.L_InfobarIconInfoAccessibleName_TXT;
        }
        else {
            className = "moe-status-secinfo-icon";
            id = "iconImg_secinfo_" + params.id;
            altText = Strings.OsfRuntime.L_InfobarIconSecInfoAccessibleName_TXT;
        }
        var re = new RegExp("MSIE ([0-9]{1,}[\.0-9]{0,})");
        if ((re.exec(navigator.userAgent) != null) && (parseFloat(RegExp.$1) == 9)) {
            className += "_ie";
        }
        statusIconImg.setAttribute("class", className);
        statusIconImg.setAttribute("id", id);
        statusIconImg.setAttribute("alt", altText);
        if (params.div.childNodes.length != 0) {
            params.div.insertBefore(backgroundDiv, params.div.childNodes[0]);
        }
        else {
            params.div.appendChild(backgroundDiv);
        }
        _focusOnNotificationUx(params.id);
        OSF.OUtil.writeProfilerMark(OSF.NotificationUxPerfMarker.RenderStage1End);
    };
    var _showInfoBar = function (params) {
        OSF.OUtil.writeProfilerMark(OSF.NotificationUxPerfMarker.RenderStage2Start);
        _cleanupDiv(params.div);
        var tooltipString = params.description;
        if (params.title.length > 100)
            params.title = params.title.substring(0, 99);
        if (params.description.length > 255)
            params.description = params.description.substring(0, 254);
        var infobarBodyId = "moe-infobar-body_" + params.id;
        var infoBarDiv = document.getElementById(infobarBodyId);
        if (infoBarDiv == undefined) {
            infoBarDiv = document.createElement('div');
            infoBarDiv.setAttribute("class", "moe-infobar-body");
            infoBarDiv.setAttribute("id", infobarBodyId);
        }
        var tooltipDiv = document.createElement("div");
        tooltipDiv.innerHTML = tooltipString;
        infoBarDiv.setAttribute("title", tooltipDiv.textContent);
        tooltipDiv = null;
        var infoTable = document.createElement('table');
        infoTable.setAttribute("class", "moe-infobar-infotable");
        infoTable.setAttribute("role", "presentation");
        var row, i;
        for (i = 0; i < 3; i++) {
            row = infoTable.insertRow(i);
            row.setAttribute("role", "presentation");
        }
        var infoTableRows = infoTable.rows;
        infoTableRows[0].insertCell(0);
        infoTableRows[0].insertCell(1);
        infoTableRows[0].insertCell(2);
        infoTableRows[0].cells[1].setAttribute("rowSpan", "2");
        infoTableRows[1].insertCell(0);
        infoTableRows[1].insertCell(1);
        infoTableRows[2].insertCell(0);
        infoTableRows[2].insertCell(1);
        infoTableRows[2].insertCell(2);
        infoTableRows[0].cells[0].setAttribute("class", "moe-infobar-top-left-cell");
        infoTableRows[0].cells[1].setAttribute("class", "moe-infobar-message-cell");
        infoTableRows[0].cells[2].setAttribute("class", "moe-infobar-top-right-cell");
        infoTableRows[2].cells[1].setAttribute("class", "moe-infobar-button-cell");
        var moeCommonImg = document.createElement("img");
        moeCommonImg.src = _contextActivationManager.getLocalizedImageFilePath("moe_status_icons.png");
        var className, altText;
        if (params.infoType === OSF.InfoType.Error) {
            className = "moe-infobar-error";
            altText = Strings.OsfRuntime.L_InfobarIconErrorAccessibleName_TXT;
        }
        else if (params.infoType === OSF.InfoType.Warning) {
            className = "moe-infobar-warning";
            altText = Strings.OsfRuntime.L_InfobarIconWarningAccessibleName_TXT;
        }
        else if (params.infoType === OSF.InfoType.Information) {
            className = "moe-infobar-info";
            altText = Strings.OsfRuntime.L_InfobarIconInfoAccessibleName_TXT;
        }
        else {
            className = "moe-infobar-secinfo";
            altText = Strings.OsfRuntime.L_InfobarIconSecInfoAccessibleName_TXT;
        }
        moeCommonImg.setAttribute("class", className);
        moeCommonImg.setAttribute("alt", altText);
        infoTableRows[0].cells[0].appendChild(moeCommonImg);
        var msgDiv = document.createElement("div");
        msgDiv.setAttribute("class", "moe-infobar-message-div");
        var titleSpan = document.createElement("span");
        titleSpan.setAttribute("class", "moe-infobar-title");
        titleSpan.innerHTML = params.title;
        var infobarMessageId = "moe-infobar-message_" + params.id;
        var descSpan = document.getElementById(infobarMessageId);
        if (descSpan == undefined) {
            descSpan = document.createElement("span");
            descSpan.setAttribute("class", "moe-infobar-message");
            descSpan.setAttribute("id", infobarMessageId);
        }
        descSpan.innerHTML = params.description;
        msgDiv.appendChild(titleSpan);
        msgDiv.appendChild(descSpan);
        infoTableRows[0].cells[1].appendChild(msgDiv);
        var logNotificationUls = function OSF__ErrorUXHelper__showInfoBar$logNotificationUls(params) {
            var osfControl = _contextActivationManager.getOsfControl(params.id);
            Telemetry.AppNotificationHelper.LogNotification(osfControl._appCorrelationId, params.sqmDWords[0], params.sqmDWords[1]);
        };
        var handleDismiss = function () {
            params.sqmDWords[1] |= 8;
            logNotificationUls(params);
            if (!params.reDisplay) {
                _removeInfoBarDiv(params.id, params.displayDeactive);
            }
            if (params.reDisplay) {
                _showNotification(params);
            }
            else if (_notificationQueues[params.id].length > 0) {
                var firstItem = _notificationQueues[params.id][0];
                _showNotification(firstItem);
            }
            if (params.dismissCallback) {
                params.dismissCallback();
            }
            _setControlFocusTrue(params.id);
        };
        var dismissIconImg = document.createElement("input");
        dismissIconImg.setAttribute("type", "image");
        dismissIconImg.setAttribute("src", _contextActivationManager.getLocalizedImageFilePath("moe_status_icons.png"));
        dismissIconImg.setAttribute("class", "moe-infobar-dismiss");
        dismissIconImg.setAttribute("id", "moe-infobar-dismiss_" + params.id);
        dismissIconImg.setAttribute("tabindex", "0");
        dismissIconImg.setAttribute("alt", Strings.OsfRuntime.L_InfobarIconCloseButtonAccessibleName_TXT);
        dismissIconImg.setAttribute("onclick", "handleDismiss();");
        OSF.OUtil.attachClickHandler(dismissIconImg, handleDismiss);
        infoTableRows[0].cells[2].appendChild(dismissIconImg);
        params.detailView = false;
        var button = document.createElement("button");
        button.setAttribute("class", "moe-infobar-button");
        button.innerHTML = params.buttonTxt;
        button.setAttribute("id", "moe-infobar-button_" + params.id);
        button.setAttribute("tabindex", "0");
        button.setAttribute("type", "button");
        if (params.buttonCallback) {
            var handleButtonClick = function () {
                params.sqmDWords[1] |= 4;
                logNotificationUls(params);
                _removeInfoBarDiv(params.id, false);
                if (_notificationQueues[params.id].length > 0) {
                    var firstItem = _notificationQueues[params.id][0];
                    _showNotification(firstItem);
                }
                if (params.retryAll === true) {
                    var osfControl = _contextActivationManager.getOsfControl(params.id);
                    osfControl._retryActivate = null;
                    _contextActivationManager.retryAll(osfControl._marketplaceID);
                }
                params.buttonCallback();
                _setControlFocusTrue(params.id);
            };
            button.setAttribute("onclick", "handleButtonClick()");
            OSF.OUtil.attachClickHandler(button, handleButtonClick);
        }
        else {
            button.setAttribute("onclick", "handleDismiss()");
            OSF.OUtil.attachClickHandler(button, handleDismiss);
        }
        infoTableRows[2].cells[1].appendChild(button);
        if (params.url) {
            var moreInfoButtonClick = function () {
                params.sqmDWords[1] |= 4;
                logNotificationUls(params);
                params.sqmDWords[1] = 1;
                window.open(params.url);
            };
            var moreInfoButton = document.createElement("button");
            moreInfoButton.setAttribute("class", "moe-infobar-button");
            moreInfoButton.innerHTML = params.urlButtonTxt ? params.urlButtonTxt : Strings.OsfRuntime.L_MoreInfoButton_TXT;
            moreInfoButton.setAttribute("id", "moe-infobar-button2_" + params.id);
            moreInfoButton.setAttribute("onclick", "moreInfoButtonClick()");
            OSF.OUtil.attachClickHandler(moreInfoButton, moreInfoButtonClick);
            moreInfoButton.setAttribute("tabindex", "0");
            moreInfoButton.setAttribute("type", "button");
            infoTableRows[2].cells[1].appendChild(moreInfoButton);
        }
        infoBarDiv.appendChild(infoTable);
        var backgroundDiv = document.createElement('div');
        backgroundDiv.setAttribute("class", "moe-background");
        backgroundDiv.setAttribute("id", "notificationbackground_" + params.id);
        if (params.displayDeactive) {
            backgroundDiv.style.backgroundImage = "url(" + _contextActivationManager.getLocalizedImageFilePath("agavedefaulticon96x96.png") + ")";
            backgroundDiv.style.backgroundColor = 'white';
            backgroundDiv.style.opacity = '1';
            backgroundDiv.style.filter = 'alpha(opacity=100)';
            backgroundDiv.style.backgroundRepeat = "no-repeat";
            backgroundDiv.style.backgroundPosition = "center";
            backgroundDiv.style.height = '100%';
        }
        backgroundDiv.appendChild(infoBarDiv);
        if (params.div.childNodes.length != 0) {
            params.div.insertBefore(backgroundDiv, params.div.childNodes[0]);
        }
        else {
            params.div.appendChild(backgroundDiv);
        }
        _focusOnNotificationUx(params.id);
        OSF.OUtil.writeProfilerMark(OSF.NotificationUxPerfMarker.RenderStage2End);
    };
    var _setControlFocusTrue = function (id) {
        var osfControl = _contextActivationManager.getOsfControl(id);
        if (osfControl) {
            osfControl._controlFocus = true;
        }
    };
    var _focusOnNotificationUx = function (id) {
        var osfControl = _contextActivationManager.getOsfControl(id);
        if (osfControl) {
            if (osfControl._controlFocus) {
                if (_notificationQueues[id] && _notificationQueues[id].length > 0) {
                    var topItem = _notificationQueues[id][0];
                    if (topItem && topItem.div) {
                        var list = topItem.div.querySelectorAll('input,a,button');
                        if (list && list.length > 0) {
                            var item;
                            if (list.length === 1) {
                                item = list[0];
                            }
                            else {
                                item = list[1];
                            }
                            if (item instanceof HTMLElement) {
                                item.focus();
                                if (_contextActivationManager._notifyHost) {
                                    _contextActivationManager._notifyHost(id, OSF.AgaveHostAction.SelectWithError);
                                }
                            }
                        }
                    }
                }
            }
        }
    };
    var _showNotification = function (params) {
        if (params.detailView == undefined || params.detailView === false) {
            params.sqmDWords[1] = 0;
            _showICon(params);
        }
        else {
            params.sqmDWords[1] = 1;
            _showInfoBar(params);
        }
    };
    var _dismissMessages = function (id) {
        if (_notificationQueues[id]) {
            if (_notificationQueues[id].length > 0) {
                var agaveDiv = _notificationQueues[id][0].div;
                _cleanupDiv(agaveDiv);
            }
            delete _notificationQueues[id];
        }
    };
    var _getHTMLEncodedString = function (str) {
        var div = document.createElement('div');
        var textNode = document.createTextNode(str);
        div.appendChild(textNode);
        return div.innerHTML;
    };
    return {
        showProgress: function OSF__ErrorUXHelper$showProgress(div, id) {
            try {
                OSF.OUtil.writeProfilerMark(OSF.NotificationUxPerfMarker.RenderLoadingAnimationStart);
                var progressDiv = document.getElementById("progress_" + id);
                if (!progressDiv) {
                    _notificationQueues[id] = [];
                    var backgroundDiv = document.createElement('div');
                    backgroundDiv.setAttribute("class", "moe-background");
                    backgroundDiv.setAttribute("id", "progress_" + id);
                    backgroundDiv.style.backgroundColor = 'rgba(255, 255, 255, 0.5)';
                    backgroundDiv.style.opacity = '1';
                    backgroundDiv.style.filter = 'alpha(opacity=100)';
                    backgroundDiv.style.height = '100%';
                    var loadingDiv = document.createElement('div');
                    loadingDiv.style.width = "100%";
                    loadingDiv.style.height = "100%";
                    loadingDiv.style.backgroundImage = "url(" + _contextActivationManager.getLocalizedImageFilePath("progress.gif") + ")";
                    loadingDiv.style.backgroundRepeat = "no-repeat";
                    loadingDiv.style.backgroundPosition = "center";
                    backgroundDiv.appendChild(loadingDiv);
                    div.appendChild(backgroundDiv);
                }
                OSF.OUtil.writeProfilerMark(OSF.NotificationUxPerfMarker.RenderLoadingAnimationEnd);
            }
            catch (ex) { }
        },
        showNotification: function OSF__ErrorUXHelper$showNotification(params) {
            params.sqmDWords = [params.errorCode, 0];
            delete params.errorCode;
            if (params.highPriority == undefined) {
                params.highPriority = params.infoType === OSF.InfoType.Error ? true : false;
            }
            if (params.reDisplay == undefined) {
                params.reDisplay = params.infoType === OSF.InfoType.Error ? true : false;
            }
            var notificationQueue = _notificationQueues[params.id];
            if (notificationQueue === undefined) {
                notificationQueue = [];
                _notificationQueues[params.id] = notificationQueue;
            }
            if (params.highPriority === false) {
                notificationQueue.push(params);
            }
            else {
                notificationQueue.splice(_highPriorityCount, 0, params);
                _highPriorityCount++;
            }
            if (_notificationQueues[params.id].length === 1 || params.highPriority) {
                _showNotification(notificationQueue[0]);
            }
        },
        showICon: function OSF__ErrorUXHelper$showICon(params) {
            _showICon(params);
        },
        removeDOMElement: function OSF__ErrorUXHelper$removeDOMElement(id) {
            _removeDOMElement(id);
        },
        removeProgressDiv: function OSF__ErrorUXHelper$revmoveProgressDiv(containerDiv, id) {
            OSF.OUtil.writeProfilerMark(OSF.NotificationUxPerfMarker.RemoveLoadingAnimationStart);
            var progressDiv = containerDiv.ownerDocument.getElementById("progress_" + id);
            if (progressDiv) {
                _cleanupDiv(containerDiv);
            }
            OSF.OUtil.writeProfilerMark(OSF.NotificationUxPerfMarker.RemoveLoadingAnimationEnd);
        },
        removeIConDiv: function OSF__ErrorUXHelper$revmoveIConDiv(id) {
            _removeIConDiv(id);
        },
        removeInfoBarDiv: function OSF__ErrorUXHelper$revmoveInfoBarDiv(id, displayDeactive) {
            _removeInfoBarDiv(id, displayDeactive);
        },
        dismissMessages: function OSF_ErrorUXHelper$dismissMessages(id) {
            _dismissMessages(id);
        },
        getHTMLEncodedString: function OSF_ErrorUXHelper$getHTMLEncodedString(str) {
            return _getHTMLEncodedString(str);
        },
        purgeOsfControlNotification: function OSF_ErrorUXHelper$purgeOsfControlNotification(id) {
            var queue = _notificationQueues[id];
            var osfControl = _contextActivationManager.getOsfControl(id);
            if (queue && queue.length > 0 && osfControl) {
                Telemetry.AppNotificationHelper.LogNotification(id, queue[0].sqmDWords[0], queue[0].sqmDWords[1]);
            }
        },
        focusOnNotificationUx: function OSF_ErrorUXHelper$focusOnNotificationUx(id) {
            _setControlFocusTrue(id);
            _focusOnNotificationUx(id);
        },
        appHasNotifications: function OSF_ErrorUXHelper$appHasNotifications(id) {
            return (_notificationQueues[id] && _notificationQueues[id].length > 0);
        }
    };
};
OSF.RequirementSetNames = {
    "ActiveView": "ActiveView",
    "BindingEvents": "BindingEvents",
    "CompressedFile": "CompressedFile",
    "CustomXmlParts": "CustomXmlParts",
    "DocumentEvents": "DocumentEvents",
    "ExcelApi": "ExcelApi",
    "File": "File",
    "HtmlCoercion": "HtmlCoercion",
    "ImageCoercion": "ImageCoercion",
    "Mailbox": "Mailbox",
    "MatrixBindings": "MatrixBindings",
    "MatrixCoercion": "MatrixCoercion",
    "OoxmlCoercion": "OoxmlCoercion",
    "PartialTableBindings": "PartialTableBindings",
    "PdfFile": "PdfFile",
    "Selection": "Selection",
    "Settings": "Settings",
    "TableBindings": "TableBindings",
    "TableCoercion": "TableCoercion",
    "TextBindings": "TextBindings",
    "TextCoercion": "TextCoercion",
    "TextFile": "TextFile",
    "WordApi": "WordApi"
};
OSF.RequirementMethodNames = {
    "Binding.addHandlerAsync": "Binding.addHandlerAsync",
    "Binding.removeHandlerAsync": "Binding.removeHandlerAsync",
    "Bindings.addFromNamedItemAsync": "Bindings.addFromNamedItemAsync",
    "Bindings.addFromPromptAsync": "Bindings.addFromPromptAsync",
    "Bindings.addFromSelectionAsync": "Bindings.addFromSelectionAsync",
    "Bindings.getAllAsync": "Bindings.getAllAsync",
    "Bindings.getByIdAsync": "Bindings.getByIdAsync",
    "Bindings.releaseByIdAsync": "Bindings.releaseByIdAsync",
    "CustomXmlParts.addAsync": "CustomXmlParts.addAsync",
    "CustomXmlParts.getByIdAsync": "CustomXmlParts.getByIdAsync",
    "CustomXmlParts.getByNamespaceAsync": "CustomXmlParts.getByNamespaceAsync",
    "CustomXmlPart.getXmlAsync": "CustomXmlPart.getXmlAsync",
    "CustomXmlPart.getNodesAsync": "CustomXmlPart.getNodesAsync",
    "CustomXmlPart.deleteAsync": "CustomXmlPart.deleteAsync",
    "CustomXmlNode.getNodesAsync": "CustomXmlNode.getNodesAsync",
    "CustomXmlNode.getNodeValueAsync": "CustomXmlNode.getNodeValueAsync",
    "CustomXmlNode.getXmlAsync": "CustomXmlNode.getXmlAsync",
    "CustomXmlNode.setNodeValueAsync": "CustomXmlNode.setNodeValueAsync",
    "CustomXmlNode.setXmlAsync": "CustomXmlNode.setXmlAsync",
    "CustomXmlPrefixMappings.addNamespaceAsync": "CustomXmlPrefixMappings.addNamespaceAsync",
    "CustomXmlPrefixMappings.getNamespaceAsync": "CustomXmlPrefixMappings.getNamespaceAsync",
    "CustomXmlPrefixMappings.getPrefixAsync": "CustomXmlPrefixMappings.getPrefixAsync",
    "Document.addHandlerAsync": "Document.addHandlerAsync",
    "Document.getActiveViewAsync": "Document.getActiveViewAsync",
    "Document.getFileAsync": "Document.getFileAsync",
    "Document.getFilePropertiesAsync": "Document.getFilePropertiesAsync",
    "Document.getSelectedDataAsync": "Document.getSelectedDataAsync",
    "Document.goToByIdAsync": "Document.goToByIdAsync",
    "Document.removeHandlerAsync": "Document.removeHandlerAsync",
    "Document.setSelectedDataAsync": "Document.setSelectedDataAsync",
    "File.closeAsync": "File.closeAsync",
    "File.getSliceAsync": "File.getSliceAsync",
    "MatrixBinding.getDataAsync": "MatrixBinding.getDataAsync",
    "MatrixBinding.setDataAsync": "MatrixBinding.setDataAsync",
    "Settings.addHandlerAsync": "Settings.addHandlerAsync",
    "Settings.get": "Settings.get",
    "Settings.refreshAsync": "Settings.refreshAsync",
    "Settings.remove": "Settings.remove",
    "Settings.removeHandlerAsync": "Settings.removeHandlerAsync",
    "Settings.saveAsync": "Settings.saveAsync",
    "Settings.set": "Settings.set",
    "TableBinding.addColumnsAsync": "TableBinding.addColumnsAsync",
    "TableBinding.addRowsAsync": "TableBinding.addRowsAsync",
    "TableBinding.clearFormatsAsync": "TableBinding.clearFormatsAsync",
    "TableBinding.deleteAllDataValuesAsync": "TableBinding.deleteAllDataValuesAsync",
    "TableBinding.getDataAsync": "TableBinding.getDataAsync",
    "TableBinding.setDataAsync": "TableBinding.setDataAsync",
    "TableBinding.setFormatsAsync": "TableBinding.setFormatsAsync",
    "TableBinding.setTableOptionsAsync": "TableBinding.setTableOptionsAsync",
    "TextBinding.getDataAsync": "TextBinding.getDataAsync",
    "TextBinding.setDataAsync": "TextBinding.setDataAsync",
    "CustomXmlPart.addHandlerAsync": "CustomXmlPart.addHandlerAsync",
    "CustomXmlPart.removeHandlerAsync": "CustomXmlPart.removeHandlerAsync"
};
OSF.DataPointNames = {
    AppManagementMenu: "appmanagementmenu",
    InsertionDialogSession: "insertiondialogsession",
    UploadFileDevCatelog: "uploadfiledevcatelog",
    UploadFileDevCatalogUsage: "uploadfiledevcatalogusage"
};
OSF.AppManagementAction = {
    Cancel: 0,
    AppDetails: 1,
    RateReview: 2,
    Remove: 3
};
OSF.UploadFileDevCatelogAction = {
    OpenUploadFileDialog: 0,
    Install: 1,
    Cancel: 2
};
OSF.OUtil.normalizeAppVersion = function OSF_WEF$normalizeAppVersion(version) {
    var items = version.split('.');
    var appVersion = version;
    for (var i = 0; i < 4 - items.length; i++) {
        appVersion += ".0";
    }
    return appVersion;
};
Microsoft.Office.WebExtension.EventType = {
    BindingSelectionChanged: "bindingSelectionChanged",
    BindingDataChanged: "bindingDataChanged"
};
Microsoft.Office.WebExtension.CoercionType = {
    Text: "text",
    Matrix: "matrix",
    Table: "table"
};
OSF.DDA.PropertyDescriptors.Subset = "subset";
Microsoft.Office.WebExtension.Table = {
    ThisRow: "thisRow"
};
OSF.AccessWebAppPolicyManager = function OSF_AccessWebAppPolicyManager(contextActivationManager, registeredActions) {
    var e = Function._validateParams(arguments, [
        { name: "contextActivationMgr", type: Object, mayBeNull: false },
        { name: "registeredActions", type: Object, mayBeNull: false }
    ]);
    if (e)
        throw e;
    this._contextActivationManager = contextActivationManager;
    this._registeredActions = {};
    var isValidPermission;
    var registeredPermission;
    for (var actionName in registeredActions) {
        isValidPermission = true;
        registeredPermission = registeredActions[actionName];
        if (typeof registeredPermission === "object") {
            for (var dispId in registeredPermission) {
                isValidPermission = this._validatePermissionInput(registeredPermission[dispId]);
                if (!isValidPermission) {
                    break;
                }
            }
        }
        else {
            isValidPermission = this._validatePermissionInput(registeredPermission);
        }
        if (!isValidPermission) {
            throw OsfMsAjaxFactory.msAjaxError.argument("registeredActions");
        }
        this._registeredActions[actionName] = registeredPermission;
    }
};
OSF.AccessWebAppPolicyManager.prototype = {
    checkPermission: function OSF_AccessWebAppPolicyManager$checkPermission(conversationID, actionName, params) {
        var permissionNeeded = this._registeredActions[actionName];
        if (!permissionNeeded) {
            return false;
        }
        if (typeof permissionNeeded === "object") {
            if (params && params.DdaMethod && params.DdaMethod.DispatchId && permissionNeeded[params.DdaMethod.DispatchId]) {
                permissionNeeded = permissionNeeded[params.DdaMethod.DispatchId];
            }
            else {
                return false;
            }
        }
        if (permissionNeeded === OSF.OsfControlPermission.Restricted) {
            return true;
        }
        var entry = this._contextActivationManager._getManifestAndTargetByConversationId(conversationID);
        if (entry && entry.manifest) {
            return entry.manifest.hasPermission(permissionNeeded);
        }
        else {
            return false;
        }
    },
    _validatePermissionInput: function OSF_AccessWebAppPolicyManager$_validatePermissionInput(permission) {
        var isValidPermission = false;
        for (var permissionName in OSF.OsfControlPermission) {
            if (OSF.OsfControlPermission[permissionName] === permission) {
                isValidPermission = true;
                break;
            }
        }
        return isValidPermission;
    }
};
OSF.AppSpecificSetup.AccessFacade = function OSF_AppSpecificSetup_AccessFacade(wnd, ewaControl) {
    var _ewaControl = ewaControl;
    var _me = this;
    var _init = function (wnd) {
        if (_ewaControl != undefined)
            return true;
        for (var i = 0; i < wnd.frames.length; i++) {
            var success = _init(wnd.frames[i]);
            if (success)
                return true;
        }
        return false;
    };
    this.executeMethod = function (params, callback) {
        if (params) {
            var dispId = params["DdaMethod"]["DispatchId"];
            if (dispId === OSF.DDA.MethodDispId.dispidGetBindingDataMethod) {
                var getBindingDataParams = params["DdaGetBindingData"];
                if (getBindingDataParams && getBindingDataParams["RowCount"] === 0) {
                    delete getBindingDataParams["RowCount"];
                }
                if (getBindingDataParams && getBindingDataParams["ColCount"] === 0) {
                    delete getBindingDataParams["ColCount"];
                }
            }
            else if (dispId == OSF.DDA.MethodDispId.dispidAddRowsMethod ||
                dispId == OSF.DDA.MethodDispId.dispidAddColumnsMethod) {
                var addTableData = params["DdaAddRowsColumns"];
            }
        }
        _ewaControl.invokeDdaMethod(params, function onComplete(payload) {
            if (payload["Error"] == 0) {
                var dispId = params["DdaMethod"]["DispatchId"];
                if (dispId == OSF.DDA.MethodDispId.dispidAddBindingFromSelectionMethod ||
                    dispId == OSF.DDA.MethodDispId.dispidAddBindingFromPromptMethod ||
                    dispId == OSF.DDA.MethodDispId.dispidGetBindingMethod ||
                    dispId == OSF.DDA.MethodDispId.dispidGetAllBindingsMethod) {
                    var registeredActions = _me.serviceEndpoint._policyManager._registeredActions;
                    for (var binding in payload["Bindings"]) {
                        var bindingId = payload["Bindings"][binding]["Name"];
                        var selectionChanged = OSF.DDA.getXdmEventName(bindingId, Microsoft.Office.WebExtension.EventType.BindingSelectionChanged);
                        if (!registeredActions[selectionChanged]) {
                            _me.serviceEndpoint.registerEventEx(selectionChanged, _me.registerEvent, Microsoft.Office.Common.InvokeType.asyncRegisterEvent, _me.unregisterEvent, Microsoft.Office.Common.InvokeType.asyncUnregisterEvent);
                            registeredActions[selectionChanged] = OSF.OsfControlPermission.Restricted;
                        }
                        var dataChanged = OSF.DDA.getXdmEventName(bindingId, Microsoft.Office.WebExtension.EventType.BindingDataChanged);
                        if (!registeredActions[dataChanged]) {
                            _me.serviceEndpoint.registerEventEx(dataChanged, _me.registerEvent, Microsoft.Office.Common.InvokeType.asyncRegisterEvent, _me.unregisterEvent, Microsoft.Office.Common.InvokeType.asyncUnregisterEvent);
                            registeredActions[dataChanged] = OSF.OsfControlPermission.Restricted;
                        }
                    }
                }
            }
            if (callback) {
                callback(payload);
            }
        }, null);
    };
    this.registerEvent = function (eventHandler, callback, params) {
        _ewaControl.registerDdaEventAsync(params["eventDispId"], params["controlId"], params["targetId"], function onEvent(eventArgs) {
            if (eventArgs.dispid == OSF.DDA.EventDispId.dispidBindingSelectionChangedEvent) {
                var subset = {};
                subset[Microsoft.Office.WebExtension.Parameters.StartRow] = eventArgs[Microsoft.Office.WebExtension.Parameters.StartRow];
                subset[Microsoft.Office.WebExtension.Parameters.StartColumn] = eventArgs[Microsoft.Office.WebExtension.Parameters.StartColumn];
                subset[Microsoft.Office.WebExtension.Parameters.RowCount] = eventArgs[Microsoft.Office.WebExtension.Parameters.RowCount];
                subset[Microsoft.Office.WebExtension.Parameters.ColumnCount] = eventArgs[Microsoft.Office.WebExtension.Parameters.ColumnCount];
                eventArgs[OSF.DDA.PropertyDescriptors.Subset] = subset;
            }
            if (eventHandler) {
                eventHandler(eventArgs);
            }
        }, function onRegister(asyncResult) {
            if (callback) {
                callback(asyncResult["Error"] == 0);
            }
        });
    };
    this.unregisterEvent = function (eventHandler, callback, params) {
        _ewaControl.unregisterDdaEventAsync(params["eventDispId"], params["controlId"], params["targetId"], function onUnregister(asyncResult) {
            if (callback) {
                callback(asyncResult["Error"] == 0);
            }
        });
    };
    _init(wnd);
};
OSF.AppSpecificSetup._setupFacade = function OSF_AppSpecificSetup__setupFacade(hostControl, contextActivationManager, serviceEndPoint, serviceEndPointInternal) {
    var accFacade = new OSF.AppSpecificSetup.AccessFacade(window, hostControl);
    accFacade.serviceEndpoint = serviceEndPoint;
    var registeredActions = {
        'ContextActivationManager_getAppContextAsync': OSF.OsfControlPermission.Restricted,
        'ContextActivationManager_notifyHost': OSF.OsfControlPermission.Restricted,
        'ContextActivationManager_openWindowInHost': OSF.OsfControlPermission.Restricted,
        'documentSelectionChanged': OSF.OsfControlPermission.ReadDocument,
        'settingsChanged': OSF.OsfControlPermission.Restricted,
        'executeMethod': {}
    };
    registeredActions.executeMethod[OSF.DDA.MethodDispId.dispidGetSelectedDataMethod] = OSF.OsfControlPermission.ReadDocument;
    registeredActions.executeMethod[OSF.DDA.MethodDispId.dispidSetSelectedDataMethod] = OSF.OsfControlPermission.WriteDocument;
    registeredActions.executeMethod[OSF.DDA.MethodDispId.dispidAddBindingFromNamedItemMethod] = OSF.OsfControlPermission.ReadWriteDocument;
    registeredActions.executeMethod[OSF.DDA.MethodDispId.dispidAddBindingFromSelectionMethod] = OSF.OsfControlPermission.ReadWriteDocument;
    registeredActions.executeMethod[OSF.DDA.MethodDispId.dispidAddBindingFromPromptMethod] = OSF.OsfControlPermission.ReadWriteDocument;
    registeredActions.executeMethod[OSF.DDA.MethodDispId.dispidGetBindingMethod] = OSF.OsfControlPermission.ReadWriteDocument;
    registeredActions.executeMethod[OSF.DDA.MethodDispId.dispidGetAllBindingsMethod] = OSF.OsfControlPermission.ReadWriteDocument;
    registeredActions.executeMethod[OSF.DDA.MethodDispId.dispidReleaseBindingMethod] = OSF.OsfControlPermission.ReadWriteDocument;
    registeredActions.executeMethod[OSF.DDA.MethodDispId.dispidGetBindingDataMethod] = OSF.OsfControlPermission.ReadWriteDocument;
    registeredActions.executeMethod[OSF.DDA.MethodDispId.dispidSetBindingDataMethod] = OSF.OsfControlPermission.ReadWriteDocument;
    registeredActions.executeMethod[OSF.DDA.MethodDispId.dispidLoadSettingsMethod] = OSF.OsfControlPermission.Restricted;
    registeredActions.executeMethod[OSF.DDA.MethodDispId.dispidSaveSettingsMethod] = OSF.OsfControlPermission.Restricted;
    registeredActions.executeMethod[OSF.DDA.MethodDispId.dispidAddRowsMethod] = OSF.OsfControlPermission.ReadWriteDocument;
    registeredActions.executeMethod[OSF.DDA.MethodDispId.dispidClearAllRowsMethod] = OSF.OsfControlPermission.ReadWriteDocument;
    registeredActions.executeMethod[OSF.DDA.MethodDispId.dispidAddColumnsMethod] = OSF.OsfControlPermission.ReadWriteDocument;
    var policyManager = new OSF.AccessWebAppPolicyManager(contextActivationManager, registeredActions);
    serviceEndPoint.setPolicyManager(policyManager);
    serviceEndPoint.registerMethod("executeMethod", accFacade.executeMethod, Microsoft.Office.Common.InvokeType.async, false);
    serviceEndPoint.registerEventEx("documentSelectionChanged", accFacade.registerEvent, Microsoft.Office.Common.InvokeType.asyncRegisterEvent, accFacade.unregisterEvent, Microsoft.Office.Common.InvokeType.asyncUnregisterEvent);
    serviceEndPoint.registerEventEx("settingsChanged", accFacade.registerEvent, Microsoft.Office.Common.InvokeType.asyncRegisterEvent, accFacade.unregisterEvent, Microsoft.Office.Common.InvokeType.asyncUnregisterEvent);
    var registeredActionsInternal = {
        'ContextActivationManager_onClickInsertOsfControl': OSF.OsfControlPermission.Restricted,
        'ContextActivationManager_onClickCancelDialog': OSF.OsfControlPermission.Restricted,
        'ContextActivationManager_getEntitlementsForInsertDialog': OSF.OsfControlPermission.Restricted,
        'ContextActivationManager_invokeSignIn': OSF.OsfControlPermission.Restricted,
        'ContextActivationManager_invokeWindowOpen': OSF.OsfControlPermission.Restricted,
        'ContextActivationManager_getOmexData': OSF.OsfControlPermission.Restricted
    };
    var policyManagerInternal = new OSF.AccessWebAppPolicyManager(contextActivationManager, registeredActionsInternal);
    serviceEndPointInternal.setPolicyManager(policyManagerInternal);
    contextActivationManager.getLocalizedImageFilePath = function (fileName) {
        return this.getLocalizedImagesUrl() + fileName;
    };
    contextActivationManager.getLocalizedCSSFilePath = function (fileName) {
        return this.getLocalizedStylesUrl() + fileName;
    };
    contextActivationManager._hostType = OSF.HostType.Access;
    contextActivationManager._hostPlatform = OSF.HostPlatform.Web;
    contextActivationManager._hostSpecificFileVersion = OSF.HostSpecificFileVersionMap["access"]["web"];
    ;
    var supportedCapabilities = [OSF.Capability.Database];
    var supportedHosts = [OSF.Capability.Database];
    var supportedSets = [{
            name: OSF.RequirementSetNames.BindingEvents
        }, {
            name: OSF.RequirementSetNames.PartialTableBindings
        }, {
            name: OSF.RequirementSetNames.Settings
        }, {
            name: OSF.RequirementSetNames.TableBindings
        }, {
            name: OSF.RequirementSetNames.TableCoercion
        }];
    var methods = OSF.RequirementMethodNames;
    var supportedMethods = [
        methods["Binding.addHandlerAsync"],
        methods["Binding.removeHandlerAsync"],
        methods["Bindings.addFromNamedItemAsync"],
        methods["Bindings.addFromPromptAsync"],
        methods["Bindings.addFromSelectionAsync"],
        methods["Bindings.getAllAsync"],
        methods["Bindings.getByIdAsync"],
        methods["Bindings.releaseByIdAsync"],
        methods["Settings.addHandlerAsync"],
        methods["Settings.get"],
        methods["Settings.refreshAsync"],
        methods["Settings.remove"],
        methods["Settings.removeHandlerAsync"],
        methods["Settings.saveAsync"],
        methods["Settings.set"],
        methods["TableBinding.addColumnsAsync"],
        methods["TableBinding.addRowsAsync"],
        methods["TableBinding.deleteAllDataValuesAsync"],
        methods["TableBinding.getDataAsync"],
        methods["TableBinding.setDataAsync"]
    ];
    var supportedControlTarget = [OSF.OsfControlTarget.InContent];
    var requirementsChecker = new OSF.RequirementsChecker(supportedCapabilities, supportedHosts, {
        sets: supportedSets,
        methods: supportedMethods
    }, supportedControlTarget);
    requirementsChecker.setFilteringEnabled(true);
    contextActivationManager.setRequirementsChecker(requirementsChecker);
};
if (typeof (Sys) !== 'undefined' && Sys && Sys.Application) {
    Sys.Application.notifyScriptLoaded();
}
if (typeof (NotifyScriptLoadedAndExecuteWaitingJobs) === 'function') {
    NotifyScriptLoadedAndExecuteWaitingJobs("OSFRuntimeAccessWAC.js");
}
