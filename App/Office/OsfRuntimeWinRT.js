/* Office runtime JavaScript library */
/* Version: 16.0.6807.3002 */
/*
	Copyright (c) Microsoft Corporation.  All rights reserved.
*/

/*
	Your use of this file is governed by the Microsoft Services Agreement http://go.microsoft.com/fwlink/?LinkId=266419.
*/

var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
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
    dispidMessageParentMethod: 144,
    dispidMethodMax: 144,
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
    dispidGetMaxResourceIndexMethod: 123
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
OSF.HostType = {
    Excel: "Excel",
    Outlook: "Outlook",
    Access: "Access",
    PowerPoint: "PowerPoint"
};
OSF.HostPlatform = {
    Web: "Web",
    WinRT: "WinRT"
};
OSF.OsfControlTarget = {
    InContent: 0,
    TaskPane: 1,
    Contextual: 2
};
OSF.OfficeAppType = {
    ContentApp: OSF.OsfControlTarget.InContent,
    TaskPaneApp: OSF.OsfControlTarget.TaskPane,
    MailApp: OSF.OsfControlTarget.Contextual
};
OSF.Constants = {
    Http: "http",
    Https: "https",
    ProtocolSeparator: "://",
    ETokenParameterName: "et"
};
var OsfRuntimeBase;
(function (OsfRuntimeBase) {
    (function (OsfControlType) {
        OsfControlType[OsfControlType["DocumentLevel"] = 0] = "DocumentLevel";
        OsfControlType[OsfControlType["ContainerLevel"] = 1] = "ContainerLevel";
    })(OsfRuntimeBase.OsfControlType || (OsfRuntimeBase.OsfControlType = {}));
    var OsfControlType = OsfRuntimeBase.OsfControlType;
    (function (OsfControlPermission) {
        OsfControlPermission[OsfControlPermission["Restricted"] = 1] = "Restricted";
        OsfControlPermission[OsfControlPermission["ReadDocument"] = 2] = "ReadDocument";
        OsfControlPermission[OsfControlPermission["WriteDocument"] = 4] = "WriteDocument";
        OsfControlPermission[OsfControlPermission["ReadWriteDocument"] = 6] = "ReadWriteDocument";
        OsfControlPermission[OsfControlPermission["ReadItem"] = 32] = "ReadItem";
        OsfControlPermission[OsfControlPermission["ReadWriteMailbox"] = 64] = "ReadWriteMailbox";
        OsfControlPermission[OsfControlPermission["ReadAllDocument"] = 131] = "ReadAllDocument";
    })(OsfRuntimeBase.OsfControlPermission || (OsfRuntimeBase.OsfControlPermission = {}));
    var OsfControlPermission = OsfRuntimeBase.OsfControlPermission;
    ;
    (function (OsfControlStatus) {
        OsfControlStatus[OsfControlStatus["NotActivated"] = 1] = "NotActivated";
        OsfControlStatus[OsfControlStatus["Activated"] = 2] = "Activated";
        OsfControlStatus[OsfControlStatus["AppStoreNotReachable"] = 3] = "AppStoreNotReachable";
        OsfControlStatus[OsfControlStatus["InvalidOsfControl"] = 4] = "InvalidOsfControl";
        OsfControlStatus[OsfControlStatus["UnsupportedStore"] = 5] = "UnsupportedStore";
        OsfControlStatus[OsfControlStatus["UnknownStore"] = 6] = "UnknownStore";
        OsfControlStatus[OsfControlStatus["ActivationFailed"] = 7] = "ActivationFailed";
        OsfControlStatus[OsfControlStatus["NotSandBoxSupported"] = 8] = "NotSandBoxSupported";
    })(OsfRuntimeBase.OsfControlStatus || (OsfRuntimeBase.OsfControlStatus = {}));
    var OsfControlStatus = OsfRuntimeBase.OsfControlStatus;
    ;
    var ContextActivationParametersBase = (function () {
        function ContextActivationParametersBase(appName, appVersion, clientMode, appUiLocale, dataLocale) {
            this.appName = appName;
            this.appVersion = appVersion;
            this.clientMode = clientMode;
            this.appUILocale = appUiLocale;
            this.dataLocale = dataLocale;
        }
        return ContextActivationParametersBase;
    })();
    OsfRuntimeBase.ContextActivationParametersBase = ContextActivationParametersBase;
    var InsertOsfControlParametersBase = (function () {
        function InsertOsfControlParametersBase(parent, id, osfControlType) {
            this.parent = parent;
            this.id = id;
            this.osfControlType = osfControlType;
        }
        return InsertOsfControlParametersBase;
    })();
    OsfRuntimeBase.InsertOsfControlParametersBase = InsertOsfControlParametersBase;
    var ContextActivationManagerBase = (function () {
        function ContextActivationManagerBase(contextActivationParametersBase) {
            this.appName = contextActivationParametersBase.appName;
            this.appVersion = contextActivationParametersBase.appVersion;
            this.appUILocale = contextActivationParametersBase.appUILocale;
            this.dataLocale = contextActivationParametersBase.dataLocale;
            this.clientMode = contextActivationParametersBase.clientMode;
        }
        ContextActivationManagerBase.prototype.insertOsfControl = function (insertOsfControlParametersBase) {
            var osfControl = new OsfControlBase(insertOsfControlParametersBase);
            osfControl.activate(null);
        };
        return ContextActivationManagerBase;
    })();
    OsfRuntimeBase.ContextActivationManagerBase = ContextActivationManagerBase;
    var OsfControlBase = (function () {
        function OsfControlBase(insertOsfControlParametersBase) {
            this.activationStatus = OsfControlStatus.NotActivated;
            this.osfControlParentElement = insertOsfControlParametersBase.parent;
            this.id = insertOsfControlParametersBase.id;
            this.osfControlType = insertOsfControlParametersBase.osfControlType;
        }
        OsfControlBase.prototype.activate = function (refreshContext) {
            this.activationStatus = OsfControlStatus.Activated;
        };
        return OsfControlBase;
    })();
    OsfRuntimeBase.OsfControlBase = OsfControlBase;
})(OsfRuntimeBase || (OsfRuntimeBase = {}));
var OsfRuntimeWinRT;
(function (OsfRuntimeWinRT) {
    var ContextActivationParametersWinRT = (function (_super) {
        __extends(ContextActivationParametersWinRT, _super);
        function ContextActivationParametersWinRT(appName, appVersion, clientMode, appUiLocale, dataLocale, hostControl) {
            _super.call(this, appName, appVersion, clientMode, appUiLocale, dataLocale);
            this.hostControl = hostControl;
        }
        return ContextActivationParametersWinRT;
    })(OsfRuntimeBase.ContextActivationParametersBase);
    OsfRuntimeWinRT.ContextActivationParametersWinRT = ContextActivationParametersWinRT;
    var ContextActivationManagerWinRT = (function (_super) {
        __extends(ContextActivationManagerWinRT, _super);
        function ContextActivationManagerWinRT(contextActivationParameters) {
            _super.call(this, contextActivationParameters);
            this.clientId = "1";
            this.serviceEndPoint = Microsoft.Office.Common.XdmCommunicationManager.createServiceEndPoint(this.clientId);
            var me = this;
            var getAppContextAsync = function OSF_ContextActivationManager$getAppContextAsync(contextId, gotAppContext) {
                var e = Function._validateParams(arguments, [{ name: "contextId", type: String, mayBeNull: false },
                    { name: "gotAppContext", type: Function, mayBeNull: false }
                ]);
                if (e)
                    throw e;
                if (!me.osfControl) {
                    OsfMsAjaxFactory.msAjaxDebug.trace("osfControl for the given ID doesn't exist.");
                    throw OsfMsAjaxFactory.msAjaxError.argument("contextId");
                }
                else {
                    var eToken = me.osfControl.getEToken();
                    var correlationId;
                    var appInstanceId;
                    var touchEnabled;
                    var commerceAllowed;
                    var minorVersion;
                    var requirementMatrix;
                    var appContext = new OSF.OfficeAppContext(contextId, me.appName, me.appVersion, me.appUILocale, me.dataLocale, window.location.href, me.clientMode, {}, Microsoft.Office.WebExtension.InitializationReason.Inserted, me.osfControl.osfControlType, me.osfControl.getEToken(), correlationId, appInstanceId, touchEnabled, commerceAllowed, minorVersion, requirementMatrix);
                    gotAppContext(appContext);
                }
            };
            this.serviceEndPoint.registerMethod("ContextActivationManager_getAppContextAsync", getAppContextAsync, Microsoft.Office.Common.InvokeType.async, false);
            this.hostControl = contextActivationParameters.hostControl;
            var winRTFacade = new WinRTFacade(this.hostControl);
            winRTFacade.serviceEndpoint = this.serviceEndPoint;
            var registeredActions = {
                'ContextActivationManager_getAppContextAsync': OsfRuntimeBase.OsfControlPermission.Restricted,
                'executeMethod': {}
            };
            registeredActions.executeMethod[OSF.DDA.MethodDispId.dispidGetSelectedDataMethod] = OsfRuntimeBase.OsfControlPermission.Restricted;
            var policyManager = new WinRTPolicyManager(this, registeredActions);
            this.serviceEndPoint.setPolicyManager(policyManager);
            this.serviceEndPoint.registerMethod("executeMethod", winRTFacade.executeMethod, Microsoft.Office.Common.InvokeType.async, false);
        }
        ContextActivationManagerWinRT.prototype.insertOsfControl = function (insertOsfControlParameters) {
            this.osfControl = new OsfControlWinRT(insertOsfControlParameters, this);
            this.osfControl.activate(null);
        };
        return ContextActivationManagerWinRT;
    })(OsfRuntimeBase.ContextActivationManagerBase);
    OsfRuntimeWinRT.ContextActivationManagerWinRT = ContextActivationManagerWinRT;
    var InsertOsfControlParametersWinRT = (function (_super) {
        __extends(InsertOsfControlParametersWinRT, _super);
        function InsertOsfControlParametersWinRT(parent, id, osfControlType, url) {
            _super.call(this, parent, id, osfControlType);
            this.osfControlContainer = this.createDefaultContainer(id);
            this.url = url;
        }
        InsertOsfControlParametersWinRT.prototype.createDefaultContainer = function (id) {
            var frame = document.createElement("iframe");
            frame.setAttribute("id", id);
            frame.setAttribute("width", "100%");
            frame.setAttribute("height", "100%");
            frame.setAttribute("frameborder", "0");
            return frame;
        };
        return InsertOsfControlParametersWinRT;
    })(OsfRuntimeBase.InsertOsfControlParametersBase);
    OsfRuntimeWinRT.InsertOsfControlParametersWinRT = InsertOsfControlParametersWinRT;
    var OsfControlWinRT = (function (_super) {
        __extends(OsfControlWinRT, _super);
        function OsfControlWinRT(insertOsfControlParameters, contextActivationMgr) {
            _super.call(this, insertOsfControlParameters);
            this._osfControlContainer = insertOsfControlParameters.osfControlContainer;
            this._osfControlContainerParent = insertOsfControlParameters.parent;
            this._id = insertOsfControlParameters.id;
            this._url = insertOsfControlParameters.url;
            this._osfControlType = insertOsfControlParameters.osfControlType;
            this._contextActivationMgr = contextActivationMgr;
            this._etoken = "";
        }
        OsfControlWinRT.prototype.activate = function (refreshContext) {
            this._activate(this._osfControlContainer, this._url);
            _super.prototype.activate.call(this, refreshContext);
        };
        OsfControlWinRT.prototype._addETokenAsQueryParameter = function (iframeUrl) {
            var aElement = document.createElement('a');
            aElement.href = iframeUrl;
            var etoken = this.getEToken();
            var etokenQueryString = OSF.Constants.ETokenParameterName + "=" + encodeURIComponent(OSF.OUtil.encodeBase64(etoken));
            var queryString = aElement.search.length > 1 ? aElement.search.substr(1) + "&" : "";
            aElement.search = queryString + etokenQueryString;
            var modifiedUrl = aElement.href;
            aElement = null;
            return modifiedUrl;
        };
        OsfControlWinRT.prototype._addHostInfoAsQueryParam = function (url, hostInfoValue) {
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
        OsfControlWinRT.prototype._activate = function (frame, iframeUrl) {
            iframeUrl = this._addETokenAsQueryParameter(iframeUrl);
            var cacheKey = this._contextActivationMgr.clientId + "_" + window.location.href + "_" + this.id;
            this._conversationId = OSF.OUtil.getFrameNameAndConversationId(cacheKey, frame);
            var newUrl = this._addHostInfoAsQueryParam(iframeUrl, this._contextActivationMgr.hostType + "|" + this._contextActivationMgr.hostPlatform + "|" + this._contextActivationMgr.hostSpecificFileVersion);
            newUrl = OSF.OUtil.addXdmInfoAsHash(newUrl, this._conversationId + "|" + this.id + "|" + window.location.href);
            this._contextActivationMgr.serviceEndPoint.registerConversation(this._conversationId);
            frame.setAttribute("src", newUrl);
            this._osfControlContainerParent.appendChild(frame);
        };
        OsfControlWinRT.prototype.getEToken = function () {
            return this._etoken;
        };
        OsfControlWinRT.prototype.getOsfControlContainer = function () {
            return this._osfControlContainer;
        };
        OsfControlWinRT.prototype.getOsfControlUrl = function () {
            return this._url;
        };
        OsfControlWinRT.prototype.getConversationId = function () {
            return this._conversationId;
        };
        OsfControlWinRT.prototype.getReasons = function () {
            return this._reasons;
        };
        OsfControlWinRT.prototype.getSettigns = function () {
            return this._settings;
        };
        OsfControlWinRT.prototype.setAgaveEndPoint = function (agaveEndPoint) {
            this._agaveEndPoint = agaveEndPoint;
        };
        return OsfControlWinRT;
    })(OsfRuntimeBase.OsfControlBase);
    OsfRuntimeWinRT.OsfControlWinRT = OsfControlWinRT;
    var WinRTFacade = (function () {
        function WinRTFacade(_hostControl) {
            this.executeMethod = function (params, callback) {
                _hostControl.InvokeDdaMethod(params, function onMethodExecution(asyncResult) {
                    callback(asyncResult);
                }, null);
            };
            this.registerEvent = function (eventHandler, callback, params) {
                _hostControl.RegisterDdaEventAsync(params['eventDispId'], params['controlId'], params['targetId'], eventHandler, callback);
            };
            this.unregisterEvent = function (eventHandler, callback, params) {
                _hostControl.UnregisterDdaEventAsync(params['eventDispId'], params['controlId'], params['targetId'], eventHandler, callback);
            };
        }
        return WinRTFacade;
    })();
    var WinRTPolicyManager = (function () {
        function WinRTPolicyManager(contextActivationManager, registeredActions) {
            var e = Function._validateParams(arguments, [
                {
                    name: "contextActivationMgr",
                    type: Object,
                    mayBeNull: false
                },
                {
                    name: "registeredActions",
                    type: Object,
                    mayBeNull: false
                }
            ]);
            if (e) {
                throw e;
            }
            this._contextActivationManager = contextActivationManager;
            this._registeredActions = {};
            var isValidPermission;
            var registeredPermission;
            for (var actionName in registeredActions) {
                isValidPermission = true;
                registeredPermission = registeredActions[actionName];
                if (typeof registeredPermission === "object") {
                    for (var dispId in registeredPermission) {
                        isValidPermission = this.validatePermissionInput(registeredPermission[dispId]);
                        if (!isValidPermission) {
                            break;
                        }
                    }
                }
                else {
                    isValidPermission = this.validatePermissionInput(registeredPermission);
                }
                if (!isValidPermission) {
                    throw OsfMsAjaxFactory.msAjaxError.argument("registeredActions");
                }
                this._registeredActions[actionName] = registeredPermission;
            }
        }
        WinRTPolicyManager.prototype.checkPermission = function (conversationID, actionName, params) {
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
            if (permissionNeeded === OsfRuntimeBase.OsfControlPermission.Restricted) {
                return true;
            }
            var entry = this._contextActivationManager._getManifestAndTargetByConversationId(conversationID);
            if (entry && entry.manifest) {
                return entry.manifest.hasPermission(permissionNeeded);
            }
            else {
                return false;
            }
        };
        WinRTPolicyManager.prototype.validatePermissionInput = function (permission) {
            var isValidPermission = false;
            for (var permissionName in OsfRuntimeBase.OsfControlPermission) {
                if (OsfRuntimeBase.OsfControlPermission[permissionName] === permission) {
                    isValidPermission = true;
                    break;
                }
            }
            return isValidPermission;
        };
        return WinRTPolicyManager;
    })();
})(OsfRuntimeWinRT || (OsfRuntimeWinRT = {}));
var cblist;
var debugConsole;
var WinRTHost;
(function (WinRTHost) {
    var HostControl = (function () {
        function HostControl(consoleParent) {
            this._debugConsole = null;
            if (consoleParent != null) {
                this._debugConsole = document.createElement("textArea");
                this._debugConsole.setAttribute("id", "debugConsole");
                this._debugConsole.setAttribute("cols", "100");
                this._debugConsole.setAttribute("rows", "10");
                consoleParent.appendChild(this._debugConsole);
                debugConsole = this._debugConsole;
                this._logLineCount = 1;
            }
        }
        HostControl.prototype.InvokeDdaMethod = function (parameters, callback, additionalParams) {
            var dispatchId = parameters["DdaMethod"]["DispatchId"];
            this.writeToLog("Method invoked! DispatchId=" + dispatchId);
            cblist = this.genericCallback;
            try {
                var methodObj = new MethodStruct(0, parameters["DdaMethod"]["DispatchId"], parameters["DdaGetBindingData"], 'cblist');
                this.sendToHostApp(methodObj);
            }
            catch (ex) {
                this.writeToLog("WebView:ScriptNotify error:" + ex.Message);
            }
        };
        HostControl.prototype.RegisterDdaEventAsync = function (eventDispId, controlId, targetId, onEvent, onRegistration) {
            var evnId = eventDispId;
            this.writeToLog("RegisterEvent invoked! eventDispId=" + eventDispId);
        };
        HostControl.prototype.UnregisterDdaEventAsync = function (eventDispId, controlId, targetId, onEvent, onUnregistration) {
            var evnId = eventDispId;
            this.writeToLog("UnRegisterEvent invoked! eventDispId=" + eventDispId);
        };
        HostControl.prototype.writeToLog = function (value) {
            if (this._debugConsole != null && value != null && value.toString() != '') {
                this._debugConsole.value = this._logLineCount + ": " + value + "\r\n" + this._debugConsole.value;
                this._logLineCount = this._logLineCount + 1;
            }
        };
        HostControl.prototype.genericCallback = function () {
            for (var i = 0; i < arguments.length; i++) {
                var value = arguments[i].toString();
                if (debugConsole != null && value != null && value != '') {
                    debugConsole.value = value + "\r\n" + debugConsole.value;
                }
            }
        };
        HostControl.prototype.sendToHostApp = function (obj) {
            var value = JSON.stringify(obj);
            this.writeToLog("JSON string to host app=" + value);
            window.external.notify(value);
        };
        return HostControl;
    })();
    WinRTHost.HostControl = HostControl;
    var MethodStruct = (function () {
        function MethodStruct(action, dispId, parameters, callback) {
            this.Action = action;
            this.DispId = dispId;
            this.Parameters = [];
            this.Parameters[0] = '"' + JSON.stringify(parameters) + '"';
            this.Callback = callback;
        }
        return MethodStruct;
    })();
})(WinRTHost || (WinRTHost = {}));
