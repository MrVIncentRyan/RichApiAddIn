/* Office iOS client gallery insertion dialog JavaScript file */
/* Version: 16.0.7504.3000 */
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
                typeof (Function._validateParams) === "function" &&
                Sys.Serialization && Sys.Serialization.JavaScriptSerializer && typeof (Sys.Serialization.JavaScriptSerializer.serialize) === "function") {
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
OSF.XdmFieldName = {
    ConversationUrl: "ConversationUrl",
    AppId: "AppId"
};
OSF.WindowNameItemKeys = {
    BaseFrameName: "baseFrameName",
    HostInfo: "hostInfo",
    XdmInfo: "xdmInfo",
    SerializerVersion: "serializerVersion",
    AppContext: "appContext"
};
OSF.OUtil = (function () {
    var _uniqueId = -1;
    var _xdmInfoKey = '&_xdm_Info=';
    var _serializerVersionKey = '&_serializer_version=';
    var _xdmSessionKeyPrefix = '_xdm_';
    var _serializerVersionKeyPrefix = '_serializer_version=';
    var _fragmentSeparator = '#';
    var _fragmentInfoDelimiter = '&';
    var _classN = "class";
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
    function _reOrderTabbableElements(elements) {
        var bucket0 = [];
        var bucketPositive = [];
        var i;
        var len = elements.length;
        var ele;
        for (i = 0; i < len; i++) {
            ele = elements[i];
            if (ele.tabIndex) {
                if (ele.tabIndex > 0) {
                    bucketPositive.push(ele);
                }
                else if (ele.tabIndex === 0) {
                    bucket0.push(ele);
                }
            }
            else {
                bucket0.push(ele);
            }
        }
        bucketPositive = bucketPositive.sort(function (left, right) {
            var diff = left.tabIndex - right.tabIndex;
            if (diff === 0) {
                diff = bucketPositive.indexOf(left) - bucketPositive.indexOf(right);
            }
            return diff;
        });
        return [].concat(bucketPositive, bucket0);
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
        getFrameName: function OSF_OUtil$getFrameName(cacheKey) {
            return _xdmSessionKeyPrefix + cacheKey + this.generateConversationId();
        },
        addXdmInfoAsHash: function OSF_OUtil$addXdmInfoAsHash(url, xdmInfoValue) {
            return OSF.OUtil.addInfoAsHash(url, _xdmInfoKey, xdmInfoValue, false);
        },
        addSerializerVersionAsHash: function OSF_OUtil$addSerializerVersionAsHash(url, serializerVersion) {
            return OSF.OUtil.addInfoAsHash(url, _serializerVersionKey, serializerVersion, true);
        },
        addInfoAsHash: function OSF_OUtil$addInfoAsHash(url, keyName, infoValue, encodeInfo) {
            url = url.trim() || '';
            var urlParts = url.split(_fragmentSeparator);
            var urlWithoutFragment = urlParts.shift();
            var fragment = urlParts.join(_fragmentSeparator);
            var newFragment;
            if (encodeInfo) {
                newFragment = [keyName, encodeURIComponent(infoValue), fragment].join('');
            }
            else {
                newFragment = [fragment, keyName, infoValue].join('');
            }
            return [urlWithoutFragment, _fragmentSeparator, newFragment].join('');
        },
        parseHostInfoFromWindowName: function OSF_OUtil$parseHostInfoFromWindowName(skipSessionStorage, windowName) {
            return OSF.OUtil.parseInfoFromWindowName(skipSessionStorage, windowName, OSF.WindowNameItemKeys.HostInfo);
        },
        parseXdmInfo: function OSF_OUtil$parseXdmInfo(skipSessionStorage) {
            var xdmInfoValue = OSF.OUtil.parseXdmInfoWithGivenFragment(skipSessionStorage, window.location.hash);
            if (!xdmInfoValue) {
                xdmInfoValue = OSF.OUtil.parseXdmInfoFromWindowName(skipSessionStorage, window.name);
            }
            return xdmInfoValue;
        },
        parseXdmInfoFromWindowName: function OSF_OUtil$parseXdmInfoFromWindowName(skipSessionStorage, windowName) {
            return OSF.OUtil.parseInfoFromWindowName(skipSessionStorage, windowName, OSF.WindowNameItemKeys.XdmInfo);
        },
        parseXdmInfoWithGivenFragment: function OSF_OUtil$parseXdmInfoWithGivenFragment(skipSessionStorage, fragment) {
            return OSF.OUtil.parseInfoWithGivenFragment(_xdmInfoKey, _xdmSessionKeyPrefix, false, skipSessionStorage, fragment);
        },
        parseSerializerVersion: function OSF_OUtil$parseSerializerVersion(skipSessionStorage) {
            var serializerVersion = OSF.OUtil.parseSerializerVersionWithGivenFragment(skipSessionStorage, window.location.hash);
            if (isNaN(serializerVersion)) {
                serializerVersion = OSF.OUtil.parseSerializerVersionFromWindowName(skipSessionStorage, window.name);
            }
            return serializerVersion;
        },
        parseSerializerVersionFromWindowName: function OSF_OUtil$parseSerializerVersionFromWindowName(skipSessionStorage, windowName) {
            return parseInt(OSF.OUtil.parseInfoFromWindowName(skipSessionStorage, windowName, OSF.WindowNameItemKeys.SerializerVersion));
        },
        parseSerializerVersionWithGivenFragment: function OSF_OUtil$parseSerializerVersionWithGivenFragment(skipSessionStorage, fragment) {
            return parseInt(OSF.OUtil.parseInfoWithGivenFragment(_serializerVersionKey, _serializerVersionKeyPrefix, true, skipSessionStorage, fragment));
        },
        parseInfoFromWindowName: function OSF_OUtil$parseInfoFromWindowName(skipSessionStorage, windowName, infoKey) {
            try {
                var windowNameObj = JSON.parse(windowName);
                var infoValue = windowNameObj != null ? windowNameObj[infoKey] : null;
                var osfSessionStorage = _getSessionStorage();
                if (!skipSessionStorage && osfSessionStorage && windowNameObj != null) {
                    var sessionKey = windowNameObj[OSF.WindowNameItemKeys.BaseFrameName] + infoKey;
                    if (infoValue) {
                        osfSessionStorage.setItem(sessionKey, infoValue);
                    }
                    else {
                        infoValue = osfSessionStorage.getItem(sessionKey);
                    }
                }
                return infoValue;
            }
            catch (Exception) {
                return null;
            }
        },
        parseInfoWithGivenFragment: function OSF_OUtil$parseInfoWithGivenFragment(infoKey, infoKeyPrefix, decodeInfo, skipSessionStorage, fragment) {
            var fragmentParts = fragment.split(infoKey);
            var infoValue = fragmentParts.length > 1 ? fragmentParts[fragmentParts.length - 1] : null;
            if (decodeInfo && infoValue != null) {
                if (infoValue.indexOf(_fragmentInfoDelimiter) >= 0) {
                    infoValue = infoValue.split(_fragmentInfoDelimiter)[0];
                }
                infoValue = decodeURIComponent(infoValue);
            }
            var osfSessionStorage = _getSessionStorage();
            if (!skipSessionStorage && osfSessionStorage) {
                var sessionKeyStart = window.name.indexOf(infoKeyPrefix);
                if (sessionKeyStart > -1) {
                    var sessionKeyEnd = window.name.indexOf(";", sessionKeyStart);
                    if (sessionKeyEnd == -1) {
                        sessionKeyEnd = window.name.length;
                    }
                    var sessionKey = window.name.substring(sessionKeyStart, sessionKeyEnd);
                    if (infoValue) {
                        osfSessionStorage.setItem(sessionKey, infoValue);
                    }
                    else {
                        infoValue = osfSessionStorage.getItem(sessionKey);
                    }
                }
            }
            return infoValue;
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
            if (typeof items[1] == "undefined") {
                items = strInfo.split("%7C");
            }
            return items;
        },
        getXdmFieldValue: function OSF_OUtil$getXdmFieldValue(xdmFieldName, skipSessionStorage) {
            var fieldValue = '';
            var xdmInfoValue = OSF.OUtil.parseXdmInfo(skipSessionStorage);
            if (xdmInfoValue) {
                var items = OSF.OUtil.getInfoItems(xdmInfoValue);
                if (items != undefined && items.length >= 3) {
                    switch (xdmFieldName) {
                        case OSF.XdmFieldName.ConversationUrl:
                            fieldValue = items[2];
                            break;
                        case OSF.XdmFieldName.AppId:
                            fieldValue = items[1];
                            break;
                    }
                }
            }
            return fieldValue;
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
        isChrome: function OSF_Outil$isChrome() {
            return (window.navigator.userAgent.indexOf("Chrome") > 0) && !OSF.OUtil.isEdge();
        },
        isEdge: function OSF_Outil$isEdge() {
            return window.navigator.userAgent.indexOf("Edge") > 0;
        },
        isIE: function OSF_Outil$isIE() {
            return window.navigator.userAgent.indexOf("Trident") > 0;
        },
        isFirefox: function OSF_Outil$isFirefox() {
            return window.navigator.userAgent.indexOf("Firefox") > 0;
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
        },
        addClass: function OSF_OUtil$addClass(elmt, val) {
            if (!OSF.OUtil.hasClass(elmt, val)) {
                var className = elmt.getAttribute(_classN);
                if (className) {
                    elmt.setAttribute(_classN, className + " " + val);
                }
                else {
                    elmt.setAttribute(_classN, val);
                }
            }
        },
        hasClass: function OSF_OUtil$hasClass(elmt, clsName) {
            var className = elmt.getAttribute(_classN);
            return className && className.match(new RegExp('(\\s|^)' + clsName + '(\\s|$)'));
        },
        focusToFirstTabbable: function OSF_OUtil$focusToFirstTabbable(all, backward) {
            var next;
            var focused = false;
            var candidate;
            var setFlag = function (e) {
                focused = true;
            };
            var findNextPos = function (allLen, currPos, backward) {
                if (currPos < 0 || currPos > allLen) {
                    return -1;
                }
                else if (currPos === 0 && backward) {
                    return -1;
                }
                else if (currPos === allLen - 1 && !backward) {
                    return -1;
                }
                if (backward) {
                    return currPos - 1;
                }
                else {
                    return currPos + 1;
                }
            };
            all = _reOrderTabbableElements(all);
            next = backward ? all.length - 1 : 0;
            if (all.length === 0) {
                return null;
            }
            while (!focused && next >= 0 && next < all.length) {
                candidate = all[next];
                window.focus();
                candidate.addEventListener('focus', setFlag);
                candidate.focus();
                candidate.removeEventListener('focus', setFlag);
                next = findNextPos(all.length, next, backward);
                if (!focused && candidate === document.activeElement) {
                    focused = true;
                }
            }
            if (focused) {
                return candidate;
            }
            else {
                return null;
            }
        },
        focusToNextTabbable: function OSF_OUtil$focusToNextTabbable(all, curr, shift) {
            var currPos;
            var next;
            var focused = false;
            var candidate;
            var setFlag = function (e) {
                focused = true;
            };
            var findCurrPos = function (all, curr) {
                var i = 0;
                for (; i < all.length; i++) {
                    if (all[i] === curr) {
                        return i;
                    }
                }
                return -1;
            };
            var findNextPos = function (allLen, currPos, shift) {
                if (currPos < 0 || currPos > allLen) {
                    return -1;
                }
                else if (currPos === 0 && shift) {
                    return -1;
                }
                else if (currPos === allLen - 1 && !shift) {
                    return -1;
                }
                if (shift) {
                    return currPos - 1;
                }
                else {
                    return currPos + 1;
                }
            };
            all = _reOrderTabbableElements(all);
            currPos = findCurrPos(all, curr);
            next = findNextPos(all.length, currPos, shift);
            if (next < 0) {
                return null;
            }
            while (!focused && next >= 0 && next < all.length) {
                candidate = all[next];
                candidate.addEventListener('focus', setFlag);
                candidate.focus();
                candidate.removeEventListener('focus', setFlag);
                next = findNextPos(all.length, next, shift);
                if (!focused && candidate === document.activeElement) {
                    focused = true;
                }
            }
            if (focused) {
                return candidate;
            }
            else {
                return null;
            }
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
            if (!instanceType || (typeof (instanceType) !== "function") || !instanceType.__typeName || instanceType.__typeName === 'Object') {
                instanceType = Object;
            }
            return !!(instanceType === type) ||
                (instanceType.__typeName && type.__typeName && instanceType.__typeName === type.__typeName);
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
        var registerTypeInternal = function registerTypeInternal(type, name, isClass) {
            if (type.__typeName === undefined) {
                type.__typeName = name;
            }
            if (type.__class === undefined) {
                type.__class = isClass;
            }
        };
        registerTypeInternal(Function, "Function", true);
        registerTypeInternal(Error, "Error", true);
        registerTypeInternal(Object, "Object", true);
        registerTypeInternal(String, "String", true);
        registerTypeInternal(Boolean, "Boolean", true);
        registerTypeInternal(Date, "Date", true);
        registerTypeInternal(Number, "Number", true);
        registerTypeInternal(RegExp, "RegExp", true);
        registerTypeInternal(Array, "Array", true);
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
var OfficeExt;
(function (OfficeExt) {
    function appSpecificCheckOriginFunction(url, eventObj, messageObj, checkOriginFunction) {
        return true;
    }
    ;
    OfficeExt.appSpecificCheckOrigin = appSpecificCheckOriginFunction;
})(OfficeExt || (OfficeExt = {}));
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
    },
    dispose: function Microsoft_Office_Common_ServiceEndPoint$dispose() {
        this._methodObjectList = null;
        this._eventHandlerProxyList = null;
        this._Id = null;
        this._conversations = null;
        this._policyManager = null;
        this._appDomains = null;
        this._onHandleRequestError = null;
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
    try {
        if (!targetWindow.postMessage) {
            throw OsfMsAjaxFactory.msAjaxError.argument("targetWindow");
        }
    }
    catch (ex) {
        if (!Object.prototype.hasOwnProperty.call(targetWindow, "postMessage")) {
            throw OsfMsAjaxFactory.msAjaxError.argument("targetWindow");
        }
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
                if (!clientEndPoint) {
                    return;
                }
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
        deleteServiceEndPoint: function Microsoft_Office_Common_XdmCommunicationManager$deleteServiceEndPoint(serviceEndPointId) {
            var e = Function._validateParams(arguments, [
                { name: "serviceEndPointId", type: String, mayBeNull: false }
            ]);
            if (e)
                throw e;
            delete _serviceEndPoints[serviceEndPointId];
        },
        checkUrlWithAppDomains: function Microsoft_Office_Common_XdmCommunicationManager$_checkUrlWithAppDomains(appDomains, origin) {
            return _checkOriginWithAppDomains(appDomains, origin);
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
    this._send = function (result, responseCode) {
        if (responseCode != undefined) {
            me._invokeResultCode = responseCode;
        }
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
var OSF = OSF || {};
var OSFWebkit;
(function (OSFWebkit) {
    var WebkitSafeArray = (function () {
        function WebkitSafeArray(data) {
            this.data = data;
            this.safeArrayFlag = this.isSafeArray(data);
        }
        WebkitSafeArray.prototype.dimensions = function () {
            var dimensions = 0;
            if (this.safeArrayFlag) {
                dimensions = this.data[0][0];
            }
            else if (this.isArray()) {
                dimensions = 2;
            }
            return dimensions;
        };
        WebkitSafeArray.prototype.getItem = function () {
            var array = [];
            var element = null;
            if (this.safeArrayFlag) {
                array = this.toArray();
            }
            else {
                array = this.data;
            }
            element = array;
            for (var i = 0; i < arguments.length; i++) {
                element = element[arguments[i]];
            }
            return element;
        };
        WebkitSafeArray.prototype.lbound = function (dimension) {
            return 0;
        };
        WebkitSafeArray.prototype.ubound = function (dimension) {
            var ubound = 0;
            if (this.safeArrayFlag) {
                ubound = this.data[0][dimension];
            }
            else if (this.isArray()) {
                if (dimension == 1) {
                    return this.data.length;
                }
                else if (dimension == 2) {
                    if (OSF.OUtil.isArray(this.data[0])) {
                        return this.data[0].length;
                    }
                    else if (this.data[0] != null) {
                        return 1;
                    }
                }
            }
            return ubound;
        };
        WebkitSafeArray.prototype.toArray = function () {
            if (this.isArray() == false) {
                return this.data;
            }
            var arr = [];
            var startingIndex = this.safeArrayFlag ? 1 : 0;
            for (var i = startingIndex; i < this.data.length; i++) {
                var element = this.data[i];
                if (this.isSafeArray(element)) {
                    arr.push(new WebkitSafeArray(element));
                }
                else {
                    arr.push(element);
                }
            }
            return arr;
        };
        WebkitSafeArray.prototype.isArray = function () {
            return OSF.OUtil.isArray(this.data);
        };
        WebkitSafeArray.prototype.isSafeArray = function (obj) {
            var isSafeArray = false;
            if (OSF.OUtil.isArray(obj) && OSF.OUtil.isArray(obj[0])) {
                var bounds = obj[0];
                var dimensions = bounds[0];
                if (bounds.length != dimensions + 1) {
                    return false;
                }
                var expectedArraySize = 1;
                for (var i = 1; i < bounds.length; i++) {
                    var dimension = bounds[i];
                    if (isFinite(dimension) == false) {
                        return false;
                    }
                    expectedArraySize = expectedArraySize * dimension;
                }
                expectedArraySize++;
                isSafeArray = (expectedArraySize == obj.length);
            }
            return isSafeArray;
        };
        return WebkitSafeArray;
    })();
    OSFWebkit.WebkitSafeArray = WebkitSafeArray;
})(OSFWebkit || (OSFWebkit = {}));
var OSFWebkit;
(function (OSFWebkit) {
    var ScriptMessaging;
    (function (ScriptMessaging) {
        var scriptMessenger = null;
        function agaveHostCallback(callbackId, params) {
            scriptMessenger.agaveHostCallback(callbackId, params);
        }
        ScriptMessaging.agaveHostCallback = agaveHostCallback;
        function agaveHostEventCallback(callbackId, params) {
            scriptMessenger.agaveHostEventCallback(callbackId, params);
        }
        ScriptMessaging.agaveHostEventCallback = agaveHostEventCallback;
        function GetScriptMessenger() {
            if (scriptMessenger == null) {
                scriptMessenger = new WebkitScriptMessaging("OSF.ScriptMessaging.agaveHostCallback", "OSF.ScriptMessaging.agaveHostEventCallback");
            }
            return scriptMessenger;
        }
        ScriptMessaging.GetScriptMessenger = GetScriptMessenger;
        var EventHandlerCallback = (function () {
            function EventHandlerCallback(id, targetId, handler) {
                this.id = id;
                this.targetId = targetId;
                this.handler = handler;
            }
            return EventHandlerCallback;
        })();
        var WebkitScriptMessaging = (function () {
            function WebkitScriptMessaging(methodCallbackName, eventCallbackName) {
                this.callingIndex = 0;
                this.callbackList = {};
                this.eventHandlerList = {};
                this.asyncMethodCallbackFunctionName = methodCallbackName;
                this.eventCallbackFunctionName = eventCallbackName;
                this.conversationId = WebkitScriptMessaging.getCurrentTimeMS().toString();
            }
            WebkitScriptMessaging.prototype.invokeMethod = function (handlerName, methodId, params, callback) {
                var messagingArgs = {};
                this.postWebkitMessage(messagingArgs, handlerName, methodId, params, callback);
            };
            WebkitScriptMessaging.prototype.registerEvent = function (handlerName, methodId, dispId, targetId, handler, callback) {
                var messagingArgs = {
                    eventCallbackFunction: this.eventCallbackFunctionName
                };
                var hostArgs = {
                    id: dispId,
                    targetId: targetId
                };
                var correlationId = this.postWebkitMessage(messagingArgs, handlerName, methodId, hostArgs, callback);
                this.eventHandlerList[correlationId] = new EventHandlerCallback(dispId, targetId, handler);
            };
            WebkitScriptMessaging.prototype.unregisterEvent = function (handlerName, methodId, dispId, targetId, callback) {
                var hostArgs = {
                    id: dispId,
                    targetId: targetId
                };
                for (var key in this.eventHandlerList) {
                    if (this.eventHandlerList.hasOwnProperty(key)) {
                        var eventCallback = this.eventHandlerList[key];
                        if (eventCallback.id == dispId && eventCallback.targetId == targetId) {
                            delete this.eventHandlerList[key];
                        }
                    }
                }
                this.invokeMethod(handlerName, methodId, hostArgs, callback);
            };
            WebkitScriptMessaging.prototype.agaveHostCallback = function (callbackId, params) {
                var callbackFunction = this.callbackList[callbackId];
                if (callbackFunction) {
                    var callbacksDone = callbackFunction(params);
                    if (callbacksDone === undefined || callbacksDone === true) {
                        delete this.callbackList[callbackId];
                    }
                }
            };
            WebkitScriptMessaging.prototype.agaveHostEventCallback = function (callbackId, params) {
                var eventCallback = this.eventHandlerList[callbackId];
                if (eventCallback) {
                    eventCallback.handler(params);
                }
            };
            WebkitScriptMessaging.prototype.postWebkitMessage = function (messagingArgs, handlerName, methodId, params, callback) {
                var correlationId = this.generateCorrelationId();
                this.callbackList[correlationId] = callback;
                messagingArgs.methodId = methodId;
                messagingArgs.params = params;
                messagingArgs.callbackId = correlationId;
                messagingArgs.callbackFunction = this.asyncMethodCallbackFunctionName;
                var invokePostMessage = function () {
                    window.webkit.messageHandlers[handlerName].postMessage(JSON.stringify(messagingArgs));
                };
                var currentTimestamp = WebkitScriptMessaging.getCurrentTimeMS();
                if (this.lastMessageTimestamp == null || (currentTimestamp - this.lastMessageTimestamp >= WebkitScriptMessaging.MESSAGE_TIME_DELTA)) {
                    invokePostMessage();
                    this.lastMessageTimestamp = currentTimestamp;
                }
                else {
                    this.lastMessageTimestamp += WebkitScriptMessaging.MESSAGE_TIME_DELTA;
                    setTimeout(function () {
                        invokePostMessage();
                    }, this.lastMessageTimestamp - currentTimestamp);
                }
                return correlationId;
            };
            WebkitScriptMessaging.prototype.generateCorrelationId = function () {
                ++this.callingIndex;
                return this.conversationId + this.callingIndex;
            };
            WebkitScriptMessaging.getCurrentTimeMS = function () {
                return (new Date).getTime();
            };
            WebkitScriptMessaging.MESSAGE_TIME_DELTA = 10;
            return WebkitScriptMessaging;
        })();
        ScriptMessaging.WebkitScriptMessaging = WebkitScriptMessaging;
    })(ScriptMessaging = OSFWebkit.ScriptMessaging || (OSFWebkit.ScriptMessaging = {}));
})(OSFWebkit || (OSFWebkit = {}));
OSF.ScriptMessaging = OSFWebkit.ScriptMessaging;
var WEF;
(function (WEF) {
    WEF.AGAVE_DEFAULT_ICON = "";
    WEF.PageTypeEnum = {
        "ManageApps": 0,
        "Recommendation": 2,
        "Landing": 3,
        "EndNode": 4,
        "Takedown": 5,
        "TermsAndConditions": 6,
        "RateAndReview": 7
    };
    WEF.PageStoreId = {
        "Recommendation": "{98143890-AC66-440E-A448-ED8771A02D52}"
    };
    WEF.StoreTypeEnum = {
        "MarketPlace": 0,
        "Catalog": 1,
        "Exchange": 3,
        "FileShare": 4,
        "Developer": 5,
        "Recommendation": 6,
        "ThisDocument": 8,
        "OneDrive": 9,
        "ExchangeCorporateCatalog": 10
    };
    WEF.AuthType = {
        "Anonymous": "0",
        "MSA": "1",
        "OrgId": "2",
        "ADAL": "3"
    };
    WEF.storeTypes = {
        0: Strings.wefgallery.L_MarketPlaceTabTxt,
        1: Strings.wefgallery.L_CatalogTabTxt,
        3: Strings.wefgallery.L_ExchangeTabTxt,
        4: Strings.wefgallery.L_FileShareTabTxt,
        6: Strings.wefgallery.L_RecommendationTabTxt,
        8: Strings.wefgallery.L_ThisDocumentTabTxt,
        9: Strings.wefgallery.L_OneDriveTabTxt,
        10: Strings.wefgallery.L_ExchangeCCTabTxt
    };
    WEF.InvokeResultCode = {
        "S_OK": 0,
        "E_REQUEST_TIME_OUT": -2147471590,
        "E_USER_NOT_SIGNED_IN": -2147208619,
        "E_CATALOG_ACCESS_DENIED": -2147471591,
        "E_CATALOG_REQUEST_FAILED": -2147471589,
        "E_OEM_NO_NETWORK_CONNECTION": -2147208640,
        "E_PROVIDER_NOT_REGISTERED": -2147208617,
        "E_OEM_CACHE_SHUTDOWN": -2147208637,
        "E_CATALOG_NO_APPS": -1,
        "S_HIDE_PROVIDER": 10,
        "E_OEM_REMOVED_FAILED": -2147209421
    };
    WEF.OemStoreStatus = {
        "ossNotReady": 0,
        "ossSignInRequired": 1,
        "ossRegisteredButNotReady": 2,
        "ossRegisteredReady": 3,
        "ossUnregistered": 4
    };
    WEF.ActionButtonGroups = {
        "InsertCancel": 0,
        "ThisDocument": 1,
        "None": 2
    };
    WEF.OmexMessage = {
        CancelDialog: "ESC_KEY",
        PreloadManifest: "PRELOAD_MANIFEST",
        RefreshRequired: "REFRESH_REQUIRED",
        WindowOpen: "WINDOW_OPEN"
    };
    (function (NodeType) {
        NodeType[NodeType["ELEMENT"] = 1] = "ELEMENT";
        NodeType[NodeType["ATTRIBUTE"] = 2] = "ATTRIBUTE";
        NodeType[NodeType["TEXT"] = 3] = "TEXT";
    })(WEF.NodeType || (WEF.NodeType = {}));
    var NodeType = WEF.NodeType;
    var AgaveInfo = (function () {
        function AgaveInfo() {
            this.displayName = "";
            this.description = "";
            this.providerName = "";
            this.id = "";
            this.width = 0;
            this.height = 0;
            this.iconUrl = "";
            this.targetType = 1;
            this.appVersion = "";
            this.assetId = "";
            this.assetStoreId = "";
            this.storeId = "";
            this.appEndNodeUrl = "";
            this.rateReviewUrl = "";
            this.authType = "";
            this.isAppCommandAddin = false;
            this.hasLoadingError = false;
        }
        AgaveInfo.cmpDisplayName = function (a, b) {
            if (a.displayName && b.displayName) {
                if (a.displayName.localeCompare(b.displayName) > 0) {
                    return 1;
                }
                else {
                    return -1;
                }
            }
            else {
                return -1;
            }
        };
        return AgaveInfo;
    })();
    WEF.AgaveInfo = AgaveInfo;
    var UI;
    (function (UI) {
        UI.SelectedItemDesciptionWidthAdjustment = 75;
        UI.DefaultGalleryWidth = 695;
        UI.DefaultTabMaxWidth = 113;
        UI.DefaultDialogBtnMaxWidth = 150;
        UI.DefaultHeaderHeight = 62;
        UI.DefaultNotificationHeight = 30;
        UI.DefaultSelectedItemHeight = 42;
        UI.DefaultSelectedDescriptionTextMaxWidth = 380;
        UI.DefaultLeftMargin = 26;
        UI.DefaultRightMargin = 25;
        UI.AdjustNotificationHeight = 9;
        UI.DefaultDPI = 96;
        UI.DefaultFooterHeight = 70;
        UI.HiddenFooterHeight = 0;
        UI.HeroMessageMarginTop = 70;
        UI.HeroBtnWidth = 100;
        UI.HeroBtnHeight = 32;
        UI.MenuButtonSide = 25;
        UI.MenuButtonBackgroundSize = 16;
        UI.OptionsMenuWidth = 120;
        UI.DismissButtonSide = 16;
        UI.ConfirmDialogWidth = 300;
        UI.DefaultSeparatorWidth = 5;
        UI.OptionBarElementMargin = 7;
        UI.OptionBarMenuGap = 20;
    })(UI = WEF.UI || (WEF.UI = {}));
    ;
})(WEF || (WEF = {}));
var WEF;
(function (WEF) {
    var WefGalleryHelper;
    (function (WefGalleryHelper) {
        var classN = "class";
        var htmlDir;
        function getDPIScaleXRatio() {
            return (window.screen.logicalXDPI ? window.screen.logicalXDPI : WEF.UI.DefaultDPI) / WEF.UI.DefaultDPI;
        }
        function getDPIScaleYRatio() {
            return (window.screen.logicalYDPI ? window.screen.logicalYDPI : WEF.UI.DefaultDPI) / WEF.UI.DefaultDPI;
        }
        function getProperSignInMessageForStoreType(storeType) {
            var signInMessage = Strings.wefgallery.L_SignInPrompt;
            if (storeType == WEF.StoreTypeEnum.MarketPlace) {
                signInMessage = Strings.wefgallery.L_SignInPromptLiveId;
            }
            return signInMessage;
        }
        function getDocumentDimension(dimensionName) {
            var doc = document.documentElement;
            var maxHeight = 0;
            if (doc != null) {
                maxHeight = Math.max(doc["offset" + dimensionName], doc["scroll" + dimensionName]);
            }
            maxHeight = Math.max(maxHeight, document.body["offset" + dimensionName], document.body["scroll" + dimensionName]);
            return maxHeight;
        }
        function addClass(elmt, val) {
            if (!hasClass(elmt, val)) {
                var className = elmt.getAttribute(classN);
                if (className) {
                    elmt.setAttribute(classN, className + " " + val);
                }
                else {
                    elmt.setAttribute(classN, val);
                }
            }
        }
        WefGalleryHelper.addClass = addClass;
        function createHtmlEncodedTextNode(parent, clsName, text) {
            var div = document.createElement("div");
            parent.appendChild(div);
            addClass(div, clsName);
            var textNode = document.createTextNode(text);
            div.appendChild(textNode);
        }
        WefGalleryHelper.createHtmlEncodedTextNode = createHtmlEncodedTextNode;
        function setHtmlEncodedText(element, text) {
            var textNode = null;
            var childNodes = element.childNodes;
            var childrenCount = childNodes.length;
            for (var j = 0; j < childrenCount; j++) {
                if (childNodes[j].nodeType === WEF.NodeType.TEXT) {
                    textNode = childNodes[j];
                    break;
                }
            }
            if (!textNode) {
                textNode = document.createTextNode(text);
                element.appendChild(textNode);
            }
            else {
                textNode.nodeValue = text;
            }
        }
        WefGalleryHelper.setHtmlEncodedText = setHtmlEncodedText;
        function hasClass(elmt, clsName) {
            var className = elmt.getAttribute(classN);
            return className && className.match(new RegExp('(\\s|^)' + clsName + '(\\s|$)'));
        }
        WefGalleryHelper.hasClass = hasClass;
        function removeClass(elmt, clsName) {
            if (hasClass(elmt, clsName)) {
                var reg = new RegExp('(\\s|^)' + clsName + '(\\s|$)');
                var className = elmt.getAttribute(classN);
                className = className.replace(reg, ' ');
                elmt.setAttribute(classN, className);
            }
        }
        WefGalleryHelper.removeClass = removeClass;
        function getWinWidth() {
            var x = 0;
            if (self.innerWidth) {
                x = self.innerWidth;
            }
            else if (document.documentElement && document.documentElement.clientHeight) {
                x = document.documentElement.clientWidth;
            }
            else if (document.body) {
                x = document.body.clientWidth;
            }
            return x;
        }
        WefGalleryHelper.getWinWidth = getWinWidth;
        function getWinHeight() {
            var y = 0;
            if (self.innerHeight) {
                y = self.innerHeight;
            }
            else if (document.documentElement && document.documentElement.clientHeight) {
                y = document.documentElement.clientHeight;
            }
            else if (document.body) {
                y = document.body.clientHeight;
            }
            return y;
        }
        WefGalleryHelper.getWinHeight = getWinHeight;
        function dpiScale(element) {
            var newWidth = element.offsetWidth * getDPIScaleXRatio();
            var newHeight = element.offsetHeight * getDPIScaleYRatio();
            element.style.width = newWidth + "px";
            element.style.height = newHeight + "px";
        }
        WefGalleryHelper.dpiScale = dpiScale;
        function dpiScaleHeight(element) {
            var newHeight = element.offsetHeight * getDPIScaleYRatio();
            element.style.height = newHeight + "px";
        }
        WefGalleryHelper.dpiScaleHeight = dpiScaleHeight;
        function dpiScaleWidth(element) {
            var newWidth = element.offsetWidth * getDPIScaleXRatio();
            element.style.width = newWidth + "px";
        }
        WefGalleryHelper.dpiScaleWidth = dpiScaleWidth;
        function dpiScaleMarginLeft(element) {
            if (window.getComputedStyle) {
                if (WEF.WefGalleryHelper.getHTMLDir() == "ltr") {
                    var newMarginLeft = parseInt(window.getComputedStyle(element).marginLeft) * getDPIScaleXRatio();
                    element.style.marginLeft = newMarginLeft + "px";
                }
                else {
                    var newMarginRight = parseInt(window.getComputedStyle(element).marginRight) * getDPIScaleXRatio();
                    element.style.marginRight = newMarginRight + "px";
                }
            }
        }
        WefGalleryHelper.dpiScaleMarginLeft = dpiScaleMarginLeft;
        function getDPIXScaledNumber(num) {
            return num * getDPIScaleXRatio();
        }
        WefGalleryHelper.getDPIXScaledNumber = getDPIXScaledNumber;
        function getDPIYScaledNumber(num) {
            return num * getDPIScaleXRatio();
        }
        WefGalleryHelper.getDPIYScaledNumber = getDPIYScaledNumber;
        function clearElementInnerHTML(elementId) {
            var element = document.getElementById(elementId);
            if (element) {
                element.innerHTML = "";
            }
        }
        WefGalleryHelper.clearElementInnerHTML = clearElementInnerHTML;
        function getHTMLDir() {
            if (!htmlDir) {
                htmlDir = document.documentElement.getAttribute("dir");
            }
            return htmlDir;
        }
        WefGalleryHelper.getHTMLDir = getHTMLDir;
        function addSpinWheel(gallery) {
            while (gallery.hasChildNodes()) {
                gallery.removeChild(gallery.firstChild);
            }
            var spinWheelDiv = document.createElement("div");
            WEF.WefGalleryHelper.addClass(spinWheelDiv, "SpinWheel");
            spinWheelDiv.style.width = "100%";
            spinWheelDiv.style.height = "100%";
            gallery.appendChild(spinWheelDiv);
            gallery.setAttribute("aria-busy", "true");
            return spinWheelDiv;
        }
        WefGalleryHelper.addSpinWheel = addSpinWheel;
        function handleErrorCode(errorCode, storeId, storeType, providerStatus) {
            var errorMessage = null;
            var skipShowApps = false;
            var signInRequired = false;
            if (providerStatus) {
                switch (providerStatus) {
                    case WEF.OemStoreStatus.ossSignInRequired:
                        errorMessage = getProperSignInMessageForStoreType(storeType);
                        signInRequired = true;
                        skipShowApps = true;
                        break;
                }
            }
            if (errorMessage == null && errorCode < 0) {
                switch (errorCode) {
                    case WEF.InvokeResultCode.E_REQUEST_TIME_OUT:
                        errorMessage = Strings.wefgallery.L_TimeOutError;
                        break;
                    case WEF.InvokeResultCode.E_USER_NOT_SIGNED_IN:
                        errorMessage = getProperSignInMessageForStoreType(storeType);
                        signInRequired = true;
                        skipShowApps = true;
                        break;
                    case WEF.InvokeResultCode.E_CATALOG_ACCESS_DENIED:
                        errorMessage = Strings.wefgallery.L_AccessDeniedError;
                        skipShowApps = true;
                        break;
                    case WEF.InvokeResultCode.E_CATALOG_REQUEST_FAILED:
                        errorMessage = Strings.wefgallery.L_RequestFailedError;
                        break;
                    case WEF.InvokeResultCode.E_CATALOG_NO_APPS:
                        errorMessage = Strings.wefgallery.L_CatalogNoAppsInstalledError;
                        skipShowApps = true;
                        break;
                    default:
                        errorMessage = Strings.wefgallery.L_GetEntitilementsGeneralError;
                        skipShowApps = true;
                        break;
                }
            }
            if (errorMessage) {
                if (signInRequired) {
                    WEF.IMPage.showError(errorMessage, storeId, Strings.wefgallery.L_SignInLinkText, WEF.IMPage.invokeSignIn);
                }
                else {
                    WEF.IMPage.showError(errorMessage, storeId);
                }
            }
            return skipShowApps;
        }
        WefGalleryHelper.handleErrorCode = handleErrorCode;
        function isHttpsUrl(url) {
            var tmpLink = document.createElement("a");
            tmpLink.href = url;
            return tmpLink.href.split("/")[0].toLowerCase() == "https:";
        }
        WefGalleryHelper.isHttpsUrl = isHttpsUrl;
        function dpiScaleHeightAndWidth(element) {
            dpiScaleHeight(element);
            dpiScaleWidth(element);
        }
        WefGalleryHelper.dpiScaleHeightAndWidth = dpiScaleHeightAndWidth;
        function getDocumentHeight() {
            return getDocumentDimension("Height");
        }
        WefGalleryHelper.getDocumentHeight = getDocumentHeight;
        function getDocumentWidth() {
            return getDocumentDimension("Width");
        }
        WefGalleryHelper.getDocumentWidth = getDocumentWidth;
        function retrieveRefreshRequired() {
            var refreshRequired;
            var retValue = false;
            try {
                if (window.localStorage) {
                    refreshRequired = window.localStorage.getItem("refreshRequired");
                    if (refreshRequired == "true") {
                        retValue = true;
                    }
                }
            }
            catch (e) {
            }
            return retValue;
        }
        WefGalleryHelper.retrieveRefreshRequired = retrieveRefreshRequired;
        function saveRefreshRequired(refreshRequired) {
            try {
                if (window.localStorage) {
                    window.localStorage.setItem("refreshRequired", refreshRequired);
                }
            }
            catch (e) {
            }
        }
        WefGalleryHelper.saveRefreshRequired = saveRefreshRequired;
        function retrieveStoreIdfromStorage() {
            var lastActiveStoreId = null;
            try {
                if (window.localStorage) {
                    lastActiveStoreId = decodeURI(window.localStorage.getItem("lastActiveStoreId"));
                }
            }
            catch (e) {
            }
            return lastActiveStoreId;
        }
        WefGalleryHelper.retrieveStoreIdfromStorage = retrieveStoreIdfromStorage;
        function addEventListener(element, eventName, listener) {
            if (element.attachEvent) {
                element.attachEvent("on" + eventName, listener);
            }
            else if (element.addEventListener) {
                element.addEventListener(eventName, listener, false);
            }
            else {
                element["on" + eventName] = listener;
            }
        }
        WefGalleryHelper.addEventListener = addEventListener;
    })(WefGalleryHelper = WEF.WefGalleryHelper || (WEF.WefGalleryHelper = {}));
})(WEF || (WEF = {}));
var WEF;
(function (WEF) {
    var GalleryItem = (function () {
        function GalleryItem(result, index, focusOnCallBack) {
            this.result = result;
            this.index = index;
            this.galleryItem = null;
            this.moeInnerDiv = null;
            this.focusOnCallBack = focusOnCallBack;
            this.appOptions = null;
            this.itemCreated = false;
        }
        GalleryItem.prototype.displayAgave = function (documentFragment) {
            var moeDiv = document.createElement("div");
            documentFragment.appendChild(moeDiv);
            WEF.WefGalleryHelper.addClass(moeDiv, "Moe");
            moeDiv.setAttribute("data-ri", this.index.toString());
            moeDiv.setAttribute("role", "option");
            var moeInnerDiv = document.createElement("div");
            moeDiv.appendChild(moeInnerDiv);
            WEF.WefGalleryHelper.addClass(moeInnerDiv, "MoeInner");
            WEF.WefGalleryHelper.dpiScale(moeInnerDiv);
            moeInnerDiv.setAttribute("title", this.result.description);
            moeInnerDiv.setAttribute("tabindex", "-1");
            moeInnerDiv.setAttribute("data-inner-moe", this.index.toString());
            this.moeInnerDiv = moeInnerDiv;
            WEF.WefGalleryHelper.dpiScale(moeDiv);
            WEF.WefGalleryHelper.dpiScaleMarginLeft(moeDiv);
            moeDiv.onfocus = function WEF_GalleryItem_displayAgave$onfocus() {
                moeDiv.setAttribute("aria-selected", "true");
            };
            moeDiv.onblur = function WEF_GalleryItem_displayAgave$onblur() {
                moeDiv.setAttribute("aria-selected", "false");
            };
            moeDiv.oncontextmenu = function WEF_GalleryItem_displayAgave$oncontextmenu() {
                return false;
            };
            this.galleryItem = moeDiv;
        };
        GalleryItem.prototype.ShowRateReviewAtGalleryItem = function () {
            return false;
        };
        GalleryItem.prototype.updateImage = function (insertHandler) {
            var _this = this;
            if (!this.galleryItem || !this.moeInnerDiv) {
                return;
            }
            if (!this.itemCreated) {
                WEF.WefGalleryHelper.addEventListener(this.moeInnerDiv, "click", function () {
                    WEF.IMPage.selectGalleryItems(_this.index);
                });
                WEF.WefGalleryHelper.addEventListener(this.moeInnerDiv, "dblclick", function () {
                    insertHandler(_this);
                });
                WEF.WefGalleryHelper.addEventListener(this.moeInnerDiv, "mousedown", function (e) {
                    if (!e)
                        e = event;
                    if (e.which === 3 || e.button === 2) {
                        if (_this.appOptions) {
                            _this.appOptions.popupMenu();
                        }
                    }
                });
                WEF.WefGalleryHelper.addEventListener(this.moeInnerDiv, "mouseover", function () {
                    WEF.WefGalleryHelper.addClass(_this.galleryItem, "mouseover");
                    _this.appOptions.showOptionsButton();
                });
                WEF.WefGalleryHelper.addEventListener(this.moeInnerDiv, "mouseout", function () {
                    WEF.WefGalleryHelper.removeClass(_this.galleryItem, "mouseover");
                    if (!WEF.WefGalleryHelper.hasClass(_this.galleryItem, "selected")) {
                        _this.appOptions.hideOptionsButton();
                    }
                });
                var agaveIconUrl = this.result.iconUrl;
                var tnDiv = document.createElement("div");
                this.moeInnerDiv.appendChild(tnDiv);
                WEF.WefGalleryHelper.addClass(tnDiv, "Tn");
                var detailsDiv = document.createElement("div");
                this.moeInnerDiv.appendChild(detailsDiv);
                WEF.WefGalleryHelper.addClass(detailsDiv, "Details");
                WEF.WefGalleryHelper.dpiScale(detailsDiv);
                WEF.WefGalleryHelper.createHtmlEncodedTextNode(detailsDiv, "Title", this.result.displayName);
                if (this.result.hasLoadingError) {
                    var reloadAnchor = document.createElement("a");
                    reloadAnchor.textContent = Strings.wefgallery.L_Reload_Button_Text;
                    reloadAnchor.onclick = function () { insertHandler(_this); };
                    detailsDiv.appendChild(reloadAnchor);
                    WEF.WefGalleryHelper.dpiScale(reloadAnchor);
                }
                else {
                    WEF.WefGalleryHelper.createHtmlEncodedTextNode(detailsDiv, "Description", this.result.providerName);
                }
                if (this.ShowRateReviewAtGalleryItem()) {
                    var rateDiv = document.createElement("div");
                    detailsDiv.appendChild(rateDiv);
                    var rateLink = document.createElement("a");
                    rateDiv.appendChild(rateLink);
                    rateLink.setAttribute("tabindex", "0");
                    rateLink.setAttribute("id", "rateReviewLink");
                    rateLink.text = Strings.wefgallery.L_OptionsMenu_RateReview_Txt;
                    WEF.WefGalleryHelper.addEventListener(rateLink, "click", function (e) {
                        e.preventDefault();
                        e.stopPropagation();
                        WEF.IMPage.invokeWindowOpen(_this.result.rateReviewUrl);
                    });
                }
                var img = document.createElement("img");
                tnDiv.appendChild(img);
                WEF.WefGalleryHelper.addClass(img, "MoeIcon");
                WEF.WefGalleryHelper.removeClass(tnDiv, "Tn");
                WEF.WefGalleryHelper.addClass(tnDiv, "TnNoBackGround");
                if (!agaveIconUrl || WEF.WefGalleryHelper.isHttpsUrl(window.location.href) && !WEF.WefGalleryHelper.isHttpsUrl(agaveIconUrl)) {
                    agaveIconUrl = WEF.AGAVE_DEFAULT_ICON;
                }
                agaveIconUrl = GalleryItem.errorIconCache[agaveIconUrl] ? GalleryItem.errorIconCache[agaveIconUrl] : agaveIconUrl;
                img.onload = function () {
                    if (img.height >= img.width) {
                        img.style.height = "100%";
                        img.style.width = "auto";
                    }
                    else {
                        img.style.height = "auto";
                        img.style.width = "100%";
                    }
                };
                var iconErrorHandler = function () {
                    var errorIconUrl = img.getAttribute("src");
                    GalleryItem.errorIconCache[errorIconUrl] = WEF.AGAVE_DEFAULT_ICON;
                    img.setAttribute("src", WEF.AGAVE_DEFAULT_ICON);
                };
                img.onerror = iconErrorHandler;
                img.onabort = iconErrorHandler;
                img.setAttribute("src", agaveIconUrl);
                this.appOptions = WEF.IMPage.menuHandler.createAppOptions(this.result);
                var optionsButton = this.appOptions.createOptionsButton(this.index, tnDiv, img);
                if (optionsButton) {
                    this.moeInnerDiv.appendChild(optionsButton);
                }
                var arialLabelDiv = this.galleryItem;
                if (window.navigator.userAgent.indexOf("AppleWebKit") > 0) {
                    arialLabelDiv = this.moeInnerDiv;
                }
                if (optionsButton) {
                    arialLabelDiv.setAttribute("aria-label", Strings.wefgallery.L_GalleryItem_Name_InsertAndOptions_Txt.replace("{0}", this.result.displayName));
                }
                else if (WEF.IMPage.currentStoreType === WEF.StoreTypeEnum.ThisDocument) {
                    arialLabelDiv.setAttribute("aria-label", this.result.displayName);
                }
                else {
                    arialLabelDiv.setAttribute("aria-label", Strings.wefgallery.L_GalleryItem_Name_InsertOnly_Txt.replace("{0}", this.result.displayName));
                }
                if (this.result.hasLoadingError) {
                    var icon = document.createElement("img");
                    icon.className = "MoeErrorStatusIcon";
                    icon.src = "moe_status_icons.png";
                    tnDiv.appendChild(icon);
                    img.style.opacity = "0.5";
                }
            }
            this.itemCreated = true;
        };
        GalleryItem.prototype.setIndex = function (index) {
            this.index = index;
            this.galleryItem.setAttribute("data-ri", index.toString());
            if (this.appOptions) {
                this.appOptions.setAppIndex(index);
            }
        };
        GalleryItem.prototype.dispose = function () {
            this.galleryItem = null;
            this.result = null;
            this.index = null;
        };
        GalleryItem.errorIconCache = {};
        return GalleryItem;
    })();
    WEF.GalleryItem = GalleryItem;
})(WEF || (WEF = {}));
var WEF;
(function (WEF) {
    var WefGalleryPage = (function () {
        function WefGalleryPage(clientFacadeCommon) {
            var _this = this;
            this.providers = {};
            this.currentStoreId = null;
            this.currentStoreType = null;
            this.omexStoreId = null;
            this.hasMarketPlace = false;
            this.currentPageUrl = null;
            this.landingPageUrl = "";
            this.appManagePageUrl = "";
            this.delaying = false;
            this.delayLoad = 200;
            this.delayTime = null;
            this.delayCallbacks = [];
            this.btnAction = null;
            this.btnCancel = null;
            this.btnDone = null;
            this.btnTrustAll = null;
            this.documentAppsMsg = null;
            this.documentAppsMsgText = null;
            this.errorMessage = null;
            this.footer = null;
            this.footerLink = null;
            this.mainPage = null;
            this.gallery = null;
            this.galleryContainer = null;
            this.header = null;
            this.mainTitle = null;
            this.manageATag = null;
            this.uploadATag = null;
            this.menuRight = null;
            this.noAppsMessage = null;
            this.noAppsMessageText = null;
            this.noAppsMessageTitle = null;
            this.notification = null;
            this.notificationDismiss = null;
            this.notificationDismissImg = null;
            this.officeStoreBtn = null;
            this.permissionATag = null;
            this.permissionTextAndLink = null;
            this.permissionTextTR = null;
            this.readMoreATag = null;
            this.selectedDescriptionReadMoreLink = null;
            this.selectedDescriptionText = null;
            this.selectedItem = null;
            this.tabs = null;
            this.uploadMenuDiv = null;
            this.refreshMenuDiv = null;
            this.refreshATag = HTMLAnchorElement = null;
            this.manageMenuDiv = null;
            this.menuRightSeparatorDiv = null;
            this.tabTitles = [];
            this.enterKeyTarget = null;
            this.menuSeparatorWidth = null;
            this.menuRightMaxPossibleWidth = null;
            this.galleryItems = null;
            this.uiState = { "Ready": false, "StoreIdBeforeReady": "", "ErrorBeforeReady": "", "ErrorLinkTextBeforeReady": "", "ErrorLinkHandlerBeforeReady": null };
            this.currentIndex = -1;
            this.currentTabIndex = -1;
            this.results = null;
            this.height = "100%";
            this.width = "100%";
            this.itemsPerRow = null;
            this.leftKeyHandler = null;
            this.rightKeyHandler = null;
            this.upKeyHandler = null;
            this.downKeyHandler = null;
            this.keyHandlers = null;
            this.keyCodePressed = -1;
            this.menuHandler = null;
            this.modalDialog = null;
            this.storeTab = null;
            this.firstTabATag = null;
            this.totalSessionTime = 0;
            this.trustPageSessionTime = 0;
            this.envSetting = {};
            this.isUploadFileDevCatalogEnabled = false;
            this.isAppCommandEnabled = false;
            this.moveLeft = function (event, eventTarget) {
                if (WEF.WefGalleryHelper.hasClass(eventTarget, "TabATag")) {
                    var targetTabIndex = _this.currentTabIndex - 2;
                    if (targetTabIndex < 0) {
                        targetTabIndex = _this.tabs.childNodes.length - 1;
                    }
                    if (targetTabIndex != _this.currentTabIndex) {
                        var targetTab = _this.tabs.childNodes[targetTabIndex];
                        _this.toggleTabSelection(targetTab, null);
                    }
                }
                else {
                    _this.currentIndex--;
                    if (_this.currentIndex >= 0) {
                        _this.selectGalleryItems(_this.currentIndex);
                        if (event.preventDefault) {
                            event.preventDefault();
                        }
                    }
                    else {
                        _this.currentIndex = 0;
                    }
                }
            };
            this.moveRight = function (event, eventTarget, numOfItems) {
                if (WEF.WefGalleryHelper.hasClass(eventTarget, "TabATag")) {
                    var targetTabIndex = _this.currentTabIndex + 2;
                    if (targetTabIndex > _this.tabs.childNodes.length - 1) {
                        targetTabIndex = 0;
                    }
                    if (targetTabIndex != _this.currentTabIndex) {
                        var targetTab = _this.tabs.childNodes[targetTabIndex];
                        _this.toggleTabSelection(targetTab, null);
                    }
                }
                else {
                    _this.currentIndex++;
                    if (_this.currentIndex < numOfItems) {
                        _this.selectGalleryItems(_this.currentIndex);
                        if (event.preventDefault) {
                            event.preventDefault();
                        }
                    }
                    else {
                        _this.currentIndex = numOfItems - 1;
                    }
                }
            };
            this.moveUp = function (event, eventTarget) {
                if (WEF.WefGalleryHelper.hasClass(eventTarget, "Moe") || WEF.WefGalleryHelper.hasClass(eventTarget, "MoeInner")) {
                    _this.currentIndex -= _this.itemsPerRow;
                    if (_this.currentIndex >= 0) {
                        _this.selectGalleryItems(_this.currentIndex);
                        if (event.preventDefault) {
                            event.preventDefault();
                        }
                    }
                    else {
                        _this.currentIndex += _this.itemsPerRow;
                    }
                }
            };
            this.moveDown = function (event, eventTarget, numOfItems) {
                if (WEF.WefGalleryHelper.hasClass(eventTarget, "Moe") || WEF.WefGalleryHelper.hasClass(eventTarget, "MoeInner")) {
                    if (_this.currentIndex >= 0) {
                        _this.currentIndex += _this.itemsPerRow;
                    }
                    else {
                        _this.currentIndex = 0;
                    }
                    if (_this.currentIndex < numOfItems) {
                        _this.selectGalleryItems(_this.currentIndex);
                        if (event.preventDefault) {
                            event.preventDefault();
                        }
                    }
                    else {
                        _this.currentIndex -= _this.itemsPerRow;
                    }
                }
            };
            this.tabKeyHandler = function (event, element) {
                if (!event.shiftKey && (element == _this.tabs.childNodes[_this.currentTabIndex] || element == _this.tabs.childNodes[_this.currentTabIndex].firstChild) && event.preventDefault && _this.currentIndex < 0 && _this.galleryItems && _this.galleryItems.length > 0) {
                    _this.currentIndex = 0;
                    _this.selectGalleryItems(_this.currentIndex, false);
                    event.preventDefault();
                }
                if (!event.shiftKey && element.getAttribute("id") == "RefreshInner" && event.preventDefault && _this.tabs && _this.currentTabIndex >= 0 && _this.currentTabIndex < _this.tabs.childNodes.length) {
                    _this.tabs.childNodes[_this.currentTabIndex].firstChild.focus();
                    event.preventDefault();
                }
                if (event.shiftKey && _this.tabs && (element == _this.tabs.childNodes[_this.currentTabIndex] || element == _this.tabs.childNodes[_this.currentTabIndex].firstChild) && event.preventDefault && _this.refreshATag) {
                    _this.refreshATag.focus();
                    event.preventDefault();
                }
            };
            this.galleryKeyDownHandler = function (e) {
                var numOfItems = 0;
                if (_this.results) {
                    numOfItems = _this.results.length;
                }
                if (!e)
                    e = event;
                for (var i = 0; i < _this.keyHandlers.length; i++) {
                    var keyHandler = _this.keyHandlers[i];
                    if (keyHandler.handleKeyDown(e)) {
                        e.stopPropagation();
                        e.preventDefault();
                        return;
                    }
                }
                var eventTarget = e.srcElement ? e.srcElement : e.target;
                switch (e.keyCode) {
                    case 9:
                        _this.tabKeyHandler(e, eventTarget);
                        break;
                    case 13:
                        _this.enterKeyTarget = eventTarget;
                        e.preventDefault();
                        break;
                    case 27:
                        _this.cancelDialog();
                        break;
                    case 32:
                        if (_this.currentIndex > -1) {
                            _this.selectGalleryItems(_this.currentIndex);
                        }
                        return;
                    case 37:
                        _this.leftKeyHandler(e, eventTarget, numOfItems);
                        break;
                    case 38:
                        _this.upKeyHandler(e, eventTarget);
                        break;
                    case 39:
                        _this.rightKeyHandler(e, eventTarget, numOfItems);
                        break;
                    case 40:
                        _this.downKeyHandler(e, eventTarget, numOfItems);
                        break;
                    default:
                        return;
                }
            };
            this.galleryKeyUpHandler = function (e) {
                if (!e)
                    e = event;
                for (var i = 0; i < _this.keyHandlers.length; i++) {
                    var keyHandler = _this.keyHandlers[i];
                    if (keyHandler.handleKeyUp(e)) {
                        e.stopPropagation();
                        e.preventDefault();
                        return;
                    }
                }
                var eventTarget = e.srcElement ? e.srcElement : e.target;
                switch (e.keyCode) {
                    case 13:
                        _this.executeButtonCommand(eventTarget, e);
                        break;
                }
            };
            this.resizeHandler = function () {
                _this.uiState.Ready = false;
                var winHeight = WEF.WefGalleryHelper.getWinHeight().toString();
                var winWidth = WEF.WefGalleryHelper.getWinWidth().toString();
                if (_this.height != winHeight || _this.width != winWidth) {
                    _this.height = winHeight;
                    _this.width = winWidth;
                    _this.setGalleryHeight();
                    _this.delayLoadVisibleImages();
                    var newMaxWidth, widthIncreaseRatio = (_this.width) / WEF.UI.DefaultGalleryWidth;
                    newMaxWidth = WEF.WefGalleryHelper.getDPIXScaledNumber(WEF.UI.DefaultDialogBtnMaxWidth) * widthIncreaseRatio;
                    _this.btnAction.style.maxWidth = newMaxWidth + "px";
                    _this.btnCancel.style.maxWidth = newMaxWidth + "px";
                    _this.btnTrustAll.style.maxWidth = newMaxWidth + "px";
                    _this.btnDone.style.maxWidth = newMaxWidth + "px";
                    _this.setOptionBarElementMaxSize(_this.tabTitles);
                    if (_this.currentIndex >= 0 && _this.galleryItems && _this.galleryItems.length > 0) {
                        var item = _this.galleryItems[_this.currentIndex];
                        if (WEF.WefGalleryHelper.hasClass(item.galleryItem, "selected")) {
                            _this.setSelectedItemWidth();
                        }
                    }
                }
                _this.menuHandler.hideMenu(true);
                _this.modalDialog.positionDialog();
                _this.uiState.Ready = true;
                if (_this.uiState.ErrorLinkTextBeforeReady && _this.uiState.ErrorLinkHandlerBeforeReady) {
                    _this.showError(_this.uiState.ErrorBeforeReady, _this.uiState.StoreIdBeforeReady, _this.uiState.ErrorLinkTextBeforeReady, _this.uiState.ErrorLinkHandlerBeforeReady);
                }
                else {
                    _this.showError(_this.uiState.ErrorBeforeReady, _this.uiState.StoreIdBeforeReady);
                }
            };
            this.loadVisibleImages = function () {
                if (new Date().getTime() - _this.delayTime < _this.delayLoad && _this.delaying) {
                    setTimeout(_this.loadVisibleImages, _this.delayLoad);
                }
                else {
                    var gallery = _this.gallery;
                    if (gallery && gallery.children.length > 0) {
                        var foundFirst = false;
                        var foundLast = false;
                        var itemsPerRow = _this.getItemsPerRow();
                        if (itemsPerRow > 0) {
                            var offset = _this.galleryItems[0].galleryItem.offsetHeight;
                            if (_this.currentIndex < itemsPerRow && _this.keyCodePressed == 40) {
                                _this.gallery.scrollTop = 0;
                                _this.keyCodePressed = -1;
                            }
                            var displayTop = gallery.scrollTop + offset;
                            var displayBottom = displayTop + (gallery.clientHeight * 2);
                            var item;
                            for (var i = 0; i < _this.galleryItems.length; i += itemsPerRow) {
                                for (var j = 0; j < itemsPerRow && (i + j) < _this.galleryItems.length; j++) {
                                    item = _this.galleryItems[i + j].galleryItem;
                                    if (item && (j > 0 || (item.offsetTop + item.clientHeight >= displayTop
                                        && item.offsetTop < displayBottom))) {
                                        if (_this.galleryItems) {
                                            _this.galleryItems[i + j].updateImage(_this.insertItem);
                                        }
                                    }
                                    else {
                                        if (foundFirst) {
                                            foundLast = true;
                                        }
                                        break;
                                    }
                                    foundFirst = true;
                                }
                                if (foundLast) {
                                    break;
                                }
                            }
                        }
                        _this.delaying = false;
                        _this.gallery.setAttribute("aria-busy", "false");
                    }
                    if (_this.delaying) {
                        setTimeout(_this.loadVisibleImages, 3000);
                        _this.delaying = false;
                    }
                    else {
                        while (_this.delayCallbacks.length > 0) {
                            var callback = _this.delayCallbacks.pop();
                            callback();
                        }
                    }
                }
            };
            this.insertItem = function (item) {
                throw "Should be implemented by WefGalleryRich.ts or WefGalleryWac.ts.";
            };
            this.showEntitlements = function (storeId, refresh, callback) {
                throw "Should be implemented by WefGalleryRich.ts or WefGalleryWac.ts.";
            };
            this.invokeSignIn = function () {
                throw "Should be implemented by WefGalleryRich.ts or WefGalleryWac.ts.";
            };
            if (WEF.WefGalleryHelper.getHTMLDir() == "ltr") {
                this.leftKeyHandler = this.moveLeft;
                this.rightKeyHandler = this.moveRight;
            }
            else {
                this.leftKeyHandler = this.moveRight;
                this.rightKeyHandler = this.moveLeft;
            }
            this.upKeyHandler = this.moveUp;
            this.downKeyHandler = this.moveDown;
            this.clientFacadeCommon = clientFacadeCommon;
            this.envSetting = this.clientFacadeCommon.getEnvSetting();
            this.isAppCommandEnabled = this.envSetting["IsAppCommandEnabled"] === true;
        }
        WefGalleryPage.prototype.showHideRightMenuButtons = function (showManageApp, showRefresh) {
            this.menuRight.style.display = "block";
            var hideRightMenu = !showManageApp && !showRefresh;
            var showUploadAddin = !hideRightMenu && !showManageApp && this.isUploadFileDevCatalogEnabled;
            if (showUploadAddin) {
                this.menuRight.children[0].style.display = 'inline-block';
            }
            else {
                this.menuRight.children[0].style.display = 'none';
            }
            if (showManageApp) {
                this.menuRight.children[1].style.display = 'inline-block';
            }
            else {
                this.menuRight.children[1].style.display = 'none';
            }
            if (showRefresh) {
                this.menuRight.children[3].style.display = 'inline-block';
            }
            else {
                this.menuRight.children[3].style.display = 'none';
            }
            if ((showUploadAddin || showManageApp) && showRefresh) {
                this.menuRight.children[2].style.display = 'inline-block';
            }
            else {
                this.menuRight.children[2].style.display = 'none';
            }
        };
        WefGalleryPage.prototype.showFooter = function () {
            this.footer.style.visibility = 'visible';
            this.footer.style.height = WEF.WefGalleryHelper.getDPIYScaledNumber(WEF.UI.DefaultFooterHeight) + "px";
        };
        WefGalleryPage.prototype.showActionButtons = function (buttonGroup) {
            if (buttonGroup == WEF.ActionButtonGroups.InsertCancel) {
                this.btnAction.style.display = "inline";
                this.btnCancel.style.display = "inline";
                this.btnTrustAll.style.display = "none";
                this.btnDone.style.display = "none";
                this.selectedDescriptionReadMoreLink.style.display = "none";
                this.permissionTextAndLink.style.display = "none";
                this.disableNarratorOnControl(this.permissionTextTR);
            }
            else if (buttonGroup == WEF.ActionButtonGroups.ThisDocument) {
                this.btnAction.style.display = "none";
                this.btnCancel.style.display = "none";
                this.btnTrustAll.style.display = "inline";
                this.btnDone.style.display = "inline";
                this.selectedDescriptionReadMoreLink.style.display = "inline-block";
                this.permissionTextAndLink.style.display = "inline-block";
                this.footerLink.style.display = 'none';
            }
            else {
                this.btnAction.style.display = "none";
                this.btnCancel.style.display = "none";
                this.btnTrustAll.style.display = "none";
                this.btnDone.style.display = "inline";
                this.selectedDescriptionReadMoreLink.style.display = "none";
                this.permissionTextAndLink.style.display = "none";
                this.footerLink.style.display = 'none';
                this.documentAppsMsg.style.display = 'none';
            }
        };
        WefGalleryPage.prototype.getTabTooltip = function (storeType) {
            var tooltip = "";
            switch (storeType) {
                case WEF.StoreTypeEnum.MarketPlace:
                    tooltip = Strings.wefgallery.L_MarketPlaceTab_Tooltip;
                    break;
                case WEF.StoreTypeEnum.Catalog:
                    tooltip = Strings.wefgallery.L_CatalogTab_Tooltip;
                    break;
                case WEF.StoreTypeEnum.FileShare:
                    tooltip = Strings.wefgallery.L_FileShareTab_Tooltip;
                    break;
                case WEF.StoreTypeEnum.Recommendation:
                    tooltip = Strings.wefgallery.L_RecommendationTab_Tooltip;
                    break;
                case WEF.StoreTypeEnum.ThisDocument:
                    tooltip = Strings.wefgallery.L_ThisDocumentTab_Tooltip;
                    break;
                case WEF.StoreTypeEnum.ExchangeCorporateCatalog:
                    tooltip = Strings.wefgallery.L_ExchangeCCTab_Tooltip;
                    break;
            }
            return tooltip;
        };
        WefGalleryPage.prototype.getCurrentProviderHResult = function () {
            var hres = 0;
            if (this.currentStoreId) {
                hres = this.providers[this.currentStoreId][2];
            }
            return hres;
        };
        WefGalleryPage.prototype.getCurrentProviderStatus = function () {
            var status = 0;
            if (this.currentStoreId) {
                status = this.providers[this.currentStoreId][1];
            }
            return status;
        };
        WefGalleryPage.prototype.getFeaturedPageUrl = function () {
            if (!this.currentPageUrl) {
                this.currentPageUrl = this.getPageUrl(WEF.PageTypeEnum.Recommendation);
            }
            return this.currentPageUrl;
        };
        WefGalleryPage.prototype.getLandingPageUrl = function () {
            if (!this.landingPageUrl) {
                this.landingPageUrl = this.getPageUrl(WEF.PageTypeEnum.Landing);
            }
            return this.landingPageUrl;
        };
        WefGalleryPage.prototype.getAppManagePageUrl = function () {
            if (!this.appManagePageUrl) {
                this.appManagePageUrl = this.getPageUrl(WEF.PageTypeEnum.ManageApps);
            }
            return this.appManagePageUrl;
        };
        WefGalleryPage.prototype.executeButtonCommand = function (element, event) {
            this.menuHandler.hideMenu(true);
            if (element != this.enterKeyTarget) {
                return;
            }
            if (WEF.WefGalleryHelper.hasClass(element, "MoeInner") || WEF.WefGalleryHelper.hasClass(element, "Moe")) {
                this.insertSelectedItem();
            }
            else if (WEF.WefGalleryHelper.hasClass(element, "TabATag")) {
                var storeId = element.parentElement.getAttribute("data-storeId");
                if (storeId) {
                    this.toggleTabSelection(element.parentElement, null);
                }
                else {
                    this.showEntitlements(this.currentStoreId, true, null);
                }
            }
            else if (element.getAttribute("id") == "BtnAction") {
                this.insertSelectedItem();
            }
            else if (element.getAttribute("id") == "BtnCancel" || element.getAttribute("id") == "BtnDone") {
                this.cancelDialog();
            }
            else if (element.getAttribute("id") == "ManageInner") {
                this.launchAppManagePage();
            }
            else if (element.getAttribute("id") == "RefreshInner") {
                this.showEntitlements(this.currentStoreId, true, null);
            }
            else if (element.getAttribute("id") == "FooterLinkATag") {
                this.gotoStore();
            }
            else if (element.getAttribute("id") == "linkId") {
                this.invokeSignIn();
            }
            else if (element.getAttribute("id") == "rateReviewLink") {
                if (this.results != null && this.results.length > 0) {
                    WEF.IMPage.invokeWindowOpen(this.results[0].rateReviewUrl);
                }
            }
            else if (WEF.WefGalleryHelper.hasClass(element, "OptionsButton")) {
                element.click();
            }
        };
        WefGalleryPage.prototype.restoreFooterLink = function () {
            WEF.WefGalleryHelper.clearElementInnerHTML('SelectedItemTitle');
            WEF.WefGalleryHelper.clearElementInnerHTML('SelectedItemDescription');
            if (this.currentStoreType != WEF.StoreTypeEnum.ThisDocument) {
                if (this.hasMarketPlace) {
                    this.footerLink.style.display = 'block';
                }
                this.selectedItem.style.display = 'none';
                if (!WEF.WefGalleryHelper.hasClass(this.btnAction, 'disabled')) {
                    WEF.WefGalleryHelper.addClass(this.btnAction, 'disabled');
                }
                this.btnAction.setAttribute('disabled', 'true');
            }
        };
        WefGalleryPage.prototype.toggleTabSelection = function (selectedTabDiv, callback) {
            this.cleanUpGallery();
            var selectedTabId = selectedTabDiv.getAttribute("id");
            var len = this.tabs.childNodes.length, i, child, tabId;
            for (i = 0; i < len; i++) {
                child = this.tabs.childNodes[i];
                if (child.attributes && WEF.WefGalleryHelper.hasClass(child, "TextNav")) {
                    WEF.WefGalleryHelper.removeClass(child.firstChild, "TabSelected");
                    child.setAttribute("tabIndex", "-1");
                    child.firstChild.setAttribute("aria-selected", "false");
                    child.firstChild.removeAttribute("aria-controls");
                    tabId = child.getAttribute("id");
                    if (tabId == selectedTabId) {
                        this.currentTabIndex = i;
                        child.setAttribute("tabIndex", "0");
                        if (child.firstChild != null) {
                            child.firstChild.focus();
                        }
                        WEF.WefGalleryHelper.addClass(child.firstChild, "TabSelected");
                        child.firstChild.setAttribute("aria-selected", "true");
                        child.firstChild.setAttribute("aria-controls", "GalleryContainer");
                        var storeId = child.getAttribute("data-storeId");
                        var storeType = parseInt(child.getAttribute("data-storeType"));
                        if (this.currentStoreId != storeId) {
                            this.currentIndex = -1;
                        }
                        this.currentStoreId = storeId;
                        this.currentStoreType = storeType;
                        this.saveStoreId(this.currentStoreId);
                        if (storeId && storeType != WEF.StoreTypeEnum.Recommendation) {
                            this.restoreFooterLink();
                            this.showFooter();
                            this.showEntitlements(storeId, false, callback);
                            this.setGalleryHeight();
                        }
                        else {
                            var pageUrl = child.getAttribute("data-PageUrl");
                            this.showContentPage(pageUrl);
                        }
                        this.refreshATag.setAttribute("title", Strings.wefgallery.L_WefDialog_RefreshButton_Tooltip.replace("{0}", child.firstChild.textContent));
                    }
                }
            }
        };
        WefGalleryPage.prototype.initializeGalleryUI = function (providers, resetToMarketPlace) {
            var _this = this;
            if (providers == undefined || providers.length === 0) {
                return false;
            }
            var provider, providersArray = [];
            var len = providers.length, tempStoreId, tempStoreType, tempStatus, tempHResult;
            var hasOneDriveCatalogProvider = false;
            for (var i = 0; i < len; i++) {
                provider = providers[i].toArray ? providers[i].toArray() : providers[i];
                tempStoreId = provider[0];
                tempStoreType = provider[1];
                tempStatus = provider[2];
                tempHResult = provider[3];
                if (tempStoreType === WEF.StoreTypeEnum.Developer) {
                    continue;
                }
                if (tempHResult != WEF.InvokeResultCode.S_HIDE_PROVIDER) {
                    if (tempStoreType === WEF.StoreTypeEnum.OneDrive) {
                        hasOneDriveCatalogProvider = true;
                    }
                    else {
                        providersArray.push([tempStoreId.toString(), tempStoreType, tempStatus, tempHResult]);
                    }
                }
                if (tempStoreType === WEF.StoreTypeEnum.MarketPlace) {
                    this.hasMarketPlace = true;
                    this.omexStoreId = tempStoreId;
                }
            }
            if (this.hasMarketPlace) {
                providersArray.push([WEF.PageStoreId.Recommendation, WEF.StoreTypeEnum.Recommendation, 0, 0]);
                this.footerLink.style.display = 'block';
            }
            else {
                this.footerLink.style.display = 'none';
            }
            if (hasOneDriveCatalogProvider) {
                providersArray.push([WEF.StoreTypeEnum.OneDrive, WEF.StoreTypeEnum.OneDrive, 0, 0]);
            }
            len = providersArray.length;
            if (len === 0) {
                return false;
            }
            var isCurrentSet = false;
            var lastStoreId;
            if (resetToMarketPlace && this.hasMarketPlace) {
                lastStoreId = this.omexStoreId;
            }
            else {
                lastStoreId = this.retrieveStoreId();
            }
            if (this.hasMarketPlace && lastStoreId && tempStoreType !== WEF.StoreTypeEnum.ThisDocument) {
                if (WEF.PageStoreId.Recommendation === lastStoreId) {
                    this.currentStoreId = lastStoreId;
                    this.currentStoreType = WEF.StoreTypeEnum.Recommendation;
                    this.currentPageUrl = this.getPageUrl(WEF.PageTypeEnum.Recommendation);
                    isCurrentSet = true;
                }
                else {
                    for (var i = 0; i < len; i++) {
                        if (providersArray[i][0] === lastStoreId) {
                            this.currentStoreId = lastStoreId;
                            this.currentStoreType = providersArray[i][1];
                            isCurrentSet = true;
                            break;
                        }
                    }
                }
            }
            this.tabs.setAttribute("role", "tablist");
            while (this.tabs.hasChildNodes()) {
                this.tabs.removeChild(this.tabs.firstChild);
            }
            if (!isCurrentSet) {
                this.currentStoreId = providersArray[0][0];
                this.currentStoreType = providersArray[0][1];
            }
            var tabOrder = 0;
            var selectedTab = null;
            var createdTab = null;
            this.tabTitles = [];
            for (var i = 0; i < len; i++) {
                provider = providersArray[i];
                tempStoreId = provider[0];
                tempStoreType = provider[1];
                tempStatus = provider[2];
                tempHResult = provider[3];
                this.providers[tempStoreId] = [tempStoreType, tempStatus, tempHResult];
                var tabName = WEF.storeTypes[tempStoreType];
                if (tabName) {
                    delete WEF.storeTypes[tempStoreType];
                    tabOrder++;
                    if (tempStoreId === WEF.StoreTypeEnum.OneDrive) {
                        this.checkAndCreateOneDriveProviderTab(this.tabs, tabOrder, tabName, tempStoreId, tempStoreType);
                    }
                    else {
                        createdTab = this.createTab(this.tabs, tabOrder, tabName, tempStoreId, tempStoreType);
                        this.tabTitles.push(createdTab);
                    }
                    if (tempStoreId === this.currentStoreId) {
                        selectedTab = createdTab;
                        WEF.WefGalleryHelper.addClass(selectedTab.firstChild, "TabSelected");
                        selectedTab.firstChild.setAttribute("aria-selected", "true");
                        selectedTab.firstChild.setAttribute("aria-controls", "GalleryContainer");
                    }
                    if (tempStoreType == WEF.StoreTypeEnum.Recommendation) {
                        this.storeTab = createdTab;
                    }
                }
            }
            this.setOptionBarElementMaxSize(this.tabTitles);
            var selectedTabTitle = null;
            if (this.tabs.childNodes.length > 0) {
                if (selectedTab) {
                    WEF.WefGalleryHelper.addClass(selectedTab.childNodes[0], "selected");
                    selectedTab.childNodes[0].focus();
                    selectedTabTitle = selectedTab.childNodes[0].textContent;
                }
                else if (this.tabs.childNodes[0].childNodes.length > 0) {
                    WEF.WefGalleryHelper.addClass(this.tabs.childNodes[0].childNodes[0], "selected");
                    selectedTabTitle = this.tabs.childNodes[0].childNodes[0].textContent;
                }
            }
            var child = null;
            for (var i = 0; i < this.tabs.childNodes.length; i++) {
                child = this.tabs.childNodes[i];
                if (child.getAttribute("data-storeId") === this.currentStoreId) {
                    this.currentTabIndex = i;
                    child.setAttribute("tabIndex", "0");
                    break;
                }
            }
            if (this.isUploadFileDevCatalogEnabled) {
                var dropDownArrow = document.createElement("img");
                dropDownArrow.setAttribute("src", "./DropDownArrow_16x16x32.png");
                dropDownArrow.id = "DropDownArrow";
                this.manageATag.appendChild(dropDownArrow);
            }
            this.manageATag.setAttribute("tabIndex", "0");
            this.manageATag.setAttribute("title", Strings.wefgallery.L_WefDialog_ManageButton_Tooltip);
            this.manageATag.setAttribute("role", "link");
            this.uploadATag.setAttribute("tabIndex", "0");
            this.uploadATag.setAttribute("title", Strings.wefgallery.L_AddinCommands_UploadMyAddin_Txt);
            var refreshCurrentTab = function () {
                _this.cleanUpGallery();
                _this.restoreFooterLink();
                _this.showContent(true);
            };
            this.refreshATag = document.getElementById('RefreshInner');
            this.refreshATag.setAttribute("title", Strings.wefgallery.L_WefDialog_RefreshButton_Tooltip.replace("{0}", selectedTabTitle));
            this.refreshATag.onclick = function WEF_WefGalleryPage_initializeGalleryUI_refreshATag$onclick() { refreshCurrentTab(); };
            this.refreshATag.setAttribute("tabIndex", "0");
            this.refreshATag.setAttribute("role", "link");
            var footerLinkATag = document.getElementById('FooterLinkATag');
            footerLinkATag.setAttribute("tabIndex", "0");
            footerLinkATag.setAttribute("title", Strings.wefgallery.L_Footer_Link_Text_Tooltip);
            footerLinkATag.setAttribute("role", "link");
            this.documentAppsMsg.setAttribute("title", Strings.wefgallery.L_TrustUx_AppsMessage);
            this.documentAppsMsg.firstChild.innerText = Strings.wefgallery.L_TrustUx_AppsMessage;
            this.readMoreATag.setAttribute("tabIndex", "0");
            this.readMoreATag.setAttribute("title", Strings.wefgallery.L_TrustUx_ReadMoreLink_Txt_Tooltip);
            this.readMoreATag.setAttribute("role", "link");
            this.permissionATag.setAttribute("tabIndex", "0");
            this.permissionATag.setAttribute("title", Strings.wefgallery.L_Permission_Link_Txt_Tooltip);
            this.permissionATag.setAttribute("role", "link");
            this.permissionTextAndLink.setAttribute("title", Strings.wefgallery.L_Permission_Link_Txt_Tooltip);
            this.btnAction.setAttribute("tabIndex", "0");
            if (this.isAppCommandEnabled) {
                this.btnAction.value = Strings.wefgallery.L_OK_Button_Txt;
                this.btnAction.title = Strings.wefgallery.L_OK_Button_Txt_Tooltip;
            }
            else {
                this.btnAction.title = Strings.wefgallery.L_Action_Button_Txt_Tooltip;
            }
            this.btnCancel.setAttribute("tabIndex", "0");
            this.btnCancel.setAttribute("title", Strings.wefgallery.L_Cancel_Button_Text_Tooltip);
            this.btnTrustAll.setAttribute("tabIndex", "0");
            this.btnTrustAll.setAttribute("title", Strings.wefgallery.L_TrustAll_Button_Txt_Tooltip);
            this.btnDone.setAttribute("tabIndex", "0");
            this.btnDone.setAttribute("title", Strings.wefgallery.L_Done_Button_Txt_Tooltip);
            this.noAppsMessage.setAttribute("title", Strings.wefgallery.L_OfficeStore_Button_Tooltip.replace("{0}", this.officeStoreBtn.value));
            this.noAppsMessage.style.marginTop = WEF.WefGalleryHelper.getDPIXScaledNumber(WEF.UI.HeroMessageMarginTop) + "px";
            this.noAppsMessageTitle.innerHTML = Strings.wefgallery.L_NoAppsMessageTitle;
            this.noAppsMessageText.innerHTML = Strings.wefgallery.L_NoAppsMessageText;
            this.officeStoreBtn.style.width = WEF.WefGalleryHelper.getDPIXScaledNumber(WEF.UI.HeroBtnWidth) + "px";
            this.officeStoreBtn.style.height = WEF.WefGalleryHelper.getDPIXScaledNumber(WEF.UI.HeroBtnHeight) + "px";
            this.overrideButtonTooltip();
            return true;
        };
        WefGalleryPage.prototype.showContent = function (forceRefresh) {
            if (this.currentStoreId == WEF.PageStoreId.Recommendation) {
                this.showContentPage(this.currentPageUrl);
            }
            else {
                this.showEntitlements(this.currentStoreId, forceRefresh, null);
            }
        };
        WefGalleryPage.prototype.saveStoreId = function (currentStoreId) {
            try {
                if (window.localStorage) {
                    window.localStorage.setItem("lastActiveStoreId", encodeURI(currentStoreId));
                }
            }
            catch (e) {
            }
        };
        WefGalleryPage.prototype.disableNarratorOnControl = function (ctl) {
            ctl.setAttribute("role", "presentation");
            ctl.setAttribute("aria-hidden", "true");
            ctl.setAttribute("tabindex", "-1");
        };
        WefGalleryPage.prototype.createTab = function (tabsDiv, tabOrder, tabName, storeId, storeType) {
            var me = this;
            if (tabsDiv.childNodes.length != 0) {
                var separatorDiv = document.createElement('div');
                WEF.WefGalleryHelper.addClass(separatorDiv, "separator");
                separatorDiv.innerHTML = "|";
                this.disableNarratorOnControl(separatorDiv);
                tabsDiv.appendChild(separatorDiv);
            }
            var pageUrl = WEF.PageStoreId.Recommendation === storeId ? this.getFeaturedPageUrl() : null;
            var tabDiv = document.createElement('div');
            WEF.WefGalleryHelper.addClass(tabDiv, "TextNav");
            tabsDiv.appendChild(tabDiv);
            var aTag = document.createElement('a');
            WEF.WefGalleryHelper.addClass(aTag, "TabATag");
            WEF.WefGalleryHelper.setHtmlEncodedText(aTag, tabName);
            var tooltip = this.getTabTooltip(storeType);
            aTag.setAttribute("title", tooltip);
            aTag.setAttribute("tabIndex", "-1");
            aTag.setAttribute("role", "tab");
            tabDiv.appendChild(aTag);
            if (tabOrder == 1) {
                aTag.focus();
                this.firstTabATag = aTag;
            }
            tabDiv.setAttribute("id", tabName);
            tabDiv.setAttribute("tabIndex", "-1");
            tabDiv.setAttribute("data-storeId", storeId);
            tabDiv.setAttribute("data-storeType", storeType.toString());
            tabDiv.setAttribute("role", "presentation");
            if (pageUrl) {
                tabDiv.setAttribute("data-pageUrl", pageUrl);
            }
            tabDiv.onclick = function WEF_WefGalleryPage_createTab_tabDiv$onclick() { me.toggleTabSelection(this, null); };
            tabDiv.onfocus = function WEF_WefGalleryPage_createTab_tabDiv$onfocus() {
                aTag.focus();
            };
            return tabDiv;
        };
        WefGalleryPage.prototype.galleryScrollHandler = function () {
            this.menuHandler.hideMenu(true);
            this.delayLoadVisibleImages();
        };
        WefGalleryPage.prototype.storeStaticElementRealSize = function () {
            this.menuSeparatorWidth = WEF.UI.DefaultSeparatorWidth;
            if (this.menuRightSeparatorDiv.offsetWidth != 0) {
                this.menuSeparatorWidth = this.menuRightSeparatorDiv.offsetWidth;
            }
            var uploadMenuWidth = 0;
            if (this.isUploadFileDevCatalogEnabled) {
                uploadMenuWidth = this.uploadMenuDiv.offsetWidth;
            }
            this.menuRightMaxPossibleWidth = Math.max(uploadMenuWidth, this.manageMenuDiv.offsetWidth) + this.menuSeparatorWidth + this.refreshMenuDiv.offsetWidth + WEF.WefGalleryHelper.getDPIXScaledNumber(WEF.UI.OptionBarElementMargin) * 3;
        };
        WefGalleryPage.prototype.setOptionBarElementMaxSize = function (tabTitles) {
            if (tabTitles == null || tabTitles.length == 0)
                return;
            for (var i = 0; i < tabTitles.length; i++) {
                tabTitles[i].style.maxWidth = "none";
            }
            this.refreshMenuDiv.style.maxWidth = "none";
            this.uploadMenuDiv.style.maxWidth = "none";
            this.manageMenuDiv.style.maxWidth = "none";
            var optionBarTotalWidth = WEF.WefGalleryHelper.getDPIXScaledNumber(WEF.UI.DefaultLeftMargin) + this.tabs.offsetWidth + WEF.WefGalleryHelper.getDPIXScaledNumber(WEF.UI.OptionBarMenuGap) + this.menuRightMaxPossibleWidth + WEF.UI.DefaultRightMargin;
            if (optionBarTotalWidth > WEF.WefGalleryHelper.getWinWidth()) {
                var widthForAllTitleText = WEF.WefGalleryHelper.getWinWidth() - WEF.WefGalleryHelper.getDPIXScaledNumber(WEF.UI.DefaultLeftMargin) -
                    WEF.UI.DefaultRightMargin - WEF.WefGalleryHelper.getDPIXScaledNumber(WEF.UI.OptionBarMenuGap);
                widthForAllTitleText -= this.menuSeparatorWidth * tabTitles.length;
                widthForAllTitleText -= WEF.WefGalleryHelper.getDPIXScaledNumber(WEF.UI.OptionBarElementMargin) * (tabTitles.length * 2 + 2);
                var titleTextMaxWidth = (widthForAllTitleText / (tabTitles.length + 2)) + "px";
                for (var i = 0; i < tabTitles.length; i++) {
                    tabTitles[i].style.maxWidth = titleTextMaxWidth;
                }
                this.refreshMenuDiv.style.maxWidth = titleTextMaxWidth;
                this.uploadMenuDiv.style.maxWidth = titleTextMaxWidth;
                this.manageMenuDiv.style.maxWidth = titleTextMaxWidth;
            }
        };
        WefGalleryPage.prototype.setGalleryHeight = function () {
            var galleryContainerHeight = WEF.WefGalleryHelper.getWinHeight() - this.header.offsetHeight - this.footer.offsetHeight;
            if (this.galleryContainer && galleryContainerHeight > 0 &&
                (galleryContainerHeight != this.galleryContainer.offsetHeight || this.footer && this.footer.style.top === "")) {
                this.galleryContainer.style.height = galleryContainerHeight + "px";
                this.galleryContainer.style.top = this.header.offsetHeight + "px";
                var galleryHeight = galleryContainerHeight;
                if (this.currentStoreType == WEF.StoreTypeEnum.ThisDocument) {
                    galleryHeight = galleryHeight - this.documentAppsMsg.offsetHeight * 2;
                }
                this.gallery.style.height = galleryHeight + "px";
                var footerTop = galleryContainerHeight + this.header.offsetHeight;
                this.footer.style.top = footerTop + "px";
            }
        };
        WefGalleryPage.prototype.setSelectedItemWidth = function () {
            var newWidth = WEF.WefGalleryHelper.getWinWidth() - WEF.WefGalleryHelper.getDPIXScaledNumber(WEF.UI.SelectedItemDesciptionWidthAdjustment);
            if (this.currentStoreType == WEF.StoreTypeEnum.ThisDocument) {
                newWidth = newWidth - this.btnTrustAll.offsetWidth - this.btnDone.offsetWidth;
            }
            else {
                newWidth = newWidth - this.btnAction.offsetWidth - this.btnCancel.offsetWidth;
            }
            this.selectedItem.style.width = newWidth + "px";
            this.selectedItem.style.height = WEF.WefGalleryHelper.getDPIYScaledNumber(WEF.UI.DefaultSelectedItemHeight) + "px";
            var marginLeft = parseInt(window.getComputedStyle ? window.getComputedStyle(this.selectedItem).marginLeft : this.selectedItem.style.marginLeft);
            this.selectedDescriptionText.style.maxWidth = (newWidth - marginLeft - this.selectedDescriptionReadMoreLink.offsetWidth) + "px";
            this.footerLink.style.width = newWidth + "px";
        };
        WefGalleryPage.prototype.deSelectBtnAction = function () {
            this.selectedItem.style.display = 'none';
            if (this.currentStoreType != WEF.StoreTypeEnum.ThisDocument) {
                if (this.hasMarketPlace) {
                    this.footerLink.style.display = 'block';
                }
                WEF.WefGalleryHelper.addClass(this.btnAction, 'disabled');
                this.btnAction.disabled = true;
            }
        };
        WefGalleryPage.prototype.cleanUpGallery = function () {
            this.menuHandler.hideMenu(true);
            this.noAppsMessage.style.display = 'none';
            this.notification.style.display = 'none';
            this.errorMessage.innerHTML = "";
            this.deSelectBtnAction();
            if (this.galleryItems != null) {
                var i;
                for (i = 0; i < this.galleryItems.length; i++) {
                    this.galleryItems[i].dispose();
                    delete this.galleryItems[i];
                }
            }
            this.galleryItems = null;
            while (this.gallery.hasChildNodes()) {
                this.gallery.removeChild(this.gallery.firstChild);
            }
            this.header.style.height = WEF.WefGalleryHelper.getDPIYScaledNumber(WEF.UI.DefaultHeaderHeight) + "px";
            this.setGalleryHeight();
            this.trustPageSessionTime = 0;
        };
        WefGalleryPage.prototype.processResults = function (results) {
            this.results = null;
            if (results == null) {
                return;
            }
            this.results = results;
            this.galleryItems = new Array(results.length);
            for (var i = 0; i < results.length; i++) {
                this.galleryItems[i] = new WEF.GalleryItem(results[i], i);
                this.galleryItems[i].displayAgave(this.gallery);
            }
            this.delayLoadVisibleImages();
        };
        WefGalleryPage.prototype.processAddinLoadingErrors = function (results) {
            for (var i = 0; i < results.length; i++) {
                if (results[i].hasLoadingError) {
                    this.showError(Strings.wefgallery.L_AddinsHasLoadingErrors, this.currentStoreId);
                    break;
                }
            }
        };
        WefGalleryPage.prototype.delayLoadVisibleImages = function (onLoadImagesComplete) {
            if (onLoadImagesComplete != null) {
                this.delayCallbacks.push(onLoadImagesComplete);
            }
            if (!this.delayTime || this.delaying == false || ((new Date().getTime() - this.delayTime) > 1000)) {
                this.delayTime = new Date().getTime();
                this.delaying = true;
                setTimeout(this.loadVisibleImages, this.delayLoad);
            }
            else {
                this.delayTime = new Date().getTime();
            }
        };
        WefGalleryPage.prototype.getItemsPerRow = function () {
            if (!this.gallery || !this.galleryItems || this.galleryItems.length == 0) {
                return false;
            }
            var itemEndOfLine = 0;
            if (WEF.WefGalleryHelper.getHTMLDir() != "ltr") {
                itemEndOfLine = this.gallery.offsetLeft + this.gallery.offsetWidth;
            }
            var itemsPerRow = 0;
            var defaultMargin = WEF.WefGalleryHelper.getDPIXScaledNumber(WEF.UI.DefaultLeftMargin);
            for (var i = 0; i < this.galleryItems.length; i++) {
                var item = this.galleryItems[i].galleryItem;
                if (item.offsetLeft == 0) {
                    itemsPerRow = 3;
                    break;
                }
                if (WEF.WefGalleryHelper.getHTMLDir() == "ltr") {
                    var left = Math.abs(item.offsetLeft - defaultMargin);
                    if (left >= itemEndOfLine) {
                        itemEndOfLine = left;
                        itemsPerRow++;
                    }
                    else {
                        break;
                    }
                }
                else {
                    var right = item.offsetLeft + item.offsetWidth + defaultMargin;
                    if ((right <= itemEndOfLine) || (Math.abs(right - itemEndOfLine) < 1)) {
                        itemEndOfLine = right;
                        itemsPerRow++;
                    }
                    else {
                        break;
                    }
                }
            }
            this.itemsPerRow = itemsPerRow;
            return itemsPerRow;
        };
        WefGalleryPage.prototype.showContentPage = function (pageUrl) {
            var _this = this;
            this.footer.style.visibility = 'hidden';
            this.documentAppsMsg.style.display = 'none';
            this.footer.style.height = WEF.WefGalleryHelper.getDPIYScaledNumber(WEF.UI.HiddenFooterHeight) + "px";
            this.setGalleryHeight();
            this.showHideRightMenuButtons(false, false);
            if (pageUrl && pageUrl != "") {
                this.gallery.style.overflowY = "hidden";
                var spinWheelDiv = WEF.WefGalleryHelper.addSpinWheel(this.gallery);
                var frame = document.createElement("iframe");
                frame.setAttribute("id", "OMEXSTORE");
                frame.setAttribute("width", "100%");
                frame.setAttribute("height", "100%");
                frame.setAttribute("frameBorder", "0");
                frame.setAttribute("scrolling", "no");
                frame.setAttribute("title", Strings.wefgallery.L_RecommendationTabIframeTitleTxt);
                var iframeOnLoad = function () {
                    if (spinWheelDiv) {
                        if (spinWheelDiv.parentNode == _this.gallery) {
                            _this.gallery.removeChild(spinWheelDiv);
                        }
                        spinWheelDiv = null;
                    }
                    _this.gallery.setAttribute("aria-busy", "false");
                    if (frame.contentWindow) {
                        frame.contentWindow.focus();
                    }
                    _this.onPageLoad();
                };
                WEF.WefGalleryHelper.addEventListener(frame, "load", iframeOnLoad);
                pageUrl += "#" + window.location.href;
                frame.setAttribute("src", pageUrl);
                this.gallery.appendChild(frame);
            }
            else {
                this.showError(Strings.wefgallery.L_NoFeaturedItemsError, WEF.PageStoreId.Recommendation);
            }
        };
        WefGalleryPage.prototype.removeGalleryItem = function (index) {
            if (this.galleryItems) {
                if (this.galleryItems.length == 1 && index == 0) {
                    this.showNoAppsError();
                }
                else if (this.galleryItems[index]) {
                    var moeDiv = this.galleryItems[index].galleryItem;
                    this.gallery.removeChild(moeDiv);
                    this.galleryItems.splice(index, 1);
                    this.results.splice(index, 1);
                    var len = this.galleryItems.length;
                    for (var i = index; i < len; i++) {
                        this.galleryItems[i].setIndex(i);
                    }
                    if (this.galleryItems.length >= 1) {
                        var indexToFocus = index;
                        if (index >= this.galleryItems.length) {
                            indexToFocus = 0;
                        }
                        this.selectGalleryItems(indexToFocus, true);
                    }
                }
                this.currentIndex = -1;
                this.deSelectBtnAction();
            }
        };
        WefGalleryPage.prototype.selectGalleryItems = function (index, forceSelected) {
            if (forceSelected === void 0) { forceSelected = false; }
            var result = this.results[index];
            var len = this.galleryItems ? this.galleryItems.length : 0;
            this.currentIndex = -1;
            for (var i = 0; i < len; i++) {
                var item = this.galleryItems[i];
                if (index == i) {
                    this.currentIndex = index;
                    if (WEF.WefGalleryHelper.hasClass(item.galleryItem, "selected")) {
                        if (forceSelected == false) {
                            WEF.WefGalleryHelper.removeClass(item.galleryItem, "selected");
                            item.galleryItem.removeAttribute("aria-selected");
                            item.galleryItem.setAttribute("tabIndex", "-1");
                            this.currentIndex = -1;
                            this.deSelectBtnAction();
                        }
                    }
                    else {
                        WEF.WefGalleryHelper.addClass(item.galleryItem, "selected");
                        WEF.WefGalleryHelper.setHtmlEncodedText(this.selectedDescriptionText, result.description);
                        this.selectedDescriptionText.setAttribute("title", result.description);
                        this.selectedItem.style.display = 'block';
                        this.footerLink.style.display = 'none';
                        if (this.currentStoreType != WEF.StoreTypeEnum.ThisDocument) {
                            WEF.WefGalleryHelper.removeClass(this.btnAction, 'disabled');
                            this.btnAction.removeAttribute('disabled');
                        }
                        item.galleryItem.setAttribute("tabIndex", "0");
                        item.galleryItem.focus();
                        item.galleryItem.setAttribute("aria-selected", "true");
                        this.setSelectedItemWidth();
                        if (item.appOptions) {
                            item.appOptions.showOptionsButton();
                        }
                        this.onItemSelect(item);
                    }
                }
                else {
                    this.unselectGalleryItems(item);
                }
            }
        };
        WefGalleryPage.prototype.unselectGalleryItems = function (item) {
            if (item && item.galleryItem) {
                WEF.WefGalleryHelper.removeClass(item.galleryItem, "selected");
                item.galleryItem.removeAttribute("aria-selected");
                item.galleryItem.setAttribute("tabIndex", "-1");
                if (item.appOptions && item.galleryItem.querySelector(":hover") == null) {
                    item.appOptions.hideOptionsButton();
                }
            }
        };
        WefGalleryPage.prototype.showNoAppsError = function () {
            this.gallery.innerHTML = "";
            if (this.currentStoreType === WEF.StoreTypeEnum.MarketPlace) {
                this.noAppsMessage.style.display = 'block';
                this.gallery.appendChild(this.noAppsMessage);
                this.officeStoreBtn.focus();
                this.footer.style.visibility = 'hidden';
                this.showHideRightMenuButtons(false, true);
            }
            else {
                this.showError(Strings.wefgallery.L_NoAgavePrompt, this.currentStoreId);
            }
        };
        WefGalleryPage.prototype.showErrorInternal = function (messageStr, linkedMessageStr, linkedCallback, showCloseButton) {
            var _this = this;
            if (this.uiState.Ready) {
                this.notification.style.display = 'block';
                if (linkedMessageStr && linkedCallback) {
                    var link = document.getElementById("linkId");
                    if (!link) {
                        this.errorMessage.innerHTML = messageStr + " <a id='linkId'>" + linkedMessageStr + "</a>";
                        link = document.getElementById("linkId");
                        link.setAttribute("tabIndex", "0");
                        link.setAttribute("role", "link");
                        link.onclick = function () {
                            linkedCallback();
                        };
                        WEF.WefGalleryHelper.addClass(link, "SignInATag");
                    }
                }
                else {
                    this.errorMessage.innerHTML = messageStr;
                }
                var notificationHeight = this.errorMessage.scrollHeight + WEF.UI.AdjustNotificationHeight;
                document.getElementById("Notification").style.height = notificationHeight + "px";
                var headerHeight = WEF.WefGalleryHelper.getDPIYScaledNumber(WEF.UI.DefaultHeaderHeight) + notificationHeight;
                document.getElementById("Header").style.height = headerHeight + "px";
                this.setGalleryHeight();
            }
            else {
                if (arguments.length >= 3 && linkedMessageStr && linkedCallback) {
                    this.uiState.ErrorLinkTextBeforeReady = linkedMessageStr;
                    this.uiState.ErrorLinkHandlerBeforeReady = linkedCallback;
                }
                this.uiState.StoreIdBeforeReady = this.currentStoreId;
                this.uiState.ErrorBeforeReady = messageStr;
            }
            if (showCloseButton) {
                this.notificationDismissImg.style.width = WEF.WefGalleryHelper.getDPIXScaledNumber(WEF.UI.DismissButtonSide) + "px";
                this.notificationDismissImg.style.height = WEF.WefGalleryHelper.getDPIYScaledNumber(WEF.UI.DismissButtonSide) + "px";
                this.notificationDismiss.style.display = "table-cell";
                this.notificationDismiss.onclick = function () {
                    _this.notification.style.display = 'none';
                    _this.notificationDismiss.style.display = "none";
                    _this.header.style.height = WEF.WefGalleryHelper.getDPIYScaledNumber(WEF.UI.DefaultHeaderHeight) + "px";
                    _this.setGalleryHeight();
                };
            }
            else {
                this.notificationDismiss.style.display = "none";
            }
        };
        WefGalleryPage.prototype.showError = function (messageStr, storeId, linkedMessageStr, linkedCallback, showCloseButton) {
            if ((storeId || storeId === "") && storeId != this.currentStoreId || !messageStr) {
                return;
            }
            if (this.gallery && this.gallery.firstChild && WEF.WefGalleryHelper.hasClass(this.gallery.firstChild, "SpinWheel")) {
                this.gallery.removeChild(this.gallery.firstChild);
            }
            this.gallery.setAttribute("aria-busy", "false");
            if (arguments.length < 4) {
                this.showErrorInternal(messageStr);
            }
            else {
                this.showErrorInternal(messageStr, linkedMessageStr, linkedCallback, showCloseButton);
            }
        };
        WefGalleryPage.prototype.gotoStore = function () {
            this.toggleTabSelection(this.storeTab, null);
        };
        WefGalleryPage.prototype.overrideButtonTooltip = function () {
        };
        WefGalleryPage.prototype.getPageUrl = function (pageType) {
            var pageUrl = this.clientFacadeCommon.getPageUrl(pageType);
            if (pageUrl == "" && pageType == WEF.PageTypeEnum.Recommendation) {
                this.showError(Strings.wefgallery.L_NoFeaturedItemsError, WEF.PageStoreId.Recommendation);
            }
            return pageUrl;
        };
        WefGalleryPage.prototype.insertSelectedItem = function () {
            if (this.allowInsertion() && this.galleryItems) {
                for (var i = 0; i < this.galleryItems.length; i++) {
                    var item = this.galleryItems[i];
                    if (WEF.WefGalleryHelper.hasClass(item.galleryItem, "selected")) {
                        this.insertItem(item);
                        break;
                    }
                }
            }
        };
        WefGalleryPage.prototype.allowInsertion = function () {
            return true;
        };
        WefGalleryPage.prototype.checkAndCreateOneDriveProviderTab = function (oneDriveTabs, oneDriveTabOrder, oneDriveTabName, oneDriveStoreId, oneDriveStoreType) {
        };
        WefGalleryPage.prototype.wefGalleryAppOnLoad = function () {
            var _this = this;
            this.galleryContainer = document.getElementById('GalleryContainer');
            this.galleryContainer.setAttribute("role", "tabpanel");
            this.mainPage = document.getElementById('MainPage');
            this.gallery = document.getElementById('InsertGallery');
            this.header = document.getElementById("Header");
            this.tabs = document.getElementById("Tabs");
            this.footer = document.getElementById('Footer');
            this.footerLink = document.getElementById('FooterLink');
            this.mainTitle = document.getElementById('MainTitle');
            this.selectedItem = document.getElementById('SelectedItem');
            this.selectedDescriptionText = document.getElementById('SelectedDescriptionText');
            this.selectedDescriptionReadMoreLink = document.getElementById('SelectedDescriptionReadMoreLink');
            this.permissionTextAndLink = document.getElementById('PermissionTextAndLink');
            this.permissionTextTR = document.getElementById('PermissionTextTR');
            this.readMoreATag = document.getElementById('ReadMoreLink');
            this.permissionATag = document.getElementById('PermissionLink');
            this.documentAppsMsg = document.getElementById('DocumentAppsMessageId');
            this.documentAppsMsgText = document.getElementById('DocumentAppsMessageText');
            this.btnAction = document.getElementById('BtnAction');
            this.btnCancel = document.getElementById('BtnCancel');
            this.btnTrustAll = document.getElementById('BtnTrustAll');
            this.btnDone = document.getElementById('BtnDone');
            this.notification = document.getElementById("Notification");
            this.notification.setAttribute("role", "alert");
            this.errorMessage = document.getElementById('ErrorMessage');
            this.errorMessage.setAttribute("role", "alert");
            this.notificationDismiss = document.getElementById('NotificationDismiss');
            this.notificationDismissImg = document.getElementById('DismissImg');
            this.menuRight = document.getElementById('MenuRight');
            this.noAppsMessage = document.getElementById('NoAppsMessage');
            this.noAppsMessage.setAttribute("role", "alert");
            this.noAppsMessageTitle = document.getElementById('NoAppsMessageTitle');
            this.noAppsMessageTitle.setAttribute("role", "alert");
            this.noAppsMessageText = document.getElementById('NoAppsMessageText');
            this.noAppsMessageText.setAttribute("role", "alert");
            this.officeStoreBtn = document.getElementById('BtnStore');
            this.officeStoreBtn.title = Strings.wefgallery.L_OfficeStore_Button_NoAddIns_Tooltip;
            this.manageATag = document.getElementById('ManageInner');
            this.uploadATag = document.getElementById('UploadMenuInner');
            this.uploadMenuDiv = document.getElementById('UploadMenu');
            this.manageMenuDiv = document.getElementById('Manage');
            this.refreshMenuDiv = document.getElementById('Refresh');
            this.menuRightSeparatorDiv = document.getElementById("MenuRightSeparator");
            var optionsDiv = document.getElementById('Options');
            this.storeStaticElementRealSize();
            WEF.WefGalleryHelper.dpiScaleHeight(this.header);
            WEF.WefGalleryHelper.dpiScaleMarginLeft(this.mainTitle);
            WEF.WefGalleryHelper.dpiScaleHeight(this.mainTitle);
            WEF.WefGalleryHelper.dpiScaleHeight(optionsDiv);
            WEF.WefGalleryHelper.dpiScaleMarginLeft(this.errorMessage);
            WEF.WefGalleryHelper.dpiScaleMarginLeft(this.tabs);
            WEF.WefGalleryHelper.dpiScaleMarginLeft(this.documentAppsMsgText);
            WEF.WefGalleryHelper.dpiScaleHeight(this.footer);
            WEF.WefGalleryHelper.dpiScaleWidth(this.footerLink);
            WEF.WefGalleryHelper.dpiScaleMarginLeft(this.footerLink);
            WEF.WefGalleryHelper.dpiScaleMarginLeft(this.selectedItem);
            WEF.WefGalleryHelper.dpiScaleHeightAndWidth(this.btnAction);
            WEF.WefGalleryHelper.dpiScaleHeightAndWidth(this.btnCancel);
            WEF.WefGalleryHelper.dpiScaleHeightAndWidth(this.btnTrustAll);
            WEF.WefGalleryHelper.dpiScaleHeightAndWidth(this.btnDone);
            WEF.WefGalleryHelper.dpiScaleHeight(this.notification);
            WEF.WefGalleryHelper.dpiScaleHeight(this.uploadMenuDiv);
            WEF.WefGalleryHelper.dpiScaleHeight(this.menuRight);
            WEF.WefGalleryHelper.dpiScaleHeight(this.manageMenuDiv);
            WEF.WefGalleryHelper.dpiScaleHeight(this.menuRightSeparatorDiv);
            WEF.WefGalleryHelper.dpiScaleHeight(this.refreshMenuDiv);
            this.menuRight.style.display = "none";
            this.gallery.onscroll = function () { _this.galleryScrollHandler(); };
            this.btnAction.onclick = function () { _this.insertSelectedItem(); };
            this.btnCancel.onclick = function () { _this.cancelDialog(); };
            this.btnDone.onclick = function () { _this.cancelDialog(); };
            this.officeStoreBtn.onclick = function () { _this.gotoStore(); };
            this.footerLink.onclick = function () { _this.gotoStore(); };
            this.manageATag.onclick = function () { _this.launchAppManagePage(); };
            this.showActionButtons(WEF.ActionButtonGroups.None);
            this.modalDialog = new WEF.AppManagement.ModalDialog(this.mainPage);
            this.menuHandler = new WEF.AppManagement.MenuHandler(this.galleryContainer, this.modalDialog);
            this.keyHandlers = [this.menuHandler, this.modalDialog];
            window.document.onkeydown = function (e) {
                _this.keyCodePressed = e.keyCode;
                _this.galleryKeyDownHandler(e);
            };
            window.document.onkeyup = function (e) {
                _this.keyCodePressed = e.keyCode;
                _this.galleryKeyUpHandler(e);
            };
            window.onresize = this.resizeHandler;
            this.uiState.Ready = true;
        };
        return WefGalleryPage;
    })();
    WEF.WefGalleryPage = WefGalleryPage;
    WEF.setupClientSpecificWefGalleryPage = null;
    WEF.showIt = function () {
        if (WEF.setupClientSpecificWefGalleryPage) {
            WEF.setupClientSpecificWefGalleryPage();
            WEF.IMPage.showItInternal();
        }
    };
})(WEF || (WEF = {}));
var WEF;
(function (WEF) {
    var AppManagement;
    (function (AppManagement) {
        var AppManagementAction;
        (function (AppManagementAction) {
            AppManagementAction[AppManagementAction["Cancel"] = 0] = "Cancel";
            AppManagementAction[AppManagementAction["AppDetails"] = 1] = "AppDetails";
            AppManagementAction[AppManagementAction["RateReview"] = 2] = "RateReview";
            AppManagementAction[AppManagementAction["Remove"] = 3] = "Remove";
        })(AppManagementAction || (AppManagementAction = {}));
        var AppManagementMenuFlags;
        (function (AppManagementMenuFlags) {
            AppManagementMenuFlags[AppManagementMenuFlags["ConfirmationDialogCancel"] = 256] = "ConfirmationDialogCancel";
            AppManagementMenuFlags[AppManagementMenuFlags["IsAnonymous"] = 1024] = "IsAnonymous";
        })(AppManagementMenuFlags || (AppManagementMenuFlags = {}));
        var MenuDirection;
        (function (MenuDirection) {
            MenuDirection[MenuDirection["Up"] = 0] = "Up";
            MenuDirection[MenuDirection["Down"] = 1] = "Down";
            MenuDirection[MenuDirection["Left"] = 2] = "Left";
            MenuDirection[MenuDirection["Right"] = 3] = "Right";
        })(MenuDirection || (MenuDirection = {}));
        var ModalDialog = (function () {
            function ModalDialog(modalDisabledDiv) {
                this.modalDisabledDiv = null;
                this.overlayDiv = null;
                this.dialogDiv = null;
                this.buttonDiv = null;
                this.confirmMessageDiv = null;
                this.buttonElements = [];
                this.enterKeyTarget = null;
                this.dialogId = "appManagementModalDialog";
                this.modalDisabledDiv = modalDisabledDiv;
                this.overlayDiv = document.createElement("div");
                WEF.WefGalleryHelper.addClass(this.overlayDiv, "Overlay");
                document.body.appendChild(this.overlayDiv);
                this.dialogDiv = document.createElement("div");
                this.dialogDiv.setAttribute("role", "dialog");
                WEF.WefGalleryHelper.addClass(this.dialogDiv, "ConfirmDialog");
                document.body.appendChild(this.dialogDiv);
                this.confirmMessageDiv = document.createElement("div");
                WEF.WefGalleryHelper.addClass(this.confirmMessageDiv, "ConfirmMessage");
                this.dialogDiv.appendChild(this.confirmMessageDiv);
                this.buttonDiv = document.createElement("div");
                WEF.WefGalleryHelper.addClass(this.buttonDiv, "ConfirmButtons");
                this.dialogDiv.appendChild(this.buttonDiv);
                this.dialogDiv.style.width = WEF.WefGalleryHelper.getDPIXScaledNumber(WEF.UI.ConfirmDialogWidth) + "px";
            }
            ModalDialog.prototype.handleKeyDown = function (ev) {
                if (!this.isDialogVisible()) {
                    return false;
                }
                var handled = false;
                switch (ev.keyCode) {
                    case 9:
                        this.onTab(ev);
                        handled = true;
                        break;
                    case 13:
                        handled = this.onEnterKeyDown(ev);
                        break;
                    case 27:
                        this.hideDialog();
                        handled = true;
                        break;
                }
                return handled;
            };
            ModalDialog.prototype.handleKeyUp = function (ev) {
                var handled = false;
                if (!this.isDialogVisible()) {
                    return handled;
                }
                switch (ev.keyCode) {
                    case 13:
                        var eventTarget = ev.srcElement ? ev.srcElement : ev.target;
                        if (eventTarget == this.enterKeyTarget) {
                            this.enterKeyTarget.click();
                            handled = true;
                        }
                        break;
                }
                return handled;
            };
            ModalDialog.prototype.hideDialog = function () {
                if (!this.isDialogVisible()) {
                    return;
                }
                var tabElements = this.getTabbableElements();
                var reFocused = false;
                for (var i = 0; i < tabElements.length; i++) {
                    var element = tabElements[i];
                    var previousTabValue = element.getAttribute("data-previous-tab");
                    var previousDisabledValue = element.getAttribute("data-previous-disabled");
                    if (previousTabValue !== null) {
                        element.setAttribute("tabindex", previousTabValue);
                    }
                    else {
                        element.removeAttribute("tabIndex");
                    }
                    if (previousDisabledValue !== null) {
                        element.disabled = (previousDisabledValue.toLowerCase() == "true");
                        if (!reFocused && !element.disabled) {
                            reFocused = true;
                            element.focus();
                        }
                    }
                }
                this.dialogDiv.style.display = "none";
                this.overlayDiv.style.display = "none";
                this.modalDisabledDiv.removeAttribute("aria-hidden");
            };
            ModalDialog.prototype.showDialog = function (message, buttonsCreationInfo) {
                if (!this.isDialogVisible()) {
                    var tabElements = this.getTabbableElements();
                    for (var i = 0; i < tabElements.length; i++) {
                        var element = tabElements[i];
                        element.setAttribute("data-previous-tab", element.getAttribute("tabindex"));
                        var disableStatusBackup = false;
                        if (element.disabled) {
                            disableStatusBackup = true;
                        }
                        element.setAttribute("data-previous-disabled", disableStatusBackup.toString());
                        element.setAttribute("tabindex", "-1");
                        element.disabled = true;
                    }
                    this.modalDisabledDiv.setAttribute("aria-hidden", "true");
                }
                this.dialogDiv.style.display = "block";
                this.overlayDiv.style.display = "block";
                this.confirmMessageDiv.innerHTML = message;
                this.buttonDiv.innerHTML = "";
                this.buttonElements = [];
                for (i = 0; i < buttonsCreationInfo.length; i++) {
                    var buttonInfo = buttonsCreationInfo[i];
                    var button = document.createElement("input");
                    button.setAttribute("id", buttonInfo.id);
                    button.setAttribute("type", "button");
                    button.setAttribute("value", buttonInfo.text);
                    button.setAttribute("title", buttonInfo.text);
                    button.setAttribute("data-buttonIndex", i.toString());
                    button.onclick = buttonInfo.onclick;
                    this.buttonDiv.appendChild(button);
                    this.buttonElements.push(button);
                    WEF.WefGalleryHelper.dpiScaleHeightAndWidth(button);
                    if (buttonInfo.hasFocus) {
                        button.focus();
                    }
                }
                this.positionDialog();
            };
            ModalDialog.prototype.positionDialog = function () {
                if (!this.isDialogVisible()) {
                    return;
                }
                var confirmDialog = this.dialogDiv;
                var top = WEF.WefGalleryHelper.getDocumentHeight() / 2 - confirmDialog.offsetHeight / 2;
                var left = WEF.WefGalleryHelper.getDocumentWidth() / 2 - confirmDialog.offsetWidth / 2;
                confirmDialog.style.top = top + "px";
                confirmDialog.style.left = left + "px";
            };
            ModalDialog.prototype.getTabbableElements = function () {
                return document.querySelectorAll("input,a,button,[tabindex]");
            };
            ModalDialog.prototype.isDialogVisible = function () {
                return this.dialogDiv.style.display != "none" && this.dialogDiv.offsetWidth > 0;
            };
            ModalDialog.prototype.onTab = function (ev) {
                var eventTarget = ev.srcElement ? ev.srcElement : ev.target;
                var buttonIndexAttribute = parseInt(eventTarget.getAttribute("data-buttonIndex"));
                var currentIndex = 0;
                if (isFinite(buttonIndexAttribute)) {
                    currentIndex = buttonIndexAttribute;
                }
                if (ev.shiftKey) {
                    if (currentIndex <= 0) {
                        this.buttonElements[this.buttonElements.length - 1].focus();
                    }
                    else {
                        this.buttonElements[currentIndex - 1].focus();
                    }
                }
                else {
                    if (currentIndex >= this.buttonElements.length - 1) {
                        this.buttonElements[0].focus();
                    }
                    else {
                        this.buttonElements[currentIndex + 1].focus();
                    }
                }
            };
            ModalDialog.prototype.onEnterKeyDown = function (ev) {
                var handled = false;
                var eventTarget = ev.srcElement ? ev.srcElement : ev.target;
                if (eventTarget == null) {
                    return handled;
                }
                for (var i = 0; i < this.buttonElements.length; i++) {
                    if (this.buttonElements[i] == eventTarget) {
                        this.enterKeyTarget = this.buttonElements[i];
                        handled = true;
                        break;
                    }
                }
                return handled;
            };
            return ModalDialog;
        })();
        AppManagement.ModalDialog = ModalDialog;
        var MenuHandler = (function () {
            function MenuHandler(containerDiv, removalConfirmationDialog) {
                var _this = this;
                this.menuDiv = null;
                this.appDetailsButton = null;
                this.rateReviewButton = null;
                this.removeAppButton = null;
                this.menuItems = null;
                this.currentMenuItemIndex = 0;
                this.currentResult = null;
                this.removalConfirmationDialog = null;
                this.enterKeyTarget = null;
                this.dialogId = "appManagementMenuDialog";
                this.menuDiv = document.createElement("ul");
                this.menuDiv.setAttribute("role", "menu");
                this.menuDiv.setAttribute("tabindex", "-1");
                this.menuDiv.setAttribute("id", "OptionsMenu");
                this.removalConfirmationDialog = removalConfirmationDialog;
                this.menuDiv.oncontextmenu = function () {
                    return false;
                };
                containerDiv.appendChild(this.menuDiv);
                this.appDetailsButton = new OptionsMenuItem(this.menuDiv, "AppDetails", Strings.wefgallery.L_OptionsMenu_AppDetails_Txt, Strings.wefgallery.L_OptionsMenu_AppDetails_Txt_Tooltip);
                this.rateReviewButton = new OptionsMenuItem(this.menuDiv, "RateReview", Strings.wefgallery.L_OptionsMenu_RateReview_Txt, Strings.wefgallery.L_OptionsMenu_RateReview_Txt_Tooltip);
                this.removeAppButton = new OptionsMenuItem(this.menuDiv, "RemoveApp", Strings.wefgallery.L_OptionsMenu_Remove_Txt, Strings.wefgallery.L_OptionsMenu_Remove_Txt_Tooltip);
                this.menuItems = [this.appDetailsButton, this.rateReviewButton, this.removeAppButton];
                var addFocusListener = function (index) {
                    WEF.WefGalleryHelper.addEventListener(_this.menuItems[index].element, "focus", function () {
                        _this.selectMenuItemAtIndex(index);
                    });
                };
                for (var i = 0; i < this.menuItems.length; i++) {
                    addFocusListener(i);
                }
                var clickInMenuCheck = function (event) {
                    if (_this.menuDiv.contains(event.target) == false) {
                        _this.hideMenu(true);
                    }
                };
                WEF.WefGalleryHelper.addEventListener(document, "click", clickInMenuCheck);
            }
            MenuHandler.prototype.createAppOptions = function (result) {
                return new AppOptions(result, this);
            };
            MenuHandler.prototype.handleKeyDown = function (ev) {
                if (this.isMenuVisible() == false) {
                    return false;
                }
                var handled = false;
                switch (ev.keyCode) {
                    case 13:
                        handled = this.onEnterKeyDown(ev);
                        break;
                    case 27:
                        this.hideMenu(true);
                        handled = true;
                        break;
                    case 38:
                        this.selectNextMenuItem(MenuDirection.Up);
                        handled = true;
                        break;
                    case 40:
                        this.selectNextMenuItem(MenuDirection.Down);
                        handled = true;
                        break;
                    case 37:
                    case 39:
                        handled = true;
                        break;
                    default:
                        handled = false;
                        break;
                }
                return handled;
            };
            MenuHandler.prototype.handleKeyUp = function (ev) {
                var handled = false;
                if (this.isMenuVisible() == false) {
                    return handled;
                }
                switch (ev.keyCode) {
                    case 13:
                        var eventTarget = ev.srcElement ? ev.srcElement : ev.target;
                        if (eventTarget == this.enterKeyTarget) {
                            this.enterKeyTarget.click();
                            handled = true;
                        }
                        break;
                }
                return handled;
            };
            MenuHandler.prototype.hideMenu = function (logData) {
                if (this.isMenuVisible()) {
                    this.menuDiv.style.display = "none";
                    if (logData) {
                        this.logData(this.currentResult, AppManagementAction.Cancel, 0);
                    }
                    if (this.removalConfirmationDialog != null) {
                        var tabElements = this.removalConfirmationDialog.getTabbableElements();
                        for (var i = 0; i < tabElements.length; i++) {
                            var element = tabElements[i];
                            if (!element.disabled) {
                                element.focus();
                                break;
                            }
                        }
                    }
                }
            };
            MenuHandler.prototype.isMenuVisible = function () {
                return this.menuDiv.style.display != "none" && this.menuDiv.offsetWidth > 0;
            };
            MenuHandler.prototype.popupMenuForApp = function (result, optionsButton, appIndex, tnDiv, img) {
                var _this = this;
                this.currentResult = result;
                this.appDetailsButton.setOnClick(function () {
                    _this.hideMenu(false);
                    WEF.IMPage.invokeWindowOpen(_this.currentResult.appEndNodeUrl);
                    _this.logData(result, AppManagementAction.AppDetails, 0);
                });
                this.rateReviewButton.setOnClick(function () {
                    _this.hideMenu(false);
                    WEF.IMPage.invokeWindowOpen(result.rateReviewUrl);
                    _this.logData(result, AppManagementAction.RateReview, 0);
                });
                this.removeAppButton.setOnClick(function () {
                    _this.hideMenu(false);
                    _this.showRemoveConfirmationDialog(result.authType, function () {
                        _this.removeAppHandler(result, appIndex, tnDiv, img);
                    }, function () {
                        _this.logData(result, AppManagementAction.Remove | AppManagementMenuFlags.ConfirmationDialogCancel, 0);
                    });
                });
                setTimeout(function () {
                    WEF.IMPage.selectGalleryItems(appIndex, true);
                    _this.positionMenu(optionsButton);
                    _this.clearMenuSelection();
                    _this.menuDiv.focus();
                }, 0);
            };
            MenuHandler.prototype.onEnterKeyDown = function (ev) {
                var handled = false;
                var eventTarget = ev.srcElement ? ev.srcElement : ev.target;
                if (eventTarget == null) {
                    return handled;
                }
                for (var i = 0; i < this.menuItems.length; i++) {
                    if (this.menuItems[i].element == eventTarget) {
                        this.enterKeyTarget = this.menuItems[i].element;
                        handled = true;
                        break;
                    }
                }
                return handled;
            };
            MenuHandler.prototype.positionMenu = function (optionsButton) {
                this.menuDiv.style.display = "block";
                var insertDialogHeight = WEF.WefGalleryHelper.getDocumentHeight();
                var insertDialogWidth = WEF.WefGalleryHelper.getDocumentWidth();
                var menuRect = this.menuDiv.getBoundingClientRect();
                var optionButtonRect = optionsButton.getBoundingClientRect();
                var menuHeight = this.menuDiv.offsetHeight;
                var menuWidth = this.menuDiv.offsetWidth;
                var offsetTop = optionsButton.offsetHeight;
                var calculatedZIndex = 1;
                var parentZIndex = parseInt(this.menuDiv.parentElement.style.zIndex);
                if (isFinite(parentZIndex)) {
                    calculatedZIndex = parentZIndex + 1;
                }
                this.menuDiv.style.zIndex = calculatedZIndex.toString();
                if (optionButtonRect.top + menuHeight <= insertDialogHeight) {
                    this.menuDiv.style.top = (optionButtonRect.top) + "px";
                }
                else {
                    this.menuDiv.style.top = (optionButtonRect.top + offsetTop - menuHeight) + "px";
                }
                if (WEF.WefGalleryHelper.getHTMLDir() == "ltr") {
                    if (optionButtonRect.left + menuWidth <= insertDialogWidth) {
                        this.menuDiv.style.left = (optionButtonRect.left) + "px";
                    }
                    else {
                        this.menuDiv.style.left = (optionButtonRect.right - menuWidth) + "px";
                    }
                }
                else {
                    if (optionButtonRect.left - menuWidth > 0) {
                        this.menuDiv.style.left = (optionButtonRect.right - menuWidth) + "px";
                    }
                    else {
                        this.menuDiv.style.left = (optionButtonRect.left) + "px";
                    }
                }
            };
            MenuHandler.prototype.removeAppHandler = function (result, appIndex, tnDiv, img) {
                var _this = this;
                var onRemoveComplete = function (status) {
                    _this.removeAppButton.setDisabled(false);
                    WEF.WefGalleryHelper.removeClass(tnDiv, "SpinWheel");
                    img.style.display = "block";
                    var errorMessage = "";
                    switch (status) {
                        case WEF.InvokeResultCode.S_OK:
                            WEF.IMPage.removeGalleryItem(appIndex);
                            break;
                        case WEF.InvokeResultCode.E_CATALOG_REQUEST_FAILED:
                            errorMessage = Strings.wefgallery.L_RequestFailedError;
                            break;
                        case WEF.InvokeResultCode.E_OEM_NO_NETWORK_CONNECTION:
                            errorMessage = Strings.wefgallery.L_RemoveAppOfflineError;
                            break;
                        default:
                            errorMessage = Strings.wefgallery.L_RemoveAppGeneralError.replace("{0}", result.displayName);
                            break;
                    }
                    if (errorMessage) {
                        WEF.IMPage.showError(errorMessage, WEF.StoreTypeEnum.MarketPlace, null, null, true);
                    }
                };
                this.removeAppButton.setDisabled(true);
                WEF.WefGalleryHelper.addClass(tnDiv, "SpinWheel");
                img.style.display = "none";
                WEF.IMPage.removeAgave(this.currentResult, onRemoveComplete);
            };
            MenuHandler.prototype.showRemoveConfirmationDialog = function (authType, removeAppHandler, cancelHandler) {
                var _this = this;
                var message = Strings.wefgallery.L_Confirmation_RemoveAppAuthenticated_Message;
                if (authType == WEF.AuthType.Anonymous) {
                    message = Strings.wefgallery.L_Confirmation_RemoveAppAnonymous_Message;
                }
                message = message.replace(/\\n/g, "<br />");
                var buttons = [];
                buttons.push({
                    id: "ConfirmRemove",
                    text: Strings.wefgallery.L_OptionsMenu_Remove_Txt,
                    hasFocus: true,
                    onclick: function () {
                        _this.removalConfirmationDialog.hideDialog();
                        removeAppHandler();
                    }
                });
                buttons.push({
                    id: "ConfirmCancel",
                    text: Strings.wefgallery.L_Confirmation_Cancel_Button_Txt,
                    hasFocus: false,
                    onclick: function () {
                        _this.removalConfirmationDialog.hideDialog();
                        cancelHandler();
                    }
                });
                this.removalConfirmationDialog.showDialog(message, buttons);
            };
            MenuHandler.prototype.selectMenuItemAtIndex = function (index) {
                if (this.menuItems[index] && this.menuItems[index].disabled == false) {
                    if (this.currentMenuItemIndex >= 0) {
                        this.menuItems[this.currentMenuItemIndex].setSelected(false);
                    }
                    this.menuItems[index].setSelected(true);
                    this.currentMenuItemIndex = index;
                    return true;
                }
                return false;
            };
            MenuHandler.prototype.clearMenuSelection = function () {
                if (this.currentMenuItemIndex >= 0) {
                    this.menuItems[this.currentMenuItemIndex].setSelected(false);
                    this.currentMenuItemIndex = -1;
                }
            };
            MenuHandler.prototype.selectNextMenuItem = function (direction) {
                if (this.currentMenuItemIndex < 0 && this.selectMenuItemAtIndex(0)) {
                    return;
                }
                else if (this.currentMenuItemIndex >= this.menuItems.length && this.selectMenuItemAtIndex(this.menuItems.length - 1)) {
                    return;
                }
                var i = this.currentMenuItemIndex;
                while (i >= 0 && i < this.menuItems.length) {
                    if (direction == MenuDirection.Up) {
                        i--;
                    }
                    else {
                        i++;
                    }
                    if (this.selectMenuItemAtIndex(i)) {
                        return;
                    }
                }
            };
            MenuHandler.prototype.logData = function (result, operationInfo, hresult) {
                var assetId = "0";
                if (result) {
                    assetId = result.id;
                }
                var maskIsAnonymous = 0;
                if (result.authType == WEF.AuthType.Anonymous) {
                    maskIsAnonymous = AppManagementMenuFlags.IsAnonymous;
                }
                WEF.IMPage.clientFacadeCommon.logAppManagementAction(assetId, operationInfo | maskIsAnonymous, hresult);
            };
            return MenuHandler;
        })();
        AppManagement.MenuHandler = MenuHandler;
        var OptionsMenuItem = (function () {
            function OptionsMenuItem(menuDiv, id, text, title) {
                this.disabled = false;
                this.element = null;
                var li = document.createElement("li");
                li.setAttribute("role", "presentation");
                this.element = document.createElement("button");
                WEF.WefGalleryHelper.setHtmlEncodedText(this.element, text);
                this.element.setAttribute("title", title);
                this.element.setAttribute("tabindex", "0");
                this.element.setAttribute("role", "menuitem");
                this.element.setAttribute("id", id);
                WEF.WefGalleryHelper.addClass(this.element, "menuOption");
                menuDiv.appendChild(li);
                li.appendChild(this.element);
                this.setSelected(false);
                this.setDisabled(false);
            }
            OptionsMenuItem.prototype.setOnClick = function (onClickHandler) {
                var _this = this;
                this.element.onclick = function () {
                    if (_this.disabled) {
                        return;
                    }
                    onClickHandler();
                };
            };
            OptionsMenuItem.prototype.setSelected = function (selected) {
                this.element.setAttribute("aria-selected", selected.toString());
                if (selected) {
                    this.element.focus();
                }
            };
            OptionsMenuItem.prototype.setDisabled = function (disabled) {
                this.element.setAttribute("aria-disabled", disabled.toString());
                this.element.disabled = disabled;
                this.disabled = disabled;
                if (disabled) {
                    WEF.WefGalleryHelper.addClass(this.element, "disabled");
                }
                else {
                    WEF.WefGalleryHelper.removeClass(this.element, "disabled");
                }
            };
            return OptionsMenuItem;
        })();
        AppManagement.OptionsMenuItem = OptionsMenuItem;
        var AppOptions = (function () {
            function AppOptions(result, menuHandler) {
                this.result = null;
                this.appIndex = null;
                this.menuHandler = null;
                this.tnDiv = null;
                this.img = null;
                this.optionsButton = null;
                this.result = result;
                this.menuHandler = menuHandler;
            }
            AppOptions.prototype.createOptionsButton = function (appIndex, tnDiv, img) {
                var _this = this;
                var optionsButton = null;
                if (WEF.IMPage.currentStoreType === WEF.StoreTypeEnum.MarketPlace && WEF.IMPage.canShowAppManagementMenu()) {
                    optionsButton = document.createElement("input");
                    WEF.WefGalleryHelper.addClass(optionsButton, "OptionsButton");
                    optionsButton.setAttribute("type", "button");
                    optionsButton.setAttribute("role", "button");
                    optionsButton.setAttribute("value", "\u22EF");
                    optionsButton.setAttribute("aria-label", Strings.wefgallery.L_OptionsMenu_Tooltip);
                    optionsButton.setAttribute("title", Strings.wefgallery.L_OptionsMenu_Tooltip);
                    optionsButton.setAttribute("tabindex", "0");
                    optionsButton.style.width = WEF.WefGalleryHelper.getDPIXScaledNumber(WEF.UI.MenuButtonSide) + "px";
                    optionsButton.style.height = WEF.WefGalleryHelper.getDPIYScaledNumber(WEF.UI.MenuButtonSide) + "px";
                    optionsButton.style.backgroundSize = WEF.WefGalleryHelper.getDPIYScaledNumber(WEF.UI.MenuButtonBackgroundSize) + "px";
                    this.optionsButton = optionsButton;
                    WEF.WefGalleryHelper.addEventListener(optionsButton, "click", function () {
                        _this.popupMenu();
                    });
                    this.appIndex = appIndex;
                    this.tnDiv = tnDiv;
                    this.img = img;
                }
                return optionsButton;
            };
            AppOptions.prototype.showOptionsButton = function () {
                if (this.optionsButton) {
                    this.optionsButton.style.display = "block";
                }
            };
            AppOptions.prototype.hideOptionsButton = function () {
                if (this.optionsButton) {
                    this.optionsButton.style.display = "none";
                }
            };
            AppOptions.prototype.popupMenu = function () {
                if (this.optionsButton) {
                    this.menuHandler.popupMenuForApp(this.result, this.optionsButton, this.appIndex, this.tnDiv, this.img);
                }
            };
            AppOptions.prototype.setAppIndex = function (appIndex) {
                if (appIndex >= 0) {
                    this.appIndex = appIndex;
                }
            };
            return AppOptions;
        })();
        AppManagement.AppOptions = AppOptions;
    })(AppManagement = WEF.AppManagement || (WEF.AppManagement = {}));
})(WEF || (WEF = {}));
WEF.AGAVE_DEFAULT_ICON = "AgaveDefaultIcon.png";
var WEF;
(function (WEF) {
    var ClientFacade_Native = (function () {
        function ClientFacade_Native() {
            var _this = this;
            this.onShowEntitlementsComplete = null;
            this.onRemoveAgaveCallback = null;
            this.envSetting = {};
            this.onGetEntitlementsInternal = function (results, hres) {
            };
            this.onGetEntitlements = function (results, hres) {
                if (_this.storeId != WEF.IMPage.currentStoreId) {
                    return;
                }
                WEF.IMPage.cleanUpGallery();
                WEF.IMPage.uiState.ErrorBeforeReady = "";
                WEF.IMPage.providers[_this.storeId][1] = 0;
                WEF.IMPage.providers[_this.storeId][2] = 0;
                if (WEF.WefGalleryHelper.handleErrorCode(hres, _this.storeId, null, null)) {
                    return;
                }
                var etsArray = results.toArray ? results.toArray() : results;
                var entitlements = new Array();
                var existingId = {};
                for (var i = 0; i < etsArray.length; i++) {
                    var etArray = etsArray[i].toArray ? etsArray[i].toArray() : etsArray[i];
                    var galleryItem = new WEF.AgaveInfo();
                    galleryItem.displayName = etArray[0];
                    galleryItem.id = etArray[1];
                    galleryItem.description = etArray[2];
                    galleryItem.targetType = etArray[3];
                    galleryItem.appVersion = etArray.length > 4 ? etArray[4] : "";
                    galleryItem.assetId = etArray.length > 5 ? etArray[5] : "";
                    galleryItem.assetStoreId = etArray.length > 6 ? etArray[6] : "";
                    galleryItem.width = etArray.length > 7 ? etArray[7] : 0;
                    galleryItem.height = etArray.length > 8 ? etArray[8] : 0;
                    galleryItem.iconUrl = etArray.length > 9 ? etArray[9] : WEF.AGAVE_DEFAULT_ICON;
                    galleryItem.providerName = etArray.length > 10 ? etArray[10] : "";
                    galleryItem.storeId = etArray.length > 11 ? etArray[11] : "";
                    galleryItem.appEndNodeUrl = etArray.length > 12 && etArray[12] !== "" ? etArray[12] : WEF.IMPage.getLandingPageUrl();
                    galleryItem.rateReviewUrl = etArray.length > 13 && etArray[13] !== "" ? etArray[13] : null;
                    galleryItem.authType = etArray.length > 14 && etArray[14] !== "" ? etArray[14] : null;
                    galleryItem.isAppCommandAddin = etArray.length > 15 && etArray[15] !== "" ? etArray[15] : null;
                    galleryItem.hasLoadingError = etArray.length > 16 && etArray[16] !== "" ? etArray[16] : null;
                    if (existingId[galleryItem.id] == null) {
                        existingId[galleryItem.id] = true;
                        entitlements.push(galleryItem);
                    }
                }
                entitlements.sort(WEF.AgaveInfo.cmpDisplayName);
                if (entitlements.length == 0) {
                    WEF.IMPage.showNoAppsError();
                    return;
                }
                if (WEF.IMPage.footer.style.visibility === 'hidden') {
                    WEF.IMPage.showFooter();
                    WEF.IMPage.showHideRightMenuButtons(true, true);
                    WEF.IMPage.setGalleryHeight();
                }
                WEF.IMPage.processResults(entitlements);
                WEF.IMPage.processAddinLoadingErrors(entitlements);
                if (_this.onShowEntitlementsComplete) {
                    _this.onShowEntitlementsComplete();
                }
                WEF.IMPage.onPageLoad();
            };
            this.onRemoveAgave = function (result, hres) {
                var ret = result.toArray();
                var status = ret[0];
                _this.onRemoveAgaveCallback(status);
            };
        }
        ClientFacade_Native.prototype.getEnvSetting = function () {
            return this.envSetting;
        };
        ClientFacade_Native.prototype.launchAppManagePage = function () {
        };
        ClientFacade_Native.prototype.onGetProviders = function (results, hres) {
            var refreshRequired = WEF.WefGalleryHelper.retrieveRefreshRequired();
            var providers = results.toArray();
            if (!providers || hres < 0 || providers.length === 0) {
                WEF.IMPage.cleanUpGallery();
                WEF.IMPage.showError(Strings.wefgallery.L_NoProviderError);
                return;
            }
            providers.sort(function (a, b) { return (a.toArray()[1] - b.toArray()[1]); });
            if (!WEF.IMPage.initializeGalleryUI(providers, false)) {
                WEF.IMPage.cleanUpGallery();
                WEF.IMPage.showError(Strings.wefgallery.L_NoProviderError);
                return;
            }
            WEF.IMPage.showContent(refreshRequired);
        };
        ClientFacade_Native.prototype.onGetProvidersShowContent = function (results, hres) {
            var providers = results.toArray();
            if (!providers || hres < 0 || providers.length === 0) {
                WEF.IMPage.cleanUpGallery();
                WEF.IMPage.showError(Strings.wefgallery.L_NoProviderError);
                return;
            }
            var len = providers.length, i, tempStoreId, tempStoreType, tempStatus, tempHResult;
            for (i = 0; i < len; i++) {
                var provider = providers[i].toArray();
                tempStoreId = provider[0];
                tempStoreType = provider[1];
                tempStatus = provider[2];
                tempHResult = provider[3];
                if (tempStoreType === WEF.StoreTypeEnum.MarketPlace) {
                    if (tempHResult == WEF.InvokeResultCode.E_USER_NOT_SIGNED_IN) {
                        WEF.IMPage.showError(Strings.wefgallery.L_SignInPromptLiveId, WEF.IMPage.currentStoreId, Strings.wefgallery.L_SignInLinkText, WEF.IMPage.invokeSignIn);
                        return;
                    }
                }
            }
            WEF.storeTypes = {
                0: Strings.wefgallery.L_MarketPlaceTabTxt,
                1: Strings.wefgallery.L_CatalogTabTxt,
                3: Strings.wefgallery.L_ExchangeTabTxt,
                4: Strings.wefgallery.L_FileShareTabTxt,
                6: Strings.wefgallery.L_RecommendationTabTxt,
                8: Strings.wefgallery.L_ThisDocumentTabTxt
            };
            providers.sort(function (a, b) { return (a.toArray()[1] - b.toArray()[1]); });
            if (WEF.IMPage.initializeGalleryUI(providers, true)) {
                WEF.IMPage.showContent(false);
            }
        };
        ClientFacade_Native.prototype.setShowEntitlementsComplete = function (onShowEntitlementsComplete) {
            this.onShowEntitlementsComplete = onShowEntitlementsComplete;
        };
        ClientFacade_Native.prototype.setStoreId = function (storeId) {
            this.storeId = storeId;
        };
        ClientFacade_Native.prototype.setOnRemoveAgaveCallback = function (callback) {
            this.onRemoveAgaveCallback = callback;
        };
        return ClientFacade_Native;
    })();
    WEF.ClientFacade_Native = ClientFacade_Native;
    var WefGalleryPage_Native = (function (_super) {
        __extends(WefGalleryPage_Native, _super);
        function WefGalleryPage_Native(clientFacade) {
            var _this = this;
            _super.call(this, clientFacade);
            this.clientFacade = null;
            this.insertItem = function (item) {
                if (_this.allowInsertion()) {
                    _this.clientFacade.insertAgave(item, _this.currentStoreType);
                }
            };
            this.showEntitlements = function (storeId, refresh, onShowEntitlementsComplete) {
                if (_this.currentStoreType === WEF.StoreTypeEnum.MarketPlace) {
                    _this.showHideRightMenuButtons(_this.footer.style.visibility != 'hidden', true);
                    _this.showActionButtons(WEF.ActionButtonGroups.InsertCancel);
                    _this.documentAppsMsg.style.display = 'none';
                }
                else if (_this.currentStoreType === WEF.StoreTypeEnum.ThisDocument) {
                    _this.showHideRightMenuButtons(false, false);
                    _this.showActionButtons(WEF.ActionButtonGroups.ThisDocument);
                    _this.documentAppsMsg.style.display = 'inline';
                }
                else {
                    _this.showHideRightMenuButtons(false, true);
                    _this.showActionButtons(WEF.ActionButtonGroups.InsertCancel);
                    _this.documentAppsMsg.style.display = 'none';
                }
                _this.hideButtons();
                if (WEF.WefGalleryHelper.handleErrorCode(_this.getCurrentProviderHResult(), _this.currentStoreId, _this.currentStoreType, _this.getCurrentProviderStatus())) {
                    if (!refresh) {
                        return;
                    }
                }
                _this.gallery.style.overflowY = "auto";
                var spinWheelDiv = WEF.WefGalleryHelper.addSpinWheel(_this.gallery);
                if (storeId != undefined) {
                    var tempStoreId = storeId;
                    _this.clientFacade.setStoreId(storeId);
                    _this.clientFacade.setShowEntitlementsComplete(onShowEntitlementsComplete);
                    setTimeout(function () {
                        _this.clientFacade.getEntitlements(storeId, refresh, _this.clientFacade.onGetEntitlementsInternal);
                    }, 0);
                    if (refresh) {
                        WEF.WefGalleryHelper.saveRefreshRequired(false);
                    }
                }
            };
            this.postMessageListener = function (e) {
                if (e.data == "REFRESH_REQUIRED") {
                    WEF.WefGalleryHelper.saveRefreshRequired(true);
                }
            };
            this.clientFacade = clientFacade;
        }
        WefGalleryPage_Native.prototype.allowInsertion = function () {
            return this.currentStoreType != WEF.StoreTypeEnum.ThisDocument;
        };
        WefGalleryPage_Native.prototype.retrieveStoreId = function () {
            var initTab = 0;
            try {
                initTab = this.clientFacade.getInitTab();
            }
            catch (ex) {
                return WEF.WefGalleryHelper.retrieveStoreIdfromStorage();
            }
            if (initTab == 0) {
                return this.omexStoreId;
            }
            else if (initTab == 1) {
                return WEF.PageStoreId.Recommendation;
            }
            else {
                return WEF.WefGalleryHelper.retrieveStoreIdfromStorage();
            }
        };
        WefGalleryPage_Native.prototype.launchAppManagePage = function () {
            this.clientFacade.launchAppManagePage();
        };
        WefGalleryPage_Native.prototype.removeAgave = function (result, callback) {
            this.clientFacade.removeAgave(result, this.currentStoreType, callback);
        };
        WefGalleryPage_Native.prototype.hideButtons = function () {
        };
        WefGalleryPage_Native.prototype.showItInternal = function () {
            this.wefGalleryAppOnLoad();
            WEF.WefGalleryHelper.addEventListener(window, "message", this.postMessageListener);
            WEF.WefGalleryHelper.addSpinWheel(this.gallery);
            this.setGalleryHeight();
            try {
                this.clientFacade.runShowIt();
            }
            catch (ex) {
                this.showError(Strings.wefgallery.L_GetEntitilementsGeneralError);
            }
            this.gallery.setAttribute("aria-busy", "false");
        };
        return WefGalleryPage_Native;
    })(WEF.WefGalleryPage);
    WEF.WefGalleryPage_Native = WefGalleryPage_Native;
})(WEF || (WEF = {}));
var WEF;
(function (WEF) {
    var InvokeType = {
        "GetProviders": 620,
        "Insert": 621,
        "GetEntitlements": 622,
        "GetLandingPageUrl": 623,
        "MountOrSignInLiveId": 624,
        "TrustAllInDocOmexApps": 625,
        "GetInitTab": 626,
        "RemoveAgave": 627,
        "LogAppManagementAction": 628,
        "GetResourceStrings": 629
    };
    var MessageHandlerName = "WefGalleryHandler";
    var ClientFacade_Ios = (function (_super) {
        __extends(ClientFacade_Ios, _super);
        function ClientFacade_Ios() {
            var _this = this;
            _super.apply(this, arguments);
            this.omexPageUrls = [];
            this.onGetEntitlementsInternal = function (results, hres) {
                hres = results[0];
                results = results[1];
                results = new OSFWebkit.WebkitSafeArray(results);
                _this.onGetEntitlements(results, hres);
            };
        }
        ClientFacade_Ios.prototype.onGetLandingPageUrl = function (results, hres) {
            results = new OSFWebkit.WebkitSafeArray(results);
            var initTabPageUrls = results.toArray();
            if (!initTabPageUrls || (hres < 0) || initTabPageUrls.length === 0) {
                WEF.IMPage.cleanUpGallery();
                WEF.IMPage.showError(Strings.wefgallery.L_NoProviderError);
                return;
            }
            var initTabId = initTabPageUrls[0];
            this.setInitTab(initTabId);
            var pageUrls = initTabPageUrls[1];
            this.setPageUrl(pageUrls);
            try {
                OSF.ScriptMessaging.GetScriptMessenger().invokeMethod(MessageHandlerName, InvokeType.GetProviders, {}, this.onGetProviders);
            }
            catch (ex) {
                WEF.IMPage.showError(Strings.wefgallery.L_GetEntitilementsGeneralError);
            }
        };
        ClientFacade_Ios.prototype.onGetProviders = function (results, hres) {
            hres = results[0];
            results = results[1];
            _super.prototype.onGetProviders.call(this, new OSFWebkit.WebkitSafeArray(results), hres);
        };
        ClientFacade_Ios.prototype.onGetProvidersShowContent = function (results, hres) {
            _super.prototype.onGetProvidersShowContent.call(this, new OSFWebkit.WebkitSafeArray(results), hres);
        };
        ClientFacade_Ios.prototype.getPageUrl = function (pageType) {
            return this.omexPageUrls[pageType];
        };
        ClientFacade_Ios.prototype.setPageUrl = function (pageUrls) {
            this.omexPageUrls = pageUrls;
        };
        ClientFacade_Ios.prototype.getInitTab = function () {
            return this.initTab;
        };
        ClientFacade_Ios.prototype.setInitTab = function (initTabId) {
            this.initTab = initTabId;
        };
        ClientFacade_Ios.prototype.runShowIt = function () {
            OSF.ScriptMessaging.GetScriptMessenger().invokeMethod(MessageHandlerName, InvokeType.GetLandingPageUrl, {}, this.onGetLandingPageUrl);
        };
        ClientFacade_Ios.prototype.getEntitlements = function (storeId, refresh, onGetEntitlements) {
            var params = { "StoreId": storeId, "refresh": refresh };
            OSF.ScriptMessaging.GetScriptMessenger().invokeMethod(MessageHandlerName, InvokeType.GetEntitlements, params, onGetEntitlements);
        };
        ClientFacade_Ios.prototype.insertAgave = function (item, currentStoreType) {
            var params = {
                "AssetId": item.result.id,
                "Target": item.result.targetType,
                "Version": item.result.appVersion,
                "OmexStore": currentStoreType,
                "StoreId": item.result.storeId,
                "AssetIdA": currentStoreType == WEF.StoreTypeEnum.MarketPlace ? item.result.id : item.result.assetId,
                "AssetIdB": currentStoreType == WEF.StoreTypeEnum.MarketPlace ? item.result.id : item.result.assetStoreId,
                "Width": item.result.width,
                "Height": item.result.height
            };
            OSF.ScriptMessaging.GetScriptMessenger().invokeMethod(MessageHandlerName, InvokeType.Insert, params, '');
        };
        ClientFacade_Ios.prototype.removeResultsDimensionInfo = function (results) {
            if (OSF.OUtil.isArray(results)) {
                results.shift();
            }
        };
        ClientFacade_Ios.prototype.removeAgave = function (result, currentStoreType, callback) {
            this.setOnRemoveAgaveCallback(callback);
            var params = {
                "AssetId": result.id, "Tartget": result.targetType, "Version": result.appVersion, "OmexStore": currentStoreType,
                "StoreId": result.storeId, "AssetIdA": result.assetId, "AssetIdB": result.assetStoreId
            };
            OSF.ScriptMessaging.GetScriptMessenger().invokeMethod(MessageHandlerName, InvokeType.RemoveAgave, params, this.onRemoveAgave);
        };
        ClientFacade_Ios.prototype.logAppManagementAction = function (assetId, operationInfo, hresult) {
            var params = { "AssetId": assetId, "operationInfo": operationInfo, "hresult": hresult };
            OSF.ScriptMessaging.GetScriptMessenger().invokeMethod(MessageHandlerName, InvokeType.LogAppManagementAction, params, '');
        };
        return ClientFacade_Ios;
    })(WEF.ClientFacade_Native);
    var WefGallertPage_Ios = (function (_super) {
        __extends(WefGallertPage_Ios, _super);
        function WefGallertPage_Ios() {
            _super.apply(this, arguments);
            this.invokeSignIn = function () {
                throw "Shouldn't call into invokeSignIn on iOS platform.";
            };
        }
        WefGallertPage_Ios.prototype.onItemSelect = function (item) {
            this.insertSelectedItem();
        };
        WefGallertPage_Ios.prototype.trustAllAgaves = function () {
            OSF.ScriptMessaging.GetScriptMessenger().invokeMethod(MessageHandlerName, InvokeType.TrustAllInDocOmexApps, {}, '');
        };
        WefGallertPage_Ios.prototype.canShowAppManagementMenu = function () {
            return false;
        };
        WefGallertPage_Ios.prototype.invokeWindowOpen = function (pageUrl) {
            throw "Shouldn't call into invokeWindowOpen on iOS platform.";
        };
        WefGallertPage_Ios.prototype.cancelDialog = function () {
            throw "Shouldn't call into cancelDialog on iOS platform.";
        };
        WefGallertPage_Ios.prototype.onPageLoad = function () {
        };
        return WefGallertPage_Ios;
    })(WEF.WefGalleryPage_Native);
    WEF.setupClientSpecificWefGalleryPage = function () {
        var clientFacade = new ClientFacade_Ios();
        WEF.IMPage = new WefGallertPage_Ios(clientFacade);
    };
})(WEF || (WEF = {}));
