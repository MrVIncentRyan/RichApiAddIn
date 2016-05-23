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
OSF.OUtil.setNamespace("OSF", window);
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
OSF.OmexAnonymousProvider = (function () {
    var _hasBeenInitialized = false;
    var _clientID = '9EA00E01-D076-4F05-88A9-A8738244BC24';
    var _cacheKeyPrefix = '__OSF_ANONYMOUS_OMEX.';
    var _manifestRefreshRate = 5 * 365;
    var _cacheManager = new OfficeExt.AppsDataCacheManager(OSF.OUtil.getLocalStorage(), new OfficeExt.SafeSerializer());
    var _appInstallInfoWS = {
        url: "/appinstall/unauthenticated?cmu={0}&assetid={1}&ret=0",
        cacheKey: _cacheKeyPrefix + 'appInstallInfo.{0}.{1}'
    };
    var _killedAppsWS = {
        url: "/appinfo/query?rt=xml",
        cacheKey: _cacheKeyPrefix + 'killedApps'
    };
    var _appStateWS = {
        url: "/appstate/query?ma={0}:{1}",
        cacheKey: _cacheKeyPrefix + 'appState.{0}.{1}'
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
    var appStatus = {
        OK: "1",
        UnknownAssetId: "2",
        KilledAsset: "3",
        NoEntitlement: "4",
        ServerError: "9",
        BadRequest: "10",
        VersionNotSupported: "15"
    };
    var authNStatus = {
        NotAttempted: "-1",
        CheckFailed: "0",
        Authenticated: "1",
        Anonymous: "2",
        Unknown: "3"
    };
    function _onGetAppInstallInfoCompleted(executor) {
        var context = executor.get_webRequest().get_userContext();
        var url = executor.get_webRequest().get_url();
        if (executor.get_timedOut()) {
            OsfMsAjaxFactory.msAjaxDebug.trace("Request timed out: " + url);
            _invokeCallbackTag(context.callback, statusCode.Failed, null, null, executor, 0x0085a297);
        }
        else if (executor.get_aborted()) {
            OsfMsAjaxFactory.msAjaxDebug.trace("Request aborted: " + url);
            _invokeCallbackTag(context.callback, statusCode.Failed, null, null, executor, 0x0085a298);
        }
        else if (executor.get_responseAvailable()) {
            var statusText = executor.get_statusText();
            OsfMsAjaxFactory.msAjaxDebug.trace("Request to " + url + " came back with the status code: " + statusText);
            var manifestAndEToken = { manifest: "", etoken: "", status: appStatus.OK };
            if (executor.get_statusCode() == 200) {
                manifestAndEToken.manifest = executor.get_responseData();
            }
            else if (executor.get_statusCode() == 204 || executor.get_statusCode() == 1223) {
                manifestAndEToken.status = appStatus.NoEntitlement;
            }
            else if (executor.get_statusCode() == 400) {
                manifestAndEToken.status = appStatus.BadRequest;
            }
            else if (executor.get_statusCode() == 410) {
                manifestAndEToken.status = appStatus.KilledAsset;
            }
            else if (executor.get_statusCode() == 404) {
                manifestAndEToken.status = appStatus.UnknownAssetId;
            }
            else if (executor.get_statusCode() == 412) {
                manifestAndEToken.status = appStatus.VersionNotSupported;
            }
            else {
                manifestAndEToken.status = appStatus.ServerError;
            }
            manifestAndEToken.cached = false;
            _invokeCallbackTag(context.callback, statusCode.Succeeded, manifestAndEToken, null, executor, 0x0080b720);
        }
        else {
            OsfMsAjaxFactory.msAjaxDebug.trace("Request failed: " + url);
            _invokeCallbackTag(context.callback, statusCode.Failed, null, null, executor, 0x0085a299);
        }
    }
    ;
    function _onGetAppStateCompleted(executor) {
        var appState = {};
        var context = executor.get_webRequest().get_userContext();
        var responseXml = executor.get_responseData();
        var xmlProcessor = new OSF.XmlProcessor(responseXml, _omexXmlNamespaces);
        var root = xmlProcessor.getDocumentElement();
        xmlProcessor.readAttributes(root, { "rr": "refreshRate" }, appState);
        var resultNode = xmlProcessor.selectSingleNode("o:results");
        var langNode = xmlProcessor.selectSingleNode("o:lang", resultNode);
        var assetNode = xmlProcessor.selectSingleNode("o:asset", langNode);
        xmlProcessor.readAttributes(assetNode, { "assetid": "assetId", "prodid": "productId", "ver": "version", "state": "state", "tdurl": "takeDownUrl", "upv": "unsafePreviousVersion", "expiry": "expirationDate" }, appState);
        _invokeCallbackTag(context.callback, statusCode.Succeeded, appState, null, executor, 0x0080b722);
    }
    ;
    function _onGetKilledAppsCompleted(executor) {
        var context = executor.get_webRequest().get_userContext();
        var responseXml = executor.get_responseData();
        var killedAppsInfo = {};
        var xmlProcessor = new OSF.XmlProcessor(responseXml, _omexXmlNamespaces);
        var root = xmlProcessor.getDocumentElement();
        xmlProcessor.readAttributes(root, { "rr": "refreshRate" }, killedAppsInfo);
        killedAppsInfo.killedApps = [];
        var assetsNode = xmlProcessor.selectSingleNode("o:assets");
        var assetNodes = xmlProcessor.selectNodes("o:asset", assetsNode);
        var assetNode;
        var killedApp;
        for (var i = 0; i < assetNodes.length; ++i) {
            assetNode = assetNodes[i];
            killedApp = {};
            xmlProcessor.readAttributes(assetNode, { "assetid": "assetId", "pid": "productId" }, killedApp);
            killedAppsInfo.killedApps.push(killedApp);
        }
        if (assetNodes.length >= 0) {
            _cacheManager.SetCacheItem(context.cacheKey, killedAppsInfo, killedAppsInfo.refreshRate / _hourToDayConversionFactor);
        }
        _invokeCallbackTag(context.callback, statusCode.Succeeded, killedAppsInfo, null, executor, 0x0080b723);
    }
    ;
    function _onGetAuthNStatusCompleted(executor) {
        var context = executor.get_webRequest().get_userContext();
        var url = executor.get_webRequest().get_url();
        if (executor.get_timedOut() || executor.get_aborted()) {
            _invokeCallbackTag(context.callback, statusCode.Failed, null, null, executor, 0x0080b740);
        }
        else if (executor.get_responseAvailable()) {
            var statusText = executor.get_statusText();
            OsfMsAjaxFactory.msAjaxDebug.trace("Request to " + url + " came back with the status code: " + statusText);
            var logOnStatus = authNStatus.Unknown;
            if (executor.get_statusCode() == 200) {
                logOnStatus = authNStatus.Authenticated;
            }
            else if (executor.get_statusCode() == 401 || executor.get_statusCode() == 403) {
                logOnStatus = authNStatus.Anonymous;
            }
            else {
                logOnStatus = authNStatus.Unknown;
            }
            _invokeCallbackTag(context.callback, statusCode.Succeeded, logOnStatus, null, logOnStatus != authNStatus.Anonymous ? executor : null, 0x0080b741);
        }
        else {
            _invokeCallbackTag(context.callback, statusCode.Failed, null, "Getting authentication status failed.", executor, 0x0080b742);
        }
    }
    ;
    return {
        initialize: function OSF_OmexAnonymousProvider$initialize() {
            if (!_hasBeenInitialized) {
                _serviceEndPoint = Microsoft.Office.Common.XdmCommunicationManager.createServiceEndPoint(_clientID);
                _serviceEndPoint.registerMethod("OMEX_getAppStateAsync", this.getAppStateAsync, Microsoft.Office.Common.InvokeType.async, false);
                _serviceEndPoint.registerMethod("OMEX_getKilledAppsAsync", this.getKilledAppsAsync, Microsoft.Office.Common.InvokeType.async, false);
                _serviceEndPoint.registerMethod("OMEX_getManifestAndETokenAsync", this.getManifestAndETokenAsync, Microsoft.Office.Common.InvokeType.async, false);
                _serviceEndPoint.registerMethod("OMEX_removeCacheAsync", this.removeCacheAsync, Microsoft.Office.Common.InvokeType.async, false);
                _serviceEndPoint.registerMethod("OMEX_clearCacheAsync", this.clearCacheAsync, Microsoft.Office.Common.InvokeType.async, false);
                _serviceEndPoint.registerMethod("OMEX_isProxyReady", _isProxyReady, Microsoft.Office.Common.InvokeType.async, false);
                _serviceEndPoint.registerMethod("OMEX_getAuthNStatus", this.getAuthNStatus, Microsoft.Office.Common.InvokeType.async, false);
                var conversationId = OSF.OUtil.getConversationId();
                var conversationUrl = OSF.OUtil.getConversationUrl();
                _serviceEndPoint.registerConversation(conversationId, conversationUrl, null, OSF.OUtil.parseSerializerVersion(true));
                _hasBeenInitialized = true;
            }
        },
        getAppStateAsync: function OSF_OmexAnonymousProvider$getAppStateAsync(params, callback) {
            OSF.OUtil.validateParamObject(params, {
                "assetID": { type: String, mayBeNull: false },
                "contentMarket": { type: String, mayBeNull: false },
                "clientName": { type: String, mayBeNull: true },
                "clientVersion": { type: String, mayBeNull: true }
            }, callback);
            params.clearAppState = params.clearAppState || false;
            try {
                var cacheKey = OSF.OUtil.formatString(_appStateWS.cacheKey, params.contentMarket, params.assetID);
                params.callback = callback;
                params.cacheKey = cacheKey;
                var requestUrl = _appStateWS.url;
                requestUrl = OSF.OUtil.formatString(requestUrl, params.contentMarket, params.assetID);
                var queryStringComponents = {};
                queryStringComponents[_queryStringParameters.clientName] = params.clientName;
                queryStringComponents[_queryStringParameters.clientVersion] = params.clientVersion;
                requestUrl += _createQueryStringFragment(queryStringComponents);
                params._onCompleteHandler = _onGetAppStateCompleted;
                _sendWebRequest(requestUrl, 'GET', { 'Content-Type': 'text/xml' }, _onCompleted, params);
            }
            catch (ex) {
                OsfMsAjaxFactory.msAjaxDebug.trace("Getting app state failed: " + ex);
                _invokeCallbackTag(callback, statusCode.Failed, null, "Getting app state failed.", null, 0x0085a29a);
            }
        },
        getKilledAppsAsync: function OSF_OmexAnonymousProvider$getKilledAppsAsync(params, callback) {
            var e = Function._validateParams(arguments, [{ name: "params", type: Object, mayBeNull: false },
                { name: "callback", type: Function, mayBeNull: false }
            ]);
            if (e)
                throw e;
            params.clearKilledApps = params.clearKilledApps || false;
            try {
                var cacheKey = _killedAppsWS.cacheKey;
                if (params.clearKilledApps) {
                    _cacheManager.RemoveCacheItem(cacheKey);
                }
                else {
                    var value = _cacheManager.GetCacheItem(cacheKey);
                    if (value) {
                        _invokeCallbackTag(callback, statusCode.Succeeded, value, null, null, 0x0080b744);
                        return;
                    }
                }
                params.callback = callback;
                params.cacheKey = cacheKey;
                var requestUrl = _killedAppsWS.url;
                var queryStringComponents = {};
                queryStringComponents[_queryStringParameters.clientName] = params.clientName;
                queryStringComponents[_queryStringParameters.clientVersion] = params.clientVersion;
                requestUrl += _createQueryStringFragment(queryStringComponents);
                params._onCompleteHandler = _onGetKilledAppsCompleted;
                _sendWebRequest(requestUrl, 'GET', { 'Content-Type': 'text/xml' }, _onCompleted, params);
            }
            catch (ex) {
                OsfMsAjaxFactory.msAjaxDebug.trace("Getting killed bits failed: " + ex);
                _invokeCallbackTag(callback, statusCode.Failed, null, "Getting killed bits failed.", null, 0x0085a29b);
            }
        },
        getManifestAndETokenAsync: function OSF_OmexAnonymousProvider$getManifestAndETokenAsync(params, callback) {
            OSF.OUtil.validateParamObject(params, {
                "assetID": { type: String, mayBeNull: false },
                "applicationName": { type: String, mayBeNull: false },
                "contentMarket": { type: String, mayBeNull: false },
                "clientName": { type: String, mayBeNull: true },
                "clientVersion": { type: String, mayBeNull: true }
            }, callback);
            params.clearManifest = params.clearManifest || false;
            try {
                var cacheKey = OSF.OUtil.formatString(_appInstallInfoWS.cacheKey, params.assetID, params.contentMarket);
                params.callback = callback;
                params.cacheKey = cacheKey;
                var requestUrl = _appInstallInfoWS.url;
                requestUrl = OSF.OUtil.formatString(requestUrl, params.contentMarket, params.assetID);
                var queryStringComponents = {};
                queryStringComponents[_queryStringParameters.clientName] = params.clientName;
                queryStringComponents[_queryStringParameters.clientVersion] = params.clientVersion;
                requestUrl += _createQueryStringFragment(queryStringComponents);
                _sendWebRequest(requestUrl, 'GET', { 'Content-Type': 'text/xml' }, _onGetAppInstallInfoCompleted, params);
            }
            catch (ex) {
                OsfMsAjaxFactory.msAjaxDebug.trace("Getting manifest and token failed: " + ex);
                _invokeCallbackTag(callback, statusCode.Failed, null, "Getting manifest and token failed.", null, 0x0085a29c);
            }
        },
        removeCacheAsync: function OSF_OmexAnonymousProvider$removeCacheAsync(params, callback) {
            OSF.OUtil.validateParamObject(params, {
                "assetID": { type: String, mayBeNull: false },
                "applicationName": { type: String, mayBeNull: false },
                "contentMarket": { type: String, mayBeNull: false }
            }, callback);
            try {
                var cacheKey;
                if (params.clearManifest) {
                    cacheKey = OSF.OUtil.formatString(_appInstallInfoWS.cacheKey, params.assetID, params.contentMarket);
                    _cacheManager.RemoveCacheItem(cacheKey);
                }
                if (params.clearAppState) {
                    cacheKey = OSF.OUtil.formatString(_appStateWS.cacheKey, params.contentMarket, params.assetID);
                    _cacheManager.RemoveCacheItem(cacheKey);
                }
                _invokeCallbackTag(callback, statusCode.Succeeded, null, null, null, 0x0080b747);
            }
            catch (ex) {
                OsfMsAjaxFactory.msAjaxDebug.trace("Removing cache failed: " + ex);
                _invokeCallbackTag(callback, statusCode.Failed, null, "Removing cache failed.", null, 0x0085a29d);
            }
        },
        clearCacheAsync: function OSF_OmexAnonymousProvider$clearCacheAsync(params, callback) {
            var e = Function._validateParams(arguments, [{ name: "params", type: Object, mayBeNull: false },
                { name: "callback", type: Function, mayBeNull: false }
            ]);
            if (e)
                throw e;
            try {
                _cacheManager.RemoveAll(_cacheKeyPrefix);
                _invokeCallbackTag(callback, statusCode.Succeeded, null, null, null, 0x0080b749);
            }
            catch (ex) {
                OsfMsAjaxFactory.msAjaxDebug.trace("Clearing cache failed: " + ex);
                _invokeCallbackTag(callback, statusCode.Failed, null, "Clearing cache failed.", null, 0x0085a29e);
            }
        },
        getAuthNStatus: function OSF_OmexAnonymousProvider$getAuthNStatus(params, callback) {
            var e = Function._validateParams(arguments, [{ name: "params", type: Object, mayBeNull: false },
                { name: "callback", type: Function, mayBeNull: false }
            ]);
            if (e)
                throw e;
            try {
                params.callback = callback;
                _sendWebRequest("/gatedserviceextension.aspx?fromAR=3", "HEAD", { "Cookie": document.cookie }, _onGetAuthNStatusCompleted, params);
            }
            catch (ex) {
                OsfMsAjaxFactory.msAjaxDebug.trace("Getting authentication status failed: " + ex);
                callback({ "status": statusCode.Failed, "result": null });
            }
        }
    };
})();
OSF.OmexAnonymousProvider.initialize();
