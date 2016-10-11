/* WinRT host page JavaScript library */
/* Version: 16.0.7504.3000 */
/*
	Copyright (c) Microsoft Corporation.  All rights reserved.
*/


/*
	Your use of this file is governed by the Microsoft Services Agreement http://go.microsoft.com/fwlink/?LinkId=266419.
*/

var CrossIFrameCommon;
(function (CrossIFrameCommon) {
    (function (CallbackType) {
        CallbackType[CallbackType["MethodCallback"] = 0] = "MethodCallback";
        CallbackType[CallbackType["EventCallback"] = 1] = "EventCallback";
    })(CrossIFrameCommon.CallbackType || (CrossIFrameCommon.CallbackType = {}));
    var CallbackType = CrossIFrameCommon.CallbackType;
    var CallbackData = (function () {
        function CallbackData(callbackType, callbackId, params) {
            this.callbackType = callbackType;
            this.callbackId = callbackId;
            this.params = params;
        }
        return CallbackData;
    })();
    CrossIFrameCommon.CallbackData = CallbackData;
})(CrossIFrameCommon || (CrossIFrameCommon = {}));
var WinRT;
(function (WinRT) {
    function Init() {
        window.addEventListener("message", WinRT.OnReceiveMessage);
    }
    WinRT.Init = Init;
    function LoadAddIn(url) {
        var frame = document.getElementById("addInFrame");
        frame.setAttribute("src", url);
    }
    WinRT.LoadAddIn = LoadAddIn;
    function agaveHostCallback(callbackId, params) {
        var cbData = new CrossIFrameCommon.CallbackData(CrossIFrameCommon.CallbackType.MethodCallback, callbackId, params);
        var frame = document.getElementById("addInFrame");
        frame.contentWindow.postMessage(JSON.stringify(cbData), "*");
    }
    WinRT.agaveHostCallback = agaveHostCallback;
    function agaveHostEventCallback(callbackId, params) {
        var cbData = new CrossIFrameCommon.CallbackData(CrossIFrameCommon.CallbackType.EventCallback, callbackId, params);
        var frame = document.getElementById("addInFrame");
        frame.contentWindow.postMessage(JSON.stringify(cbData), "*");
    }
    WinRT.agaveHostEventCallback = agaveHostEventCallback;
    function OnReceiveMessage(event) {
        var frame = document.getElementById("addInFrame");
        if (event.source != frame.contentWindow) {
            return;
        }
        window.external.notify(event.data);
    }
    WinRT.OnReceiveMessage = OnReceiveMessage;
})(WinRT || (WinRT = {}));
function agaveHostCallback(callbackId, params) {
    WinRT.agaveHostCallback(callbackId, params);
}
function agaveHostEventCallback(callbackId, params) {
    WinRT.agaveHostEventCallback(callbackId, params);
}
function LoadAddIn(url) {
    WinRT.LoadAddIn(url);
}
