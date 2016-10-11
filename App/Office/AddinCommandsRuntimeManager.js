var OfficeExt;
(function (OfficeExt) {
    OfficeExt.AddinActionContextMap = {};
    var taskpaneControlMap = {};
    (function (CommandInvokeMode) {
        CommandInvokeMode[CommandInvokeMode["Default"] = 1] = "Default";
        CommandInvokeMode[CommandInvokeMode["ExclusiveForCur"] = 2] = "ExclusiveForCur";
        CommandInvokeMode[CommandInvokeMode["ExclusiveForNew"] = 3] = "ExclusiveForNew";
    })(OfficeExt.CommandInvokeMode || (OfficeExt.CommandInvokeMode = {}));
    var CommandInvokeMode = OfficeExt.CommandInvokeMode;
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
        ControlBoundActionBuilder.prototype.buildShowTaskpane = function (sourceLocation, title, taskpaneId) {
            return this.CreateManifestAndActivationContextForTaskpaneCommand(sourceLocation, title, taskpaneId);
        };
        ControlBoundActionBuilder.prototype.buildCallFunction = function (functionFile, functionName, controlId) {
            return this.CreateManifestAndActivationContextForUILessCommand(functionFile, functionName, controlId);
        };
        ControlBoundActionBuilder.prototype.CreateManifestAndActivationContextForTaskpaneCommand = function (sourceLocation, title, taskpaneId) {
            var onHostReady = function (osfControlId, result) {
                if (result != ErrorCodes.ooeSuccess) {
                    return;
                }
                if (taskpaneId != null) {
                    taskpaneControlMap[taskpaneId] = osfControlId;
                }
            };
            var renderInExistingOsfControl = function (taskpaneId, contextActivationMgr) {
                var osfControlId = taskpaneControlMap[taskpaneId];
                if (osfControlId == null)
                    return false;
                var osfControl = contextActivationMgr.getOsfControl(osfControlId);
                if (osfControl == null)
                    return false;
                osfControl._refresh();
                return true;
            };
            return this.CreateManifestAndActivationContext(sourceLocation, title, taskpaneId, function (host, contextActivationMgr, entitlement, actionManifest) {
                if (taskpaneId != null) {
                    renderInExistingOsfControl(taskpaneId, contextActivationMgr);
                }
                host.showTaskpane(entitlement, actionManifest, onHostReady);
            });
        };
        ControlBoundActionBuilder.prototype.CreateManifestAndActivationContext = function (sourceLocation, title, taskpaneId, callback) {
            var _this = this;
            var actionManifest = OfficeExt.AddinCommandsManifestManager.createManifestForAddinAction(this.manifest, sourceLocation, title, null);
            return function (controlID) {
                var context = _this.contextProvider.createActionContext(controlID, function (host, contextActivationMgr) {
                    var entitlement = {
                        assetId: _this.entitlement.assetId + (taskpaneId || controlID),
                        appVersion: _this.entitlement.appVersion,
                        storeId: _this.entitlement.storeId,
                        storeType: OSF.StoreType.InMemory,
                        targetType: OSF.OsfControlTarget.TaskPane
                    };
                    OfficeExt.AddinCommandsManifestManager.cacheManifestForAction(actionManifest, entitlement.assetId, entitlement.appVersion);
                    callback(host, contextActivationMgr, entitlement, actionManifest);
                });
                OfficeExt.AddinActionContextMap[controlID] = context;
            };
        };
        ControlBoundActionBuilder.prototype.CreateManifestAndActivationContextForUILessCommand = function (functionFile, functionName, manifestControlId) {
            var _this = this;
            var onHostReady = function (osfControlId, result) {
                if (result != ErrorCodes.ooeSuccess) {
                    return;
                }
                var args = {
                    source: {
                        id: manifestControlId
                    }
                };
                AddinCommandsRuntimeManager.invokeAppCommand(osfControlId, functionName, args, function (status, data) {
                });
            };
            var actionManifest = OfficeExt.AddinCommandsManifestManager.createManifestForAddinAction(this.manifest, functionFile, null, null);
            return function (ribbonControlId) {
                var context = _this.contextProvider.createActionContext(ribbonControlId, function (host) {
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
                OfficeExt.AddinActionContextMap[ribbonControlId] = context;
            };
        };
        return ControlBoundActionBuilder;
    })();
    var ControlDeleterActionBuilder = (function () {
        function ControlDeleterActionBuilder(contextProvider, entitlement) {
            this.contextProvider = contextProvider;
            this.entitlement = entitlement;
        }
        ControlDeleterActionBuilder.prototype.buildShowTaskpane = function (sourceLocation, title) {
            var _this = this;
            return function (controlID) {
                delete OfficeExt.AddinActionContextMap[controlID];
                OfficeExt.AddinCommandsManifestManager.purgeManifestForAction(_this.entitlement.assetId + controlID, _this.entitlement.appVersion);
            };
        };
        ControlDeleterActionBuilder.prototype.buildCallFunction = function (functionFile, functionName, manifestControlId) {
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
        OsfHostEntry.prototype.invokeAppCommand = function (appCommandId, callbackName, eventObj, onComplete, timeout) {
            var args = {
                dispid: OSF.DDA.EventDispId.dispidAppCommandInvokedEvent,
                controlId: this.osfControlId
            };
            args[0] = appCommandId;
            args[1] = callbackName;
            args[2] = JSON.stringify(eventObj);
            timeout = timeout ? timeout : OsfHostEntry.defaultTimeout;
            var invokeMode = (eventObj && eventObj.commandMode) ? eventObj.commandMode : CommandInvokeMode.Default;
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
            if (!(invokeMode == CommandInvokeMode.ExclusiveForCur && queue.length > 0) &&
                !(invokeMode == CommandInvokeMode.ExclusiveForNew && queue.length > 1)) {
                queue.push(e);
            }
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
            entry.invokeAppCommand(appCommandId, callbackName, eventObj, onComplete);
        };
        AddinCommandsRuntimeManager.invocationCompleted = function (osfControlId, args) {
            var entry = AddinCommandsRuntimeManager.getOrCreateOsfHostEntry(osfControlId);
            var appCommandId = args[0];
            var status = args[1];
            var data = args[2] ? JSON.parse(args[2]) : null;
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
