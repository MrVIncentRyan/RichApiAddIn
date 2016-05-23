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
