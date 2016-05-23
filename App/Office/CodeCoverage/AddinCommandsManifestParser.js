var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
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
