/* Office rich client gallery insertion dialog JavaScript file */
/* Version: 16.0.6723.3002 */
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
        UI.DefaultHeaderHeight = 52;
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
            this.focusOnCallBack = focusOnCallBack;
            this.appOptions = null;
            this.itemCreated = false;
        }
        GalleryItem.prototype.displayAgave = function (documentFragment) {
            var moeDiv = document.createElement("div");
            documentFragment.appendChild(moeDiv);
            WEF.WefGalleryHelper.addClass(moeDiv, "Moe");
            moeDiv.setAttribute("data-ri", this.index.toString());
            WEF.WefGalleryHelper.dpiScale(moeDiv);
            WEF.WefGalleryHelper.dpiScaleMarginLeft(moeDiv);
            moeDiv.oncontextmenu = function WEF_GalleryItem_displayAgave$oncontextmenu() {
                return false;
            };
            this.galleryItem = moeDiv;
        };
        GalleryItem.prototype.updateImage = function (insertHandler) {
            var _this = this;
            if (!this.galleryItem) {
                return;
            }
            var moeDiv = this.galleryItem;
            if (!this.itemCreated) {
                moeDiv.onclick = function () {
                    WEF.IMPage.selectGalleryItems(_this.index);
                };
                moeDiv.ondblclick = function () {
                    insertHandler(_this);
                };
                moeDiv.onmousedown = function (e) {
                    if (!e)
                        e = event;
                    if (e.which === 3 || e.button === 2) {
                        if (_this.appOptions) {
                            _this.appOptions.popupMenu();
                        }
                    }
                };
                moeDiv.onmouseover = function () {
                    WEF.WefGalleryHelper.addClass(_this.galleryItem, "mouseover");
                    _this.appOptions.showOptionsButton();
                };
                moeDiv.onmouseout = function () {
                    WEF.WefGalleryHelper.removeClass(_this.galleryItem, "mouseover");
                    if (!WEF.WefGalleryHelper.hasClass(_this.galleryItem, "selected")) {
                        _this.appOptions.hideOptionsButton();
                    }
                };
                var agaveIconUrl = this.result.iconUrl;
                var moeInnerDiv = document.createElement("div");
                moeDiv.appendChild(moeInnerDiv);
                WEF.WefGalleryHelper.addClass(moeInnerDiv, "MoeInner");
                WEF.WefGalleryHelper.dpiScale(moeInnerDiv);
                moeInnerDiv.setAttribute("title", this.result.description);
                moeInnerDiv.setAttribute("tabindex", "-1");
                var tnDiv = document.createElement("div");
                moeInnerDiv.appendChild(tnDiv);
                WEF.WefGalleryHelper.addClass(tnDiv, "Tn");
                var detailsDiv = document.createElement("div");
                moeInnerDiv.appendChild(detailsDiv);
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
                moeInnerDiv.setAttribute("aria-label", this.result.displayName);
                moeInnerDiv.setAttribute("role", "Option");
                this.appOptions = WEF.IMPage.menuHandler.createAppOptions(this.result);
                var optionsButton = this.appOptions.createOptionsButton(this.index, tnDiv, img);
                if (optionsButton) {
                    moeInnerDiv.appendChild(optionsButton);
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
            this.btnAction = null;
            this.btnCancel = null;
            this.btnDone = null;
            this.btnTrustAll = null;
            this.documentAppsMsg = null;
            this.documentAppsMsgText = null;
            this.errorMessage = null;
            this.footer = null;
            this.footerLink = null;
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
            this.readMoreATag = null;
            this.selectedDescriptionReadMoreLink = null;
            this.selectedDescriptionText = null;
            this.selectedItem = null;
            this.tabs = null;
            this.uploadMenuDiv = null;
            this.refreshMenuDiv = null;
            this.manageMenuDiv = null;
            this.menuRightSeparatorDiv = null;
            this.tabTitles = [];
            this.menuSeparatorWidth = null;
            this.menuRightMaxPossibleWidth = null;
            this.galleryItems = null;
            this.uiState = { "Ready": false, "StoreIdBeforeReady": "", "ErrorBeforeReady": "", "ErrorLinkTextBeforeReady": "", "ErrorLinkHandlerBeforeReady": null };
            this.currentIndex = -1;
            this.results = null;
            this.height = "100%";
            this.width = "100%";
            this.itemsPerRow = null;
            this.leftKeyHandler = null;
            this.rightKeyHandler = null;
            this.keyHandlers = null;
            this.menuHandler = null;
            this.modalDialog = null;
            this.storeTab = null;
            this.firstTabATag = null;
            this.totalSessionTime = 0;
            this.trustPageSessionTime = 0;
            this.envSetting = {};
            this.isUploadFileDevCatalogEnabled = false;
            this.isAppCommandEnabled = false;
            this.moveLeft = function () {
                _this.currentIndex--;
                if (_this.currentIndex >= 0) {
                    _this.selectGalleryItems(_this.currentIndex);
                }
                else {
                    _this.currentIndex = 0;
                }
            };
            this.moveRight = function (numOfItems) {
                _this.currentIndex++;
                if (_this.currentIndex < numOfItems) {
                    _this.selectGalleryItems(_this.currentIndex);
                }
                else {
                    _this.currentIndex = numOfItems - 1;
                }
            };
            this.upKeyHandler = function () {
                _this.currentIndex -= _this.itemsPerRow;
                if (_this.currentIndex >= 0) {
                    _this.selectGalleryItems(_this.currentIndex);
                }
                else {
                    _this.currentIndex += _this.itemsPerRow;
                }
            };
            this.downKeyHandler = function (numOfItems) {
                if (_this.currentIndex >= 0) {
                    _this.currentIndex += _this.itemsPerRow;
                }
                else {
                    _this.currentIndex = 0;
                }
                if (_this.currentIndex < numOfItems) {
                    _this.selectGalleryItems(_this.currentIndex);
                }
                else {
                    _this.currentIndex -= _this.itemsPerRow;
                }
            };
            this.tabKeyHandler = function (event, element) {
                if (!event.shiftKey && element.getAttribute("id") == "BtnCancel" && event.preventDefault && _this.firstTabATag) {
                    _this.firstTabATag.focus();
                    event.preventDefault();
                }
            };
            this.galleryKeyHandler = function (e) {
                var numOfItems = 0;
                if (_this.results) {
                    numOfItems = _this.results.length;
                }
                if (!e)
                    e = event;
                for (var i = 0; i < _this.keyHandlers.length; i++) {
                    var keyHandler = _this.keyHandlers[i];
                    if (keyHandler.handleKey(e)) {
                        return;
                    }
                }
                var eventTarget = e.srcElement ? e.srcElement : e.target;
                switch (e.keyCode) {
                    case 9:
                        _this.tabKeyHandler(e, eventTarget);
                        break;
                    case 13:
                        _this.executeButtonCommand(eventTarget);
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
                        _this.leftKeyHandler(numOfItems);
                        break;
                    case 38:
                        _this.upKeyHandler();
                        break;
                    case 39:
                        _this.rightKeyHandler(numOfItems);
                        break;
                    case 40:
                        _this.downKeyHandler(numOfItems);
                        break;
                    default:
                        return;
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
                    _this.delayFunction(_this.loadVisibleImages);
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
                    }
                    if (_this.delaying) {
                        setTimeout(_this.loadVisibleImages, 3000);
                        _this.delaying = false;
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
        WefGalleryPage.prototype.executeButtonCommand = function (element) {
            if (WEF.WefGalleryHelper.hasClass(element, "MoeInner")) {
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
                    tabId = child.getAttribute("id");
                    if (tabId == selectedTabId) {
                        WEF.WefGalleryHelper.addClass(child.firstChild, "TabSelected");
                        var storeId = child.getAttribute("data-storeId");
                        var storeType = parseInt(child.getAttribute("data-storeType"));
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
                    }
                    if (tempStoreType == WEF.StoreTypeEnum.Recommendation) {
                        this.storeTab = createdTab;
                    }
                }
            }
            this.setOptionBarElementMaxSize(this.tabTitles);
            if (this.tabs.childNodes.length > 0) {
                if (selectedTab) {
                    WEF.WefGalleryHelper.addClass(selectedTab.childNodes[0], "selected");
                }
                else if (this.tabs.childNodes[0].childNodes.length > 0) {
                    WEF.WefGalleryHelper.addClass(this.tabs.childNodes[0].childNodes[0], "selected");
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
            this.uploadATag.setAttribute("tabIndex", "0");
            this.uploadATag.setAttribute("title", Strings.wefgallery.L_AddinCommands_UploadMyAddin_Txt);
            var refreshCurrentTab = function () {
                _this.cleanUpGallery();
                _this.restoreFooterLink();
                _this.showContent(true);
            };
            var refreshATag = document.getElementById('RefreshInner');
            refreshATag.setAttribute("title", Strings.wefgallery.L_WefDialog_RefreshButton_Tooltip);
            refreshATag.onclick = function WEF_WefGalleryPage_initializeGalleryUI_refreshATag$onclick() { refreshCurrentTab(); };
            refreshATag.setAttribute("tabIndex", "0");
            var footerLinkATag = document.getElementById('FooterLinkATag');
            footerLinkATag.setAttribute("tabIndex", "0");
            footerLinkATag.setAttribute("title", Strings.wefgallery.L_Footer_Link_Text_Tooltip);
            this.documentAppsMsg.setAttribute("title", Strings.wefgallery.L_TrustUx_AppsMessage);
            this.documentAppsMsg.firstChild.innerText = Strings.wefgallery.L_TrustUx_AppsMessage;
            this.readMoreATag.setAttribute("tabIndex", "0");
            this.readMoreATag.setAttribute("title", Strings.wefgallery.L_TrustUx_ReadMoreLink_Txt_Tooltip);
            this.permissionATag.setAttribute("tabIndex", "0");
            this.permissionATag.setAttribute("title", Strings.wefgallery.L_Permission_Link_Txt_Tooltip);
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
            this.noAppsMessage.setAttribute("title", Strings.wefgallery.L_OfficeStore_Button_Tooltip);
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
        WefGalleryPage.prototype.createTab = function (tabsDiv, tabOrder, tabName, storeId, storeType) {
            var me = this;
            if (tabsDiv.childNodes.length != 0) {
                var separatorDiv = document.createElement('div');
                WEF.WefGalleryHelper.addClass(separatorDiv, "separator");
                separatorDiv.innerHTML = "|";
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
            aTag.setAttribute("tabIndex", "0");
            tabDiv.appendChild(aTag);
            if (tabOrder == 1) {
                aTag.focus();
                this.firstTabATag = aTag;
            }
            tabDiv.setAttribute("id", tabName);
            tabDiv.setAttribute("data-storeId", storeId);
            tabDiv.setAttribute("data-storeType", storeType.toString());
            if (pageUrl) {
                tabDiv.setAttribute("data-pageUrl", pageUrl);
            }
            tabDiv.onclick = function WEF_WefGalleryPage_createTab_tabDiv$onclick() { me.toggleTabSelection(this, null); };
            return tabDiv;
        };
        WefGalleryPage.prototype.galleryScrollHandler = function () {
            this.menuHandler.hideMenu(true);
            this.delayFunction(this.loadVisibleImages);
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
            this.delayFunction(this.loadVisibleImages);
        };
        WefGalleryPage.prototype.processAddinLoadingErrors = function (results) {
            for (var i = 0; i < results.length; i++) {
                if (results[i].hasLoadingError) {
                    this.showError(Strings.wefgallery.L_AddinsHasLoadingErrors, this.currentStoreId);
                    break;
                }
            }
        };
        WefGalleryPage.prototype.delayFunction = function (delayFunction) {
            if (!this.delayTime || this.delaying == false || ((new Date().getTime() - this.delayTime) > 1000)) {
                this.delayTime = new Date().getTime();
                this.delaying = true;
                setTimeout(delayFunction, this.delayLoad);
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
                    if (WEF.WefGalleryHelper.hasClass(item.galleryItem, "selected")) {
                        if (forceSelected == false) {
                            WEF.WefGalleryHelper.removeClass(item.galleryItem, "selected");
                            this.deSelectBtnAction();
                        }
                        else {
                            this.currentIndex = index;
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
                        this.currentIndex = index;
                        if (item.galleryItem.children.length > 0) {
                            item.galleryItem.children[0].focus();
                            item.galleryItem.children[0].setAttribute("aria-selected", "true");
                        }
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
                if (item.galleryItem.children.length > 0) {
                    item.galleryItem.children[0].removeAttribute("aria-selected");
                }
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
            this.readMoreATag = document.getElementById('ReadMoreLink');
            this.permissionATag = document.getElementById('PermissionLink');
            this.documentAppsMsg = document.getElementById('DocumentAppsMessageId');
            this.documentAppsMsgText = document.getElementById('DocumentAppsMessageText');
            this.btnAction = document.getElementById('BtnAction');
            this.btnCancel = document.getElementById('BtnCancel');
            this.btnTrustAll = document.getElementById('BtnTrustAll');
            this.btnDone = document.getElementById('BtnDone');
            this.notification = document.getElementById("Notification");
            this.errorMessage = document.getElementById('ErrorMessage');
            this.notificationDismiss = document.getElementById('NotificationDismiss');
            this.notificationDismissImg = document.getElementById('DismissImg');
            this.menuRight = document.getElementById('MenuRight');
            this.noAppsMessage = document.getElementById('NoAppsMessage');
            this.noAppsMessageTitle = document.getElementById('NoAppsMessageTitle');
            this.noAppsMessageText = document.getElementById('NoAppsMessageText');
            this.officeStoreBtn = document.getElementById('BtnStore');
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
            this.modalDialog = new WEF.AppManagement.ModalDialog();
            this.menuHandler = new WEF.AppManagement.MenuHandler(this.galleryContainer, this.modalDialog);
            this.keyHandlers = [this.menuHandler, this.modalDialog];
            window.document.onkeydown = function (e) {
                _this.galleryKeyHandler(e);
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
            function ModalDialog() {
                this.overlayDiv = null;
                this.dialogDiv = null;
                this.buttonDiv = null;
                this.confirmMessageDiv = null;
                this.buttonElements = [];
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
            ModalDialog.prototype.handleKey = function (ev) {
                if (!this.isDialogVisible()) {
                    return false;
                }
                var handled = true;
                switch (ev.keyCode) {
                    case 9:
                        this.onTab(ev);
                        break;
                    case 27:
                        this.hideDialog();
                        break;
                }
                return handled;
            };
            ModalDialog.prototype.hideDialog = function () {
                if (!this.isDialogVisible()) {
                    return;
                }
                var tabElements = this.getTabbableElements();
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
                    }
                }
                this.dialogDiv.style.display = "none";
                this.overlayDiv.style.display = "none";
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
                ev.preventDefault();
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
                var focusInMenuCheck = function (event) {
                    var hideMenu = false;
                    if (event.relatedTarget !== undefined) {
                        if (_this.menuDiv.contains(event.relatedTarget) == false) {
                            hideMenu = true;
                        }
                    }
                    else if (!_this.menuDiv.querySelector(":focus")) {
                        hideMenu = true;
                    }
                    if (hideMenu) {
                        _this.hideMenu(true);
                    }
                };
                WEF.WefGalleryHelper.addEventListener(this.menuDiv, "focusout", focusInMenuCheck);
            }
            MenuHandler.prototype.createAppOptions = function (result) {
                return new AppOptions(result, this);
            };
            MenuHandler.prototype.handleKey = function (ev) {
                if (this.isMenuVisible() == false) {
                    return false;
                }
                var handled = false;
                switch (ev.keyCode) {
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
            MenuHandler.prototype.hideMenu = function (logData) {
                if (this.isMenuVisible()) {
                    this.menuDiv.style.display = "none";
                    if (logData) {
                        this.logData(this.currentResult, AppManagementAction.Cancel, 0);
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
                    optionsButton = document.createElement("button");
                    WEF.WefGalleryHelper.addClass(optionsButton, "OptionsButton");
                    optionsButton.setAttribute("type", "button");
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
        };
        return WefGalleryPage_Native;
    })(WEF.WefGalleryPage);
    WEF.WefGalleryPage_Native = WefGalleryPage_Native;
})(WEF || (WEF = {}));
var WEF;
(function (WEF) {
    var ClientFacade_RichClient = (function (_super) {
        __extends(ClientFacade_RichClient, _super);
        function ClientFacade_RichClient(external) {
            _super.call(this);
            this.onGetEntitlementsInternal = this.onGetEntitlements;
            this.external = external;
            try {
                var queryStrObj = ClientFacade_RichClient.parseUrlQueryString(window.location.search);
                this.envSetting = JSON.parse(queryStrObj["env_setting"]);
            }
            catch (e) {
            }
        }
        ClientFacade_RichClient.prototype.runShowIt = function () {
            if (typeof this.external.GetProviders !== 'undefined') {
                this.external.GetProviders(this.onGetProviders);
            }
        };
        ClientFacade_RichClient.prototype.getEntitlements = function (storeId, refresh, onGetEntitlements) {
            if (typeof this.external.GetEntitlements !== 'undefined') {
                this.external.GetEntitlements(storeId, refresh, onGetEntitlements);
            }
        };
        ClientFacade_RichClient.prototype.insertAgave = function (item, currentStoreType) {
            if (typeof this.external.Insert !== 'undefined') {
                this.external.Insert(item.result.id, item.result.targetType, item.result.appVersion, currentStoreType, item.result.storeId, currentStoreType == WEF.StoreTypeEnum.MarketPlace ? item.result.id : item.result.assetId, currentStoreType == WEF.StoreTypeEnum.MarketPlace ? item.result.id : item.result.assetStoreId, item.result.width, item.result.height);
            }
        };
        ClientFacade_RichClient.prototype.getInitTab = function () {
            return this.external.GetInitTab();
        };
        ClientFacade_RichClient.prototype.getPageUrl = function (pageType) {
            var pageUrl = "";
            if (typeof this.external.GetLandingPageUrl !== 'undefined') {
                try {
                    pageUrl = this.external.GetLandingPageUrl(pageType);
                }
                catch (ex) {
                    pageUrl = "";
                }
            }
            return pageUrl;
        };
        ClientFacade_RichClient.prototype.launchAppManagePage = function () {
            if (typeof this.external.LaunchUrl !== 'undefined' && typeof WEF.IMPage !== 'undefined') {
                WEF.WefGalleryHelper.saveRefreshRequired(true);
                this.external.LaunchUrl(WEF.IMPage.getAppManagePageUrl(), true);
            }
        };
        ClientFacade_RichClient.prototype.removeResultsDimensionInfo = function (results) {
        };
        ClientFacade_RichClient.prototype.removeAgave = function (result, currentStoreType, callback) {
            if (typeof this.external.RemoveAgave != 'undefined') {
                this.setOnRemoveAgaveCallback(callback);
                this.external.RemoveAgave(result.id, result.targetType, result.appVersion, currentStoreType, result.storeId, result.assetId, result.assetStoreId, this.onRemoveAgave);
            }
        };
        ClientFacade_RichClient.prototype.launchUrl = function (url, closeWindow) {
            if (typeof this.external.LaunchUrl != 'undefined') {
                this.external.LaunchUrl(url, closeWindow);
            }
        };
        ClientFacade_RichClient.prototype.preloadManifest = function (item, currentStoreType) {
            if (typeof this.external.PreloadManifest != 'undefined' && item) {
                if (currentStoreType == WEF.StoreTypeEnum.MarketPlace) {
                    this.external.PreloadManifest(item.result.id, item.result.appVersion, currentStoreType, item.result.storeId, item.result.assetId, item.result.assetId);
                }
                else {
                    this.external.PreloadManifest(item.result.id, item.result.appVersion, currentStoreType, item.result.storeId, item.result.assetId, item.result.assetStoreId);
                }
            }
        };
        ClientFacade_RichClient.prototype.logAppManagementAction = function (assetId, operationInfo, hresult) {
            if (typeof this.external.LogAppManagementAction != 'undefined') {
                this.external.LogAppManagementAction(assetId, operationInfo, hresult);
            }
        };
        ClientFacade_RichClient.parseUrlQueryString = function (queryStr) {
            if (queryStr.length <= 1) {
                return {};
            }
            queryStr = queryStr.substring(1);
            var searchRegex = new RegExp("([^&=]+)=?([^&]*)", "g");
            var queryStrObj = {}, match;
            while (match = searchRegex.exec(queryStr)) {
                queryStrObj[decodeURIComponent(match[1])] = decodeURIComponent(match[2]);
            }
            return queryStrObj;
        };
        return ClientFacade_RichClient;
    })(WEF.ClientFacade_Native);
    WEF.ClientFacade_RichClient = ClientFacade_RichClient;
    var WefGalleryPage_RichClient = (function (_super) {
        __extends(WefGalleryPage_RichClient, _super);
        function WefGalleryPage_RichClient() {
            var _this = this;
            _super.apply(this, arguments);
            this.invokeSignIn = function () {
                if (typeof _this.clientFacade.external.MountOrSignInLiveId !== 'undefined') {
                    _this.errorMessage.innerHTML = "";
                    _this.notification.style.visibility = 'hidden';
                    try {
                        _this.clientFacade.external.MountOrSignInLiveId();
                    }
                    catch (ex) {
                        _this.notification.style.visibility = 'visible';
                        _this.showError(Strings.wefgallery.L_SignInPromptLiveId, _this.currentStoreId, Strings.wefgallery.L_SignInLinkText, _this.invokeSignIn);
                        return;
                    }
                    _this.notification.style.visibility = 'visible';
                    _this.checkOmexProviderInfoShowContent();
                }
            };
        }
        WefGalleryPage_RichClient.prototype.onPageLoad = function () {
            if (typeof this.clientFacade.external.OnPageLoad != 'undefined') {
                this.clientFacade.external.OnPageLoad();
            }
        };
        WefGalleryPage_RichClient.prototype.onItemSelect = function (item) {
            this.clientFacade.preloadManifest(item, this.currentStoreType);
        };
        WefGalleryPage_RichClient.prototype.cancelDialog = function () {
            if (typeof this.clientFacade.external.CancelDialog != 'undefined') {
                this.clientFacade.external.CancelDialog();
            }
        };
        WefGalleryPage_RichClient.prototype.onMoreInfo = function () {
            var url = this.results[this.currentIndex].appEndNodeUrl;
            this.clientFacade.launchUrl(url, false);
        };
        WefGalleryPage_RichClient.prototype.trustAllAgaves = function () {
            if (typeof this.clientFacade.external.TrustAllInDocOmexApps != 'undefined') {
                this.clientFacade.external.TrustAllInDocOmexApps();
            }
        };
        WefGalleryPage_RichClient.prototype.canShowAppManagementMenu = function () {
            return true;
        };
        WefGalleryPage_RichClient.prototype.invokeWindowOpen = function (pageUrl) {
            this.clientFacade.launchUrl(pageUrl, false);
        };
        WefGalleryPage_RichClient.prototype.hideButtons = function () {
        };
        WefGalleryPage_RichClient.prototype.checkOmexProviderInfoShowContent = function () {
            try {
                this.clientFacade.external.GetProviders(this.clientFacade.onGetProvidersShowContent);
            }
            catch (ex) {
                this.showError(Strings.wefgallery.L_GetEntitilementsGeneralError);
            }
        };
        WefGalleryPage_RichClient.prototype.executeButtonCommand = function (element) {
            _super.prototype.executeButtonCommand.call(this, element);
            if (element.getAttribute("id") == "BtnTrustAll") {
                this.trustAllAgaves();
            }
        };
        WefGalleryPage_RichClient.prototype.wefGalleryAppOnLoad = function () {
            var _this = this;
            _super.prototype.wefGalleryAppOnLoad.call(this);
            this.btnTrustAll.onclick = function () { _this.trustAllAgaves(); };
            this.readMoreATag.onclick = function () { _this.onMoreInfo(); };
            this.permissionATag.onclick = function () { _this.onMoreInfo(); };
        };
        return WefGalleryPage_RichClient;
    })(WEF.WefGalleryPage_Native);
    WEF.WefGalleryPage_RichClient = WefGalleryPage_RichClient;
    WEF.setupClientSpecificWefGalleryPage = function () {
        var clientFacade = new ClientFacade_RichClient(window.external);
        WEF.IMPage = new WefGalleryPage_RichClient(clientFacade);
    };
})(WEF || (WEF = {}));
