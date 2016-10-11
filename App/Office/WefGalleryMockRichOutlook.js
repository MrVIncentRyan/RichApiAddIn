var WEF;
(function (WEF) {
    WEF.AGAVE_DEFAULT_ICON = "";
    /**
     * Page type enum values
     * this type must match the enum OmexMarketplacePage at ..\..\osfclient\osf\OfficeExtensionManager.h
     */
    WEF.PageTypeEnum = {
        "ManageApps": 0,
        "Recommendation": 2,
        "Landing": 3,
        "EndNode": 4,
        "Takedown": 5,
        "TermsAndConditions": 6,
        "RateAndReview": 7
    };
    /**
     * Page store Id values
     */
    WEF.PageStoreId = {
        "Recommendation": "{98143890-AC66-440E-A448-ED8771A02D52}"
    };
    /**
     * Enum for store type
     */
    WEF.StoreTypeEnum = {
        "MarketPlace": 0,
        "Catalog": 1,
        "Exchange": 3,
        "FileShare": 4,
        "Developer": 5,
        "Recommendation": 6,
        "ThisDocument": 8,
        "OneDrive": 9,
        "ExchangeCorporateCatalog": 10 // New Exchange corporate catalog introduced after App Command
    };
    /**
     * Enum for auth type used for a provider
     */
    WEF.AuthType = {
        "Anonymous": "0",
        "MSA": "1",
        "OrgId": "2",
        "ADAL": "3"
    };
    /**
     * Mapping for store types to localized strings
     */
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
    /**
     * Rich client native code HRESULT values
     */
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
        "E_OEM_REMOVED_FAILED": -2147209421 // Needs to match osfclient/inc/OsfErrorCodes.h
    };
    /**
     * Rich client native code OemStoreStatus. Mirrors definition in OfficeExtensionManager.h
     */
    WEF.OemStoreStatus = {
        "ossNotReady": 0,
        "ossSignInRequired": 1,
        "ossRegisteredButNotReady": 2,
        "ossRegisteredReady": 3,
        "ossUnregistered": 4 // Was registered as a store but has since been unregistered, so should be treated as absent.
    };
    /**
     * Different button groups will be shown for different tab scenario.
     */
    WEF.ActionButtonGroups = {
        "InsertCancel": 0,
        "ThisDocument": 1,
        "None": 2
    };
    /**
     * Event Message posted from omex
     */
    WEF.OmexMessage = {
        CancelDialog: "ESC_KEY",
        PreloadManifest: "PRELOAD_MANIFEST",
        RefreshRequired: "REFRESH_REQUIRED",
        WindowOpen: "WINDOW_OPEN"
    };
    /**
     * The node type of HTML node.
     */
    (function (NodeType) {
        NodeType[NodeType["ELEMENT"] = 1] = "ELEMENT";
        NodeType[NodeType["ATTRIBUTE"] = 2] = "ATTRIBUTE";
        NodeType[NodeType["TEXT"] = 3] = "TEXT";
    })(WEF.NodeType || (WEF.NodeType = {}));
    var NodeType = WEF.NodeType;
    /**
     * Encapsulate the meta data of an Agave
     */
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
        /**
         * Sort Agaves based on displayName in lexicographic(dictionary) order
         */
        AgaveInfo.cmpDisplayName = function (a, b) {
            if (a.displayName && b.displayName) {
                if (a.displayName.localeCompare(b.displayName) > 0) {
                    return 1; //a is sorted as higher index than b
                }
                else {
                    return -1; //a is sorted as lower index than b
                }
            }
            else {
                return -1; //a is sorted as lower index than b
            }
        };
        return AgaveInfo;
    })();
    WEF.AgaveInfo = AgaveInfo;
    /**
     * Constant values with user friendly names
     */
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
        /** A pre-calculated value for the width of separator in the option bar */
        UI.DefaultSeparatorWidth = 5;
        /** The margin between two elements in the option bar */
        UI.OptionBarElementMargin = 7;
        /** The gap between the tabs and menu in the option bar */
        UI.OptionBarMenuGap = 20;
    })(UI = WEF.UI || (WEF.UI = {}));
    ;
})(WEF || (WEF = {}));
/**
 * Common helper modules and util functions for WefGallery among all clients.
 */
var WEF;
(function (WEF) {
    var WefGalleryHelper;
    (function (WefGalleryHelper) {
        // IE uses className while other browser use class
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
        /**
         * Handle the error code from clients
         * @returns true tells the caller to just return and skip the following code
         */
        function handleErrorCode(errorCode, storeId, storeType, providerStatus) {
            var errorMessage = null;
            var skipShowApps = false;
            var signInRequired = false;
            // Check error scenario from provider status first
            if (providerStatus) {
                switch (providerStatus) {
                    case WEF.OemStoreStatus.ossSignInRequired:
                        errorMessage = getProperSignInMessageForStoreType(storeType);
                        signInRequired = true;
                        skipShowApps = true;
                        break;
                }
            }
            // If there is no error found in provider Status, then check hresult to find error.
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
        /**
         * Retrieve the RefreshRequired flag value
         */
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
        /**
         * Store the the RefreshRequired flag value
         */
        function saveRefreshRequired(refreshRequired) {
            //Note: localStorage doesn't work in IE for file:// protocol
            try {
                if (window.localStorage) {
                    window.localStorage.setItem("refreshRequired", refreshRequired);
                }
            }
            catch (e) {
            }
        }
        WefGalleryHelper.saveRefreshRequired = saveRefreshRequired;
        /**
         * Try to access the local storage and retrieve the cached store ID.
         */
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
        /**
         * addEventListener with IE 8 support.
         * Only use it for rich client's codes, or common codes used by rich client.
         * Once we deprecate IE 8, this function will be removed.
         */
        function addEventListener(element, eventName, listener) {
            // for IE 8
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
    /**
     * Class that hold the state and functions for a Add-in/Agave/Moe item inside the gallery.
     */
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
        /**
         * Insert the GalleryItem placeholder into the DOM.
         */
        GalleryItem.prototype.displayAgave = function (documentFragment) {
            // Create MOE div with spin wheel on it and index and put it on the parent div for now
            var moeDiv = document.createElement("div");
            documentFragment.appendChild(moeDiv);
            WEF.WefGalleryHelper.addClass(moeDiv, "Moe");
            // "data-ri" attribute is used by automation.
            moeDiv.setAttribute("data-ri", this.index.toString());
            moeDiv.setAttribute("role", "option");
            // Construct the MOE item in the gallery using the information from result.
            var moeInnerDiv = document.createElement("div");
            moeDiv.appendChild(moeInnerDiv);
            WEF.WefGalleryHelper.addClass(moeInnerDiv, "MoeInner");
            WEF.WefGalleryHelper.dpiScale(moeInnerDiv);
            moeInnerDiv.setAttribute("title", this.result.description);
            moeInnerDiv.setAttribute("tabindex", "-1");
            // "data-inner-moe" attribute so automation can detect click events etc.
            moeInnerDiv.setAttribute("data-inner-moe", this.index.toString());
            this.moeInnerDiv = moeInnerDiv;
            // Scale to fit the right DPI.
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
        /**
         * When the GalleryItem is visible on UI, download the image and attach necessary event handlers.
         */
        GalleryItem.prototype.updateImage = function (insertHandler) {
            var _this = this;
            if (!this.galleryItem || !this.moeInnerDiv) {
                return;
            }
            if (!this.itemCreated) {
                // When MOE is click by the user, select and deselect the items
                WEF.WefGalleryHelper.addEventListener(this.moeInnerDiv, "click", function () {
                    WEF.IMPage.selectGalleryItems(_this.index);
                });
                // When MOE is double clicked by the user, insert the app
                WEF.WefGalleryHelper.addEventListener(this.moeInnerDiv, "dblclick", function () {
                    insertHandler(_this);
                });
                // When MOE is right clicked, show the option menu
                WEF.WefGalleryHelper.addEventListener(this.moeInnerDiv, "mousedown", function (e) {
                    if (!e)
                        e = event;
                    if (e.which === 3 /*rightMouse*/ || e.button === 2 /*rightMouse*/) {
                        if (_this.appOptions) {
                            _this.appOptions.popupMenu();
                        }
                    }
                });
                // When mouse over, change the style of the MOE.
                WEF.WefGalleryHelper.addEventListener(this.moeInnerDiv, "mouseover", function () {
                    WEF.WefGalleryHelper.addClass(_this.galleryItem, "mouseover");
                    _this.appOptions.showOptionsButton();
                });
                // When mouse out, change the style back.
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
                // Create icon image.
                var img = document.createElement("img");
                tnDiv.appendChild(img);
                WEF.WefGalleryHelper.addClass(img, "MoeIcon");
                WEF.WefGalleryHelper.removeClass(tnDiv, "Tn");
                WEF.WefGalleryHelper.addClass(tnDiv, "TnNoBackGround");
                if (!agaveIconUrl || WEF.WefGalleryHelper.isHttpsUrl(window.location.href) && !WEF.WefGalleryHelper.isHttpsUrl(agaveIconUrl)) {
                    agaveIconUrl = WEF.AGAVE_DEFAULT_ICON; // When no IconUrl. or mixed content page, then use default icon
                }
                agaveIconUrl = GalleryItem.errorIconCache[agaveIconUrl] ? GalleryItem.errorIconCache[agaveIconUrl] : agaveIconUrl;
                img.onload = function () {
                    // We will scale images to fit within 32x32, by scaling down/up keeping the aspect ratio
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
                // Assign the icon url, the image will start loading. If onload event is failed, then we just do nothing and show the gray square as place holder
                img.setAttribute("src", agaveIconUrl);
                this.appOptions = WEF.IMPage.menuHandler.createAppOptions(this.result);
                var optionsButton = this.appOptions.createOptionsButton(this.index, tnDiv, img);
                if (optionsButton) {
                    this.moeInnerDiv.appendChild(optionsButton);
                }
                // Enable Narrator reading
                var arialLabelDiv = this.galleryItem;
                if (window.navigator.userAgent.indexOf("AppleWebKit") > 0) {
                    // Set aria-label attribute on InnerMoe when it is Safari or Chrome,
                    // somehow for Safari, when the arial-label is set on the Moe, it prevents arrow key from working when voiceover is on 
                    // and chrome works for setting arial-lable on either Moe or InnerMoe
                    arialLabelDiv = this.moeInnerDiv;
                }
                // set the arial-label, when it gets focus, narrator will read the string out.
                if (optionsButton) {
                    arialLabelDiv.setAttribute("aria-label", Strings.wefgallery.L_GalleryItem_Name_InsertAndOptions_Txt.replace("{0}", this.result.displayName));
                }
                else if (WEF.IMPage.currentStoreType === WEF.StoreTypeEnum.ThisDocument) {
                    arialLabelDiv.setAttribute("aria-label", this.result.displayName); // In [View all add-ins] page, there is no Insert/Options
                }
                else {
                    arialLabelDiv.setAttribute("aria-label", Strings.wefgallery.L_GalleryItem_Name_InsertOnly_Txt.replace("{0}", this.result.displayName));
                }
                // Special UI change for the app with loading error (app command scenario)
                if (this.result.hasLoadingError) {
                    var icon = document.createElement("img");
                    icon.className = "MoeErrorStatusIcon";
                    icon.src = "moe_status_icons.png";
                    tnDiv.appendChild(icon);
                    img.style.opacity = "0.5";
                }
            }
            // Mark it so no need to create it anymore.
            this.itemCreated = true;
        };
        /**
         * Set galley item relative indexs.
         */
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
/*****************************************************************************
WefGallery.ts

Web Extension Gallery Page surfaced in Office Client applications.
Code style: https://github.com/Microsoft/TypeScript/wiki/Coding-guidelines
******************************************************************************/
var WEF;
(function (WEF) {
    /**
     * Class that holds the state and functions for the WefGallery page.
     */
    var WefGalleryPage = (function () {
        function WefGalleryPage(clientFacadeCommon) {
            var _this = this;
            this.providers = {};
            /**
             * The Id of current store. In rich client, it's a string like "Omex". In WAC, it's a number in string like "0".
             */
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
            // UI HTML elements object. Make them all public and sort them in alphabetical order
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
            /** The width of a separator in the option bar (excluding margin) */
            this.menuSeparatorWidth = null;
            /** The max possible width of the menu right in the option bar. */
            this.menuRightMaxPossibleWidth = null;
            // Gallery status controlling variables
            this.galleryItems = null;
            this.uiState = { "Ready": false, "StoreIdBeforeReady": "", "ErrorBeforeReady": "", "ErrorLinkTextBeforeReady": "", "ErrorLinkHandlerBeforeReady": null };
            // currentIndex: it is current index of the gallery items that can be selected or de-selected, and it is set by keyboard arrows or tab keys;
            // -1 means no gallery item can be selected/de-selected triggered by spacebar
            this.currentIndex = -1;
            this.currentTabIndex = -1;
            this.results = null;
            this.height = "100%";
            this.width = "100%";
            this.itemsPerRow = null;
            // Gallery keyboard handlers
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
            // Performance measuring varaibles
            this.totalSessionTime = 0;
            this.trustPageSessionTime = 0;
            // Configuration and flightings from the client
            this.envSetting = {};
            this.isUploadFileDevCatalogEnabled = false;
            this.isAppCommandEnabled = false;
            this.moveLeft = function (event, eventTarget) {
                // Handle keyevent when Tabpanel item has focus
                if (WEF.WefGalleryHelper.hasClass(eventTarget, "TabATag")) {
                    var targetTabIndex = _this.currentTabIndex - 2; // Skip over text separator '|'
                    if (targetTabIndex < 0) {
                        targetTabIndex = _this.tabs.childNodes.length - 1; // Loop
                    }
                    if (targetTabIndex != _this.currentTabIndex) {
                        var targetTab = _this.tabs.childNodes[targetTabIndex];
                        _this.toggleTabSelection(targetTab, null /*calback*/);
                    }
                }
                else {
                    // Handle keyevent when gallery list item has focus
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
                // Handle keyevent when Tabpanel item has focus
                if (WEF.WefGalleryHelper.hasClass(eventTarget, "TabATag")) {
                    var targetTabIndex = _this.currentTabIndex + 2; // Skip over text separator '|'
                    if (targetTabIndex > _this.tabs.childNodes.length - 1) {
                        targetTabIndex = 0; // Loop
                    }
                    if (targetTabIndex != _this.currentTabIndex) {
                        var targetTab = _this.tabs.childNodes[targetTabIndex];
                        _this.toggleTabSelection(targetTab, null /*calback*/);
                    }
                }
                else {
                    // Handle keyevent when gallery list item has focus
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
                // Handle keyevent only for gallery list item
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
                // Handle keyevent only for gallery list item
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
                // Select first gallery list item if none is selected on tab from tab panel item
                if (!event.shiftKey && (element == _this.tabs.childNodes[_this.currentTabIndex] || element == _this.tabs.childNodes[_this.currentTabIndex].firstChild) && event.preventDefault && _this.currentIndex < 0 && _this.galleryItems && _this.galleryItems.length > 0) {
                    _this.currentIndex = 0;
                    _this.selectGalleryItems(_this.currentIndex, false);
                    event.preventDefault();
                }
                // OM 1113515: a TAB-loop is created here to prevent tab exiting the insertion dialog.
                // If BUG 438605 is resolved, we don't need this process any more.
                // Put focus on selected tab panel item when focus is on refresh link and vice-versa when shift-tab is used
                if (!event.shiftKey && element.getAttribute("id") == "RefreshInner" && event.preventDefault && _this.tabs && _this.currentTabIndex >= 0 && _this.currentTabIndex < _this.tabs.childNodes.length) {
                    _this.tabs.childNodes[_this.currentTabIndex].firstChild.focus();
                    event.preventDefault();
                }
                if (event.shiftKey && _this.tabs && (element == _this.tabs.childNodes[_this.currentTabIndex] || element == _this.tabs.childNodes[_this.currentTabIndex].firstChild) && event.preventDefault && _this.refreshATag) {
                    _this.refreshATag.focus();
                    event.preventDefault();
                }
            };
            /**
             * Handle key down events.
             */
            this.galleryKeyDownHandler = function (e) {
                var numOfItems = 0;
                if (_this.results) {
                    numOfItems = _this.results.length;
                }
                if (!e)
                    e = event;
                // Give other handlers a chance to handle the keystroke before defaulting to generic handler
                for (var i = 0; i < _this.keyHandlers.length; i++) {
                    var keyHandler = _this.keyHandlers[i];
                    if (keyHandler.handleKeyDown(e)) {
                        e.stopPropagation();
                        e.preventDefault();
                        return;
                    }
                }
                var eventTarget = e.srcElement ? e.srcElement : e.target; // e.srcElement is for IE 8
                switch (e.keyCode) {
                    case 9:
                        _this.tabKeyHandler(e, eventTarget);
                        break;
                    case 13:
                        // Keep the target element of key down event, then when key up event is triggered,
                        // use this kept target element to decide whether to execute key up handlers.
                        // Besides that, prevent browser's default action for this enter key down event.
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
            /**
             * Handle key up events.
             */
            this.galleryKeyUpHandler = function (e) {
                if (!e)
                    e = event;
                // Give other handlers a chance to handle the keystroke before defaulting to generic handler
                for (var i = 0; i < _this.keyHandlers.length; i++) {
                    var keyHandler = _this.keyHandlers[i];
                    if (keyHandler.handleKeyUp(e)) {
                        e.stopPropagation();
                        e.preventDefault();
                        return;
                    }
                }
                var eventTarget = e.srcElement ? e.srcElement : e.target; // e.srcElement is for IE 8
                switch (e.keyCode) {
                    case 13:
                        _this.executeButtonCommand(eventTarget, e);
                        break;
                }
            };
            /**
             * We need to resize the gallery area when the dialog is resized.
             */
            this.resizeHandler = function () {
                _this.uiState.Ready = false;
                // In IE, resize is fired when any item on the page is resized. We limit the resize event to
                // when the window size is changed.
                var winHeight = WEF.WefGalleryHelper.getWinHeight().toString();
                var winWidth = WEF.WefGalleryHelper.getWinWidth().toString();
                if (_this.height != winHeight || _this.width != winWidth) {
                    _this.height = winHeight;
                    _this.width = winWidth;
                    _this.setGalleryHeight();
                    _this.delayLoadVisibleImages();
                    // Adjust maxWidth of buttons
                    var newMaxWidth, widthIncreaseRatio = (_this.width) / WEF.UI.DefaultGalleryWidth;
                    newMaxWidth = WEF.WefGalleryHelper.getDPIXScaledNumber(WEF.UI.DefaultDialogBtnMaxWidth) * widthIncreaseRatio;
                    _this.btnAction.style.maxWidth = newMaxWidth + "px";
                    _this.btnCancel.style.maxWidth = newMaxWidth + "px";
                    _this.btnTrustAll.style.maxWidth = newMaxWidth + "px";
                    _this.btnDone.style.maxWidth = newMaxWidth + "px";
                    _this.setOptionBarElementMaxSize(_this.tabTitles);
                    // Resize the description for the selected item.
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
            /**
             * Load the visible images.
             */
            this.loadVisibleImages = function () {
                if (new Date().getTime() - _this.delayTime < _this.delayLoad && _this.delaying) {
                    // if not enough time has passed since this method was called, reset the timer.
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
                            if (_this.currentIndex < itemsPerRow && _this.keyCodePressed == 40 /* down arrow pressed */) {
                                _this.gallery.scrollTop = 0; // down arrow is used to access the gallery items, set the scrollTop to 0 if it is row 0 to avoid scroll to the next row
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
                    // if gallery.children.length==0, wait for some more time.
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
            /*****************************************************************************
             Abstract handlers that are going to be implemented in the subclasses for native client and WAC.
            *****************************************************************************/
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
            // Inject the envrionment settings
            this.envSetting = this.clientFacadeCommon.getEnvSetting();
            this.isAppCommandEnabled = this.envSetting["IsAppCommandEnabled"] === true;
        }
        /**
         * Control the visibility of the menu in the top-right of the gallery
         */
        WefGalleryPage.prototype.showHideRightMenuButtons = function (showManageApp, showRefresh) {
            this.menuRight.style.display = "block";
            var hideRightMenu = !showManageApp && !showRefresh;
            var showUploadAddin = !hideRightMenu && !showManageApp && this.isUploadFileDevCatalogEnabled; // Upload addin menu is the replacement for mangage app menu
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
        /**
         * Show the footer and adjust its height properly.
         */
        WefGalleryPage.prototype.showFooter = function () {
            this.footer.style.visibility = 'visible';
            this.footer.style.height = WEF.WefGalleryHelper.getDPIYScaledNumber(WEF.UI.DefaultFooterHeight) + "px";
        };
        /**
         * Control the visibility of the buttons in the bottom
         */
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
        /**
         * Get HResult for the Provider of the current storeid.
         */
        WefGalleryPage.prototype.getCurrentProviderHResult = function () {
            var hres = 0;
            if (this.currentStoreId) {
                hres = this.providers[this.currentStoreId][2];
            }
            return hres;
        };
        /**
         * Get status for the Provider of the current storeid.
         */
        WefGalleryPage.prototype.getCurrentProviderStatus = function () {
            var status = 0;
            if (this.currentStoreId) {
                status = this.providers[this.currentStoreId][1];
            }
            return status;
        };
        /**
         * Get the Featured Item page url.
         */
        WefGalleryPage.prototype.getFeaturedPageUrl = function () {
            if (!this.currentPageUrl) {
                this.currentPageUrl = this.getPageUrl(WEF.PageTypeEnum.Recommendation);
            }
            return this.currentPageUrl;
        };
        /**
         * Get the Featured Item page url.
         */
        WefGalleryPage.prototype.getLandingPageUrl = function () {
            if (!this.landingPageUrl) {
                this.landingPageUrl = this.getPageUrl(WEF.PageTypeEnum.Landing);
            }
            return this.landingPageUrl;
        };
        /**
         * Get the Featured Item page url.
         */
        WefGalleryPage.prototype.getAppManagePageUrl = function () {
            if (!this.appManagePageUrl) {
                this.appManagePageUrl = this.getPageUrl(WEF.PageTypeEnum.ManageApps);
            }
            return this.appManagePageUrl;
        };
        /**
         * Handle enter key events
         */
        WefGalleryPage.prototype.executeButtonCommand = function (element, event) {
            this.menuHandler.hideMenu(true);
            if (element != this.enterKeyTarget) {
                return;
            }
            if (WEF.WefGalleryHelper.hasClass(element, "MoeInner") || WEF.WefGalleryHelper.hasClass(element, "Moe")) {
                // insert the MOE when a MOE item is selected.
                this.insertSelectedItem();
            }
            else if (WEF.WefGalleryHelper.hasClass(element, "TabATag")) {
                var storeId = element.parentElement.getAttribute("data-storeId");
                if (storeId) {
                    this.toggleTabSelection(element.parentElement, null /*calback*/);
                }
                else {
                    this.showEntitlements(this.currentStoreId, true, null /*calback*/);
                }
            }
            else if (element.getAttribute("id") == "BtnAction") {
                // Insert button
                this.insertSelectedItem();
            }
            else if (element.getAttribute("id") == "BtnCancel" || element.getAttribute("id") == "BtnDone") {
                // Cancel button
                this.cancelDialog();
            }
            else if (element.getAttribute("id") == "ManageInner") {
                // Manage my apps button
                this.launchAppManagePage();
            }
            else if (element.getAttribute("id") == "RefreshInner") {
                // Refresh button
                this.showEntitlements(this.currentStoreId, true, null /*calback*/);
            }
            else if (element.getAttribute("id") == "FooterLinkATag") {
                // Footer link
                this.gotoStore();
            }
            else if (element.getAttribute("id") == "linkId") {
                // Signin link
                this.invokeSignIn();
            }
            else if (element.getAttribute("id") == "rateReviewLink") {
                if (this.results != null && this.results.length > 0) {
                    WEF.IMPage.invokeWindowOpen(this.results[0].rateReviewUrl);
                }
            }
            else if (WEF.WefGalleryHelper.hasClass(element, "OptionsButton")) {
                // Gallery item option button
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
        /**
         * toggle between tab buttons
         */
        WefGalleryPage.prototype.toggleTabSelection = function (selectedTabDiv, callback) {
            this.cleanUpGallery();
            var selectedTabId = selectedTabDiv.getAttribute("id");
            var len = this.tabs.childNodes.length, i, child, tabId;
            for (i = 0; i < len; i++) {
                child = this.tabs.childNodes[i];
                if (child.attributes && WEF.WefGalleryHelper.hasClass(child, "TextNav")) {
                    WEF.WefGalleryHelper.removeClass(child.firstChild, "TabSelected");
                    // Make Tab panel item non-tabbable
                    child.setAttribute("tabIndex", "-1");
                    child.firstChild.setAttribute("aria-selected", "false");
                    child.firstChild.removeAttribute("aria-controls");
                    tabId = child.getAttribute("id");
                    if (tabId == selectedTabId) {
                        // Make selected tab panel item part of tab-loop
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
                            this.currentIndex = -1; // reset the index when switch tabs
                        }
                        this.currentStoreId = storeId;
                        this.currentStoreType = storeType;
                        this.saveStoreId(this.currentStoreId);
                        if (storeId && storeType != WEF.StoreTypeEnum.Recommendation) {
                            //Show footer and upper right menu
                            this.restoreFooterLink();
                            this.showFooter();
                            this.showEntitlements(storeId, false, callback);
                            this.setGalleryHeight();
                        }
                        else {
                            // Show Featured page.  
                            var pageUrl = child.getAttribute("data-PageUrl");
                            this.showContentPage(pageUrl);
                        }
                        // Update the tooltip of  Refresh button with the tab name, e.g. Refresh 'My orgagnization'/ Refresh 'My AddIns'
                        this.refreshATag.setAttribute("title", Strings.wefgallery.L_WefDialog_RefreshButton_Tooltip.replace("{0}", child.firstChild.textContent));
                    }
                }
            }
        };
        /**
         * Intialize Gallery UI: Creating tabs for the given providers, hooking up events and setting up tab order
         */
        WefGalleryPage.prototype.initializeGalleryUI = function (providers, resetToMarketPlace) {
            var _this = this;
            // Do nothing when there is no providers.
            if (providers == undefined || providers.length === 0) {
                return false;
            }
            // Make providersArray for tab construction.
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
                        // If OneDrive then don't add now, add provider at the end
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
                // Put recommendaton featured tab the last one
                providersArray.push([WEF.PageStoreId.Recommendation, WEF.StoreTypeEnum.Recommendation, 0, 0]);
                this.footerLink.style.display = 'block';
            }
            else {
                this.footerLink.style.display = 'none';
            }
            // If OneDriveCatalog is added then make it the last tab. This is available to developers only.
            if (hasOneDriveCatalogProvider) {
                providersArray.push([WEF.StoreTypeEnum.OneDrive, WEF.StoreTypeEnum.OneDrive, 0, 0]);
            }
            len = providersArray.length;
            if (len === 0) {
                // No valid provider
                return false;
            }
            // Set Current
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
            // Clean up tabs
            while (this.tabs.hasChildNodes()) {
                this.tabs.removeChild(this.tabs.firstChild);
            }
            if (!isCurrentSet) {
                // The first provider is selected.
                this.currentStoreId = providersArray[0][0];
                this.currentStoreType = providersArray[0][1];
            }
            // Construct tabs
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
                // {StoreId, [StoreType, Status, HResult]}
                this.providers[tempStoreId] = [tempStoreType, tempStatus, tempHResult];
                var tabName = WEF.storeTypes[tempStoreType]; // The tabNames of storeTypes
                if (tabName) {
                    delete WEF.storeTypes[tempStoreType]; // Current design is that we only show the first store of the same storeType.
                    tabOrder++;
                    if (tempStoreId === WEF.StoreTypeEnum.OneDrive) {
                        // Special case for OneDrive provider, check if user has consented for app and can get access token.
                        this.checkAndCreateOneDriveProviderTab(this.tabs, tabOrder, tabName, tempStoreId, tempStoreType);
                    }
                    else {
                        // For all other providers create Tab
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
            // Set current selected tab
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
            // Make selected tab panel item part of tab-loop
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
            // hook up refresh button.
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
            // TODO: The localization tag was added after the localization deadline. We can remove this after SP1. The localized htm page will contain the localized string.
            this.documentAppsMsg.firstChild.innerText = Strings.wefgallery.L_TrustUx_AppsMessage;
            this.readMoreATag.setAttribute("tabIndex", "0");
            this.readMoreATag.setAttribute("title", Strings.wefgallery.L_TrustUx_ReadMoreLink_Txt_Tooltip);
            this.readMoreATag.setAttribute("role", "link");
            this.permissionATag.setAttribute("tabIndex", "0");
            this.permissionATag.setAttribute("title", Strings.wefgallery.L_Permission_Link_Txt_Tooltip);
            this.permissionATag.setAttribute("role", "link");
            this.permissionTextAndLink.setAttribute("title", Strings.wefgallery.L_Permission_Link_Txt_Tooltip);
            this.btnAction.setAttribute("tabIndex", "0");
            // When app command is enabled, the text in action button is different.
            if (this.isAppCommandEnabled) {
                this.btnAction.value = Strings.wefgallery.L_OK_Button_Txt;
                this.btnAction.title = Strings.wefgallery.L_OK_Button_Txt_Tooltip;
            }
            else {
                // The value of the action button is defined in the HTML page
                this.btnAction.title = Strings.wefgallery.L_Action_Button_Txt_Tooltip;
            }
            this.btnCancel.setAttribute("tabIndex", "0");
            this.btnCancel.setAttribute("title", Strings.wefgallery.L_Cancel_Button_Text_Tooltip);
            this.btnTrustAll.setAttribute("tabIndex", "0");
            this.btnTrustAll.setAttribute("title", Strings.wefgallery.L_TrustAll_Button_Txt_Tooltip);
            this.btnDone.setAttribute("tabIndex", "0");
            this.btnDone.setAttribute("title", Strings.wefgallery.L_Done_Button_Txt_Tooltip);
            // Initialize Hero message and button
            this.noAppsMessage.setAttribute("title", Strings.wefgallery.L_OfficeStore_Button_Tooltip.replace("{0}", this.officeStoreBtn.value));
            this.noAppsMessage.style.marginTop = WEF.WefGalleryHelper.getDPIXScaledNumber(WEF.UI.HeroMessageMarginTop) + "px";
            this.noAppsMessageTitle.innerHTML = Strings.wefgallery.L_NoAppsMessageTitle;
            this.noAppsMessageText.innerHTML = Strings.wefgallery.L_NoAppsMessageText;
            this.officeStoreBtn.style.width = WEF.WefGalleryHelper.getDPIXScaledNumber(WEF.UI.HeroBtnWidth) + "px";
            this.officeStoreBtn.style.height = WEF.WefGalleryHelper.getDPIXScaledNumber(WEF.UI.HeroBtnHeight) + "px";
            this.overrideButtonTooltip();
            return true;
        };
        /**
         * Show proper content based on current store Id
         */
        WefGalleryPage.prototype.showContent = function (forceRefresh) {
            if (this.currentStoreId == WEF.PageStoreId.Recommendation) {
                this.showContentPage(this.currentPageUrl);
            }
            else {
                this.showEntitlements(this.currentStoreId, forceRefresh, null /*callback*/);
            }
        };
        /**
         * Store the storeId for current active tab
         */
        WefGalleryPage.prototype.saveStoreId = function (currentStoreId) {
            // Note: localStorage doesn't work in IE for file:// protocol
            try {
                if (window.localStorage) {
                    window.localStorage.setItem("lastActiveStoreId", encodeURI(currentStoreId));
                }
            }
            catch (e) {
            }
        };
        /**
         * No narrator announcement on presentation only control (non-functional control such as separator |)
         * @param ctl
         */
        WefGalleryPage.prototype.disableNarratorOnControl = function (ctl) {
            ctl.setAttribute("role", "presentation");
            ctl.setAttribute("aria-hidden", "true");
            ctl.setAttribute("tabindex", "-1");
        };
        /**
         * Create a Tab Div with the specified tabName, storeId, storeType, pageUrl
         */
        WefGalleryPage.prototype.createTab = function (tabsDiv, tabOrder, tabName, storeId, storeType) {
            var me = this;
            if (tabsDiv.childNodes.length != 0) {
                // Add a separator if there is existing tab
                var separatorDiv = document.createElement('div');
                WEF.WefGalleryHelper.addClass(separatorDiv, "separator");
                separatorDiv.innerHTML = "|";
                this.disableNarratorOnControl(separatorDiv);
                tabsDiv.appendChild(separatorDiv);
            }
            var pageUrl = WEF.PageStoreId.Recommendation === storeId ? this.getFeaturedPageUrl() : null;
            var tabDiv = document.createElement('div');
            WEF.WefGalleryHelper.addClass(tabDiv, "TextNav");
            //tabDiv.style.maxWidth = WEF.WefGalleryHelper.getDPIXScaledNumber(UI.DefaultTabMaxWidth) + "px";
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
            tabDiv.onclick = function WEF_WefGalleryPage_createTab_tabDiv$onclick() { me.toggleTabSelection(this, null /*callback*/); };
            // Always put focus on link to make tab and link behave as single control
            tabDiv.onfocus = function WEF_WefGalleryPage_createTab_tabDiv$onfocus() {
                aTag.focus();
            };
            return tabDiv;
        };
        /**
         * When the gallery has been scrolled, we need to load images.
         */
        WefGalleryPage.prototype.galleryScrollHandler = function () {
            this.menuHandler.hideMenu(true);
            this.delayLoadVisibleImages();
        };
        /**
         * Store the size of some elements when they are visible right after the page is loaded.
         */
        WefGalleryPage.prototype.storeStaticElementRealSize = function () {
            this.menuSeparatorWidth = WEF.UI.DefaultSeparatorWidth;
            if (this.menuRightSeparatorDiv.offsetWidth != 0) {
                this.menuSeparatorWidth = this.menuRightSeparatorDiv.offsetWidth;
            }
            var uploadMenuWidth = 0;
            if (this.isUploadFileDevCatalogEnabled) {
                uploadMenuWidth = this.uploadMenuDiv.offsetWidth;
            }
            this.menuRightMaxPossibleWidth = Math.max(uploadMenuWidth, this.manageMenuDiv.offsetWidth) +
                this.menuSeparatorWidth +
                this.refreshMenuDiv.offsetWidth +
                WEF.WefGalleryHelper.getDPIXScaledNumber(WEF.UI.OptionBarElementMargin) * 3; // Don't forget the left margin in all elements
        };
        /**
         * Set a max width for the elements in the option bar when overflow happens.
         */
        WefGalleryPage.prototype.setOptionBarElementMaxSize = function (tabTitles) {
            /*
             The structure of the option bar:
            | (DefaultLeftMargin)  |  (Menu Left)                                 (OptionBarMenuGap)        (Menu Right)                                       |  (DefaultRightPadding)  |
            |  <-    26 px     ->  |  TAB 1 ~ | ~ TAB 2 ~ | ~ ... ~ | ~ TAB n ~   <-     20 px    ->    ~ Upload Menu or Management Menu ~ | ~ Refresh button  |  <-      25 px      ->  |
            Note: ~ represents a OptionBarElementMargin, whose width is 7px.
            */
            if (tabTitles == null || tabTitles.length == 0)
                return;
            // Reset the max width for all the tabs because we want to get their real offsetWidth.
            for (var i = 0; i < tabTitles.length; i++) {
                tabTitles[i].style.maxWidth = "none";
            }
            this.refreshMenuDiv.style.maxWidth = "none";
            this.uploadMenuDiv.style.maxWidth = "none";
            this.manageMenuDiv.style.maxWidth = "none";
            /** The width of all elements and margins in the option bar */
            var optionBarTotalWidth = WEF.WefGalleryHelper.getDPIXScaledNumber(WEF.UI.DefaultLeftMargin) +
                this.tabs.offsetWidth +
                WEF.WefGalleryHelper.getDPIXScaledNumber(WEF.UI.OptionBarMenuGap) +
                this.menuRightMaxPossibleWidth +
                WEF.UI.DefaultRightMargin;
            // Only set the max width when overflow happens based on the calculation
            if (optionBarTotalWidth > WEF.WefGalleryHelper.getWinWidth()) {
                /** The usable space for text for all elements in option bar */
                var widthForAllTitleText = WEF.WefGalleryHelper.getWinWidth() - WEF.WefGalleryHelper.getDPIXScaledNumber(WEF.UI.DefaultLeftMargin) -
                    WEF.UI.DefaultRightMargin - WEF.WefGalleryHelper.getDPIXScaledNumber(WEF.UI.OptionBarMenuGap);
                // There are (n-1) separators if there are n tabs, and 1 separator in the right menu. Exclude them.
                widthForAllTitleText -= this.menuSeparatorWidth * tabTitles.length;
                // There are (n * 2 - 1) margins if there are n tabs, and 3 margins in right menus. Exclude them.
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
        /**
         * We set the gallery height, leaving room for the provider and refresh button on top,
         * and the insert and cancel button on the bottom.
         */
        WefGalleryPage.prototype.setGalleryHeight = function () {
            // The height of the Gallery is Windows height - header height - footer height.
            var galleryContainerHeight = WEF.WefGalleryHelper.getWinHeight() - this.header.offsetHeight - this.footer.offsetHeight;
            if (this.galleryContainer && galleryContainerHeight > 0 &&
                (galleryContainerHeight != this.galleryContainer.offsetHeight || this.footer && this.footer.style.top === "")) {
                this.galleryContainer.style.height = galleryContainerHeight + "px";
                // Put GalleryContainer right below Header div.
                this.galleryContainer.style.top = this.header.offsetHeight + "px";
                var galleryHeight = galleryContainerHeight;
                if (this.currentStoreType == WEF.StoreTypeEnum.ThisDocument) {
                    galleryHeight = galleryHeight - this.documentAppsMsg.offsetHeight * 2;
                }
                this.gallery.style.height = galleryHeight + "px";
                // Put footer div at the right position
                var footerTop = galleryContainerHeight + this.header.offsetHeight;
                this.footer.style.top = footerTop + "px"; //current code assumes Footer existence in DOM
            }
        };
        /**
         * We set the width of selectedItem and its children, when the UI is resized.
         */
        WefGalleryPage.prototype.setSelectedItemWidth = function () {
            var newWidth = WEF.WefGalleryHelper.getWinWidth() - WEF.WefGalleryHelper.getDPIXScaledNumber(WEF.UI.SelectedItemDesciptionWidthAdjustment);
            if (this.currentStoreType == WEF.StoreTypeEnum.ThisDocument) {
                newWidth = newWidth - this.btnTrustAll.offsetWidth - this.btnDone.offsetWidth;
            }
            else {
                newWidth = newWidth - this.btnAction.offsetWidth - this.btnCancel.offsetWidth;
            }
            // Adjust the width of selected item description to fit the right DPI
            this.selectedItem.style.width = newWidth + "px";
            this.selectedItem.style.height = WEF.WefGalleryHelper.getDPIYScaledNumber(WEF.UI.DefaultSelectedItemHeight) + "px";
            // Adjust the maxWidth of description text based on the width of UI
            var marginLeft = parseInt(window.getComputedStyle ? window.getComputedStyle(this.selectedItem).marginLeft : this.selectedItem.style.marginLeft);
            this.selectedDescriptionText.style.maxWidth = (newWidth - marginLeft - this.selectedDescriptionReadMoreLink.offsetWidth) + "px";
            this.footerLink.style.width = newWidth + "px";
        };
        /**
         *  Change the footer bar info and disable the action button when there is no gallery item is selected.
         */
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
        /**
         * Clean up any child elements in the Gallery area and clean up the any cached GalleryItems
         */
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
            this.header.style.height = WEF.WefGalleryHelper.getDPIYScaledNumber(WEF.UI.DefaultHeaderHeight) + "px"; // 62px is the default height of header. This will make sure notification is not shown.
            this.setGalleryHeight();
            this.trustPageSessionTime = 0;
        };
        /**
         * Process and show the entitlements retrieved from the provider inside the gallery.
         */
        WefGalleryPage.prototype.processResults = function (results) {
            // TODO: to reuse the GalleryItems already retrieved and constructed for the providers. This could be a good perf improvement.
            // clean up the old data
            this.results = null;
            // If an error occurred on the server side, do not attempt to process results.
            if (results == null) {
                return;
            }
            this.results = results;
            // Construct and display enititlements in the gallery.
            this.galleryItems = new Array(results.length);
            for (var i = 0; i < results.length; i++) {
                this.galleryItems[i] = new WEF.GalleryItem(results[i], i);
                this.galleryItems[i].displayAgave(this.gallery);
            }
            // Start loading images.
            this.delayLoadVisibleImages();
        };
        /**
         * Check whether there is any add-ins marked with "has loading error".
         * If any, show the error message bar.
         */
        WefGalleryPage.prototype.processAddinLoadingErrors = function (results) {
            for (var i = 0; i < results.length; i++) {
                if (results[i].hasLoadingError) {
                    this.showError(Strings.wefgallery.L_AddinsHasLoadingErrors, this.currentStoreId);
                    break;
                }
            }
        };
        /**
         * Used to throttle events that may be called multiple times in quick succession such as
         * Scrolling and Resizing. The delayFunction must either reset the delay timer or execute
         * the method and set this.delaying to false.
         */
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
        /**
         * Calculate the items in the row by iterating through the first several. This is used
         * for keyboard navigation and displaying the visible images.
         */
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
                // This fix the issue when UI is not ready the offsetLeft is always 0, the default itemsPerRow is set to 3.
                if (item.offsetLeft == 0) {
                    itemsPerRow = 3;
                    break;
                }
                if (WEF.WefGalleryHelper.getHTMLDir() == "ltr") {
                    var left = Math.abs(item.offsetLeft - defaultMargin); //left margin to the edge of the gallery.
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
            // Set the variable for keyboard handling.
            this.itemsPerRow = itemsPerRow;
            return itemsPerRow;
        };
        /**
         * Rich Client and WAC code share the same code for showContentPage action.
         */
        WefGalleryPage.prototype.showContentPage = function (pageUrl) {
            var _this = this;
            // Hide footer and re-adjust footer height and gallery height.
            this.footer.style.visibility = 'hidden';
            this.documentAppsMsg.style.display = 'none';
            this.footer.style.height = WEF.WefGalleryHelper.getDPIYScaledNumber(WEF.UI.HiddenFooterHeight) + "px";
            this.setGalleryHeight();
            // Hide upper right menu.
            this.showHideRightMenuButtons(false, false);
            if (pageUrl && pageUrl != "") {
                this.gallery.style.overflowY = "hidden"; // No scolling bar is needed. It will be handled by the iframe itself.
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
                // Pass the origin for the Recommendation page to call back.
                pageUrl += "#" + window.location.href;
                frame.setAttribute("src", pageUrl);
                this.gallery.appendChild(frame);
            }
            else {
                this.showError(Strings.wefgallery.L_NoFeaturedItemsError, WEF.PageStoreId.Recommendation);
            }
        };
        /**
         * Remove a single gallery item, and update the remaining items' index.
         * @param index CurrentIndex passed as a parameter
         */
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
                    // after remove an item, set foucs on the next item 
                    if (this.galleryItems.length >= 1) {
                        var indexToFocus = index;
                        //if item deleted is the last item, set focus on the first item 
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
        /**
         * Select and toggle to user selected gallery item.
         * Default behavior: if sepecified item is selected, it will be toggled to "unselected".
         * @param index The item to be selected/toggled
         * @param forceSelected Force to select the specified item instead of toggle
         */
        WefGalleryPage.prototype.selectGalleryItems = function (index, forceSelected) {
            if (forceSelected === void 0) { forceSelected = false; }
            var result = this.results[index];
            var len = this.galleryItems ? this.galleryItems.length : 0;
            this.currentIndex = -1;
            for (var i = 0; i < len; i++) {
                var item = this.galleryItems[i];
                if (index == i) {
                    this.currentIndex = index;
                    // Deselect the current selected item.
                    if (WEF.WefGalleryHelper.hasClass(item.galleryItem, "selected")) {
                        if (forceSelected == false) {
                            WEF.WefGalleryHelper.removeClass(item.galleryItem, "selected");
                            item.galleryItem.removeAttribute("aria-selected");
                            // Remove unselected gallery list item from tab-loop
                            item.galleryItem.setAttribute("tabIndex", "-1");
                            this.currentIndex = -1;
                            this.deSelectBtnAction();
                        }
                    }
                    else {
                        // Select the user newly selected item.
                        WEF.WefGalleryHelper.addClass(item.galleryItem, "selected");
                        WEF.WefGalleryHelper.setHtmlEncodedText(this.selectedDescriptionText, result.description);
                        this.selectedDescriptionText.setAttribute("title", result.description);
                        this.selectedItem.style.display = 'block';
                        this.footerLink.style.display = 'none';
                        if (this.currentStoreType != WEF.StoreTypeEnum.ThisDocument) {
                            WEF.WefGalleryHelper.removeClass(this.btnAction, 'disabled');
                            this.btnAction.removeAttribute('disabled');
                        }
                        // Make selected gallery list item part of tab-loop
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
                    // Make sure all other items are unselected.
                    this.unselectGalleryItems(item);
                }
            }
        };
        /**
         * Unselect and toggle to the gallery item to be unselected.
         * @param item The item to be unselected
         */
        WefGalleryPage.prototype.unselectGalleryItems = function (item) {
            if (item && item.galleryItem) {
                WEF.WefGalleryHelper.removeClass(item.galleryItem, "selected");
                item.galleryItem.removeAttribute("aria-selected");
                // Remove unselected gallery list item from tab-loop
                item.galleryItem.setAttribute("tabIndex", "-1");
                if (item.appOptions && item.galleryItem.querySelector(":hover") == null) {
                    item.appOptions.hideOptionsButton();
                }
            }
        };
        /**
         * Shows the no apps error message
         */
        WefGalleryPage.prototype.showNoAppsError = function () {
            this.gallery.innerHTML = "";
            if (this.currentStoreType === WEF.StoreTypeEnum.MarketPlace) {
                // Show no apps message and the button for Office Store link.
                this.noAppsMessage.style.display = 'block';
                this.gallery.appendChild(this.noAppsMessage);
                this.officeStoreBtn.focus();
                // Hide footer and Manage My App.  
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
                // Adjust notification height according to the height of the message.
                var notificationHeight = this.errorMessage.scrollHeight + WEF.UI.AdjustNotificationHeight;
                document.getElementById("Notification").style.height = notificationHeight + "px";
                // Show notification by increase the height of Header
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
            // If storeId is provided including "", then it must be the same as the current storeId
            // in order to show the error. Otherwise the error is not for the current 
            // tab hence return without showing the error
            if ((storeId || storeId === "") && storeId != this.currentStoreId || !messageStr) {
                return;
            }
            // Remove Spin wheel if there is one.
            if (this.gallery && this.gallery.firstChild && WEF.WefGalleryHelper.hasClass(this.gallery.firstChild, "SpinWheel")) {
                this.gallery.removeChild(this.gallery.firstChild);
            }
            this.gallery.setAttribute("aria-busy", "false");
            // Show Error message or Show Error message, plus a Linked button with event handler
            if (arguments.length < 4) {
                this.showErrorInternal(messageStr);
            }
            else {
                this.showErrorInternal(messageStr, linkedMessageStr, linkedCallback, showCloseButton);
            }
        };
        /**
         * Goto the store tab.
         */
        WefGalleryPage.prototype.gotoStore = function () {
            this.toggleTabSelection(this.storeTab, null /*callback*/);
        };
        /**
         * Empty method that is going to be implemented in WefGalleryRichOutlook.js.
         */
        WefGalleryPage.prototype.overrideButtonTooltip = function () {
            // Nothing happens. This is only implemented in Outlook rich client.
        };
        WefGalleryPage.prototype.getPageUrl = function (pageType) {
            var pageUrl = this.clientFacadeCommon.getPageUrl(pageType);
            if (pageUrl == "" && pageType == WEF.PageTypeEnum.Recommendation) {
                this.showError(Strings.wefgallery.L_NoFeaturedItemsError, WEF.PageStoreId.Recommendation);
            }
            return pageUrl;
        };
        /**
         * Insert the selected agave in the gallery to the host.
         */
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
        /*****************************************************************************
         Abstract methods that have empty implementation by default
        *****************************************************************************/
        /**
         * Determine whether insertion operation is allowed in current condition
         */
        WefGalleryPage.prototype.allowInsertion = function () {
            return true;
        };
        /**
        * Check the status of OneDrive provider and create the tab if it's available
        */
        WefGalleryPage.prototype.checkAndCreateOneDriveProviderTab = function (oneDriveTabs, oneDriveTabOrder, oneDriveTabName, oneDriveStoreId, oneDriveStoreType) {
        };
        /**
         * Setting up page level and hooking up events
         */
        WefGalleryPage.prototype.wefGalleryAppOnLoad = function () {
            var _this = this;
            // Intialize this.gallery
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
            this.notification.setAttribute("role", "alert"); // This makes narrator to announce it when this control changes from invisible to visible.
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
            this.officeStoreBtn.title = Strings.wefgallery.L_OfficeStore_Button_NoAddIns_Tooltip; // tooltip for narrator to announce
            this.manageATag = document.getElementById('ManageInner');
            this.uploadATag = document.getElementById('UploadMenuInner');
            this.uploadMenuDiv = document.getElementById('UploadMenu');
            this.manageMenuDiv = document.getElementById('Manage');
            this.refreshMenuDiv = document.getElementById('Refresh');
            this.menuRightSeparatorDiv = document.getElementById("MenuRightSeparator");
            var optionsDiv = document.getElementById('Options');
            // Store the size of visible elements in the page.
            this.storeStaticElementRealSize();
            // Scale to fit the right DPI
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
            // Hide the menu right bar after its children elements are prepared.
            this.menuRight.style.display = "none";
            // Setting up InsertGallery and hooking up events.
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
            // UI that implement HandleKey
            this.keyHandlers = [this.menuHandler, this.modalDialog];
            // Handle key stroke
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
    /**
     * Constructing the WefGallery page by setting up necessary objects.
     * Different clients will have its special implementation.
     */
    WEF.setupClientSpecificWefGalleryPage = null;
    /**
     * The first function to call in page load event.
     * It will triggered by the "onload" event of <body> element in HTML page.
     */
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
        /**
         * Operation code used for telemetry for app management menu interactions
         */
        var AppManagementAction;
        (function (AppManagementAction) {
            AppManagementAction[AppManagementAction["Cancel"] = 0] = "Cancel";
            AppManagementAction[AppManagementAction["AppDetails"] = 1] = "AppDetails";
            AppManagementAction[AppManagementAction["RateReview"] = 2] = "RateReview";
            AppManagementAction[AppManagementAction["Remove"] = 3] = "Remove";
        })(AppManagementAction || (AppManagementAction = {}));
        /**
         * Bit masks for telemetry flags
         */
        var AppManagementMenuFlags;
        (function (AppManagementMenuFlags) {
            AppManagementMenuFlags[AppManagementMenuFlags["ConfirmationDialogCancel"] = 256] = "ConfirmationDialogCancel";
            AppManagementMenuFlags[AppManagementMenuFlags["IsAnonymous"] = 1024] = "IsAnonymous";
        })(AppManagementMenuFlags || (AppManagementMenuFlags = {}));
        /**
         * Menu arrow key directions
         */
        var MenuDirection;
        (function (MenuDirection) {
            MenuDirection[MenuDirection["Up"] = 0] = "Up";
            MenuDirection[MenuDirection["Down"] = 1] = "Down";
            MenuDirection[MenuDirection["Left"] = 2] = "Left";
            MenuDirection[MenuDirection["Right"] = 3] = "Right";
        })(MenuDirection || (MenuDirection = {}));
        /**
         * Class for showing confirmation message in the insertion dialog
         */
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
            /**
             * Blocks default key handler while confirmation dialog is opened
             */
            ModalDialog.prototype.handleKeyDown = function (ev) {
                if (!this.isDialogVisible()) {
                    return false;
                }
                // handle all keys by default
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
            /**
             * Blocks default key handler while confirmation dialog is opened
             */
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
            /**
             * Hides the confirmation dialog
             */
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
                    // All tab index were removed previously to prevent tabbing while dialog is open
                    // Need to reset everything to correct attributes before the dialog was opened
                    if (previousTabValue !== null) {
                        element.setAttribute("tabindex", previousTabValue);
                    }
                    else {
                        element.removeAttribute("tabIndex");
                    }
                    if (previousDisabledValue !== null) {
                        element.disabled = (previousDisabledValue.toLowerCase() == "true");
                        // Find the first not disabled element to set new focus.
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
            /**
             * Shows the confirmation dialog, with message and buttons specified
             */
            ModalDialog.prototype.showDialog = function (message, buttonsCreationInfo) {
                if (!this.isDialogVisible()) {
                    // Prevent tabbing away from dialog while it is open
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
                        // It is best to set disabled on input that cannot be clicked or tabbed to for assistive technology
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
            /**
             * Repositions modal dialog in the center of the page
             */
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
            /**
             * Gets all the elements which can take keyboard focus on the page
             */
            ModalDialog.prototype.getTabbableElements = function () {
                return document.querySelectorAll("input,a,button,[tabindex]");
            };
            /**
             * Returns if the modal dialog is currently being displayed
             */
            ModalDialog.prototype.isDialogVisible = function () {
                return this.dialogDiv.style.display != "none" && this.dialogDiv.offsetWidth > 0;
            };
            /**
             * Tab key is pressed within the modal dialog
             */
            ModalDialog.prototype.onTab = function (ev) {
                var eventTarget = ev.srcElement ? ev.srcElement : ev.target;
                var buttonIndexAttribute = parseInt(eventTarget.getAttribute("data-buttonIndex"));
                var currentIndex = 0;
                if (isFinite(buttonIndexAttribute)) {
                    currentIndex = buttonIndexAttribute;
                }
                // WAC needs special logic to prevent keyboard focus to go out of insertion dialog
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
            /*
             * enter key is pressed within the modal dialog
             */
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
        /**
         * Class which handles the popup options menu
         */
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
                /** The modal dialog shown for confirming Agave removal. */
                this.removalConfirmationDialog = null;
                this.enterKeyTarget = null;
                this.dialogId = "appManagementMenuDialog";
                this.menuDiv = document.createElement("ul");
                this.menuDiv.setAttribute("role", "menu");
                this.menuDiv.setAttribute("tabindex", "-1");
                this.menuDiv.setAttribute("id", "OptionsMenu");
                this.removalConfirmationDialog = removalConfirmationDialog;
                // Do not show default context menu, it is triggered by right click
                this.menuDiv.oncontextmenu = function () {
                    return false;
                };
                // Note: the position within the DOM is important for WAC tab order.  The menu must be drawn before the cancel button.
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
            /**
             * Factory method for app options button
             */
            MenuHandler.prototype.createAppOptions = function (result) {
                return new AppOptions(result, this);
            };
            /**
             * Certain keys now have different behavior while options menu is being displayed, and should not use default behavior in the wef gallery
             */
            MenuHandler.prototype.handleKeyDown = function (ev) {
                // Only handle keys when the menu is visible
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
                    case 37: // Key.left:
                    case 39:
                        handled = true;
                        break;
                    default:
                        handled = false;
                        break;
                }
                return handled;
            };
            /**
             * Certain keys now have different behavior while options menu is being displayed, and should not use default behavior in the wef gallery
             */
            MenuHandler.prototype.handleKeyUp = function (ev) {
                var handled = false;
                // Only handle keys when the menu is visible
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
            /**
             * Hides the menu
             */
            MenuHandler.prototype.hideMenu = function (logData) {
                if (this.isMenuVisible()) {
                    this.menuDiv.style.display = "none";
                    // Log everytime menu is opened unless told otherwise
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
            /**
             * Returns if the menu is currently being displayed
             */
            MenuHandler.prototype.isMenuVisible = function () {
                return this.menuDiv.style.display != "none" && this.menuDiv.offsetWidth > 0;
            };
            /**
             * Displays the menu for a given app
             */
            MenuHandler.prototype.popupMenuForApp = function (result, optionsButton, appIndex, tnDiv, img) {
                var _this = this;
                this.currentResult = result;
                // Provide handlers for all menu buttons
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
                        // When user clicks remove app
                        _this.removeAppHandler(result, appIndex, tnDiv, img);
                    }, function () {
                        // When user clicks to cancel remove
                        _this.logData(result, AppManagementAction.Remove | AppManagementMenuFlags.ConfirmationDialogCancel, 0);
                    });
                });
                // Draw menu after all event bubbling has occured
                setTimeout(function () {
                    // The app opening the menu should always be selected, disable the force selection after all event bubbling occurs
                    WEF.IMPage.selectGalleryItems(appIndex, true /*forceSelected*/);
                    _this.positionMenu(optionsButton);
                    _this.clearMenuSelection();
                    _this.menuDiv.focus();
                }, 0);
            };
            /*
             * enter key is pressed within the menu dialog
             */
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
            /**
             * Computes the correct position to display the menu
             */
            MenuHandler.prototype.positionMenu = function (optionsButton) {
                // Draw the menu
                // Determine whether the menu needs to open up/down left/right
                this.menuDiv.style.display = "block";
                var insertDialogHeight = WEF.WefGalleryHelper.getDocumentHeight();
                var insertDialogWidth = WEF.WefGalleryHelper.getDocumentWidth();
                var menuRect = this.menuDiv.getBoundingClientRect();
                var optionButtonRect = optionsButton.getBoundingClientRect();
                var menuHeight = this.menuDiv.offsetHeight;
                var menuWidth = this.menuDiv.offsetWidth;
                var offsetTop = optionsButton.offsetHeight;
                // Calculate z-index to draw menu
                var calculatedZIndex = 1;
                var parentZIndex = parseInt(this.menuDiv.parentElement.style.zIndex);
                if (isFinite(parentZIndex)) {
                    calculatedZIndex = parentZIndex + 1;
                }
                this.menuDiv.style.zIndex = calculatedZIndex.toString();
                // Check if drawing the menu with the top starting at the options button will fit in insertion dialog
                if (optionButtonRect.top + menuHeight <= insertDialogHeight) {
                    // Draw the menu with the top starting at the options button
                    this.menuDiv.style.top = (optionButtonRect.top) + "px";
                }
                else {
                    // Draw the menu with the bottom starting at the options button
                    this.menuDiv.style.top = (optionButtonRect.top + offsetTop - menuHeight) + "px";
                }
                // Menu should open right by default unless it will be clipped.  This is reversed in rtl case.
                if (WEF.WefGalleryHelper.getHTMLDir() == "ltr") {
                    // Check if the menu will be clipped if drawn to the right
                    if (optionButtonRect.left + menuWidth <= insertDialogWidth) {
                        // Draw the menu with the left starting at the options button
                        this.menuDiv.style.left = (optionButtonRect.left) + "px";
                    }
                    else {
                        // Draw the menu with the right starting at the options button
                        this.menuDiv.style.left = (optionButtonRect.right - menuWidth) + "px";
                    }
                }
                else {
                    // Reverse due to RTL
                    // Check if the menu will be clipped if drawn to the left
                    if (optionButtonRect.left - menuWidth > 0) {
                        // Draw the menu with the right starting at the options button
                        this.menuDiv.style.left = (optionButtonRect.right - menuWidth) + "px";
                    }
                    else {
                        // Draw the menu with the left starting at the options button
                        this.menuDiv.style.left = (optionButtonRect.left) + "px";
                    }
                }
            };
            /**
             * Gets function to be called after user has confirmed they want to remove an app.
             */
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
            /**
             * Shows confirmation dialog to display before removing an app.
             */
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
            /**
             * Select menu item at a given index
             */
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
            /**
             * Clears selected menu item
             */
            MenuHandler.prototype.clearMenuSelection = function () {
                if (this.currentMenuItemIndex >= 0) {
                    this.menuItems[this.currentMenuItemIndex].setSelected(false);
                    this.currentMenuItemIndex = -1;
                }
            };
            /**
             * Selects next available menu item (skips disabled elements)
             */
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
            /**
             * Logs app management menu action
             */
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
        /**
         * Class which stores/sets state on an app management menu item
         */
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
            /**
             * Sets on click event for menu item
             */
            OptionsMenuItem.prototype.setOnClick = function (onClickHandler) {
                var _this = this;
                this.element.onclick = function () {
                    // Do not allow clicking menu item while disabled
                    if (_this.disabled) {
                        return;
                    }
                    onClickHandler();
                };
            };
            /**
             * Sets relevant properties to a menu item being selected
             */
            OptionsMenuItem.prototype.setSelected = function (selected) {
                this.element.setAttribute("aria-selected", selected.toString());
                if (selected) {
                    this.element.focus();
                }
            };
            /**
             * Sets relevant menu item disabled properties
             */
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
        /**
         * Class for operations with the app options button
         */
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
            /**
             * Factory method to create the options button
             */
            AppOptions.prototype.createOptionsButton = function (appIndex, tnDiv, img) {
                var _this = this;
                var optionsButton = null;
                // Only create the options menu if OMEX, and the host application can show the menu (not n-1 case or unsupported host)
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
            /**
             * Show the options button on the app
             */
            AppOptions.prototype.showOptionsButton = function () {
                if (this.optionsButton) {
                    this.optionsButton.style.display = "block";
                }
            };
            /**
             * Hide the options button on the app
             */
            AppOptions.prototype.hideOptionsButton = function () {
                if (this.optionsButton) {
                    this.optionsButton.style.display = "none";
                }
            };
            /**
             * Signals the menu handler to popup the menu for this app
             */
            AppOptions.prototype.popupMenu = function () {
                if (this.optionsButton) {
                    this.menuHandler.popupMenuForApp(this.result, this.optionsButton, this.appIndex, this.tnDiv, this.img);
                }
            };
            /**
             * Set appIndex
             */
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
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
WEF.AGAVE_DEFAULT_ICON = "AgaveDefaultIcon.png";
var WEF;
(function (WEF) {
    /**
     * Client facade supporting the communication between insertion dialog and the native client (rich client and iOS)
     * Abstracct class. Don't instantiate it.
     */
    var ClientFacade_Native = (function () {
        function ClientFacade_Native() {
            var _this = this;
            this.onShowEntitlementsComplete = null;
            this.onRemoveAgaveCallback = null;
            /** Env setting information in object */
            this.envSetting = {};
            /**
             * Abstract method that is going to be implemented in each platform client version of .ts file.
             */
            this.onGetEntitlementsInternal = function (results, hres) {
            };
            this.onGetEntitlements = function (results, hres) {
                if (_this.storeId != WEF.IMPage.currentStoreId) {
                    return;
                }
                // remove spin wheel and cleanup UI after GetEntitleMents call comes back
                WEF.IMPage.cleanUpGallery();
                // GetEntitlement error will override provider status.
                WEF.IMPage.uiState.ErrorBeforeReady = "";
                WEF.IMPage.providers[_this.storeId][1] = 0;
                WEF.IMPage.providers[_this.storeId][2] = 0;
                if (WEF.WefGalleryHelper.handleErrorCode(hres, _this.storeId, null /*storeType*/, null /*providerStatus*/)) {
                    // Return since there is an error.
                    return;
                }
                var etsArray = results.toArray ? results.toArray() : results;
                var entitlements = new Array();
                var existingId = {}; // Helper id set to prevent duplicate
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
                    // Prevent duplicate
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
                // Entitlement count is not 0, then show footer and Manage My App/Refresh buttons
                if (WEF.IMPage.footer.style.visibility === 'hidden') {
                    WEF.IMPage.showFooter();
                    WEF.IMPage.showHideRightMenuButtons(true /* showManageApp */, true /* showRefresh */);
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
        /**
         * Abstract method that is going to be implemented in each platform client version of .ts file.
         */
        ClientFacade_Native.prototype.launchAppManagePage = function () {
        };
        /**
         * Callback for get provider operation. Prepare the provider data and try to show them in gallery.
         */
        ClientFacade_Native.prototype.onGetProviders = function (results, hres) {
            var refreshRequired = WEF.WefGalleryHelper.retrieveRefreshRequired();
            var providers = results.toArray();
            // If there is no providers then show error and return.
            if (!providers || hres < 0 || providers.length === 0) {
                WEF.IMPage.cleanUpGallery();
                WEF.IMPage.showError(Strings.wefgallery.L_NoProviderError);
                return;
            }
            // Sorted based on storetype
            providers.sort(function (a, b) { return (a.toArray()[1] - b.toArray()[1]); });
            // If there is no providers then show error and return.
            if (!WEF.IMPage.initializeGalleryUI(providers, false /*let initTab to decide the display tab*/)) {
                WEF.IMPage.cleanUpGallery();
                WEF.IMPage.showError(Strings.wefgallery.L_NoProviderError);
                return;
            }
            WEF.IMPage.showContent(refreshRequired);
        };
        ClientFacade_Native.prototype.onGetProvidersShowContent = function (results, hres) {
            // Check the result of get provider info
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
            // init page internal page structure
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
    /**
     * WefGallery page shared by rich clients and iOS
     * Abstracct class. Don't instantiate it.
     */
    var WefGalleryPage_Native = (function (_super) {
        __extends(WefGalleryPage_Native, _super);
        function WefGalleryPage_Native(clientFacade) {
            var _this = this;
            _super.call(this, clientFacade);
            this.clientFacade = null;
            /**
             * Override with client code to enable Insert Item action.
             */
            this.insertItem = function (item) {
                if (_this.allowInsertion()) {
                    _this.clientFacade.insertAgave(item, _this.currentStoreType);
                }
            };
            /**
             * Override with shared client code to enable showEntitlements action.
             */
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
                // Let specific rich client provider custom hide behavior
                _this.hideButtons();
                // If the provider has error and the flag refresh is not set to true, then just show error and return.
                if (WEF.WefGalleryHelper.handleErrorCode(_this.getCurrentProviderHResult(), _this.currentStoreId, _this.currentStoreType, _this.getCurrentProviderStatus())) {
                    if (!refresh) {
                        return;
                    }
                }
                _this.gallery.style.overflowY = "auto"; // Make the gallery vertically scrollable.
                // Add Spinwheel
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
            /**
             * Listener to receive featured page's posting message
             */
            this.postMessageListener = function (e) {
                if (e.data == "REFRESH_REQUIRED") {
                    WEF.WefGalleryHelper.saveRefreshRequired(true);
                }
            };
            this.clientFacade = clientFacade;
        }
        /**
         * Determine whether insertion operation is allowed in current condition
         */
        WefGalleryPage_Native.prototype.allowInsertion = function () {
            return this.currentStoreType != WEF.StoreTypeEnum.ThisDocument;
        };
        /**
         * Override with shared client code to Retrieve StoreID.
         * TODO: Verify whether this function is useful. If not, remove it.
         */
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
        /**
         * Launch the app management page.
         */
        WefGalleryPage_Native.prototype.launchAppManagePage = function () {
            this.clientFacade.launchAppManagePage();
        };
        /**
         * Remove an Agave.
         */
        WefGalleryPage_Native.prototype.removeAgave = function (result, callback) {
            this.clientFacade.removeAgave(result, this.currentStoreType, callback);
        };
        /**
         * Define the client specific behavior to hide some UI elements.
         */
        WefGalleryPage_Native.prototype.hideButtons = function () {
            // Nothing happen
        };
        /**
         * Implement the client specific behavior to show the WefGallery.
         */
        WefGalleryPage_Native.prototype.showItInternal = function () {
            this.wefGalleryAppOnLoad();
            WEF.WefGalleryHelper.addEventListener(window, "message", this.postMessageListener); // IE, register listener to receive Featured Page's posting message    
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
/*************************************************************
  Copyright (c) 2015 Microsoft Corporation
**************************************************************/
var WEF;
(function (WEF) {
    /**
     * Facade class that handles the the inter-boundary communication with rich client.
     */
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
        /**
         * Override with Rich client code to run the first step.
         */
        ClientFacade_RichClient.prototype.runShowIt = function () {
            if (typeof this.external.GetProviders !== 'undefined') {
                this.external.GetProviders(this.onGetProviders);
            }
        };
        /**
         * Override with Rich client code to Get Entitlements.
         * @param refresh Whether force to refresh
         */
        ClientFacade_RichClient.prototype.getEntitlements = function (storeId, refresh, onGetEntitlements) {
            if (typeof this.external.GetEntitlements !== 'undefined') {
                this.external.GetEntitlements(storeId, refresh /*force to refresh*/, onGetEntitlements);
            }
        };
        /**
         * Override with Rich client code to enable Insert Item action.
         */
        ClientFacade_RichClient.prototype.insertAgave = function (item, currentStoreType) {
            if (typeof this.external.Insert !== 'undefined') {
                this.external.Insert(item.result.id, item.result.targetType, item.result.appVersion, currentStoreType, item.result.storeId, currentStoreType == WEF.StoreTypeEnum.MarketPlace ? item.result.id : item.result.assetId, currentStoreType == WEF.StoreTypeEnum.MarketPlace ? item.result.id : item.result.assetStoreId, item.result.width, item.result.height);
            }
        };
        /**
         * Override with Rich client code to get Initial Tab.
         */
        ClientFacade_RichClient.prototype.getInitTab = function () {
            return this.external.GetInitTab();
        };
        /**
         * Override with Rich client code to get Landing Page Url.
         */
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
        /**
         * Override with Rich client code to launch App Management Page.
         */
        ClientFacade_RichClient.prototype.launchAppManagePage = function () {
            if (typeof this.external.LaunchUrl !== 'undefined' && typeof WEF.IMPage !== 'undefined') {
                WEF.WefGalleryHelper.saveRefreshRequired(true);
                this.external.LaunchUrl(WEF.IMPage.getAppManagePageUrl(), true);
            }
        };
        /**
         * Override with Rich client code to remove the dimension information element.
         */
        ClientFacade_RichClient.prototype.removeResultsDimensionInfo = function (results) {
            // Do nothing needed on Rich Client
        };
        /**
         * Override with Rich client code to remove an Agave.
         */
        ClientFacade_RichClient.prototype.removeAgave = function (result, currentStoreType, callback) {
            if (typeof this.external.RemoveAgave != 'undefined') {
                this.setOnRemoveAgaveCallback(callback);
                this.external.RemoveAgave(result.id, result.targetType, result.appVersion, currentStoreType, result.storeId, result.assetId, result.assetStoreId, this.onRemoveAgave);
            }
        };
        /**
         * Launch the browser window to open a URL
         */
        ClientFacade_RichClient.prototype.launchUrl = function (url, closeWindow) {
            if (typeof this.external.LaunchUrl != 'undefined') {
                this.external.LaunchUrl(url, closeWindow);
            }
        };
        /**
         * Preload the manifest to accelerate the app loading speed
         */
        ClientFacade_RichClient.prototype.preloadManifest = function (item, currentStoreType) {
            if (typeof this.external.PreloadManifest != 'undefined' && item) {
                if (currentStoreType == WEF.StoreTypeEnum.MarketPlace) {
                    this.external.PreloadManifest(item.result.id, item.result.appVersion, currentStoreType, item.result.storeId, item.result.assetId, item.result.assetId); // Use as assetStoreId
                }
                else {
                    this.external.PreloadManifest(item.result.id, item.result.appVersion, currentStoreType, item.result.storeId, item.result.assetId, item.result.assetStoreId);
                }
            }
        };
        /**
         * Log usage of app management menu
         */
        ClientFacade_RichClient.prototype.logAppManagementAction = function (assetId, operationInfo, hresult) {
            if (typeof this.external.LogAppManagementAction != 'undefined') {
                this.external.LogAppManagementAction(assetId, operationInfo, hresult);
            }
        };
        /**
         * Parse the given url query string to a key-value dictionary
         * @param queryStr the query string to be parse
         */
        ClientFacade_RichClient.parseUrlQueryString = function (queryStr) {
            if (queryStr.length <= 1) {
                return {};
            }
            queryStr = queryStr.substring(1); // Remove the leading "?"
            var searchRegex = new RegExp("([^&=]+)=?([^&]*)", "g");
            var queryStrObj = {}, match;
            while (match = searchRegex.exec(queryStr)) {
                // Need to decode the input, which is encoded by rich client
                queryStrObj[decodeURIComponent(match[1])] = decodeURIComponent(match[2]);
            }
            return queryStrObj;
        };
        return ClientFacade_RichClient;
    })(WEF.ClientFacade_Native);
    WEF.ClientFacade_RichClient = ClientFacade_RichClient;
    /**
     * The WefGallery page for rich clients.
     */
    var WefGalleryPage_RichClient = (function (_super) {
        __extends(WefGalleryPage_RichClient, _super);
        function WefGalleryPage_RichClient() {
            var _this = this;
            _super.apply(this, arguments);
            /**
             * Override method with Rich client code to call sign In
             */
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
        /**
         * Override with Rich client code to enable page load action.
         */
        WefGalleryPage_RichClient.prototype.onPageLoad = function () {
            if (typeof this.clientFacade.external.OnPageLoad != 'undefined') {
                this.clientFacade.external.OnPageLoad();
            }
        };
        /**
         * Override with Rich client code to enable select action.
         * Hook up with the preloading feature in rich client.
         */
        WefGalleryPage_RichClient.prototype.onItemSelect = function (item) {
            this.clientFacade.preloadManifest(item, this.currentStoreType);
        };
        /**
         * Override with Rich client code to enable cancel action.
         */
        WefGalleryPage_RichClient.prototype.cancelDialog = function () {
            if (typeof this.clientFacade.external.CancelDialog != 'undefined') {
                this.clientFacade.external.CancelDialog();
            }
        };
        /**
         * Override with Rich client code to enable more info action.
         */
        WefGalleryPage_RichClient.prototype.onMoreInfo = function () {
            var url = this.results[this.currentIndex].appEndNodeUrl;
            this.clientFacade.launchUrl(url, false);
        };
        /**
         * Override with Rich client code to enable Trust All action.
         */
        WefGalleryPage_RichClient.prototype.trustAllAgaves = function () {
            if (typeof this.clientFacade.external.TrustAllInDocOmexApps != 'undefined') {
                this.clientFacade.external.TrustAllInDocOmexApps();
            }
        };
        /**
         * Performs n-1 check to see if all the dlls are the correct version to properly remove an app.
         */
        WefGalleryPage_RichClient.prototype.canShowAppManagementMenu = function () {
            return true;
        };
        /**
         * Launches the url in a new IE window
         */
        WefGalleryPage_RichClient.prototype.invokeWindowOpen = function (pageUrl) {
            this.clientFacade.launchUrl(pageUrl, false);
        };
        /**
         * Override method to hide buttons.
         */
        WefGalleryPage_RichClient.prototype.hideButtons = function () {
            // Nothing happen
        };
        /**
         * Check to see if Omex provider is ready with right status, show entitlement content if it is
         */
        WefGalleryPage_RichClient.prototype.checkOmexProviderInfoShowContent = function () {
            try {
                this.clientFacade.external.GetProviders(this.clientFacade.onGetProvidersShowContent);
            }
            catch (ex) {
                this.showError(Strings.wefgallery.L_GetEntitilementsGeneralError);
            }
        };
        /**
         * Handle enter key events
         */
        WefGalleryPage_RichClient.prototype.executeButtonCommand = function (element, event) {
            if (event === void 0) { event = null; }
            _super.prototype.executeButtonCommand.call(this, element, event);
            // Handle the key event only available in rich client
            if (element.getAttribute("id") == "BtnTrustAll") {
                this.trustAllAgaves();
            }
        };
        /**
         * Setting up page level and hooking up events
         */
        WefGalleryPage_RichClient.prototype.wefGalleryAppOnLoad = function () {
            var _this = this;
            _super.prototype.wefGalleryAppOnLoad.call(this);
            // Special setup for rich client page.
            this.btnTrustAll.onclick = function () { _this.trustAllAgaves(); };
            this.readMoreATag.onclick = function () { _this.onMoreInfo(); };
            this.permissionATag.onclick = function () { _this.onMoreInfo(); };
        };
        return WefGalleryPage_RichClient;
    })(WEF.WefGalleryPage_Native);
    WEF.WefGalleryPage_RichClient = WefGalleryPage_RichClient;
    /**
     * Setup the client specific classes for rich client environment.
     */
    WEF.setupClientSpecificWefGalleryPage = function () {
        // Create Page object.
        var clientFacade = new ClientFacade_RichClient(window.external);
        WEF.IMPage = new WefGalleryPage_RichClient(clientFacade);
    };
})(WEF || (WEF = {}));
var WEF;
(function (WEF) {
    /**
     * Mock the safe array used by rich client
     */
    var MockSafeArray = (function () {
        function MockSafeArray(data) {
            this.data = data;
        }
        MockSafeArray.prototype.toArray = function () {
            return this.data;
        };
        /**
         * Construct a fake saft array from a 2-d array.
         */
        MockSafeArray.constructFrom2dArray = function (input) {
            var output = [];
            for (var i = 0; i < input.length; i++) {
                var item = [];
                for (var j = 0; j < input[i].length; j++) {
                    item.push(input[i][j]);
                }
                output.push(new MockSafeArray(item));
            }
            return new MockSafeArray(output);
        };
        return MockSafeArray;
    })();
    WEF.MockSafeArray = MockSafeArray;
    /**
     * Mock the window.external project for rich client.
     */
    var MockExternal = (function () {
        function MockExternal() {
            var _this = this;
            this.conversationId = 0;
            this.callbacks = {};
            this.receiveMessage = function (e) {
                var data = e.data;
                data = JSON.parse(data);
                if (data.action == "GetProviders") {
                    var callback = _this.callbacks[data.conversationId];
                    _this.callbacks[data.conversationId] = null;
                    var providersSafeArray = MockSafeArray.constructFrom2dArray(data.providers);
                    callback(providersSafeArray, data.status);
                }
                else if (data.action == "GetEntitlements") {
                    var callback = _this.callbacks[data.conversationId];
                    _this.callbacks[data.conversationId] = null;
                    var entitlementsSafeArray = MockSafeArray.constructFrom2dArray(data.entitlements);
                    callback(entitlementsSafeArray, data.status);
                }
                else if (data.action == "RemoveAgave") {
                    var callback = _this.callbacks[data.conversationId];
                    _this.callbacks[data.conversationId] = null;
                    callback(new MockSafeArray(data.result), 0);
                }
            };
        }
        // Mock up methods for window.external
        MockExternal.prototype.CancelDialog = function () {
            var msg = { action: "CancelDialog" };
            this.invokeMethod(msg, null);
        };
        MockExternal.prototype.GetEntitlements = function (storeId, refresh, onGetEntitlements) {
            var msg = { action: "GetEntitlements", storeId: storeId, refresh: refresh };
            this.invokeMethod(msg, onGetEntitlements);
        };
        MockExternal.prototype.GetProviders = function (callback) {
            var msg = { action: "GetProviders" };
            this.invokeMethod(msg, callback);
        };
        MockExternal.prototype.Insert = function (id, targetType, appVersion, storeType, storeId, assetId, assetStoreId, width, height) {
            var msg = {
                action: "Insert",
                id: id,
                targeType: targetType,
                appVersion: appVersion,
                storeType: storeType,
                storeId: storeId,
                assetId: assetId,
                assetStoreId: assetStoreId,
                width: width,
                height: height
            };
            this.invokeMethod(msg, null);
        };
        MockExternal.prototype.GetLandingPageUrl = function () {
            var msg = {
                action: "GetLandingPageUrl"
            };
            return "http://o15.officeredir.microsoft.com/r/rlidMktplcOSFRecs?ver=16&app=winwordd.exe&clid=1033&p1=16.0.6025.1000&p2=6&p3=en-US%2Fwa104099688&p4=0&p5=0&lidhelp=0409&liduser=0409&lidui=0409&client=Win32_Word&cv=16.0.0.0&authtype=0&pm=0&lcid=1033&syslcid=1033&uilcid=1033";
        };
        MockExternal.prototype.LaunchUrl = function (url, closeWindow) {
            var msg = {
                action: "LaunchUrl",
                url: url,
                closeWindow: closeWindow
            };
            this.invokeMethod(msg, null);
        };
        MockExternal.prototype.MountOrSignInLiveId = function () {
            var msg = {
                action: "MountOrSignInLiveId"
            };
            this.invokeMethod(msg, null);
        };
        MockExternal.prototype.OnPageLoad = function () {
            var msg = {
                action: "OnPageLoad"
            };
            this.invokeMethod(msg, null);
        };
        MockExternal.prototype.PreloadManifest = function (id, appVersion, storeType, storeId, assetId, assetStoreId) {
            var msg = {
                action: "PreloadManifest",
                id: id,
                appVersion: appVersion,
                storeType: storeType,
                storeId: storeId,
                assetId: assetId,
                assetStoreId: assetStoreId
            };
            this.invokeMethod(msg, null);
        };
        MockExternal.prototype.RemoveAgave = function (id, targetType, appVersion, storeType, storeId, assetId, assetStoreId, onRemoveAgave) {
            var msg = {
                action: "RemoveAgave",
                id: id,
                targeType: targetType,
                appVersion: appVersion,
                storeType: storeType,
                storeId: storeId,
                assetId: assetId,
                assetStoreId: assetStoreId
            };
            this.invokeMethod(msg, onRemoveAgave);
        };
        MockExternal.prototype.init = function () {
            // This mock-up class works with its parent frame. It sends the request to parent via postMessage,
            // and picks up the response via the event listener.
            window.addEventListener("message", this.receiveMessage, false);
        };
        MockExternal.prototype.invokeMethod = function (msg, callback) {
            msg.conversationId = this.conversationId;
            if (callback) {
                this.callbacks[this.conversationId.toString()] = callback;
            }
            this.conversationId++;
            window.parent.postMessage(JSON.stringify(msg), "*");
        };
        return MockExternal;
    })();
    WEF.MockExternal = MockExternal;
})(WEF || (WEF = {}));
var WEF;
(function (WEF) {
    var WefGalleryPage_RichClient_Outlook = (function (_super) {
        __extends(WefGalleryPage_RichClient_Outlook, _super);
        function WefGalleryPage_RichClient_Outlook() {
            _super.apply(this, arguments);
        }
        WefGalleryPage_RichClient_Outlook.prototype.hideButtons = function () {
            this.showHideRightMenuButtons(true, false);
            if (this.getLandingPageUrl() == undefined) {
                this.footerLink.style.visibility = 'hidden';
            }
            var permissionTextAndLink = document.getElementById('PermissionTextAndLink');
            if (permissionTextAndLink != undefined) {
                permissionTextAndLink.style.visibility = 'hidden';
            }
        };
        WefGalleryPage_RichClient_Outlook.prototype.retrieveStoreId = function () {
            // We want legacy behavior. Returning null ensures that we get legacy behavior.
            return null;
        };
        WefGalleryPage_RichClient_Outlook.prototype.gotoStore = function () {
            // Outlook has only one tab. So, we need to fallback to launching browser for Outlook.
            this.launchLandingPage();
        };
        WefGalleryPage_RichClient_Outlook.prototype.launchLandingPage = function () {
            // Currently launch Url is blocked by MSO/WebDialog infrastructure
            // We have to use its custom command.
            if (typeof window.external.LaunchUrl !== 'undefined' && typeof WEF.IMPage !== 'undefined') {
                WEF.WefGalleryHelper.saveRefreshRequired(true);
                window.external.LaunchUrl(WEF.IMPage.getLandingPageUrl(), true);
            }
        };
        return WefGalleryPage_RichClient_Outlook;
    })(WEF.WefGalleryPage_RichClient);
    WEF.WefGalleryPage_RichClient_Outlook = WefGalleryPage_RichClient_Outlook;
    /**
     * Setup the client specific classes for Outlook rich client environment.
     */
    WEF.setupClientSpecificWefGalleryPage = function () {
        var clientFacade = new WEF.ClientFacade_RichClient(window.external);
        WEF.IMPage = new WefGalleryPage_RichClient_Outlook(clientFacade);
    };
})(WEF || (WEF = {}));
var WEF;
(function (WEF) {
    /**
     * Setup the client specific classes for the mock-up rich client environment.
     */
    WEF.setupClientSpecificWefGalleryPage = function () {
        var mockExternal = new WEF.MockExternal();
        mockExternal.init();
        var clientFacade = new WEF.ClientFacade_RichClient(mockExternal);
        WEF.IMPage = new WEF.WefGalleryPage_RichClient_Outlook(clientFacade);
    };
})(WEF || (WEF = {}));
