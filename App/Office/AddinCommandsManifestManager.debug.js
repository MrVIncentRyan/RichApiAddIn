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
