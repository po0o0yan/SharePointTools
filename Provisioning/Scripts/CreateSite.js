/// <reference path="jquery-1.7.1.js" />
/// <reference name="MicrosoftAjax.js" />
/// <reference path="~/_layouts/15/init.js" />
/// <reference path="~/_layouts/15/SP.Core.js" />
/// <reference path="~/_layouts/15/SP.Runtime.js" />
/// <reference path="~/_layouts/15/SP.UI.Dialog.js" />
/// <reference path="~/_layouts/15/SP.js" />

'use strict';

window.PM = window.PM || {};

window.PM.createSiteInHostWeb = function () {
    var createSite = function (title, description, url, template) {
        var webCreationInfo = new SP.WebCreationInformation();
        webCreationInfo.set_title(title);
        webCreationInfo.set_description(description);
        webCreationInfo.set_language(1033);
        webCreationInfo.set_url(url);
        webCreationInfo.set_useSamePermissionsAsParentSite(true);
        webCreationInfo.set_webTemplate(template);
        web.get_webs().add(webCreationInfo);
        web.update();
        hostWebContext.executeQueryAsync(
                                         Function.createDelegate(this, function () { writeSuccessMessage("Site has been created successfully"); }),
                                         Function.createDelegate(this, function (sender, args) { writeErrorMessage("Error Creating Site </br>" + args.get_message()); }));
    }
    return {
        execute: function (title, description, url, template) {
            createSite(title, description, url, template );
        }
    }
}();