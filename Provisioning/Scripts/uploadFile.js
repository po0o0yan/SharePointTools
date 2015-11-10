/// <reference path="jquery-1.7.1.js" />
/// <reference name="MicrosoftAjax.js" />
/// <reference path="~/_layouts/15/init.js" />
/// <reference path="~/_layouts/15/SP.Core.js" />
/// <reference path="~/_layouts/15/SP.Runtime.js" />
/// <reference path="~/_layouts/15/SP.UI.Dialog.js" />
/// <reference path="~/_layouts/15/SP.js" />

'use strict';

window.PM = window.PM || {};

window.PM.uploadFileToHostWeb = function () {
    var hostWebUrl,
        appWebUrl,
        hostWebContext,
        destinationServerRelativeUrl,
        destinationFileName,
        factory,
        appContextSite,
        web,

    // locate a file in the app web and retrieve the contents. If successful, provision to host web..
    readFromAppWebAndProvisionToHost = function (appFileUrl, hostWebServerRelativeUrl, hostWebFileName, fileType) {
        destinationServerRelativeUrl = hostWebServerRelativeUrl;
        destinationFileName = hostWebFileName;
        switch (fileType) {
            case 'text':
                {
                    var req = $.ajax({
                        url: appFileUrl,
                        type: "GET",
                        cache: false
                    }).done(function (fileContents) {
                        if (fileContents !== undefined && fileContents.length > 0) {
                            uploadTextFileToHostWeb(destinationServerRelativeUrl, destinationFileName, fileContents);
                        }
                        else {
                            writeErrorMessage('Failed to read file from app web, so not uploading to host web..');
                        }
                    }).fail(function (jqXHR, textStatus) {
                        writeErrorMessage("Request for page in app web failed: " + textStatus);
                    });
                    break;
                }
            case 'binary':
                {
                    debugger;
                    $.ajax({
                        url: appFileUrl,
                        type: "GET",
                        dataType: "binary",
                        processData: false,
                        responseType: 'arraybuffer',
                        cache: false
                    }).done(function (fileContents) {
                        uploadBinaryFileToHostWeb(destinationServerRelativeUrl, destinationFileName, fileContents);
                    }).fail(function (jqXHR, textStatus) {
                        writeErrorMessage(textStatus);
                    });
                    break;
                }
        }
    },
    uploadBinaryFileToHostWeb = function (serverRelativeUrl, filename, contents) {
        var createInfo = new SP.FileCreationInformation();
        createInfo.set_content(arrayBufferToBase64(contents));
        createInfo.set_overwrite(true);
        createInfo.set_url(filename);
        web = appContextSite.get_web();
        var files = web.getFolderByServerRelativeUrl(serverRelativeUrl).get_files();
        hostWebContext.load(files);
        files.add(createInfo);
        hostWebContext.executeQueryAsync(onProvisionFileSuccess, onProvisionFileFail);
    },
    uploadTextFileToHostWeb = function (serverRelativeUrl, filename, contents) {
        var createInfo = new SP.FileCreationInformation();
        createInfo.set_content(new SP.Base64EncodedByteArray());
        for (var i = 0; i < contents.length; i++) {
            createInfo.get_content().append(contents.charCodeAt(i));
        }
        createInfo.set_overwrite(true);
        createInfo.set_url(filename);
        web = appContextSite.get_web();
        var files = web.getFolderByServerRelativeUrl(serverRelativeUrl).get_files();
        hostWebContext.load(files);
        files.add(createInfo);
        hostWebContext.executeQueryAsync(onProvisionFileSuccess, onProvisionFileFail);
    },
    onProvisionFileSuccess = function () {
        $('#message').append('<br /><div>File provisioned in host web successfully: ' + hostWebUrl + '/' + destinationServerRelativeUrl + '/' + destinationFileName + '</div>');
        setMaster('/' + destinationServerRelativeUrl + '/' + destinationFileName);
    },
    onProvisionFileFail = function (sender, args) {
        writeErrorMessage('Failed to provision file into host web. Error:' + args.get_message() + '\n' + args.get_stackTrace());
    },

    // set file on host web..
    setMaster = function (masterUrl) {
        var hostWeb = hostWebContext.get_web();
        hostWeb.set_masterUrl(masterUrl);
        hostWeb.update();
        hostWebContext.load(hostWeb);
        hostWebContext.executeQueryAsync(onSetMasterSuccess, onSetMasterFail);
    },
    onSetMasterSuccess = function () {
        $('#message').append('<br /><div>File updated successfully..</div>');
    },
    onSetMasterFail = function (sender, args) {
        writeErrorMessage('Failed to update file on host web. Error:' + args.get_message());
    },

    init = function () {
        var hostWebUrlFromQS = $.getUrlVar("SPHostUrl");
        hostWebUrl = (hostWebUrlFromQS !== undefined) ? decodeURIComponent(hostWebUrlFromQS) : undefined;

        var appWebUrlFromQS = $.getUrlVar("SPAppWebUrl");
        appWebUrl = (appWebUrlFromQS !== undefined) ? decodeURIComponent(appWebUrlFromQS) : undefined;
        hostWebContext = new SP.ClientContext(appWebUrl);
        factory = new SP.ProxyWebRequestExecutorFactory(appWebUrl);
        hostWebContext.set_webRequestExecutorFactory(factory);
        appContextSite = new SP.AppContextSite(hostWebContext, hostWebUrl);
    }

    return {
        execute: function (sourceFileRelativePath, destinationListRelativePath, destinationFileName, fileType) {
            init();
            readFromAppWebAndProvisionToHost(appWebUrl + sourceFileRelativePath, destinationListRelativePath, destinationFileName, fileType);
        }
    }
}();

window.PM.AppHelper = {
    getRelativeUrlFromAbsolute: function (absoluteUrl) {
        absoluteUrl = absoluteUrl.replace('https://', '');
        var parts = absoluteUrl.split('/');
        var relativeUrl = '/';
        for (var i = 1; i < parts.length; i++) {
            relativeUrl += parts[i] + '/';
        }
        return relativeUrl;
    },
};

$.ajaxTransport("+binary", function (options, originalOptions, jqXHR) {
    // check for conditions and support for blob / arraybuffer response type
    if (window.FormData && ((options.dataType && (options.dataType == 'binary')) || (options.data && ((window.ArrayBuffer && options.data instanceof ArrayBuffer) || (window.Blob && options.data instanceof Blob))))) {
        return {
            // create new XMLHttpRequest
            send: function (headers, callback) {
                // setup all variables
                var xhr = new XMLHttpRequest(),
		url = options.url,
		type = options.type,
		async = options.async || true,
		// blob or arraybuffer. Default is blob
		dataType = options.responseType || "blob",
		data = options.data || null,
		username = options.username || null,
		password = options.password || null;

                xhr.addEventListener('load', function () {
                    var data = {};
                    data[options.dataType] = xhr.response;
                    // make callback and send data
                    callback(xhr.status, xhr.statusText, data, xhr.getAllResponseHeaders());
                });

                xhr.open(type, url, async, username, password);

                // setup custom headers
                for (var i in headers) {
                    xhr.setRequestHeader(i, headers[i]);
                }

                xhr.responseType = dataType;
                xhr.send(data);
            },
            abort: function () {
                jqXHR.abort();
            }
        };
    }
});
var arrayBufferToBase64 = function (buffer) {
    var binary = '';
    var bytes = new Uint8Array(buffer);
    var len = bytes.byteLength;
    for (var i = 0; i < len; i++) {
        binary += String.fromCharCode(bytes[i]);
    }
    return window.btoa(binary);
}