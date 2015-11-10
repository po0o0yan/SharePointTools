/// <reference path="jquery-1.7.1.js" />
/// <reference name="MicrosoftAjax.js" />
/// <reference path="~/_layouts/15/init.js" />
/// <reference path="~/_layouts/15/SP.Core.js" />
/// <reference path="~/_layouts/15/SP.Runtime.js" />
/// <reference path="~/_layouts/15/SP.UI.Dialog.js" />
/// <reference path="~/_layouts/15/SP.js" />

'use strict';

window.PM = window.PM || {};

window.PM.ProvisionContentTypes = function () {
    var hostWebUrl,
        appWebUrl,
        hostWebContext,
        hostWebContentTypes,
        createdField,
        contentTypeName = 'Announcements',
        contentTypeDescription = 'Content Type for Announcements',
        contentTypeGroupName = 'Dynacare',
        listName = 'Announcements',
        listDescription = 'List for Announcements',
        factory,
        web,
        fields = new Array(),
        createdContentType,

    readFromAppWebAndProvisionToHost = function (appPageUrl) {
        var req = $.ajax({
            url: appPageUrl,
            type: "GET",
            cache: false
        }).done(function (fileContents) {
            if (fileContents !== undefined && fileContents.length > 0) {
                provisionFields(fileContents);
            }
            else {
                writeSuccessMessage('Failed to read file from app web, so not uploading to host web.');
            }
        }).fail(function (jqXHR, textStatus) {
            writeErrorMessage('<br /> Request for page in app web failed: ' + textStatus);
        });
    },
    provisionFields = function (fileContents) {
        $.each(fileContents.split(','), function () { createField(this) })
    },
    createField = function (fieldXml) {
        var fields = web.get_fields();
        createdField = fields.addFieldAsXml(fieldXml, false, SP.AddFieldOptions.AddToNoContentType);
        var currentFieldName = fieldXml.split('=')[2].split('"')[1];

        hostWebContext.load(fields);
        hostWebContext.load(createdField);
        hostWebContext.executeQueryAsync(
                                        Function.createDelegate(this, function () { onProvisionFieldSuccess(currentFieldName) }),
                                        Function.createDelegate(this, function () { onProvisionFieldFail(this, currentFieldName) }));
    },
    onProvisionFieldSuccess = function (currentFieldName) {
        fields.push(currentFieldName);
        writeSuccessMessage('Field provisioned in host web successfully: ' + currentFieldName);
    },
    onProvisionFieldFail = function (currentFieldName) {
        writeErrorMessage('Failed to provision field into host web: ' + currentFieldName);
    },

    createContentType = function (ctypeName, ctypeDescription, ctypeGroup) {
       
        hostWebContentTypes = web.get_contentTypes();
        var cTypeInfo = new SP.ContentTypeCreationInformation();
        cTypeInfo.set_name(ctypeName);
        cTypeInfo.set_description(ctypeDescription);
        cTypeInfo.set_group(ctypeGroup);
        hostWebContentTypes.add(cTypeInfo);
        hostWebContext.load(hostWebContentTypes);
        hostWebContext.executeQueryAsync(onProvisionContentTypeSuccess, onProvisionContentTypeFail);
    },
    onProvisionContentTypeSuccess = function () {
        writeSuccessMessage('Content type provisioned in host web successfully.');
        addFieldToContentType(fields);
    },
    onProvisionContentTypeFail = function (sender, args) {
        writeErrorMessage('<strong>Failed to provision content type into host web. Error:' + sender.statusCode);
    },
    addFieldToContentType = function (fields) {
        var cTypeFound = false;

        var contentTypeEnumerator = hostWebContentTypes.getEnumerator();
        while (contentTypeEnumerator.moveNext()) {
            var contentType = contentTypeEnumerator.get_current();
            if (contentType.get_name() === contentTypeName) {
                cTypeFound = true;
                createdContentType = contentType;
                break;
            }
        }

        if (cTypeFound) {
            // - NOT the below line - SP.FieldCollection doesn't appear to have an add() method when fetched from content type..
            //contentType.get_fields.add(fieldInternalName)
            // - instead, this..

            $.each(fields, function () {
                createdField = web.get_fields().getByInternalNameOrTitle(this);
                hostWebContext.load(createdField);
                var fieldRef = new SP.FieldLinkCreationInformation();
                fieldRef.set_field(createdField);
                createdContentType.get_fieldLinks().add(fieldRef);

            });
            createdContentType.update(true);
            hostWebContext.load(createdContentType);
            hostWebContext.executeQueryAsync(onAddFieldToContentTypeSuccess, onAddFieldToContentTypeFail);
        }
        else {
            writeErrorMessage('Failed to add field to content type - check the content type exists!');
        }
    },
        onAddFieldToContentTypeSuccess = function (fieldInternalName) {
            writeSuccessMessage('Field added to content type in host web successfully');
            createList();
        },
        onAddFieldToContentTypeFail = function (sender, args) {
            writeErrorMessage('Failed to add field to content type: ' + args.$2J_2);
        },
        createList = function () {
            var listCreationInfo = new SP.ListCreationInformation();
            listCreationInfo.set_title(listName);
            listCreationInfo.set_description(listDescription);
            listCreationInfo.set_templateType(SP.ListTemplateType.webPageLibrary);
            web.get_lists().add(listCreationInfo);
            hostWebContext.executeQueryAsync(onListCreationSuccess, onListCreationFail);
        },
    onListCreationSuccess = function () {
        writeSuccessMessage("List Created");
        addContentTypeToList();
    },
    onListCreationFail = function (sender, args) {
        writeErrorMessage("List Failed");
    },

    addContentTypeToList = function () {
        var list = web.get_lists().getByTitle(listName);
        var cTypes = list.get_contentTypes();
        cTypes.addExistingContentType(createdContentType);
        hostWebContext.load(cTypes);
        hostWebContext.executeQueryAsync(onAddContentTypeToListSuccess, onAddContentTypeToListFail);
    },
    onAddContentTypeToListSuccess = function () {
        writeSuccessMessage("Content Type association successful")
    },
    onAddContentTypeToListFail = function (sender, args) {
        writeErrorMessage("Content Type association failed");
    },

    init = function () {
        var hostWebUrlFromQS = $.getUrlVar("SPHostUrl");
        hostWebUrl = (hostWebUrlFromQS !== undefined) ? decodeURIComponent(hostWebUrlFromQS) : undefined;
        var appWebUrlFromQS = $.getUrlVar("SPAppWebUrl");
        appWebUrl = (appWebUrlFromQS !== undefined) ? decodeURIComponent(appWebUrlFromQS) : undefined;
        hostWebContext = new SP.ClientContext(appWebUrl);
        factory = new SP.ProxyWebRequestExecutorFactory(appWebUrl);
        hostWebContext.set_webRequestExecutorFactory(factory);
        var appContextSite = new SP.AppContextSite(hostWebContext, hostWebUrl);
        web = appContextSite.get_web();
    }

    return {
        execute: function () {
            init();
            readFromAppWebAndProvisionToHost(appWebUrl + '/Announcements/Fields.txt');
            createContentType(contentTypeName, contentTypeDescription, contentTypeGroupName);
        }
    }
}();