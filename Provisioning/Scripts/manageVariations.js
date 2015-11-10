/// <reference path="jquery-1.7.1.js" />
/// <reference name="MicrosoftAjax.js" />
/// <reference path="~/_layouts/15/init.js" />
/// <reference path="~/_layouts/15/SP.Core.js" />
/// <reference path="~/_layouts/15/SP.Runtime.js" />
/// <reference path="~/_layouts/15/SP.UI.Dialog.js" />
/// <reference path="~/_layouts/15/SP.js" />

'use strict'

window.PM = window.PM || {};

window.PM.createVariationLabel = function () {
    var
        labelItems,
        labelsList,
        variationLabelFieldNames = {
            isSource: "Is_x0020_Source",
            language: "Language",
            locale: "Locale",
            title: "Title",
            topWebUrl: "Top_x0020_Web_x0020_URL",
            displayName: "Flag_x0020_Control_x0020_Display",
            hierarchyIsCreated: "Hierarchy_x0020_Is_x0020_Created",
            notificationMode: "NotificationMode",
            isMachineTranslationEnabled: "IsMachineTranslationEnabledField",
            isHumanTranslationEnabled: "IsHumanTranslationEnabledFieldNa",
            machineTranslationLanguage: "MachineTranslationLanguageFieldN",
            humanTranslationLanguage: "HumanTranslationLanguageFieldNam",
            hierarchyCreationMode: "Hierarchy_x0020_Creation_x0020_M",
        },
        createLabel = function (locale, title, topWebUrl, language, displayName) {
            var varLabelsListId = webProperties.get_item('_VarLabelsListId');
            labelsList = web.get_lists().getById(varLabelsListId);
            labelItems = labelsList.getItems(SP.CamlQuery.createAllItemsQuery());
            hostWebContext.load(labelsList);

            var itemCreateInfo = new SP.ListItemCreationInformation();
            var labelItem = labelsList.addItem(itemCreateInfo);
            labelItem.set_item(variationLabelFieldNames.isSource, true);
            labelItem.set_item(variationLabelFieldNames.locale, locale);
            labelItem.set_item(variationLabelFieldNames.title, title);
            labelItem.set_item(variationLabelFieldNames.topWebUrl, topWebUrl);
            labelItem.set_item(variationLabelFieldNames.language, language);
            labelItem.set_item(variationLabelFieldNames.displayName, displayName);
            labelItem.set_item(variationLabelFieldNames.hierarchyIsCreated, false);
            labelItem.set_item(variationLabelFieldNames.notificationMode, false);

            //labelItem[variationLabelFieldNames.hierarchyCreationMode] = SP.HierarchyCreationMode.PublishingSitesAndAllPages;
            labelItem.update();
            hostWebContext.load(labelItems);
            hostWebContext.executeQueryAsync(onCreateLabelSuccess, onCreateLabelFail);
        },
        onCreateLabelSuccess = function () {
            var variationLabels = [];
            var e = labelItems.getEnumerator();
            while (e.moveNext()) {
                var labelItem = e.get_current();
                variationLabels.push({
                    'IsSource': labelItem.get_item('Is_x0020_Source'),
                    'Language': labelItem.get_item('Language'),
                    'Locale': labelItem.get_item('Locale'),
                    'Title': labelItem.get_item('Title'),
                    'TopWebUrl': labelItem.get_item('Top_x0020_Web_x0020_URL'),
                });
            }
            writeSuccessMessage('There is/are ' + variationLabels.length + ' variation label(s) in this Site Collection');
        },
        onCreateLabelFail = function (sender, args) {
            writeErrorMessage(args.get_message());
        }

    return {
        execute: function (locale, title, topWebUrl, language, displayName) {
            createLabel(locale, title, topWebUrl, language, displayName);
        }
    }
}();