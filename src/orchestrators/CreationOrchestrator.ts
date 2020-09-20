// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { setProgressState } from "./../actions/CreationActions";
import { toJS } from "mobx";
import { Localizer } from "../utils/Localizer";
import { orchestrator } from "satcheljs";
import { setContext, initialize, callActionInstanceCreationAPI, updateTitle, updateChoiceText, setSendingFlag, shouldValidateUI } from "../actions/CreationActions";
import { ProgressState } from "../utils/SharedEnum";
import getStore from "../store/CreationStore";
import { Utils } from "../utils/Utils";
import * as actionSDK from "@microsoft/m365-action-sdk";
import { ActionSdkHelper } from "../helper/ActionSdkHelper";

/**
 * Creation view orchestrators to do API calls, perform any action on data and dispatch further actions to modify stores in case of any change
 */

function validateActionInstance(actionInstance: actionSDK.Action): boolean {
    if (actionInstance == null) { return false; }

    let dataColumns = actionInstance.dataTables[0].dataColumns;
    if (!dataColumns || dataColumns.length <= 0 || !dataColumns[0].displayName || dataColumns[0].displayName == "" ||
        !dataColumns[0].options || dataColumns[0].options.length < 2) {
        return false;
    }

    for (let option of dataColumns[0].options) {
        if (!option.displayName || option.displayName == "") {
            return false;
        }
    }
    return true;
}

/**
 * Initialization of createion view fetching action context and localization details
 */
orchestrator(initialize, async () => {
    setProgressState(ProgressState.InProgress);
    let actionContext = await ActionSdkHelper.getActionContext();
    if (actionContext.success) {
        setContext(actionContext.context);
        let response = await Localizer.initialize();
        setProgressState(response ? ProgressState.Completed : ProgressState.Failed);
    }
});

orchestrator(callActionInstanceCreationAPI, async () => {
    let actionInstance: actionSDK.Action = {
        displayName: "Poll",
        expiryTime: getStore().settings.dueDate,
        dataTables: [
            {
                name: "",
                dataColumns: [],
                attachments: [],
            },
        ],
    };

    // create poll question
    updateTitle(getStore().title.trim());

    let pollQuestion: actionSDK.ActionDataColumn = {
        name: "0",
        valueType: actionSDK.ActionDataColumnValueType.SingleOption,
        displayName: getStore().title,
    };
    actionInstance.dataTables[0].dataColumns.push(pollQuestion);
    actionInstance.dataTables[0].dataColumns[0].options = [];

    // Create poll options
    for (let index = 0; index < getStore().options.length; index++) {
        updateChoiceText(index, getStore().options[index].trim());

        let pollChoice: actionSDK.ActionDataColumnOption = {
            name: `${index}`,
            displayName: getStore().options[index],
        };
        actionInstance.dataTables[0].dataColumns[0].options.push(pollChoice);
    }

    // Set poll responses visibility
    actionInstance.dataTables[0].rowsVisibility = getStore().settings.resultVisibility === actionSDK.Visibility.Sender ?
        actionSDK.Visibility.Sender : actionSDK.Visibility.All;

    if (validateActionInstance(actionInstance)) {
        setSendingFlag();
        prepareActionInstance(actionInstance, toJS(getStore().context));
        await ActionSdkHelper.createActionInstance(actionInstance);
    } else {
        shouldValidateUI(true);
    }
});

function prepareActionInstance(actionInstance: actionSDK.Action, actionContext: actionSDK.ActionSdkContext) {
    if (Utils.isEmpty(actionInstance.id)) {
        actionInstance.id = Utils.generateGUID();
        actionInstance.createTime = Date.now();
    }
    actionInstance.updateTime = Date.now();
    actionInstance.creatorId = actionContext.userId;
    actionInstance.actionPackageId = actionContext.actionPackageId;
    actionInstance.version = actionInstance.version || 1;
    actionInstance.dataTables[0].rowsEditable = actionInstance.dataTables[0].rowsEditable || true;
    actionInstance.dataTables[0].canUserAddMultipleRows = actionInstance.dataTables[0].canUserAddMultipleRows || false;
    actionInstance.dataTables[0].rowsVisibility = actionInstance.dataTables[0].rowsVisibility || actionSDK.Visibility.All;

    let isPropertyExists: boolean = false;

    if (actionInstance.customProperties && actionInstance.customProperties.length > 0) {
        for (let property of actionInstance.customProperties) {
            if (property.name == "Locale") {
                isPropertyExists = true;
            }
        }
    }

    if (!isPropertyExists) {
        actionInstance.customProperties = actionInstance.customProperties || [];
        actionInstance.customProperties.push({
            name: "Locale",
            valueType: actionSDK.ActionPropertyValueType.Text,
            value: actionContext.locale,
        });
    }
}
