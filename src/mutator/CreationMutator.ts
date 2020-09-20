// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { mutator } from "satcheljs";
import {
    setContext, setSendingFlag, goToPage, updateTitle, updateSettings, updateChoiceText, deleteChoice, shouldValidateUI,
    addChoice, setProgressState
} from "./../actions/CreationActions";
import * as actionSDK from "@microsoft/m365-action-sdk";
import { Utils } from "../utils/Utils";
import getStore from "../store/CreationStore";

/**
 * Creation view mutators to modify store data on which create view relies
 */

mutator(setContext, (msg) => {
    const store = getStore();
    store.context = msg.context;
    if (!Utils.isEmpty(store.context.lastSessionData)) {
        const lastSessionData = store.context.lastSessionData;
        const actionInstance: actionSDK.Action = lastSessionData.action;
        getStore().title = actionInstance.dataTables[0].dataColumns[0].displayName;
        let options = actionInstance.dataTables[0].dataColumns[0].options;
        // clearing the options since it is always initialize with 2 empty options.
        getStore().options = [];
        options.forEach((option) => {
            getStore().options.push(option.displayName);
        });

        getStore().settings.resultVisibility = (actionInstance.dataTables[0].rowsVisibility === actionSDK.Visibility.Sender) ?
            actionSDK.Visibility.Sender : actionSDK.Visibility.All;

        getStore().settings.dueDate = actionInstance.expiryTime;
    }
});

mutator(setSendingFlag, () => {
    const store = getStore();
    store.sendingAction = true;
});

mutator(goToPage, (msg) => {
    const store = getStore();
    store.currentPage = msg.page;
});

mutator(addChoice, () => {
    const store = getStore();
    const optionsCopy = [...store.options];
    optionsCopy.push("");
    store.options = optionsCopy;
});

mutator(shouldValidateUI, (msg) => {
    const store = getStore();
    store.shouldValidate = msg.shouldValidate;
});

mutator(deleteChoice, (msg) => {
    const store = getStore();
    const optionsCopy = [...store.options];
    optionsCopy.splice(msg.index, 1);
    store.options = optionsCopy;
});

mutator(updateChoiceText, (msg) => {
    const store = getStore();
    const optionsCopy = [...store.options];
    optionsCopy[msg.index] = msg.text;
    store.options = optionsCopy;
});

mutator(updateTitle, (msg) => {
    const store = getStore();
    store.title = msg.title;
});

mutator(updateSettings, (msg) => {
    const store = getStore();
    store.settings = msg.settingProps;
});

mutator(setProgressState, (msg) => {
    const store = getStore();
    store.progressState = msg.state;
});
