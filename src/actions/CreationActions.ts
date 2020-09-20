// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { action } from "satcheljs";
import { Page } from "../store/CreationStore";
import * as actionSDK from "@microsoft/m365-action-sdk";
import { ISettingsComponentProps } from "./../components/Creation/Settings";
import { ProgressState } from "./../utils/SharedEnum";

export enum PollCreationAction {
    initialize = "initialize",
    setContext = "setContext",
    addChoice = "addChoice",
    deleteChoice = "deleteChoice",
    updateChoiceText = "updateChoiceText",
    updateTitle = "updateTitle",
    updateSettings = "updateSettings",
    shouldValidateUI = "shouldValidateUI",
    setSendingFlag = "setSendingFlag",
    setProgressState = "setProgressState",
    goToPage = "goToPage",
    callActionInstanceCreationAPI = "callActionInstanceCreationAPI"
}

export let initialize = action(PollCreationAction.initialize);

export let callActionInstanceCreationAPI = action(PollCreationAction.callActionInstanceCreationAPI);

export let setContext = action(PollCreationAction.setContext, (context: actionSDK.ActionSdkContext) => ({
    context: context
}));

export let setSendingFlag = action(PollCreationAction.setSendingFlag);

export let goToPage = action(PollCreationAction.goToPage, (page: Page) => ({
    page: page
}));

export let addChoice = action(PollCreationAction.addChoice);

export let deleteChoice = action(PollCreationAction.deleteChoice, (index: number) => ({
    index: index
}));

export let updateChoiceText = action(PollCreationAction.updateChoiceText, (index: number, text: string) => ({
    index: index,
    text: text
}));

export let updateTitle = action(PollCreationAction.updateTitle, (title: string) => ({
    title: title
}));

export let updateSettings = action(PollCreationAction.updateSettings, (settingProps: ISettingsComponentProps) => ({
    settingProps: settingProps
}));

export let setProgressState = action(PollCreationAction.setProgressState, (state: ProgressState) => ({
    state: state
}));

export let shouldValidateUI = action(PollCreationAction.shouldValidateUI, (shouldValidate: boolean) => ({
    shouldValidate: shouldValidate
}));
