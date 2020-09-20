// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { createStore } from "satcheljs";
import * as actionSDK from "@microsoft/m365-action-sdk";
import { Utils } from "../utils/Utils";
import { ISettingsComponentProps } from "./../components/Creation/Settings";
import { ProgressState } from "./../utils/SharedEnum";
import "./../orchestrators/CreationOrchestrator";
import "./../mutator/CreationMutator";

/**
 * Creation view store containing all the required data
 */

/**
 * Enum for two main component of creation view (main page and settings page)
 */
export enum Page {
    Main,
    Settings,
}

interface IPollCreationStore {
    context: actionSDK.ActionSdkContext;
    progressState: ProgressState;
    title: string;
    maxOptions: number;
    options: string[];
    settings: ISettingsComponentProps;
    shouldValidate: boolean;
    sendingAction: boolean;
    currentPage: Page;
}

const store: IPollCreationStore = {
    context: null,
    title: "",
    maxOptions: 10, // max choice we can have in poll
    options: ["", ""],
    settings: {
        resultVisibility: actionSDK.Visibility.All,   // result of poll will be visible to everyone
        dueDate: Utils.getDefaultExpiry(7).getTime(), // default due date for poll is one week
        strings: null,
    },
    shouldValidate: false,
    sendingAction: false,
    currentPage: Page.Main,  // change currentPage value to switch b/w diff components
    progressState: ProgressState.NotStarted
};

export default createStore<IPollCreationStore>("cerationStore", store);
