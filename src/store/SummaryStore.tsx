// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { createStore } from "satcheljs";
import { ProgressState } from "./../utils/SharedEnum";
import * as actionSDK from "@microsoft/m365-action-sdk";
import { Utils } from "../utils/Utils";
import "./../orchestrators/SummaryOrchectrator";
import "./../mutator/SummaryMutator";

/**
 * Summary view store containing all the required data
 */

/**
 * Enum to define three component of summary view (main page, responder and non-responder tab)
 */
export enum ViewType {
    Main,
    ResponderView,
    NonResponderView
}

export interface SummaryProgressStatus {
    actionInstance: ProgressState;
    actionInstanceSummary: ProgressState;
    memberCount: ProgressState;
    nonResponder: ProgressState;
    localizationState: ProgressState;
    actionInstanceRow: ProgressState;
    myActionInstanceRow: ProgressState;
    downloadData: ProgressState;
    closeActionInstance: ProgressState;
    deleteActionInstance: ProgressState;
    updateActionInstance: ProgressState;
    currentContext: ProgressState;
}

interface IPollSummaryStore {
    context: actionSDK.ActionSdkContext;
    actionInstance: actionSDK.Action;
    actionSummary: actionSDK.ActionDataRowsSummary;
    dueDate: number;
    currentView: ViewType;
    continuationToken: string;
    actionInstanceRows: actionSDK.ActionDataRow[];
    myRow: actionSDK.ActionDataRow;
    userProfile: { [key: string]: actionSDK.SubscriptionMember };
    nonResponders: actionSDK.SubscriptionMember[];
    memberCount: number;
    showMoreOptionsList: boolean;
    isPollCloseAlertOpen: boolean;
    isChangeExpiryAlertOpen: boolean;
    isDeletePollAlertOpen: boolean;
    progressStatus: SummaryProgressStatus;
    isActionDeleted: boolean;
}

const store: IPollSummaryStore = {
    context: null,
    actionInstance: null,
    actionSummary: null,
    myRow: null,
    dueDate: Utils.getDefaultExpiry(7).getTime(),
    currentView: ViewType.Main,
    actionInstanceRows: [],
    continuationToken: null,
    showMoreOptionsList: false,
    isPollCloseAlertOpen: false,
    isChangeExpiryAlertOpen: false,
    isDeletePollAlertOpen: false,
    userProfile: {},
    nonResponders: null,
    memberCount: null,
    progressStatus: {
        actionInstance: ProgressState.NotStarted,
        actionInstanceSummary: ProgressState.NotStarted,
        memberCount: ProgressState.NotStarted,
        nonResponder: ProgressState.NotStarted,
        localizationState: ProgressState.NotStarted,
        actionInstanceRow: ProgressState.NotStarted,
        myActionInstanceRow: ProgressState.NotStarted,
        downloadData: ProgressState.NotStarted,
        closeActionInstance: ProgressState.NotStarted,
        deleteActionInstance: ProgressState.NotStarted,
        updateActionInstance: ProgressState.NotStarted,
        currentContext: ProgressState.NotStarted,
    },
    isActionDeleted: false
};

export default createStore<IPollSummaryStore>("summaryStore", store);
