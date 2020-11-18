// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as actionSDK from "@microsoft/m365-action-sdk";
import { Logger } from "./../utils/Logger";

export class ActionSdkHelper {

    /**
     * API to fetch current action context
     */
    public static async getActionContext() {
        let request = new actionSDK.GetContext.Request();
        let response = await actionSDK.executeApi(request) as actionSDK.GetContext.Response;
        if (!response.error) {
            Logger.logInfo(`fetchCurrentContext success - Request: ${JSON.stringify(request)} Response: ${JSON.stringify(response)}`);
            return { success: true, context: response.context };
        }
        else {
            Logger.logError(`fetchCurrentContext failed, Error: ${response.error.category}, ${response.error.code}, ${response.error.message}`);
            return { success: false, error: response.error };
        }
    }

    /*
    * @desc Service Request to create new Action Instance
    * @param {actionSDK.Action} action instance which need to get created
    */
    public static async createActionInstance(action: actionSDK.Action) {
        let request = new actionSDK.CreateAction.Request(action);
        let response = await actionSDK.executeApi(request) as actionSDK.GetContext.Response;
        if (!response.error) {
            Logger.logInfo(`createActionInstance success - Request: ${JSON.stringify(request)} Response: ${JSON.stringify(response)}`);
        }
        else {
            Logger.logError(`createActionInstance failed, Error: ${response.error.category}, ${response.error.code}, ${response.error.message}`);
        }
    }

    /**
     * Function to get for data rows
     * @param actionId action instance id
     * @param createrId created id
     * @param continuationToken
     * @param pageSize
     */
    public static async getActionDataRows(actionId: string, creatorId?: string, continuationToken?: string, pageSize?: number) {
        let request = new actionSDK.GetActionDataRows.Request(actionId, creatorId, continuationToken, pageSize);
        let response = await actionSDK.executeApi(request) as actionSDK.GetActionDataRows.Response;
        if (!response.error) {
            Logger.logInfo(`getActionDataRows success - Request: ${JSON.stringify(request)} Response: ${JSON.stringify(response)}`);
            return { success: true, dataRows: response.dataRows, continuationToken: response.continuationToken };
        }
        else {
            Logger.logError(`getActionDataRows failed, Error: ${response.error.category}, ${response.error.code}, ${response.error.message}`);
            return { success: false, error: response.error };
        }
    }

    /*
    *   @desc Service API Request for getting the membercount
    *   @param subscription - action subscription: actionSDK.ActionSdkContext.subscription
    */
    public static async getSubscriptionMemberCount(subscription: actionSDK.Subscription) {
        let request = new actionSDK.GetSubscriptionMemberCount.Request(subscription);
        let response = await actionSDK.executeApi(request) as actionSDK.GetSubscriptionMemberCount.Response;
        if (!response.error) {
            Logger.logInfo(`getSubscriptionMemberCount success - Request: ${JSON.stringify(request)} Response: ${JSON.stringify(response)}`);
            return { success: true, memberCount: response.memberCount };
        }
        else {
            Logger.logError(`getSubscriptionMemberCount failed, Error: ${response.error.category}, ${response.error.code}, ${response.error.message}`);
            return { success: false, error: response.error };
        }
    }

    /*
    * @desc Service API Request for fetching action instance
    * @param {actionId} action id for which we want to get details
    */
    public static async getAction(actionId?: string) {
        let request = new actionSDK.GetAction.Request(actionId);
        let response = await actionSDK.executeApi(request) as actionSDK.GetAction.Response;
        if (!response.error) {
            Logger.logInfo(`getAction success - Request: ${JSON.stringify(request)} Response: ${JSON.stringify(response)}`);
            return { success: true, action: response.action };
        }
        else {
            Logger.logError(`getAction failed, Error: ${response.error.category}, ${response.error.code}, ${response.error.message}`);
            return { success: false, error: response.error };
        }
    }

    /**
     * Funtion to get action data summary
     * @param actionId action id
     * @param addDefaultAggregates
     */
    public static async getActionDataRowsSummary(actionId: string, addDefaultAggregates?: boolean) {
        let request = new actionSDK.GetActionDataRowsSummary.Request(actionId, addDefaultAggregates);
        let response = await actionSDK.executeApi(request) as actionSDK.GetActionDataRowsSummary.Response;
        if (!response.error) {
            Logger.logInfo(`getActionDataRowsSummary success - Request: ${JSON.stringify(request)} Response: ${JSON.stringify(response)}`);
            return { success: true, summary: response.summary };
        }
        else {
            Logger.logError(`getActionDataRowsSummary failed, Error: ${response.error.category}, ${response.error.code}, ${response.error.message}`);
            return { success: false, error: response.error };
        }

    }

    /**
     * Method to get details of member of subscription
     * @param subscription subscription
     * @param userId user id to get details
     */
    public static async getSubscriptionMembers(subscription, userIds) {
        let request = new actionSDK.GetSubscriptionMembers.Request(subscription, userIds);
        let response = await actionSDK.executeApi(request) as actionSDK.GetSubscriptionMembers.Response;
        if (!response.error) {
            Logger.logInfo(`getSubscriptionMembers success - Request: ${JSON.stringify(request)} Response: ${JSON.stringify(response)}`);
            return { success: true, members: response.members, memberIdsNotFound: response.memberIdsNotFound };
        }
        else {
            Logger.logError(`getSubscriptionMembers failed, Error: ${response.error.category}, ${response.error.code}, ${response.error.message}`);
            return { success: false, error: response.error };
        }
    }

    /**
     * @desc Service API Request for getting the nonResponders details
     * @param actionId actionId
     * @param subscriptionId subscriptionId
     */
    public static async getNonResponders(actionId: string, subscriptionId: string) {
        let request = new actionSDK.GetActionSubscriptionNonParticipants.Request(actionId, subscriptionId);
        let response = await actionSDK.executeApi(request) as actionSDK.GetActionSubscriptionNonParticipants.Response;
        if (!response.error) {
            Logger.logInfo(`getNonResponders success - Request: ${JSON.stringify(request)} Response: ${JSON.stringify(response)}`);
            return { success: true, nonParticipants: response.nonParticipants };
        }
        else {
            Logger.logError(`getNonResponders failed, Error: ${response.error.category}, ${response.error.code}, ${response.error.message}`);
            return { sucess: false, error: response.error };
        }
    }

    /**
     * Method to update action instance data
     * @param data object of data we want modify
     */
    public static async updateActionInstance(actionUpdateInfo: actionSDK.ActionUpdateInfo) {
        let request = new actionSDK.UpdateAction.Request(actionUpdateInfo);
        let response = await actionSDK.executeApi(request) as actionSDK.UpdateAction.Response;
        if (!response.error) {
            Logger.logInfo(`updateActionInstance success - Request: ${JSON.stringify(request)} Response: ${JSON.stringify(response)}`);
            return { success: true, updateSuccess: response.success };
        }
        else {
            Logger.logError(`updateActionInstance failed, Error: ${response.error.category}, ${response.error.code}, ${response.error.message}`);
            return { success: false, error: response.error };
        }
    }

    /**
     * API to close current view
     */
    public static async closeView() {
        let closeViewRequest = new actionSDK.CloseView.Request();
        await actionSDK.executeApi(closeViewRequest);
    }

    /**
     * Method to delete action instance
     * @param actionId action instance id
     */
    public static async deleteActionInstance(actionId) {
        let request = new actionSDK.DeleteAction.Request(actionId);
        let response = await actionSDK.executeApi(request) as actionSDK.DeleteAction.Response;
        if (!response.error) {
            Logger.logInfo(`deleteActionInstance success - Request: ${JSON.stringify(request)} Response: ${JSON.stringify(response)}`);
            return { success: true, deleteSuccess: response.success };
        } else {
            Logger.logError(`deleteActionInstance failed, Error: ${response.error.category}, ${response.error.code}, ${response.error.message}`);
            return { success: false, error: response.error };
        }
    }

    /**
     * API to download CSV for the current action instance summary
     * @param actionId actionID
     * @param fileName filename of csv
     */
    public static async downloadCSV(actionId, fileName) {
        let request = new actionSDK.DownloadActionDataRowsResult.Request(actionId, fileName);
        try {
            let response = actionSDK.executeApi(request);
            Logger.logInfo(`downloadCSV success - Request: ${JSON.stringify(request)} Response: ${JSON.stringify(response)}`);
            return { success: true };
        } catch (error) {
            Logger.logError(`downloadCSV failed, Error: ${error.category}, ${error.code}, ${error.message}`);
            return { success: false, error: error };
        }
    }

    /*
    * @desc Gets the localized strings in which the app is rendered
    */
    public static async getLocalizedStrings() {
        let request = new actionSDK.GetLocalizedStrings.Request();
        let response = await actionSDK.executeApi(request) as actionSDK.GetLocalizedStrings.Response;
        if (!response.error) {
            return { success: true, strings: response.strings };
        }
        else {
            Logger.logError(`fetchLocalization failed, Error: ${response.error.category}, ${response.error.code}, ${response.error.message}`);
        }
    }

    /**
     * Method to hide loading indicater
     */
    public static hideLoadingIndicator() {
        actionSDK.executeApi(new actionSDK.HideLoadingIndicator.Request());
    }
}
