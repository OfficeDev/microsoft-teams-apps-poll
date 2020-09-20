// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import { observer } from "mobx-react";
import getStore, { ViewType } from "./../../store/SummaryStore";
import "./summary.scss";
import SummaryView from "./SummaryView";
import { TabView } from "./TabView";
import { Localizer } from "../../utils/Localizer";
import { ErrorView } from "../ErrorView";
import { ProgressState } from "./../../utils/SharedEnum";
import { ActionSdkHelper } from "../../helper/ActionSdkHelper";

/**
 * <SummaryPage> component to render data for summary page
 * @observer decorator on the component this is what tells MobX to rerender the component whenever the data it relies on changes.
 */
@observer
export default class SummaryPage extends React.Component<any, any> {
    render() {
        if (getStore().isActionDeleted) {
            ActionSdkHelper.hideLoadingIndicator();
            return (
                <ErrorView
                    title={Localizer.getString("PollDeletedError")}
                    subtitle={Localizer.getString("PollDeletedErrorDescription")}
                    buttonTitle={Localizer.getString("Close")}
                    image={"./images/actionDeletedError.png"}
                />
            );
        }

        let progressStatus = getStore().progressStatus;
        if (progressStatus.actionInstance == ProgressState.Failed || progressStatus.actionInstanceSummary == ProgressState.Failed ||
            progressStatus.localizationState == ProgressState.Failed || progressStatus.memberCount == ProgressState.Failed) {
            ActionSdkHelper.hideLoadingIndicator();
            return (
                <ErrorView
                    title={Localizer.getString("GenericError")}
                    buttonTitle={Localizer.getString("Close")}
                />
            );
        }

        ActionSdkHelper.hideLoadingIndicator();
        return this.getView();
    }

    /**
     * Method to return the view based on the user selection
     */
    private getView(): JSX.Element {
        let currentView = getStore().currentView;
        if (currentView == ViewType.Main) {
            return <SummaryView />;
        } else if (currentView == ViewType.ResponderView || currentView == ViewType.NonResponderView) {
            return <TabView />;
        }
    }
}
