// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import "./summary.scss";
import getStore from "./../../store/SummaryStore";
import { observer } from "mobx-react";
import { fetchActionInstanceRows, fetchUserDetails } from "./../../actions/SummaryActions";
import { Loader, Flex, Text, FocusZone, RetryIcon, FlexItem, Avatar } from "@fluentui/react-northstar";
import * as actionSDK from "@microsoft/m365-action-sdk";
import { Utils } from "../../utils/Utils";
import { Localizer } from "../../utils/Localizer";
import { RecyclerViewType, RecyclerViewComponent } from "../RecyclerViewComponent";
import { ProgressState } from "./../../utils/SharedEnum";
import { UxUtils } from "./../../utils/UxUtils";

interface IUserInfoViewProps {
    userName: string;
    subtitle?: string;
    date?: string;
    accessibilityLabel?: string;
}

/**
 * <ResponderView> component for the responder tab
 */
@observer
export class ResponderView extends React.Component<any, any> {
    private threshHoldRow: number = 5;
    private isAnyUserProfilePending: boolean = false;
    private rowsWithUser: IUserInfoViewProps[] = [];

    componentWillMount() {
        let userIds: string[] = [];
        for (let row of getStore().actionInstanceRows) {
            userIds.push(row.creatorId);
        }
        fetchUserDetails(userIds);
        if (getStore().actionInstanceRows.length == 0) {
            fetchActionInstanceRows(true);
        }
    }

    render() {
        this.isAnyUserProfilePending = false;
        this.rowsWithUser = [];
        for (let row of getStore().actionInstanceRows) {
            this.addUserInfoProps(row);
        }

        return (
            <FocusZone className="zero-padding" isCircularNavigation={true}>
                <Flex column className="list-container" gap="gap.small">
                    <RecyclerViewComponent
                        data={this.rowsWithUser}
                        rowHeight={48}
                        showFooter={getStore().progressStatus.actionInstanceRow.toString()}
                        onRowRender={(
                            type: RecyclerViewType,
                            index: number,
                            props: IUserInfoViewProps
                        ): JSX.Element => {
                            return this.onRowRender(type, index, props);
                        }}
                    />
                </Flex>
            </FocusZone>
        );
    }

    private onRowRender(type: RecyclerViewType, index: number, userProps: IUserInfoViewProps): JSX.Element {
        if (index + this.threshHoldRow > getStore().actionInstanceRows.length &&
            getStore().progressStatus.actionInstanceRow != ProgressState.Failed) {
            fetchActionInstanceRows(true);
        }

        if (type == RecyclerViewType.Footer) {
            if (getStore().progressStatus.actionInstanceRow == ProgressState.Failed) {
                return (
                    <Flex
                        vAlign="center"
                        hAlign="center"
                        gap="gap.small"
                        {...UxUtils.getTabKeyProps()}
                        onClick={() => {
                            fetchActionInstanceRows(true);
                        }}
                    >
                        <Text content={Localizer.getString("ResponseFetchError")}></Text>
                        <RetryIcon />
                    </Flex>
                );
            } else if (getStore().progressStatus.actionInstanceRow == ProgressState.InProgress || this.isAnyUserProfilePending) {
                return <Loader />;
            }
        } else {
            return (
                <Flex aria-label={userProps.accessibilityLabel} className="user-info-view overflow-hidden" vAlign="center" gap="gap.small" {...UxUtils.getListItemProps()}>
                    <Avatar className="user-profile-pic" name={userProps.userName} size="medium" aria-hidden="true" />
                    <Flex aria-hidden={!Utils.isEmpty(userProps.accessibilityLabel)} column className="overflow-hidden">
                        <Text truncated size="medium" content={userProps.userName} />
                        <Text truncated size="small" title={userProps.subtitle} content={userProps.subtitle} />
                    </Flex>
                    <FlexItem push>
                        <Text aria-hidden={!Utils.isEmpty(userProps.accessibilityLabel)} timestamp className="nowrap" size="small" content={userProps.date} />
                    </FlexItem>
                </Flex>
            );
        }
    }

    private findSubtitle(id: string): string {
        for (let item of getStore().actionInstance.dataTables[0].dataColumns[0].options) {
            if (item.name === id) {
                return item.displayName;
            }
        }
        return null;
    }

    private addUserInfoProps(row: actionSDK.ActionDataRow): void {
        if (row && getStore().actionInstance) {
            let userProfile: actionSDK.SubscriptionMember = getStore().userProfile[row.creatorId];
            let optionId = row.columnValues[getStore().actionInstance.dataTables[0].dataColumns[0].name];
            let dateOptions: Intl.DateTimeFormatOptions = {
                year: "numeric",
                month: "long",
                day: "numeric",
                hour: "numeric",
                minute: "numeric",
            };

            let userProps: Partial<IUserInfoViewProps> = {
                subtitle: this.findSubtitle(optionId),
                date: UxUtils.formatDate(new Date(row.updateTime),
                    getStore().context ? getStore().context.locale : Utils.DEFAULT_LOCALE, dateOptions)
            };

            if (userProfile) {
                userProps.userName = userProfile.displayName || Localizer.getString("UnknownMember");
                userProps.accessibilityLabel = Localizer.getString("ResponderAccessibilityLabel",
                    userProps.userName, userProps.subtitle, userProps.date);
                this.rowsWithUser.push(userProps as IUserInfoViewProps);
            } else {
                this.isAnyUserProfilePending = true;
            }
        }
    }
}
