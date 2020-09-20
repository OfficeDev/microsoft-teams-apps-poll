// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import "./summary.scss";
import getStore from "../../store/SummaryStore";
import { Flex, Loader, FocusZone, Text, Avatar } from "@fluentui/react-northstar";
import { observer } from "mobx-react";
import { fetchNonReponders } from "../../actions/SummaryActions";
import { ProgressState } from "./../../utils/SharedEnum";
import { RecyclerViewComponent, RecyclerViewType } from "../RecyclerViewComponent";
import { UxUtils } from "./../../utils/UxUtils";
import { Utils } from "./../../utils/Utils";

interface IUserInfoViewProps {
    userName: string;
    accessibilityLabel?: string;
}

/**
 * <NonResponderView> component for the non-responders tab
 */
@observer
export class NonResponderView extends React.Component {
    componentWillMount() {
        fetchNonReponders();
    }

    render() {
        let rowsWithUser: IUserInfoViewProps[] = [];
        if (getStore().progressStatus.nonResponder == ProgressState.InProgress) {
            return <Loader />;
        }
        if (getStore().progressStatus.nonResponder == ProgressState.Completed) {
            for (let userProfile of getStore().nonResponders) {
                userProfile = getStore().userProfile[userProfile.id];

                if (userProfile) {
                    rowsWithUser.push({
                        userName: userProfile.displayName,
                        accessibilityLabel: userProfile.displayName,
                    });
                }
            }
        }
        return (
            <FocusZone className="zero-padding" isCircularNavigation={true}>
                <Flex column className="list-container" gap="gap.small">
                    <RecyclerViewComponent
                        data={rowsWithUser}
                        rowHeight={48}
                        onRowRender={(type: RecyclerViewType, index: number, userProps: IUserInfoViewProps): JSX.Element => {
                            return (
                                <Flex aria-label={userProps.accessibilityLabel} className="user-info-view overflow-hidden" vAlign="center" gap="gap.small" {...UxUtils.getListItemProps()}>
                                    <Avatar className="user-profile-pic" name={userProps.userName} size="medium" aria-hidden="true" />
                                    <Flex aria-hidden={!Utils.isEmpty(userProps.accessibilityLabel)} column className="overflow-hidden">
                                         <Text truncated size="medium">{userProps.userName}</Text>
                                    </Flex>
                                </Flex>
                            );
                        }}
                    />
                </Flex>
            </FocusZone>
        );
    }
}
