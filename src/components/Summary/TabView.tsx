// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import { ResponderView } from "./ResponderView";
import getStore, { ViewType } from "./../../store/SummaryStore";
import { NonResponderView } from "./NonResponderView";
import { Flex, ChevronStartIcon, Text, Menu, ArrowLeftIcon } from "@fluentui/react-northstar";
import { setCurrentView, goBack } from "./../../actions/SummaryActions";
import { observer } from "mobx-react";
import { Localizer } from "../../utils/Localizer";
import { Constants } from "./../../utils/Constants";
import { UxUtils } from "./../..//utils/UxUtils";
import { INavBarComponentProps, NavBarComponent } from "../NavBarComponent";

/**
 * <TabView> component that shows responder and non responder tabs
 */
@observer
export class TabView extends React.Component<any, any> {
    componentDidMount() {
        UxUtils.setFocus(document.body, Constants.FOCUSABLE_ITEMS.All);
    }

    render() {
        let participation: string =
            getStore().actionSummary.rowCount == 1
                ? Localizer.getString("ParticipationIndicatorSingular", getStore().actionSummary.rowCount, getStore().memberCount)
                : Localizer.getString("ParticipationIndicatorPlural", getStore().actionSummary.rowCount, getStore().memberCount);

        return (
            <Flex column className="tabview-container no-mobile-footer">
                {this.getNavBar()}
                <Text className="participation-title" size="small" weight="bold">
                    {participation}
                </Text>
                <Menu
                    role="tablist"
                    fluid
                    defaultActiveIndex={0}
                    items={this.getItems()}
                    underlined
                    primary
                />
                {getStore().currentView == ViewType.ResponderView ? (<ResponderView />) : (<NonResponderView />)}
                {this.getFooterElement()}
            </Flex>
        );
    }

    private getItems() {
        return [
            {
                key: "responders",
                role: "tab",
                "aria-selected": getStore().currentView == ViewType.ResponderView,
                "aria-label": Localizer.getString("Responders"),
                content: Localizer.getString("Responders"),
                onClick: () => {
                    setCurrentView(ViewType.ResponderView);
                },
            },
            {
                key: "nonResponders",
                role: "tab",
                "aria-selected": getStore().currentView == ViewType.NonResponderView,
                "aria-label": Localizer.getString("NonResponders"),
                content: Localizer.getString("NonResponders"),
                onClick: () => {
                    setCurrentView(ViewType.NonResponderView);
                },
            },
        ];
    }

    private getFooterElement() {
        if (UxUtils.renderingForMobile()) {
            return null;
        }
        return (
            <Flex className="footer-layout tab-view-footer" gap={"gap.smaller"}>
                <Flex
                    vAlign="center"
                    className="pointer-cursor"
                    {...UxUtils.getTabKeyProps()}
                    onClick={() => {
                        goBack();
                    }}
                >
                    <ChevronStartIcon xSpacing="after" size="small" />
                    <Text content={Localizer.getString("Back")} />
                </Flex>
            </Flex>
        );
    }

    private getNavBar() {
        if (!UxUtils.renderingForMobile()) {
            return null;
        }
        let navBarComponentProps: INavBarComponentProps = {
            title: Localizer.getString("ViewResponses"),
            leftNavBarItem: {
                icon: <ArrowLeftIcon />,
                ariaLabel: Localizer.getString("Back"),
                onClick: () => {
                    goBack();
                }
            },
        };

        return <NavBarComponent {...navBarComponentProps} />;
    }
}
