// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import { Menu, Text, Flex, Dialog } from "@fluentui/react-northstar";
import "./AdaptiveMenu.scss";

export enum AdaptiveMenuRenderStyle {
    MENU,
    ACTIONSHEET
}

export interface IAdaptiveMenuProps {
    key: string;
    content: React.ReactNode;
    menuItems: AdaptiveMenuItem[];
    renderAs: AdaptiveMenuRenderStyle;
    className?: string;
    dismissMenuAriaLabel?: string;
}

export interface IAdaptiveMenuState {
    menuOpen: boolean;
}

export interface AdaptiveMenuItem {
    key: string;
    content: React.ReactNode;
    icon?: React.ReactNode;
    onClick: (event?) => void;
    className?: string;
}

/**
 * <AdaptiveMenu> component to provide dropdown
 */
export class AdaptiveMenu extends React.Component<IAdaptiveMenuProps, IAdaptiveMenuState> {

    constructor(props) {
        super(props);
        this.state = {
            menuOpen: false
        };
    }

    render() {
        switch (this.props.renderAs) {
            case AdaptiveMenuRenderStyle.ACTIONSHEET:
                return this.getActionSheet();
            case AdaptiveMenuRenderStyle.MENU:
            default:
                return this.getMenu();
        }
    }

    getAdaptiveMenuItemComponent(menuItem) {
        return (
            <div role="menuitem" tabIndex={0}
                className="actionsheet-item-container" key={menuItem.key}
                onClick={() => { menuItem.onClick(); }}
            >
                {menuItem.icon}
                <Text className="actionsheet-item-label" content={menuItem.content} />
            </div>
        );
    }

    private getActionSheet() {
        return (
            <>
                <Flex className="actionsheet-trigger-bg" onClick={() => { this.setState({ menuOpen: !this.state.menuOpen }); }}>
                    {this.props.content}
                </Flex>
                <Dialog
                    open={this.state.menuOpen}
                    className="hide-default-dialog-container"
                    content={
                        <Flex className="actionsheet-view-bg" onClick={() => { this.setState({ menuOpen: !this.state.menuOpen }); }}>
                            {this.getDismissButtonForActionSheet()}
                            <Flex role="menu" column className="actionsheet-view-container">
                                {this.getActionSheetItems()}
                            </Flex>
                        </Flex>
                    }
                />
            </>
        );
    }

    private getActionSheetItems() {
        let actionSheetItems = [];
        this.props.menuItems.forEach((menuItem) => {
            actionSheetItems.push(this.getAdaptiveMenuItemComponent(menuItem));
        });
        return actionSheetItems;
    }

    private getDismissButtonForActionSheet() {
        // Hidden Dismiss button for accessibility
        return (
            <Flex
                className="hidden-dismiss-button"
                role="button"
                aria-hidden={false}
                tabIndex={0}
                aria-label={this.props.dismissMenuAriaLabel}
                onClick={() => {
                    this.setState({ menuOpen: !this.state.menuOpen });
                }}
            />
        );
    }

    private getMenu() {
        let menuItems: AdaptiveMenuItem[];
        menuItems = Object.assign([], this.props.menuItems);
        for (let i = 0; i < menuItems.length; i++) {
            menuItems[i].className = "menu-item " + menuItems[i].className;
        }
        return (
            <Menu
                defaultActiveIndex={0}
                className={(this.props.className ? this.props.className : "") + " menu-default"}
                items={
                    [
                        {
                            key: this.props.key,
                            "aria-hidden": true,
                            content: this.props.content,
                            className: "menu-items",
                            indicator: null,
                            menu: {
                                items: menuItems
                            }
                        }
                    ]
                }
            />
        );
    }

}
