// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import { Flex } from "@fluentui/react-northstar";
import "./NavBarComponent.scss";

export interface INavBarItem {
    title?: string;
    icon?: React.ReactNode;
    ariaLabel?: string;
    className?: string;
    onClick?: () => void;
}

export interface INavBarComponentProps {
    title: string;
    leftNavBarItem?: INavBarItem;
    rightNavBarItem?: INavBarItem;
}

/**
 * <NavBarComponent> component to provide left right nav bar
 */
export class NavBarComponent extends React.PureComponent<INavBarComponentProps> {
    render() {
        return (
            <>
                <div className="nav-container">
                    {this.getNavBarItem(this.props.leftNavBarItem)}
                    <label className="nav-title">{this.props.title}</label>
                    {this.getNavBarItem(this.props.rightNavBarItem)}
                </div>
                <div className="nav-bar-offset-height" />
            </>
        );
    }

    private getNavBarItem(navBarItem: INavBarItem) {
        if (!navBarItem) {
            return null;
        }
        let navBarItemClassName = "nav-bar-item-default" + (navBarItem.className ? " " + navBarItem.className : "");
        return (
            <Flex vAlign="center"
                className={navBarItemClassName}
                role="button"
                aria-label={navBarItem.ariaLabel}
                onClick={() => { navBarItem.onClick(); }}
                tabIndex={0}
            >
                {navBarItem.icon}
            </Flex>
        );
    }
}
