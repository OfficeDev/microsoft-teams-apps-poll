// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import "./Shimmer.scss";

export interface IShimmerProps {
    // Profile image circular shimmer will be shown with radius 32px
    showProfilePic?: boolean;

    // Shimmer will be shown with 100% height and width given in 0th element in width prop
    fill?: boolean;

    // Number of line to be shown
    lines?: number;

    // Width of each line and default is 100% if it is not given
    width?: string[];

    // If true or not given, shimmer will be shown else the child componnent will be shown
    showShimmer?: boolean;
}

/**
 * <ShimmerContainer> component that simulates a shimmer effect for the children elements.
 */

export class ShimmerContainer extends React.PureComponent<IShimmerProps> {

    render() {
        if (this.props.showShimmer != undefined && !this.props.showShimmer) {
            return this.props.children;
        }
        return (
            <div className="shimmer-container">
                <div className="container-shimmer-child">
                    {this.props.children}
                </div>
                <div className="container-shimmer-loader">
                    {this.getShimmerLoader()}
                </div>
            </div>
        );
    }

    getShimmerLoader() {
        let lineShimmer: JSX.Element[] = [];
        if (this.props.lines) {
            for (let i = 0; i < this.props.lines; i++) {
                if (i != 0) {
                    lineShimmer.push(<div className="height20"></div>);
                }
                let width = this.props.width && (this.props.width.length > i) && this.props.width[i];
                lineShimmer.push(<div className="comment shim-br animate" style={{ width: width || "100%" }}></div>);
            }
        }
        return (
            <div className="card shim-br">
                <div className="wrapper">
                    {this.props.showProfilePic ? <div className="profilePic animate"></div> : null}
                    {this.props.fill ? <div className="comment-full animate" style={{
                        width: (this.props.width && this.props.width.length > 0 && this.props.width[0] ? this.props.width[0] : "100%")
                    }}></div> : lineShimmer}
                </div>
            </div>
        );
    }
}
