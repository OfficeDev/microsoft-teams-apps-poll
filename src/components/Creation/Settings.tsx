// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import { UxUtils } from "./../../utils/UxUtils";
import "./Settings.scss";
import { DateTimePickerView } from "../DateTime";
import { RadioGroupMobile } from "../RadioGroupMobile";
import * as actionSDK from "@microsoft/m365-action-sdk";
import { Flex, Text, ChevronStartIcon, RadioGroup } from "@fluentui/react-northstar";
import { Localizer } from "../../utils/Localizer";

export interface ISettingsComponentProps {
    dueDate: number;
    locale?: string;
    resultVisibility: actionSDK.Visibility;
    renderForMobile?: boolean;
    strings: ISettingsComponentStrings;
    onChange?: (props: ISettingsComponentProps) => void;
    onMount?: () => void;
    onBack?: () => void;
}

export interface ISettingsComponentStrings {
    dueBy?: string;
    resultsVisibleTo?: string;
    resultsVisibleToAll?: string;
    resultsVisibleToSender?: string;
    datePickerPlaceholder?: string;
    timePickerPlaceholder?: string;
}

/**
 * <Settings> Settings component of creation view of poll
 */
export class Settings extends React.PureComponent<ISettingsComponentProps> {
    private settingProps: ISettingsComponentProps;
    constructor(props: ISettingsComponentProps) {
        super(props);
    }

    componentDidMount() {
        if (this.props.onMount) {
            this.props.onMount();
        }
    }

    render() {
        this.settingProps = {
            dueDate: this.props.dueDate,
            locale: this.props.locale,
            resultVisibility: this.props.resultVisibility,
            strings: this.props.strings
        };

        if (this.props.renderForMobile) {
            return this.renderSettings();
        } else {
            return (
                <Flex className="body-container" column gap="gap.medium">
                    {this.renderSettings()}
                    {this.props.onBack && this.getBackElement()}
                </Flex>
            );
        }
    }

    /**
     * Common to render settings view for both mobile and web
     */
    private renderSettings() {
        return (
            <Flex column>
                {this.renderDueBySection()}
                {this.renderResultVisibilitySection()}
            </Flex>
        );
    }

    /**
     * Rendering due date section for settings view
     */
    private renderDueBySection() {
        // handling mobile view differently
        let className = this.props.renderForMobile ? "due-by-pickers-container date-time-equal" : "settings-indentation";
        return (
            <Flex className="settings-item-margin" role="group" aria-label={this.getString("dueBy")} column gap="gap.smaller">
                <label className="settings-item-title">{this.getString("dueBy")}</label>
                <div className={className}>
                    <DateTimePickerView
                        minDate={new Date()}
                        locale={this.props.locale}
                        value={new Date(this.props.dueDate)}
                        placeholderDate={this.getString("datePickerPlaceholder")}
                        placeholderTime={this.getString("timePickerPlaceholder")}
                        renderForMobile={this.props.renderForMobile}
                        onSelect={(date: Date) => {
                            this.settingProps.dueDate = date.getTime();
                            this.props.onChange(this.settingProps);
                        }} />
                </div>
            </Flex>
        );
    }

    /**
     * Rendering result visiblity radio button
     */
    private renderResultVisibilitySection() {

        let radioProps = {
            checkedValue: this.settingProps.resultVisibility,
            items: this.getVisibilityItems(this.getString("resultsVisibleToAll"), this.getString("resultsVisibleToSender")),
        };

        // handling radio group differently for mobile by using custom RadioGroupMobile component
        let radioComponent = this.props.renderForMobile ?
            <RadioGroupMobile {...radioProps} onCheckedValueChange={(value) => {
                this.settingProps.resultVisibility = value as actionSDK.Visibility;
                this.props.onChange(this.settingProps);
            }}></RadioGroupMobile> :
            <RadioGroup vertical {...radioProps} onCheckedValueChange={(e, props) => {
                this.settingProps.resultVisibility = props.value as actionSDK.Visibility;
                this.props.onChange(this.settingProps);
            }} />;

        return (
            <Flex
                className="settings-item-margin"
                role="group"
                aria-label={this.getString("resultsVisibleTo")}
                column gap="gap.smaller">
                <label className="settings-item-title">{this.getString("resultsVisibleTo")}</label>
                <div className="settings-indentation">
                    {
                        radioComponent
                    }
                </div>
            </Flex>
        );
    }

    /**
     * Footer for settings view
     */
    private getBackElement() {
        return (
            <Flex className="footer-layout" gap={"gap.smaller"}>
                <Flex vAlign="center" className="pointer-cursor" {...UxUtils.getTabKeyProps()}
                    onClick={() => {
                        this.props.onBack();
                    }}
                >
                    <ChevronStartIcon xSpacing="after" size="small" />
                    <Text content={Localizer.getString("Back")} />
                </Flex>
            </Flex>
        );
    }

    private getString(key: string): string {
        if (this.props.strings && this.props.strings.hasOwnProperty(key)) {
            return this.props.strings[key];
        }
        return key;
    }

    private getVisibilityItems(resultsVisibleToAllLabel: string, resultsVisibleToSenderLabel: string) {
        return [
            {
                key: "1",
                label: resultsVisibleToAllLabel,
                value: actionSDK.Visibility.All,
                className: "settings-radio-item"
            },
            {
                key: "2",
                label: resultsVisibleToSenderLabel,
                value: actionSDK.Visibility.Sender,
                className: "settings-radio-item-last"
            },
        ];
    }
}
