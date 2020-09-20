// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import {
    addChoice, updateTitle, deleteChoice, updateChoiceText, callActionInstanceCreationAPI, updateSettings, goToPage, shouldValidateUI
} from "./../../actions/CreationActions";
import "./creation.scss";
import getStore, { Page } from "./../../store/CreationStore";
import { observer } from "mobx-react";
import { Flex, FlexItem, CircleIcon, Button, Loader, ArrowLeftIcon, SettingsIcon, Text } from "@fluentui/react-northstar";
import * as actionSDK from "@microsoft/m365-action-sdk";
import { Localizer } from "../../utils/Localizer";
import { Utils } from "../../utils/Utils";
import { ProgressState } from "./../../utils/SharedEnum";
import { ErrorView } from "../ErrorView";
import { UxUtils } from "./../../utils/UxUtils";
import { Settings, ISettingsComponentProps, ISettingsComponentStrings } from "./Settings";
import { IChoiceContainerOption, ChoiceContainer } from "../ChoiceContainer";
import { InputBox } from "../InputBox";
import { INavBarComponentProps, NavBarComponent } from "../NavBarComponent";
import { Constants } from "./../../utils/Constants";
import { ActionSdkHelper } from "../../helper/ActionSdkHelper";

/**
 * <CreationPage> component for create view of poll app
 * @observer decorator on the component this is what tells MobX to rerender the component whenever the data it relies on changes.
 */
@observer
export default class CreationPage extends React.Component<any, any> {

    private settingsFooterComponentRef: HTMLElement;
    private validationErrorQuestionRef: HTMLElement;

    render() {
        let progressState = getStore().progressState;
        if (progressState === ProgressState.NotStarted || progressState == ProgressState.InProgress) {
            return <Loader />;
        } else if (progressState === ProgressState.Failed) {
            ActionSdkHelper.hideLoadingIndicator();
            return (
                <ErrorView
                    title={Localizer.getString("GenericError")}
                    buttonTitle={Localizer.getString("Close")}
                />
            );
        } else {
            // Render View
            ActionSdkHelper.hideLoadingIndicator();
            if (UxUtils.renderingForMobile()) {
                // this will load the setting view where user can change due date and result visibility
                if (getStore().currentPage === Page.Settings) {
                    return this.renderSettingsPageForMobile();
                } else {
                    return (
                        <Flex className="body-container no-mobile-footer">
                            {this.renderChoicesSection()}
                            <div className="settings-summary-mobile-container">
                                {this.renderFooterSection(true)}
                            </div>
                        </Flex>
                    );
                }
            } else {
                if (getStore().currentPage == Page.Settings) {
                    let settingsProps: ISettingsComponentProps = {
                        ...this.getCommonSettingsProps(),
                        onBack: () => {
                            goToPage(Page.Main);
                            setTimeout(
                                function () {
                                    if (this.settingsFooterComponentRef) {
                                        this.settingsFooterComponentRef.focus();
                                    }
                                }.bind(this),
                                0
                            );
                        }
                    };
                    return <Settings {...settingsProps} />;
                } else if (getStore().currentPage == Page.Main) {
                    return (
                        <>
                            <Flex gap="gap.medium" column className="body-container">
                                {this.renderChoicesSection()}
                            </Flex>
                            {this.renderFooterSection()}
                        </>
                    );
                }
            }
        }
    }

    /**
     * Method to render the input title box and choice input box
     */
    renderChoicesSection() {
        let questionEmptyError: string;
        let optionsError: string[] = [];
        let choiceOptions = [];
        let accessibilityAnnouncementString: string = "";
        let focusChoiceOnError: boolean = false;
        // validation of title and choices that it should not be blank setting this flag to true while creating action instance only
        if (getStore().shouldValidate) {
            questionEmptyError = getStore().title == "" ? Localizer.getString("TitleBlankError") : null;
            if (getStore().options.length >= 2) {
                for (let option of getStore().options) {
                    optionsError.push((option == null || option == "") ? Localizer.getString("BlankChoiceError") : "");
                }
            }

            if (questionEmptyError) {
                accessibilityAnnouncementString = questionEmptyError;
                if (this.validationErrorQuestionRef) {
                    this.validationErrorQuestionRef.focus();
                }
            } else {
                for (let error in optionsError) {
                    if (!Utils.isEmpty(error)) {
                        accessibilityAnnouncementString = Localizer.getString("BlankChoiceError");
                        focusChoiceOnError = true;
                        break;
                    }
                }
            }
        }

        const choicePrefix = <CircleIcon outline size="small" className="choice-item-circle" disabled />;
        let i = 0;
        getStore().options.forEach((option) => {
            const choiceOption: IChoiceContainerOption = {
                value: option,
                choicePrefix: choicePrefix,
                choicePlaceholder: Localizer.getString("Choice", i + 1),
                deleteChoiceLabel: Localizer.getString("DeleteChoiceX", i + 1),
            };
            choiceOptions.push(choiceOption);
            i++;
        });
        Utils.announceText(accessibilityAnnouncementString);
        return (
            <Flex column>
                <InputBox
                    fluid multiline
                    maxLength={Constants.POLL_TITLE_MAX_LENGTH}
                    inputRef={(element) => {
                        this.validationErrorQuestionRef = element;
                    }}
                    input={{
                        className: "title-box"
                    }}
                    showError={questionEmptyError != null}
                    errorText={questionEmptyError}
                    value={getStore().title}
                    className="title-box"
                    placeholder={Localizer.getString("PollTitlePlaceholder")}
                    aria-placeholder={Localizer.getString("PollTitlePlaceholder")}
                    onChange={(e) => {
                        updateTitle((e.target as HTMLInputElement).value);
                        shouldValidateUI(false); // setting this flag to false to not validate input everytime value changes
                    }}
                />
                <div className="indentation">
                    <ChoiceContainer optionsError={optionsError} options={choiceOptions} limit={getStore().maxOptions}
                        focusOnError={focusChoiceOnError}
                        renderForMobile={UxUtils.renderingForMobile()}
                        maxLength={Constants.POLL_CHOICE_MAX_LENGTH}
                        onDeleteChoice={(i) => {
                            shouldValidateUI(false);
                            deleteChoice(i);
                        }}
                        onUpdateChoice={(i, value) => {
                            updateChoiceText(i, value);
                            shouldValidateUI(false);
                        }}
                        onAddChoice={() => {
                            addChoice();
                            shouldValidateUI(false);
                        }} />
                </div>
            </Flex>
        );
    }

    /**
     * Setting summary and button to switch from main view to settings view
     */
    renderFooterSettingsSection() {
        return (
            <div className="settings-summary-footer" {...UxUtils.getTabKeyProps()}
                ref={(element) => {
                    this.settingsFooterComponentRef = element;
                }}
                onClick={() => {
                    goToPage(Page.Settings);
                }}>
                <SettingsIcon className="settings-icon" outline={true} styles={({ theme: { siteVariables } }) => ({
                    color: siteVariables.colorScheme.brand.foreground,
                })} />
                <Text content={this.getSettingsSummary()} size="small" color="brand" />
            </div>
        );
    }

    /**
     * Settings page view for mobile
     */
    renderSettingsPageForMobile() {
        let navBarComponentProps: INavBarComponentProps = {
            title: Localizer.getString("Settings"),
            leftNavBarItem: {
                icon: <ArrowLeftIcon />,
                ariaLabel: Localizer.getString("Back"),
                onClick: () => {
                    goToPage(Page.Main);
                    setTimeout(() => {
                        if (this.settingsFooterComponentRef) {
                            this.settingsFooterComponentRef.focus();
                        }
                    }, 0);
                },
            },
        };

        return (
            <Flex className="body-container no-mobile-footer" column>
                <NavBarComponent {...navBarComponentProps} />
                <Settings {...this.getCommonSettingsProps()} />
            </Flex>
        );
    }

    /**
     * Helper function to provide footer for main page
     * @param isMobileView true or false based of whether its for mobile view or not
     */
    renderFooterSection(isMobileView?: boolean) {
        let className = isMobileView ? "" : "footer-layout";
        return (
            <Flex className={className} gap={"gap.smaller"}>
                {this.renderFooterSettingsSection()}
                <FlexItem push>
                    <Button
                        primary
                        loading={getStore().sendingAction}
                        disabled={getStore().sendingAction}
                        content={Localizer.getString("Next")}
                        onClick={() => {
                            callActionInstanceCreationAPI();
                        }}>
                    </Button>
                </FlexItem>
            </Flex>
        );
    }

    /**
     * method to get the setting summary from selected due date and result visibility
     */
    getSettingsSummary(): string {
        let settingsStrings: string[] = [];
        let dueDate = new Date(getStore().settings.dueDate);
        let resultVisibility = getStore().settings.resultVisibility;
        if (dueDate) {
            let dueDateString: string;
            let dueDateValues: number[];
            let dueIn: {} = Utils.getTimeRemaining(dueDate);
            if (dueIn[Utils.YEARS] > 0) {
                dueDateString = dueIn[Utils.YEARS] == 1 ? "DueInYear" : "DueInYears";
                dueDateValues = [dueIn[Utils.YEARS]];
            } else if (dueIn[Utils.MONTHS] > 0) {
                dueDateString = dueIn[Utils.MONTHS] == 1 ? "DueInMonth" : "DueInMonths";
                dueDateValues = [dueIn[Utils.MONTHS]];
            } else if (dueIn[Utils.WEEKS] > 0) {
                dueDateString = dueIn[Utils.WEEKS] == 1 ? "DueInWeek" : "DueInWeeks";
                dueDateValues = [dueIn[Utils.WEEKS]];
            } else if (dueIn[Utils.DAYS] > 0) {
                dueDateString = dueIn[Utils.DAYS] == 1 ? "DueInDay" : "DueInDays";
                dueDateValues = [dueIn[Utils.DAYS]];
            } else if (dueIn[Utils.HOURS] > 0 && dueIn[Utils.MINUTES] > 0) {
                if (dueIn[Utils.HOURS] == 1 && dueIn[Utils.MINUTES] == 1) {
                    dueDateString = "DueInHourAndMinute";
                } else if (dueIn[Utils.HOURS] == 1) {
                    dueDateString = "DueInHourAndMinutes";
                } else if (dueIn[Utils.MINUTES] == 1) {
                    dueDateString = "DueInHoursAndMinute";
                } else {
                    dueDateString = "DueInHoursAndMinutes";
                }
                dueDateValues = [dueIn[Utils.HOURS], dueIn[Utils.MINUTES]];
            } else if (dueIn[Utils.HOURS] > 0) {
                dueDateString = dueIn[Utils.HOURS] == 1 ? "DueInHour" : "DueInHours";
                dueDateValues = [dueIn[Utils.HOURS]];
            } else {
                dueDateString = dueIn[Utils.MINUTES] == 1 ? "DueInMinute" : "DueInMinutes";
                dueDateValues = [dueIn[Utils.MINUTES]];
            }
            settingsStrings.push(Localizer.getString(dueDateString, ...dueDateValues));
        }

        if (resultVisibility) {
            let visibilityString: string = resultVisibility == actionSDK.Visibility.All
                ? "ResultsVisibilitySettingsSummaryEveryone" : "ResultsVisibilitySettingsSummarySenderOnly";
            settingsStrings.push(Localizer.getString(visibilityString));
        }

        return settingsStrings.join(". ");
    }

    /**
     * Helper method to provide strings for settings view
     */
    getStringsForSettings(): ISettingsComponentStrings {
        let settingsComponentStrings: ISettingsComponentStrings = {
            dueBy: Localizer.getString("dueBy"),
            resultsVisibleTo: Localizer.getString("resultsVisibleTo"),
            resultsVisibleToAll: Localizer.getString("resultsVisibleToAll"),
            resultsVisibleToSender: Localizer.getString("resultsVisibleToSender"),
            datePickerPlaceholder: Localizer.getString("datePickerPlaceholder"),
            timePickerPlaceholder: Localizer.getString("timePickerPlaceholder"),
        };
        return settingsComponentStrings;
    }

    /**
     * Helper method to provide common settings props for both mobile and web view
     */
    getCommonSettingsProps() {
        return {
            resultVisibility: getStore().settings.resultVisibility,
            dueDate: getStore().settings.dueDate,
            locale: getStore().context.locale,
            renderForMobile: UxUtils.renderingForMobile(),
            strings: this.getStringsForSettings(),
            onChange: (props: ISettingsComponentProps) => {
                updateSettings(props);
            },
            onMount: () => {
                UxUtils.setFocus(document.body, Constants.FOCUSABLE_ITEMS.All);
            },
        };
    }
}
