// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as uuid from "uuid";

export namespace Utils {
    export let YEARS: string = "YEARS";
    export let MONTHS: string = "MONTHS";
    export let WEEKS: string = "WEEKS";
    export let DAYS: string = "DAYS";
    export let HOURS: string = "HOURS";
    export let MINUTES: string = "MINUTES";
    export let DEFAULT_LOCALE: string = "en";

    /**
     * Method to check whether the obj param is empty or not
     * @param obj
     */
    export function isEmpty(obj: any): boolean {
        if (obj == undefined || obj == null) {
            return true;
        }

        let isEmpty = false; // isEmpty will be false if obj type is number or boolean so not adding a check for that

        if (typeof obj === "string") {
            isEmpty = (obj.trim().length == 0);
        } else if (Array.isArray(obj)) {
            isEmpty = (obj.length == 0);
        } else if (typeof obj === "object") {
            isEmpty = (JSON.stringify(obj) == "{}");
        }
        return isEmpty;
    }

    /**
     * Method to get the time diff between the time passed as param and the current time
     * @param deadLineDate
     */
    export function getTimeRemaining(deadLineDate: Date): {} {
        let now = new Date().getTime();
        let deadLineTime = deadLineDate.getTime();

        let diff = Math.abs(deadLineTime - now);
        return {
            [Utils.MINUTES] : Math.floor((diff % (1000 * 60 * 60)) / (1000 * 60)),
            [Utils.HOURS]   : Math.floor((diff % (1000 * 60 * 60 * 24)) / (1000 * 60 * 60)),
            [Utils.DAYS]    : Math.floor((diff % (1000 * 60 * 60 * 24 * 7)) / (1000 * 60 * 60 * 24)),
            [Utils.WEEKS]   : Math.floor((diff % (1000 * 60 * 60 * 24 * 30)) / (1000 * 60 * 60 * 24 * 7)),
            [Utils.MONTHS]  : Math.floor((diff % (1000 * 60 * 60 * 24 * 365)) / (1000 * 60 * 60 * 24 * 30)),
            [Utils.YEARS]   : Math.floor(diff / (1000 * 60 * 60 * 24 * 365))
        };
    }

    /**
     * Method to get the expiry date
     * @param activeDays number of days action will be active
     */
    export function getDefaultExpiry(activeDays: number): Date {
        let date: Date = new Date();
        date.setDate(date.getDate() + activeDays);

        // round off to next 30 minutes time multiple
        if (date.getMinutes() > 30) {
            date.setMinutes(0);
            date.setHours(date.getHours() + 1);
        } else {
            date.setMinutes(30);
        }
        return date;
    }

    /**
     * Method to get the unique identifier
     */
    export function generateGUID(): string {
        return uuid.v4();
    }

    /**
     * Method to download content
     * @param fileName
     * @param data
     */
    export function downloadContent(fileName: string, data: string) {
        if (data && fileName) {
            let a = document.createElement("a");
            a.href = data;
            a.download = fileName;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
        }
    }

    /**
     * Method to check whether the text direction is right-to-left or not
     * @param locale
     */
    export function isRTL(locale: string): boolean {
        let rtlLang: string[] = ["ar", "he", "fl"];
        if (locale && rtlLang.indexOf(locale.split("-")[0]) !== -1) {
            return true;
        } else {
            return false;
        }
    }

    /**
     * Method to provide accessibility
     * @param text
     */
    export function announceText(text: string) {
        let ariaLiveSpan: HTMLSpanElement = document.getElementById(
            "aria-live-span"
        );
        if (ariaLiveSpan) {
            ariaLiveSpan.innerText = text;
        } else {
            ariaLiveSpan = document.createElement("SPAN");
            ariaLiveSpan.style.cssText =
                "position: fixed; overflow: hidden; width: 0px; height: 0px;";
            ariaLiveSpan.id = "aria-live-span";
            ariaLiveSpan.innerText = "";
            ariaLiveSpan.setAttribute("aria-live", "polite");
            ariaLiveSpan.tabIndex = -1;
            document.body.appendChild(ariaLiveSpan);
            setTimeout(() => {
                ariaLiveSpan.innerText = text;
            }, 50);
        }
    }
}
