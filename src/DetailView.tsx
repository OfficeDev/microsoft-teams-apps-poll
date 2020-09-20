// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import * as ReactDom from "react-dom";
import SummaryPage from "./components/Summary/SummaryPage";
import { initialize } from "./actions/SummaryActions";
import { ActionRootView } from "./components/ActionRootView";

initialize();
ReactDom.render(
    <ActionRootView>
        <SummaryPage />
    </ActionRootView>,
    document.getElementById("root"));
