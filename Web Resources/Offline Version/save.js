/***************************************
 * HELPERS
 ***************************************/
function isOffline() {
    try {
        if (Xrm.Utility.getGlobalContext().client.isOffline()) return true;
        if (Xrm.Utility.getGlobalContext().client.getClientState() === "Offline") return true;
    } catch (e) { }
    return false;
}

function getUserLanguage() {
    return Xrm.Utility.getGlobalContext().userSettings.languageId === 1025 ? "ar" : "en";
}

function setSectionVisible(formContext, tabName, sectionName, visible) {
    try {
        var tab = formContext.ui.tabs.get(tabName);
        if (!tab) return;
        var section = tab.sections.get(sectionName);
        if (section) section.setVisible(visible);
    } catch (e) {
        console.warn("[setSectionVisible] " + tabName + " > " + sectionName + ": " + e.message);
    }
}

/***************************************
 * SECTION VISIBILITY FROM CONFIG
 * Works online + offline:
 * - Online:  _duc_incidenttype_value eq <GUID>
 * - Offline: duc_incidenttype eq <GUID>
 ***************************************/
async function toggleSectionFromConfigByCode(formContext, tabName, sectionName, configCode) {
    try {
        var isOff = isOffline();

        // Hide section by default
        setSectionVisible(formContext, tabName, sectionName, false);

        // Step 1: Get Work Order Incident lookup
        var woiAttr = formContext.getAttribute("msdyn_workorderincident");
        if (!woiAttr || !woiAttr.getValue()) {
            console.warn("[toggleSection] No Work Order Incident on form for section: " + sectionName);
            return;
        }
        var woiId = woiAttr.getValue()[0].id.replace(/[{}]/g, "");

        // Step 2: Retrieve Incident Type from Work Order Incident
        // Online uses _msdyn_incidenttype_value, offline reads it from the record object directly
        var woiRecord = await Xrm.WebApi.retrieveRecord(
            "msdyn_workorderincident",
            woiId,
            "?$select=_msdyn_incidenttype_value"
        );

        var incidentTypeId = woiRecord._msdyn_incidenttype_value;
        if (!incidentTypeId) {
            console.warn("[toggleSection] No Incident Type on Work Order Incident.");
            return;
        }
        incidentTypeId = incidentTypeId.replace(/[{}]/g, "");

        // Step 3: Query config — use schema field for offline, _value alias for online
        var incidentTypeField = isOff ? "duc_incidenttype" : "_duc_incidenttype_value";

        var query = "?$select=duc_actionvalue" +
            "&$filter=duc_code eq '" + configCode + "'" +
            " and " + incidentTypeField + " eq " + incidentTypeId;

        var result = await Xrm.WebApi.retrieveMultipleRecords("duc_incidenttypeconfigurations", query);

        if (!result.entities || result.entities.length === 0) {
            console.warn("[toggleSection] No config found for code: " + configCode);
            return;
        }

        var show = (result.entities[0].duc_actionvalue === 1);

        // Special rule: Penalties section requires duc_question1 = "مخالفة"
        if (show && sectionName === "Work_Order_Penalties_Section") {
            var q1 = formContext.getAttribute("duc_question1");
            var q1Val = q1 ? (q1.getValue() || "") : "";
            show = q1Val.toString().trim().indexOf("مخالفة") > -1;
        }

        setSectionVisible(formContext, tabName, sectionName, show);

    } catch (e) {
        var errMsg = "[toggleSectionFromConfigByCode] Error for code: " + configCode
            + "\nSection: " + sectionName
            + "\nOffline: " + isOffline()
            + "\nError: " + (e.message || JSON.stringify(e));
        console.error(errMsg, e);
        Xrm.Navigation.openAlertDialog({ text: errMsg });
    }
}

/***************************************
 * SECTION TOGGLE: Run all sections
 ***************************************/
async function applyAllSectionVisibilities(formContext) {
    var sectionsToToggle = [
        { tab: "GeneralTab", section: "Samples_Section", code: "Work Order Service Task - Sample Section" },
        { tab: "GeneralTab", section: "Products_Production_Capacity_Section", code: "Work Order Service Task - Products Section" },
        { tab: "GeneralTab", section: "Raw_Materials_Section", code: "Work Order Service Task - Raw Materials Section" },
        { tab: "GeneralTab", section: "Risk_Level_Section", code: "Work Order Service Task - Risk_Level_Section" },
        { tab: "GeneralTab", section: "Customer_Assets_Section", code: "Work Order Service Task - Work_Order_Customer_Assets_Section" },
        { tab: "GeneralTab", section: "Property_Assets_Section", code: "Work Order Service Task - Property_Assets_Section" },
        { tab: "GeneralTab", section: "Work_Order_Penalties_Section", code: "Work Order Service Task - Work_Order_Penalties_Section" },
        { tab: "GeneralTab", section: "GeneralTab_section_permitsDetails", code: "GeneralTab_section_permitsDetails" }
    ];

    // Run all in parallel for speed — each handles its own errors
    await Promise.all(
        sectionsToToggle.map(function (x) {
            return toggleSectionFromConfigByCode(formContext, x.tab, x.section, x.code);
        })
    );
}

/***************************************
 * SAVE + REFRESH
 ***************************************/
async function saveAndRefresh(formContext) {
    var step = "";
    try {
        Xrm.Utility.showProgressIndicator(
            getUserLanguage() === "ar" ? "جاري الحفظ..." : "Saving..."
        );

        // Save only if form is dirty
        step = "Saving form";
        if (formContext.data.entity.getIsDirty()) {
            await formContext.data.save();
        }

        step = "Applying section visibilities";
        await applyAllSectionVisibilities(formContext);

        step = "Running onLoad logic after save";
        await runOnLoadLogicAfterSave(formContext);

        step = "Refreshing form";
        Xrm.Utility.closeProgressIndicator();

        // Use data.refresh() instead of openForm() —
        // openForm with an entityId can fail in offline mode (platform can't locate
        // record in local cache by ID). data.refresh(false) reloads the current form
        // from the offline cache without navigating away — same pattern as wo_finish.js.
        setTimeout(function () {
            try { formContext.data.refresh(false); } catch (e) {
                console.warn("[saveAndRefresh] data.refresh failed: " + e.message);
            }
        }, 1500);

    } catch (e) {
        Xrm.Utility.closeProgressIndicator();
        var errMsg = "[saveAndRefresh] Error at step: " + step
            + "\nOffline: " + isOffline()
            + "\nError: " + (e.message || JSON.stringify(e));
        console.error(errMsg, e);
        Xrm.Navigation.openAlertDialog({ text: errMsg });
    }
}

/***************************************
 * ONLOAD LOGIC (copied from WOST/onload.js)
 * Re-applies form state after save
 ***************************************/
async function runOnLoadLogicAfterSave(formContext) {
    try {
        // 1. Hide inspection controls
        setTimeout(function () {
            try {
                var inspCtrl = formContext.getControl("msdyn_inspection");
                if (inspCtrl) inspCtrl.setVisible(false);
                var resultCtrl = formContext.getControl("msdyn_inspectiontaskresult");
                if (resultCtrl) resultCtrl.setVisible(false);
            } catch (e) {
                console.warn("[runOnLoadLogicAfterSave] Hide inspection controls: " + e.message);
            }
        }, 3000);

        // 2. Hide ribbon
        try {
            formContext.ui.headerSection.setBodyVisible(false);
            formContext.ui.headerSection.setCommandBarVisible(false);
            formContext.ui.headerSection.setTabNavigatorVisible(false);
        } catch (e) {
            console.warn("[runOnLoadLogicAfterSave] hideRibbon: " + e.message);
        }

        // 3. Disable survey form if needed
        try {
            if (formContext.getAttribute("msdyn_surveyboundedoutput")) {
                var enableForAdmin = formContext.getAttribute("duc_enableforadmin");
                if (enableForAdmin && !enableForAdmin.getValue()) {
                    disableSurveyAfterSave(formContext);
                }
            }
        } catch (e) {
            console.warn("[runOnLoadLogicAfterSave] disableSurvey: " + e.message);
        }

        // 4. HideGrid logic
        try {
            hideGridAfterSave(formContext);
        } catch (e) {
            console.warn("[runOnLoadLogicAfterSave] HideGrid: " + e.message);
        }

        // 5. Toggle sections from config
        var sectionsToToggle = [
            { tab: "GeneralTab", section: "Samples_Section", code: "Work Order Service Task - Sample Section" },
            { tab: "GeneralTab", section: "Products_Production_Capacity_Section", code: "Work Order Service Task - Products Section" },
            { tab: "GeneralTab", section: "Raw_Materials_Section", code: "Work Order Service Task - Raw Materials Section" },
            { tab: "GeneralTab", section: "Risk_Level_Section", code: "Work Order Service Task - Risk_Level_Section" },
            { tab: "GeneralTab", section: "Customer_Assets_Section", code: "Work Order Service Task - Work_Order_Customer_Assets_Section" },
            { tab: "GeneralTab", section: "Property_Assets_Section", code: "Work Order Service Task - Property_Assets_Section" },
            { tab: "GeneralTab", section: "Work_Order_Penalties_Section", code: "Work Order Service Task - Work_Order_Penalties_Section" },
            { tab: "GeneralTab", section: "VIOLATING_VEHICLES_Section", code: "Work Order Service Task - VIOLATING_VEHICLES_Section" },
            { tab: "GeneralTab", section: "GeneralTab_section_permitsDetails", code: "GeneralTab_section_permitsDetails" }
        ];

        sectionsToToggle.forEach(function (x) {
            toggleSectionFromConfigByCode_OnLoad(formContext, x.tab, x.section, x.code);
        });

        // 6. Toggle next button from inspection result
        await toggleNextButtonAfterSave(formContext);

    } catch (e) {
        console.warn("[runOnLoadLogicAfterSave] Error: " + (e.message || JSON.stringify(e)));
    }
}

/***************************************
 * ONLOAD HELPERS (adapted from WOST/onload.js
 * to work with formContext directly)
 ***************************************/
function toggleSectionFromConfigByCode_OnLoad(formContext, tabName, sectionName, configCode) {
    try {
        setSectionVisible(formContext, tabName, sectionName, false);

        var woiAttr = formContext.getAttribute("msdyn_workorderincident");
        if (!woiAttr || !woiAttr.getValue()) return;

        var woiId = woiAttr.getValue()[0].id.replace(/[{}]/g, "");

        Xrm.WebApi.retrieveRecord(
            "msdyn_workorderincident",
            woiId,
            "?$select=_msdyn_incidenttype_value"
        ).then(function (woiResult) {
            var incidentTypeId = woiResult._msdyn_incidenttype_value;
            if (!incidentTypeId) return;

            var isOff = isOffline();
            var incTypeFilterField = isOff ? "duc_incidenttype" : "_duc_incidenttype_value";

            var query = "?$select=duc_actionvalue" +
                "&$filter=duc_code eq '" + configCode + "'" +
                " and " + incTypeFilterField + " eq " + incidentTypeId;

            Xrm.WebApi.retrieveMultipleRecords("duc_incidenttypeconfigurations", query).then(
                function (result) {
                    if (result.entities.length === 0) return;

                    var show = (result.entities[0].duc_actionvalue === 1);

                    if (show && sectionName === "Work_Order_Penalties_Section") {
                        var q1 = formContext.getAttribute("duc_question1");
                        var q1Val = q1 ? (q1.getValue() || "") : "";
                        var hasViolation = (q1Val.toString().trim().indexOf("مخالفة") > -1)
                            || (q1Val.toString().trim().indexOf("غير مستوف الشروط") > -1);
                        show = hasViolation;
                    }

                    setSectionVisible(formContext, tabName, sectionName, show);
                },
                function (err) {
                    console.warn("[toggleSectionFromConfigByCode_OnLoad] " + configCode + ": " + (err.message || err));
                }
            );
        }, function (error) {
            console.warn("[toggleSectionFromConfigByCode_OnLoad] WOI retrieval: " + (error.message || error));
        });

    } catch (e) {
        console.warn("[toggleSectionFromConfigByCode_OnLoad] " + configCode + ": " + (e.message || e));
    }
}

async function toggleNextButtonAfterSave(formContext) {
    var nextCtrl = formContext.getControl("duc_nextbutton");
    if (!nextCtrl) return;

    // Check auto-calculation on incident type
    try {
        var incTypeLookup = formContext.getAttribute("duc_incidenttype");
        if (incTypeLookup) {
            var incVal = incTypeLookup.getValue();
            if (incVal && incVal.length > 0) {
                var incId = (incVal[0].id || "").replace(/[{}]/g, "");
                if (incId) {
                    var incRec = await Xrm.WebApi.retrieveRecord(
                        "msdyn_incidenttype", incId, "?$select=duc_autostatuscalculation"
                    );
                    if (incRec.duc_autostatuscalculation) {
                        nextCtrl.setVisible(true);
                        return;
                    }
                }
            }
        }
    } catch (e) { /* non-critical */ }

    // Default hide
    nextCtrl.setVisible(false);

    try {
        var woLookup = formContext.getAttribute("msdyn_workorder")?.getValue();
        if (!woLookup || woLookup.length === 0 || !woLookup[0].id) return;

        var woId = woLookup[0].id.replace(/[{}]/g, "");
        var isOff = isOffline();
        var woFilterField = isOff ? "duc_workorder" : "_duc_workorder_value";

        var res = await Xrm.WebApi.retrieveMultipleRecords(
            "duc_inspectionsurveyresult",
            "?$select=duc_answer1" +
            "&$filter=" + woFilterField + " eq " + woId +
            "&$orderby=createdon desc&$top=1"
        );

        if (!res || !res.entities || res.entities.length === 0) return;

        var answer1 = res.entities[0].duc_answer1;
        var isFilled = answer1 !== null && answer1 !== undefined && answer1.toString().trim() !== "";
        nextCtrl.setVisible(isFilled);
    } catch (e) {
        console.warn("[toggleNextButtonAfterSave] " + (e.message || e));
    }
}

function disableSurveyAfterSave(formContext) {
    if (!formContext.getAttribute("msdyn_percentcomplete") ||
        formContext.getAttribute("msdyn_percentcomplete").getValue() !== 100) {
        return;
    }

    var langId = Xrm.Utility.getGlobalContext().userSettings.languageId;
    var message = (langId === 1025) ? "جاري المعالجة..." : "Processing...";
    Xrm.Utility.showProgressIndicator(message);

    setTimeout(function () {
        try {
            // Try to find and disable survey UI elements
            var selectors = [
                '[data-control-name="msdyn_surveyboundedoutput"]',
                '.customControl[data-control-name*="survey"]'
            ];

            for (var s = 0; s < selectors.length; s++) {
                var container = document.querySelector(selectors[s]);
                if (!container) {
                    try { container = window.parent.document.querySelector(selectors[s]); } catch (e) { }
                }
                if (container) {
                    var elements = container.querySelectorAll("input, textarea, select, button");
                    elements.forEach(function (el) { el.disabled = true; });

                    var col = document.getElementsByClassName("ms-control-align");
                    Object.keys(col).forEach(function (key) { col[key].style.visibility = "hidden"; });
                    break;
                }
            }
        } catch (e) {
            console.warn("[disableSurveyAfterSave] " + e.message);
        }
        Xrm.Utility.closeProgressIndicator();
    }, 6000);
}

function hideGridAfterSave(formContext) {
    var percentAttr = formContext.getAttribute("msdyn_percentcomplete");
    if (!percentAttr) return;

    var grid = formContext.getControl("cc_1765579582463");
    var tab = formContext.ui.tabs.get("GeneralTab");
    if (!tab) return;

    var section = tab.sections.get("InspectionTerms");
    var saveButton = formContext.getControl("duc_savebutton");

    if (grid && grid.getGrid) {
        try {
            var count = grid.getGrid().getTotalRecordCount();
            if (section) section.setVisible(count > 0);
        } catch (e) { /* grid may not be loaded yet */ }
    }

    if (saveButton) {
        var percentValue = percentAttr.getValue();
        saveButton.setVisible(percentValue !== 100);
    }
}

/***************************************
 * ENTRY POINT
 ***************************************/
(async function () {
    var formContext = Xrm.Page;
    await saveAndRefresh(formContext);
})();