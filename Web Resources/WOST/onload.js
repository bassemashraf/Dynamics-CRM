var surveyObserver = null;
var surveyLocked = false;
const PERMIT_LOOKUP_FIELD = "duc_permit";
const PERMIT_ENTITY_NAME = "duc_permit";
const PERMIT_TYPE_FIELD = "duc_permittype";
const CITES_OPTION_VALUE = 100000002;

function onLoad(executionContext) {
    var formContext = executionContext.getFormContext();
    setTimeout(() => {
        formContext.getControl("msdyn_inspection").setVisible(false);

        formContext.getControl("msdyn_inspectiontaskresult").setVisible(false);
    }, 3000);

    hideRibbonOnLoad(executionContext); disablesurveyform(executionContext);

    HideGrid(executionContext);
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
        toggleSectionFromConfigByCode(executionContext, x.tab, x.section, x.code);
    });

    toggleNextButtonFromInspectionResult(executionContext);

    registerSurveyTabReLock(executionContext);

    //Add Permit Prefilter
    const ctrl = formContext.getControl(PERMIT_LOOKUP_FIELD);

    if (!ctrl) return;

    ctrl.removePreSearch(addPermitCitesFilter);

    ctrl.addPreSearch(addPermitCitesFilter);
}


async function toggleNextButtonFromInspectionResult(executionContext) {
    var formContext = executionContext.getFormContext();

    var nextCtrl = formContext.getControl("duc_nextbutton");
    if (!nextCtrl) return;

    var autoCalcuateIncidentType = getIncidentTypeAutoStatusCalculation(executionContext);
    if (autoCalcuateIncidentType) {
        nextCtrl.setVisible(true);
        return;
    }

    // Default hide
    nextCtrl.setVisible(false);

    try {
        // 1) Read WO from the lookup field msdyn_workorder
        var woLookup = formContext.getAttribute("msdyn_workorder")?.getValue();
        if (!woLookup || woLookup.length === 0 || !woLookup[0].id) return;

        var woId = woLookup[0].id.replace(/[{}]/g, "");

        // 2) Query Inspection Survey Result linked to this WO
        // TODO: replace with your actual Inspection Survey Result entity logical name
        var inspectionResultEntity = "duc_inspectionsurveyresult";

        // IMPORTANT:
        // This assumes the lookup on Inspection Survey Result to Work Order is ALSO named "msdyn_workorder"
        // so the Web API field is: _msdyn_workorder_value
        // If different, replace _msdyn_workorder_value with _<yourlookup>_value
        var query =
            "?$select=duc_answer1" +
            "&$filter=_duc_workorder_value eq " + woId +
            "&$orderby=createdon desc" +
            "&$top=1";

        var res = await Xrm.WebApi.retrieveMultipleRecords(inspectionResultEntity, query);
        if (!res || !res.entities || res.entities.length === 0) return;

        // 3) Check if duc_answer1 is filled
        var answer1 = res.entities[0].duc_answer1;
        var isFilled =
            answer1 !== null &&
            answer1 !== undefined &&
            answer1.toString().trim() !== "";

        nextCtrl.setVisible(isFilled);
    } catch (e) {
        console.error("toggleNextButtonFromInspectionResult error:", e);
        // keep hidden on error
    }
}


function toggleSectionFromConfigByCode(executionContext, tabName, sectionName, configCode) {

    try {
        var formContext = executionContext.getFormContext();

        var configEntity = "duc_incidenttypeconfigurations";
        var visibilityField = "duc_actionvalue";
        var codeField = "duc_code";

        // hide by default
        setSectionVisible(formContext, tabName, sectionName, false);

        // 1) Work Order Incident
        var woiAttr = formContext.getAttribute("msdyn_workorderincident");
        if (!woiAttr || !woiAttr.getValue()) {
            console.warn("No Work Order Incident on form");
            return;
        }

        var woiId = woiAttr.getValue()[0].id.replace(/[{}]/g, "");

        // 2) retrieve Incident Type
        Xrm.WebApi.retrieveRecord(
            "msdyn_workorderincident",
            woiId,
            "?$select=_msdyn_incidenttype_value"
        ).then(function (woiResult) {

            var incidentTypeId = woiResult._msdyn_incidenttype_value;
            if (!incidentTypeId) {
                console.warn("No Incident Type on Work Order Incident");
                return;
            }

            // 3) config by code + incident type
            var query =
                "?$select=" + visibilityField +
                "&$filter=" + codeField + " eq '" + configCode + "'" +
                " and _duc_incidenttype_value eq " + incidentTypeId;

            Xrm.WebApi.retrieveMultipleRecords(configEntity, query).then(
                function success(result) {

                    if (result.entities.length === 0) {
                        console.warn("No matching configuration found");
                        return;
                    }

                    var show = (result.entities[0][visibilityField] === 1);

                    // CONDITION ONLY FOR Work_Order_Penalties_Section
                    if (show && tabName === "GeneralTab" && sectionName === "Work_Order_Penalties_Section") {
                        var q1 = formContext.getAttribute("duc_question1");
                        var q1Val = q1 ? (q1.getValue() || "") : "";
                        var hasViolation = (q1Val.toString().trim().indexOf("مخالفة") > -1) || (q1Val.toString().trim().indexOf("غير مستوف الشروط") > -1);

                        show = hasViolation; // must be true
                    }

                    setSectionVisible(formContext, tabName, sectionName, show);
                },
                function error(err) {
                    console.error("Config retrieve error:", err.message);
                }
            );

        }, function (error) {
            console.error("WOI retrieve error:", error.message);
        });

    } catch (e) {
        console.error("toggleSectionFromConfigByCode error:", e);
    }
}



function hideRibbonOnLoad(executionContext) {
    var formContext = executionContext.getFormContext();

    formContext.ui.headerSection.setBodyVisible(false);

    formContext.ui.headerSection.setCommandBarVisible(false);

    formContext.ui.headerSection.setTabNavigatorVisible(false);
}

function disablesurveyform(executionContext) {
    var f = executionContext.getFormContext();

    if (!f.getAttribute("msdyn_surveyboundedoutput")) {
        console.warn("Survey field not available yet.");
        return;
    }

    if (!f.getAttribute("duc_enableforadmin").getValue()) {
        disablesurvey(executionContext);
    }

    var tab = f.ui.tabs.get("GeneralTab");
    tab.addTabStateChange(disablesurvey);

}

function disablesurvey(executionContext) {
    var f = executionContext.getFormContext();
    if (!f.getAttribute("msdyn_percentcomplete") || f.getAttribute("msdyn_percentcomplete").getValue() !== 100) {
        return;
    }

    surveyLocked = true;

    var globalContext = Xrm.Utility.getGlobalContext();
    var langId = globalContext.userSettings.languageId;
    var message = (langId === 1025) ? "جاري المعالجة..." : "Processing...";

    Xrm.Utility.showProgressIndicator(message);

    setTimeout(function () {
        lockInspectionSurveyWithObserver();

        Xrm.Utility.closeProgressIndicator();
    }, 6000);
}

function lockInspectionSurveyWithObserver() {
    if (!surveyLocked || surveyObserver) return;

    //Xrm.Utility.showProgressIndicator("Please Wait");
    let attempts = 0;
    const maxAttempts = 50;

    function attemptDisable() {
        // Try multiple selectors and parent contexts
        const selectors = [
            ".MscrmControls\\\\.InspectionControls\\\\.SurveyControl",
            ".InspectionControls\\\\.SurveyControl",
            '[data-control-name="msdyn_surveyboundedoutput"]',
            '.customControl[data-control-name*="survey"]'
        ];

        for (const sel of selectors) {
            // Current document
            let container = document.querySelector(sel);
            if (container) {
                disableSurveyUI(container);
                cleanup();
                return;
            }

            // Parent document (form header/quick view)
            try {
                container = window.parent.document.querySelector(sel);
                if (container) {
                    disableSurveyUI(container);
                    cleanup();
                    return;
                }
            } catch (e) {
                Xrm.Utility.closeProgressIndicator();
            }

            // Top window (full app)
            try {
                container = window.top.document.querySelector(sel);
                if (container) {
                    disableSurveyUI(container);
                    cleanup();
                    return;
                }
            } catch (e) {
                Xrm.Utility.closeProgressIndicator();
            }
        }

        if (attempts++ < maxAttempts) {
            setTimeout(attemptDisable, 100);
        } else {
            console.warn("Survey not found after polling");
            Xrm.Utility.closeProgressIndicator();
        }
    }

    function cleanup() {
        if (surveyObserver) {
            surveyObserver.disconnect();
            surveyObserver = null;
        }
        //Xrm.Utility.closeProgressIndicator();
    }

    // Start aggressive polling
    attemptDisable();

    // Backup: Observe entire document for survey class additions
    surveyObserver = new MutationObserver((mutations) => {
        for (const mutation of mutations) {
            if (mutation.type === 'childList') {
                for (const node of mutation.addedNodes) {
                    if (node.nodeType === 1) { // Element
                        const container = node.matches?.(".MscrmControls\\\\.InspectionControls\\\\.SurveyControl, .InspectionControls\\\\.SurveyControl")
                            || node.querySelector?.(".MscrmControls\\\\.InspectionControls\\\\.SurveyControl, .InspectionControls\\\\.SurveyControl");
                        if (container) {
                            disableSurveyUI(container);
                            cleanup();
                            return;
                        }
                    }
                }
            }
        }
    });

    // Observe key areas: body, form container, tabs
    surveyObserver.observe(document.body, {
        childList: true,
        subtree: true
    });

    // Also observe form controls area if available
    const formContext = Xrm.Page?.getFormContext?.() || (typeof GetGlobalContext === 'function' ? GetGlobalContext() : null);
    if (formContext) {
        try {
            const controlsArea = document.querySelector('.form-selector > div, .tab-canvas');
            if (controlsArea) {
                surveyObserver.observe(controlsArea, { childList: true, subtree: true });
            }
        } catch (e) {
            Xrm.Utility.closeProgressIndicator();
        }
    }

    Xrm.Utility.closeProgressIndicator();
}

function disableSurveyUI(container) {

    var elements = container.querySelectorAll(
        "input, textarea, select, button"
    );

    elements.forEach(function (el) {
        el.disabled = true;
    });

    hideSurveyButtons();
}

function hideSurveyButtons() {
    var col = document.getElementsByClassName("ms-control-align");
    Object.keys(col).forEach(key => col[key].style.visibility = "hidden");

    col = document.getElementsByClassName("ms-Button ms-Button--default default-Button root-266");
    Object.keys(col).forEach(key => col[key].style.visibility = "hidden");
}

function registerSurveyTabReLock(executionContext) {
    var formContext = executionContext.getFormContext();
    var tab = formContext.ui.tabs.get("GeneralTab");

    if (!tab) return;

    tab.addTabStateChange(function () {
        if (surveyLocked) {
            lockInspectionSurveyWithObserver();
        }
    });
}

function HideGrid(executionContext) {
    var formContext = executionContext.getFormContext();

    var percentAttr = formContext.getAttribute("msdyn_percentcomplete");
    if (!percentAttr) return;

    var grid = formContext.getControl("cc_1765579582463");
    var tab = formContext.ui.tabs.get("GeneralTab");
    var section = tab.sections.get("InspectionTerms");
    var saveButton = formContext.getControl("duc_savebutton");

    // ---- Grid logic ----
    if (grid && grid.getGrid) {
        grid.addOnLoad(function () {
            var count = grid.getGrid().getTotalRecordCount();
            section.setVisible(count > 0);
        });
    }

    // ---- Save button visibility logic ----
    function updateSaveButton() {
        var percentValue = percentAttr.getValue();

        if (!saveButton) return;

        if (percentValue === 100) {
            saveButton.setVisible(false);
        } else {
            saveButton.setVisible(true);
        }
    }

    // Initial evaluation
    updateSaveButton();

    // Re-evaluate when percent changes
    percentAttr.addOnChange(updateSaveButton);
}

function toggleSectionFromConfig(executionContext, sectionName, tabName) {
    debugger;
    console.log("Function started");
    try {
        var formContext = executionContext.getFormContext();
        console.log("=== toggleSectionFromConfig START ===");
        console.log("Target Tab:", tabName, "Target Section:", sectionName);

        // --- CONFIG ---
        var lookupFieldName = "msdyn_workorderincident"; // lookup on current form
        var configEntity = "duc_incidenttypeconfigurations"; // new config entity
        var visibilityField = "duc_actionvalue"; // option set: 0=hide,1=show
        // --------------

        // Default: hide section
        console.log("Hiding section by default...");
        setSectionVisible(formContext, tabName, sectionName, false);

        // Get lookup attribute
        var lookupAttr = formContext.getAttribute(lookupFieldName);
        if (!lookupAttr) {
            console.warn("Lookup attribute not found:", lookupFieldName);
            return;
        }

        var lookupVal = lookupAttr.getValue();
        if (!lookupVal || lookupVal.length === 0) {
            console.log("Lookup is empty. Section remains hidden.");
            return;
        }

        var incidentTypeId = (lookupVal[0].id || "").replace(/[{}]/g, "");
        if (!incidentTypeId) {
            console.warn("Incident type ID is empty.");
            return;
        }

        console.log("Lookup value found:", lookupVal);
        console.log("Incident type ID:", incidentTypeId);

        // Retrieve the configuration for this incident type
        Xrm.WebApi.retrieveMultipleRecords(
            configEntity,
            "?$select=duc_actionvalue&$filter=_duc_incidenttype_value eq " + incidentTypeId
        )
            .then(function (result) {
                console.log("Configuration query returned:", result.entities.length, "records");
                if (result.entities.length > 0) {
                    var visibility = result.entities[0][visibilityField];
                    console.log("Retrieved visibility from config:", visibility);
                    setSectionVisible(formContext, tabName, sectionName, visibility === 1);
                    console.log("Section visibility set to:", visibility === 1 ? "VISIBLE" : "HIDDEN");
                } else {
                    console.log("No configuration found for this incident type. Section remains hidden.");
                }
            })
            .catch(function (error) {
                console.error("Error retrieving configuration:", error);
            });

    } catch (e) {
        console.error("toggleSectionFromConfig_OnLoad error:", e);
    }
}

// Helper: show/hide a section (safe)
function setSectionVisible(formContext, tabName, sectionName, isVisible) {
    try {
        var tab = formContext.ui.tabs.get(tabName);
        if (!tab) {
            console.warn("Tab not found:", tabName);
            return;
        }

        var section = tab.sections.get(sectionName);
        if (!section) {
            console.warn("Section not found:", sectionName);
            return;
        }

        section.setVisible(!!isVisible);
    }
    catch (e) {
        console.error("setSectionVisible error:", e);
    }
}

//function disablesurveyform(executionContext) {
//    var f = executionContext.getFormContext();

//    //try {
//    //    var formFactor = Xrm.Utility.getGlobalContext().client.getFormFactor();
//    //    if (formFactor > 1) return;  // Phone/Tablet
//    //} catch (ex) {
//    //    return;  // Legacy mobile
//    //}

//    if (f.data.entity.attributes.getByName("msdyn_surveyboundedoutput") != null) {
//        if (!f.getAttribute("duc_enableforadmin").getValue()) {
//            disablesurvey(executionContext);
//        }
//        var tab = f.ui.tabs.get("GeneralTab");
//        tab.addTabStateChange(disablesurvey);
//    }
//    else {
//        disablesurveyform(executionContext);
//    }
//}

//function disablesurvey(executionContext) {
//    var f = executionContext.getFormContext();
//    if (f.getAttribute("msdyn_percentcomplete") !== undefined && f.getAttribute("msdyn_percentcomplete").getValue() == 100) {
//        var elems = [];
//        if (elems.length == 0) {
//            Xrm.Utility.showProgressIndicator("Please Wait");
//            setTimeout(function () {
//                elems = window.parent.document.getElementsByClassName("customControl MscrmControls InspectionControls.SurveyControl MscrmControls.InspectionControls.SurveyControl");
//                if (elems.length == 0) {
//                    setTimeout(function () { disablesurveycontrols(); }, 1000);
//                } else {
//                    let elmtypes = ["input", "text", "textarea", "select", "radio"];
//                    for (var i = 0; i < elmtypes.length; i++) {
//                        for (var j = 0; j < elems[0].getElementsByTagName(elmtypes[i]).length; j++) {
//                            elems[0].getElementsByTagName(elmtypes[i])[j].disabled = true;
//                        }
//                    }

//                    var col = document.getElementsByClassName("ms-control-align");
//                    Object.keys(col).forEach((key) => col[key].style.visibility = 'hidden');
//                    col = document.getElementsByClassName("ms-Button ms-Button--default default-Button root-266");
//                    Object.keys(col).forEach((key) => col[key].style.visibility = 'hidden');

//                    Xrm.Utility.closeProgressIndicator();
//                }
//            }, 5000);
//        }
//    }
//}

//function disablesurveycontrols() {
//    var elems = window.parent.document.getElementsByClassName("customControl MscrmControls InspectionControls.SurveyControl MscrmControls.InspectionControls.SurveyControl");

//    if (elems.length == 0) {
//        setTimeout(function () { disablesurveycontrols(); }, 1000);
//    }
//    else {
//        let elmtypes = ["input", "text", "textarea", "select", "radio"];
//        for (var i = 0; i < elmtypes.length; i++) {
//            for (var j = 0; j < elems[0].getElementsByTagName(elmtypes[i]).length; j++) {
//                elems[0].getElementsByTagName(elmtypes[i])[j].disabled = true;
//            }
//        }
//        var col = document.getElementsByClassName("ms-control-align");
//        Object.keys(col).forEach((key) => col[key].style.visibility = 'hidden');
//        col = document.getElementsByClassName("ms-Button ms-Button--default default-Button root-266");
//        Object.keys(col).forEach((key) => col[key].style.visibility = 'hidden');

//        Xrm.Utility.closeProgressIndicator();
//    }
//}
function hideFieldOnWeb(executionContext) {
    var formContext = executionContext.getFormContext();

    var client = formContext.context.client;
    var clientType = client.getClient();

    console.log("Client Type: ", clientType);

    // Array of field names to hide
    var fieldNames = ["duc_homebutton"];

    // Loop through each field and hide/show based on client type
    fieldNames.forEach(function (fieldName) {
        var field = formContext.getControl(fieldName);

        if (field) {
            if (clientType === "Web" || clientType === "Outlook") {
                field.setVisible(false);
                console.log("Field '" + fieldName + "' hidden - Client is: " + clientType);
            } else {
                field.setVisible(true);
                console.log("Field '" + fieldName + "' visible - Client is: " + clientType);
            }
        } else {
            console.error("Field not found: " + fieldName);
        }
    });
}

function lockColumnsInEditableGrid(executionContext) {
    var rowFormContext = executionContext.getFormContext();
    if (!rowFormContext) return;

    var percentageAttr = Xrm.Page.getAttribute("msdyn_percentcomplete");

    var percentage = percentageAttr ? percentageAttr.getValue() : null;

    var columnsToLock = ["duc_questioncategory", "duc_questionname", "statuscode"];

    if (Math.round(percentage || 0) === 100) {
        columnsToLock.push("duc_applicableanswer");
        columnsToLock.push("duc_inspectorcomment");
    }
    columnsToLock.forEach(function (col) {
        var attribute = rowFormContext.getAttribute(col);
        if (attribute) {
            var control = attribute.controls.get(0);
            if (control && typeof control.setDisabled === 'function') {
                control.setDisabled(true);
            }
        }
    });
}

function addColor(rowData) {
    if (rowData && rowData !== "") {
        let gridRow = JSON.parse(rowData);
        let rowId = gridRow.RowId;
        let rowSelector = parent.document.querySelectorAll(`[row-id="${rowId}"]`)[1];
        if (rowSelector && gridRow.statuscode_Value != null) {
            switch (gridRow.statuscode_Value) {
                case 100000001:
                case "100000001": {
                    //Updated
                    rowSelector.style.backgroundColor = "lightgreen";
                    break;
                }
                case 100000002:
                case "100000002": {
                    //Not Updated
                    rowSelector.style.backgroundColor = "lightgoldenrodyellow";
                    break;
                }
                case 1:
                case "1": {
                    //Active
                    rowSelector.style.backgroundColor = "lightgoldenrodyellow";
                    break;
                }
            }
        }
    }
}



function getIncidentTypeAutoStatusCalculation(executionContext) {
    debugger;
    try {
        var formContext = executionContext.getFormContext();

        var lookupFieldName = "msdyn_workorderincident";     // lookup on current form
        var incidentTypeEntity = "msdyn_incidenttype";       // entity logical name
        var booleanFieldName = "duc_autostatuscalculation";  // boolean on Incident Type

        var lookupAttr = formContext.getAttribute(lookupFieldName);
        if (!lookupAttr) {
            console.warn("Lookup attribute not found:", lookupFieldName);
            return null;
        }

        var lookupVal = lookupAttr.getValue();
        if (!lookupVal || lookupVal.length === 0) {
            console.log("Incident Type lookup is empty.");
            return null;
        }

        var incidentTypeId = (lookupVal[0].id || "").replace(/[{}]/g, "");
        if (!incidentTypeId) {
            console.warn("Incident type ID is empty.");
            return null;
        }

        // Retrieve the Incident Type record to get the boolean
        return Xrm.WebApi.retrieveRecord(
            incidentTypeEntity,
            incidentTypeId,
            "?$select=" + booleanFieldName
        ).then(function (record) {
            var val = record[booleanFieldName]; // true/false (or null if not set)
            console.log(booleanFieldName + ":", val);
            return val;
        }).catch(function (error) {
            console.error("Error retrieving Incident Type:", error);
            return null;
        });

    } catch (e) {
        console.error("getIncidentTypeAutoStatusCalculation error:", e);
        return null;
    }
}

function addPermitCitesFilter(executionContext) {
    const formContext = executionContext.getFormContext();
    const ctrl = formContext.getControl(PERMIT_LOOKUP_FIELD);
    if (!ctrl) return;

    const filterXml =
        `<filter type="and">
        <condition attribute="${PERMIT_TYPE_FIELD}" operator="eq" value="${CITES_OPTION_VALUE}" />
     </filter>`;

    ctrl.addCustomFilter(filterXml, PERMIT_ENTITY_NAME);
}