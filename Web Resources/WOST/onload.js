var surveyObserver = null;
var surveyLocked = false;
var shouldLockEditableGridRow = false;
const PERMIT_LOOKUP_FIELD = "duc_permit";
const PERMIT_ENTITY_NAME = "duc_permit";
const PERMIT_TYPE_FIELD = "duc_permittype";
const CITES_OPTION_VALUE = 100000002;

async function onLoad(executionContext) {
    var formContext = executionContext.getFormContext();
    setTimeout(() => {
        formContext.getControl("msdyn_inspection").setVisible(false);

        formContext.getControl("msdyn_inspectiontaskresult").setVisible(false);
    }, 3000);

    var isOff = isOffline();

    hideRibbonOnLoad(executionContext);

    disablesurveyform(executionContext);

    await HideGrid(executionContext);
    var sectionsToToggle = [
        { tab: "GeneralTab", section: "Samples_Section", code: "Work Order Service Task - Sample Section" },
        { tab: "GeneralTab", section: "Products_Production_Capacity_Section", code: "Work Order Service Task - Products Section" },
        { tab: "GeneralTab", section: "Raw_Materials_Section", code: "Work Order Service Task - Raw Materials Section" },
        { tab: "GeneralTab", section: "Risk_Level_Section", code: "Work Order Service Task - Risk_Level_Section" },
        { tab: "GeneralTab", section: "Customer_Assets_Section", code: "Work Order Service Task - Work_Order_Customer_Assets_Section" },
        { tab: "GeneralTab", section: "Property_Assets_Section", code: "Work Order Service Task - Property_Assets_Section" },
        { tab: "GeneralTab", section: "Work_Order_Penalties_Section", code: "Work Order Service Task - Work_Order_Penalties_Section" },
        { tab: "GeneralTab", section: "VIOLATING_VEHICLES_Section", code: "Work Order Service Task - VIOLATING_VEHICLES_Section" },
        { tab: "GeneralTab", section: "Vehicle_Information_Section", code: "Work Order Service Task - Vehicle_Information_Section" },
        { tab: "GeneralTab", section: "GeneralTab_section_permitsDetails", code: "GeneralTab_section_permitsDetails" },
        { tab: "GeneralTab", section: "InspectionTerms", code: "GeneralTab_section_InspectionTerms" },
        { tab: "GeneralTab", section: "InspectionTerms_Chemical", code: "GeneralTab_section_InspectionTerms_Chemical" },
        { tab: "GeneralTab", section: "Radiation_Devices_Section", code: "Work Order Service Task - Radiation_Devices_Section" },
        { tab: "GeneralTab", section: "Additional_Inspectors", code: "Work Order Service Task - Additional_Inspectors" },
        { tab: "GeneralTab", section: "Additional_Violators_Section", code: "Work Order Service Task - Additional_Violators" },
        { tab: "GeneralTab", section: "nadeeb_section", code: "Work Order Service Task - nadeeb_section" }
    ];

    sectionsToToggle.forEach(function (x) {
        if (isOff) {
            toggleSectionFromConfigByCode_Offline(executionContext, x.tab, x.section, x.code);
        }
        else {
            toggleSectionFromConfigByCode(executionContext, x.tab, x.section, x.code);
        }

    });

    toggleNextButtonFromInspectionResult(executionContext);

    registerSurveyTabReLock(executionContext);

    // Toggle specific offline buttons

    toggleSpecificOfflineFields(executionContext, "duc_nextbutton", "duc_nextbuttonoffline");
    toggleSpecificOfflineFields(executionContext, "duc_savebutton", "duc_savebuttonoffline");


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

        // 1) Incident Type
        var incidentTypeAttr = formContext.getAttribute("duc_incidenttype");
        if (!incidentTypeAttr || !incidentTypeAttr.getValue()) {
            console.warn("No Incident Type  on form");
            return;
        }

        var incidentTypeId = incidentTypeAttr.getValue()[0].id.replace(/[{}]/g, "");
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
            async function success(result) {

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

                // CONDITION ONLY FOR Vehicle_Information_Section
                if (show && tabName === "GeneralTab" && sectionName === "Vehicle_Information_Section") {
                    var WOSTId = formContext.data.entity.getId();
                    WOSTId = WOSTId ? WOSTId.replace(/[{}]/g, "") : "";

                    var q5Val = "";

                    try {
                        var results = await Xrm.WebApi.retrieveMultipleRecords(
                            "duc_inspectionsurveyresult",
                            "?$select=duc_answer5,duc_question5&$filter=_duc_workorderservicetask_value eq " + WOSTId
                        );

                        if (results.entities.length > 0) {
                            q5Val = results.entities[0]["duc_answer5"] || "";
                        }
                    } catch (error) {
                        console.log(error.message);
                    }

                    var hasVehicle = q5Val.toString().trim().indexOf("طلب نقل نفايات") > -1;

                    show = hasVehicle;
                }

                setSectionVisible(formContext, tabName, sectionName, show);
            },
            function error(err) {
                console.error("Config retrieve error:", err.message);
            }
        );

    } catch (e) {
        console.error("toggleSectionFromConfigByCode error:", e);
    }
}

function toggleSectionFromConfigByCode_Offline(executionContext, tabName, sectionName, configCode) {

    try {

        var formContext = executionContext.getFormContext();

        var configEntity = "duc_incidenttypeconfigurations";
        var visibilityField = "duc_actionvalue";
        var codeField = "duc_code";

        // hide by default
        setSectionVisible(formContext, tabName, sectionName, false);

        // 1) Incident Type
        var incidentTypeAttr = formContext.getAttribute("duc_incidenttype");
        if (!incidentTypeAttr || !incidentTypeAttr.getValue()) {
            console.warn("No Incident Type  on form");
            return;
        }

        var incidentTypeId = incidentTypeAttr.getValue()[0].id.replace(/[{}]/g, "");

        // 2) config by code + incident type
        var query =
            "?$select=" + visibilityField +
            "&$filter=" + codeField + " eq '" + configCode + "'" +
            " and duc_incidenttype eq " + incidentTypeId;

        Xrm.WebApi.retrieveMultipleRecords(configEntity, query).then(
            async function success(result) {

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

                    //alert("Browser: has violation? " + show);
                }

                // CONDITION ONLY FOR Vehicle_Information_Section
                if (show && tabName === "GeneralTab" && sectionName === "Vehicle_Information_Section") {
                    var WOSTId = formContext.data.entity.getId();
                    WOSTId = WOSTId ? WOSTId.replace(/[{}]/g, "") : "";

                    var q5Val = "";

                    try {
                        var results = await Xrm.WebApi.retrieveMultipleRecords(
                            "duc_inspectionsurveyresult",
                            "?$select=duc_answer5,duc_question5&$filter=_duc_workorderservicetask_value eq " + WOSTId
                        );

                        if (results.entities.length > 0) {
                            q5Val = results.entities[0]["duc_answer5"] || "";
                        }
                    } catch (error) {
                        console.log(error.message);
                    }

                    var hasVehicle = q5Val.toString().trim().indexOf("طلب نقل نفايات") > -1;

                    show = hasVehicle;
                }

                setSectionVisible(formContext, tabName, sectionName, show);
            },
            function error(err) {
                console.error("Config retrieve error:", err.message);
            }
        );

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

    var globalContext = Xrm.Utility.getGlobalContext();
    var langId = globalContext.userSettings.languageId;
    var message = (langId === 1025) ? "جاري المعالجة..." : "Processing...";

    Xrm.Utility.showProgressIndicator(message);

    if (!f.getAttribute("msdyn_surveyboundedoutput")) {
        console.warn("Survey field not available yet.");
        Xrm.Utility.closeProgressIndicator();

        return;
    }

    setTimeout(function () {
        if (!f.getAttribute("duc_enableforadmin").getValue()) {
            disablesurvey(executionContext);
        }

        var tab = f.ui.tabs.get("GeneralTab");
        tab.addTabStateChange(disablesurvey);

        prepareEditableGridLocking(executionContext);

    }, 10000);
}

async function disablesurvey(executionContext) {
    var f = executionContext.getFormContext();

    var workOrderAttr = f.getAttribute("msdyn_workorder");

    var workOrderId = workOrderAttr.getValue()[0].id.replace(/[{}]/g, "");

    var currentUserId = Xrm.Utility.getGlobalContext().userSettings.userId.replace(/[{}]/g, "");

    // Only these statuses will keep the form enabled
    var allowedStatuses = [
        690970002
    ];

    await Xrm.WebApi.retrieveRecord("msdyn_workorder", workOrderId, "?$select=msdyn_systemstatus,_duc_assignedinspector_value").then(
        function (result) {
            var woStatus = result.msdyn_systemstatus;

            var assignedInspectorId = result._duc_assignedinspector_value;

            var isAllowedStatus = allowedStatuses.indexOf(woStatus) !== -1;

            var isAssignedInspector = assignedInspectorId && assignedInspectorId.toLowerCase() === currentUserId.toLowerCase();

            if (!isAllowedStatus || !isAssignedInspector) {
                surveyLocked = true;
                lockInspectionSurveyWithObserver();

                Xrm.Utility.closeProgressIndicator();
                return;
            }
        },
        function (error) {
            console.error("Failed to retrieve Work Order status: " + error.message);

            Xrm.Utility.closeProgressIndicator();
        }
    );

    if (!f.getAttribute("msdyn_percentcomplete") || f.getAttribute("msdyn_percentcomplete").getValue() !== 100) {
        Xrm.Utility.closeProgressIndicator();

        return;
    }

    surveyLocked = true;

    lockInspectionSurveyWithObserver();

    Xrm.Utility.closeProgressIndicator();
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

    hideSurveyDeleteButtons(container);

    if (!container.getAttribute("data-survey-delete-observer-attached")) {
        var surveyObserverInner = new MutationObserver(function (mutations) {
            hideSurveyDeleteButtons(container);
        });

        surveyObserverInner.observe(container, { childList: true, subtree: true });
        container.setAttribute("data-survey-delete-observer-attached", "true");
    }

    hideSurveyButtons();
}

function hideSurveyDeleteButtons(container) {
    var allElements = container.querySelectorAll("button, [role='button'], a, i, span, svg, div[class*='Button'], div[class*='button'], div[class*='Icon'], *[data-icon-name]");
    var deleteKeywords = ["delete", "remove", "حذف", "إزالة", "clear", "cancel"];

    allElements.forEach(function (btn) {
        var title = (btn.getAttribute("title") || "").toLowerCase();
        var ariaLabel = (btn.getAttribute("aria-label") || "").toLowerCase();
        var iconName = (btn.getAttribute("data-icon-name") || "").toLowerCase();
        var className = (typeof btn.className === 'string') ? btn.className.toLowerCase() : "";
        var text = (btn.innerText || btn.textContent || "").toLowerCase().trim();

        var isDeleteBtn = false;

        // Substring match for attributes
        isDeleteBtn = deleteKeywords.some(function (kw) {
            return (title && title.indexOf(kw) > -1) ||
                (ariaLabel && ariaLabel.indexOf(kw) > -1) ||
                (iconName && iconName.indexOf(kw) > -1);
        });

        // Exact or reasonable substring match for text inside small elements
        if (!isDeleteBtn && text.length > 0 && text.length < 30) {
            // Avoid matching if it explicitly contains edit or edit arabic to be super safe
            if (text.indexOf("edit") === -1 && text.indexOf("تعديل") === -1 && text.indexOf("تحرير") === -1) {
                if (deleteKeywords.some(function (kw) { return text.indexOf(kw) > -1; })) {
                    isDeleteBtn = true;
                }
            }
        }

        // Check for child icon
        if (!isDeleteBtn && btn.querySelector) {
            if (btn.querySelector('[data-icon-name="Delete"], [data-icon-name="Remove"], [data-icon-name="Cancel"], [data-icon-name="Clear"], [aria-label*="Delete"], [aria-label*="delete"], [aria-label*="حذف"]')) {
                isDeleteBtn = true;
            }
        }

        // Use class name heuristically for specific tags
        if (!isDeleteBtn && (btn.tagName.toLowerCase() === 'button' || btn.getAttribute('role') === 'button' || btn.tagName.toLowerCase() === 'i' || btn.tagName.toLowerCase() === 'svg')) {
            if (className.indexOf("delete") > -1 || className.indexOf("remove") > -1 || className.indexOf("cancel") > -1) {
                isDeleteBtn = true;
            }
        }

        if (isDeleteBtn) {
            // Avoid hiding parent container if it has too many children (e.g., misidentified container)
            if (btn.children && btn.children.length > 5) {
                return;
            }
            btn.style.setProperty("display", "none", "important");
            btn.style.setProperty("visibility", "hidden", "important");
            btn.style.setProperty("pointer-events", "none", "important");
            btn.disabled = true;
        }
    });

    // Event capture blocker on container as an ultimate fallback against dynamic clicks
    if (!container.getAttribute("data-survey-click-blocker")) {
        container.addEventListener("click", function (e) {
            var target = e.target;
            var isDeleteClick = false;

            while (target && target !== container && target.nodeType === 1) {
                var title = (target.getAttribute("title") || "").toLowerCase();
                var ariaLabel = (target.getAttribute("aria-label") || "").toLowerCase();
                var iconName = (target.getAttribute("data-icon-name") || "").toLowerCase();
                var className = (typeof target.className === 'string') ? target.className.toLowerCase() : "";
                var text = (target.innerText || target.textContent || "").toLowerCase().trim();

                if (deleteKeywords.some(function (kw) {
                    return (title && title.indexOf(kw) > -1) ||
                        (ariaLabel && ariaLabel.indexOf(kw) > -1) ||
                        (iconName && iconName.indexOf(kw) > -1);
                })) {
                    isDeleteClick = true;
                    break;
                }

                if (text.length > 0 && text.length < 30 && text.indexOf("edit") === -1 && text.indexOf("تعديل") === -1 && text.indexOf("تحرير") === -1) {
                    if (deleteKeywords.some(function (kw) { return text.indexOf(kw) > -1; })) {
                        isDeleteClick = true;
                        break;
                    }
                }

                if ((target.tagName.toLowerCase() === 'button' || target.getAttribute('role') === 'button' || target.tagName.toLowerCase() === 'i' || target.tagName.toLowerCase() === 'svg') &&
                    (className.indexOf("delete") > -1 || className.indexOf("remove") > -1 || className.indexOf("cancel") > -1)) {
                    isDeleteClick = true;
                    break;
                }

                target = target.parentNode;
            }

            if (isDeleteClick) {
                e.preventDefault();
                e.stopPropagation();
                e.stopImmediatePropagation();
            }
        }, true); // Important: capture phase
        container.setAttribute("data-survey-click-blocker", "true");
    }
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

async function HideGrid(executionContext) {
    var formContext = executionContext.getFormContext();

    var percentAttr = formContext.getAttribute("msdyn_percentcomplete");
    if (!percentAttr) return;

    var tab = formContext.ui.tabs.get("GeneralTab");
    var section = tab.sections.get("InspectionTerms");
    var saveButton = formContext.getControl("duc_savebutton");

    try {
        var entityId = formContext.data.entity.getId();
        var entityName = formContext.data.entity.getEntityName();

        if (!entityId || entityName !== 'msdyn_workorderservicetask') {
            section.setVisible(false);
            return;
        }

        // Remove curly braces
        var result = await Xrm.WebApi.retrieveMultipleRecords(
            'duc_questionanswersconfiguration',
            `?$top=1&$select=duc_questionanswersconfigurationid&$filter=_duc_msdyn_workorderservicetask_value eq ${entityId.replace(/[{}]/g, "")}`
        );

        section.setVisible(result.entities.length > 0);
    } catch (error) {
        console.error("Error checking question answers configuration:", error);
        section.setVisible(false);
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

        // Incident Type
        var incidentTypeAttr = formContext.getAttribute("duc_incidenttype");
        if (!incidentTypeAttr || !incidentTypeAttr.getValue()) {
            console.warn("No Incident Type  on form");
            return;
        }

        var incidentTypeId = incidentTypeAttr.getValue()[0].id.replace(/[{}]/g, "");
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

function lockColumnsInEditableGrid(executionContext) { // Question Answer Configuration subgrid cc_1765579582463
    var rowFormContext = executionContext.getFormContext();
    if (!rowFormContext) return;

    var percentageAttr = Xrm.Page.getAttribute("msdyn_percentcomplete");

    var percentage = percentageAttr ? percentageAttr.getValue() : null;

    var columnsToLock = ["duc_questioncategory", "duc_questionname", "statuscode", "duc_order"];

    if (Math.round(percentage || 0) === 100 || shouldLockEditableGridRow) {
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

function lockColumnsInEditableGrid_Chemical(executionContext) { // Question Answer Configuration subgrid Subgrid_new_8
    var rowFormContext = executionContext.getFormContext();
    if (!rowFormContext) return;

    var percentageAttr = Xrm.Page.getAttribute("msdyn_percentcomplete");

    var percentage = percentageAttr ? percentageAttr.getValue() : null;

    var columnsToLock = ["duc_questioncategory", "duc_questionname", "statuscode", "duc_order"];

    if (Math.round(percentage || 0) === 100 || shouldLockEditableGridRow) {
        columnsToLock.push("duc_inspectorcomments");
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

function lockWorkOrderPenaltiesGrid(executionContext) { //Work Order Penalties subgrid Subgrid_new_4
    var rowFormContext = executionContext.getFormContext();
    if (!rowFormContext) return;

    var lockForm = rowFormContext.getAttribute("duc_lockform")?.getValue();

    var columnsToLock = ["duc_penalty", "duc_lockform", "duc_penaltyorder"];

    if (lockForm || shouldLockEditableGridRow) {
        columnsToLock.push("duc_compliancestatus");
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

        // Incident Type
        var incidentTypeAttr = formContext.getAttribute("duc_incidenttype");
        if (!incidentTypeAttr || !incidentTypeAttr.getValue()) {
            console.warn("No Incident Type  on form");
            return;
        }

        var incidentTypeId = incidentTypeAttr.getValue()[0].id.replace(/[{}]/g, "");
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

var shouldLockEditableGridRow = false;

async function prepareEditableGridLocking(executionContext) {

    var formContext = executionContext.getFormContext();

    if (!formContext) return;

    var workOrderAttr = formContext.getAttribute("msdyn_workorder");

    if (
        !workOrderAttr ||
        !workOrderAttr.getValue() ||
        workOrderAttr.getValue().length === 0
    ) {
        shouldLockEditableGridRow = false;
        return;
    }

    var workOrderId = workOrderAttr.getValue()[0].id.replace(/[{}]/g, "");

    var currentUserId = Xrm.Utility.getGlobalContext().userSettings.userId.replace(/[{}]/g, "");

    // Only these statuses keep the row editable
    var allowedStatuses = [690970002];

    try {

        var result = await Xrm.WebApi.retrieveRecord("msdyn_workorder", workOrderId, "?$select=msdyn_systemstatus,_duc_assignedinspector_value");

        var woStatus = result.msdyn_systemstatus;

        var assignedInspectorId = result._duc_assignedinspector_value;

        var isAllowedStatus = allowedStatuses.indexOf(woStatus) !== -1;

        var isAssignedInspector = assignedInspectorId && assignedInspectorId.toLowerCase() === currentUserId.toLowerCase();

        shouldLockEditableGridRow = !isAllowedStatus || !isAssignedInspector;

    } catch (e) {

        console.error("Failed to retrieve Work Order details: " + e.message);

        shouldLockEditableGridRow = false;
    }
}

function isOffline() {
    try {
        if (Xrm.Utility.getGlobalContext().client.isOffline()) return true;
        if (Xrm.Utility.getGlobalContext().client.getClientState() === "Offline") return true;
    } catch (e) { }
    return false;
}

async function SubgridButtonEnableRule(f) {

    var workOrderAttr = f.getAttribute("msdyn_workorder");

    var workOrderId = workOrderAttr.getValue()[0].id.replace(/[{}]/g, "");

    // Only these statuses will keep the form enabled
    var allowedStatuses = [
        690970002
    ];

    await Xrm.WebApi.retrieveRecord("msdyn_workorder", workOrderId, "?$select=msdyn_systemstatus").then(
        function (result) {
            var woStatus = result.msdyn_systemstatus;

            var isAllowedStatus = allowedStatuses.indexOf(woStatus) !== -1;

            if (!isAllowedStatus) {

                return false;
            }
        },
        function (error) {
            console.error("Failed to retrieve Work Order status: " + error.message);
        }
    );

    if (f.getAttribute("msdyn_percentcomplete") && f.getAttribute("msdyn_percentcomplete").getValue() === 100) {

        return false;
    }

    return true;
}

function toggleSpecificOfflineFields(executionContext, onlineFieldName, offlineFieldName) {
    var formContext = executionContext.getFormContext();

    try {
        var isCurrentlyOffline = isUserOffline();

        var onlineField = formContext.getControl(onlineFieldName);
        var offlineField = formContext.getControl(offlineFieldName);

        if (isCurrentlyOffline) {
            // --- OFFLINE mode ---
            if (onlineField.getVisible()) {
                offlineField.setVisible(true);
                onlineField.setVisible(false);
                console.log("[toggleSpecificOfflineFields] Offline mode detected — showing " + offlineFieldName + ", hiding " + onlineFieldName + ".");
            }
        }
    } catch (e) {
        console.log("[toggleSpecificOfflineFields] error: " + e.message);
    }
}
function isUserOffline() {
    try {
        // Check if user is on an offline profile (works even WITH internet connection).
        // isAvailableOffline returns true when the entity is part of the active
        // Mobile Offline profile — meaning the user is on the offline-first app.
        if (Xrm.WebApi.isAvailableOffline &&
            Xrm.WebApi.isAvailableOffline("msdyn_workorder")) return true;
        if (Xrm.Utility.getGlobalContext().client.isOffline()) return true;
        if (Xrm.Utility.getGlobalContext().client.getClientState() === "Offline") return true;
    } catch (e) { }
    return false;
}