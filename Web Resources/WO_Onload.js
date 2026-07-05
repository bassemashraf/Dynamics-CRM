// JavaScript source code

function lockSystemStatus() {
    //debugger;
    // In some orgs the form runs in an iframe; try parent.document too
    const getDoc = () => {
        try {
            if (window.parent && window.parent.document) return window.parent.document;
        } catch (e) { }
        return document;
    };

    const selectors = [
        '[data-id="msdyn_systemstatus.fieldControl-pcf-container-id"]',
        '[data-id="msdyn_systemstatus2.fieldControl-pcf-container-id"]',
        '[data-control-name="msdyn_systemstatus"]',
        '[data-control-name="msdyn_systemstatus2"]'
    ];

    const addOverlay = () => {
        const doc = getDoc();

        let container = null;
        //for (const s of selectors) {
        //    container = doc.querySelector(s);
        //    if (container) break;
        //}
        for (const s of selectors) {
            const matches = doc.querySelectorAll(s);

            for (const m of matches) {
                // skip the Booking Card / Quick Action section
                if (!m.closest('section[data-id="bookingcard_section"]')) {
                    container = m;
                    break;
                }
            }

            if (container) break;
        }
        if (!container) return false;

        // Pick the smallest stable host for the field only (avoid FieldSectionItemContainer)
        const host =
            container.closest('[data-id$=".fieldControl-pcf-container-id"]') ||
            container;

        if (!host) return false;

        // Ensure overlay doesn't already exist
        if (host.querySelector(":scope > .wo-systemstatus-lock-overlay")) return true;

        // Make host a positioning context
        if (!host.style.position) host.style.position = "relative";

        const overlay = doc.createElement("div");
        overlay.className = "wo-systemstatus-lock-overlay";
        overlay.style.position = "absolute";
        overlay.style.top = "0";
        overlay.style.left = "0";
        overlay.style.width = "100%";
        overlay.style.height = "100%";
        overlay.style.zIndex = "999999";
        overlay.style.background = "transparent";
        overlay.style.cursor = "not-allowed";

        overlay.addEventListener(
            "pointerdown",
            (e) => {
                e.preventDefault();
                e.stopPropagation();
                e.stopImmediatePropagation();
            },
            true
        );

        overlay.addEventListener(
            "click",
            (e) => {
                e.preventDefault();
                e.stopPropagation();
                e.stopImmediatePropagation();
            },
            true
        );

        host.appendChild(overlay);

        return true;
    };

    let tries = 0;
    const interval = setInterval(() => {
        tries++;
        const ok = addOverlay();
        if (ok || tries >= 30) clearInterval(interval);
    }, 500);

    try {
        const doc = getDoc();
        const obs = new MutationObserver(() => addOverlay());
        obs.observe(doc.body, { childList: true, subtree: true });
    } catch (e) { }
}




function lockWorkOrderServiceTaskButton() {
    const getDoc = () => {
        try {
            if (window.parent && window.parent.document) return window.parent.document;
        } catch (e) { }
        return document;
    };

    const addOverlay = () => {
        const doc = getDoc();

        const buttonSelectors = [
            'button[data-id^="CC_"][role="checkbox"]',
            'button[role="checkbox"].checkbox',
            '.ServiceTabViewControl button[role="checkbox"]',
            '[data-id^="Subgrid_"] button[role="checkbox"]'
        ];

        let lockedAny = false;

        for (const s of buttonSelectors) {
            const buttons = doc.querySelectorAll(s);

            for (const button of buttons) {
                if (!button) continue;

                if (button.querySelector(":scope > .wo-servicetask-lock-overlay")) {
                    lockedAny = true;
                    continue;
                }

                const currentPosition = window.getComputedStyle(button).position;
                if (!currentPosition || currentPosition === "static") {
                    button.style.position = "relative";
                }

                const overlay = doc.createElement("div");
                overlay.className = "wo-servicetask-lock-overlay";
                overlay.style.position = "absolute";
                overlay.style.top = "0";
                overlay.style.left = "0";
                overlay.style.width = "100%";
                overlay.style.height = "100%";
                overlay.style.zIndex = "999999";
                overlay.style.background = "transparent";
                overlay.style.cursor = "not-allowed";
                overlay.style.pointerEvents = "auto";

                overlay.addEventListener(
                    "pointerdown",
                    (e) => {
                        e.preventDefault();
                        e.stopPropagation();
                        e.stopImmediatePropagation();
                    },
                    true
                );

                overlay.addEventListener(
                    "click",
                    (e) => {
                        e.preventDefault();
                        e.stopPropagation();
                        e.stopImmediatePropagation();
                    },
                    true
                );

                button.appendChild(overlay);
                lockedAny = true;
            }

            if (lockedAny) break;
        }

        return lockedAny;
    };

    let tries = 0;
    const interval = setInterval(() => {
        tries++;
        const ok = addOverlay();
        if (ok || tries >= 30) clearInterval(interval);
    }, 500);

    try {
        const doc = getDoc();
        const obs = new MutationObserver(() => addOverlay());
        obs.observe(doc.body, { childList: true, subtree: true });
    } catch (e) { }
}

function hideServiceTaskAddButton() {

    const getDoc = () => {
        try {
            if (window.parent && window.parent.document) return window.parent.document;
        } catch (e) { }
        return document;
    };

    const hideButton = () => {
        const doc = getDoc();
        const buttons = doc.querySelectorAll('.createNewButton');
        let found = false;

        for (const btn of buttons) {
            if (btn.getAttribute('aria-label') === 'New Work Order Service Task' || btn.getAttribute('aria-label') === 'مهمة خدمة أمر عمل جديد') {
                btn.style.display = 'none';
                found = true;
            }
        }
        return found;
    };

    const styleLabels = () => {
        const doc = getDoc();

        // Targets the task name text inside each service task row
        // Change selector below if you're targeting a different label
        const labels = doc.querySelectorAll('p.label');

        labels.forEach(label => {
            label.style.color = '#1160B7';
            label.style.textDecoration = 'underline';
            label.style.textUnderlineOffset = '3px';
            label.style.fontWeight = 'bold';
            label.style.cursor = 'pointer';
        });

        return labels.length > 0;
    };

    const applyAll = () => {
        const btnDone = hideButton();
        const lblDone = styleLabels();
        return btnDone && lblDone;
    };

    // Interval (500ms × 30 tries)
    let tries = 0;
    const interval = setInterval(() => {
        tries++;
        const ok = applyAll();
        if (ok || tries >= 30) clearInterval(interval);
    }, 500);

    // MutationObserver — catches dynamic renders & new rows added later
    try {
        const doc = getDoc();
        const obs = new MutationObserver(() => applyAll());
        obs.observe(doc.body, { childList: true, subtree: true });
    } catch (e) { }
}

//##########################################
//##########################################

async function toggleFinishInspectionSection(executionContext) {
    var formContext = executionContext.getFormContext();
    if (!formContext) return;

    var statusAttr = formContext.getAttribute("msdyn_systemstatus");
    if (!statusAttr) {
        console.log("msdyn_systemstatus attribute not found");
        return;
    }

    var statusValue = statusAttr.getValue();
    var isTargetStatus = Number(statusValue) === 690970002;

    console.log("statusValue =", statusValue);
    console.log("isTargetStatus =", isTargetStatus);

    var userId = Xrm.Utility.getGlobalContext().userSettings.userId.replace(/[{}]/g, "");
    var isInspector = false;

    try {
        var user = await Xrm.WebApi.retrieveRecord(
            "systemuser",
            userId,
            "?$select=duc_isinspector"
        );

        isInspector = user["duc_isinspector"] === true || user["duc_isinspector"] === 1;
        console.log("isInspector =", isInspector);

    } catch (error) {
        console.log("Error retrieving user:", error.message);
        return;
    }

    var tab = formContext.ui.tabs.get("ResponsibleEmployee");
    if (!tab) {
        console.log("Tab ResponsibleEmployee not found");
        return;
    }

    var finishSection = tab.sections.get("finishinspection");
    if (finishSection) {
        var showFinishInspection = isTargetStatus && isInspector;
        finishSection.setVisible(showFinishInspection);

        finishSection.controls.forEach(function (control) {
            try {
                if (control && typeof control.setDisabled === "function") {
                    control.setDisabled(!showFinishInspection);
                    console.log("finishinspection control:", control.getName(), "disabled =", !showFinishInspection);
                }
            } catch (e) {
                console.log("Error in finishinspection control:", e.message);
            }
        });
    }

    var targetSectionNames = [
        "ResponsibleEmployee_section_5",
        "{b8e326ee-5c21-4a18-ba55-e3b56966c249}_responsible_employee_tab_section_3",
        "{b8e326ee-5c21-4a18-ba55-e3b56966c249}_section_10"
    ];

    targetSectionNames.forEach(function (sectionName) {
        var sec = tab.sections.get(sectionName);

        if (!sec) {
            console.log("Section not found:", sectionName);
            return;
        }

        console.log("Processing section:", sectionName);

        sec.controls.forEach(function (control) {
            try {
                if (control && typeof control.setDisabled === "function") {
                    control.setDisabled(!isTargetStatus);
                    console.log(
                        "Control:", control.getName(),
                        "| Section:", sectionName,
                        "| Disabled set to:", !isTargetStatus
                    );
                }
            } catch (e) {
                console.log("Error disabling control in section " + sectionName + ":", e.message);
            }
        });
    });

    if (isTargetStatus && isInspector) {
        await toggleVehicleOwnerButton(executionContext);
        await toggleSelectEstablishmentButton(executionContext);
    } else {
        var btn1 = formContext.getControl("duc_assignvehicleowner");
        if (btn1) btn1.setVisible(false);

        var btn2 = formContext.getControl("duc_selectestablishmentaccount");
        if (btn2) btn2.setVisible(false);
    }
}

async function toggleVehicleOwnerButton(executionContext) {
    var formContext = executionContext.getFormContext();

    // Get the button control
    var btnCtrl = formContext.getControl("duc_assignvehicleowner");
    if (!btnCtrl) {
        console.log("Vehicle owner button control not found");
        return;
    }

    // Default: hide the button
    btnCtrl.setVisible(false);

    // Get service account
    var serviceAccountAttr = formContext.getAttribute("msdyn_serviceaccount");
    if (!serviceAccountAttr) return;

    var serviceAccount = serviceAccountAttr.getValue();
    if (!serviceAccount || !serviceAccount[0] || !serviceAccount[0].id) return;

    var accountId = serviceAccount[0].id.replace(/[{}]/g, "");

    try {
        // Filter for accounts with account type = 1 (Individual/Vehicle Owner)
        // or type = 3 Cabin
        //new change on 5 july 2026 , no need to appears for both 3 & 6
        var showButton = false;

        if (isOffline()) {
            // OFFLINE: navigation property filters not supported, use two-step retrieve
            var accountRecord = await Xrm.WebApi.retrieveRecord(
                "account", accountId,
                "?$select=accountid,_parentaccountid_value,_duc_newaccounttype_value"
            );
            if (accountRecord._parentaccountid_value === null && accountRecord._duc_newaccounttype_value) {
                var accountTypeRecord = await Xrm.WebApi.retrieveRecord(
                    "duc_accounttype", accountRecord._duc_newaccounttype_value,
                    "?$select=duc_accounttype"
                );
                var at = accountTypeRecord.duc_accounttype;
                showButton = (at === 1 || at === 3);
            }
        } else {
            // ONLINE: use navigation property filter
            var result = await Xrm.WebApi.retrieveMultipleRecords(
                "account",
                `?$select=accountid&$filter=accountid eq ${accountId} and _parentaccountid_value eq null and (duc_NewAccountType/duc_accounttype eq 1  or duc_NewAccountType/duc_accounttype eq 3)&$top=1`
            );
            showButton = result.entities.length > 0;
        }

        if (showButton) {
            btnCtrl.setVisible(true);
        }
    } catch (error) {
        console.error("Error checking account type for vehicle owner button:", error);
    }
}

async function toggleSelectEstablishmentButton(executionContext) {
    var formContext = executionContext.getFormContext();

    var isIndividual = false;
    // Get the button control
    var btnCtrl = formContext.getControl("duc_selectestablishmentaccount");
    if (!btnCtrl) {
        console.log("Select Establishment button control not found");
        return;
    }

    // Default: hide the button
    btnCtrl.setVisible(false);

    // Get service account
    var serviceAccountAttr = formContext.getAttribute("msdyn_serviceaccount");
    if (!serviceAccountAttr) return;

    var serviceAccount = serviceAccountAttr.getValue();
    if (!serviceAccount || !serviceAccount[0] || !serviceAccount[0].id) return;

    var accountId = serviceAccount[0].id.replace(/[{}]/g, "");


    var departmentAttr = formContext.getAttribute("duc_department");
    if (!departmentAttr) return;

    var department = departmentAttr.getValue();
    if (!department || !department[0] || !department[0].id) return;

    var departmentId = department[0].id.replace(/[{}]/g, "");

    try {
        // Filter for accounts with account type = 2 (Individual Owner)

        var result = await Xrm.WebApi.retrieveMultipleRecords(
            "account",
            `?$select=accountid&$filter=accountid eq ${accountId} and _parentaccountid_value eq null and (duc_NewAccountType/duc_accounttype eq 2)&$top=1`
        );

        if (result.entities.length > 0) {

            isIndividual = true;
        }

        var result1 = await Xrm.WebApi.retrieveMultipleRecords(
            "msdyn_organizationalunit",
            `?$select=duc_englishname&$filter=(contains(duc_englishname,'Radiation') and msdyn_organizationalunitid eq ${departmentId})&$top=1`
        );

        if (result1.entities.length > 0 && isIndividual) {

            btnCtrl.setVisible(true);
        }

    } catch (error) {
        console.error("Error checking account type for vehicle owner button:", error);
    }
}

async function toggleVehicleOwnerTab(executionContext) {
    var formContext = executionContext.getFormContext();

    // Get the tab
    var tab = formContext.ui.tabs.get("tab_17_Vehicle_Owner");
    if (!tab) {
        console.log("Vehicle Owner tab not found");
        return;
    }

    // Default: hide the tab
    tab.setVisible(false);

    // Get service account
    var serviceAccountAttr = formContext.getAttribute("msdyn_serviceaccount");
    if (!serviceAccountAttr) return;

    var serviceAccount = serviceAccountAttr.getValue();
    if (!serviceAccount || !serviceAccount[0] || !serviceAccount[0].id) return;

    var accountId = serviceAccount[0].id.replace(/[{}]/g, "");

    try {
        var result = await Xrm.WebApi.retrieveMultipleRecords(
            "account",
            `?$select=accountid&$filter=accountid eq ${accountId} 
            and _parentaccountid_value eq null 
            and (duc_NewAccountType/duc_accounttype eq 1 
                 or duc_NewAccountType/duc_accounttype eq 3 
                 or duc_NewAccountType/duc_accounttype eq 6)
            &$top=1`
        );

        // Show tab if condition matches
        if (result.entities.length > 0) {
            tab.setVisible(true);
        }

    } catch (error) {
        console.error("Error checking account type for vehicle owner tab:", error);
    }
}

function onFormLoad(executionContext) {
    lockSystemStatus();
    lockWorkOrderServiceTaskButton();
    hideServiceTaskAddButton();
    setDepartmentFromUserOU(executionContext);
    initializeResponsibleEmployeeValidations(executionContext);
    resizeSection(executionContext);
    toggleResponsibleEmployeeSections(executionContext);
    toggleSendToTawasolButton(executionContext);
    toggleVehicleOwnerTab(executionContext);
    setTimeout(function () {
        hideSubOrderTypeField(executionContext);
        //     toggleReopenWorkOrderSection(executionContext);
    }, 300);

    var formContext = executionContext.getFormContext();

    // department change
    var deptAttr = formContext.getAttribute("duc_department");
    if (deptAttr) {
        deptAttr.addOnChange(function () {
            setTimeout(function () {
                hideSubOrderTypeField(executionContext);
            }, 200);
        });
    }


    var peAttr = formContext.getAttribute("duc_processextension");
    if (peAttr) {
        peAttr.addOnChange(function () {
            toggleSendToTawasolButton(executionContext);
        });
    }

    // substatus change
    var subStatusAttr = formContext.getAttribute("msdyn_substatus");
    if (subStatusAttr) {
        subStatusAttr.addOnChange(function () {
            //           toggleReopenWorkOrderSection(executionContext);
        });
    }
}

async function toggleSendToTawasolButton(executionContext) {
    console.log("=== toggleSendToTawasolButton START ===");
    //debugger;
    var formContext = executionContext.getFormContext();

    // ---- Get Tab & Section ----
    var tab = formContext.ui.tabs.get("Inspection_action_tab");
    var section = tab ? tab.sections.get("Send_To_Tawasol_Section") : null;

    // ---- Get Button Control (optional) ----
    var btnCtrl = formContext.getControl("duc_sendtotawasolbutton");

    // default hidden
    try { if (btnCtrl) btnCtrl.setVisible(false); } catch (e) { }
    try { if (section) section.setVisible(false); } catch (e) { }
    // (don't hide the tab unless you want it hidden completely)
    // try { if (tab) tab.setVisible(false); } catch (e) {}

    // ---- (1) Check Work Order system status contains Violation ----
    var sysAttr = formContext.getAttribute("msdyn_systemstatus");
    var sysText = sysAttr ? sysAttr.getText() : "";
    var isViolation = (sysText && sysText.toLowerCase().indexOf("violation") > -1);
    if (!isViolation) return;

    // ---- (2) Check OU english name = Monitoring and Inspection - Section ----
    var deptAttr = formContext.getAttribute("duc_department");
    var deptVal = deptAttr ? deptAttr.getValue() : null;
    if (!deptVal || !deptVal[0] || !deptVal[0].id) return;

    var ouId = deptVal[0].id.replace(/[{}]/g, "");
    var ou = await Xrm.WebApi.retrieveRecord("msdyn_organizationalunit", ouId, "?$select=duc_englishname");
    if (ou.duc_englishname !== "Monitoring and Inspection - Section") return;

    // ---- (3) Check Process Extension status/substatus is in allowed list ----
    var peAttr = formContext.getAttribute("duc_processextension");
    var peVal = peAttr ? peAttr.getValue() : null;
    if (!peVal || !peVal[0] || !peVal[0].id) return;

    var peId = peVal[0].id.replace(/[{}]/g, "");

    var pe = await Xrm.WebApi.retrieveRecord(
        "duc_processextension",
        peId,
        "?$select=_duc_status_value,_duc_substatus_value"
    );

    var statusText = pe["_duc_status_value@OData.Community.Display.V1.FormattedValue"] || "";
    var subStatusText = pe["_duc_substatus_value@OData.Community.Display.V1.FormattedValue"] || "";

    // normalize (optional but recommended)
    statusText = statusText.trim();
    subStatusText = subStatusText.trim();

    // make everything lowercase to avoid case/spaces issues
    var statusKey = (statusText || "").trim().toLowerCase();
    var subStatusKey = (subStatusText || "").trim().toLowerCase();

    var allowed = new Set([
        "pending customer",
        "open settlement in progress",
        "setlement - pending section head",
        "settelment in progress",
        "pending settlement",
        "closed settlement",
        "rejected",
        "pending section head",
        "settlement - pending section hed",
        "pending department manager",
        "payment completed"
    ]);

    var isAllowed = allowed.has(statusKey) || allowed.has(subStatusKey);
    if (!isAllowed) return;


    // SHOW
    try { if (tab) tab.setVisible(true); } catch (e) { }
    try { if (section) section.setVisible(true); } catch (e) { }
    try { if (btnCtrl) btnCtrl.setVisible(true); } catch (e) { }
}

function toggleReopenWorkOrderSection(executionContext) {
    var formContext = executionContext.getFormContext();

    var tabName = "General_Tab";
    var sectionName = "Reopen_WorkOrder_Section";

    var section = formContext.ui.tabs.get(tabName)?.sections.get(sectionName);
    if (!section) return;

    var subStatusAttr = formContext.getAttribute("msdyn_substatus");
    var isClosedArchived = false;
    var deptAttr = formContext.getAttribute("duc_department");
    var deptVal = deptAttr ? deptAttr.getValue() : null;

    var ouId = deptVal[0].id.replace(/[{}]/g, "");

    if (subStatusAttr && subStatusAttr.getValue()) {
        var lookup = subStatusAttr.getValue()[0];
        if (lookup && lookup.name && lookup.name.toLowerCase() === "closed - archived") {
            isClosedArchived = true;
        }
    }


    if (isClosedArchived == true) {
        Xrm.WebApi.retrieveRecord("msdyn_organizationalunit", ouId, "?$select=duc_englishname")
            .then(function (result) {
                isClosedArchived = (result.duc_englishname !== "Chemical and Hazardous Waste - Section");
            })
            .catch(function (error) {
                console.log("Error retrieving Organizational Unit: " + error.message);
            });

    }
    // show section only when status = Closed - Archived
    section.setVisible(isClosedArchived);
}

function hideSubOrderTypeField(executionContext) {
    var formContext = executionContext.getFormContext();

    var deptAttr = formContext.getAttribute("duc_department");
    var deptVal = deptAttr ? deptAttr.getValue() : null;

    var subTypeCtrl =
        formContext.getControl("header_duc_workordersubtype") ||
        formContext.getControl("duc_workordersubtype"); // fallback if it’s not actually header

    if (!subTypeCtrl) {
        console.log("SubType control not found (header/body). Check control name on the form.");
        return;
    }

    setControlVisibleSafe(subTypeCtrl, false);

    if (!deptVal || !deptVal[0] || !deptVal[0].id) return;

    var ouId = deptVal[0].id.replace(/[{}]/g, "");

    Xrm.WebApi.retrieveRecord("msdyn_organizationalunit", ouId, "?$select=duc_englishname")
        .then(function (result) {
            var show = (result.duc_englishname === "Wildlife - Animal Section" || result.duc_englishname === "Wildlife - Plant Life Section" || result.duc_englishname === "Wildlife - Natural Resources Section");
            setControlVisibleSafe(subTypeCtrl, show);
        })
        .catch(function (error) {
            console.log("Error retrieving Organizational Unit: " + error.message);
            setControlVisibleSafe(subTypeCtrl, false);
        });
}

function setControlVisibleSafe(ctrl, visible, attemptsLeft) {
    attemptsLeft = (attemptsLeft === undefined) ? 10 : attemptsLeft;

    try {
        ctrl.setVisible(visible);
    } catch (e) {
        if (attemptsLeft <= 0) return;
        setTimeout(function () {
            setControlVisibleSafe(ctrl, visible, attemptsLeft - 1);
        }, 200);
    }
}

function resizeSection(executionContext) {

    const attempt = () => {
        const sectionName = "MainPCFSection";
        const secondSectionName = "Details"; // logical name
        const newWidth = "194.6%";
        const newMargin = "36%";
        // Get first section
        const elMain = document.querySelector('[aria-label="MainPCFSection"]');
        if (elMain) {
            elMain.style.width = newWidth;
            elMain.style.maxWidth = newWidth;
            elMain.style.flex = `0 0 ${newWidth}`;
            console.log("Main section resized");
        }

        // Get second section
        const elSecond = document.querySelector('[aria-label="Details"]');
        if (elSecond) {
            elSecond.style.marginTop = newMargin;
            console.log("Second section margin updated");
        }

        // Retry if either section is not yet in DOM
        if (!elMain || !elSecond) {
            setTimeout(attempt, 500);
        }
    };

    attempt();
}

function searchAndPopulateServiceAccount(executionContext) {
    var formContext = executionContext.getFormContext();
    formContext.getControl("duc_searchserviceaccount").clearNotification("NoAccountFound");
    // 1. Get the search text
    var searchText = formContext.getAttribute("duc_searchserviceaccount").getValue();

    // If search field is empty, clear the Service Account and return
    if (searchText == null || searchText === "") {
        formContext.getAttribute("duc_subaccount").setValue(null);
        return;
    }

    // 2. PRIORITY 1: Search directly on the Account Entity
    var accountOptions = "?$select=_duc_account_value&$filter=duc_name eq '" + searchText + "'";

    Xrm.WebApi.retrieveMultipleRecords("duc_accountpermit", accountOptions).then(
        function success(result) {
            if (result.entities.length > 0) {
                // MATCH FOUND IN ACCOUNT
                var accountRecord = result.entities[0];
                var duc_accountpermitid = accountRecord["duc_accountpermitid"]; // Guid
                var duc_account = accountRecord["_duc_account_value"]; // Lookup
                var duc_account_formatted = accountRecord["_duc_account_value@OData.Community.Display.V1.FormattedValue"];
                var duc_account_lookuplogicalname = accountRecord["_duc_account_value@Microsoft.Dynamics.CRM.lookuplogicalname"];
                setServiceAccountLookup(formContext, duc_account, duc_account_formatted);
            } else {
                // NO MATCH IN ACCOUNT -> TRIGGER PRIORITY 2
                searchPermitEntity(formContext, searchText);
            }
        },
        function (error) {
            console.error("Error searching Account: " + error.message);
        }
    );
}

function searchPermitEntity(formContext, searchText) {
    // 3. PRIORITY 2: Search on the Permit Entity
    // UPDATED: Using 'duc_serviceaccount' as the lookup field name

    var permitOptions = "?$select=_duc_serviceaccount_value&$filter=duc_number eq '" + searchText + "'";

    Xrm.WebApi.retrieveMultipleRecords("duc_permit", permitOptions).then(
        function success(result) {
            if (result.entities.length > 0) {
                // MATCH FOUND IN PERMIT
                var permitRecord = result.entities[0];

                // Get the Account ID from the Permit record (using the updated schema name)
                var accountId = permitRecord["_duc_serviceaccount_value"];
                var accountName = permitRecord["_duc_serviceaccount_value@OData.Community.Display.V1.FormattedValue"];

                if (accountId) {
                    setServiceAccountLookup(formContext, accountId, accountName);
                }
            } else {
                formContext.getControl("duc_searchserviceaccount").setNotification("This value is invalid", "NoAccountFound");
            }
        },
        function (error) {
            console.error("Error searching Permit: " + error.message);
        }
    );
}

/**
 * Show/hide fields and rename them based on user's department
 * @param {object} executionContext - Form execution context
 */
/**
 * Show/hide fields and rename them based on user's department (partial, case-insensitive match)
 * @param {object} executionContext - Form execution context
 */
/**
 * Customize fields based on department with per-department configuration
 * @param {object} executionContext - Form execution context
 */
/**
 * Customize fields based on department with per-department config and AR/EN labels
 * @param {object} executionContext - Form execution context
 */
function customizeFieldsByDepartment(executionContext, deptName) {
    var formContext = executionContext.getFormContext();

    // 1️⃣ Configuration: per field, per department, with labels in AR/EN
    var fieldConfig = {
        "duc_authority": {
            "التقييم والتصاريح": { visible: true, enabled: true, label: { en: "Project", ar: "المشروع" } },
        },
        "duc_anonymouscustomer": {
            "المحميات": { visible: false }
        },
        "new_campaign:": {
            "المحميات": { visible: true, enabled: true, label: { en: "Patrol", ar: "الدورية" } },
        }
    };

    // 2️⃣ Get current user's department


    var deptLower = deptName.toLowerCase(); // for case-insensitive match

    // Determine user's language
    var userLang = Xrm.Utility.getGlobalContext().userSettings.languageId;
    var isArabic = (userLang === 1025); // 1025 = Arabic

    // 3️⃣ Loop through fields and apply customization
    for (var field in fieldConfig) {
        if (!fieldConfig.hasOwnProperty(field)) continue;

        var control = formContext.getControl(field);
        if (!control) continue;

        var config = fieldConfig[field];
        var matchedDept = Object.keys(config).find(function (dept) {
            return deptLower.includes(dept.toLowerCase());
        });

        if (matchedDept) {
            var deptSettings = config[matchedDept];

            // Apply visibility
            if (deptSettings.visible !== undefined) {
                control.setVisible(deptSettings.visible);
            }

            // Apply enabled/disabled
            if (deptSettings.enabled !== undefined) {
                control.setDisabled(!deptSettings.enabled);
            }

            // Apply label based on language
            if (deptSettings.label) {
                var label = isArabic ? deptSettings.label.ar : deptSettings.label.en;
                if (label) control.setLabel(label);
            }
        } else {
            // Default: hide and disable if no matching department
            control.setVisible(false);
            control.setDisabled(true);
        }
    }

}

function setServiceAccountLookup(formContext, id, name) {
    // Helper function to set the lookup value
    var lookupValue = [];
    lookupValue[0] = {};
    lookupValue[0].id = id;
    lookupValue[0].name = name;
    lookupValue[0].entityType = "account"; // The target entity is always Account

    formContext.getAttribute("duc_subaccount").setValue(lookupValue);
    formContext.getAttribute("msdyn_serviceaccount").setValue(lookupValue);
}

function setDepartmentFromUserOU(executionContext) {
    var formContext = executionContext.getFormContext();
    var departmentField = formContext.getAttribute("duc_department");

    if (formContext.ui.getFormType() !== 1 || !departmentField || departmentField.getValue() !== null) {
        return;
    }

    var userId = Xrm.Utility.getGlobalContext().userSettings.userId.replace(/[{}]/g, "");

    Xrm.WebApi.retrieveRecord(
        "systemuser",
        userId,
        "?$select=_duc_organizationalunitid_value"
    ).then(function (result) {

        var orgUnitId = result._duc_organizationalunitid_value;
        var orgUnitName =
            result["_duc_organizationalunitid_value@OData.Community.Display.V1.FormattedValue"];

        if (!orgUnitId) return;

        departmentField.setValue([{
            id: orgUnitId,
            name: orgUnitName,
            entityType: "msdyn_organizationalunit"
        }]);
        customizeFieldsByDepartment(executionContext, orgUnitName);

    }, function (error) {
        console.error(error.message);
    });
}

function filtercampaign() {
    var filterc = "<filter type='and'><condition attribute='new_inspectioncampaignid' operator='in'>";
    var filter = "";
    var loggedinid = Xrm.Utility.getGlobalContext().userSettings.userId;
    var fetcxml = "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>" +
        "<entity name='new_inspectioncampaign'>" +
        "<attribute name='new_inspectioncampaignid' />" +
        " <attribute name='new_name' />" +
        "<attribute name='createdon' />" +
        "<order attribute='new_name' descending='false' />" +
        "  <link-entity name='msdyn_organizationalunit' from='msdyn_organizationalunitid' to='duc_organizationalunitid' link-type='inner' alias='ac'>" +
        " <link-entity name='systemuser' from='duc_organizationalunitid' to='msdyn_organizationalunitid' link-type='inner' alias='ad'>" +
        "  <filter type='and'>" +
        "   <condition attribute='systemuserid' operator='eq' uitype='systemuser' value='" + loggedinid + "' />" +
        " </filter>" +
        "</link-entity>" +
        "</link-entity>" +
        "</entity>" +
        "</fetch>";
    var globalcontext = Xrm.Utility.getGlobalContext();
    var clientUrl = globalcontext.getClientUrl();
    clientUrl = clientUrl + "/api/data/v9.2/new_inspectioncampaigns?fetchXml=" + fetcxml;
    var $ = parent.$;
    $.ajax({
        url: clientUrl,
        type: "GET",
        contentType: "application/json;charset=utf-8",
        datatype: "json",
        async: false,
        beforeSend: function (Xmlhttprequest) { Xmlhttprequest.setRequestHeader("Accept", "application/json"); },
        success: function (result, textStatus, xhr) {

            if (result.value.length > 0) {
                for (var i = 0; i < result.value.length; i++) {
                    filterc += "<value>{" + result.value[i].new_inspectioncampaignid + "}</value>";
                }
            }
            else {
                filterc += "<value>{00000000-0000-0000-0000-000000000000}</value>";
            }
            filterc += "</condition></filter>";
        },
        error: function (error) {
            console.log("Error retrieving data from campaign");
            filterc += "<value>{00000000-0000-0000-0000-000000000000}</value></condition></filter>";
        }
    });
    formContextCallback.getControl("new_campaign").addCustomFilter(filterc);
}

// JavaScript source code
function filterincidenttype() {
    filteri = "<filter type='and'><condition attribute='msdyn_incidenttypeid' operator='in'>";

    var loggedinid = Xrm.Utility.getGlobalContext().userSettings.userId;

    //Get M:M
    var linkQuery = "<link-entity name='duc_msdyn_incidenttype_msdyn_organizational' from='msdyn_incidenttypeid' to='msdyn_incidenttypeid' visible='false' intersect='true'>" +
        "<link-entity name='msdyn_organizationalunit' from='msdyn_organizationalunitid' to='msdyn_organizationalunitid' alias='ae'>" +
        "<link-entity name='systemuser' from='duc_organizationalunitid' to='msdyn_organizationalunitid' link-type='inner' alias='af'>" +
        "<filter type='and'>" +
        "<condition attribute='systemuserid' operator='eq' uitype='systemuser' value='" + loggedinid + "' />" +
        "</filter>" +
        "</link-entity>" +
        "</link-entity>" +
        "</link-entity>";
    filteri += filterincidenttypeHelper(linkQuery);

    //Get N:1
    linkQuery = "  <link-entity name='msdyn_organizationalunit' from='msdyn_organizationalunitid' to='duc_organizationalunitid' link-type='inner' alias='ac'>" +
        " <link-entity name='systemuser' from='duc_organizationalunitid' to='msdyn_organizationalunitid' link-type='inner' alias='ad'>" +
        "  <filter type='and'>" +
        "   <condition attribute='systemuserid' operator='eq' uitype='systemuser' value='" + loggedinid + "' />" +
        " </filter>" +
        "</link-entity>" +
        "</link-entity>";
    filteri += filterincidenttypeHelper(linkQuery);
    //filteri += filterincidenttypeHelper("duc_organizationunitsid", "msdyn_incidenttypeid");
    //filteri += filterincidenttypeHelper("msdyn_organizationalunitid", "duc_organizationalunitid");

    filteri += "</condition></filter>";
    formContextCallback.getControl("msdyn_primaryincidenttype").addCustomFilter(filteri);
}

function filterincidenttypeHelper(linkQuery) {
    var filter = "";
    var fetcxml = "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>" +
        "<entity name='msdyn_incidenttype'>" +
        "<attribute name='msdyn_incidenttypeid' />" +


        "<order attribute='msdyn_name' descending='false' />" +
        linkQuery +
        "</entity>" +
        "</fetch>";

    var globalcontext = Xrm.Utility.getGlobalContext();
    var clientUrl = globalcontext.getClientUrl();
    clientUrl = clientUrl + "/api/data/v9.2/msdyn_incidenttypes?fetchXml=" + fetcxml;
    var $ = parent.$;
    $.ajax({
        url: clientUrl,
        type: "GET",
        contentType: "application/json;charset=utf-8",
        datatype: "json",
        async: false,
        beforeSend: function (Xmlhttprequest) { Xmlhttprequest.setRequestHeader("Accept", "application/json"); },
        success: function (result, textStatus, xhr) {

            if (result.value.length > 0) {
                for (var i = 0; i < result.value.length; i++) {
                    filter += "<value>{" + result.value[i].msdyn_incidenttypeid + "}</value>";
                }
            }
            else {
                filter += "<value>{00000000-0000-0000-0000-000000000000}</value>";
            }
            return filter;
        },
        error: function (error) {
            console.log("Error retrieving data from campaign");
            filter += "<value>{00000000-0000-0000-0000-000000000000}</value>";
            return filter;
        }
    });
    return filter;
}

var formContextCallback = null;
function setPreSearch(executionContext) {
    formContextCallback = executionContext.getFormContext();
    formContextCallback.getControl("new_campaign").addPreSearch(filtercampaign);
    formContextCallback.getControl("msdyn_primaryincidenttype").addPreSearch(filterincidenttype);
}

function onchangeincidentype(executionContext) {
    formContextCallback = executionContext.getFormContext();
    if (formContextCallback.getAttribute("msdyn_primaryincidenttype").getValue() != null)
        formContextCallback.getControl("msdyn_workordertype").setDisabled(true);
    else
        formContextCallback.getControl("msdyn_workordertype").setDisabled(false);

}

function setclienttype(e) {
    var f = e.getFormContext();
    if (f.ui.getFormType() == 1) {
        var ct = Xrm.Utility.getGlobalContext().client.getClient();
        if (ct == "Mobile") {
            f.getAttribute("duc_clientdevice").setValue(1);
        }
    }
}

function setDepartmentFromIncidentType(executionContext) {
    var formContext = executionContext.getFormContext();
    var incidentTypeLookup = formContext.getAttribute("msdyn_primaryincidenttype").getValue();

    // 1. Check if we are on a Create form (Form Type 1)
    if (formContext.ui.getFormType() !== 1) {
        return; // Only run on creation
    }

    if (incidentTypeLookup != null && incidentTypeLookup[0] != null) {
        var incidentTypeId = incidentTypeLookup[0].id.replace("{", "").replace("}", "");
        var departmentField = formContext.getAttribute("duc_department");

        // 2. Clear department field if needed and check if incident type has an ID
        if (departmentField == null) return;
        departmentField.setValue(null); // Clear previous value

        // 3. Fetch the duc_organizationalunitid lookup value from the Incident Type (msdyn_incidenttype)
        Xrm.WebApi.online.retrieveMultipleRecords("msdyn_incidenttype",
            "?$select=_duc_organizationalunitid_value" +
            "&$expand=duc_organizationalunitid($select=msdyn_organizationalunitid,msdyn_name)" +
            "&$filter=msdyn_incidenttypeid eq " + incidentTypeId)
            .then(
                function success(results) {
                    if (results != null && results.entities.length > 0) {
                        var incidentTypeRecord = results.entities[0];
                        var orgUnitLookup = incidentTypeRecord["duc_organizationalunitid"];

                        if (orgUnitLookup != null) {
                            // 4. Construct the lookup object for the duc_department field
                            var departmentValue = [];
                            departmentValue[0] = {
                                id: orgUnitLookup.msdyn_organizationalunitid,
                                name: orgUnitLookup.msdyn_name,
                                entityType: "msdyn_organizationalunit" // Assuming this is the target entity type
                            };

                            // 5. Set the Work Order's duc_department field
                            departmentField.setValue(departmentValue);
                        }
                    }
                },
                function (error) {
                    console.error("Error retrieving Organizational Unit: " + error.message);
                }
            );
    } else {
        // If incident type is cleared, clear the department field
        formContext.getAttribute("duc_department").setValue(null);
    }
}

// =======================================
// AUTO REFRESH ON BACKEND STATUS CHANGE
// =======================================

var autoRefreshInterval = null;

function startStatusMonitor(executionContext) {
    const formContext = executionContext.getFormContext();

    // Avoid creating multiple intervals
    if (autoRefreshInterval !== null) {
        return;
    }

    autoRefreshInterval = setInterval(async function () {
        try {
            const workOrderId = formContext.data.entity.getId().replace(/[{}]/g, "");
            if (!workOrderId) return;

            // 1) Get FRONTEND value
            const frontStatus = formContext.getAttribute("msdyn_systemstatus")?.getValue();

            // 2) Query the BACKEND value
            const backend = await Xrm.WebApi.retrieveRecord(
                "msdyn_workorder",
                workOrderId,
                "?$select=msdyn_systemstatus"
            );

            const backendStatus = backend.msdyn_systemstatus;

            // 3) Compare → if different, refresh
            if (backendStatus !== frontStatus) {
                console.log("Status changed in backend. Refreshing form...");

                if (formContext.data.entity.getEntityName() === "msdyn_workorder") {
                    await Xrm.Navigation.openForm({
                        entityName: "msdyn_workorder",
                        entityId: formContext.data.entity.getId()
                    });
                }
            }

        } catch (err) {
            console.error("Status monitor failed:", err);
        }
    }, 5000); // 5 seconds
}

// OPTIONAL: call this on form OnLoad
function onLoad2(executionContext) {
    startStatusMonitor(executionContext);
}

function controlVisibilityOnMobile(executionContext) {
    var formContext = executionContext.getFormContext();
    var client = Xrm.Utility.getGlobalContext().client;
    var clientType = client.getClient();        // "Web", "Mobile", "Outlook"
    var formFactor = client.getFormFactor();    // 1=Desktop,2=Tablet,3=Phone

    var isMobile = (formFactor === 2 || formFactor === 3 || clientType === "Mobile");

    // Tabs to hide on mobile
    var tabNames = [
        "ActionLogs",
        "tab_10"
    ];

    for (var i = 0; i < tabNames.length; i++) {
        var tab = formContext.ui.tabs.get(tabNames[i]);
        if (tab) {
            tab.setVisible(!isMobile);  // hide on mobile, show on web
        }
    }

    // Hide Inspection_action_tab only if duc_primaryinspectionaction doesn't have data and isMobile is true
    var inspectionActionTab = formContext.ui.tabs.get("Inspection_action_tab");
    if (inspectionActionTab) {
        var primaryInspectionAction = formContext.getAttribute("duc_primaryinspectionaction");
        var hasData = primaryInspectionAction && primaryInspectionAction.getValue();
        var shouldHide = isMobile && !hasData;
        inspectionActionTab.setVisible(!shouldHide);
    }

    // Section bookingcard_section under General_Tab
    var generalTab = formContext.ui.tabs.get("General_Tab");
    if (generalTab) {
        var sections = generalTab.sections;
        var bookingSection = sections.get("bookingcard_section");
        if (bookingSection) {
            bookingSection.setVisible(!isMobile);  // hide on mobile, show on web
        }
    }
}

function hideFieldOnWeb(executionContext) {
    var formContext = executionContext.getFormContext();

    var client = formContext.context.client;
    var clientType = client.getClient();

    console.log("Client Type: ", clientType);

    // Array of field names to hide
    var fieldNames = ["duc_mapsredirect", "duc_accountcr"];

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

function WO_SetCoordinatesFromAddress(executionContext) {
    var formContext = executionContext.getFormContext();

    // Get current Work Order coordinates
    var latitudeAttr = formContext.getAttribute("msdyn_latitude");
    var longitudeAttr = formContext.getAttribute("msdyn_longitude");

    // If both coordinates are already set, don't overwrite them
    if (
        latitudeAttr &&
        longitudeAttr &&
        latitudeAttr.getValue() != null &&
        longitudeAttr.getValue() != null
    ) {
        console.log("Work Order coordinates already exist. Skipping update.");
        return;
    }

    // Get the Address lookup
    var addressAttr = formContext.getAttribute("duc_address");
    if (!addressAttr) {
        console.log("duc_address field not found on form.");
        return;
    }

    var addressVal = addressAttr.getValue();

    if (!addressVal || !addressVal[0] || !addressVal[0].id) {
        console.log("No address selected in duc_address lookup.");
        return;
    }

    var addressId = addressVal[0].id.replace(/[{}]/g, "");

    // Retrieve Address coordinates
    Xrm.WebApi.retrieveRecord(
        "duc_addressinformation",
        addressId,
        "?$select=duc_latitude,duc_longitude"
    ).then(
        function (addressRecord) {

            if (
                addressRecord.duc_latitude == null ||
                addressRecord.duc_longitude == null
            ) {
                console.log("Address record does not contain valid coordinates.");
                return;
            }

            // Only populate fields that are currently empty
            if (latitudeAttr && latitudeAttr.getValue() == null) {
                latitudeAttr.setValue(addressRecord.duc_latitude);
                console.log("Set msdyn_latitude to: " + addressRecord.duc_latitude);
            }

            if (longitudeAttr && longitudeAttr.getValue() == null) {
                longitudeAttr.setValue(addressRecord.duc_longitude);
                console.log("Set msdyn_longitude to: " + addressRecord.duc_longitude);
            }
        },
        function (error) {
            console.error("Error retrieving address record: " + (error.message || "Unknown error"));
        }
    );
}

function WO_ManageAddressSections(executionContext) {
    var formContext = executionContext.getFormContext();

    // Get the address tab
    var addressTab = formContext.ui.tabs.get("address");
    if (!addressTab) {
        console.log("Address tab not found.");
        return;
    }

    // Get the sections
    var accountAddressesSection = addressTab.sections.get("account_addresses");
    var workOrderAddressesSection = addressTab.sections.get("workorder_adresses");

    if (!accountAddressesSection || !workOrderAddressesSection) {
        console.log("One or both address sections not found.");
        return;
    }

    // Get the work order ID
    var workOrderId = formContext.data.entity.getId().replace(/[{}]/g, "");

    // Get the sub-account lookup field
    var subAccountAttr = formContext.getAttribute("duc_subaccount");
    var subAccountVal = subAccountAttr ? subAccountAttr.getValue() : null;

    // Function to hide both sections initially
    function hideAllSections() {
        accountAddressesSection.setVisible(false);
        workOrderAddressesSection.setVisible(false);
    }

    // Function to show account addresses section
    function showAccountAddresses() {
        accountAddressesSection.setVisible(true);
        workOrderAddressesSection.setVisible(false);
        console.log("Showing account addresses section.");
    }

    // Function to show work order addresses section
    function showWorkOrderAddresses() {
        accountAddressesSection.setVisible(false);
        workOrderAddressesSection.setVisible(true);
        console.log("Showing work order addresses section.");
    }

    // Function to check for addresses related to work order
    function checkWorkOrderAddresses() {
        var woOptions = "?$select=duc_addressinformationid";
        woOptions += "&$filter=_duc_msdyn_workorder_value eq " + workOrderId;
        woOptions += "&$top=1"; // We only need to know if at least one exists

        Xrm.WebApi.retrieveMultipleRecords("duc_addressinformation", woOptions).then(
            function (result) {
                if (result.entities && result.entities.length > 0) {
                    // Found addresses related to work order
                    showWorkOrderAddresses();
                } else {
                    // No addresses found anywhere
                    hideAllSections();
                    console.log("No addresses found for account or work order. Hiding both sections.");
                }
            },
            function (error) {
                console.error("Error checking work order addresses: " + (error.message || "Unknown error"));
                hideAllSections();
            }
        );
    }

    // Function to check for addresses related to sub-account
    function checkAccountAddresses(accountId) {
        var accountOptions = "?$select=duc_addressinformationid";
        accountOptions += "&$filter=_duc_account_value eq " + accountId;
        accountOptions += "&$top=1"; // We only need to know if at least one exists

        Xrm.WebApi.retrieveMultipleRecords("duc_addressinformation", accountOptions).then(
            function (result) {
                if (result.entities && result.entities.length > 0) {
                    // Found addresses related to sub-account
                    showAccountAddresses();
                } else {
                    // No addresses found for sub-account, check work order
                    checkWorkOrderAddresses();
                }
            },
            function (error) {
                console.error("Error checking account addresses: " + (error.message || "Unknown error"));
                // On error, still try to check work order addresses
                checkWorkOrderAddresses();
            }
        );
    }

    // Start the check process
    if (subAccountVal && subAccountVal[0] && subAccountVal[0].id) {
        // Sub-account is selected, check for addresses related to it first
        var accountId = subAccountVal[0].id.replace(/[{}]/g, "");
        checkAccountAddresses(accountId);
    } else {
        // No sub-account selected, check work order addresses directly
        checkWorkOrderAddresses();
    }
}

function toggleResponsibleEmployeeSections(executionContext) {
    console.log("=== toggleResponsibleEmployeeSections START ===");

    var formContext = executionContext.getFormContext();

    var flagAttr = formContext.getAttribute("duc_responsibleemployeeisnotavailable");
    if (!flagAttr) {
        console.warn("Field duc_responsibleemployeeisnotavailable not found");
        return;
    }

    var flagValue = flagAttr.getValue(); // true / false
    console.log("Flag value:", flagValue);

    // true = NOT available → hide sections
    var showSections = flagValue === false;

    var tab = formContext.ui.tabs.get("ResponsibleEmployee");
    if (tab) {
        var sectionNames = [
            "{b8e326ee-5c21-4a18-ba55-e3b56966c249}_responsible_employee_tab_section_3",
            "{b8e326ee-5c21-4a18-ba55-e3b56966c249}_section_10"
        ];

        sectionNames.forEach(function (sectionName) {
            var section = tab.sections.get(sectionName);
            if (section) {
                section.setVisible(showSections);
            }
        });
    }

    /* if (flagValue === true) {
  
      var refusedAttr = formContext.getAttribute("duc_responsibleemployeerefusedtosign");
      if (refusedAttr && refusedAttr.getValue() !== false) {
          refusedAttr.setValue(false);
          console.log("duc_responsibleemployeerefusedtosign set to FALSE");
      }
  }*/



    console.log("=== toggleResponsibleEmployeeSections END ===");
}

async function handleBookingSuggestionVisiblity(executionContext) {
    console.log("=== handleBookingSuggestionVisiblity start ===");

    const formContext = executionContext.getFormContext();
    const tab = formContext.ui.tabs.get("General_Tab");
    if (!tab) return;

    const section = tab.sections.get("bookingcard_section");
    if (!section) return;

    // Default hide
    section.setVisible(false);

    // ---------------- STATUS CONDITION ----------------
    const statusAttr = formContext.getAttribute("msdyn_systemstatus");
    if (!statusAttr) return;

    const hideValues = [
        690970006, 100000007, 100000001, 100000000,
        690970003, 100000004, 100000002, 100000009,
        690970004, 100000003, 100000005, 100000008,
        690970005, 100000010, 100000011, 100000012,
        100000013, 100000014, 100000015, 100000016,
        100000017, 100000018, 100000019, 100000020,
        100000021, 100000022
    ];

    const statusValue = statusAttr.getValue();
    console.log("Booking status value:", statusValue);
    if (statusValue === null || hideValues.includes(statusValue)) {
        console.log("Status not allowed, hiding section.");
        return;
    }

    // ---------------- OWNER CONDITION ----------------
    const ownerAttr = formContext.getAttribute("ownerid");
    if (!ownerAttr || !ownerAttr.getValue()) return;

    const owner = ownerAttr.getValue()[0];
    const loggedUserId = Xrm.Utility.getGlobalContext().userSettings.userId.replace(/[{}]/g, "");
    console.log("Owner info:", owner, "Logged-in User ID:", loggedUserId);

    if (owner.entityType === "systemuser") {
        // Owner is a user
        const ownerId = owner.id.replace(/[{}]/g, "");
        if (ownerId.toLowerCase() === loggedUserId.toLowerCase()) {
            console.log("Owner is the logged-in user. Showing section.");
            section.setVisible(true);
        } else {
            console.log("Owner is a different user. Section remains hidden.");
        }
        return;
    }

    if (owner.entityType === "team") {
        // Owner is a team, check membership
        const teamId = owner.id.replace(/[{}]/g, "");
        try {
            const membersResponse = await Xrm.WebApi.retrieveMultipleRecords(
                "teammembership",
                `?$filter=teamid eq ${teamId} and systemuserid eq ${loggedUserId}&$select=systemuserid`
            );

            if (membersResponse.entities.length > 0) {
                console.log("Logged-in user is a team member. Showing section.");
                section.setVisible(true);
            } else {
                section.setVisible(false);
                console.log("Logged-in user is NOT a team member. Section remains hidden.");
            }
        } catch (e) {
            console.error("Error retrieving team membership info:", e);
        }
    }

    console.log("=== handleBookingSuggestionVisiblity end ===");
}

function handleResponsibleEmployeeTabVisiblity(executionContext) {
    const formContext = executionContext.getFormContext();
    const statusAttr = formContext.getAttribute("msdyn_systemstatus");
    if (!statusAttr) return;

    const statusValue = statusAttr.getValue();
    const tab = formContext.ui.tabs.get("ResponsibleEmployee");
    if (!tab) return;

    const hideValues = [690970000, 690970001];

    if (statusValue !== null && hideValues.includes(statusValue)) {
        console.log("Status matches hideValues. Hiding tab.");
        tab.setVisible(false);
    } else {
        console.log("Status does not match hideValues. Showing tab.");
        tab.setVisible(true);
    }
}

function toggleBookingCardSection(executionContext) {
    var formContext = executionContext.getFormContext();

    var STATUS_TO_SHOW = 690970000;

    var tabName = "General_Tab";
    var sectionName = "bookingcard_section";

    var statusAttr = formContext.getAttribute("msdyn_systemstatus");
    if (!statusAttr) return;

    var statusValue = statusAttr.getValue();

    var tab = formContext.ui.tabs.get(tabName);
    if (!tab) return;

    var section = tab.sections.get(sectionName);
    if (!section) return;

    section.setVisible(statusValue === STATUS_TO_SHOW);
}

function showTabsOnlyForSystemAdmin(executionContext) {
    var formContext = executionContext.getFormContext();

    var tabsToControl = [
        "ActionLogs",
        "InteractionTab"
    ];

    // Hide all by default
    tabsToControl.forEach(function (tabName) {
        var tab = formContext.ui.tabs.get(tabName);
        if (tab) {
            tab.setVisible(false);
        }
    });

    var roles = Xrm.Utility.getGlobalContext().userSettings.roles;
    var isAdmin = false;

    if (roles && roles.getAll) {
        var allRoles = roles.getAll();
        for (var i = 0; i < allRoles.length; i++) {
            var roleName = (allRoles[i].name || "").toLowerCase();
            if (roleName === "system administrator") {
                isAdmin = true;
                break;
            }
        }
    }

    // Show tabs only for System Admin
    if (isAdmin) {
        tabsToControl.forEach(function (tabName) {
            var tab = formContext.ui.tabs.get(tabName);
            if (tab) {
                tab.setVisible(true);
            }
        });
    }
}
function initializeResponsibleEmployeeValidations(executionContext) {
    var formContext = executionContext.getFormContext();

    // Employee ID: only numbers
    var empId = formContext.getAttribute("duc_employeeid");
    if (empId) {
        empId.addOnChange(function () {
            var value = empId.getValue();
            if (value && /\D/.test(value)) {
                empId.setValue(value.replace(/\D/g, ""));
            }
        });
    }

    // Employee Name: only letters
    var empName = formContext.getAttribute("duc_employeename");
    if (empName) {
        empName.addOnChange(function () {
            var value = empName.getValue();
            if (value && /[^a-zA-Zء-ي\s]/.test(value)) {
                empName.setValue(value.replace(/[^a-zA-Zء-ي\s]/g, ""));
            }
        });
    }

    // Nationality: only letters
    var nationality = formContext.getAttribute("duc_nationality");
    if (nationality) {
        nationality.addOnChange(function () {
            var value = nationality.getValue();
            if (value && /[^a-zA-Zء-ي\s]/.test(value)) {
                nationality.setValue(value.replace(/[^a-zA-Zء-ي\s]/g, ""));
            }
        });
    }
}



async function lockFormIfDepartmentMismatch(executionContext) {
    var formContext = executionContext.getFormContext();

    try {
        var globalContext = Xrm.Utility.getGlobalContext();
        var userId = globalContext.userSettings.userId.replace(/[{}]/g, "");

        // Get current user's duc_department
        var userResult = await Xrm.WebApi.retrieveRecord(
            "systemuser",
            userId,
            "?$select=_duc_department_value"
        );

        var userDepartmentId = userResult._duc_department_value
            ? userResult._duc_department_value.toLowerCase()
            : null;

        // If user has no department, lock the form
        if (!userDepartmentId) {
            disableAllFields(formContext);
            return;
        }

        // Get current form duc_department lookup
        var formDepartmentAttr = formContext.getAttribute("duc_department");
        if (!formDepartmentAttr) {
            console.log("Field duc_department not found on form.");
            return;
        }

        var formDepartmentValue = formDepartmentAttr.getValue();

        // If form department is empty, lock the form
        if (!formDepartmentValue || formDepartmentValue.length === 0) {
            disableAllFields(formContext);
            return;
        }

        var organizationalUnitId = formDepartmentValue[0].id.replace(/[{}]/g, "");

        // Get duc_department from selected msdyn_organizationalunit
        var organizationalUnitResult = await Xrm.WebApi.retrieveRecord(
            "msdyn_organizationalunit",
            organizationalUnitId,
            "?$select=_duc_department_value"
        );

        var organizationUnitDepartmentId = organizationalUnitResult._duc_department_value
            ? organizationalUnitResult._duc_department_value.toLowerCase()
            : null;

        // Compare user department with organizational unit's duc_department
        if (!organizationUnitDepartmentId || organizationUnitDepartmentId !== userDepartmentId) {
            disableAllFields(formContext);
        }

    } catch (e) {
        console.error("Error in lockFormIfDepartmentMismatch: " + e.message);
    }
}

function disableAllFields(formContext) {
    formContext.ui.controls.forEach(function (control) {
        try {
            if (control && typeof control.setDisabled === "function") {
                control.setDisabled(true);
            }
        } catch (e) {
            // Ignore controls that can't be disabled
        }
    });
}

function setFirstOpenedByInspector(executionContext) {
    var formContext = executionContext.getFormContext();

    var currentUserId = Xrm.Utility.getGlobalContext().userSettings.userId.replace("{", "").replace("}", "").toLowerCase();

    var assignedInspectorAttr = formContext.getAttribute("duc_assignedinspector");
    var firstOpenedAttr = formContext.getAttribute("duc_firstopenedbyinspector");

    if (!assignedInspectorAttr || !firstOpenedAttr) return;

    if (firstOpenedAttr.getValue()) return;

    var assignedInspectorLookup = assignedInspectorAttr.getValue();
    if (!assignedInspectorLookup || assignedInspectorLookup.length === 0) return;

    var assignedInspectorId = assignedInspectorLookup[0].id.replace("{", "").replace("}", "").toLowerCase();

    if (assignedInspectorId === currentUserId) {

        var now = new Date();

        // Set on main form
        firstOpenedAttr.setValue(now);
        firstOpenedAttr.setSubmitMode("always");

        // Update related duc_workorderdetails
        updateWorkOrderDetails(formContext.data.entity.getId(), {
            "duc_firstopenedbyinspector": now
        });

        console.log("duc_firstopenedbyinspector set for current user");
    }
}

function setFirstOpenedByHead(executionContext) {
    var formContext = executionContext.getFormContext();

    var currentUserId = Xrm.Utility.getGlobalContext().userSettings.userId.replace("{", "").replace("}", "").toLowerCase();
    var firstOpenedByHeadAttr = formContext.getAttribute("duc_firstopenedbyhead");
    var incidentTypeAttr = formContext.getAttribute("msdyn_primaryincidenttype");

    if (!firstOpenedByHeadAttr || !incidentTypeAttr) return;
    if (firstOpenedByHeadAttr.getValue()) return;

    var incidentTypeLookup = incidentTypeAttr.getValue();
    if (!incidentTypeLookup || incidentTypeLookup.length === 0) return;

    var incidentTypeId = incidentTypeLookup[0].id.replace("{", "").replace("}", "");

    Xrm.WebApi.retrieveRecord("msdyn_incidenttype", incidentTypeId, "?$select=_duc_headsofdepartment_value").then(
        function (result) {

            var headId = result["_duc_headsofdepartment_value"];
            if (!headId) return;

            headId = headId.toLowerCase();

            if (headId === currentUserId) {

                var now = new Date();

                // Set on main form
                firstOpenedByHeadAttr.setValue(now);
                firstOpenedByHeadAttr.setSubmitMode("always");

                // Update related duc_workorderdetails
                updateWorkOrderDetails(formContext.data.entity.getId(), {
                    "duc_firstopenedbyhead": now
                });

                console.log("duc_firstopenedbyhead set");
            }
        },
        function (error) {
            console.error(error.message);
        }
    );
}
function updateWorkOrderDetails(workOrderId, dataToUpdate) {

    if (!workOrderId) return;

    workOrderId = workOrderId.replace("{", "").replace("}", "");

    var fetchXml = `
        <fetch>
            <entity name="duc_workorderdetails">
                <attribute name="duc_workorderdetailsid" />
                <filter>
                    <condition attribute="duc_relatedworkorder" operator="eq" value="${workOrderId}" />
                </filter>
            </entity>
        </fetch>`;

    Xrm.WebApi.retrieveMultipleRecords("duc_workorderdetails", "?fetchXml=" + encodeURIComponent(fetchXml)).then(
        function (result) {

            if (result.entities.length === 0) {
                console.log("No related duc_workorderdetails found");
                return;
            }

            result.entities.forEach(function (record) {

                Xrm.WebApi.updateRecord("duc_workorderdetails", record.duc_workorderdetailsid, dataToUpdate).then(
                    function () {
                        console.log("Updated duc_workorderdetails:", record.duc_workorderdetailsid);
                    },
                    function (error) {
                        console.error("Update failed:", error.message);
                    }
                );
            });
        },
        function (error) {
            console.error("Retrieve failed:", error.message);
        }
    );
}


function handleProcessExtensionVisibility(executionContext) {
    console.log("handleProcessExtensionVisibility triggered");
    const formContext = executionContext.getFormContext();

    const attr = formContext.getAttribute("duc_processextension");
    const value = attr ? attr.getValue() : null;

    const tab = formContext.ui.tabs.get("tab_13");

    if (!tab) {
        console.log("Tab not found: tab_13");
        return;
    }

    const section = tab.sections.get("tab_13_section_1");

    if (!section) {
        console.log("Section not found: tab_13_section_1");
        return;
    }

    // hide if null
    section.setVisible(value ? true : false);
}


async function hideSectionIfJsonFormEmpty(executionContext) {
    try {
        var formContext = executionContext.getFormContext();

        // Get tab
        var tab = formContext.ui.tabs.get("tab_13");
        if (!tab) {
            console.log("tab_13 not found");
            return;
        }

        // Get section
        var section = tab.sections.get("tab_13_section_1");
        if (!section) {
            console.log("tab_13_section_1 not found");
            return;
        }

        // Get Process Extension lookup
        var peAttr = formContext.getAttribute("duc_processextension");
        var peVal = peAttr ? peAttr.getValue() : null;

        // No process extension → hide section
        if (!peVal || !peVal[0] || !peVal[0].id) {
            section.setVisible(false);
            return;
        }

        var peId = peVal[0].id.replace(/[{}]/g, "");

        // Retrieve duc_jsonform
        var pe = await Xrm.WebApi.retrieveRecord(
            "duc_processextension",
            peId,
            "?$select=duc_jsonform"
        );

        console.log("Process Extension:", pe);

        var jsonFormValue = pe["duc_jsonform"];

        console.log("duc_jsonform value:", jsonFormValue);

        // Check empty correctly
        var hasValue =
            jsonFormValue !== null &&
            jsonFormValue !== undefined &&
            jsonFormValue.toString().trim() !== "";

        // Show only when value exists
        section.setVisible(hasValue);

    } catch (e) {
        console.error("Error in hideSectionIfJsonFormEmpty:", e);
    }
}

function validateQID(executionContext) {
    validateField(executionContext, "duc_employeeid", "NumbersOnlyRegex", "NumbersOnlyMessageEN", "NumbersOnlyMessageAR");
}

function validateEmployeeName(executionContext) {
    validateField(executionContext, "duc_employeename", "AlphaRegex", "AlphaMessageEN", "AlphaMessageAR");
}

var configCache = {};

async function getConfigValue(configName) {

    if (configCache[configName]) {
        return configCache[configName];
    }

    var result = await Xrm.WebApi.retrieveMultipleRecords(
        "duc_processautomationconfigurations",
        `?$select=duc_value&$filter=duc_key eq '${configName}'`
    );

    if (result.entities.length === 0) {
        return null;
    }

    var value = result.entities[0].duc_value;

    configCache[configName] = value;

    return value;
}

async function validateField(executionContext, fieldName, regexConfigName, messageEnConfigName, messageArConfigName) {

    var formContext = executionContext.getFormContext();

    var attribute = formContext.getAttribute(fieldName);

    if (!attribute) return;

    var value = attribute.getValue();

    if (!value) return;

    try {

        var regexPattern = await getConfigValue(regexConfigName);

        if (!regexPattern) return;

        var regex = new RegExp(regexPattern);

        var userLanguage = Xrm.Utility.getGlobalContext().userSettings.languageId;

        var message = userLanguage === 1025 ? await getConfigValue(messageArConfigName) : await getConfigValue(messageEnConfigName);

        var control = formContext.getControl(fieldName);

        if (!regex.test(value)) {
            attribute.setValue(null);

            control.clearNotification();

            var alertStrings = { text: message };

            Xrm.Navigation.openAlertDialog(alertStrings).then(
                function (success) {
                },
                function (error) {
                    console.error("Alert dialog error: " + error.message);
                }
            );
        } else {
            control.clearNotification();
        }
    } catch (e) {

        console.error("Validation failed: " + e.message);
    }
}


async function toggleCaptainFieldVisibility(executionContext) {

    try {

        const formContext = executionContext.getFormContext();

        console.log("=== toggleCaptainFieldVisibility START ===");

        // CONTROLS
        const captainCtrl = formContext.getControl("duc_captain");
        const unitMeccCtrl = formContext.getControl("duc_unitsmecc");
        const targetedAreaCtrl = formContext.getControl("duc_targetedarea");
        const subAreaCtrl = formContext.getControl("duc_subarea");

        if (!captainCtrl && !targetedAreaCtrl && !subAreaCtrl && !unitMeccCtrl) {
            console.log(" Controls not found");
            return;
        }

        console.log(" Controls found");

        // Hide by default
        if (captainCtrl) {
            captainCtrl.setVisible(false);
        }

        if (unitMeccCtrl) {
            unitMeccCtrl.setVisible(false);
        }

        if (targetedAreaCtrl) {
            targetedAreaCtrl.setVisible(false);
        }

        if (subAreaCtrl) {
            subAreaCtrl.setVisible(false);
        }

        console.log("Fields hidden by default");

        // DEPARTMENT FIELD
        const deptAttr = formContext.getAttribute("duc_department");

        if (!deptAttr) {
            console.log(" duc_department attribute not found");
            return;
        }

        const deptVal = deptAttr.getValue();

        console.log("Department Value:", deptVal);

        if (!deptVal || deptVal.length === 0) {
            console.log("No department selected");
            return;
        }

        const departmentId = deptVal[0].id.replace(/[{}]/g, "");

        console.log("Department ID:", departmentId);

        // Retrieve Organizational Unit
        const result = await Xrm.WebApi.retrieveRecord(
            "msdyn_organizationalunit",
            departmentId,
            "?$select=duc_englishname"
        );

        console.log("Retrieve Result:", result);

        const englishName = result.duc_englishname || "";

        console.log("English Name:", englishName);

        // Check if contains "marine"
        const showFields =
            englishName.toLowerCase().includes("marine");

        console.log("Show Fields:", showFields);

        // SHOW / HIDE FIELDS
        if (captainCtrl) {
            captainCtrl.setVisible(showFields);
        }

        if (unitMeccCtrl) {
            unitMeccCtrl.setVisible(showFields);
        }

        if (targetedAreaCtrl) {
            targetedAreaCtrl.setVisible(showFields);
        }

        if (subAreaCtrl) {
            subAreaCtrl.setVisible(showFields);
        }

        console.log("Fields visibility updated");

    } catch (error) {

        console.log(" ERROR in toggleCaptainFieldVisibility");
        console.log(error);
        console.log(error.message);

    }
}
function isOffline() {
    try {
        var ctx = Xrm.Utility.getGlobalContext().client;
        // Actually check if an entity is available offline — this returns true
        // ONLY when a Mobile Offline profile is enabled and includes the entity.
        // Just checking whether the API *exists* is not enough because the
        // Xrm.WebApi.offline namespace is present on all mobile apps.
        if (Xrm.WebApi.offline &&
            typeof Xrm.WebApi.offline.isAvailableOffline === "function" &&
            Xrm.WebApi.offline.isAvailableOffline("msdyn_workorder")) return true;
        if (Xrm.WebApi.isAvailableOffline &&
            Xrm.WebApi.isAvailableOffline("msdyn_workorder")) return true;
        if (ctx.isOffline()) return true;
        if (ctx.getClientState() === "Offline") return true;
    } catch (e) { }
    return false;

}