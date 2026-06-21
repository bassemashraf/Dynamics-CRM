// ============================================================
// Edited Functions - Offline Query Fixes (Inline Branching)
// Date: 2026-06-16
// 
// Each function uses if(isOffline()) {...} else {...} inline.
// isOffline() is defined in hideheader.js (loaded on same form).
// No function names were changed.
//
// ISSUES FIXED:
// 1. Navigation property in $filter (duc_NewAccountType/duc_accounttype)
//    → Not supported offline → replaced with two-step retrieveRecord
// 2. Lookup fields in $filter (_fieldname_value)
//    → Offline requires schema name → switched with isOffline() check
// 3. teammembership intersect entity
//    → May not be available offline → skipped when offline
// ============================================================


// ============================================================
// FUNCTION 1: toggleVehicleOwnerButton
// Location: WO_Onload.js ~line 368
// Issue: duc_NewAccountType/duc_accounttype navigation property in $filter
// ============================================================

async function toggleVehicleOwnerButton(executionContext) {
    var formContext = executionContext.getFormContext();

    var btnCtrl = formContext.getControl("duc_assignvehicleowner");
    if (!btnCtrl) {
        console.log("Vehicle owner button control not found");
        return;
    }

    btnCtrl.setVisible(false);

    var serviceAccountAttr = formContext.getAttribute("msdyn_serviceaccount");
    if (!serviceAccountAttr) return;

    var serviceAccount = serviceAccountAttr.getValue();
    if (!serviceAccount || !serviceAccount[0] || !serviceAccount[0].id) return;

    var accountId = serviceAccount[0].id.replace(/[{}]/g, "");

    try {
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


// ============================================================
// FUNCTION 2: toggleSelectEstablishmentButton
// Location: WO_Onload.js ~line 427
// Issue: duc_NewAccountType/duc_accounttype navigation property in $filter
// ============================================================

async function toggleSelectEstablishmentButton(executionContext) {
    var formContext = executionContext.getFormContext();

    var isIndividual = false;
    var btnCtrl = formContext.getControl("duc_selectestablishmentaccount");
    if (!btnCtrl) {
        console.log("Select Establishment button control not found");
        return;
    }

    btnCtrl.setVisible(false);

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
                if (accountTypeRecord.duc_accounttype === 2) {
                    isIndividual = true;
                }
            }
        } else {
            // ONLINE: use navigation property filter
            var result = await Xrm.WebApi.retrieveMultipleRecords(
                "account",
                `?$select=accountid&$filter=accountid eq ${accountId} and _parentaccountid_value eq null and (duc_NewAccountType/duc_accounttype eq 2)&$top=1`
            );
            if (result.entities.length > 0) {
                isIndividual = true;
            }
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


// ============================================================
// FUNCTION 3: toggleVehicleOwnerTab
// Location: WO_Onload.js ~line 480
// Issue: duc_NewAccountType/duc_accounttype navigation property in $filter
// ============================================================

async function toggleVehicleOwnerTab(executionContext) {
    var formContext = executionContext.getFormContext();

    var tab = formContext.ui.tabs.get("tab_17_Vehicle_Owner");
    if (!tab) {
        console.log("Vehicle Owner tab not found");
        return;
    }

    tab.setVisible(false);

    var serviceAccountAttr = formContext.getAttribute("msdyn_serviceaccount");
    if (!serviceAccountAttr) return;

    var serviceAccount = serviceAccountAttr.getValue();
    if (!serviceAccount || !serviceAccount[0] || !serviceAccount[0].id) return;

    var accountId = serviceAccount[0].id.replace(/[{}]/g, "");

    try {
        var showTab = false;

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
                showTab = (at === 1 || at === 3 || at === 6);
            }
        } else {
            // ONLINE: use navigation property filter
            var result = await Xrm.WebApi.retrieveMultipleRecords(
                "account",
                `?$select=accountid&$filter=accountid eq ${accountId} and _parentaccountid_value eq null and (duc_NewAccountType/duc_accounttype eq 1 or duc_NewAccountType/duc_accounttype eq 3 or duc_NewAccountType/duc_accounttype eq 6)&$top=1`
            );
            showTab = result.entities.length > 0;
        }

        if (showTab) {
            tab.setVisible(true);
        }

    } catch (error) {
        console.error("Error checking account type for vehicle owner tab:", error);
    }
}


// ============================================================
// FUNCTION 4: WO_ManageAddressSections (inner functions)
// Location: WO_Onload.js ~line 1321
// Issue: _duc_msdyn_workorder_value and _duc_account_value in $filter
//        → Offline requires schema name (duc_msdyn_workorder, duc_account)
// ============================================================

// checkWorkOrderAddresses (inner function):
//   var woFilterField = isOffline() ? "duc_msdyn_workorder" : "_duc_msdyn_workorder_value";
//   woOptions += "&$filter=" + woFilterField + " eq " + workOrderId;

// checkAccountAddresses (inner function):
//   var accFilterField = isOffline() ? "duc_account" : "_duc_account_value";
//   accountOptions += "&$filter=" + accFilterField + " eq " + accountId;


// ============================================================
// FUNCTION 5: handleBookingSuggestionVisiblity
// Location: WO_Onload.js ~line 1472
// Issue: teammembership intersect entity not available offline
//        → Skip team membership check when offline
// ============================================================

// Inside the team owner branch:
//   if (isOffline()) {
//       console.log("Offline mode: skipping team membership check. Section remains hidden.");
//       return;
//   }
