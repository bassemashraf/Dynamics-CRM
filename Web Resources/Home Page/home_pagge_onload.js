function hideFormHeader(executionContext) {
    var formContext = executionContext.getFormContext();

    try {
        // Hide header items (works on web, partial support on mobile)
        formContext.ui.headerSection.setBodyVisible(false);
        formContext.ui.headerSection.setCommandBarVisible(false);
        formContext.ui.headerSection.setTabNavigatorVisible(false);

        // Force ribbon update
        formContext.ui.refreshRibbon();
    } catch (e) {
        console.log("Header manipulation error (may be expected on mobile): " + e.message);
    }

    // // Get global context
    // var globalContext = Xrm.Utility.getGlobalContext();

    // // Detect client type
    // var clientType = globalContext.client.getClient();
    // var isMobile = (clientType === "Mobile");
    // var isWeb = (clientType === "Web");

    // // Get environment identifier - works on both mobile and web
    // var orgUniqueName = globalContext.organizationSettings.uniqueName.toLowerCase();

    // // Alternative: Use organization URL as fallback
    // var orgUrl = globalContext.getClientUrl().toLowerCase();

    // // Determine environment
    // var isInspectionImport = orgUniqueName.includes("unq2e5feb323b91f011a700000d3a3a0") || orgUrl.includes("inspectionimport");
    // var isMecc = orgUniqueName.includes("mecc") || orgUrl.includes("mecc");

    // // Get field controls
    // var nameField = formContext.getControl("duc_name");
    // var gtaHomeField = formContext.getControl("duc_gtahome");

    // // Apply visibility rules based on environment
    // if (isInspectionImport) {
    //     // inspectionimport: hide duc_name, show duc_gtahome
    //     if (nameField) {
    //         nameField.setVisible(false);
    //     }
    //     if (gtaHomeField) {
    //         gtaHomeField.setVisible(true);
    //     }

    //     console.log("Environment: InspectionImport | Client: " + clientType);
    // }
    // else {
    //     // mecc: show duc_name, hide duc_gtahome
    //     if (nameField) {
    //         nameField.setVisible(true);
    //     }
    //     if (gtaHomeField) {
    //         gtaHomeField.setVisible(false);
    //     }

    //     console.log("Environment: MECC | Client: " + clientType);
    // }

    // // Mobile-specific adjustments (if needed)
    // if (isMobile) {
    //     // Additional mobile-specific logic can go here
    //     // For example, adjust field layouts or add mobile-friendly messages
    // }
}

/**
 * Toggles visibility between the offline and online PCF home fields
 * based on the current connectivity state of the client.
 *
 * - Offline: shows duc_name (offline PCF), hides duc_onlinehome (online PCF)
 * - Online:  shows duc_onlinehome (online PCF), hides duc_name (offline PCF)
 *
 * @param {Object} executionContext - The form execution context passed by CRM.
 */
function isOffline() {
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

function toggleOfflineFields(executionContext) {
    var formContext = executionContext.getFormContext();

    try {
        var isCurrentlyOffline = isOffline();

        var offlineField = formContext.getControl("duc_name");       // Offline PCF component
        var onlineField  = formContext.getControl("duc_onlinehome"); // Online PCF component

        if (isCurrentlyOffline) {
            // --- OFFLINE mode ---
            if (offlineField) offlineField.setVisible(true);
            if (onlineField)  onlineField.setVisible(false);
            console.log("[HomeOnload] Offline mode detected — showing duc_name, hiding duc_onlinehome.");
        } else {
            // --- ONLINE mode ---
            if (offlineField) offlineField.setVisible(false);
            if (onlineField)  onlineField.setVisible(true);
            console.log("[HomeOnload] Online mode detected — showing duc_onlinehome, hiding duc_name.");
        }
    } catch (e) {
        console.log("[HomeOnload] toggleOfflineFields error: " + e.message);
    }
}