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