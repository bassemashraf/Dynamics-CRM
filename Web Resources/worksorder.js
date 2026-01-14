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
}function WO_SetCoordinatesFromAddress(executionContext) {
    var formContext = executionContext.getFormContext();
    
    // Get the duc_address lookup field
    var addressAttr = formContext.getAttribute("duc_address");
    if (!addressAttr) {
        console.log("duc_address field not found on form.");
        return;
    }
    
    var addressVal = addressAttr.getValue();
    
    // Check if the lookup has a value
    if (!addressVal || !addressVal[0] || !addressVal[0].id) {
        console.log("No address selected in duc_address lookup.");
        return;
    }
    
    var addressId = addressVal[0].id.replace(/[{}]/g, "");
    
    // Retrieve the address information record to get coordinates
    Xrm.WebApi.retrieveRecord("duc_addressinformation", addressId, "?$select=duc_latitude,duc_longitude").then(
        function (addressRecord) {
            // Check if coordinates exist
            if (addressRecord.duc_latitude != null && addressRecord.duc_longitude != null) {
                // Set msdyn_latitude
                var latitudeAttr = formContext.getAttribute("msdyn_latitude");
                if (latitudeAttr) {
                    latitudeAttr.setValue(addressRecord.duc_latitude);
                    console.log("Set msdyn_latitude to: " + addressRecord.duc_latitude);
                }
                
                // Set msdyn_longitude
                var longitudeAttr = formContext.getAttribute("msdyn_longitude");
                if (longitudeAttr) {
                    longitudeAttr.setValue(addressRecord.duc_longitude);
                    console.log("Set msdyn_longitude to: " + addressRecord.duc_longitude);
                }
            } else {
                console.log("Address record does not have valid coordinates.");
            }
        },
        function (error) {
            console.error("Error retrieving address record: " + (error.message || "Unknown error"));
        }
    );
}