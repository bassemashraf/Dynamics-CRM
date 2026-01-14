async function filterCampaignTypeOnLoad(executionContext) {
    try {
        const formContext = executionContext.getFormContext();
        
        const orgUnitLookup = formContext.getAttribute("duc_organizationalunitid").getValue();
        
        if (!orgUnitLookup || orgUnitLookup.length === 0) {
            console.log("No organizational unit selected");
            return;
        }
        
        const orgUnitId = orgUnitLookup[0].id.replace(/[{}]/g, ""); // Remove braces from GUID
        
        const orgUnitResult = await Xrm.WebApi.retrieveRecord(
            "msdyn_organizationalunit",
            orgUnitId,
            "?$select=duc_englishname"
        );
        
        const englishName = orgUnitResult.duc_englishname;
        console.log("English Name:", englishName);
        
        const campaignTypeControl = formContext.getControl("duc_campaigntype");
        
        if (!campaignTypeControl) {
            console.error("Campaign type control not found");
            return;
        }
        
        // Clear any existing filters
        campaignTypeControl.clearOptions();
        
        // Define which options to show based on English name
        let allowedValues = [];
        
        if (englishName === "Wildlife - Plant Life Section" || 
            englishName === "Wildlife - Natural Resources Section") {

            allowedValues = [100000000, 100000001, 100000002, 100000004];
        } 
        else if (englishName === "Wildlife - Animal Section") {
            allowedValues = [100000000, 100000001, 100000002];
        }
        else {
            console.log("Other section - no filtering applied");
            return;
        }
        
        // Get all available options from the attribute
        const campaignTypeAttribute = formContext.getAttribute("duc_campaigntype");
        const allOptions = campaignTypeAttribute.getOptions();
        
        // Add only the allowed options back
        allOptions.forEach(option => {
            if (allowedValues.includes(option.value)) {
                campaignTypeControl.addOption(option);
            }
        });
        
        // Check if current value is still valid
        const currentValue = campaignTypeAttribute.getValue();
        if (currentValue !== null && !allowedValues.includes(currentValue)) {
            // Clear the value if it's not in the allowed list
            campaignTypeAttribute.setValue(null);
            console.log("Current value cleared as it's not allowed for this organizational unit");
        }
        
    } catch (error) {
        console.error("Error in filterCampaignTypeOnLoad:", error);
        // Xrm.Navigation.openAlertDialog({
        //     text: "An error occurred while filtering campaign types: " + error.message
        // });
    }
}

async function onOrganizationalUnitChange(executionContext) {
    await filterCampaignTypeOnLoad(executionContext);
}
