var currentOrgUnitId = null;

function onLoad(executionContext) {
    LookupFilterInit(executionContext);

    SetDepartment(executionContext);

    setIncidentTypeRequirement(executionContext);
}

function LookupFilterInit(executionContext) {
    var formContext = executionContext.getFormContext();
    var userId = Xrm.Utility.getGlobalContext().userSettings.userId.replace("{", "").replace("}", "");

    Xrm.WebApi.retrieveRecord("systemuser", userId, "?$select=_duc_organizationalunitid_value").then(
        function success(result) {
            currentOrgUnitId = result._duc_organizationalunitid_value;

            Xrm.WebApi.retrieveRecord("msdyn_organizationalunit", currentOrgUnitId, "?$select=duc_englishname").then(
                function success2(orgUnitResult) {
                    var OUEnglishName = orgUnitResult["duc_englishname"];

                    if (OUEnglishName == "Inspection Section – Natural Reserves") {

                        var filter = "duc_campaignstatus eq 2 and duc_campaigninternaltype eq 100000004 and duc_campaigntype eq 100000004 and _ownerid_value eq " + userId + " and _duc_organizationalunitid_value eq " + currentOrgUnitId;

                        Xrm.WebApi.retrieveMultipleRecords("new_inspectioncampaign", "?$select=new_inspectioncampaignid,new_name&$filter=" + encodeURIComponent(filter)).then(
                            function success3(results) {
                                if (results.entities && results.entities.length > 0) {
                                    var result = results.entities[0];
                                    formContext.getAttribute("new_campaign").setValue([{
                                        id: result["new_inspectioncampaignid"],
                                        name: result["new_name"],
                                        entityType: "new_inspectioncampaign"
                                    }]);
                                }
                            },
                            function (error) { console.error("Campaign lookup error: " + error.message); }
                        );
                    }
                    else {
                        AddCampaignLookupFilter(formContext);
                    }
                },
                function (error) { console.error("Org unit lookup error: " + error.message); }
            );
        },
        function (error) { console.error("User org unit error: " + error.message); }
    );
}

function AddCampaignLookupFilter(formContext) {

    var campaignStatusValue = 2;

    var campaignTypeValue = 100000000;

    var filterXml =
        "<filter type='and'>" +
        "<condition attribute='duc_campaignstatus' operator='eq' value='" + campaignStatusValue + "' />" +
        "<condition attribute='duc_campaigntype' operator='eq' value='" + campaignTypeValue + "' />" +
        "<condition attribute='duc_organizationalunitid' operator='eq' value='" + currentOrgUnitId + "' />" +
        "</filter>";

    var lookupControl = formContext.getControl("new_campaign");

    if (lookupControl) {
        lookupControl.addPreSearch(function () {
            lookupControl.addCustomFilter(filterXml, "new_inspectioncampaign");
        });
    }
}

function SetParentCampaign(executionContext) {
    var formContext = executionContext.getFormContext();

    var campaignAttr = formContext.getAttribute("new_campaign");

    var parentCampaignAttr = formContext.getAttribute("duc_parentcampaign");

    var incidentTypeAttr = formContext.getAttribute("msdyn_primaryincidenttype");

    var workOrderTypeAttr = formContext.getAttribute("msdyn_workordertype");

    var campaignValue = campaignAttr.getValue();

    if (!campaignValue) {
        parentCampaignAttr.setValue(null);

        incidentTypeAttr.setValue(null);

        formContext.getControl("msdyn_primaryincidenttype").setDisabled(false);

        workOrderTypeAttr.setValue(null);

        return;
    }

    parentCampaignAttr.setValue([{
        id: campaignValue[0].id,
        name: campaignValue[0].name,
        entityType: campaignValue[0].entityType
    }]);

    Xrm.WebApi.retrieveRecord(
        "new_inspectioncampaign",
        campaignValue[0].id,
        "?$select=_duc_incidenttype_value"
    ).then(function success(result) {

        var id = result["_duc_incidenttype_value"];

        if (!id) return;

        var name = result["_duc_incidenttype_value@OData.Community.Display.V1.FormattedValue"];

        var entity = result["_duc_incidenttype_value@Microsoft.Dynamics.CRM.lookuplogicalname"];

        incidentTypeAttr.setValue([
            { id: id, name: name, entityType: entity }
        ]);

        formContext.getControl("msdyn_primaryincidenttype").setDisabled(true);

        incidentTypeAttr.fireOnChange();

    }, function (error) {
        console.log(error);
    });
}

function SetDepartment(executionContext) {
    var formContext = executionContext.getFormContext();

    var userId = Xrm.Utility.getGlobalContext().userSettings.userId.replace("{", "").replace("}", "");

    Xrm.WebApi.retrieveRecord(
        "systemuser",
        userId,
        "?$select=_duc_organizationalunitid_value"
    ).then(
        function success(result) {
            var OUId = result._duc_organizationalunitid_value;
            if (!OUId) return;

            var name = result["_duc_organizationalunitid_value@OData.Community.Display.V1.FormattedValue"];

            var entity = result["_duc_organizationalunitid_value@Microsoft.Dynamics.CRM.lookuplogicalname"];

            formContext.getAttribute("duc_department").setValue([
                { id: OUId, name: name, entityType: entity }
            ]);
        },
        function (error) {
            console.error(error);
        }
    );
}

function SetWorkOrderType(executionContext) {
    var formContext = executionContext.getFormContext();

    var incidentTypeAttr = formContext.getAttribute("msdyn_primaryincidenttype");

    var workOrderTypeAttr = formContext.getAttribute("msdyn_workordertype");

    var incidentVal = incidentTypeAttr.getValue();

    if (!incidentVal) {
        workOrderTypeAttr.setValue(null);

        return;
    }

    Xrm.WebApi.retrieveRecord(
        "msdyn_incidenttype",
        incidentVal[0].id,
        "?$select=_msdyn_defaultworkordertype_value"
    ).then(function success(result) {

        var id = result["_msdyn_defaultworkordertype_value"];
        if (!id) return;

        var name = result["_msdyn_defaultworkordertype_value@OData.Community.Display.V1.FormattedValue"];

        var entity = result["_msdyn_defaultworkordertype_value@Microsoft.Dynamics.CRM.lookuplogicalname"];

        workOrderTypeAttr.setValue([
            { id: id, name: name, entityType: entity }
        ]);

    }, function (error) {
        console.log(error);
    });
}

function onAnonymousCustomerChange(executionContext) {
    debugger
    var formContext = executionContext.getFormContext();

    var anonymousAttr = formContext.getAttribute("duc_anonymouscustomer");
    var subAccountAttr = formContext.getAttribute("duc_subaccount");
    var serviceAccountAttr = formContext.getAttribute("msdyn_serviceaccount");

    if (!anonymousAttr || !subAccountAttr || !serviceAccountAttr) {
        return;
    }

    var anonymousValue = anonymousAttr.getValue();

    // If false, clear both fields
    if (anonymousValue != true) {
        subAccountAttr.setValue(null);
        serviceAccountAttr.setValue(null);
        return;
    }

    // Get current user's ID
    var currentUserId = Xrm.Utility.getGlobalContext().userSettings.userId.replace(/[{}]/g, "");

    // Retrieve the current user's organizational unit ID
    Xrm.WebApi.retrieveRecord(
        "systemuser",
        currentUserId,
        "?$select=_duc_organizationalunitid_value"
    ).then(
        function successUser(user) {
            if (!user._duc_organizationalunitid_value) {
                console.log("No organizational unit found for the user.");
                return;
            }

            var orgUnitId = user._duc_organizationalunitid_value;

            // Retrieve the organizational unit and get the unknown account lookup
            Xrm.WebApi.retrieveRecord(
                "msdyn_organizationalunit",
                orgUnitId,
                "?$select=_duc_unknownaccount_value"
            ).then(
                function successOrgUnit(orgUnit) {
                    if (!orgUnit._duc_unknownaccount_value) {
                        console.log("No unknown account found in the organizational unit.");
                        return;
                    }

                    // Set the unknown account lookup value
                    var lookupValue = [{
                        id: orgUnit._duc_unknownaccount_value,
                        name: orgUnit["_duc_unknownaccount_value@OData.Community.Display.V1.FormattedValue"] || "",
                        entityType: "account"
                    }];

                    // Set both duc_subaccount and msdyn_serviceaccount
                    subAccountAttr.setValue(lookupValue);
                    subAccountAttr.setSubmitMode("always");

                    serviceAccountAttr.setValue(lookupValue);
                    serviceAccountAttr.setSubmitMode("always");

                    // Hide both fields
                    var subAccountControl = formContext.getControl("duc_subaccount");
                    var searchServiceAccountControl = formContext.getControl("duc_searchserviceaccount");

                    if (subAccountControl != null) {
                        subAccountControl.setVisible(false);
                    }

                    if (searchServiceAccountControl != null) {
                        searchServiceAccountControl.setVisible(false);
                    }
                },
                function errorOrgUnit(error) {
                    console.error("Error retrieving organizational unit:", error.message);
                }
            );
        },
        function errorUser(error) {
            console.error("Error retrieving user information:", error.message);
        }
    );
}

function setIncidentTypeRequirement(executionContext) {
    var formContext = executionContext.getFormContext();

    var incidentTypeAttr = formContext.getAttribute("msdyn_primaryincidenttype");

    if (incidentTypeAttr) {
        incidentTypeAttr.setRequiredLevel("required");
    }
}

function fireIncidentTypeOnChange(executionContext) {
    try {
        var formContext = executionContext.getFormContext();
        
        var primaryIncidentTypeAttr = formContext.getAttribute("msdyn_primaryincidenttype");
        
        if (primaryIncidentTypeAttr == null) {
            console.log("msdyn_primaryincidenttype field not found on form");
            return;
        }

        // Check if the field has a value (set by default from PCF)
        var primaryIncidentTypeValue = primaryIncidentTypeAttr.getValue();
        
        if (primaryIncidentTypeValue != null && primaryIncidentTypeValue.length > 0) {
            console.log("Primary Incident Type found with default value:", primaryIncidentTypeValue[0].id);
            
            // Trigger the onChange event to execute any business logic
            // This will fire any OnChange handlers registered for this field
            primaryIncidentTypeAttr.fireOnChange();
            
            console.log("Primary Incident Type onChange event triggered successfully");
        } else {
            console.log("No default value found for Primary Incident Type");
        }


    } catch (error) {
        console.error("Error in onLoadQuickCreate:", error.message);
    }
}