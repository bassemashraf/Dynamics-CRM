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

            AddCampaignLookupFilter(formContext);
        },
        function (error) {
            console.error("Error retrieving user's organization unit: " + JSON.stringify(error));
        }
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
        "?$select=_duc_department_value"
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

    if (!anonymousAttr || !subAccountAttr) {
        return;
    }

    var anonymousValue = anonymousAttr.getValue();

    if (anonymousValue != true) {
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

            // Retrieve the organizational unit and get business unit ID
            Xrm.WebApi.retrieveRecord(
                "msdyn_organizationalunit",
                orgUnitId,
                "?$select=_duc_businessunit_value"
            ).then(
                function successOrgUnit(orgUnit) {
                    if (!orgUnit._duc_businessunit_value) {
                        console.log("No business unit found for the organizational unit.");
                        return;
                    }

                    var businessUnitId = orgUnit._duc_businessunit_value;

                    // Retrieve Account where duc_isunknown = true AND owningbusinessunit matches the business unit
                    Xrm.WebApi.retrieveMultipleRecords(
                        "account",
                        "?$select=accountid,name&$filter=duc_isunknown eq true and _owningbusinessunit_value eq " + businessUnitId
                    ).then(
                        function successAccount(result) {
                            if (result.entities.length === 0) {
                                console.log("No unknown account found for the business unit.");
                                return;
                            }

                            var account = result.entities[0];

                            var lookupValue = [{
                                id: account.accountid,
                                name: account.name,
                                entityType: "account"
                            }];

                            subAccountAttr.setValue(lookupValue);
                            subAccountAttr.setSubmitMode("always");

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
                        function errorAccount(error) {
                            console.error("Error retrieving unknown account:", error.message);
                        }
                    );
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
