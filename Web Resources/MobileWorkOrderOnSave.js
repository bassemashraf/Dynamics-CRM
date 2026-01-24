function DetectMObileCreation(executionContext) {
    try {

        const formContext = executionContext.getFormContext();
        const context = Xrm.Utility.getGlobalContext().client;
        const isMobileApp = context.getClient() === "Mobile";

        if (!isMobileApp) {
            return;
        }


        formContext.getAttribute("duc_createdfrommobile").setValue(true);

        formContext.getAttribute("duc_createdfrommobile").setSubmitMode("always");

    } catch (e) {
        console.error("Error in DetectMObileCreation:", e);
    }
}

async function quickCreateonSave(executionContext) {

    var formContext = executionContext.getFormContext();
    var serviceAccount = formContext.getAttribute("duc_subaccount").getValue();
    var PrimaryIncidentType = formContext.getAttribute("msdyn_primaryincidenttype").getValue();

    if (serviceAccount != null && PrimaryIncidentType != null) {
        formContext.data.save().then(
            async function (data) {
                var entityFormOptions = {};
                entityFormOptions["entityName"] = formContext.data.entity.getEntityName();
                entityFormOptions["entityId"] = data.savedEntityReference.id;
                entityFormOptions["openInNewWindow"] = true;
                await createAutoBookingOnWorkOrderCreate(executionContext, entityFormOptions["entityId"]);
                Xrm.Navigation.openForm(entityFormOptions).then(
                    function (success) {
                        console.log(success);
                    },
                    function (error) {
                        console.log(error);
                    });

            },
            function () {
                console.log("error callback");
            }
        );
    }
}

async function createAutoBookingOnWorkOrderCreate(executionContext, workorderId) {
    let progressIndicator = null;
    try {
        const formContext = executionContext.getFormContext();


        if (formContext.ui.getFormType() == 1) return;

        const booked = formContext.getAttribute("duc_createdfrommobile")?.getValue();
        if (booked !== true) {
            console.log("Work order not created from mobile. Exiting booking logic.");
            return;
        }

        if (formContext.getAttribute("msdyn_bookingcreated")?.getValue() === true) {
            console.log("Booking already created for this work order");
            return;
        }

        const userLanguageCode = Xrm.Utility.getGlobalContext().userSettings.languageId;

        // Translation messages
        const translations = {
            1025: "جاري إنشاء الحجز، يرجى الانتظار...",  // Arabic
            1033: "Creating booking, please wait...",      // English
            // Add more language codes as needed
        };

        // Get translated message or default to English
        const loadingMessage = translations[userLanguageCode] || translations[1033];

        // Show loading indicator with translated message
        Xrm.Utility.showProgressIndicator(loadingMessage);
        progressIndicator = true;

        const createdBy = Xrm.Utility.getGlobalContext().userSettings.userId.replace(/[{}]/g, "");

        // Check if user is inspector
        const userResults = await Xrm.WebApi.retrieveMultipleRecords(
            "systemuser",
            `?$select=duc_isinspector&$filter=(systemuserid eq ${createdBy})`
        );

        if (userResults.entities.length === 0 || !userResults.entities[0].duc_isinspector) {
            console.log("User is not an inspector. Booking not created.");
            return;
        }

        // Get resource for the user
        const resourceRes = await Xrm.WebApi.retrieveMultipleRecords(
            "bookableresource",
            `?$select=bookableresourceid&$filter=_userid_value eq ${createdBy} and resourcetype eq 3`
        );

        if (resourceRes.entities.length === 0) {
            console.error("No bookable resource found for user");
            return;
        }

        const resourceId = resourceRes.entities[0].bookableresourceid;

        // Get "Scheduled" booking status
        const statusQuery = "?$select=bookingstatusid,name&$filter=name eq 'Scheduled'";
        const statusResults = await Xrm.WebApi.retrieveMultipleRecords("bookingstatus", statusQuery);

        if (statusResults.entities.length === 0) {
            console.error("Booking status 'Scheduled' not found.");
            return;
        }

        const bookingStatusId = statusResults.entities[0].bookingstatusid;

        // Get the work order ID from the current form (after save)

        if (!workorderId) {
            console.error("Work order ID not available");
            return;
        }

        const now = new Date();
        const end = new Date(Date.now() + 60 * 60 * 1000);

        const booking = {
            "ownerid@odata.bind": `/systemusers(${createdBy})`,
            "starttime": now.toISOString(),
            "endtime": end.toISOString(),
            "duration": 1,
            "msdyn_workorder@odata.bind": `/msdyn_workorders(${workorderId})`,
            "Resource@odata.bind": `/bookableresources(${resourceId})`,
            "BookingStatus@odata.bind": `/bookingstatuses(${bookingStatusId})`
        };

        const result = await Xrm.WebApi.createRecord("bookableresourcebooking", booking);
        console.log("Booking created successfully. Booking ID:", result.id);

        // Mark that booking was created to prevent duplicate attempts
        // const updateData = {
        //     "duc_createdfrommobile": false
        // };

        // await Xrm.WebApi.updateRecord("msdyn_workorder", workorderId, updateData);

    } catch (e) {
        alert("Error in booking logic. See console for details. " + e.message);
        console.error("Error in JS booking logic:", e);
    }
    finally {
        // Always close the loading indicator
        if (progressIndicator) {
            Xrm.Utility.closeProgressIndicator();
            // Xrm.Navigation.navigateTo(
            //     {
            //         pageType: "entityrecord",
            //         entityName: "msdyn_workorder",
            //         entityId: workorderId
            //     },
            //     {
            //         target: 1
            //     }
            // );
        }
    }
}

async function onSubaccountChange(executionContext) {

    var formContext = executionContext.getFormContext();

    // Get the selected subaccount
    var subaccountLookup = formContext.getAttribute("duc_subaccount").getValue();

    if (subaccountLookup && subaccountLookup.length > 0) {
        var subaccountId = subaccountLookup[0].id.replace(/[{}]/g, "");
        var serviceAccount = formContext.getAttribute("msdyn_serviceaccount");

        // Retrieve the selected account to check parent account and address lookup
        await Xrm.WebApi.retrieveRecord("account", subaccountId, "?$select=parentaccountid").then(
            async function success(result) {
                var serviceAccountValue;

                // Check if parentaccountid has a value
                if (result.parentaccountid) {
                    // Set parent account on msdyn_serviceaccount
                    serviceAccountValue = [{
                        id: result.parentaccountid,
                        name: result["_parentaccountid_value@OData.Community.Display.V1.FormattedValue"],
                        entityType: "account"
                    }];
                } else {
                    // Set the subaccount itself on msdyn_serviceaccount
                    serviceAccountValue = subaccountLookup;
                }

                formContext.getAttribute("msdyn_serviceaccount").setValue(serviceAccountValue);

                // Handle address lookup if duc_address has a value
                // if (result._duc_address_value) {
                //     var addressId = result._duc_address_value;
                //     var addressName = result["_duc_address_value@OData.Community.Display.V1.FormattedValue"];

                //     // Set the address lookup on the current form
                //     formContext.getAttribute("duc_address").setValue([{
                //         id: addressId,
                //         name: addressName,
                //         entityType: "duc_addressinformation"
                //     }]);

                //     // Now retrieve the address information to get longitude and latitude
                //     await Xrm.WebApi.retrieveRecord("duc_addressinformation", addressId, "?$select=duc_longitude,duc_latitude").then(
                //         function successAddress(addressResult) {
                //             // Set longitude if available
                //             if (addressResult.duc_longitude != null) {
                //                 formContext.getAttribute("msdyn_longitude").setValue(addressResult.duc_longitude);
                //             }

                //             // Set latitude if available
                //             if (addressResult.duc_latitude != null) {
                //                 formContext.getAttribute("msdyn_latitude").setValue(addressResult.duc_latitude);
                //             }
                //         },
                //         function errorAddress(error) {
                //             console.log("Error retrieving address information: " + error.message);
                //         }
                //     );

                // } else {
                //     // Clear address fields if no address found
                //     formContext.getAttribute("duc_address").setValue(null);
                //     formContext.getAttribute("msdyn_longitude").setValue(null);
                //     formContext.getAttribute("msdyn_latitude").setValue(null);
                // }
            },
            function error(error) {
                console.log("Error retrieving account: " + error.message);
            }
        );
        if (serviceAccount != null) {
            serviceAccount.fireOnChange();
        }
    } else {
        // Clear all fields if duc_subaccount is cleared
        formContext.getAttribute("msdyn_serviceaccount").setValue(null);
        formContext.getAttribute("duc_address").setValue(null);
        formContext.getAttribute("msdyn_longitude").setValue(null);
        formContext.getAttribute("msdyn_latitude").setValue(null);
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