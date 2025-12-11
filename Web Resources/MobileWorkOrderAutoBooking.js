async function createAutoBookingOnWorkOrderCreate(executionContext) {
    let progressIndicator = null;
    let workorderId = null;
    try {
        const formContext = executionContext.getFormContext();
        debugger

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

        // Show loading indicator
        Xrm.Utility.showProgressIndicator("Creating booking, please wait...");
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
        workorderId = formContext.data.entity.getId().replace(/[{}]/g, "");

        if (!workorderId) {
            console.error("Work order ID not available");
            return;
        }

        const now = new Date();
        const end = new Date(Date.now() + 60 * 60 * 1000);

        const booking = {
            "ownerid@odata.bind": `/systemusers(${createdBy})`,
            "starttime": now,
            "endtime": end,
            "duration": 1,
            "msdyn_workorder@odata.bind": `/msdyn_workorders(${workorderId})`,
            "Resource@odata.bind": `/bookableresources(${resourceId})`,
            "BookingStatus@odata.bind": `/bookingstatuses(${bookingStatusId})`
        };

        const result = await Xrm.WebApi.createRecord("bookableresourcebooking", booking);
        console.log("Booking created successfully. Booking ID:", result.id);

        // Mark that booking was created to prevent duplicate attempts
        const updateData = {
            "duc_createdfrommobile": false
        };

        await Xrm.WebApi.updateRecord("msdyn_workorder", workorderId, updateData);

    } catch (e) {
        alert("Error in booking logic. See console for details. " + e.message);
        console.error("Error in JS booking logic:", e);
    } finally {
        // Always close the loading indicator
        if (progressIndicator) {
            Xrm.Utility.closeProgressIndicator();
            Xrm.Navigation.navigateTo(
                {
                    pageType: "entityrecord",
                    entityName: "msdyn_workorder",
                    entityId: workorderId
                },
                {
                    target: 1   
                }
            );
        }
    }
}