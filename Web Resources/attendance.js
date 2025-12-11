async function createDailyInspection() {
    try {
        Xrm.Utility.showProgressIndicator("Creating inspection record...");
 
        const workOrderId = Xrm.Page.data.entity.getId().replace(/[{}]/g, "");
       
        if (!workOrderId) {
            Xrm.Navigation.openAlertDialog({
                text: "Unable to retrieve Work Order ID. Please save the form first."
            });
            Xrm.Utility.closeProgressIndicator();
            return;
        }
 
        const createdBy = Xrm.Utility.getGlobalContext().userSettings.userId.replace(/[{}]/g, "");
       
        const resourceRes = await Xrm.WebApi.retrieveMultipleRecords(
            "bookableresource",
            `?$select=bookableresourceid&$filter=_userid_value eq ${createdBy} and resourcetype eq 3`
        );
 
        if (!resourceRes.entities || resourceRes.entities.length === 0) {
            Xrm.Navigation.openAlertDialog({
                text: "No bookable resource found for the current user."
            });
            Xrm.Utility.closeProgressIndicator();
            return;
        }
 
        const resourceId = resourceRes.entities[0].bookableresourceid;
 
        const currentDateTime = new Date();
        
        // Get today's date without time
        const todayStart = new Date(currentDateTime.getFullYear(), currentDateTime.getMonth(), currentDateTime.getDate());
        
        // Check if attendance record exists for today for this user
        const attendanceRes = await Xrm.WebApi.retrieveMultipleRecords(
            "duc_attendance",
            `?$select=duc_attendanceid&$filter=_duc_user_value eq ${createdBy} and Microsoft.Dynamics.CRM.On(PropertyName='createdon',PropertyValue='${todayStart.toISOString()}')`
        );
        
        let attendanceId;
        
        if (attendanceRes.entities && attendanceRes.entities.length > 0) {
            // Use existing attendance record
            attendanceId = attendanceRes.entities[0].duc_attendanceid;
            console.log("Using existing attendance record: " + attendanceId);
        } else {
            // Create new attendance record
            const attendanceData = {
                "duc_User@odata.bind": `/systemusers(${createdBy})`
            };
            
            const attendanceResult = await Xrm.WebApi.createRecord("duc_attendance", attendanceData);
            attendanceId = attendanceResult.id;
            console.log("Created new attendance record: " + attendanceId);
        }
 
        // Create daily inspection record with attendance lookup
        const recordData = {
            "duc_WorkOrder@odata.bind": `/msdyn_workorders(${workOrderId})`,
            "duc_BookableResource@odata.bind": `/bookableresources(${resourceId})`,
            "duc_Attendance@odata.bind": `/duc_attendances(${attendanceId})`,
            "duc_startinspectiontime": currentDateTime
        };
 
        const result = await Xrm.WebApi.createRecord("duc_dailyinspectorinspections", recordData);
 
        Xrm.Page.data.refresh(false);
 
        console.log("Daily inspection record created with ID: " + result.id);
        
        Xrm.Utility.closeProgressIndicator();
 
    } catch (error) {
        Xrm.Utility.closeProgressIndicator();
 
        Xrm.Navigation.openAlertDialog({
            text: "Error creating daily inspection record: " + error.message
        });
 
        console.error("Error creating daily inspection:", error);
    }
}