// navigate to the relevant services task 
function oFST(){try{var f=Xrm.Page,w=f.data.entity.getId();if(!w)return Xrm.Navigation.openAlertDialog({text:"Work Order ID not found."});w=w.replace(/[{}]/g,"");Xrm.WebApi.retrieveMultipleRecords("msdyn_workorderservicetask","?$select=msdyn_workorderservicetaskid&$filter=_msdyn_workorder_value eq "+w+"&$orderby=createdon asc&$top=1").then(r=>{if(!r.entities.length)return Xrm.Navigation.openAlertDialog({text:"No related service tasks found."});Xrm.Navigation.openForm({entityName:"msdyn_workorderservicetask",entityId:r.entities[0].msdyn_workorderservicetaskid});});}catch(e){}}



async function uLA() {
    try {
      
	  		const woId = Xrm.Page.data.entity.getId().replace(/[{}]/g, "");


        var woRecord = await Xrm.WebApi.retrieveRecord(
            "msdyn_workorder",
            woId,
            "?$select=_duc_processextension_value"
        );

        if (!woRecord._duc_processextension_value) return;

        var peId = woRecord._duc_processextension_value;

        var peRecord = await Xrm.WebApi.retrieveRecord(
            "duc_processextension",
            peId,
            "?$select=_duc_processdefinition_value,_duc_currentstage_value"
        );

        if (!peRecord._duc_processdefinition_value) return;

        var processDefinitionId = peRecord._duc_processdefinition_value;
        var currentStageId = peRecord._duc_currentstage_value;

        var stageActions = await Xrm.WebApi.retrieveMultipleRecords(
            "duc_stageaction",
            `?$select=duc_stageactionid,_duc_defaultstatus_value
              &$filter=duc_canbetriggeredbytarget eq true
              and _duc_process_value eq ${processDefinitionId}
              and _duc_relatedstage_value eq ${currentStageId}
              &$top=10`
        );

        var actionIdToSet = null;

        for (let item of stageActions.entities) {
            let statusId = item._duc_defaultstatus_value;

            if (statusId) {
                let status = await Xrm.WebApi.retrieveRecord(
                    "duc_processstatus",
                    statusId,
                    "?$select=duc_value"
                );

                if (status.duc_value == 690970002) {
                    actionIdToSet = item.duc_stageactionid;
                    break;
                }
            }
        }

        if (!actionIdToSet) return;

        await Xrm.WebApi.updateRecord("duc_processextension", peId, {
            "duc_LastActionTaken_duc_ProcessExtension@odata.bind":
                `/duc_stageactions(${actionIdToSet})`
        });

    } catch (e) {
        console.warn("uLA error:", e);
    }
}

async function sBIP() {
    try {
          
		const woId = Xrm.Page.data.entity.getId().replace(/[{}]/g, "");
        var bookingResult = await Xrm.WebApi.retrieveMultipleRecords(
            "bookableresourcebooking",
            `?$select=bookableresourcebookingid
              &$filter=_msdyn_workorder_value eq ${woId}
              &$orderby=createdon asc
              &$top=1`
        );

        if (!bookingResult.entities.length) return;

        var bookingId = bookingResult.entities[0].bookableresourcebookingid;

        var statusResult = await Xrm.WebApi.retrieveMultipleRecords(
            "bookingstatus",
            `?$select=bookingstatusid&$filter=name eq 'In Progress'`
        );

        if (!statusResult.entities.length) {
            return Xrm.Navigation.openAlertDialog({ text: "Booking Status 'In Progress' not found." });
        }

        var statusId = statusResult.entities[0].bookingstatusid;

        await Xrm.WebApi.updateRecord("bookableresourcebooking", bookingId, {
            "BookingStatus@odata.bind": `/bookingstatuses(${statusId})`
        });

    } catch (e) {
        Xrm.Navigation.openAlertDialog({ text: e.message });
    }
}

function IsFromMobile() {
	if (Xrm.Page.context.client.getClient() == "Mobile"
&& (Xrm.Page.context.client.getFormFactor() == 2 || Xrm.Page.context.client.getFormFactor() == 3)
	) {
		return true;
	}
	return false;
}


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

async function hasWorkOrderServiceTask(workOrderId) {
    if (!workOrderId) {
        return false;
    }

    const cleanWorkOrderId = workOrderId.replace(/[{}]/g, "");

    const result = await Xrm.WebApi.retrieveMultipleRecords(
        "msdyn_workorderservicetask",
        `?$select=msdyn_workorderservicetaskid&$filter=_msdyn_workorder_value eq ${cleanWorkOrderId}&$top=1`
    );

    return result.entities && result.entities.length > 0;
}


const workOrderId = Xrm.Page.data.entity.getId();

const hasTasks = await hasWorkOrderServiceTask(workOrderId);

if(!hasTasks)
{
  Xrm.Navigation.openAlertDialog({
        text: "No Service Tasks found for this Work Order, please try again in few moments"
    });
    Xrm.Utility.closeProgressIndicator();

return;
}
  if(IsFromMobile())
	{

	createDailyInspection();	
	}
oFST();
  var upadtebookableResourcesBooking = await sBIP();
if(upadtebookableResourcesBooking )
{  
 uLA();
}