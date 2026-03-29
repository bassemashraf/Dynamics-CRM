function getMSG() {
    var lang = Xrm.Utility.getGlobalContext().userSettings.languageId;
    return (lang === 1025) ? {
        waitingTask: "جاري إنشاء مهام الخدمة، يرجى الانتظار...",
        taskTimeout: "لم يتم إنشاء مهام الخدمة بعد. يرجى المحاولة مرة أخرى بعد قليل.",
        taskNotFound: "لم يتم العثور على مهام الخدمة.",
        woNotFound: "لم يتم العثور على رقم أمر العمل."
    } : {
        waitingTask: "Creating service tasks, please wait...",
        taskTimeout: "Service tasks are not created yet. Please try again in a few moments.",
        taskNotFound: "No service tasks found.",
        woNotFound: "Work Order ID not found."
    };
}
async function waitForWorkOrderServiceTask(workOrderId, maxWaitMs = 30000, intervalMs = 2000) {
    try {
        const start = Date.now();
        const MSG = getMSG();
        const isOff = isOffline();
        const woField = isOff ? "msdyn_workorder" : "_msdyn_workorder_value";

        Xrm.Utility.showProgressIndicator(MSG.waitingTask);

        // Offline: tasks are either already on device or not — no point polling
        if (isOff) {
            const result = await Xrm.WebApi.retrieveMultipleRecords(
                "msdyn_workorderservicetask",
                `?$select=msdyn_workorderservicetaskid
                 &$filter=${woField} eq ${workOrderId}
                 &$orderby=createdon asc
                 &$top=1`
            );
            Xrm.Utility.closeProgressIndicator();
            return (result.entities && result.entities.length > 0)
                ? result.entities[0].msdyn_workorderservicetaskid
                : null;
        }

        // Online: poll until tasks are created by server plugin
        while ((Date.now() - start) < maxWaitMs) {
            const result = await Xrm.WebApi.retrieveMultipleRecords(
                "msdyn_workorderservicetask",
                `?$select=msdyn_workorderservicetaskid
                 &$filter=${woField} eq ${workOrderId}
                 &$orderby=createdon asc
                 &$top=1`
            );

            if (result.entities && result.entities.length > 0) {
                Xrm.Utility.closeProgressIndicator();
                return result.entities[0].msdyn_workorderservicetaskid;
            }

            await new Promise(r => setTimeout(r, intervalMs));
        }

        Xrm.Utility.closeProgressIndicator();
        return null;
    } catch (e) {
        Xrm.Utility.closeProgressIndicator();
        var errMsg = "[waitForWorkOrderServiceTask] Error while retrieving service tasks."
            + "\nWorkOrderId: " + workOrderId
            + "\nOffline: " + isOffline()
            + "\nError: " + (e.message || JSON.stringify(e));
        console.error(errMsg, e);
        Xrm.Navigation.openAlertDialog({ text: errMsg });
        return null;
    }
}




// navigate to the relevant services task 
async function oFST(woId) {
    try {
        const MSG = getMSG();
        var form = Xrm.Page;
        var accAttr = form.getAttribute("msdyn_serviceaccount");
        var accountId = null;
        if (accAttr && accAttr.getValue()) {
            accountId = accAttr.getValue()[0].id.replace(/[{}]/g, "");
        }
        console.log("accountId is:::" + accountId);
        if (!woId) {
            return Xrm.Navigation.openAlertDialog({ text: MSG.woNotFound });
        }

        woId = woId.replace(/[{}]/g, "");

        const serviceTaskId = await waitForWorkOrderServiceTask(woId);

        if (!serviceTaskId) {
            return Xrm.Navigation.openAlertDialog({
                text: MSG.taskTimeout + " " + woId
            });
        }
        Xrm.Navigation.openForm(
            {
                entityName: "msdyn_workorderservicetask",
                entityId: serviceTaskId
            },
            {
                "duc_account": accountId
            }
        );


    } catch (e) {
        var errMsg = "[oFST] Error while navigating to service task."
            + "\nOffline: " + isOffline()
            + "\nError: " + (e.message || JSON.stringify(e));
        console.error(errMsg, e);
        Xrm.Navigation.openAlertDialog({ text: errMsg });
    }
}

async function uLA(woId) {
    var step = "";
    try {
        woId = woId.replace(/[{}]/g, "");

        var isOff = isOffline();

        step = "Retrieving work order process extension";
        var woRecord = await Xrm.WebApi.retrieveRecord(
            "msdyn_workorder",
            woId,
            "?$select=_duc_processextension_value"
        );

        var peId = woRecord._duc_processextension_value;
        if (!peId) return;

        step = "Retrieving process extension record";
        var peRecord = await Xrm.WebApi.retrieveRecord(
            "duc_processextension",
            peId,
            "?$select=_duc_processdefinition_value,_duc_currentstage_value"
        );

        var processDefinitionId = peRecord._duc_processdefinition_value;
        var currentStageId = peRecord._duc_currentstage_value;

        if (!processDefinitionId) return;

        step = "Retrieving stage actions";
        var processField = isOff ? "duc_process" : "_duc_process_value";
        var relatedStageField = isOff ? "duc_relatedstage" : "_duc_relatedstage_value";

        var stageActions = await Xrm.WebApi.retrieveMultipleRecords(
            "duc_stageaction",
            "?$select=duc_stageactionid,_duc_defaultstatus_value" +
            "&$filter=duc_canbetriggeredbytarget eq true" +
            " and " + processField + " eq " + processDefinitionId +
            " and " + relatedStageField + " eq " + currentStageId +
            "&$top=10"
        );

        var actionIdToSet = null;

        step = "Looping stage actions to find matching status";
        for (var i = 0; i < stageActions.entities.length; i++) {
            var item = stageActions.entities[i];
            var statusId = item._duc_defaultstatus_value;

            if (statusId) {
                var status = await Xrm.WebApi.retrieveRecord(
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

        step = "Updating process extension with LastActionTaken";
        await Xrm.WebApi.updateRecord("duc_processextension", peId, {
            "duc_LastActionTaken_duc_ProcessExtension@odata.bind":
                "/duc_stageactions(" + actionIdToSet + ")"
        });

        // Run offline plugin logic inline (no external dependency)
        step = "Running offline action logic";
        await runOfflineActionLogic(actionIdToSet, peId, woId);

    } catch (e) {
        var errMsg = "[uLA] Error at step: " + step
            + "\nOffline: " + isOffline()
            + "\nError: " + (e.message || JSON.stringify(e));
        console.warn(errMsg, e);
        Xrm.Navigation.openAlertDialog({ text: errMsg });
    }
}

// Self-contained offline plugin logic for start_insoection_offline.js
// Replicates what the server-side plugin does when duc_lastactiontaken changes
async function runOfflineActionLogic(actionId, processExtensionId, workOrderId) {
    var step = "";
    try {
        var userId = Xrm.Utility.getGlobalContext().userSettings.userId.replace(/[{}]/g, "");

        // 1. Get action info from duc_stageaction
        step = "Retrieving stage action details";
        var action = await Xrm.WebApi.retrieveRecord("duc_stageaction", actionId,
            "?$select=duc_name,_duc_actiontype_value,_duc_nextstage_value,_duc_relatedstage_value," +
            "_duc_defaultstatus_value,_duc_defaultsubstatus_value"
        );

        var actionTypeId = action._duc_actiontype_value || null;
        var nextStageId = action._duc_nextstage_value || null;
        var statusId = action._duc_defaultstatus_value || null;
        var subStatusId = action._duc_defaultsubstatus_value || null;
        var relatedStageId = action._duc_relatedstage_value || null;

        // 2. Get action type info (for sendtocustomer flag)
        var sendToCustomer = false;
        if (actionTypeId) {
            step = "Retrieving action type";
            try {
                var actionType = await Xrm.WebApi.retrieveRecord("duc_actiontype", actionTypeId,
                    "?$select=duc_sendtocustomer"
                );
                sendToCustomer = actionType.duc_sendtocustomer || false;
            } catch (e) { /* ignore - non-critical */ }
        }

        // 3. Get process definition from related stage (for target entity status mapping)
        var processDefId = null;
        var targetStatusField = null;
        var targetSubStatusField = null;
        var existingStatusEntity = null;
        var existingSubStatusEntity = null;
        var statusLookupType = null;
        var subStatusLookupType = null;

        if (relatedStageId) {
            step = "Retrieving related stage";
            try {
                var relatedStage = await Xrm.WebApi.retrieveRecord("duc_processstage", relatedStageId,
                    "?$select=_duc_relatedprocess_value"
                );
                processDefId = relatedStage._duc_relatedprocess_value || null;
            } catch (e) { /* ignore */ }
        }

        if (processDefId) {
            step = "Retrieving process definition";
            try {
                var procDef = await Xrm.WebApi.retrieveRecord("duc_processdefinition", processDefId,
                    "?$select=duc_targetentitystatuslookupname,duc_targetentitysubstatuslookupname," +
                    "duc_existingstatusentity,duc_existingsubstatusentity,duc_statuslookuptype,duc_substatuslookuptype"
                );
                targetStatusField = procDef.duc_targetentitystatuslookupname || null;
                targetSubStatusField = procDef.duc_targetentitysubstatuslookupname || null;
                existingStatusEntity = procDef.duc_existingstatusentity || null;
                existingSubStatusEntity = procDef.duc_existingsubstatusentity || null;
                statusLookupType = procDef.duc_statuslookuptype || null;
                subStatusLookupType = procDef.duc_substatuslookuptype || null;
            } catch (e) { /* ignore */ }
        }

        // 4. Resolve status/substatus values
        var statusValue = null;
        var subStatusValue = null;

        if (statusId) {
            step = "Retrieving process status value";
            try {
                var statusRec = await Xrm.WebApi.retrieveRecord("duc_processstatus", statusId, "?$select=duc_value");
                statusValue = statusRec.duc_value || null;
            } catch (e) { /* ignore */ }
        }

        if (subStatusId) {
            step = "Retrieving process substatus value";
            try {
                var subStatusRec = await Xrm.WebApi.retrieveRecord("duc_processsubstatus", subStatusId, "?$select=duc_value");
                subStatusValue = subStatusRec.duc_value || null;
            } catch (e) { /* ignore */ }
        }

        // // 5. Create action log
        // step = "Creating action log";
        // var logData = {
        //     "subject": "Process Action",
        //     "actualstart": new Date(),
        //     "duc_sendtocustomer": sendToCustomer,
        //     "duc_islastactiontakenoffline": true
        // };
        // logData["ownerid_duc_processactionlog@odata.bind"] = "/systemusers(" + userId + ")";
        // logData["duc_AuthorId_duc_processActionLog_systemuser@odata.bind"] = "/systemusers(" + userId + ")";
        // logData["regardingobjectid_msdyn_workorder_duc_processactionlog@odata.bind"] = "/msdyn_workorders(" + workOrderId + ")";
        // logData["duc_Action_duc_processActionLog@odata.bind"] = "/duc_stageactions(" + actionId + ")";
        // if (actionTypeId) logData["duc_ActionType_duc_processActionLog@odata.bind"] = "/duc_actiontypes(" + actionTypeId + ")";
        // if (relatedStageId) logData["duc_processStage_duc_processActionLog@odata.bind"] = "/duc_processstages(" + relatedStageId + ")";
        // if (processDefId) logData["duc_process_duc_processActionLog@odata.bind"] = "/duc_processdefinitions(" + processDefId + ")";

        // await Xrm.WebApi.createRecord("duc_processactionlog", logData);

        // 6. Update process extension fields
        step = "Updating process extension fields";
        var peUpdate = { "duc_islastactiontakenoffline": true };
        if (nextStageId) peUpdate["duc_CurrentStage_duc_ProcessExtension@odata.bind"] = "/duc_processstages(" + nextStageId + ")";
        peUpdate["duc_LastApprovedById_duc_ProcessExtension_systemuser@odata.bind"] = "/systemusers(" + userId + ")";
        if (statusId) peUpdate["duc_Status_duc_ProcessExtension@odata.bind"] = "/duc_processstatuses(" + statusId + ")";
        if (subStatusId) peUpdate["duc_SubStatus_duc_ProcessExtension@odata.bind"] = "/duc_processsubstatuses(" + subStatusId + ")";

        await Xrm.WebApi.updateRecord("duc_processextension", processExtensionId, peUpdate);

        // 7. Update work order status/substatus (if process definition maps them)
        step = "Updating work order status";
        var woUpdate = {};
        var hasWoUpdate = false;

        if (targetStatusField && statusValue) {
            if (statusLookupType === 780500002) {
                woUpdate[targetStatusField] = parseInt(statusValue, 10);
            } else if (existingStatusEntity) {
                woUpdate[targetStatusField + "@odata.bind"] = "/" + _entitySetName(existingStatusEntity) + "(" + statusValue + ")";
            }
            hasWoUpdate = true;
        }
        if (targetSubStatusField && subStatusValue) {
            if (subStatusLookupType === 780500002) {
                woUpdate[targetSubStatusField] = parseInt(subStatusValue, 10);
            } else if (existingSubStatusEntity) {
                woUpdate[targetSubStatusField + "@odata.bind"] = "/" + _entitySetName(existingSubStatusEntity) + "(" + subStatusValue + ")";
            }
            hasWoUpdate = true;
        }

        if (hasWoUpdate) {
            await Xrm.WebApi.updateRecord("msdyn_workorder", workOrderId, woUpdate);
        }

    } catch (e) {
        var errMsg = "[runOfflineActionLogic] Error at step: " + step
            + "\nActionId: " + actionId
            + "\nPE Id: " + processExtensionId
            + "\nOffline: " + isOffline()
            + "\nError: " + (e.message || JSON.stringify(e));
        console.error(errMsg, e);
        Xrm.Navigation.openAlertDialog({ text: errMsg });
    }
}

function _entitySetName(logicalName) {
    var map = {
        "msdyn_workordersubstatus": "msdyn_workordersubstatuses",
        "msdyn_workorderstatus": "msdyn_workorderstatuses",
        "duc_processstatus": "duc_processstatuses",
        "duc_processsubstatus": "duc_processsubstatuses",
        "duc_processstage": "duc_processstages",
        "msdyn_workorder": "msdyn_workorders"
    };
    if (map[logicalName]) return map[logicalName];
    if (/(?:s|x|z|ch|sh)$/.test(logicalName)) return logicalName + "es";
    return logicalName + "s";
}

async function sBIP(woId) {
    var step = "";
    try {
        woId = woId.replace(/[{}]/g, "");
        var isOff = isOffline();
        var woFilterField = isOff ? "msdyn_workorder" : "_msdyn_workorder_value";

        step = "Retrieving bookable resource booking for work order";
        var bookingResult = await Xrm.WebApi.retrieveMultipleRecords(
            "bookableresourcebooking",
            "?$select=bookableresourcebookingid" +
            "&$filter=" + woFilterField + " eq " + woId +
            "&$orderby=createdon asc&$top=1"
        );

        if (!bookingResult.entities.length) return false;

        var bookingId = bookingResult.entities[0].bookableresourcebookingid;

        step = "Retrieving booking status 'In Progress'";
        var statusResult = await Xrm.WebApi.retrieveMultipleRecords(
            "bookingstatus",
            "?$select=bookingstatusid&$filter=name eq 'In Progress'"
        );

        if (!statusResult.entities.length) {
            Xrm.Navigation.openAlertDialog({ text: "[sBIP] Booking Status 'In Progress' not found." });
            return false;
        }

        var statusId = statusResult.entities[0].bookingstatusid;

        step = "Updating booking record with 'In Progress' status";
        await Xrm.WebApi.updateRecord("bookableresourcebooking", bookingId, {
            "BookingStatus@odata.bind": "/bookingstatuses(" + statusId + ")"
        });

        return true;

    } catch (e) {
        var errMsg = "[sBIP] Error at step: " + step
            + "\nOffline: " + isOffline()
            + "\nError: " + (e.message || JSON.stringify(e));
        console.error(errMsg, e);
        Xrm.Navigation.openAlertDialog({ text: errMsg });
        return false;
    }
}

async function hasWorkOrderServiceTask(workOrderId) {
    try {
        if (!workOrderId) {
            return false;
        }

        const cleanWorkOrderId = workOrderId.replace(/[{}]/g, "");

        const isOff = isOffline();
        const woFilterField = isOff ? "msdyn_workorder" : "_msdyn_workorder_value";

        const result = await Xrm.WebApi.retrieveMultipleRecords(
            "msdyn_workorderservicetask",
            `?$select=msdyn_workorderservicetaskid&$filter=${woFilterField} eq ${cleanWorkOrderId}&$top=1`
        );

        return result.entities && result.entities.length > 0;
    } catch (e) {
        var errMsg = "[hasWorkOrderServiceTask] Error checking service tasks."
            + "\nWorkOrderId: " + workOrderId
            + "\nOffline: " + isOffline()
            + "\nError: " + (e.message || JSON.stringify(e));
        console.error(errMsg, e);
        Xrm.Navigation.openAlertDialog({ text: errMsg });
        return false;
    }
}

//bassem
function IsFromMobile() {
    if (Xrm.Page.context.client.getClient() == "Mobile"
        && (Xrm.Page.context.client.getFormFactor() == 2 || Xrm.Page.context.client.getFormFactor() == 3)
    ) {
        return true;
    }
    return false;
}
//bassem

function isOffline() {
    return Xrm.Utility.getGlobalContext().client.isOffline();
}
//bassem

async function createDailyInspection() {
    var step = "";
    try {
        Xrm.Utility.showProgressIndicator("جاري الإنشاء الآن");

        const workOrderId = Xrm.Page.data.entity.getId().replace(/[{}]/g, "");

        if (!workOrderId) {
            Xrm.Navigation.openAlertDialog({
                text: "[createDailyInspection] Unable to retrieve Work Order ID. Please save the form first."
            });
            Xrm.Utility.closeProgressIndicator();
            return;
        }

        const createdBy = Xrm.Utility.getGlobalContext().userSettings.userId.replace(/[{}]/g, "");
        const offline = isOffline();

        // Bookable resource query — offline uses 'userid', online uses '_userid_value'
        step = "Retrieving bookable resource for current user (userId=" + createdBy + ")";
        const resourceFilter = offline
            ? `?$select=bookableresourceid&$filter=userid eq ${createdBy} and resourcetype eq 3`
            : `?$select=bookableresourceid&$filter=_userid_value eq ${createdBy} and resourcetype eq 3`;

        const resourceRes = await Xrm.WebApi.retrieveMultipleRecords("bookableresource", resourceFilter);

        if (!resourceRes.entities || resourceRes.entities.length === 0) {
            Xrm.Navigation.openAlertDialog({
                text: "[createDailyInspection] No bookable resource found for the current user."
                    + "\nUserId: " + createdBy
                    + "\nOffline: " + offline
            });
            Xrm.Utility.closeProgressIndicator();
            return;
        }

        const resourceId = resourceRes.entities[0].bookableresourceid;

        const currentDateTime = new Date();

        // Get today's date without time
        const todayStart = new Date(currentDateTime.getFullYear(), currentDateTime.getMonth(), currentDateTime.getDate());

        // Attendance query — offline uses 'duc_user' + ge/lt date range, online uses '_duc_user_value' + Microsoft.Dynamics.CRM.On
        step = "Retrieving attendance record for today";
        let attendanceFilter;
        if (offline) {
            const tomorrowStart = new Date(todayStart);
            tomorrowStart.setDate(tomorrowStart.getDate() + 1);
            attendanceFilter = `?$select=duc_attendanceid&$filter=duc_user eq ${createdBy} and createdon ge ${todayStart.toISOString()} and createdon lt ${tomorrowStart.toISOString()}`;
        } else {
            attendanceFilter = `?$select=duc_attendanceid&$filter=_duc_user_value eq ${createdBy} and Microsoft.Dynamics.CRM.On(PropertyName='createdon',PropertyValue='${todayStart.toISOString()}')`;
        }

        const attendanceRes = await Xrm.WebApi.retrieveMultipleRecords("duc_attendance", attendanceFilter);

        let attendanceId;

        if (attendanceRes.entities && attendanceRes.entities.length > 0) {
            // Use existing attendance record
            attendanceId = attendanceRes.entities[0].duc_attendanceid;
            console.log("Using existing attendance record: " + attendanceId);
        } else {
            // Create new attendance record
            step = "Creating new attendance record";
            const attendanceData = {
                "duc_User@odata.bind": `/systemusers(${createdBy})`,
                "duc_createdoffline": offline
            };

            const attendanceResult = await Xrm.WebApi.createRecord("duc_attendance", attendanceData);
            attendanceId = attendanceResult.id;
            console.log("Created new attendance record: " + attendanceId);
        }

        // Create daily inspection record with attendance lookup
        step = "Creating daily inspection record";
        const recordData = {
            "duc_WorkOrder@odata.bind": `/msdyn_workorders(${workOrderId})`,
            "duc_BookableResource@odata.bind": `/bookableresources(${resourceId})`,
            "duc_Attendance@odata.bind": `/duc_attendances(${attendanceId})`,
            "duc_startinspectiontime": currentDateTime,
            "duc_createdoffline": offline
        };

        const result = await Xrm.WebApi.createRecord("duc_dailyinspectorinspections", recordData);

        // Safely refresh the form if available
        if (Xrm.Page && Xrm.Page.data && Xrm.Page.data.refresh) {
            Xrm.Page.data.refresh(false);
        }

        console.log("Daily inspection record created with ID: " + result.id);

        Xrm.Utility.closeProgressIndicator();

    } catch (error) {
        Xrm.Utility.closeProgressIndicator();

        var errMsg = "[createDailyInspection] Error at step: " + step
            + "\nOffline: " + isOffline()
            + "\nError: " + (error.message || JSON.stringify(error));
        console.error(errMsg, error);
        Xrm.Navigation.openAlertDialog({ text: errMsg });
    }
}


//bassem
function WO_CheckAllowedDistance() {
    var lang = Xrm.Utility.getGlobalContext().userSettings.languageId;
    var MSG = (lang === 1025) ? {
        cannotGetLoc: "تعذر استرداد الموقع الحالي، يرجى مراجعة إعدادات الموقع في الهاتف.",
        checking: "جاري التحقق من الموقع...",
        outOfRange: "الموقع الحالي يبعد بمسافة {0} متر عن موقع التفتيش المسجل وأقصى مسافة مسموحة هي {1}.",
        outOfRangeTitle: "خارج المسافة المسموحة",
        updateError: "فشل تحديث الموقع على أمر العمل.",
        noSubAccount: "لا يوجد حساب فرعي محدد في أمر العمل.",
        noAddresses: "لم يتم العثور على عناوين مرتبطة بالحساب الفرعي."
    } : {
        cannotGetLoc: "Can't get current location. Please check your mobile location settings.",
        checking: "Checking location...",
        outOfRange: "The current location is {0} meters away from the registered location and max allowed distance is {1}.",
        outOfRangeTitle: "Out of Range",
        updateError: "Failed to update location on work order.",
        noSubAccount: "No sub-account is selected on the Work Order.",
        noAddresses: "No addresses found linked to the sub-account."
    };

    Xrm.Utility.showProgressIndicator(MSG.checking);

    var workorderId = Xrm.Page.data.entity.getId().replace(/[{}]/g, "");
    var isOff = isOffline();

    return Xrm.WebApi.retrieveRecord("msdyn_workorder", workorderId,
        "?$select=duc_details_islandmark,duc_details_alloweddistance,duc_channel,_duc_department_value"
    ).then(function (workOrder) {

        var allowedDist = null;
        var isLandmark = workOrder.duc_details_islandmark;
        var isBSS = workOrder.duc_channel === 100000002;

        if (isLandmark && workOrder.duc_details_alloweddistance) {
            allowedDist = workOrder.duc_details_alloweddistance;
            return proceedWithDistanceCheck(allowedDist);
        }
        else if (workOrder._duc_department_value) {
            var ouId = workOrder._duc_department_value.replace(/[{}]/g, "");
            var fieldName = isBSS ? "duc_bssdistanceinmeter" : "duc_distanceinmeter";

            return Xrm.WebApi.retrieveRecord("msdyn_organizationalunit", ouId,
                "?$select=" + fieldName
            ).then(function (ou) {
                if (ou[fieldName] != null) {
                    allowedDist = ou[fieldName] - 100000000;
                } else {
                    allowedDist = 0;
                }
                return proceedWithDistanceCheck(allowedDist);
            }).catch(function (error) {
                var errMsg = "[WO_CheckAllowedDistance] Error retrieving organizational unit."
                    + "\nOU Id: " + ouId
                    + "\nFieldName: " + fieldName
                    + "\nOffline: " + isOff
                    + "\nError: " + ((error && error.message) ? error.message : JSON.stringify(error));
                console.error(errMsg, error);
                Xrm.Navigation.openAlertDialog({ text: errMsg });
                allowedDist = 0;
                return proceedWithDistanceCheck(allowedDist);
            });
        }
        else {
            allowedDist = 0;
            return proceedWithDistanceCheck(allowedDist);
        }

    }).catch(function (error) {
        Xrm.Utility.closeProgressIndicator();
        var errMsg = "[WO_CheckAllowedDistance] Error retrieving work order."
            + "\nWorkOrderId: " + workorderId
            + "\nOffline: " + isOff
            + "\nError: " + ((error && error.message) ? error.message : JSON.stringify(error));
        console.error(errMsg, error);
        Xrm.Navigation.openAlertDialog({ text: errMsg });
        return Promise.resolve(false);
    });

    function proceedWithDistanceCheck(allowedDistance) {
        if (!allowedDistance || allowedDistance <= 0) {
            return getCurrentLocationAndUpdate(null, null, null);
        }

        // First, try to get address from duc_address lookup
        var addrAttr = Xrm.Page.getAttribute("duc_address");
        var addrVal = addrAttr ? addrAttr.getValue() : null;

        if (addrVal && addrVal[0] && addrVal[0].id) {
            var addressId = addrVal[0].id.replace(/[{}]/g, "");
            var addressEntity = "duc_addressinformation";

            return Xrm.WebApi.retrieveRecord(addressEntity, addressId, "?$select=duc_latitude,duc_longitude").then(
                function (result) {
                    var destLat = result.duc_latitude;
                    var destLng = result.duc_longitude;

                    if (destLat != null && destLng != null) {
                        return getCurrentLocationAndUpdate(destLat, destLng, allowedDistance);
                    } else {
                        // If address doesn't have coordinates, fall back to account addresses
                        return checkNearestAccountAddress(allowedDistance);
                    }
                },
                function (error) {
                    var errMsg = "[WO_CheckAllowedDistance > proceedWithDistanceCheck] Error retrieving address."
                        + "\nAddressId: " + addressId
                        + "\nOffline: " + isOff
                        + "\nError: " + ((error && error.message) ? error.message : JSON.stringify(error));
                    console.error(errMsg, error);
                    Xrm.Navigation.openAlertDialog({ text: errMsg });
                    // If error retrieving address, fall back to account addresses
                    return checkNearestAccountAddress(allowedDistance);
                }
            );
        } else {
            // No duc_address lookup, check account addresses
            return checkNearestAccountAddress(allowedDistance);
        }
    }

    function checkNearestAccountAddress(allowedDistance) {
        // Get the sub-account from the work order
        var subAccountAttr = Xrm.Page.getAttribute("duc_subaccount");
        var subAccountVal = subAccountAttr ? subAccountAttr.getValue() : null;

        if (!subAccountVal || !subAccountVal[0] || !subAccountVal[0].id) {
            // No sub-account, proceed without distance check
            return getCurrentLocationAndUpdate(null, null, null);
        }

        var accountId = subAccountVal[0].id.replace(/[{}]/g, "");

        // Build OData query to fetch addresses linked to the account
        var accField = isOff ? "duc_account" : "_duc_account_value";

        var options = "?$select=duc_addressinformationid,duc_latitude,duc_longitude";
        options += "&$filter=" + accField + " eq " + accountId;
        options += " and duc_latitude ne null and duc_longitude ne null";

        return Xrm.WebApi.retrieveMultipleRecords("duc_addressinformation", options).then(
            function (result) {
                if (!result.entities || result.entities.length === 0) {
                    // No addresses found, proceed without distance check
                    return getCurrentLocationAndUpdate(null, null, null);
                }

                // Filter addresses with valid coordinates
                var addresses = result.entities.filter(function (addr) {
                    return addr.duc_latitude != null && addr.duc_longitude != null;
                });

                if (addresses.length === 0) {
                    // No valid addresses, proceed without distance check
                    return getCurrentLocationAndUpdate(null, null, null);
                }

                // Get current location first to find nearest address
                return Xrm.Device.getCurrentPosition().then(function (location) {
                    var originLat = location.coords.latitude;
                    var originLng = location.coords.longitude;

                    // Calculate distance to each address and find the nearest
                    var nearestAddress = null;
                    var minDistance = Infinity;

                    addresses.forEach(function (addr) {
                        var distance = GetDistance(originLat, originLng, addr.duc_latitude, addr.duc_longitude);
                        if (distance < minDistance) {
                            minDistance = distance;
                            nearestAddress = {
                                lat: addr.duc_latitude,
                                lng: addr.duc_longitude,
                                distance: distance
                            };
                        }
                    });

                    // Check if nearest address is within allowed distance
                    if (nearestAddress && allowedDistance > 0 && nearestAddress.distance > allowedDistance) {
                        Xrm.Utility.closeProgressIndicator();

                        var errorMsg = MSG.outOfRange
                            .replace("{0}", Math.round(nearestAddress.distance))
                            .replace("{1}", allowedDistance);

                        Xrm.Navigation.openAlertDialog({
                            title: MSG.outOfRangeTitle,
                            text: errorMsg
                        });

                        return Promise.resolve(false);
                    }

                    // Within range or no distance check needed, update work order
                    return updateWorkOrderLocation(originLat, originLng);

                }, function (error) {
                    Xrm.Utility.closeProgressIndicator();
                    var errMsg = "[WO_CheckAllowedDistance > checkNearestAccountAddress] Cannot get current position."
                        + "\nOffline: " + isOff
                        + "\nError: " + ((error && error.message) ? error.message : JSON.stringify(error));
                    console.error(errMsg, error);
                    Xrm.Navigation.openAlertDialog({ text: errMsg });
                    return Promise.resolve(false);
                });
            },
            function (error) {
                var errMsg = "[WO_CheckAllowedDistance > checkNearestAccountAddress] Error retrieving addresses."
                    + "\nAccountId: " + accountId
                    + "\nOffline: " + isOff
                    + "\nError: " + ((error && error.message) ? error.message : JSON.stringify(error));
                console.error(errMsg, error);
                Xrm.Navigation.openAlertDialog({ text: errMsg });
                // Error retrieving addresses, proceed without distance check
                return getCurrentLocationAndUpdate(null, null, null);
            }
        );
    }

    function getCurrentLocationAndUpdate(destLat, destLng, allowedDistance) {
        return Xrm.Device.getCurrentPosition().then(function (location) {
            var originLat = location.coords.latitude;
            var originLng = location.coords.longitude;

            if (destLat != null && destLng != null && allowedDistance != null && allowedDistance > 0) {
                var distance = GetDistance(originLat, originLng, destLat, destLng);

                if (distance > allowedDistance) {
                    Xrm.Utility.closeProgressIndicator();

                    var errorMsg = MSG.outOfRange.replace("{0}", Math.round(distance)).replace("{1}", allowedDistance);

                    Xrm.Navigation.openAlertDialog({
                        title: MSG.outOfRangeTitle,
                        text: errorMsg
                    });

                    return Promise.resolve(false);
                }
            }

            return updateWorkOrderLocation(originLat, originLng);

        }, function (error) {
            Xrm.Utility.closeProgressIndicator();
            var errMsg = "[WO_CheckAllowedDistance > getCurrentLocationAndUpdate] Cannot get current position."
                + "\nOffline: " + isOff
                + "\nError: " + ((error && error.message) ? error.message : JSON.stringify(error));
            console.error(errMsg, error);
            Xrm.Navigation.openAlertDialog({ text: errMsg });
            return Promise.resolve(false);
        });
    }

    function updateWorkOrderLocation(originLat, originLng) {
        var updateData = {
            "msdyn_latitude": originLat,
            "msdyn_longitude": originLng
        };

        return Xrm.WebApi.updateRecord("msdyn_workorder", workorderId, updateData).then(
            function success(result) {
                Xrm.Utility.closeProgressIndicator();

                if (Xrm.Page && Xrm.Page.data && Xrm.Page.data.refresh) {
                    Xrm.Page.data.refresh(false);
                }

                return Promise.resolve(true);
            },
            function (error) {
                Xrm.Utility.closeProgressIndicator();
                var errMsg = "[WO_CheckAllowedDistance > updateWorkOrderLocation] " + MSG.updateError
                    + "\nWorkOrderId: " + workorderId
                    + "\nLat: " + originLat + ", Lng: " + originLng
                    + "\nOffline: " + isOff
                    + "\nError: " + ((error && error.message) ? error.message : JSON.stringify(error));
                console.error(errMsg, error);
                Xrm.Navigation.openAlertDialog({ text: errMsg });
                return Promise.resolve(false);
            }
        );
    }
}
//bassem
function GetDistance(lat1, lon1, lat2, lon2) {
    var R = 6371000;
    var dLat = (lat2 - lat1) * Math.PI / 180;
    var dLon = (lon2 - lon1) * Math.PI / 180;
    var a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
        Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
        Math.sin(dLon / 2) * Math.sin(dLon / 2);
    var c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
    var distance = R * c;
    return distance;
}
// ============================================
// MAIN EXECUTION - ONLY CHECK DISTANCE ON MOBILE
// ============================================
(async function () {
    try {
        const workOrderId = Xrm.Page.data.entity.getId();
        const isMobile = IsFromMobile();
        const isOff = isOffline();

        // Step 1: Check if service tasks exist


        // Step 2: Check distance ONLY if on mobile
        if (isMobile) {
            const distanceCheckPassed = await WO_CheckAllowedDistance();

            // Step 3: Only create daily inspection if distance check passed
            if (distanceCheckPassed) {
                await createDailyInspection();
            } else {
                // Distance check failed - stop execution
                return;
            }
        }

        // Step 4: Capture work order ID BEFORE redirect
        var cleanWoId = workOrderId.replace(/[{}]/g, "");

        // Step 5: Navigate to service task (this redirects the page)
        oFST(cleanWoId);

        // Step 6: Use captured work order ID (page may have changed after oFST)
        var updateBookableResourcesBooking = await sBIP(cleanWoId);

        if (updateBookableResourcesBooking) {
            await uLA(cleanWoId);
        }

    } catch (error) {
        // If any error occurs, stop execution
        var errMsg = "[MAIN EXECUTION] Execution stopped due to error."
            + "\nOffline: " + isOffline()
            + "\nIsMobile: " + IsFromMobile()
            + "\nError: " + (error.message || JSON.stringify(error));
        console.error(errMsg, error);
        Xrm.Navigation.openAlertDialog({ text: errMsg });
    }
})();