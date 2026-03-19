// ============================================
// Booking -> WorkOrder Full Flow (Same Behavior)
// ============================================

function getMSG() {
    var lang = Xrm.Utility.getGlobalContext().userSettings.languageId;
    return (lang === 1025) ? {
        waitingTask: "جاري إنشاء مهام الخدمة، يرجى الانتظار...",
        taskTimeout: "لم يتم إنشاء مهام الخدمة بعد. يرجى المحاولة مرة أخرى بعد قليل.",
        taskNotFound: "لم يتم العثور على مهام الخدمة.",
        woNotFound: "لم يتم العثور على رقم أمر العمل.",
        cannotGetLoc: "تعذر استرداد الموقع الحالي، يرجى مراجعة إعدادات الموقع في الهاتف.",
        checking: "جاري التحقق من الموقع...",
        outOfRange: "الموقع الحالي يبعد بمسافة {0} متر عن موقع التفتيش المسجل وأقصى مسافة مسموحة هي {1}.",
        outOfRangeTitle: "خارج المسافة المسموحة",
        updateError: "فشل تحديث الموقع على أمر العمل.",
        creatingInspection: "جاري الإنشاء الآن",
        noBookableResource: "لم يتم العثور على Bookable Resource للمستخدم الحالي.",
        bookingStatusNotFound: "لم يتم العثور على Booking Status باسم 'In Progress'."
    } : {
        waitingTask: "Creating service tasks, please wait...",
        taskTimeout: "Service tasks are not created yet. Please try again in a few moments.",
        taskNotFound: "No service tasks found.",
        woNotFound: "Work Order ID not found.",
        cannotGetLoc: "Can't get current location. Please check your mobile location settings.",
        checking: "Checking location...",
        outOfRange: "The current location is {0} meters away from the registered location and max allowed distance is {1}.",
        outOfRangeTitle: "Out of Range",
        updateError: "Failed to update location on work order.",
        creatingInspection: "Creating now...",
        noBookableResource: "No bookable resource found for the current user.",
        bookingStatusNotFound: "Booking Status 'In Progress' not found."
    };
}

function isOffline() {
    return Xrm.Utility.getGlobalContext().client.isOffline();
}

function IsFromMobile() {
    try {
        return (Xrm.Page.context.client.getClient() === "Mobile" &&
            (Xrm.Page.context.client.getFormFactor() === 2 || Xrm.Page.context.client.getFormFactor() === 3));
    } catch (e) {
        return false;
    }
}

function GetDistance(lat1, lon1, lat2, lon2) {
    var R = 6371000;
    var dLat = (lat2 - lat1) * Math.PI / 180;
    var dLon = (lon2 - lon1) * Math.PI / 180;
    var a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
        Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
        Math.sin(dLon / 2) * Math.sin(dLon / 2);
    var c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
    return R * c;
}

function getWorkOrderIdFromBooking() {
    var attr = Xrm.Page.getAttribute("msdyn_workorder");
    var val = attr ? attr.getValue() : null;
    if (!val || !val[0] || !val[0].id) return null;
    return val[0].id.replace(/[{}]/g, "");
}

// ==================================================
// 1) Wait for Work Order Service Task (same logic)
// ==================================================
async function waitForWorkOrderServiceTask(workOrderId, maxWaitMs, intervalMs) {
    maxWaitMs = maxWaitMs || 30000;
    intervalMs = intervalMs || 2000;

    try {
        const start = Date.now();
        const MSG = getMSG();
        const isOff = isOffline();
        const woFilterField = isOff ? "msdyn_workorder" : "_msdyn_workorder_value";

        Xrm.Utility.showProgressIndicator(MSG.waitingTask);

        // Offline: tasks are either already on device or not — no point polling
        if (isOff) {
            const result = await Xrm.WebApi.retrieveMultipleRecords(
                "msdyn_workorderservicetask",
                `?$select=msdyn_workorderservicetaskid
                 &$filter=${woFilterField} eq ${workOrderId}
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
                 &$filter=${woFilterField} eq ${workOrderId}
                 &$orderby=createdon asc
                 &$top=1`
            );

            if (result.entities && result.entities.length > 0) {
                Xrm.Utility.closeProgressIndicator();
                return result.entities[0].msdyn_workorderservicetaskid;
            }

            await new Promise(function (r) { setTimeout(r, intervalMs); });
        }

        Xrm.Utility.closeProgressIndicator();
        return null;
    } catch (e) {
        Xrm.Utility.closeProgressIndicator();
        var errMsg = "[waitForWorkOrderServiceTask] Error retrieving service tasks."
            + "\nWorkOrderId: " + workOrderId
            + "\nOffline: " + isOffline()
            + "\nError: " + (e.message || JSON.stringify(e));
        console.error(errMsg, e);
        Xrm.Navigation.openAlertDialog({ text: errMsg });
        return null;
    }
}

// ==================================================
// 2) Open First Service Task (from Booking)
// ==================================================
async function oFST_FromBooking(workOrderId) {
    try {
        const MSG = getMSG();
        if (!workOrderId) {
            return Xrm.Navigation.openAlertDialog({ text: MSG.woNotFound });
        }

        // Get Service Account from Work Order
        var wo = await Xrm.WebApi.retrieveRecord(
            "msdyn_workorder",
            workOrderId,
            "?$select=_msdyn_serviceaccount_value"
        );

        var accountId = wo._msdyn_serviceaccount_value ? wo._msdyn_serviceaccount_value.replace(/[{}]/g, "") : null;

        const serviceTaskId = await waitForWorkOrderServiceTask(workOrderId);

        if (!serviceTaskId) {
            return Xrm.Navigation.openAlertDialog({ text: MSG.taskTimeout });
        }

        var params = {};
        if (accountId) params["duc_account"] = accountId;

        return Xrm.Navigation.openForm(
            { entityName: "msdyn_workorderservicetask", entityId: serviceTaskId },
            params
        );

    } catch (e) {
        var errMsg = "[oFST_FromBooking] Error navigating to service task."
            + "\nWorkOrderId: " + workOrderId
            + "\nOffline: " + isOffline()
            + "\nError: " + (e.message || JSON.stringify(e));
        console.error(errMsg, e);
        Xrm.Navigation.openAlertDialog({ text: errMsg });
    }
}

// ==================================================
// 3) Set current Booking status = In Progress
// ==================================================
async function setCurrentBookingInProgress() {
    const MSG = getMSG();
    var step = "";
    try {
        var bookingId = Xrm.Page.data.entity.getId().replace(/[{}]/g, "");

        step = "Retrieving booking status 'In Progress'";
        var statusResult = await Xrm.WebApi.retrieveMultipleRecords(
            "bookingstatus",
            `?$select=bookingstatusid&$filter=name eq 'In Progress'&$top=1`
        );

        if (!statusResult.entities || !statusResult.entities.length) {
            await Xrm.Navigation.openAlertDialog({ text: MSG.bookingStatusNotFound });
            return false;
        }

        var statusId = statusResult.entities[0].bookingstatusid;

        step = "Updating booking record with 'In Progress' status";
        await Xrm.WebApi.updateRecord("bookableresourcebooking", bookingId, {
            "BookingStatus@odata.bind": `/bookingstatuses(${statusId})`
        });

        return true;

    } catch (e) {
        var errMsg = "[setCurrentBookingInProgress] Error at step: " + step
            + "\nOffline: " + isOffline()
            + "\nError: " + (e.message || JSON.stringify(e));
        console.warn(errMsg, e);
        Xrm.Navigation.openAlertDialog({ text: errMsg });
        return false;
    }
}

// ==================================================
// 4) Create Daily Inspection (from Booking -> WorkOrderId)
// ==================================================
async function createDailyInspection_FromBooking(workOrderId) {
    const MSG = getMSG();
    var step = "";
    try {
        Xrm.Utility.showProgressIndicator(MSG.creatingInspection);

        if (!workOrderId) {
            Xrm.Utility.closeProgressIndicator();
            await Xrm.Navigation.openAlertDialog({ text: MSG.woNotFound });
            return;
        }

        const createdBy = Xrm.Utility.getGlobalContext().userSettings.userId.replace(/[{}]/g, "");
        const offline = isOffline();

        // Bookable resource query — offline uses 'userid', online uses '_userid_value'
        step = "Retrieving bookable resource for current user (userId=" + createdBy + ")";
        const resourceFilter = offline
            ? `?$select=bookableresourceid&$filter=userid eq ${createdBy} and resourcetype eq 3&$top=1`
            : `?$select=bookableresourceid&$filter=_userid_value eq ${createdBy} and resourcetype eq 3&$top=1`;

        const resourceRes = await Xrm.WebApi.retrieveMultipleRecords("bookableresource", resourceFilter);

        if (!resourceRes.entities || resourceRes.entities.length === 0) {
            Xrm.Utility.closeProgressIndicator();
            await Xrm.Navigation.openAlertDialog({
                text: "[createDailyInspection_FromBooking] " + MSG.noBookableResource
                    + "\nUserId: " + createdBy
                    + "\nOffline: " + offline
            });
            return;
        }

        const resourceId = resourceRes.entities[0].bookableresourceid;
        const currentDateTime = new Date();
        const todayStart = new Date(currentDateTime.getFullYear(), currentDateTime.getMonth(), currentDateTime.getDate());

        // Attendance query — offline uses 'duc_user' + ge/lt date range, online uses '_duc_user_value' + Microsoft.Dynamics.CRM.On
        step = "Retrieving attendance record for today";
        let attendanceFilter;
        if (offline) {
            const tomorrowStart = new Date(todayStart);
            tomorrowStart.setDate(tomorrowStart.getDate() + 1);
            attendanceFilter = `?$select=duc_attendanceid&$filter=duc_user eq ${createdBy} and createdon ge ${todayStart.toISOString()} and createdon lt ${tomorrowStart.toISOString()}&$top=1`;
        } else {
            attendanceFilter = `?$select=duc_attendanceid&$filter=_duc_user_value eq ${createdBy} and Microsoft.Dynamics.CRM.On(PropertyName='createdon',PropertyValue='${todayStart.toISOString()}')&$top=1`;
        }

        const attendanceRes = await Xrm.WebApi.retrieveMultipleRecords("duc_attendance", attendanceFilter);

        let attendanceId;

        if (attendanceRes.entities && attendanceRes.entities.length > 0) {
            attendanceId = attendanceRes.entities[0].duc_attendanceid;
        } else {
            step = "Creating new attendance record";
            const attendanceResult = await Xrm.WebApi.createRecord("duc_attendance", {
                "duc_User@odata.bind": `/systemusers(${createdBy})`,
                "duc_createdoffline": offline
            });
            attendanceId = attendanceResult.id;
        }

        step = "Creating daily inspection record";
        const recordData = {
            "duc_WorkOrder@odata.bind": `/msdyn_workorders(${workOrderId})`,
            "duc_BookableResource@odata.bind": `/bookableresources(${resourceId})`,
            "duc_Attendance@odata.bind": `/duc_attendances(${attendanceId})`,
            "duc_startinspectiontime": currentDateTime,
            "duc_createdoffline": offline
        };

        await Xrm.WebApi.createRecord("duc_dailyinspectorinspections", recordData);

        Xrm.Utility.closeProgressIndicator();

        if (Xrm.Page && Xrm.Page.data && Xrm.Page.data.refresh) {
            Xrm.Page.data.refresh(false);
        }

    } catch (error) {
        Xrm.Utility.closeProgressIndicator();
        var errMsg = "[createDailyInspection_FromBooking] Error at step: " + step
            + "\nOffline: " + isOffline()
            + "\nError: " + (error.message || JSON.stringify(error));
        console.error(errMsg, error);
        Xrm.Navigation.openAlertDialog({ text: errMsg });
    }
}

// ==================================================
// 5) uLA (from Booking -> WorkOrderId)
// ==================================================
async function uLA_FromBooking(workOrderId) {
    var step = "";
    try {
        if (!workOrderId) return;

        var isOff = isOffline();

        step = "Retrieving work order process extension";
        var woRecord = await Xrm.WebApi.retrieveRecord(
            "msdyn_workorder",
            workOrderId,
            "?$select=_duc_processextension_value"
        );

        if (!woRecord._duc_processextension_value) return;

        var peId = woRecord._duc_processextension_value;

        step = "Retrieving process extension record";
        var peRecord = await Xrm.WebApi.retrieveRecord(
            "duc_processextension",
            peId,
            "?$select=_duc_processdefinition_value,_duc_currentstage_value"
        );

        if (!peRecord._duc_processdefinition_value) return;

        var processDefinitionId = peRecord._duc_processdefinition_value;
        var currentStageId = peRecord._duc_currentstage_value;

        step = "Retrieving stage actions";
        var processField = isOff ? "duc_process" : "_duc_process_value";
        var relatedStageField = isOff ? "duc_relatedstage" : "_duc_relatedstage_value";

        var stageActions = await Xrm.WebApi.retrieveMultipleRecords(
            "duc_stageaction",
            `?$select=duc_stageactionid,_duc_defaultstatus_value
              &$filter=duc_canbetriggeredbytarget eq true
              and ${processField} eq ${processDefinitionId}
              and ${relatedStageField} eq ${currentStageId}
              &$top=10`
        );

        var actionIdToSet = null;

        step = "Looping stage actions to find matching status";
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

        step = "Updating process extension with LastActionTaken";
        await Xrm.WebApi.updateRecord("duc_processextension", peId, {
            "duc_LastActionTaken_duc_ProcessExtension@odata.bind":
                `/duc_stageactions(${actionIdToSet})`
        });

    } catch (e) {
        var errMsg = "[uLA_FromBooking] Error at step: " + step
            + "\nWorkOrderId: " + workOrderId
            + "\nOffline: " + isOffline()
            + "\nError: " + (e.message || JSON.stringify(e));
        console.warn(errMsg, e);
        Xrm.Navigation.openAlertDialog({ text: errMsg });
    }
}

// ==================================================
// 6) Distance check (from Booking, reads WO fields via WebApi)
// ==================================================
async function WO_CheckAllowedDistance_FromBooking(workOrderId) {
    var MSG = getMSG();
    var isOff = isOffline();
    Xrm.Utility.showProgressIndicator(MSG.checking);

    try {
        var workOrder = await Xrm.WebApi.retrieveRecord(
            "msdyn_workorder",
            workOrderId,
            "?$select=duc_details_islandmark,duc_details_alloweddistance,duc_channel,_duc_department_value,_duc_subaccount_value,_duc_address_value"
        );

        var allowedDist = 0;
        var isLandmark = workOrder.duc_details_islandmark;
        var isBSS = workOrder.duc_channel === 100000002;

        // Calculate allowed distance
        if (isLandmark && workOrder.duc_details_alloweddistance) {
            allowedDist = workOrder.duc_details_alloweddistance;
        } else if (workOrder._duc_department_value) {
            var ouId = workOrder._duc_department_value.replace(/[{}]/g, "");
            var fieldName = isBSS ? "duc_bssdistanceinmeter" : "duc_distanceinmeter";

            try {
                var ou = await Xrm.WebApi.retrieveRecord("msdyn_organizationalunit", ouId, "?$select=" + fieldName);
                allowedDist = (ou[fieldName] != null) ? (ou[fieldName] - 100000000) : 0;
            } catch (ouErr) {
                var errMsg = "[WO_CheckAllowedDistance_FromBooking] Error retrieving organizational unit."
                    + "\nOU Id: " + ouId
                    + "\nFieldName: " + fieldName
                    + "\nOffline: " + isOff
                    + "\nError: " + ((ouErr && ouErr.message) ? ouErr.message : JSON.stringify(ouErr));
                console.error(errMsg, ouErr);
                allowedDist = 0;
            }
        } else {
            allowedDist = 0;
        }

        // If no distance checking required -> just update WO current location
        if (!allowedDist || allowedDist <= 0) {
            return await updateWOWithCurrentLocation(workOrderId, null, null, null);
        }

        // 1) Try Work Order address lookup (duc_address)
        if (workOrder._duc_address_value) {
            var addressId = workOrder._duc_address_value.replace(/[{}]/g, "");
            try {
                var addr = await Xrm.WebApi.retrieveRecord("duc_addressinformation", addressId, "?$select=duc_latitude,duc_longitude");
                if (addr.duc_latitude != null && addr.duc_longitude != null) {
                    return await updateWOWithCurrentLocation(workOrderId, addr.duc_latitude, addr.duc_longitude, allowedDist);
                }
            } catch (addrErr) {
                var errMsg2 = "[WO_CheckAllowedDistance_FromBooking] Error retrieving address."
                    + "\nAddressId: " + addressId
                    + "\nOffline: " + isOff
                    + "\nError: " + ((addrErr && addrErr.message) ? addrErr.message : JSON.stringify(addrErr));
                console.error(errMsg2, addrErr);
                // fallback to subaccount addresses
            }
        }

        // 2) Fallback: nearest address from SubAccount addresses
        if (workOrder._duc_subaccount_value) {
            var accountId = workOrder._duc_subaccount_value.replace(/[{}]/g, "");

            var accField = isOff ? "duc_account" : "_duc_account_value";
            var options = "?$select=duc_addressinformationid,duc_latitude,duc_longitude";
            options += "&$filter=" + accField + " eq " + accountId;
            options += " and duc_latitude ne null and duc_longitude ne null";

            var addrList = await Xrm.WebApi.retrieveMultipleRecords("duc_addressinformation", options);

            if (addrList.entities && addrList.entities.length > 0) {
                var loc = await Xrm.Device.getCurrentPosition();
                var originLat = loc.coords.latitude;
                var originLng = loc.coords.longitude;

                var minDistance = Infinity;
                var nearest = null;

                addrList.entities.forEach(function (a) {
                    var d = GetDistance(originLat, originLng, a.duc_latitude, a.duc_longitude);
                    if (d < minDistance) {
                        minDistance = d;
                        nearest = { lat: a.duc_latitude, lng: a.duc_longitude, distance: d };
                    }
                });

                if (nearest && nearest.distance > allowedDist) {
                    Xrm.Utility.closeProgressIndicator();
                    var errorMsg = MSG.outOfRange.replace("{0}", Math.round(nearest.distance)).replace("{1}", allowedDist);
                    await Xrm.Navigation.openAlertDialog({ title: MSG.outOfRangeTitle, text: errorMsg });
                    return false;
                }

                // within range
                return await updateWorkOrderLocationOnly(workOrderId, originLat, originLng);
            }
        }

        // If nothing found, proceed without blocking (same idea as your original fallback)
        return await updateWOWithCurrentLocation(workOrderId, null, null, null);

    } catch (error) {
        Xrm.Utility.closeProgressIndicator();
        var errMsg = "[WO_CheckAllowedDistance_FromBooking] Error."
            + "\nWorkOrderId: " + workOrderId
            + "\nOffline: " + isOff
            + "\nError: " + ((error && error.message) ? error.message : JSON.stringify(error));
        console.error(errMsg, error);
        Xrm.Navigation.openAlertDialog({ text: errMsg });
        return false;
    }

    async function updateWOWithCurrentLocation(workOrderId, destLat, destLng, allowedDistance) {
        try {
            var location = await Xrm.Device.getCurrentPosition();
            var originLat = location.coords.latitude;
            var originLng = location.coords.longitude;

            if (destLat != null && destLng != null && allowedDistance != null && allowedDistance > 0) {
                var distance = GetDistance(originLat, originLng, destLat, destLng);
                if (distance > allowedDistance) {
                    Xrm.Utility.closeProgressIndicator();
                    var msg = MSG.outOfRange.replace("{0}", Math.round(distance)).replace("{1}", allowedDistance);
                    await Xrm.Navigation.openAlertDialog({ title: MSG.outOfRangeTitle, text: msg });
                    return false;
                }
            }

            return await updateWorkOrderLocationOnly(workOrderId, originLat, originLng);

        } catch (e) {
            Xrm.Utility.closeProgressIndicator();
            var errMsg = "[WO_CheckAllowedDistance_FromBooking > updateWOWithCurrentLocation] Cannot get position."
                + "\nOffline: " + isOff
                + "\nError: " + ((e && e.message) ? e.message : JSON.stringify(e));
            console.error(errMsg, e);
            Xrm.Navigation.openAlertDialog({ text: errMsg });
            return false;
        }
    }

    async function updateWorkOrderLocationOnly(workOrderId, originLat, originLng) {
        try {
            await Xrm.WebApi.updateRecord("msdyn_workorder", workOrderId, {
                "msdyn_latitude": originLat,
                "msdyn_longitude": originLng
            });

            Xrm.Utility.closeProgressIndicator();
            return true;

        } catch (e) {
            Xrm.Utility.closeProgressIndicator();
            var errMsg = "[WO_CheckAllowedDistance_FromBooking > updateWorkOrderLocationOnly] " + MSG.updateError
                + "\nWorkOrderId: " + workOrderId
                + "\nOffline: " + isOff
                + "\nError: " + ((e && e.message) ? e.message : JSON.stringify(e));
            console.error(errMsg, e);
            Xrm.Navigation.openAlertDialog({ text: errMsg });
            return false;
        }
    }
}

// ==================================================
// MAIN EXECUTION (Run this from Booking)
// ==================================================
async function Booking_RunFullFlow_SameAsWorkOrder() {
    try {
        const MSG = getMSG();
        const workOrderId = getWorkOrderIdFromBooking();
        const isMobile = IsFromMobile();

        if (!workOrderId) {
            return Xrm.Navigation.openAlertDialog({ text: MSG.woNotFound });
        }

        // 1) Mobile-only: distance check + create inspection
        if (isMobile) {
            const passed = await WO_CheckAllowedDistance_FromBooking(workOrderId);
            if (!passed) return;
            await createDailyInspection_FromBooking(workOrderId);
        }

        // 2) Open first service task (wait until created)
        await oFST_FromBooking(workOrderId);

        // 3) Update CURRENT booking to In Progress
        var updated = await setCurrentBookingInProgress();

        // 4) If updated, update last action
        if (updated) {
            await uLA_FromBooking(workOrderId);
        }

    } catch (error) {
        var errMsg = "[Booking_RunFullFlow_SameAsWorkOrder] Execution stopped due to error."
            + "\nOffline: " + isOffline()
            + "\nIsMobile: " + IsFromMobile()
            + "\nError: " + (error.message || JSON.stringify(error));
        console.error(errMsg, error);
        Xrm.Navigation.openAlertDialog({ text: errMsg });
    }
}

// ======= Call it =======
Booking_RunFullFlow_SameAsWorkOrder();