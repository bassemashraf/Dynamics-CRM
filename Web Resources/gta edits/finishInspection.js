/***************************************
 * LANGUAGE & TRANSLATION HELPERS
 ***************************************/
function getUserLanguage() {
    // Arabic LCID = 1025
    return Xrm.Utility.getGlobalContext().userSettings.languageId === 1025
        ? "ar"
        : "en";
}

const Messages = {
    RequiredFields: {
        en: "Please fill all required fields of the responsible employee data before continuing.",
        ar: "يرجى تعبئة جميع الحقول الإلزامية الخاصة ببيانات الموظف المسؤول قبل المتابعة."
    },
    Processing: {
        en: "Processing...",
        ar: "جاري المعالجة..."
    },
    ProcessingWO: {
        en: "Processing Work Order...",
        ar: "جاري معالجة أمر العمل..."
    },
    WorkOrderNotFound: {
        en: "Work Order ID not found.",
        ar: "لم يتم العثور على رقم أمر العمل."
    },
    BookingCompletedNotFound: {
        en: "Booking Status 'Completed' not found.",
        ar: "لم يتم العثور على حالة الحجز (مكتمل)."
    },
    BookingUpdateError: {
        en: "Error updating booking: ",
        ar: "خطأ أثناء تحديث الحجز: "
    },
    ValidationError: {
        en: "Validation error: ",
        ar: "خطأ في التحقق: "
    }
};

function t(key) {
    const lang = getUserLanguage();
    return Messages[key]?.[lang] || key;
}

/***************************************
 * VALIDATION
 ***************************************/
async function validateRequiredFields() {
    try {
        var form = Xrm.Page;

        var requiredFields = [
            "duc_employeeid",
            "duc_employeename",
            "duc_idexpirydate",
            "duc_nationality",
            "duc_emobile",
            "duc_eaddress",
            "duc_eemail",
            "duc_customersignature"
        ];

        var refusedToSign = form.getAttribute("duc_responsibleemployeerefusedtosign")?.getValue();
        var notAvailable = form.getAttribute("duc_responsibleemployeeisnotavailable")?.getValue();

        if (refusedToSign === true || notAvailable === true)
            return true;

        for (let field of requiredFields) {
            let attr = form.getAttribute(field);
            if (!attr || attr.getValue() == null || attr.getValue() === "") {
                Xrm.Utility.closeProgressIndicator();
                await Xrm.Navigation.openAlertDialog({
                    text: t("RequiredFields")
                });
                return false;
            }
        }

        return true;
    }
    catch (e) {
        Xrm.Utility.closeProgressIndicator();
        await Xrm.Navigation.openAlertDialog({
            text: t("ValidationError") + e.message
        });
        return false;
    }
}

/***************************************
 * SET BOOKING TO COMPLETED
 ***************************************/
async function setBookingCompletedForWorkOrder() {
    try {
        var form = Xrm.Page;
        var workOrderId = form.data.entity.getId();

        if (!workOrderId) {
            await Xrm.Navigation.openAlertDialog({
                text: t("WorkOrderNotFound")
            });
            return false;
        }

        workOrderId = workOrderId.replace(/[{}]/g, "");

        var bookings = await Xrm.WebApi.retrieveMultipleRecords(
            "bookableresourcebooking",
            `?$select=bookableresourcebookingid,starttime
              &$filter=_msdyn_workorder_value eq ${workOrderId}
              &$orderby=createdon asc`
        );

        if (!bookings.entities || bookings.entities.length === 0)
            return true;

        var statusResult = await Xrm.WebApi.retrieveMultipleRecords(
            "bookingstatus",
            "?$select=bookingstatusid&$filter=name eq 'Completed'"
        );

        if (!statusResult.entities || statusResult.entities.length === 0) {
            await Xrm.Navigation.openAlertDialog({
                text: t("BookingCompletedNotFound")
            });
            return false;
        }

        var completedStatusId = statusResult.entities[0].bookingstatusid;

        // Update all bookings to Completed status
        for (let booking of bookings.entities) {
            await Xrm.WebApi.updateRecord(
                "bookableresourcebooking",
                booking.bookableresourcebookingid,
                {
                    "BookingStatus@odata.bind": `/bookingstatuses(${completedStatusId})`
                }
            );
        }

        return true;
    }
    catch (e) {
        Xrm.Utility.closeProgressIndicator();
        await Xrm.Navigation.openAlertDialog({
            text: t("BookingUpdateError") + e.message
        });
        return false;
    }
}

/***************************************
 * PROCESS EXTENSION UPDATE
 ***************************************/
async function updateLastActionOnProcessExtension() {
    try {
        var form = Xrm.Page;
        var lookup = form.getAttribute("duc_processextension");

        if (!lookup || !lookup.getValue()) return;

        var extId = lookup.getValue()[0].id.replace(/[{}]/g, "");

        var record = await Xrm.WebApi.retrieveRecord(
            "duc_processextension",
            extId,
            "?$select=_duc_processdefinition_value,_duc_currentstage_value"
        );

        if (!record._duc_processdefinition_value) return;

        var actions = await Xrm.WebApi.retrieveMultipleRecords(
            "duc_stageaction",
            "?$select=duc_stageactionid,_duc_defaultstatus_value" +
            "&$filter=duc_canbetriggeredbytarget eq true" +
            " and _duc_process_value eq " + record._duc_processdefinition_value +
            " and _duc_relatedstage_value eq " + record._duc_currentstage_value
        );

        for (var i = 0; i < actions.entities.length; i++) {
            var act = actions.entities[i];
            if (!act._duc_defaultstatus_value) continue;

            var status = await Xrm.WebApi.retrieveRecord(
                "duc_processstatus",
                act._duc_defaultstatus_value,
                "?$select=duc_value"
            );

            if (status.duc_value === 100000006) {
                await Xrm.WebApi.updateRecord(
                    "duc_processextension",
                    extId,
                    {
                        "duc_LastActionTaken_duc_ProcessExtension@odata.bind":
                            "/duc_stageactions(" + act.duc_stageactionid + ")"
                    }
                );

                // Run offline plugin logic inline
                var woId = form.data.entity.getId().replace(/[{}]/g, "");
                await runOfflineActionLogic_Finish(act.duc_stageactionid, extId, woId);
                break;
            }
        }
    }
    catch (e) {
        Xrm.Utility.closeProgressIndicator();
        console.error("[updateLastActionOnProcessExtension]", e);
    }
}

// Self-contained offline plugin logic for finishInspection.js
async function runOfflineActionLogic_Finish(actionId, processExtensionId, workOrderId) {
    var step = "";
    try {
        var userId = Xrm.Utility.getGlobalContext().userSettings.userId.replace(/[{}]/g, "");

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

        var sendToCustomer = false;
        if (actionTypeId) {
            try {
                var actionType = await Xrm.WebApi.retrieveRecord("duc_actiontype", actionTypeId, "?$select=duc_sendtocustomer");
                sendToCustomer = actionType.duc_sendtocustomer || false;
            } catch (e) { /* non-critical */ }
        }

        var processDefId = null;
        var targetStatusField = null, targetSubStatusField = null;
        var existingStatusEntity = null, existingSubStatusEntity = null;
        var statusLookupType = null, subStatusLookupType = null;

        if (relatedStageId) {
            try {
                var rs = await Xrm.WebApi.retrieveRecord("duc_processstage", relatedStageId, "?$select=_duc_relatedprocess_value");
                processDefId = rs._duc_relatedprocess_value || null;
            } catch (e) { /* ignore */ }
        }

        if (processDefId) {
            try {
                var pd = await Xrm.WebApi.retrieveRecord("duc_processdefinition", processDefId,
                    "?$select=duc_targetentitystatuslookupname,duc_targetentitysubstatuslookupname," +
                    "duc_existingstatusentity,duc_existingsubstatusentity,duc_statuslookuptype,duc_substatuslookuptype"
                );
                targetStatusField = pd.duc_targetentitystatuslookupname || null;
                targetSubStatusField = pd.duc_targetentitysubstatuslookupname || null;
                existingStatusEntity = pd.duc_existingstatusentity || null;
                existingSubStatusEntity = pd.duc_existingsubstatusentity || null;
                statusLookupType = pd.duc_statuslookuptype || null;
                subStatusLookupType = pd.duc_substatuslookuptype || null;
            } catch (e) { /* ignore */ }
        }

        var statusValue = null, subStatusValue = null;
        if (statusId) {
            try { statusValue = (await Xrm.WebApi.retrieveRecord("duc_processstatus", statusId, "?$select=duc_value")).duc_value || null; } catch (e) { /* ignore */ }
        }
        if (subStatusId) {
            try { subStatusValue = (await Xrm.WebApi.retrieveRecord("duc_processsubstatus", subStatusId, "?$select=duc_value")).duc_value || null; } catch (e) { /* ignore */ }
        }

        // step = "Creating action log";
        // var logData = {
        //     "subject": "Process Action", "actualstart": new Date(),
        //     "duc_sendtocustomer": sendToCustomer, "duc_islastactiontakenoffline": true
        // };
        // logData["ownerid_duc_processactionlog@odata.bind"] = "/systemusers(" + userId + ")";
        // logData["duc_AuthorId_duc_processActionLog_systemuser@odata.bind"] = "/systemusers(" + userId + ")";
        // logData["regardingobjectid_msdyn_workorder_duc_processactionlog@odata.bind"] = "/msdyn_workorders(" + workOrderId + ")";
        // logData["duc_Action_duc_processActionLog@odata.bind"] = "/duc_stageactions(" + actionId + ")";
        // if (actionTypeId) logData["duc_ActionType_duc_processActionLog@odata.bind"] = "/duc_actiontypes(" + actionTypeId + ")";
        // if (relatedStageId) logData["duc_processStage_duc_processActionLog@odata.bind"] = "/duc_processstages(" + relatedStageId + ")";
        // if (processDefId) logData["duc_process_duc_processActionLog@odata.bind"] = "/duc_processdefinitions(" + processDefId + ")";
        // await Xrm.WebApi.createRecord("duc_processactionlog", logData);

        step = "Updating process extension fields";
        var peUpdate = { "duc_islastactiontakenoffline": true };
        if (nextStageId) peUpdate["duc_CurrentStage_duc_ProcessExtension@odata.bind"] = "/duc_processstages(" + nextStageId + ")";
        peUpdate["duc_LastApprovedById_duc_ProcessExtension_systemuser@odata.bind"] = "/systemusers(" + userId + ")";
        if (statusId) peUpdate["duc_Status_duc_ProcessExtension@odata.bind"] = "/duc_processstatuses(" + statusId + ")";
        if (subStatusId) peUpdate["duc_SubStatus_duc_ProcessExtension@odata.bind"] = "/duc_processsubstatuses(" + subStatusId + ")";
        await Xrm.WebApi.updateRecord("duc_processextension", processExtensionId, peUpdate);

        step = "Updating work order status";
        var woUpdate = {};
        var hasWoUpdate = false;
        if (targetStatusField && statusValue) {
            if (statusLookupType === 780500002) { woUpdate[targetStatusField] = parseInt(statusValue, 10); }
            else if (existingStatusEntity) { woUpdate[targetStatusField + "@odata.bind"] = "/" + _esn(existingStatusEntity) + "(" + statusValue + ")"; }
            hasWoUpdate = true;
        }
        if (targetSubStatusField && subStatusValue) {
            if (subStatusLookupType === 780500002) { woUpdate[targetSubStatusField] = parseInt(subStatusValue, 10); }
            else if (existingSubStatusEntity) { woUpdate[targetSubStatusField + "@odata.bind"] = "/" + _esn(existingSubStatusEntity) + "(" + subStatusValue + ")"; }
            hasWoUpdate = true;
        }
        if (hasWoUpdate) { await Xrm.WebApi.updateRecord("msdyn_workorder", workOrderId, woUpdate); }

    } catch (e) {
        var errMsg = "[runOfflineActionLogic_Finish] Error at step: " + step
            + "\nActionId: " + actionId + "\nPE Id: " + processExtensionId
            + "\nError: " + (e.message || JSON.stringify(e));
        console.error(errMsg, e);
    }
}

function _esn(logicalName) {
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

/***************************************
 * SERVICE TASKS UPDATE
 ***************************************/
async function updateServiceTasksPercent() {
    try {
        var id = Xrm.Page.data.entity.getId();
        if (!id) return;

        id = id.replace(/[{}]/g, "");

        var tasks = await Xrm.WebApi.retrieveMultipleRecords(
            "msdyn_workorderservicetask",
            "?$select=msdyn_workorderservicetaskid&$filter=_msdyn_workorder_value eq " + id
        );

        for (var i = 0; i < tasks.entities.length; i++) {
            await Xrm.WebApi.updateRecord(
                "msdyn_workorderservicetask",
                tasks.entities[i].msdyn_workorderservicetaskid,
                { msdyn_percentcomplete: 100 }
            );
        }
    }
    catch (e) {
        Xrm.Utility.closeProgressIndicator();
    }
}

/***************************************
 * WORK ORDER UPDATE
 ***************************************/
async function setWorkOrderEndTime() {
    try {
        var form = Xrm.Page;
        var id = form.data.entity.getId();
        if (!id) return;

        id = id.replace(/[{}]/g, "");

        await Xrm.WebApi.updateRecord(
            form.data.entity.getEntityName(),
            id,
            {
                duc_workorderendtime: new Date().toISOString(),
                duc_completiondate: new Date().toISOString()
            }
        );
    }
    catch (e) {
        Xrm.Utility.closeProgressIndicator();
    }
}

/***************************************
 * UI HELPERS
 ***************************************/
function clickRefreshButton(refreshbtnId) {
    var btns = document.querySelectorAll("button[data-id*='" + refreshbtnId + "']");
    if (btns.length > 0) btns[0].click();
}

async function navigateToWorkORderTab() {
    var tab = Xrm.Page.ui.tabs.get("General_Tab");
    if (tab) tab.setFocus();
}

/***************************************
 * MAIN EXECUTION
 ***************************************/
async function runProcess() {
    Xrm.Utility.showProgressIndicator(t("Processing"));

    var valid = await validateRequiredFields();
    if (!valid) return;

    var booking = await setBookingCompletedForWorkOrder();
    if (!booking) return;

    await updateLastActionOnProcessExtension();
    await updateServiceTasksPercent();
    await setWorkOrderEndTime();
    await navigateToWorkORderTab();

    setTimeout(function () {
        try { Xrm.Page.data.refresh(true); } catch (e) { }
        Xrm.Utility.closeProgressIndicator();
    }, 3000);
}

runProcess();