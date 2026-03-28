/***************************************
 * LANGUAGE & TRANSLATION HELPERS
 ***************************************/
function getUserLanguage() {
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

function isOffline() {
    try {
        if (Xrm.Utility.getGlobalContext().client.isOffline()) return true;
        if (Xrm.Utility.getGlobalContext().client.getClientState() === "Offline") return true;
    } catch (e) { }
    return false;
}

/**
 * Helper: Fetches the process extension record and matching stage actions fresh from the API.
 * Each caller gets its own live snapshot — no shared cache.
 */
async function _fetchProcessContext(extId, stepPrefix) {
    var isOff = isOffline();

    var peRecord = await Xrm.WebApi.retrieveRecord(
        "duc_processextension",
        extId,
        "?$select=_duc_processdefinition_value,_duc_currentstage_value"
    );

    if (!peRecord._duc_processdefinition_value) {
        console.warn("[" + stepPrefix + "] No process definition on PE record.");
        return null;
    }

    var processField = isOff ? "duc_process" : "_duc_process_value";
    var relatedStageField = isOff ? "duc_relatedstage" : "_duc_relatedstage_value";

    var actionsResult = await Xrm.WebApi.retrieveMultipleRecords(
        "duc_stageaction",
        "?$select=duc_stageactionid,_duc_nextstage_value,_duc_defaultstatus_value,_duc_defaultsubstatus_value," +
        "_duc_actiontype_value,_duc_relatedstage_value,duc_name" +
        "&$filter=duc_canbetriggeredbytarget eq true" +
        " and " + processField + " eq " + peRecord._duc_processdefinition_value +
        " and " + relatedStageField + " eq " + peRecord._duc_currentstage_value
    );

    return {
        peRecord: peRecord,
        stageActions: actionsResult.entities || []
    };
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
        var errMsg = "[validateRequiredFields] Error."
            + "\nOffline: " + isOffline()
            + "\nError: " + (e.message || JSON.stringify(e));
        console.error(errMsg, e);
        Xrm.Navigation.openAlertDialog({ text: errMsg });
        return false;
    }
}

/***************************************
 * SET BOOKING TO COMPLETED
 ***************************************/
async function setBookingCompletedForWorkOrder() {
    var step = "";
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
        var isOff = isOffline();
        var woFilterField = isOff ? "msdyn_workorder" : "_msdyn_workorder_value";

        step = "Retrieving bookings for work order";
        var bookings = await Xrm.WebApi.retrieveMultipleRecords(
            "bookableresourcebooking",
            "?$select=bookableresourcebookingid,starttime" +
            "&$filter=" + woFilterField + " eq " + workOrderId +
            "&$orderby=createdon asc"
        );

        if (!bookings.entities || bookings.entities.length === 0)
            return true;

        step = "Retrieving 'Completed' booking status";
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

        step = "Updating bookings to Completed";
        for (var bi = 0; bi < bookings.entities.length; bi++) {
            await Xrm.WebApi.updateRecord(
                "bookableresourcebooking",
                bookings.entities[bi].bookableresourcebookingid,
                {
                    "BookingStatus@odata.bind": "/bookingstatuses(" + completedStatusId + ")"
                }
            );
        }

        return true;
    }
    catch (e) {
        Xrm.Utility.closeProgressIndicator();
        var errMsg = "[setBookingCompletedForWorkOrder] Error at step: " + step
            + "\nOffline: " + isOffline()
            + "\nError: " + (e.message || JSON.stringify(e));
        console.error(errMsg, e);
        Xrm.Navigation.openAlertDialog({ text: errMsg });
        return false;
    }
}

/***************************************
 * UPDATE CURRENT STAGE ON PROCESS EXTENSION
 * Fetches fresh data each time — no caching
 ***************************************/
async function updateCurrentStageOnProcessExtension(extId) {
    var step = "";
    try {
        step = "Fetching fresh process context";
        var ctx = await _fetchProcessContext(extId, "updateCurrentStageOnProcessExtension");
        if (!ctx || ctx.stageActions.length === 0) {
            console.warn("[updateCurrentStageOnProcessExtension] No stage actions found — skipping.");
            return;
        }

        step = "Finding action with status 690970004";
        var nextStageId = null;

        for (var i = 0; i < ctx.stageActions.length; i++) {
            var act = ctx.stageActions[i];
            var sid = act._duc_defaultstatus_value;
            if (!sid) continue;

            var statusRec = await Xrm.WebApi.retrieveRecord("duc_processstatus", sid, "?$select=duc_value");
            var statusVal = statusRec.duc_value;

            if (statusVal == 690970004) {
                nextStageId = act._duc_nextstage_value;
                break;
            }
        }

        if (!nextStageId) {
            console.warn("[updateCurrentStageOnProcessExtension] No action with status 690970004 found — skipping update.");
            return;
        }

        step = "Updating current stage on process extension";
        await Xrm.WebApi.updateRecord(
            "duc_processextension",
            extId,
            {
                "duc_CurrentStage_duc_ProcessExtension@odata.bind":
                    "/duc_processstages(" + nextStageId + ")"
            }
        );

        console.log("[updateCurrentStageOnProcessExtension] Current stage updated to: " + nextStageId);

    } catch (e) {
        var errMsg = "[updateCurrentStageOnProcessExtension] Error at step: " + step
            + "\nOffline: " + isOffline()
            + "\nError: " + (e.message || JSON.stringify(e));
        console.error(errMsg, e);
        Xrm.Navigation.openAlertDialog({ text: errMsg });
    }
}

/***************************************
 * PROCESS EXTENSION UPDATE (LastActionTaken)
 * Fetches fresh data after stage update — sees latest DB state
 ***************************************/
async function updateLastActionOnProcessExtension(extId, workOrderId) {
    var step = "";
    try {
        step = "Fetching fresh process context (after stage update)";
        var ctx = await _fetchProcessContext(extId, "updateLastActionOnProcessExtension");
        if (!ctx || ctx.stageActions.length === 0) {
            console.warn("[updateLastActionOnProcessExtension] No stage actions found — skipping.");
            return;
        }

        step = "Finding action with status 100000006";
        for (var ai = 0; ai < ctx.stageActions.length; ai++) {
            var act = ctx.stageActions[ai];
            var sid = act._duc_defaultstatus_value;
            if (!sid) continue;

            var statusRec = await Xrm.WebApi.retrieveRecord("duc_processstatus", sid, "?$select=duc_value");
            var statusVal = statusRec.duc_value;

            if (statusVal == 100000006) {
                step = "Updating process extension with LastActionTaken";
                await Xrm.WebApi.updateRecord(
                    "duc_processextension",
                    extId,
                    {
                        "duc_LastActionTaken_duc_ProcessExtension@odata.bind":
                            "/duc_stageactions(" + act.duc_stageactionid + ")"
                    }
                );

                // Run offline plugin logic
                step = "Running offline action logic";
                await runOfflineActionLogic_Finish(act, extId, workOrderId);

                break;
            }
        }
    }
    catch (e) {
        Xrm.Utility.closeProgressIndicator();
        var errMsg = "[updateLastActionOnProcessExtension] Error at step: " + step
            + "\nOffline: " + isOffline()
            + "\nError: " + (e.message || JSON.stringify(e));
        console.error(errMsg, e);
        Xrm.Navigation.openAlertDialog({ text: errMsg });
    }
}

/**
 * Offline plugin logic — replicates server-side plugin behavior.
 */
async function runOfflineActionLogic_Finish(action, processExtensionId, workOrderId) {
    var step = "";
    try {
        var userId = Xrm.Utility.getGlobalContext().userSettings.userId.replace(/[{}]/g, "");

        var actionId = action.duc_stageactionid;
        var actionTypeId = action._duc_actiontype_value || null;
        var nextStageId = action._duc_nextstage_value || null;
        var statusId = action._duc_defaultstatus_value || null;
        var subStatusId = action._duc_defaultsubstatus_value || null;
        var relatedStageId = action._duc_relatedstage_value || null;

        // 1. Get action type info (sendtocustomer flag)
        var sendToCustomer = false;
        if (actionTypeId) {
            step = "Retrieving action type";
            try {
                var actionType = await Xrm.WebApi.retrieveRecord(
                    "duc_actiontype", actionTypeId, "?$select=duc_sendtocustomer"
                );
                sendToCustomer = actionType.duc_sendtocustomer || false;
            } catch (e) { /* non-critical */ }
        }

        // 2. Get process definition via related stage
        var processDefId = null;
        var targetStatusField = null, targetSubStatusField = null;
        var existingStatusEntity = null, existingSubStatusEntity = null;
        var statusLookupType = null, subStatusLookupType = null;

        if (relatedStageId) {
            step = "Retrieving related stage";
            try {
                var rs = await Xrm.WebApi.retrieveRecord(
                    "duc_processstage", relatedStageId, "?$select=_duc_relatedprocess_value"
                );
                processDefId = rs._duc_relatedprocess_value || null;
            } catch (e) { /* non-critical */ }
        }

        if (processDefId) {
            step = "Retrieving process definition";
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
            } catch (e) { /* non-critical */ }
        }

        // 3. Resolve status/substatus values — always fresh from API, no caching
        var statusValue = null, subStatusValue = null;

        if (statusId) {
            step = "Retrieving process status value";
            try {
                var sRec = await Xrm.WebApi.retrieveRecord("duc_processstatus", statusId, "?$select=duc_value");
                statusValue = sRec.duc_value || null;
            } catch (e) { /* non-critical */ }
        }

        if (subStatusId) {
            step = "Retrieving process substatus value";
            try {
                subStatusValue = (await Xrm.WebApi.retrieveRecord(
                    "duc_processsubstatus", subStatusId, "?$select=duc_value"
                )).duc_value || null;
            } catch (e) { /* non-critical */ }
        }

        // 4. Create action log
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

        // 5. Update process extension fields
        step = "Updating process extension fields";
        var peUpdate = { "duc_islastactiontakenoffline": true };
        if (nextStageId) peUpdate["duc_CurrentStage_duc_ProcessExtension@odata.bind"] = "/duc_processstages(" + nextStageId + ")";
        peUpdate["duc_LastApprovedById_duc_ProcessExtension_systemuser@odata.bind"] = "/systemusers(" + userId + ")";
        if (statusId) peUpdate["duc_Status_duc_ProcessExtension@odata.bind"] = "/duc_processstatuses(" + statusId + ")";
        if (subStatusId) peUpdate["duc_SubStatus_duc_ProcessExtension@odata.bind"] = "/duc_processsubstatuses(" + subStatusId + ")";

        await Xrm.WebApi.updateRecord("duc_processextension", processExtensionId, peUpdate);

        // 6. Update work order status/substatus
        step = "Updating work order status";
        var woUpdate = {};
        var hasWoUpdate = false;

        if (targetStatusField && statusValue) {
            if (statusLookupType === 780500002) {
                woUpdate[targetStatusField] = parseInt(statusValue, 10);
            } else if (existingStatusEntity) {
                woUpdate[targetStatusField + "@odata.bind"] = "/" + _entitySetNameFinish(existingStatusEntity) + "(" + statusValue + ")";
            }
            hasWoUpdate = true;
        }
        if (targetSubStatusField && subStatusValue) {
            if (subStatusLookupType === 780500002) {
                woUpdate[targetSubStatusField] = parseInt(subStatusValue, 10);
            } else if (existingSubStatusEntity) {
                woUpdate[targetSubStatusField + "@odata.bind"] = "/" + _entitySetNameFinish(existingSubStatusEntity) + "(" + subStatusValue + ")";
            }
            hasWoUpdate = true;
        }

        if (hasWoUpdate) {
            await Xrm.WebApi.updateRecord("msdyn_workorder", workOrderId, woUpdate);
        }

    } catch (e) {
        var errMsg = "[runOfflineActionLogic_Finish] Error at step: " + step
            + "\nActionId: " + (action ? action.duc_stageactionid : "N/A")
            + "\nPE Id: " + processExtensionId
            + "\nOffline: " + isOffline()
            + "\nError: " + (e.message || JSON.stringify(e));
        console.error(errMsg, e);
        Xrm.Navigation.openAlertDialog({ text: errMsg });
    }
}

function _entitySetNameFinish(logicalName) {
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
    var step = "";
    try {
        var form = Xrm.Page;
        var id = form.data.entity.getId();
        if (!id) return 0;

        id = id.replace(/[{}]/g, "");
        var isOff = isOffline();
        var woFilterField = isOff ? "msdyn_workorder" : "_msdyn_workorder_value";

        step = "Retrieving service tasks for work order";
        var tasks = await Xrm.WebApi.retrieveMultipleRecords(
            "msdyn_workorderservicetask",
            "?$select=msdyn_workorderservicetaskid&$filter=" + woFilterField + " eq " + id
        );

        step = "Updating service tasks to 100%";
        for (var ti = 0; ti < tasks.entities.length; ti++) {
            await Xrm.WebApi.updateRecord(
                "msdyn_workorderservicetask",
                tasks.entities[ti].msdyn_workorderservicetaskid,
                { msdyn_percentcomplete: 100 }
            );
        }

        return 1;
    }
    catch (e) {
        Xrm.Utility.closeProgressIndicator();
        var errMsg = "[updateServiceTasksPercent] Error at step: " + step
            + "\nOffline: " + isOffline()
            + "\nError: " + (e.message || JSON.stringify(e));
        console.error(errMsg, e);
        Xrm.Navigation.openAlertDialog({ text: errMsg });
        return 0;
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
        var now = new Date().toISOString();

        await Xrm.WebApi.updateRecord(
            form.data.entity.getEntityName(),
            id,
            {
                duc_workorderendtime: now,
                duc_completiondate: now,
                duc_inspectioncompletiondate: now
            }
        );
    }
    catch (e) {
        var errMsg = "[setWorkOrderEndTime] Error."
            + "\nOffline: " + isOffline()
            + "\nError: " + (e.message || JSON.stringify(e));
        console.error(errMsg, e);
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
    if (!valid) {
        Xrm.Utility.closeProgressIndicator();
        return;
    }

    var validatePercentage = await updateServiceTasksPercent();
    if (!validatePercentage) {
        Xrm.Utility.closeProgressIndicator();
        return;
    }

    var booking = await setBookingCompletedForWorkOrder();
    if (!booking) {
        Xrm.Utility.closeProgressIndicator();
        return;
    }

    // Resolve process extension ID and work order ID once, then pass to each function
    var form = Xrm.Page;
    var peLookup = form.getAttribute("duc_processextension");
    if (peLookup && peLookup.getValue()) {
        var extId = peLookup.getValue()[0].id.replace(/[{}]/g, "");
        var workOrderId = form.data.entity.getId().replace(/[{}]/g, "");

        // Step 1: update stage — writes new current stage to DB
        await updateCurrentStageOnProcessExtension(extId);

        // Step 2: fetch fresh context AFTER stage update so it sees the new stage
        await updateLastActionOnProcessExtension(extId, workOrderId);
    }

    await setWorkOrderEndTime();
    await navigateToWorkORderTab();

    setTimeout(function () {
        try { Xrm.Page.data.refresh(true); } catch (e) { }
        Xrm.Utility.closeProgressIndicator();
    }, 3000);
}

runProcess();