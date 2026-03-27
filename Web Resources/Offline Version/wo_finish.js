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

function isOffline() {
    try {
        if (Xrm.Utility.getGlobalContext().client.isOffline()) return true;
        if (Xrm.Utility.getGlobalContext().client.getClientState() === "Offline") return true;
    } catch (e) { }
    return false;
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

        await Xrm.Navigation.openAlertDialog({ text: "[validateRequiredFields] SUCCESS" });
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

        // Update all bookings to Completed status
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

        await Xrm.Navigation.openAlertDialog({ text: "[setBookingCompletedForWorkOrder] SUCCESS" });
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
 ***************************************/
async function updateCurrentStageOnProcessExtension() {
    var step = "";
    try {
        var form = Xrm.Page;
        var lookup = form.getAttribute("duc_processextension");

        if (!lookup || !lookup.getValue()) return;

        var extId = lookup.getValue()[0].id.replace(/[{}]/g, "");
        var isOff = isOffline();

        step = "Retrieving process extension record";
        var record = await Xrm.WebApi.retrieveRecord(
            "duc_processextension",
            extId,
            "?$select=_duc_processdefinition_value,_duc_currentstage_value"
        );

        if (!record._duc_processdefinition_value || !record._duc_currentstage_value) return;

        step = "Retrieving stage actions for current stage";
        var processField = isOff ? "duc_process" : "_duc_process_value";
        var relatedStageField = isOff ? "duc_relatedstage" : "_duc_relatedstage_value";

        var actions = await Xrm.WebApi.retrieveMultipleRecords(
            "duc_stageaction",
            "?$select=duc_stageactionid,_duc_nextstage_value,_duc_defaultstatus_value" +
            "&$filter=duc_canbetriggeredbytarget eq true" +
            " and " + processField + " eq " + record._duc_processdefinition_value +
            " and " + relatedStageField + " eq " + record._duc_currentstage_value
        );

        if (!actions.entities || actions.entities.length === 0) {
            console.warn("[updateCurrentStageOnProcessExtension] No matching stage action found.");
            return;
        }

        step = "Looping stage actions to find status 690970004";
        var nextStageId = null;

        for (var i = 0; i < actions.entities.length; i++) {
            var act = actions.entities[i];
            if (!act._duc_defaultstatus_value) continue;

            var actStatus = await Xrm.WebApi.retrieveRecord(
                "duc_processstatus",
                act._duc_defaultstatus_value,
                "?$select=duc_value"
            );

            if (actStatus.duc_value === 690970004) {
                nextStageId = act._duc_nextstage_value;
                break;
            }
        }

        if (!nextStageId) {
            console.warn("[updateCurrentStageOnProcessExtension] No action with status 690970004 found, or action has no next stage.");
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
 * PROCESS EXTENSION UPDATE
 ***************************************/
async function updateLastActionOnProcessExtension() {
    debugger
    var step = "";
    try {
        var form = Xrm.Page;
        var lookup = form.getAttribute("duc_processextension");

        if (!lookup || !lookup.getValue()) return;

        var extId = lookup.getValue()[0].id.replace(/[{}]/g, "");
        alert("extId: " + extId);
        var isOff = isOffline();

        step = "Retrieving process extension record";
        alert(step);
        var record = await Xrm.WebApi.retrieveRecord(
            "duc_processextension",
            extId,
            "?$select=_duc_processdefinition_value,_duc_currentstage_value"
        );
        alert("record: " + JSON.stringify(record));

        if (!record._duc_processdefinition_value) return;

        step = "Retrieving stage actions";
        alert(step);
        var processField = isOff ? "duc_process" : "_duc_process_value";
        var relatedStageField = isOff ? "duc_relatedstage" : "_duc_relatedstage_value";

        var actions = await Xrm.WebApi.retrieveMultipleRecords(
            "duc_stageaction",
            "?$select=duc_stageactionid,_duc_defaultstatus_value" +
            "&$filter=duc_canbetriggeredbytarget eq true" +
            " and " + processField + " eq " + record._duc_processdefinition_value +
            " and " + relatedStageField + " eq " + record._duc_currentstage_value
        );
        alert("actions count: " + actions.entities.length);

        step = "Looping stage actions to find matching status";
        alert(step);
        for (var ai = 0; ai < actions.entities.length; ai++) {
            var act = actions.entities[ai];
            if (!act._duc_defaultstatus_value) continue;

            var actStatus = await Xrm.WebApi.retrieveRecord(
                "duc_processstatus",
                act._duc_defaultstatus_value,
                "?$select=duc_value"
            );
            alert("actStatus value: " + actStatus.duc_value);

            if (actStatus.duc_value === 100000006) {
                step = "Updating process extension with LastActionTaken";
                alert(step);
                await Xrm.WebApi.updateRecord(
                    "duc_processextension",
                    extId,
                    {
                        "duc_LastActionTaken_duc_ProcessExtension@odata.bind":
                            "/duc_stageactions(" + act.duc_stageactionid + ")"
                    }
                );

                // Run offline plugin logic inline (no external dependency)
                step = "Running offline action logic";
                alert(step);
                await runOfflineActionLogic_Finish(act.duc_stageactionid, extId, form.data.entity.getId().replace(/[{}]/g, ""));

                break;
            }
        }
        await Xrm.Navigation.openAlertDialog({ text: "[updateLastActionOnProcessExtension] SUCCESS" });
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

// Self-contained offline plugin logic for wo_finish.js
async function runOfflineActionLogic_Finish(actionId, processExtensionId, workOrderId) {
    var step = "";
    try {
        var userId = Xrm.Utility.getGlobalContext().userSettings.userId.replace(/[{}]/g, "");
        Xrm.Navigation.openAlertDialog({ text: "[runOfflineActionLogic_Finish] START\nactionId: " + actionId + "\nPE: " + processExtensionId + "\nWO: " + workOrderId + "\nuserId: " + userId });

        step = "Retrieving stage action details";
        var action = await Xrm.WebApi.retrieveRecord("duc_stageaction", actionId,
            "?$select=duc_name,_duc_actiontype_value,_duc_nextstage_value,_duc_relatedstage_value," +
            "_duc_defaultstatus_value,_duc_defaultsubstatus_value"
        );
        Xrm.Navigation.openAlertDialog({ text: "[step 1 OK] action: " + JSON.stringify(action) });

        var actionTypeId = action._duc_actiontype_value || null;
        var nextStageId = action._duc_nextstage_value || null;
        var statusId = action._duc_defaultstatus_value || null;
        var subStatusId = action._duc_defaultsubstatus_value || null;
        var relatedStageId = action._duc_relatedstage_value || null;

        var sendToCustomer = false;
        if (actionTypeId) {
            step = "Retrieving action type";
            try {
                var actionType = await Xrm.WebApi.retrieveRecord("duc_actiontype", actionTypeId, "?$select=duc_sendtocustomer");
                sendToCustomer = actionType.duc_sendtocustomer || false;
                Xrm.Navigation.openAlertDialog({ text: "[step 2 OK] actionType sendToCustomer: " + sendToCustomer });
            } catch (e) {
                Xrm.Navigation.openAlertDialog({ text: "[step 2 WARN] Failed to get action type: " + (e.message || e) });
            }
        }

        var processDefId = null;
        var targetStatusField = null, targetSubStatusField = null;
        var existingStatusEntity = null, existingSubStatusEntity = null;
        var statusLookupType = null, subStatusLookupType = null;

        if (relatedStageId) {
            step = "Retrieving related stage";
            try {
                var rs = await Xrm.WebApi.retrieveRecord("duc_processstage", relatedStageId, "?$select=_duc_relatedprocess_value");
                processDefId = rs._duc_relatedprocess_value || null;
                Xrm.Navigation.openAlertDialog({ text: "[step 3 OK] processDefId: " + processDefId });
            } catch (e) {
                Xrm.Navigation.openAlertDialog({ text: "[step 3 WARN] Failed to get related stage: " + (e.message || e) });
            }
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
                Xrm.Navigation.openAlertDialog({ text: "[step 4 OK] procDef: targetStatus=" + targetStatusField + " targetSubStatus=" + targetSubStatusField });
            } catch (e) {
                Xrm.Navigation.openAlertDialog({ text: "[step 4 WARN] Failed to get process def: " + (e.message || e) });
            }
        }

        var statusValue = null, subStatusValue = null;
        if (statusId) {
            step = "Retrieving process status value";
            try {
                statusValue = (await Xrm.WebApi.retrieveRecord("duc_processstatus", statusId, "?$select=duc_value")).duc_value || null;
                Xrm.Navigation.openAlertDialog({ text: "[step 5a OK] statusValue: " + statusValue });
            } catch (e) {
                Xrm.Navigation.openAlertDialog({ text: "[step 5a WARN] Failed to get status: " + (e.message || e) });
            }
        }
        if (subStatusId) {
            step = "Retrieving process substatus value";
            try {
                subStatusValue = (await Xrm.WebApi.retrieveRecord("duc_processsubstatus", subStatusId, "?$select=duc_value")).duc_value || null;
                Xrm.Navigation.openAlertDialog({ text: "[step 5b OK] subStatusValue: " + subStatusValue });
            } catch (e) {
                Xrm.Navigation.openAlertDialog({ text: "[step 5b WARN] Failed to get substatus: " + (e.message || e) });
            }
        }

        step = "Creating action log";
        var logData = {
            "subject": "Process Action", "actualstart": new Date(),
            "duc_sendtocustomer": sendToCustomer, "duc_islastactiontakenoffline": true
        };
        logData["ownerid_duc_processactionlog@odata.bind"] = "/systemusers(" + userId + ")";
        logData["duc_AuthorId_duc_processActionLog_systemuser@odata.bind"] = "/systemusers(" + userId + ")";
        logData["regardingobjectid_msdyn_workorder_duc_processactionlog@odata.bind"] = "/msdyn_workorders(" + workOrderId + ")";
        logData["duc_Action_duc_processActionLog@odata.bind"] = "/duc_stageactions(" + actionId + ")";
        if (actionTypeId) logData["duc_ActionType_duc_processActionLog@odata.bind"] = "/duc_actiontypes(" + actionTypeId + ")";
        if (relatedStageId) logData["duc_processStage_duc_processActionLog@odata.bind"] = "/duc_processstages(" + relatedStageId + ")";
        if (processDefId) logData["duc_process_duc_processActionLog@odata.bind"] = "/duc_processdefinitions(" + processDefId + ")";
        await Xrm.WebApi.createRecord("duc_processactionlog", logData);
        Xrm.Navigation.openAlertDialog({ text: "[step 6 OK] Action log created" });

        step = "Updating process extension fields";
        var peUpdate = { "duc_islastactiontakenoffline": true };
        if (nextStageId) peUpdate["duc_CurrentStage_duc_ProcessExtension@odata.bind"] = "/duc_processstages(" + nextStageId + ")";
        peUpdate["duc_LastApprovedById_duc_ProcessExtension_systemuser@odata.bind"] = "/systemusers(" + userId + ")";
        if (statusId) peUpdate["duc_Status_duc_ProcessExtension@odata.bind"] = "/duc_processstatuses(" + statusId + ")";
        if (subStatusId) peUpdate["duc_SubStatus_duc_ProcessExtension@odata.bind"] = "/duc_processsubstatuses(" + subStatusId + ")";
        await Xrm.WebApi.updateRecord("duc_processextension", processExtensionId, peUpdate);
        Xrm.Navigation.openAlertDialog({ text: "[step 7 OK] Process extension updated" });

        step = "Updating work order status";
        var woUpdate = {};
        var hasWoUpdate = false;
        if (targetStatusField && statusValue) {
            if (statusLookupType === 780500002) { woUpdate[targetStatusField] = parseInt(statusValue, 10); }
            else if (existingStatusEntity) { woUpdate[targetStatusField + "@odata.bind"] = "/" + _entitySetNameFinish(existingStatusEntity) + "(" + statusValue + ")"; }
            hasWoUpdate = true;
        }
        if (targetSubStatusField && subStatusValue) {
            if (subStatusLookupType === 780500002) { woUpdate[targetSubStatusField] = parseInt(subStatusValue, 10); }
            else if (existingSubStatusEntity) { woUpdate[targetSubStatusField + "@odata.bind"] = "/" + _entitySetNameFinish(existingSubStatusEntity) + "(" + subStatusValue + ")"; }
            hasWoUpdate = true;
        }
        if (hasWoUpdate) {
            await Xrm.WebApi.updateRecord("msdyn_workorder", workOrderId, woUpdate);
            Xrm.Navigation.openAlertDialog({ text: "[step 8 OK] Work order status/substatus updated\n" + JSON.stringify(woUpdate) });
        } else {
            Xrm.Navigation.openAlertDialog({ text: "[step 8 SKIP] No WO status update needed (targetStatusField=" + targetStatusField + " statusValue=" + statusValue + ")" });
        }

        Xrm.Navigation.openAlertDialog({ text: "[runOfflineActionLogic_Finish] DONE successfully" });

    } catch (e) {
        var errMsg = "[runOfflineActionLogic_Finish] Error at step: " + step
            + "\nActionId: " + actionId + "\nPE Id: " + processExtensionId
            + "\nOffline: " + isOffline() + "\nError: " + (e.message || JSON.stringify(e));
        console.error(errMsg, e);
        Xrm.Navigation.openAlertDialog({ text: errMsg });
    }
}

function _entitySetNameFinish(logicalName) {
    var map = {
        "msdyn_workordersubstatus": "msdyn_workordersubstatuses", "msdyn_workorderstatus": "msdyn_workorderstatuses",
        "duc_processstatus": "duc_processstatuses", "duc_processsubstatus": "duc_processsubstatuses",
        "duc_processstage": "duc_processstages", "msdyn_workorder": "msdyn_workorders"
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

        step = "Retrieving work order subaccount";
        var wo = await Xrm.WebApi.retrieveRecord(
            "msdyn_workorder",
            id,
            "?$select=_duc_subaccount_value"
        );

        var accountId = wo._duc_subaccount_value;
        alert(accountId);
        step = "Retrieving service tasks for work order";
        alert(step);
        var woFilterField = isOff ? "msdyn_workorder" : "_msdyn_workorder_value";

        var tasks = await Xrm.WebApi.retrieveMultipleRecords(
            "msdyn_workorderservicetask",
            "?$select=msdyn_workorderservicetaskid&$filter=" + woFilterField + " eq " + id
        );

        step = "Updating service tasks to 100%";
        alert(step);

        for (var ti = 0; ti < tasks.entities.length; ti++) {
            await Xrm.WebApi.updateRecord(
                "msdyn_workorderservicetask",
                tasks.entities[ti].msdyn_workorderservicetaskid,
                { msdyn_percentcomplete: 100 }
            );
        }
        alert(step);
        await Xrm.Navigation.openAlertDialog({ text: "[updateServiceTasksPercent] SUCCESS" });
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

        await Xrm.WebApi.updateRecord(
            form.data.entity.getEntityName(),
            id,
            {
                duc_workorderendtime: new Date().toISOString(),
                duc_completiondate: new Date().toISOString(),
                duc_inspectioncompletiondate: new Date().toISOString()
            }
        );
        await Xrm.Navigation.openAlertDialog({ text: "[setWorkOrderEndTime] SUCCESS" });
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
    var btns = document.querySelectorAll(`button[data-id*='${refreshbtnId}']`);
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
    debugger
    await Xrm.Navigation.openAlertDialog({ text: "[runProcess] START" });
    Xrm.Utility.showProgressIndicator(t("Processing"));

    var valid = await validateRequiredFields();
    if (!valid) {
        alert("Please make sure to fill all required fields in the inspection template, يرجى التأكد من تعبئة جميع الحقول الإلزامية في نموذج التفتيش.");
        Xrm.Utility.closeProgressIndicator();
        return;
    }

    var validatePercentage = await updateServiceTasksPercent();
    if (!validatePercentage) {
        alert("Please make sure to fill all required fields in the inspection template, يرجى التأكد من تعبئة جميع الحقول الإلزامية في نموذج التفتيش.");
        Xrm.Utility.closeProgressIndicator();
        return;
    }
    debugger

    var booking = await setBookingCompletedForWorkOrder();
    if (!booking) {
        await Xrm.Navigation.openAlertDialog({ text: "[runProcess] FAILED at setBookingCompletedForWorkOrder" });
        return;
    }

    await updateCurrentStageOnProcessExtension();
    await updateLastActionOnProcessExtension();
    await setWorkOrderEndTime();
    await navigateToWorkORderTab();

    setTimeout(async function () {
        try { Xrm.Page.data.refresh(true); } catch (e) { }
        Xrm.Utility.closeProgressIndicator();
        await Xrm.Navigation.openAlertDialog({ text: "[runProcess] DONE / SUCCESS" });
    }, 3000);
}

runProcess();