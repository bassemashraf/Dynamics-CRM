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
            `?$select=bookableresourcebookingid,starttime
              &$filter=${woFilterField} eq ${workOrderId}
              &$orderby=createdon asc`
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
        var errMsg = "[setBookingCompletedForWorkOrder] Error at step: " + step
            + "\nOffline: " + isOffline()
            + "\nError: " + (e.message || JSON.stringify(e));
        console.error(errMsg, e);
        Xrm.Navigation.openAlertDialog({ text: errMsg });
        return false;
    }
}

/***************************************
 * PROCESS EXTENSION UPDATE
 ***************************************/
async function updateLastActionOnProcessExtension() {
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

        if (!record._duc_processdefinition_value) return;

        step = "Retrieving stage actions";
        var processField = isOff ? "duc_process" : "_duc_process_value";
        var relatedStageField = isOff ? "duc_relatedstage" : "_duc_relatedstage_value";

        var actions = await Xrm.WebApi.retrieveMultipleRecords(
            "duc_stageaction",
            `?$select=duc_stageactionid,_duc_defaultstatus_value
              &$filter=duc_canbetriggeredbytarget eq true
              and ${processField} eq ${record._duc_processdefinition_value}
              and ${relatedStageField} eq ${record._duc_currentstage_value}`
        );

        step = "Looping stage actions to find matching status";
        for (let act of actions.entities) {
            if (!act._duc_defaultstatus_value) continue;

            let status = await Xrm.WebApi.retrieveRecord(
                "duc_processstatus",
                act._duc_defaultstatus_value,
                "?$select=duc_value"
            );

            if (status.duc_value === 100000006) {
                step = "Updating process extension with LastActionTaken";
                await Xrm.WebApi.updateRecord(
                    "duc_processextension",
                    extId,
                    {
                        "duc_LastActionTaken_duc_ProcessExtension@odata.bind":
                            `/duc_stageactions(${act.duc_stageactionid})`
                    }
                );
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

        step = "Retrieving service tasks for work order";
        var woFilterField = isOff ? "msdyn_workorder" : "_msdyn_workorder_value";

        var tasks = await Xrm.WebApi.retrieveMultipleRecords(
            "msdyn_workorderservicetask",
            `?$select=msdyn_workorderservicetaskid
              &$filter=${woFilterField} eq ${id}`
        );

        step = "Updating service tasks to 100%";
        for (let task of tasks.entities) {

            let updateObj = {
                msdyn_percentcomplete: 100
            };


            await Xrm.WebApi.updateRecord(
                "msdyn_workorderservicetask",
                task.msdyn_workorderservicetaskid,
                updateObj
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

        await Xrm.WebApi.updateRecord(
            form.data.entity.getEntityName(),
            id,
            {
                duc_workorderendtime: new Date().toISOString(),
                duc_completiondate: new Date().toISOString(),
                duc_inspectioncompletiondate: new Date().toISOString()
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

    var booking = await setBookingCompletedForWorkOrder();
    if (!booking) return;

    await updateLastActionOnProcessExtension();
    await setWorkOrderEndTime();
    await navigateToWorkORderTab();

    setTimeout(function () {
        try { Xrm.Page.data.refresh(true); } catch (e) { }
        Xrm.Utility.closeProgressIndicator();
    }, 3000);
}

runProcess();