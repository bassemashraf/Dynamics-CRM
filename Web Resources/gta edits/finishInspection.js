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
            `?$select=duc_stageactionid,_duc_defaultstatus_value
              &$filter=duc_canbetriggeredbytarget eq true
              and _duc_process_value eq ${record._duc_processdefinition_value}
              and _duc_relatedstage_value eq ${record._duc_currentstage_value}`
        );

        for (let act of actions.entities) {
            if (!act._duc_defaultstatus_value) continue;

            let status = await Xrm.WebApi.retrieveRecord(
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
                            `/duc_stageactions(${act.duc_stageactionid})`
                    }
                );
                break;
            }
        }
    }
    catch (e) {
        Xrm.Utility.closeProgressIndicator();
        console.error(e);
    }
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
            `?$select=msdyn_workorderservicetaskid
              &$filter=_msdyn_workorder_value eq ${id}`
        );

        for (let task of tasks.entities) {
            await Xrm.WebApi.updateRecord(
                "msdyn_workorderservicetask",
                task.msdyn_workorderservicetaskid,
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