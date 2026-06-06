// Localization dictionary
const labels = {
    processing: {
        1033: "Processing...", // English
        1025: "جارٍ المعالجة..." // Arabic
    },
    processingWorkOrder: {
        1033: "Processing Work Order...",
        1025: "جارٍ معالجة أمر العمل..."
    },
    mandatoryFields: {
        1033: "Please fill the mandatory fields:\n",
        1025: "يرجى تعبئة الحقول الإلزامية التالية:\n"
    },
    complianceCalculationFailed: {
        1033: "Failed to calculate compliance levels.",
        1025: "فشل في حساب مستويات الالتزام."
    },
    unexpectedError: {
        1033: "Unexpected error: ",
        1025: "حدث خطأ غير متوقع: "
    },
    noWorkOrderFound: {
        1033: "No related Work Order found.",
        1025: "لم يتم العثور على أمر العمل المرتبط."
    },
    bookingStatusNotFound: {
        1033: "Booking Status 'In Progress' not found.",
        1025: "حالة الحجز 'قيد التنفيذ' غير موجودة."
    }
};

// Helper function to get localized string
function getLocalizedString(key) {
    const langId = Xrm.Utility.getGlobalContext().userSettings.languageId;
    return labels[key] && labels[key][langId] ? labels[key][langId] : labels[key][1033];
}

function showLoader(messageKey) {
    Xrm.Utility.showProgressIndicator(getLocalizedString(messageKey) || getLocalizedString("processing"));
}

function hideLoader() {
    Xrm.Utility.closeProgressIndicator();
}

async function openWorkOrderAndFocusTab() {
    var formContext = Xrm.Page;
    if (!validateRequiredFields(formContext)) return;
    showLoader("processingWorkOrder");
    await calculateComplianceLevel(formContext);
    await runAsyncOpenWorkOrder(formContext);
}

async function calculateComplianceLevel(formContext) {
    try {
        var recordId = formContext.data.entity.getId();
        if (!recordId) return;
        var woId = recordId.replace(/[{}]/g, "");
        var counts = { 1: 0, 2: 0, 3: 0, 4: 0 };

        const result = await Xrm.WebApi.retrieveMultipleRecords("duc_questionanswersconfiguration",
            `?$select=duc_sequencefromcompliancelevel&$filter=_duc_msdyn_workorderservicetask_value eq ${woId}`);
        result.entities.forEach(e => {
            var level = e.duc_sequencefromcompliancelevel;
            if (counts.hasOwnProperty(level) && level) counts[level]++;
        });

        await Xrm.WebApi.updateRecord("msdyn_workorderservicetask", recordId, {
            duc_compliancelevel1count: counts[1],
            duc_compliancelevel2count: counts[2],
            duc_compliancelevel3count: counts[3],
            duc_compliancelevel4count: counts[4]
        });
    } catch (error) {
        console.error("Compliance calculation failed:", error.message);
        hideLoader();
        Xrm.Utility.alertDialog(getLocalizedString("complianceCalculationFailed"));
    }
}

function validateRequiredFields(formContext) {
    var missingFields = [];
    formContext.ui.controls.forEach(function (control) {
        if (typeof control.getAttribute !== "function") return;
        var attribute = control.getAttribute();
        if (!attribute) return;
        if (attribute.getRequiredLevel() === "required") {
            var value = attribute.getValue();
            if (value === null || value === "" || (Array.isArray(value) && value.length === 0)) {
                missingFields.push(control.getLabel());
            }
        }
    });

    if (missingFields.length > 0) {
        var msg = getLocalizedString("mandatoryFields") + missingFields.join(", ") + "\n\n" +
            getLocalizedString("mandatoryFields") + missingFields.join("، ");
        Xrm.Navigation.openAlertDialog({ text: msg });
        return false;
    }
    return true;
}
function waitForSaveToReflect() {
    return new Promise(resolve => {
        setTimeout(resolve, 2000);
    });
}
async function runAsyncOpenWorkOrder(formContext) {
    debugger;
    try {
        showLoader("processingWorkOrder");
        //  await formContext.data.save();
        await waitForSaveToReflect();
        if (formContext.data.entity.getIsDirty()) {
            await formContext.data.save();
            await waitForSaveToReflect();
        }

        // Only when question = "مخالفة"
        if (isViolationSelected(formContext)) {

            const complianceResult = await hasValidComplianceStatus(formContext);

            // Block ONLY when penalties exist AND invalid
            if (complianceResult === false) {
                hideLoader();

                await Xrm.Navigation.openAlertDialog({ text: "يجب تسجيل مخالفة واحدة على الأقل." }).then(
                    async function (success) {
                        await Xrm.Navigation.openForm({
                            entityName: formContext.data.entity.getEntityName(),
                            entityId: formContext.data.entity.getId()
                        });
                    },
                );
                return;
            }

        }

        hideLoader();
        await redirectToWorkOrder(formContext);

    } catch (e) {
        hideLoader();
        Xrm.Navigation.openAlertDialog({
            text: getLocalizedString("unexpectedError") + e.message
        });
    }
}
async function redirectToWorkOrder(formContext) {
    var woAttr = formContext.getAttribute("msdyn_workorder");
    if (!woAttr || !woAttr.getValue()) {
        return Xrm.Navigation.openAlertDialog({ text: getLocalizedString("noWorkOrderFound") });
    }
    var woRef = woAttr.getValue()[0];
    var woId = woRef.id.replace(/[{}]/g, "");
    var woEntity = woRef.entityType;

    return Xrm.Navigation.openForm({
        entityName: woEntity,
        entityId: woId,
        pageType: "entityrecord"
    }, {
        focusTab: "ResponsibleEmployee"
    });
}

async function sBIP(formContext) {
    try {
        var woAttr = formContext.getAttribute("msdyn_workorder");
        if (!woAttr || !woAttr.getValue()) return;

        var woId = woAttr.getValue()[0].id.replace(/[{}]/g, "");
        var bookingResult = await Xrm.WebApi.retrieveMultipleRecords(
            "bookableresourcebooking",
            `?$select=bookableresourcebookingid&$filter=_msdyn_workorder_value eq ${woId}&$orderby=createdon asc&$top=1`
        );

        if (!bookingResult.entities.length) return;

        var bookingId = bookingResult.entities[0].bookableresourcebookingid;
        var statusResult = await Xrm.WebApi.retrieveMultipleRecords(
            "bookingstatus",
            `?$select=bookingstatusid&$filter=name eq 'In Progress'`
        );

        if (!statusResult.entities.length) {
            return Xrm.Navigation.openAlertDialog({ text: getLocalizedString("bookingStatusNotFound") });
        }

        var statusId = statusResult.entities[0].bookingstatusid;
        await Xrm.WebApi.updateRecord("bookableresourcebooking", bookingId, {
            "BookingStatus@odata.bind": `/bookingstatuses(${statusId})`
        });
    } catch (e) {
        Xrm.Navigation.openAlertDialog({ text: e.message });
    }
}


function isViolationSelected(formContext) {
    const attr = formContext.getAttribute("duc_question1");
    if (!attr) return false;

    const value = attr.getValue();
    if (!value) return false;

    const normalizedValue = value.trim();

    return normalizedValue === "مخالفة" ||
        normalizedValue === "غير مستوف الشروط";
}


async function hasValidComplianceStatus(formContext) {
    try {
        const taskId = formContext.data.entity.getId();
        if (!taskId) return null;

        const cleanId = taskId.replace(/[{}]/g, "");

        const result = await Xrm.WebApi.retrieveMultipleRecords(
            "duc_workorderpenalties",
            `?$select=duc_compliancestatus&$filter=_duc_workorderservicetask_value eq ${cleanId}`
        );

        // NO penalties → do NOT block
        if (!result.entities.length) {
            return null;
        }

        // penalties exist → check compliance
        return result.entities.some(r => r.duc_compliancestatus !== null);

    } catch (e) {
        console.error("Penalty compliance check failed", e);
        return null; // fail-safe: do not block
    }
}

// Call the main function
openWorkOrderAndFocusTab();