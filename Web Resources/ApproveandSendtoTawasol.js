// ============================================================================
// CREATE TAWASOL REQUEST RECORD (Correct Binding Syntax)
// ============================================================================
async function createTawasolRequest() {
    try {

        debugger;

        var form = Xrm.Page;

        // ---------------------------
        // 1. Read values
        // ---------------------------
        var department = form.getAttribute("duc_department")?.getValue();

        var urlParams = new URLSearchParams(window.location.search);
        var currentEntity = urlParams.get("etn");
        var relatedInspectionAction = null;
        var relatedWorkOrder = null;

        //handle if the action from work order entity or from inspection action
        if (currentEntity == "msdyn_workorder") {
            relatedInspectionAction = form.getAttribute("duc_primaryinspectionaction")?.getValue();
            relatedWorkOrder = urlParams.get("id");


        }
        else if (currentEntity == "duc_inspectionaction") {
            relatedInspectionAction = urlParams.get("id");
            relatedWorkOrder = form.getAttribute("duc_relatedworkorderid")?.getValue();


        }


        if (!department || !relatedWorkOrder) {
            await Xrm.Navigation.openAlertDialog({
                text:
                    "Missing required data (Department or Work Order).\n" +
                    "البيانات المطلوبة غير مكتملة (القسم أو أمر العمل)."
            });
            return false;
        }

        var departmentId = Array.isArray(department)
            ? department[0]?.id?.replace(/[{}]/g, "")
            : department;

        var workOrderId = Array.isArray(relatedWorkOrder)
            ? relatedWorkOrder[0]?.id?.replace(/[{}]/g, "")
            : relatedWorkOrder;


        var relatedInspectionActionId = Array.isArray(relatedInspectionAction)
            ? relatedInspectionAction[0]?.id?.replace(/[{}]/g, "")
            : relatedInspectionAction;

        if (!relatedInspectionAction) {
            await Xrm.Navigation.openAlertDialog({
                text:
                    "Inspection Action ID is missing from URL.\n" +
                    "معرّف إجراء التفتيش غير موجود في الرابط."
            });
            return false;
        }


        // ---------------------------
        // 3. Prepare record using CORRECT binding
        // ---------------------------
        var record = {};
        record["duc_name"] = "Violation";

        // Correct lookup binding
        record["duc_OU@odata.bind"] = "/msdyn_organizationalunits(" + departmentId + ")";
        record["duc_WorkOrder@odata.bind"] = "/msdyn_workorders(" + workOrderId + ")";
        record["duc_InspectionAction@odata.bind"] = "/duc_inspectionactions(" + relatedInspectionActionId + ")";

        // ---------------------------
        // 4. Create record
        // ---------------------------
        var result = await Xrm.WebApi.createRecord("duc_tawasolrequest", record);

        // ---------------------------
        // 5. Success message
        // ---------------------------
        await Xrm.Navigation.openAlertDialog({
            text:
                "Tawasol Request has been created successfully.\n" +
                "تم إنشاء طلب التواصل بنجاح."
        });

        return true;
    }
    catch (e) {

        await Xrm.Navigation.openAlertDialog({
            text:
                "Error creating Tawasol Request: " + e.message + "\n" +
                "حدث خطأ أثناء إنشاء طلب التواصل."
        });

        return false;
    }
}
async function openTawasolRequestQuickCreate() {
    try {
        debugger;

        var form = Xrm.Page;

        // ---------------------------
        // 1. Get the related inspection action
        // ---------------------------
        var urlParams = new URLSearchParams(window.location.search);
        var currentEntity = urlParams.get("etn");
        var relatedInspectionAction = null;

        // Handle if the action from work order entity or from inspection action
        if (currentEntity == "msdyn_workorder") {
            relatedInspectionAction = form.getAttribute("duc_primaryinspectionaction")?.getValue();
        }
        else if (currentEntity == "duc_inspectionaction") {
            relatedInspectionAction = urlParams.get("id");
        }

        if (!relatedInspectionAction) {
            await Xrm.Navigation.openAlertDialog({
                text:
                    "Inspection Action ID is missing.\n" +
                    "معرّف إجراء التفتيش غير موجود."
            });
            return false;
        }

        // Extract the ID
        var relatedInspectionActionId = Array.isArray(relatedInspectionAction)
            ? relatedInspectionAction[0]?.id?.replace(/[{}]/g, "")
            : relatedInspectionAction?.replace(/[{}]/g, "");

        // Get the name for the lookup (if it's an array with lookup value)
        var relatedInspectionActionName = Array.isArray(relatedInspectionAction)
            ? relatedInspectionAction[0]?.name
            : "";

        // If we got the ID from URL, we might need to retrieve the name
        if (!relatedInspectionActionName && relatedInspectionActionId) {
            try {
                var inspectionActionRecord = await Xrm.WebApi.retrieveRecord(
                    "duc_inspectionaction",
                    relatedInspectionActionId,
                    "?$select=duc_name"
                );
                relatedInspectionActionName = inspectionActionRecord.duc_name || "";
            } catch (e) {
                console.error("Error retrieving inspection action name:", e.message);
            }
        }

        // ---------------------------
        // 2. Set up default values for quick create form
        // ---------------------------
        var defaultValues = {};

        // Set the inspection action lookup field
        defaultValues["duc_inspectionaction"] = [{
            id: relatedInspectionActionId,
            name: relatedInspectionActionName,
            entityType: "duc_inspectionaction"
        }];

        // ---------------------------
        // 3. Open quick create form
        // ---------------------------
        var entityFormOptions = {
            entityName: "duc_tawasolrequest",
            useQuickCreateForm: true
        };

        var result = await Xrm.Navigation.openForm(entityFormOptions, defaultValues);
        if (result && result.savedEntityReference && result.savedEntityReference.length > 0) {
            // Record was created successfully
            var createdRecordId = result.savedEntityReference[0].id;
            console.log("Tawasol Request created with ID: " + createdRecordId);

            // Now call the update function
            await updateLastActionOnProcessExtension();
        }

        return true;
    }
    catch (e) {
        await Xrm.Navigation.openAlertDialog({
            text:
                "Error opening Tawasol Request form: " + e.message + "\n" +
                "حدث خطأ أثناء فتح نموذج طلب التواصل."
        });

        return false;
    }
}

async function updateLastActionOnProcessExtension() {
    try {
        debugger;

        var urlParams = new URLSearchParams(window.location.search);
        var currentEntity = urlParams.get("etn");

        var processExtensionId = null;

        // 1. Get Process Extension ID
        if (currentEntity === "msdyn_workorder") {

            var ins = Xrm.Page.getAttribute("duc_primaryinspectionaction");

            if (ins && ins.getValue()) {
                var inspectionActionId = ins.getValue()[0].id.replace(/[{}]/g, "");

                var inspectionRecord = await Xrm.WebApi.retrieveRecord(
                    "duc_inspectionaction",
                    inspectionActionId,
                    "?$select=_duc_processextension_value"
                );

                processExtensionId = inspectionRecord._duc_processextension_value;
            }
        }
        else if (currentEntity === "duc_inspectionaction") {

            var lookup = Xrm.Page.getAttribute("duc_processextension");

            if (lookup && lookup.getValue()) {
                processExtensionId = lookup.getValue()[0].id.replace(/[{}]/g, "");
            }
        }

        if (!processExtensionId) {
            console.log("No Process Extension found.");
            return;
        }

        // 2. Get Process Definition and Current Stage
        var record = await Xrm.WebApi.retrieveRecord(
            "duc_processextension",
            processExtensionId,
            "?$select=_duc_processdefinition_value,_duc_currentstage_value"
        );

        var processDef = record._duc_processdefinition_value;
        var currentStage = record._duc_currentstage_value;

        if (!processDef) return;

        // 3. Get Stage Actions
        var query =
            "?$select=duc_stageactionid,_duc_defaultstatus_value" +
            "&$filter=duc_canbetriggeredbytarget eq true" +
            " and _duc_process_value eq " + processDef +
            " and _duc_relatedstage_value eq " + currentStage +
            "&$top=20";

        var actions = await Xrm.WebApi.retrieveMultipleRecords("duc_stageaction", query);

        if (!actions.entities || actions.entities.length === 0) {
            console.log("No actions found for this stage.");
            return;
        }

        // 4. Find Action with Status = 100000007
        var matchedAction = null;

        for (let act of actions.entities) {

            let statusId = act._duc_defaultstatus_value;

            if (statusId) {

                let statusRec = await Xrm.WebApi.retrieveRecord(
                    "duc_processstatus",
                    statusId,
                    "?$select=duc_value"
                );

                if (statusRec.duc_value == 100000007) {
                    matchedAction = act.duc_stageactionid;
                    break;
                }
            }
        }

        if (!matchedAction) {
            console.log("No action with status 100000007 found.");
            return;
        }

        // 5. Update Last Action Taken
        await Xrm.WebApi.updateRecord(
            "duc_processextension",
            processExtensionId,
            {
                "duc_LastActionTaken_duc_ProcessExtension@odata.bind":
                    "/duc_stageactions(" + matchedAction + ")"
            }
        );

        // 6. Show message
        var msg = await getMessage("InspectionFinishedMessage");
        Xrm.Utility.alertDialog(msg);

    } catch (e) {
        console.error("Error:", e);
    }
}

(async function () {
    // var created = await createTawasolRequest();

    // if (created) {
    //     await updateLastActionOnProcessExtension();
    // }
    await openTawasolRequestQuickCreate();
})();