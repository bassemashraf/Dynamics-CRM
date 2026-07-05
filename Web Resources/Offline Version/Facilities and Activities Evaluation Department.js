// ============================================================
// Facilities and Activities Evaluation Department
// Offline/Online Compatible Version
// Date: 2026-07-04
//
// Changes from Online-only version:
// 1. Added isOffline() helper (full version from hideheader.js)
// 2. Replaced FetchXml query with OData $filter (characteristic lookup)
// 3. All $filter lookup fields branch on isOffline():
//    - msdyn_resourcerequirement:  msdyn_workorder / _msdyn_workorder_value
//    - msdyn_requirementcharacteristic: msdyn_resourcerequirement / _msdyn_resourcerequirement_value
//    - duc_processstage: duc_relatedprocess / _duc_relatedprocess_value
//    - duc_stageaction: duc_nextstage / _duc_nextstage_value
//    - bookableresource: userid / _userid_value
//    - bookableresourcebooking: msdyn_resourcerequirement / _msdyn_resourcerequirement_value
// 4. Ribbon button clicks (executeBookOnly, executeBookResource) guarded offline
// 5. Enhanced error messages include offline status for debugging
// ============================================================

(async function () {
    removeProcessingOverlay();
    await openSelectionPopup();
})();

function isOffline() {
    try {
        var ctx = Xrm.Utility.getGlobalContext().client;
        // Actually check if an entity is available offline — this returns true
        // ONLY when a Mobile Offline profile is enabled and includes the entity.
        // Just checking whether the API *exists* is not enough because the
        // Xrm.WebApi.offline namespace is present on all mobile apps.
        if (Xrm.WebApi.offline &&
            typeof Xrm.WebApi.offline.isAvailableOffline === "function" &&
            Xrm.WebApi.offline.isAvailableOffline("msdyn_workorder")) return true;
        if (Xrm.WebApi.isAvailableOffline &&
            Xrm.WebApi.isAvailableOffline("msdyn_workorder")) return true;
        if (ctx.isOffline()) return true;
        if (ctx.getClientState() === "Offline") return true;
    } catch (e) { }
    return false;
}

async function openSelectionPopup() {

    removeProcessingOverlay();

    if (document.getElementById("ducActionPopup"))
        return;

    var permissions = await getCurrentUserPermissions();

    var inspectorScope = permissions.duc_inspectorassignmentscope;

    var showUnitManager =
        permissions.duc_transfertounithead === true;

    var showDepartmentManager =
        permissions.duc_transfertosectionhead === true;

    var showBookResource =
        permissions.duc_selfassignworkorder === true &&
        inspectorScope === 100000000;

    var showAssignOnly =
        permissions.duc_selfassignworkorder === true &&
        inspectorScope === 100000001;

    var showAssignExcludingCurrentUser =
        permissions.duc_selfassignworkorder !== true &&
        inspectorScope === 100000001;

    if (
        !showUnitManager &&
        !showDepartmentManager &&
        !showBookResource &&
        !showAssignOnly
    ) {
        Xrm.Navigation.openAlertDialog({
            text: "ليس لديك صلاحية لتنفيذ أي إجراء."
        });
        return;
    }

    var popup = document.createElement("div");
    popup.id = "ducActionPopup";

    popup.innerHTML = `
<div style="
    position:fixed;
    inset:0;
    background:rgba(0,0,0,.32);
    display:flex;
    align-items:center;
    justify-content:center;
    z-index:999999;
">

    <div style="
        background:#ffffff;
        width:420px;
        border-radius:8px;
        border:1px solid #edebe9;
        box-shadow:0 6px 18px rgba(0,0,0,.18);
        font-family:'Segoe UI',Tahoma,sans-serif;
        overflow:hidden;
    ">

        <div style="
            padding:20px 24px;
            border-bottom:1px solid #edebe9;
        ">
            <div style="
                font-size:20px;
                font-weight:600;
                color:#323130;
            ">
                اختر الإجراء المطلوب
            </div>
        </div>

        <div style="padding:20px;">

            <button id="btnUnitManager"
                style="
                    width:100%;
                    padding:10px 16px;
                    margin-bottom:10px;
                    background:#ffffff;
                    color:#323130;
                    border:1px solid #8a8886;
                    border-radius:4px;
                    font-size:14px;
                    font-weight:600;
                    cursor:pointer;
                    display:${showUnitManager ? "block" : "none"};
                ">
                إرسال إلى رئيس الوحدة
            </button>

            <button id="btnDepartmentManager"
                style="
                    width:100%;
                    padding:10px 16px;
                    margin-bottom:10px;
                    background:#ffffff;
                    color:#323130;
                    border:1px solid #8a8886;
                    border-radius:4px;
                    font-size:14px;
                    font-weight:600;
                    cursor:pointer;
                    display:${showDepartmentManager ? "block" : "none"};
                ">
                إرسال إلى رئيس القسم
            </button>

            <button id="btnBookResource"
                style="
                    width:100%;
                    padding:10px 16px;
                    margin-bottom:10px;
                    background:#0078d4;
                    color:#ffffff;
                    border:none;
                    border-radius:4px;
                    font-size:14px;
                    font-weight:600;
                    cursor:pointer;
                    display:${showBookResource ? "block" : "none"};
                ">
                تعيين ذاتي
            </button>

            <button id="btnAssignOnly"
                style="
                    width:100%;
                    padding:10px 16px;
                    margin-bottom:10px;
                    background:#0078d4;
                    color:#ffffff;
                    border:none;
                    border-radius:4px;
                    font-size:14px;
                    font-weight:600;
                    cursor:pointer;
                    display:${showAssignOnly ? "block" : "none"};
                ">
                تعيين
            </button>

            <button id="btnCancel"
                style="
                    width:100%;
                    padding:10px 16px;
                    background:#f3f2f1;
                    color:#323130;
                    border:1px solid #d2d0ce;
                    border-radius:4px;
                    font-size:14px;
                    font-weight:600;
                    cursor:pointer;
                ">
                إلغاء
            </button>

        </div>

    </div>

</div>
`;

    document.body.appendChild(popup);

    if (showUnitManager) {
        document.getElementById("btnUnitManager").onclick = async function () {
            closePopup();

            await executeStageAction(
                "قسم تقييم المنشآت والأنشطة - وحدة المعاينة",
                "Facilities and Activities Evaluation Unit Head"
            );
        };
    }

    if (showDepartmentManager) {
        document.getElementById("btnDepartmentManager").onclick = async function () {
            closePopup();

            await executeStageAction(
                [
                    "قسم تقييم المنشآت والأنشطة"
                ],
                "Facilities and Activities Dept - Section Head"

            );
        };
    }

    if (showBookResource) {
        document.getElementById("btnBookResource").onclick = async function () {
            closePopup();
            await executeBookResource();
        };
    }

    if (showAssignOnly) {
        document.getElementById("btnAssignOnly").onclick = async function () {
            closePopup();
            await executeAssignOnlyWithCharacteristic();

        };
    }

    document.getElementById("btnCancel").onclick = closePopup;
}

function closePopup() {
    var popup = document.getElementById("ducActionPopup");

    if (popup)
        popup.remove();
}

async function executeStageAction(characteristicNames, targetStageName) {

    try {

        const isArabic =
            Xrm.Utility.getGlobalContext().userSettings.languageId === 1025;

        removeProcessingOverlay();

        createProcessingOverlay(
            isArabic ? "جاري التنفيذ..." : "Processing..."
        );

        if (!Array.isArray(characteristicNames)) {
            characteristicNames = [characteristicNames];
        }

        await createRequirementCharacteristics(characteristicNames);

        await moveToStageByName(targetStageName);

        removeProcessingOverlay();

        Xrm.Navigation.openForm({
            entityName: Xrm.Page.data.entity.getEntityName(),
            entityId: Xrm.Page.data.entity.getId()
        });

    } catch (e) {

        var errMsg = "[executeStageAction] Error."
            + "\nStage: " + targetStageName
            + "\nOffline: " + isOffline()
            + "\nError: " + (e.message || String(e));

        console.error(errMsg, e);

        Xrm.Navigation.openErrorDialog({
            message: errMsg
        });

    } finally {

        removeProcessingOverlay();
    }
}

async function createRequirementCharacteristics(characteristicNames) {

    for (var i = 0; i < characteristicNames.length; i++) {
        await createRequirementCharacteristic(
            characteristicNames[i],
            i === 0
        );
    }
}

async function createRequirementCharacteristic(characteristicName, deleteOld) {

    var formContext = Xrm.Page;

    const isArabic =
        Xrm.Utility.getGlobalContext().userSettings.languageId === 1025;

    var workOrderId = formContext.data.entity.getId();

    if (!workOrderId) {
        throw new Error(
            isArabic
                ? "لم يتم العثور على أمر العمل"
                : "Work Order Id not found."
        );
    }

    workOrderId = workOrderId.replace(/[{}]/g, "");

    var isOff = isOffline();

    var safeCharacteristicName = characteristicName.replace(/'/g, "''");

    // OFFLINE FIX: FetchXml is NOT supported offline → use OData $filter instead
    var characteristicResult =
        await Xrm.WebApi.retrieveMultipleRecords(
            "characteristic",
            "?$select=characteristicid&$filter=name eq '" + safeCharacteristicName + "'&$top=1"
        );

    if (!characteristicResult.entities.length) {
        throw new Error(
            isArabic
                ? "لم يتم العثور على الخاصية: " + characteristicName
                : "Characteristic not found: " + characteristicName
        );
    }

    var characteristicId =
        characteristicResult.entities[0].characteristicid;

    // OFFLINE FIX: $filter lookup field — offline needs schema name
    var woFilterField = isOff ? "msdyn_workorder" : "_msdyn_workorder_value";

    var requirementResult =
        await Xrm.WebApi.retrieveMultipleRecords(
            "msdyn_resourcerequirement",
            "?$select=msdyn_resourcerequirementid&$filter=" + woFilterField + " eq " +
            workOrderId +
            "&$top=1"
        );

    if (!requirementResult.entities.length) {
        throw new Error(
            isArabic
                ? "لم يتم العثور على متطلب المورد"
                : "Resource Requirement not found."
        );
    }

    var resourceRequirementId =
        requirementResult.entities[0].msdyn_resourcerequirementid;

    if (deleteOld !== false) {

        // OFFLINE FIX: $filter lookup field — offline needs schema name
        var reqCharFilterField = isOff ? "msdyn_resourcerequirement" : "_msdyn_resourcerequirement_value";

        var deleteResult =
            await Xrm.WebApi.retrieveMultipleRecords(
                "msdyn_requirementcharacteristic",
                "?$select=msdyn_requirementcharacteristicid&$filter=" + reqCharFilterField + " eq " +
                resourceRequirementId +
                " and duc_manuallyadded eq true"
            );

        for (var i = 0; i < deleteResult.entities.length; i++) {
            await Xrm.WebApi.deleteRecord(
                "msdyn_requirementcharacteristic",
                deleteResult.entities[i].msdyn_requirementcharacteristicid
            );
        }
    }

    await Xrm.WebApi.createRecord(
        "msdyn_requirementcharacteristic",
        {
            "msdyn_Characteristic@odata.bind":
                "/characteristics(" + characteristicId + ")",

            "msdyn_ResourceRequirement@odata.bind":
                "/msdyn_resourcerequirements(" + resourceRequirementId + ")",

            "duc_manuallyadded": true
        }
    );
}

async function moveToStageByName(targetStageName) {

    var woId = Xrm.Page.data.entity.getId().replace(/[{}]/g, "");

    var isOff = isOffline();

    var woRecord = await Xrm.WebApi.retrieveRecord(
        "msdyn_workorder",
        woId,
        "?$select=_duc_processextension_value"
    );

    if (!woRecord._duc_processextension_value)
        throw new Error("Process Extension not found.");

    var peId = woRecord._duc_processextension_value;

    var peRecord = await Xrm.WebApi.retrieveRecord(
        "duc_processextension",
        peId,
        "?$select=_duc_processdefinition_value"
    );

    var processDefinitionId =
        peRecord._duc_processdefinition_value;

    if (!processDefinitionId)
        throw new Error("Process Definition not found.");

    var safeStageName = targetStageName.replace(/'/g, "''");

    // OFFLINE FIX: $filter lookup field — offline needs schema name
    var processFilterField = isOff ? "duc_relatedprocess" : "_duc_relatedprocess_value";

    var stageResult = await Xrm.WebApi.retrieveMultipleRecords(
        "duc_processstage",
        "?$select=duc_processstageid,duc_name" +
        "&$filter=duc_name eq '" + safeStageName + "'" +
        " and " + processFilterField + " eq " + processDefinitionId +
        "&$top=1"
    );

    if (!stageResult.entities.length)
        throw new Error("Stage '" + targetStageName + "' not found.");

    var stageId = stageResult.entities[0].duc_processstageid;

    // OFFLINE FIX: $filter lookup field — offline needs schema name
    var nextStageFilterField = isOff ? "duc_nextstage" : "_duc_nextstage_value";

    var actionResult =
        await Xrm.WebApi.retrieveMultipleRecords(
            "duc_stageaction",
            "?$select=duc_stageactionid" +
            "&$filter=" + nextStageFilterField + " eq " + stageId +
            "&$top=1"
        );

    if (!actionResult.entities.length)
        throw new Error(
            "Stage Action for '" + targetStageName + "' not found."
        );

    var actionId = actionResult.entities[0].duc_stageactionid;

    var updateRecord = {};

    updateRecord["duc_CurrentStage_duc_ProcessExtension@odata.bind"] =
        "/duc_processstages(" + stageId + ")";

    updateRecord["duc_LastActionTaken_duc_ProcessExtension@odata.bind"] =
        "/duc_stageactions(" + actionId + ")";

    await Xrm.WebApi.updateRecord(
        "duc_processextension",
        peId,
        updateRecord
    );
}

async function executeBookOnly() {

    try {

        removeProcessingOverlay();

        const isArabic =
            Xrm.Utility.getGlobalContext().userSettings.languageId === 1025;

        // OFFLINE FIX: Ribbon buttons are not available in the mobile offline app
        if (isOffline()) {
            await Xrm.Navigation.openAlertDialog({
                text: isArabic
                    ? "هذا الإجراء غير متاح في وضع عدم الاتصال."
                    : "This action is not available in offline mode."
            });
            return;
        }

        var bookbtnId = "msdyn_workorder.Form.BookResource.Button";

        var bookbuttons = document.querySelectorAll(
            "button[data-id*='" + bookbtnId + "']"
        );

        if (bookbuttons.length === 0) {
            await Xrm.Navigation.openAlertDialog({
                text: "Page component loading in progress, please try again in a moment"
            });
            return;
        }

        bookbuttons[0].click();

    } catch (e) {

        var errMsg = "[executeBookOnly] Error."
            + "\nOffline: " + isOffline()
            + "\nError: " + (e.message || String(e));

        console.error(errMsg, e);

        Xrm.Navigation.openErrorDialog({
            message: errMsg
        });

    } finally {

        removeProcessingOverlay();
    }
}

async function executeBookResource() {

    try {

        removeProcessingOverlay();

        const isArabic =
            Xrm.Utility.getGlobalContext().userSettings.languageId === 1025;

        // OFFLINE FIX: Ribbon buttons are not available in the mobile offline app
        if (isOffline()) {
            await Xrm.Navigation.openAlertDialog({
                text: isArabic
                    ? "هذا الإجراء غير متاح في وضع عدم الاتصال."
                    : "This action is not available in offline mode."
            });
            return;
        }

        var bookbtnId = "msdyn_workorder.Form.BookResource.Button";
        var refreshbtnId = "msdyn_workorder.RefreshModernButton";

        var formContext = Xrm.Page;

        var workOrderId =
            formContext.data.entity.getId().replace(/[{}]/g, "");

        var workOrderName =
            formContext.getAttribute("msdyn_name").getValue();

        var uniqueSuffix = new Date().getTime();

        var characteristicName =
            workOrderName + " - " + uniqueSuffix;

        var userId =
            Xrm.Utility.getGlobalContext().userSettings.userId.replace(/[{}]/g, "");

        var isOff = isOffline();

        var characteristicResult = await Xrm.WebApi.createRecord(
            "characteristic",
            {
                name: characteristicName
            }
        );

        var characteristicId =
            characteristicResult.id.replace(/[{}]/g, "");

        // OFFLINE FIX: $filter lookup field — offline needs schema name
        var woFilterField = isOff ? "msdyn_workorder" : "_msdyn_workorder_value";

        var requirementResult =
            await Xrm.WebApi.retrieveMultipleRecords(
                "msdyn_resourcerequirement",
                "?$filter=" + woFilterField + " eq " + workOrderId
            );

        if (!requirementResult.entities.length) {
            throw new Error("Resource Requirement not found.");
        }

        var requirementId =
            requirementResult.entities[0].msdyn_resourcerequirementid;

        // OFFLINE FIX: $filter lookup field — offline needs schema name
        var reqCharFilterField = isOff ? "msdyn_resourcerequirement" : "_msdyn_resourcerequirement_value";

        var oldLinks =
            await Xrm.WebApi.retrieveMultipleRecords(
                "msdyn_requirementcharacteristic",
                "?$filter=" + reqCharFilterField + " eq " + requirementId +
                " and duc_manuallyadded eq true"
            );

        for (var i = 0; i < oldLinks.entities.length; i++) {
            await Xrm.WebApi.deleteRecord(
                "msdyn_requirementcharacteristic",
                oldLinks.entities[i].msdyn_requirementcharacteristicid
            );
        }

        await Xrm.WebApi.createRecord(
            "msdyn_requirementcharacteristic",
            {
                "msdyn_Characteristic@odata.bind":
                    "/characteristics(" + characteristicId + ")",

                "msdyn_ResourceRequirement@odata.bind":
                    "/msdyn_resourcerequirements(" + requirementId + ")",

                "duc_manuallyadded": true
            }
        );

        // OFFLINE FIX: $filter lookup field — offline needs schema name
        var userFilterField = isOff ? "userid" : "_userid_value";

        var brResult =
            await Xrm.WebApi.retrieveMultipleRecords(
                "bookableresource",
                "?$filter=" + userFilterField + " eq " + userId
            );

        if (brResult.entities.length > 0) {

            var bookableResourceId =
                brResult.entities[0].bookableresourceid;

            await Xrm.WebApi.createRecord(
                "bookableresourcecharacteristic",
                {
                    "Characteristic@odata.bind":
                        "/characteristics(" + characteristicId + ")",

                    "Resource@odata.bind":
                        "/bookableresources(" + bookableResourceId + ")"
                }
            );
        }

        var bookbuttons = document.querySelectorAll(
            "button[data-id*='" + bookbtnId + "']"
        );

        if (bookbuttons.length === 0) {
            await Xrm.Navigation.openAlertDialog({
                text: "Page component loading in progress, please try again in a moment"
            });
            return;
        }

        bookbuttons[0].click();

        // OFFLINE FIX: $filter lookup field — offline needs schema name
        var reqBookingFilterField = isOff ? "msdyn_resourcerequirement" : "_msdyn_resourcerequirement_value";

        var checkBooking = function () {

            Xrm.WebApi.retrieveMultipleRecords(
                "bookableresourcebooking",
                "?$filter=" + reqBookingFilterField + " eq " + requirementId
            ).then(function (result) {

                if (result.entities.length > 0) {

                    var refreshbuttons = document.querySelectorAll(
                        "button[data-id*='" + refreshbtnId + "']"
                    );

                    if (refreshbuttons.length > 0) {
                        refreshbuttons[0].click();
                    }

                } else {
                    setTimeout(checkBooking, 5000);
                }

            }).catch(console.error);
        };

        setTimeout(checkBooking, 5000);

    } catch (e) {

        var errMsg = "[executeBookResource] Error."
            + "\nOffline: " + isOffline()
            + "\nError: " + (e.message || String(e));

        console.error(errMsg, e);

        Xrm.Navigation.openErrorDialog({
            message: errMsg
        });

    } finally {

        removeProcessingOverlay();
    }
}

function createProcessingOverlay(message) {

    removeProcessingOverlay();

    var overlay = document.createElement("div");

    overlay.id = "ducProcessingOverlay";

    overlay.style.position = "fixed";
    overlay.style.inset = "0";
    overlay.style.background = "rgba(0,0,0,.35)";
    overlay.style.display = "flex";
    overlay.style.alignItems = "center";
    overlay.style.justifyContent = "center";
    overlay.style.zIndex = "999999";

    overlay.innerHTML = `
        <div style="
            background:#ffffff;
            border-radius:16px;
            padding:24px 40px;
            min-width:250px;
            text-align:center;
            box-shadow:0 8px 25px rgba(0,0,0,.2);
        ">
            <div style="
                width:30px;
                height:30px;
                margin:0 auto 12px auto;
                border:4px solid #e5e5e5;
                border-top:4px solid #0078d4;
                border-radius:50%;
                animation:spinOverlay 1s linear infinite;
            "></div>

            <div style="
                font-family:Segoe UI,Tahoma,Arial,sans-serif;
                font-size:14px;
            ">
                ${message}
            </div>
        </div>
    `;

    if (!document.getElementById("spinOverlayStyle")) {

        var style = document.createElement("style");

        style.id = "spinOverlayStyle";

        style.innerHTML = `
            @keyframes spinOverlay {
                from { transform: rotate(0deg); }
                to { transform: rotate(360deg); }
            }
        `;

        document.head.appendChild(style);
    }

    document.body.appendChild(overlay);

    return overlay;
}

function removeProcessingOverlay() {

    var overlay = document.getElementById("ducProcessingOverlay");

    if (overlay) {
        overlay.remove();
    }
}

async function getCurrentUserPermissions() {

    var userId =
        Xrm.Utility.getGlobalContext().userSettings.userId.replace(/[{}]/g, "");

    return await Xrm.WebApi.retrieveRecord(
        "systemuser",
        userId,
        "?$select=duc_assignreassigninspector,duc_inspectorassignmentscope,duc_selfassignworkorder,duc_transfertosectionhead,duc_transfertounithead"
    );
}

async function executeAssignOnlyWithCharacteristic() {

    try {

        const isArabic =
            Xrm.Utility.getGlobalContext().userSettings.languageId === 1025;

        removeProcessingOverlay();

        createProcessingOverlay(
            isArabic ? "جاري التنفيذ..." : "Processing..."
        );

        await createRequirementCharacteristics([
            "قسم تقييم المنشآت والأنشطة"
        ]);

        removeProcessingOverlay();

        await executeBookOnly();

    } catch (e) {

        var errMsg = "[executeAssignOnlyWithCharacteristic] Error."
            + "\nOffline: " + isOffline()
            + "\nError: " + (e.message || String(e));

        console.error(errMsg, e);

        Xrm.Navigation.openErrorDialog({
            message: errMsg
        });

    } finally {

        removeProcessingOverlay();
    }
}
