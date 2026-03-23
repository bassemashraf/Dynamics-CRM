/***************************************
 * HELPERS
 ***************************************/
function isOffline() {
    try {
        if (Xrm.Utility.getGlobalContext().client.isOffline()) return true;
        if (Xrm.Utility.getGlobalContext().client.getClientState() === "Offline") return true;
    } catch (e) { }
    return false;
}

function getUserLanguage() {
    return Xrm.Utility.getGlobalContext().userSettings.languageId === 1025 ? "ar" : "en";
}

function setSectionVisible(formContext, tabName, sectionName, visible) {
    try {
        var tab = formContext.ui.tabs.get(tabName);
        if (!tab) return;
        var section = tab.sections.get(sectionName);
        if (section) section.setVisible(visible);
    } catch (e) {
        console.warn("[setSectionVisible] " + tabName + " > " + sectionName + ": " + e.message);
    }
}

/***************************************
 * SECTION VISIBILITY FROM CONFIG
 * Works online + offline:
 * - Online:  _duc_incidenttype_value eq <GUID>
 * - Offline: duc_incidenttype eq <GUID>
 ***************************************/
async function toggleSectionFromConfigByCode(formContext, tabName, sectionName, configCode) {
    try {
        var isOff = isOffline();

        // Hide section by default
        setSectionVisible(formContext, tabName, sectionName, false);

        // Step 1: Get Work Order Incident lookup
        var woiAttr = formContext.getAttribute("msdyn_workorderincident");
        if (!woiAttr || !woiAttr.getValue()) {
            console.warn("[toggleSection] No Work Order Incident on form for section: " + sectionName);
            return;
        }
        var woiId = woiAttr.getValue()[0].id.replace(/[{}]/g, "");

        // Step 2: Retrieve Incident Type from Work Order Incident
        // Online uses _msdyn_incidenttype_value, offline reads it from the record object directly
        var woiRecord = await Xrm.WebApi.retrieveRecord(
            "msdyn_workorderincident",
            woiId,
            "?$select=_msdyn_incidenttype_value"
        );

        var incidentTypeId = woiRecord._msdyn_incidenttype_value;
        if (!incidentTypeId) {
            console.warn("[toggleSection] No Incident Type on Work Order Incident.");
            return;
        }
        incidentTypeId = incidentTypeId.replace(/[{}]/g, "");

        // Step 3: Query config — use schema field for offline, _value alias for online
        var incidentTypeField = isOff ? "duc_incidenttype" : "_duc_incidenttype_value";

        var query = "?$select=duc_actionvalue" +
            "&$filter=duc_code eq '" + configCode + "'" +
            " and " + incidentTypeField + " eq " + incidentTypeId;

        var result = await Xrm.WebApi.retrieveMultipleRecords("duc_incidenttypeconfigurations", query);

        if (!result.entities || result.entities.length === 0) {
            console.warn("[toggleSection] No config found for code: " + configCode);
            return;
        }

        var show = (result.entities[0].duc_actionvalue === 1);

        // Special rule: Penalties section requires duc_question1 = "مخالفة"
        if (show && sectionName === "Work_Order_Penalties_Section") {
            var q1 = formContext.getAttribute("duc_question1");
            var q1Val = q1 ? (q1.getValue() || "") : "";
            show = q1Val.toString().trim().indexOf("مخالفة") > -1;
        }

        setSectionVisible(formContext, tabName, sectionName, show);

    } catch (e) {
        var errMsg = "[toggleSectionFromConfigByCode] Error for code: " + configCode
            + "\nSection: " + sectionName
            + "\nOffline: " + isOffline()
            + "\nError: " + (e.message || JSON.stringify(e));
        console.error(errMsg, e);
        Xrm.Navigation.openAlertDialog({ text: errMsg });
    }
}

/***************************************
 * SECTION TOGGLE: Run all sections
 ***************************************/
async function applyAllSectionVisibilities(formContext) {
    var sectionsToToggle = [
        { tab: "GeneralTab", section: "Samples_Section", code: "Work Order Service Task - Sample Section" },
        { tab: "GeneralTab", section: "Products_Production_Capacity_Section", code: "Work Order Service Task - Products Section" },
        { tab: "GeneralTab", section: "Raw_Materials_Section", code: "Work Order Service Task - Raw Materials Section" },
        { tab: "GeneralTab", section: "Risk_Level_Section", code: "Work Order Service Task - Risk_Level_Section" },
        { tab: "GeneralTab", section: "Customer_Assets_Section", code: "Work Order Service Task - Work_Order_Customer_Assets_Section" },
        { tab: "GeneralTab", section: "Property_Assets_Section", code: "Work Order Service Task - Property_Assets_Section" },
        { tab: "GeneralTab", section: "Work_Order_Penalties_Section", code: "Work Order Service Task - Work_Order_Penalties_Section" },
        { tab: "GeneralTab", section: "GeneralTab_section_permitsDetails", code: "GeneralTab_section_permitsDetails" }
    ];

    // Run all in parallel for speed — each handles its own errors
    await Promise.all(
        sectionsToToggle.map(function (x) {
            return toggleSectionFromConfigByCode(formContext, x.tab, x.section, x.code);
        })
    );
}

/***************************************
 * SAVE + REFRESH
 ***************************************/
async function saveAndRefresh(formContext) {
    var step = "";
    try {
        Xrm.Utility.showProgressIndicator(
            getUserLanguage() === "ar" ? "جاري الحفظ..." : "Saving..."
        );

        // Save only if form is dirty
        step = "Saving form";
        if (formContext.data.entity.getIsDirty()) {
            await formContext.data.save();
        }

        step = "Applying section visibilities";
        await applyAllSectionVisibilities(formContext);

        step = "Refreshing form";
        Xrm.Utility.closeProgressIndicator();

        // Use data.refresh() instead of openForm() —
        // openForm with an entityId can fail in offline mode (platform can't locate
        // record in local cache by ID). data.refresh(false) reloads the current form
        // from the offline cache without navigating away — same pattern as wo_finish.js.
        setTimeout(function () {
            try { formContext.data.refresh(false); } catch (e) {
                console.warn("[saveAndRefresh] data.refresh failed: " + e.message);
            }
        }, 1500);

    } catch (e) {
        Xrm.Utility.closeProgressIndicator();
        var errMsg = "[saveAndRefresh] Error at step: " + step
            + "\nOffline: " + isOffline()
            + "\nError: " + (e.message || JSON.stringify(e));
        console.error(errMsg, e);
        Xrm.Navigation.openAlertDialog({ text: errMsg });
    }
}

/***************************************
 * ENTRY POINT
 ***************************************/
(async function () {
    var formContext = Xrm.Page;
    await saveAndRefresh(formContext);
})();