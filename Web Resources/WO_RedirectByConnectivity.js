/**
 * WO_RedirectByConnectivity.js
 * ---------------------------------------------------------------
 * Work Order Form — OnLoad handler
 *
 * Redirects the user to the correct Work Order form based on
 * their current connectivity state:
 *
 *   OFFLINE  →  Offline Work Order form  (OFFLINE_FORM_ID)
 *   ONLINE   →  Online  Work Order form  (ONLINE_FORM_ID)
 *
 * This prevents an online user from seeing the offline form and
 * an offline user from seeing the online form.
 *
 * HOW TO USE:
 *   1. Register this function on the Work Order form's OnLoad event.
 *   2. Replace the placeholder GUIDs below with the real form IDs
 *      from Settings → Customizations → Work Order → Forms.
 * ---------------------------------------------------------------
 */

// ---------------------------------------------------------------
// ⚙️  CONFIGURATION — replace these GUIDs with your real form IDs
// ---------------------------------------------------------------
var WO_FORM_IDS = {
    offline: "b5ea80ff-301f-f111-88b1-6045bd8e2841", // TODO: replace with real Offline form GUID
    online: "eded7d77-6dc4-ed11-b596-6045bdf00fa1"  // TODO: replace with real Online  form GUID
};
// ---------------------------------------------------------------

/**
 * Runs on the Work Order form OnLoad event.
 * Checks connectivity and redirects to the appropriate form
 * if the user is on the wrong one.
 *
 * @param {Object} executionContext - CRM form execution context.
 */
function isOffline() {
    try {
        // Check if user is on an offline profile (works even WITH internet connection).
        // isAvailableOffline returns true when the entity is part of the active
        // Mobile Offline profile — meaning the user is on the offline-first app.
        if (Xrm.WebApi.isAvailableOffline &&
            Xrm.WebApi.isAvailableOffline("msdyn_workorder")) return true;
        if (Xrm.Utility.getGlobalContext().client.isOffline()) return true;
        if (Xrm.Utility.getGlobalContext().client.getClientState() === "Offline") return true;
    } catch (e) { }
    return false;
}

function WO_RedirectByConnectivity(executionContext) {
    var formContext = executionContext.getFormContext();

    try {
        var isCurrentlyOffline = isOffline();

        // Get the ID of the form that is currently open
        var currentFormId = formContext.ui.formSelector.getCurrentItem().getId().toLowerCase();

        var targetFormId = isCurrentlyOffline
            ? WO_FORM_IDS.offline.toLowerCase()
            : WO_FORM_IDS.online.toLowerCase();

        // If the user is already on the correct form — do nothing
        if (currentFormId === targetFormId) {
            console.log("[WO_Redirect] User is on the correct form. No redirect needed. (offline=" + isCurrentlyOffline + ")");
            return;
        }

        console.log("[WO_Redirect] Wrong form detected. Redirecting... (offline=" + isCurrentlyOffline + ", target=" + targetFormId + ")");

        // Get the current record ID to reopen the same record on the target form
        var entityId = formContext.data.entity.getId().replace(/[{}]/g, "");
        var entityName = formContext.data.entity.getEntityName();

        // Navigate to the correct form for the same record
        var pageInput = {
            pageType: "entityrecord",
            entityName: entityName,
            entityId: entityId,
            formId: targetFormId
        };

        // target: 1  = current window (inline, no popup)
        var navigationOptions = {
            target: 1
        };

        Xrm.Navigation.navigateTo(pageInput, navigationOptions).then(
            function () {
                console.log("[WO_Redirect] Successfully redirected to " + (isCurrentlyOffline ? "Offline" : "Online") + " form.");
            },
            function (error) {
                console.error("[WO_Redirect] Navigation failed: " + (error.message || error));
            }
        );

    } catch (e) {
        console.error("[WO_Redirect] Unexpected error: " + e.message);
    }
}
