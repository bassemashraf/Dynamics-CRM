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

function isOffline() {
	try {
		if (Xrm.Utility.getGlobalContext().client.isOffline()) return true;
		if (Xrm.Utility.getGlobalContext().client.getClientState() === "Offline") return true;
	} catch (e) { }
	return false;
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

		var isOff = isOffline();
		var wostFilterField = isOff ? "duc_msdyn_workorderservicetask" : "_duc_msdyn_workorderservicetask_value";

		const result = await Xrm.WebApi.retrieveMultipleRecords("duc_questionanswersconfiguration",
			`?$select=duc_sequencefromcompliancelevel&$filter=${wostFilterField} eq ${woId}`);
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
		var errMsg = "[calculateComplianceLevel] Error."
			+ "\nRecordId: " + (formContext.data.entity.getId() || "N/A")
			+ "\nOffline: " + isOffline()
			+ "\nError: " + (error.message || JSON.stringify(error));
		console.error(errMsg, error);
		hideLoader();
		Xrm.Navigation.openAlertDialog({ text: errMsg });
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

				await Xrm.Navigation.openAlertDialog({ text: "يجب تسجيل مخالفة واحد على الأقل." }).then(
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
		var errMsg = "[runAsyncOpenWorkOrder] Error."
			+ "\nOffline: " + isOffline()
			+ "\nError: " + (e.message || JSON.stringify(e));
		console.error(errMsg, e);
		Xrm.Navigation.openAlertDialog({ text: errMsg });
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
	var step = "";
	try {
		var woAttr = formContext.getAttribute("msdyn_workorder");
		if (!woAttr || !woAttr.getValue()) return;

		var woId = woAttr.getValue()[0].id.replace(/[{}]/g, "");
		var isOff = isOffline();
		var woFilterField = isOff ? "msdyn_workorder" : "_msdyn_workorder_value";

		step = "Retrieving booking for work order";
		var bookingResult = await Xrm.WebApi.retrieveMultipleRecords(
			"bookableresourcebooking",
			`?$select=bookableresourcebookingid&$filter=${woFilterField} eq ${woId}&$orderby=createdon asc&$top=1`
		);

		if (!bookingResult.entities.length) return;

		var bookingId = bookingResult.entities[0].bookableresourcebookingid;

		step = "Retrieving 'In Progress' booking status";
		var statusResult = await Xrm.WebApi.retrieveMultipleRecords(
			"bookingstatus",
			`?$select=bookingstatusid&$filter=name eq 'In Progress'`
		);

		if (!statusResult.entities.length) {
			return Xrm.Navigation.openAlertDialog({ text: getLocalizedString("bookingStatusNotFound") });
		}

		var statusId = statusResult.entities[0].bookingstatusid;

		step = "Updating booking status to 'In Progress'";
		await Xrm.WebApi.updateRecord("bookableresourcebooking", bookingId, {
			"BookingStatus@odata.bind": `/bookingstatuses(${statusId})`
		});
	} catch (e) {
		var errMsg = "[sBIP] Error at step: " + step
			+ "\nOffline: " + isOffline()
			+ "\nError: " + (e.message || JSON.stringify(e));
		console.error(errMsg, e);
		Xrm.Navigation.openAlertDialog({ text: errMsg });
	}
}


function isViolationSelected(formContext) {
	const attr = formContext.getAttribute("duc_question1");
	if (!attr) return false;

	const value = attr.getValue();
	if (!value) return false;

	return value.trim() === "مخالفة";
}


async function hasValidComplianceStatus(formContext) {
	try {
		const taskId = formContext.data.entity.getId();
		if (!taskId) return null;

		const cleanId = taskId.replace(/[{}]/g, "");
		var isOff = isOffline();
		var wostFilterField = isOff ? "duc_workorderservicetask" : "_duc_workorderservicetask_value";

		const result = await Xrm.WebApi.retrieveMultipleRecords(
			"duc_workorderpenalties",
			`?$select=duc_compliancestatus&$filter=${wostFilterField} eq ${cleanId}`
		);

		// NO penalties → do NOT block
		if (!result.entities.length) {
			return null;
		}

		// penalties exist → check compliance
		return result.entities.some(r => r.duc_compliancestatus !== null);

	} catch (e) {
		var errMsg = "[hasValidComplianceStatus] Error checking penalty compliance."
			+ "\nTaskId: " + (formContext.data.entity.getId() || "N/A")
			+ "\nOffline: " + isOffline()
			+ "\nError: " + (e.message || JSON.stringify(e));
		console.error(errMsg, e);
		return null; // fail-safe: do not block
	}
}

// Call the main function
openWorkOrderAndFocusTab();