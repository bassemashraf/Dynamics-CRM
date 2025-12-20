// JavaScript source code

function onFormLoad(executionContext) {
	var formContext = executionContext.getFormContext();

	// 1. Fill the Department field from the logged-in user's OU (New Logic)
	setDepartmentFromUserOU(executionContext);
	// customizeFieldsByDepartment(executionContext, deptName);
	resizeSection(executionContext);
}
function resizeSection(executionContext) {

	const attempt = () => {
		const sectionName = "MainPCFSection";
		const secondSectionName = "Details"; // logical name
		const newWidth = "194.6%";
		const newMargin = "36%";
		// Get first section
		const elMain = document.querySelector('[aria-label="MainPCFSection"]');
		if (elMain) {
			elMain.style.width = newWidth;
			elMain.style.maxWidth = newWidth;
			elMain.style.flex = `0 0 ${newWidth}`;
			console.log("Main section resized");
		}

		// Get second section
		const elSecond = document.querySelector('[aria-label="Details"]');
		if (elSecond) {
			elSecond.style.marginTop = newMargin;
			console.log("Second section margin updated");
		}

		// Retry if either section is not yet in DOM
		if (!elMain || !elSecond) {
			setTimeout(attempt, 500);
		}
	};

	attempt();
}


function searchAndPopulateServiceAccount(executionContext) {
	var formContext = executionContext.getFormContext();
	formContext.getControl("duc_searchserviceaccount").clearNotification("NoAccountFound");
	// 1. Get the search text
	var searchText = formContext.getAttribute("duc_searchserviceaccount").getValue();

	// If search field is empty, clear the Service Account and return
	if (searchText == null || searchText === "") {
		formContext.getAttribute("duc_subaccount").setValue(null);
		return;
	}

	// 2. PRIORITY 1: Search directly on the Account Entity
	var accountOptions = "?$select=_duc_account_value&$filter=duc_name eq '" + searchText + "'";

	Xrm.WebApi.retrieveMultipleRecords("duc_accountpermit", accountOptions).then(
		function success(result) {
			if (result.entities.length > 0) {
				// MATCH FOUND IN ACCOUNT
				var accountRecord = result.entities[0];
				var duc_accountpermitid = accountRecord["duc_accountpermitid"]; // Guid
				var duc_account = accountRecord["_duc_account_value"]; // Lookup
				var duc_account_formatted = accountRecord["_duc_account_value@OData.Community.Display.V1.FormattedValue"];
				var duc_account_lookuplogicalname = accountRecord["_duc_account_value@Microsoft.Dynamics.CRM.lookuplogicalname"];
				setServiceAccountLookup(formContext, duc_account, duc_account_formatted);
			} else {
				// NO MATCH IN ACCOUNT -> TRIGGER PRIORITY 2
				searchPermitEntity(formContext, searchText);
			}
		},
		function (error) {
			console.error("Error searching Account: " + error.message);
		}
	);
}

function searchPermitEntity(formContext, searchText) {
	// 3. PRIORITY 2: Search on the Permit Entity
	// UPDATED: Using 'duc_serviceaccount' as the lookup field name

	var permitOptions = "?$select=_duc_serviceaccount_value&$filter=duc_number eq '" + searchText + "'";

	Xrm.WebApi.retrieveMultipleRecords("duc_permit", permitOptions).then(
		function success(result) {
			if (result.entities.length > 0) {
				// MATCH FOUND IN PERMIT
				var permitRecord = result.entities[0];

				// Get the Account ID from the Permit record (using the updated schema name)
				var accountId = permitRecord["_duc_serviceaccount_value"];
				var accountName = permitRecord["_duc_serviceaccount_value@OData.Community.Display.V1.FormattedValue"];

				if (accountId) {
					setServiceAccountLookup(formContext, accountId, accountName);
				}
			} else {
				formContext.getControl("duc_searchserviceaccount").setNotification("This value is invalid", "NoAccountFound");
			}
		},
		function (error) {
			console.error("Error searching Permit: " + error.message);
		}
	);
}

/**
 * Show/hide fields and rename them based on user's department
 * @param {object} executionContext - Form execution context
 */
/**
 * Show/hide fields and rename them based on user's department (partial, case-insensitive match)
 * @param {object} executionContext - Form execution context
 */
/**
 * Customize fields based on department with per-department configuration
 * @param {object} executionContext - Form execution context
 */
/**
 * Customize fields based on department with per-department config and AR/EN labels
 * @param {object} executionContext - Form execution context
 */
function customizeFieldsByDepartment(executionContext, deptName) {
	var formContext = executionContext.getFormContext();

	// 1️⃣ Configuration: per field, per department, with labels in AR/EN
	var fieldConfig = {
		"duc_authority": {
			"التقييم والتصاريح": { visible: true, enabled: true, label: { en: "Project", ar: "المشروع" } },
		},
		"duc_anonymouscustomer": {
			"المحميات": { visible: true }
		},
		"new_campaign:": {
			"المحميات": { visible: true, enabled: true, label: { en: "Patrol", ar: "الدورية" } },
		}
	};

	// 2️⃣ Get current user's department


	var deptLower = deptName.toLowerCase(); // for case-insensitive match

	// Determine user's language
	var userLang = Xrm.Utility.getGlobalContext().userSettings.languageId;
	var isArabic = (userLang === 1025); // 1025 = Arabic

	// 3️⃣ Loop through fields and apply customization
	for (var field in fieldConfig) {
		if (!fieldConfig.hasOwnProperty(field)) continue;

		var control = formContext.getControl(field);
		if (!control) continue;

		var config = fieldConfig[field];
		var matchedDept = Object.keys(config).find(function (dept) {
			return deptLower.includes(dept.toLowerCase());
		});

		if (matchedDept) {
			var deptSettings = config[matchedDept];

			// Apply visibility
			if (deptSettings.visible !== undefined) {
				control.setVisible(deptSettings.visible);
			}

			// Apply enabled/disabled
			if (deptSettings.enabled !== undefined) {
				control.setDisabled(!deptSettings.enabled);
			}

			// Apply label based on language
			if (deptSettings.label) {
				var label = isArabic ? deptSettings.label.ar : deptSettings.label.en;
				if (label) control.setLabel(label);
			}
		} else {
			// Default: hide and disable if no matching department
			control.setVisible(false);
			control.setDisabled(true);
		}
	}

}



function setServiceAccountLookup(formContext, id, name) {
	// Helper function to set the lookup value
	var lookupValue = [];
	lookupValue[0] = {};
	lookupValue[0].id = id;
	lookupValue[0].name = name;
	lookupValue[0].entityType = "account"; // The target entity is always Account

	formContext.getAttribute("duc_subaccount").setValue(lookupValue);
	formContext.getAttribute("msdyn_serviceaccount").setValue(lookupValue);
}

function setDepartmentFromUserOU(executionContext) {
	var formContext = executionContext.getFormContext();
	var departmentField = formContext.getAttribute("duc_department");

	if (formContext.ui.getFormType() !== 1 || !departmentField || departmentField.getValue() !== null) {
		return;
	}

	var userId = Xrm.Utility.getGlobalContext().userSettings.userId.replace(/[{}]/g, "");

	Xrm.WebApi.retrieveRecord(
		"systemuser",
		userId,
		"?$select=_duc_organizationalunitid_value"
	).then(function (result) {

		var orgUnitId = result._duc_organizationalunitid_value;
		var orgUnitName =
			result["_duc_organizationalunitid_value@OData.Community.Display.V1.FormattedValue"];

		if (!orgUnitId) return;

		departmentField.setValue([{
			id: orgUnitId,
			name: orgUnitName,
			entityType: "msdyn_organizationalunit"
		}]);
		customizeFieldsByDepartment(executionContext, orgUnitName);

	}, function (error) {
		console.error(error.message);
	});
}


function filtercampaign() {
	var filterc = "<filter type='and'><condition attribute='new_inspectioncampaignid' operator='in'>";
	var filter = "";
	var loggedinid = Xrm.Utility.getGlobalContext().userSettings.userId;
	var fetcxml = "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>" +
		"<entity name='new_inspectioncampaign'>" +
		"<attribute name='new_inspectioncampaignid' />" +
		" <attribute name='new_name' />" +
		"<attribute name='createdon' />" +
		"<order attribute='new_name' descending='false' />" +
		"  <link-entity name='msdyn_organizationalunit' from='msdyn_organizationalunitid' to='duc_organizationalunitid' link-type='inner' alias='ac'>" +
		" <link-entity name='systemuser' from='duc_organizationalunitid' to='msdyn_organizationalunitid' link-type='inner' alias='ad'>" +
		"  <filter type='and'>" +
		"   <condition attribute='systemuserid' operator='eq' uitype='systemuser' value='" + loggedinid + "' />" +
		" </filter>" +
		"</link-entity>" +
		"</link-entity>" +
		"</entity>" +
		"</fetch>";
	var globalcontext = Xrm.Utility.getGlobalContext();
	var clientUrl = globalcontext.getClientUrl();
	clientUrl = clientUrl + "/api/data/v9.2/new_inspectioncampaigns?fetchXml=" + fetcxml;
	var $ = parent.$;
	$.ajax({
		url: clientUrl,
		type: "GET",
		contentType: "application/json;charset=utf-8",
		datatype: "json",
		async: false,
		beforeSend: function (Xmlhttprequest) { Xmlhttprequest.setRequestHeader("Accept", "application/json"); },
		success: function (result, textStatus, xhr) {

			if (result.value.length > 0) {
				for (var i = 0; i < result.value.length; i++) {
					filterc += "<value>{" + result.value[i].new_inspectioncampaignid + "}</value>";
				}
			}
			else {
				filterc += "<value>{00000000-0000-0000-0000-000000000000}</value>";
			}
			filterc += "</condition></filter>";
		},
		error: function (error) {
			console.log("Error retrieving data from campaign");
			filterc += "<value>{00000000-0000-0000-0000-000000000000}</value></condition></filter>";
		}
	});
	formContextCallback.getControl("new_campaign").addCustomFilter(filterc);
}

// JavaScript source code
function filterincidenttype() {
	filteri = "<filter type='and'><condition attribute='msdyn_incidenttypeid' operator='in'>";

	var loggedinid = Xrm.Utility.getGlobalContext().userSettings.userId;

	//Get M:M
	var linkQuery = "<link-entity name='duc_msdyn_incidenttype_msdyn_organizational' from='msdyn_incidenttypeid' to='msdyn_incidenttypeid' visible='false' intersect='true'>" +
		"<link-entity name='msdyn_organizationalunit' from='msdyn_organizationalunitid' to='msdyn_organizationalunitid' alias='ae'>" +
		"<link-entity name='systemuser' from='duc_organizationalunitid' to='msdyn_organizationalunitid' link-type='inner' alias='af'>" +
		"<filter type='and'>" +
		"<condition attribute='systemuserid' operator='eq' uitype='systemuser' value='" + loggedinid + "' />" +
		"</filter>" +
		"</link-entity>" +
		"</link-entity>" +
		"</link-entity>";
	filteri += filterincidenttypeHelper(linkQuery);

	//Get N:1
	linkQuery = "  <link-entity name='msdyn_organizationalunit' from='msdyn_organizationalunitid' to='duc_organizationalunitid' link-type='inner' alias='ac'>" +
		" <link-entity name='systemuser' from='duc_organizationalunitid' to='msdyn_organizationalunitid' link-type='inner' alias='ad'>" +
		"  <filter type='and'>" +
		"   <condition attribute='systemuserid' operator='eq' uitype='systemuser' value='" + loggedinid + "' />" +
		" </filter>" +
		"</link-entity>" +
		"</link-entity>";
	filteri += filterincidenttypeHelper(linkQuery);
	//filteri += filterincidenttypeHelper("duc_organizationunitsid", "msdyn_incidenttypeid");
	//filteri += filterincidenttypeHelper("msdyn_organizationalunitid", "duc_organizationalunitid");

	filteri += "</condition></filter>";
	formContextCallback.getControl("msdyn_primaryincidenttype").addCustomFilter(filteri);
}
function filterincidenttypeHelper(linkQuery) {
	var filter = "";
	var fetcxml = "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>" +
		"<entity name='msdyn_incidenttype'>" +
		"<attribute name='msdyn_incidenttypeid' />" +


		"<order attribute='msdyn_name' descending='false' />" +
		linkQuery +
		"</entity>" +
		"</fetch>";

	var globalcontext = Xrm.Utility.getGlobalContext();
	var clientUrl = globalcontext.getClientUrl();
	clientUrl = clientUrl + "/api/data/v9.2/msdyn_incidenttypes?fetchXml=" + fetcxml;
	var $ = parent.$;
	$.ajax({
		url: clientUrl,
		type: "GET",
		contentType: "application/json;charset=utf-8",
		datatype: "json",
		async: false,
		beforeSend: function (Xmlhttprequest) { Xmlhttprequest.setRequestHeader("Accept", "application/json"); },
		success: function (result, textStatus, xhr) {

			if (result.value.length > 0) {
				for (var i = 0; i < result.value.length; i++) {
					filter += "<value>{" + result.value[i].msdyn_incidenttypeid + "}</value>";
				}
			}
			else {
				filter += "<value>{00000000-0000-0000-0000-000000000000}</value>";
			}
			return filter;
		},
		error: function (error) {
			console.log("Error retrieving data from campaign");
			filter += "<value>{00000000-0000-0000-0000-000000000000}</value>";
			return filter;
		}
	});
	return filter;
}
var formContextCallback = null;
function setPreSearch(executionContext) {
	formContextCallback = executionContext.getFormContext();
	formContextCallback.getControl("new_campaign").addPreSearch(filtercampaign);
	formContextCallback.getControl("msdyn_primaryincidenttype").addPreSearch(filterincidenttype);
}


function onchangeincidentype(executionContext) {
	formContextCallback = executionContext.getFormContext();
	if (formContextCallback.getAttribute("msdyn_primaryincidenttype").getValue() != null)
		formContextCallback.getControl("msdyn_workordertype").setDisabled(true);
	else
		formContextCallback.getControl("msdyn_workordertype").setDisabled(false);

}


function setclienttype(e) {
	var f = e.getFormContext();
	if (f.ui.getFormType() == 1) {
		var ct = Xrm.Utility.getGlobalContext().client.getClient();
		if (ct == "Mobile") {
			f.getAttribute("duc_clientdevice").setValue(1);
		}
	}
}

function setDepartmentFromIncidentType(executionContext) {
	var formContext = executionContext.getFormContext();
	var incidentTypeLookup = formContext.getAttribute("msdyn_primaryincidenttype").getValue();

	// 1. Check if we are on a Create form (Form Type 1)
	if (formContext.ui.getFormType() !== 1) {
		return; // Only run on creation
	}

	if (incidentTypeLookup != null && incidentTypeLookup[0] != null) {
		var incidentTypeId = incidentTypeLookup[0].id.replace("{", "").replace("}", "");
		var departmentField = formContext.getAttribute("duc_department");

		// 2. Clear department field if needed and check if incident type has an ID
		if (departmentField == null) return;
		departmentField.setValue(null); // Clear previous value

		// 3. Fetch the duc_organizationalunitid lookup value from the Incident Type (msdyn_incidenttype)
		Xrm.WebApi.online.retrieveMultipleRecords("msdyn_incidenttype",
			"?$select=_duc_organizationalunitid_value" +
			"&$expand=duc_organizationalunitid($select=msdyn_organizationalunitid,msdyn_name)" +
			"&$filter=msdyn_incidenttypeid eq " + incidentTypeId)
			.then(
				function success(results) {
					if (results != null && results.entities.length > 0) {
						var incidentTypeRecord = results.entities[0];
						var orgUnitLookup = incidentTypeRecord["duc_organizationalunitid"];

						if (orgUnitLookup != null) {
							// 4. Construct the lookup object for the duc_department field
							var departmentValue = [];
							departmentValue[0] = {
								id: orgUnitLookup.msdyn_organizationalunitid,
								name: orgUnitLookup.msdyn_name,
								entityType: "msdyn_organizationalunit" // Assuming this is the target entity type
							};

							// 5. Set the Work Order's duc_department field
							departmentField.setValue(departmentValue);
						}
					}
				},
				function (error) {
					console.error("Error retrieving Organizational Unit: " + error.message);
				}
			);
	} else {
		// If incident type is cleared, clear the department field
		formContext.getAttribute("duc_department").setValue(null);
	}
}



// =======================================
// AUTO REFRESH ON BACKEND STATUS CHANGE
// =======================================

var autoRefreshInterval = null;

function startStatusMonitor(executionContext) {
	const formContext = executionContext.getFormContext();

	// Avoid creating multiple intervals
	if (autoRefreshInterval !== null) {
		return;
	}

	autoRefreshInterval = setInterval(async function () {
		try {
			const workOrderId = formContext.data.entity.getId().replace(/[{}]/g, "");
			if (!workOrderId) return;

			// 1) Get FRONTEND value
			const frontStatus = formContext.getAttribute("msdyn_systemstatus")?.getValue();

			// 2) Query the BACKEND value
			const backend = await Xrm.WebApi.retrieveRecord(
				"msdyn_workorder",
				workOrderId,
				"?$select=msdyn_systemstatus"
			);

			const backendStatus = backend.msdyn_systemstatus;

			// 3) Compare → if different, refresh
			if (backendStatus !== frontStatus) {
				console.log("Status changed in backend. Refreshing form...");
				formContext.data.refresh(true);
			}

		} catch (err) {
			console.error("Status monitor failed:", err);
		}
	}, 5000); // 5 seconds
}


// OPTIONAL: call this on form OnLoad
function onLoad2(executionContext) {
	startStatusMonitor(executionContext);
}

function controlVisibilityOnMobile(executionContext) {
	var formContext = executionContext.getFormContext();
	var client = Xrm.Utility.getGlobalContext().client;
	var clientType = client.getClient();        // "Web", "Mobile", "Outlook"
	var formFactor = client.getFormFactor();    // 1=Desktop,2=Tablet,3=Phone

	var isMobile = (formFactor === 2 || formFactor === 3 || clientType === "Mobile");

	// Tabs to hide on mobile
	var tabNames = [
		"Inspection_action_tab",
		"tab_14",
		"ActionLogs",
		"tab_10"
	];

	for (var i = 0; i < tabNames.length; i++) {
		var tab = formContext.ui.tabs.get(tabNames[i]);
		if (tab) {
			tab.setVisible(!isMobile);  // hide on mobile, show on web
		}
	}

	// Section bookingcard_section under General_Tab
	var generalTab = formContext.ui.tabs.get("General_Tab");
	if (generalTab) {
		var sections = generalTab.sections;
		var bookingSection = sections.get("bookingcard_section");
		if (bookingSection) {
			bookingSection.setVisible(!isMobile);  // hide on mobile, show on web
		}
	}
}
