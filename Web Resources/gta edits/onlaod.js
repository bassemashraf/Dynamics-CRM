function controlVisibilityOnMobile(executionContext) {
	var formContext = executionContext.getFormContext();
	var client = Xrm.Utility.getGlobalContext().client;
	var clientType = client.getClient();        // "Web", "Mobile", "Outlook"
	var formFactor = client.getFormFactor();    // 1=Desktop,2=Tablet,3=Phone

	var isMobile = (formFactor === 2 || formFactor === 3 || clientType === "Mobile");

	// Tabs to hide on mobile
	var tabNames = [
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

	// Hide Inspection_action_tab only if duc_primaryinspectionaction doesn't have data and isMobile is true
	var inspectionActionTab = formContext.ui.tabs.get("Inspection_action_tab");
	if (inspectionActionTab) {
		var primaryInspectionAction = formContext.getAttribute("duc_primaryinspectionaction");
		var hasData = primaryInspectionAction && primaryInspectionAction.getValue();
		var shouldHide = isMobile && !hasData;
		inspectionActionTab.setVisible(!shouldHide);
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

//on file  duc_workorder.js should appl on mecc too 
