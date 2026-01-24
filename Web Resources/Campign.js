//#region From duc_advancedfind.js
function Onload(executionContext) {
    

    var formContext = executionContext.getFormContext();
    var inspectionCampaignId = formContext.data.entity.getId();

    if (inspectionCampaignId != null) {
        inspectionCampaignId = inspectionCampaignId.replace('{', '').replace('}', '');
        sessionStorage.setItem('InspectionCampaignId', inspectionCampaignId);
    }

    var incidentType = formContext.getAttribute("duc_incidenttype").getValue();
    if (incidentType != null) {
        var incidentTypeId = incidentType[0].id;
        sessionStorage.setItem('incidentTypeId', incidentTypeId);
    }

    var accountType = false; //Commercial
    var workOrderType = formContext.getAttribute("duc_workordertype").getValue();
    if (workOrderType != null) {
        var workOrderTypeId = workOrderType[0].id;
        sessionStorage.setItem('workOrderTypeId', workOrderTypeId);

        var _accountType = GetWOType(workOrderTypeId);
    }

    //----------------------------------------------19/12---------//
    var campaignTypeAttr = formContext.getAttribute("duc_campaigntype");
    var campaignInternalTypeAttr = formContext.getAttribute("duc_campaigninternaltype");
    var organizationalUnitAttr = formContext.getAttribute("duc_organizationalunitid");

    if (campaignTypeAttr != null && campaignInternalTypeAttr != null) {

        if (campaignTypeAttr.getValue() == 100000004 && campaignInternalTypeAttr.getValue() == 100000005) //Main Patrol
        {
            var tab = formContext.ui.tabs.get("inspection_campaigns");
            tab.setVisible(true);
        }
    }

    // Check organizational unit and remove option if needed
    if (organizationalUnitAttr != null && organizationalUnitAttr.getValue() != null) {
        var organizationalUnitLookup = organizationalUnitAttr.getValue();
        var organizationalUnitId = organizationalUnitLookup[0].id;

        // Retrieve the organizational unit record to check duc_englishname
        Xrm.WebApi.retrieveRecord("msdyn_organizationalunit", organizationalUnitId, "?$select=duc_englishname").then(
            function success(result) {
                if (result.duc_englishname != "Inspection Section – Natural Reserves") {
                    // Remove option 100000004 from duc_campaigntype
                    if (campaignTypeAttr != null) {
                        var campaignTypeControl = campaignTypeAttr.controls.get(0);
                        campaignTypeControl.removeOption(100000004);
                    }
                }
            },
            function (error) {
                console.log("Error retrieving organizational unit: " + error.message);
            }
        );
    }

    // Section visibility logic for Patrol_Section
    
    var tab = formContext.ui.tabs.get("{dfbcbc11-d392-495f-a96a-11c14d55af9e}")
    var sections = tab.sections;
    var patrolSection = sections.get("Patrol_Section");

    if (campaignTypeAttr != null && campaignInternalTypeAttr != null) {
        var campaignTypeValue = campaignTypeAttr.getValue();
        var campaignInternalTypeValue = campaignInternalTypeAttr.getValue();

        // Case 1: campaignInternalType == 100000005 AND campaignType == 100000004
        if (campaignInternalTypeValue == 100000005 && campaignTypeValue == 100000004) {
            console.log("Case 1: Showing patrol section with both subgrids");

            if (patrolSection != null) {
                patrolSection.setVisible(true);
            }

            var campaignLogGrid = formContext.getControl("Campaign_Log_Grid");
            if (campaignLogGrid != null) {
                campaignLogGrid.setVisible(true);
            }

            var subgridNew5 = formContext.getControl("Subgrid_new_5");
            if (subgridNew5 != null) {
                subgridNew5.setVisible(true);
            }
        }
        // Case 2: campaignType == 100000004 AND campaignInternalType == 100000004
        else if (campaignTypeValue == 100000004 && campaignInternalTypeValue == 100000004) {
            console.log("Case 2: Showing patrol section, hiding Subgrid_new_5");

            if (patrolSection != null) {
                patrolSection.setVisible(true);
            }

            var campaignLogGrid = formContext.getControl("Campaign_Log_Grid");
            if (campaignLogGrid != null) {
                campaignLogGrid.setVisible(true);
            }

            var subgridNew5 = formContext.getControl("Subgrid_new_5");
            if (subgridNew5 != null) {
                subgridNew5.setVisible(false);
            }
        }
        // Any other case: hide the section
        else {
            console.log("Other case: Hiding patrol section");
            if (patrolSection != null) {
                patrolSection.setVisible(false);
            }
        }
    }
    else {
        console.log("Attributes are null: Hiding patrol section");
        // If attributes are null, hide the section
        if (patrolSection != null) {
            patrolSection.setVisible(false);
        }
    }

}

function GetWOType(woTypeId) {
    Xrm.WebApi.retrieveRecord("msdyn_workordertype", woTypeId, "?$select=duc_type").then(
        function success(result) {
            
            var msdyn_workordertypeid = result["msdyn_workordertypeid"];
            var duc_type = result["duc_type"];

            sessionStorage.setItem('accountType', duc_type);

            return duc_type;
        },
        function (error) {
            console.log(error.message);
        }
    );
}
//#endregion

//#region From duc_Inspection_Campaign_KPI.js
function OnChangeKPIsTab(executionContext) {
    
    RefreshKPIs(executionContext);
    waitForElementsToExist(['[data-id*="COMPLETED_INSPECTIONS_section"]'], () => KPIStyle(), { checkFrequency: 4000, timeout: 10000 });
}

function waitForElementsToExist(elementIds, callback, options) {

    // poll every X amount of ms for all DOM nodes
    let intervalHandle = setInterval(() => {
        let doElementsExist = true;
        for (let elementId of elementIds) {
            let element = window.top.document.querySelector(elementId) //window.top.document.getElementById(elementId);
            if (!element) {
                // if element does not exist, set doElementsExist to false and stop the loop
                doElementsExist = false;
                break;
            }
        }
        // if all elements exist, stop polling and invoke the callback function 
        if (doElementsExist) {
            clearInterval(intervalHandle);
            if (callback) {
                callback();
            }
        }
    }, options.checkFrequency);

    if (options.timeout != null) {
        setTimeout(() => clearInterval(intervalHandle), options.timeout);
    }
}

function KPIStyle() {

    const badgeControls = Array.from(parent.document.getElementsByClassName("fy9rknc"));

    badgeControls.forEach(b => {

        if (b.nodeName === "DIV") {
            b.style.fontSize = "24px";
            b.style.width = "90%";
            b.style.height = "50px";
        }
    });

    AlignRollupLastUpdated();
}

function AlignRollupLastUpdated() {

    const NumberOfInspection_Label = parent.document.querySelector('[data-id="duc_noofinspections-calculatedLabel"]');
    if (NumberOfInspection_Label != null)
        NumberOfInspection_Label.style.paddingTop = "15px";

    const NumberOfInspection_Date = parent.document.querySelector('[data-id="duc_noofinspections-calculatedDateLabel"]');
    if (NumberOfInspection_Date != null)
        NumberOfInspection_Date.style.paddingTop = "15px";

    const scheduledInspections_Label = parent.document.querySelector('[data-id="duc_scheduledinspections-calculatedLabel"]');
    if (scheduledInspections_Label != null)
        scheduledInspections_Label.style.paddingTop = "15px";

    const scheduledInspections_Date = parent.document.querySelector('[data-id="duc_scheduledinspections-calculatedDateLabel"]');
    if (scheduledInspections_Date != null)
        scheduledInspections_Date.style.paddingTop = "15px";

    const completedInspections_Label = parent.document.querySelector('[data-id="duc_completedinspections-calculatedLabel"]');
    if (completedInspections_Label != null)
        completedInspections_Label.style.paddingTop = "15px";

    const completedInspections_Date = parent.document.querySelector('[data-id="duc_completedinspections-calculatedDateLabel"]');
    if (completedInspections_Date != null)
        completedInspections_Date.style.paddingTop = "15px";
}

async function RefreshKPIs(executionContext) {
    
    var formContext = executionContext.getFormContext();

    var entityName = 'new_inspectioncampaigns'; // Name of Entity 
    var entityId = formContext.data.entity.getId().replace('{', '').replace('}', ''); // GUID of current record

    if (entityId != null && entityId != "") {
        var noOfInspections = 'duc_noofinspections';
        var scheduledInspections = 'duc_scheduledinspections';
        var completedInspections = 'duc_completedinspections';

        // Function Call to Update Rollup Field Values
        refreshRoleupField(formContext, entityName, entityId, noOfInspections);
        refreshRoleupField(formContext, entityName, entityId, scheduledInspections);
        refreshRoleupField(formContext, entityName, entityId, completedInspections);

        var recalculate = formContext.getAttribute("duc_recalculate");
        var statecode = formContext.getAttribute("statecode").getValue();

        // if (recalculate != undefined && statecode == 0) {

        //     //get tab
        //     let tabObj = formContext.ui.tabs.get("{dfbcbc11-d392-495f-a96a-11c14d55af9e}");
        //     //get section
        //     let sectionObj = tabObj.sections.get("Refresh_KPIs_section1");

        //     //show the section
        //     sectionObj.setVisible(true);

        //     recalculate.setValue(true);
        //     formContext.data.entity.save(1);

        //     //hide the section
        //     sectionObj.setVisible(false);

        //     //formContext.data.refresh(true);
        // }

        try {
            const accCount = await getDistinctAccountsCount(entityId);
            const currentAcc = formContext.getAttribute("duc_noofaccounts").getValue();
            if (currentAcc !== accCount) {
                formContext.getAttribute("duc_noofaccounts").setValue(accCount);
                formContext.getAttribute("duc_noofaccounts").setSubmitMode("always");
            }
        } catch (e) {
            console.error(e);
        }

        try {
            const inspCount = await getDistinctUserBookableResourceIds(entityId);
            const currentInsp = formContext.getAttribute("duc_noofinspectores").getValue();
            if (currentInsp !== inspCount) {
                formContext.getAttribute("duc_noofinspectores").setValue(inspCount);
                formContext.getAttribute("duc_noofinspectores").setSubmitMode("always");
            }
        } catch (e) {
            console.error(e);
        }

        try {
            const iaCount = await getTotalInspectionActions(entityId);
            const currentAcc = formContext.getAttribute("duc_noofinspectionactions").getValue();
            if (currentAcc !== iaCount) {
                formContext.getAttribute("duc_noofinspectionactions").setValue(iaCount);
                formContext.getAttribute("duc_noofinspectionactions").setSubmitMode("always");
            }
        } catch (e) {
            console.error(e);
        }

        if (formContext.data.getIsDirty()) {
            formContext.data.save();
        }
    }
}

function refreshRoleupField(formContext, entityName, entityId, rollup_fieldName) {
    
    //var formContext = executionContext.getFormContext();
    var clientUrl = formContext.context.getClientUrl();

    // Method Calling and defining parameter
    var rollupAPIMethod = "/api/data/v9.2/CalculateRollupField(Target=@recordID,FieldName=@field_Name)";

    // Passing Parameter Values
    rollupAPIMethod += "?@recordID={'@odata.id':'" + entityName + "(" + entityId + ")'}&@field_Name='" + rollup_fieldName + "'";

    var req = new XMLHttpRequest();
    req.open("GET", clientUrl + rollupAPIMethod, false);
    req.onreadystatechange = function () {
        if (this.readyState === 4) {
            req.onreadystatechange = null;
            if (this.status === 200) {
                console.log("Field " + rollup_fieldName + " Recalculated successfully");
            }
        }
    };

    req.send();

}


async function getDistinctAccountsCount(parentCampaignId) {

    async function getChildCampaignIds(parentId) {
        const fetchXml = [
            "<fetch mapping='logical'>",
            "<entity name='new_inspectioncampaign'>",
            "<attribute name='new_inspectioncampaignid' />",
            "<filter>",
            `<condition attribute='duc_parentcampaign' operator='eq' value='${parentId}' />`,
            "</filter>",
            "</entity>",
            "</fetch>"
        ].join("");

        let records = [];
        let page = 1;
        let pagingCookie = null;
        let hasMore = true;

        while (hasMore) {
            let fetchWithPaging = fetchXml.replace("<fetch", `<fetch page='${page}' count='5000'`);
            if (pagingCookie) {
                fetchWithPaging = fetchWithPaging.replace("<fetch", `<fetch page='${page}' count='5000' paging-cookie='${pagingCookie.replace(/"/g, "&quot;")}'`);
            }
            let response = await Xrm.WebApi.online.retrieveMultipleRecords(
                "new_inspectioncampaign",
                `?fetchXml=${encodeURIComponent(fetchWithPaging)}`
            );

            records = records.concat(response.entities);

            if (response["@odata.nextLink"]) {
                page++;
                pagingCookie = response["@Microsoft.Dynamics.CRM.fetchxmlpagingcookie"];
            } else {
                hasMore = false;
            }
        }
        return records.map(r => r.new_inspectioncampaignid);
    }

    async function getAccountIdsForCampaign(campaignId) {
        const fetchXml = [
            "<fetch mapping='logical'>",
            "<entity name='new_new_inspectioncampaign_account'>",
            "<attribute name='accountid' />",
            "<filter>",
            `<condition attribute='new_inspectioncampaignid' operator='eq' value='${campaignId}' />`,
            "</filter>",
            "</entity>",
            "</fetch>"
        ].join("");

        let accountIds = [];
        let page = 1;
        let pagingCookie = null;
        let hasMore = true;

        while (hasMore) {
            let fetchWithPaging = fetchXml.replace("<fetch", `<fetch page='${page}' count='5000'`);
            if (pagingCookie) {
                fetchWithPaging = fetchWithPaging.replace("<fetch", `<fetch page='${page}' count='5000' paging-cookie='${pagingCookie.replace(/"/g, "&quot;")}'`);
            }
            let response = await Xrm.WebApi.online.retrieveMultipleRecords(
                "new_new_inspectioncampaign_account",
                `?fetchXml=${encodeURIComponent(fetchWithPaging)}`
            );

            accountIds = accountIds.concat(response.entities.map(e => e.accountid));

            if (response["@odata.nextLink"]) {
                page++;
                pagingCookie = response["@Microsoft.Dynamics.CRM.fetchxmlpagingcookie"];
            } else {
                hasMore = false;
            }
        }
        return accountIds;
    }

    let allCampaignIds = [parentCampaignId];

    const childCampaignIds = await getChildCampaignIds(parentCampaignId);

    allCampaignIds = allCampaignIds.concat(childCampaignIds);

    const distinctAccountIds = new Set();

    for (const campaignId of allCampaignIds) {
        const accounts = await getAccountIdsForCampaign(campaignId);

        accounts.forEach(id => distinctAccountIds.add(id));
    }

    return distinctAccountIds.size;
}

async function getDistinctUserBookableResourceIds(parentCampaignId) {

    async function getChildCampaignIds(parentId) {
        const fetchXml = [
            "<fetch mapping='logical'>",
            "<entity name='new_inspectioncampaign'>",
            "<attribute name='new_inspectioncampaignid' />",
            "<filter>",
            `<condition attribute='duc_parentcampaign' operator='eq' value='${parentId}' />`,
            "</filter>",
            "</entity>",
            "</fetch>"
        ].join("");

        let records = [];
        let page = 1;
        let pagingCookie = null;
        let hasMore = true;

        while (hasMore) {
            let fetchWithPaging = fetchXml.replace("<fetch", `<fetch page='${page}' count='5000'`);
            if (pagingCookie) {
                fetchWithPaging = fetchWithPaging.replace("<fetch", `<fetch page='${page}' count='5000' paging-cookie='${pagingCookie.replace(/"/g, "&quot;")}'`);
            }
            let response = await Xrm.WebApi.online.retrieveMultipleRecords(
                "new_inspectioncampaign",
                `?fetchXml=${encodeURIComponent(fetchWithPaging)}`
            );
            records = records.concat(response.entities);

            if (response["@odata.nextLink"]) {
                page++;
                pagingCookie = response["@Microsoft.Dynamics.CRM.fetchxmlpagingcookie"];
            } else {
                hasMore = false;
            }
        }
        return records.map(r => r.new_inspectioncampaignid);
    }

    async function getUserBookableResourceIdsForCampaign(campaignId) {
        const fetchXml = [
            "<fetch mapping='logical'>",
            "<entity name='new_inspectioncampaign_bookableresource'>",
            "<attribute name='bookableresourceid' />",
            "<link-entity name='bookableresource' from='bookableresourceid' to='bookableresourceid' alias='br'>",
            "<filter>",
            "<condition attribute='resourcetype' operator='eq' value='3' />", // User type = 3
            "</filter>",
            "</link-entity>",
            "<filter>",
            `<condition attribute='new_inspectioncampaignid' operator='eq' value='${campaignId}' />`,
            "</filter>",
            "</entity>",
            "</fetch>"
        ].join("");

        let resourceIds = [];
        let page = 1;
        let pagingCookie = null;
        let hasMore = true;

        while (hasMore) {
            let fetchWithPaging = fetchXml.replace("<fetch", `<fetch page='${page}' count='5000'`);
            if (pagingCookie) {
                fetchWithPaging = fetchWithPaging.replace("<fetch", `<fetch page='${page}' count='5000' paging-cookie='${pagingCookie.replace(/"/g, "&quot;")}'`);
            }
            let response = await Xrm.WebApi.online.retrieveMultipleRecords(
                "new_inspectioncampaign_bookableresource",
                `?fetchXml=${encodeURIComponent(fetchWithPaging)}`
            );
            resourceIds = resourceIds.concat(response.entities.map(e => e.bookableresourceid));

            if (response["@odata.nextLink"]) {
                page++;
                pagingCookie = response["@Microsoft.Dynamics.CRM.fetchxmlpagingcookie"];
            } else {
                hasMore = false;
            }
        }
        return resourceIds;
    }

    let allCampaignIds = [parentCampaignId];

    const childCampaignIds = await getChildCampaignIds(parentCampaignId);

    allCampaignIds = allCampaignIds.concat(childCampaignIds);

    const distinctResourceIds = new Set();

    for (const campaignId of allCampaignIds) {
        const resourceIds = await getUserBookableResourceIdsForCampaign(campaignId);

        resourceIds.forEach(id => distinctResourceIds.add(id));
    }

    return distinctResourceIds.size;
}

async function getTotalInspectionActions(parentCampaignId) {

    async function getChildCampaignIds(parentId) {
        const fetchXml = [
            "<fetch mapping='logical'>",
            "<entity name='new_inspectioncampaign'>",
            "<attribute name='new_inspectioncampaignid' />",
            "<filter>",
            `<condition attribute='duc_parentcampaign' operator='eq' value='${parentId}' />`,
            "</filter>",
            "</entity>",
            "</fetch>"
        ].join("");

        let records = [];
        let page = 1;
        let pagingCookie = null;
        let hasMore = true;

        while (hasMore) {
            let fetchWithPaging = fetchXml.replace("<fetch", `<fetch page='${page}' count='5000'`);
            if (pagingCookie) {
                fetchWithPaging = fetchWithPaging.replace("<fetch", `<fetch page='${page}' count='5000' paging-cookie='${pagingCookie.replace(/"/g, "&quot;")}'`);
            }
            let response = await Xrm.WebApi.online.retrieveMultipleRecords(
                "new_inspectioncampaign",
                `?fetchXml=${encodeURIComponent(fetchWithPaging)}`
            );
            records = records.concat(response.entities);

            if (response["@odata.nextLink"]) {
                page++;
                pagingCookie = response["@Microsoft.Dynamics.CRM.fetchxmlpagingcookie"];
            } else {
                hasMore = false;
            }
        }
        return records.map(r => r.new_inspectioncampaignid);
    }

    async function getInspectionActionCountForCampaign(campaignId) {
        const fetchXml = [
            "<fetch mapping='logical'>",
            "<entity name='duc_inspectionaction'>",
            "<attribute name='duc_inspectionactionid' />",
            "<link-entity name='msdyn_workorder' from='msdyn_workorderid' to='duc_relatedworkorderid' alias='wo'>",
            `<filter><condition attribute='new_campaign' operator='eq' value='${campaignId}' /></filter>`,
            "</link-entity>",
            "</entity>",
            "</fetch>"
        ].join("");

        let count = 0;
        let page = 1;
        let pagingCookie = null;
        let hasMore = true;

        while (hasMore) {
            let fetchWithPaging = fetchXml.replace("<fetch", `<fetch page='${page}' count='5000'`);
            if (pagingCookie) {
                fetchWithPaging = fetchWithPaging.replace(
                    "<fetch",
                    `<fetch page='${page}' count='5000' paging-cookie='${pagingCookie.replace(/"/g, "&quot;")}'`
                );
            }

            const response = await Xrm.WebApi.online.retrieveMultipleRecords(
                "duc_inspectionaction",
                `?fetchXml=${encodeURIComponent(fetchWithPaging)}`
            );

            count += response.entities.length;

            if (response["@odata.nextLink"]) {
                page++;
                pagingCookie = response["@Microsoft.Dynamics.CRM.fetchxmlpagingcookie"];
            } else {
                hasMore = false;
            }
        }

        return count;
    }

    let allCampaignIds = [parentCampaignId];

    const childCampaignIds = await getChildCampaignIds(parentCampaignId);

    allCampaignIds = allCampaignIds.concat(childCampaignIds);

    let totalInspectionActions = 0;

    for (const campaignId of allCampaignIds) {
        const cnt = await getInspectionActionCountForCampaign(campaignId);

        totalInspectionActions += cnt;
    }

    return totalInspectionActions;
}

var inspectionCampaignStatusInterval = null;
function refreshFormOnStatusChange(executionContext) {
    const formContext = executionContext.getFormContext();

    if (inspectionCampaignStatusInterval !== null) {
        return;
    }

    inspectionCampaignStatusInterval = setInterval(function () {
        checkInspectionCampaignStatus(formContext);
    }, 5000);

    if (formContext.ui.getFormType() === 1) return;
}

async function checkInspectionCampaignStatus(formContext) {
    try {
        const recordId = formContext.data.entity.getId();

        if (!recordId) return;

        const cleanId = recordId.replace(/[{}]/g, "");

        const frontStatus = formContext.getAttribute("duc_campaignstatus")?.getValue();

        const backendRecord = await Xrm.WebApi.retrieveRecord(
            "new_inspectioncampaign",
            cleanId,
            "?$select=duc_campaignstatus"
        );

        const backendStatus = backendRecord.duc_campaignstatus;

        if (backendStatus !== frontStatus) {
            formContext.data.refresh(true);
        }

    } catch (e) {
        console.error("Status check failed:", e);
    }
}
//#endregion

//#region From duc_InspectioncampaignLibrary
var controlName = "duc_incidenttype";
function setPreSearch(executionContext) {
    formContextCallback = executionContext.getFormContext();
    var incidentTypeC = formContextCallback.getControl(controlName);
    if (incidentTypeC != null)
        incidentTypeC.addPreSearch(filterincidenttype);
}

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
    formContextCallback.getControl(controlName).addCustomFilter(filteri);
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

function setorgunit(executionContext) {
    fc = executionContext.getFormContext();
    var dept = fc.getAttribute("duc_organizationalunitid").getValue();
    if (dept == null || dept == '') {
        

        var parentCampaginAttr = fc.getAttribute("duc_parentcampagin");
        var parentCampaign = parentCampaginAttr ? parentCampaginAttr.getValue() : null;
        if (parentCampaign !== null && parentCampaign !== '') {
            var descriptionAttr = fc.getAttribute("duc_campaigndescription");
            var description = descriptionAttr ? descriptionAttr.getValue() : null;
            if (description.indexOf("Auto Generated") === -1) {
                return;
            }
        }

        var userid = Xrm.Utility.getGlobalContext().userSettings.userId;
        Xrm.WebApi.retrieveRecord("systemuser", userid.replace("{", "").replace("}", ""), "?$select=_duc_organizationalunitid_value").then(
            function success(result) {
                var id = result._duc_organizationalunitid_value;
                var obj = new Array();
                obj[0] = new Object();
                obj[0].id = id;
                obj[0].name = result["_duc_organizationalunitid_value@OData.Community.Display.V1.FormattedValue"];
                obj[0].entityType = "msdyn_organizationalunit";
                fc.getAttribute("duc_organizationalunitid").setValue(obj);
            },

            function (error) {
                console.log(error.message);
            }
        );
    }
}

function ValidateFromDate(executionContext) {
    var fromDateMsg = "";
    var userSettings = Xrm.Utility.getGlobalContext().userSettings;

    var userLanguage = userSettings.languageId; // (1025 : ar-SA  && 1033: en-US)
    if (userLanguage != null && userLanguage == 1033) {
        fromDateMsg = "Date can not be less than today";
    }
    else if (userLanguage != null && userLanguage == 1025) {
        fromDateMsg = "لا يمكن اختيار تاريخ في وقت سابق";
    }

    fc = executionContext.getFormContext();
    var fromDate = fc.getAttribute('duc_fromdate').getValue();
    if (fromDate != null) {
        fromDate.setHours(0, 0, 0, 0);
        // rest of the code
        var today = new Date();
        today.setHours(0, 0, 0, 0);
        if (fromDate < today) {
            alert(fromDateMsg);
            fc.getAttribute('duc_fromdate').setValue();
        }
    }
}

function ValidateToDate(executionContext) {
    var toDateMsg = "";
    var userSettings = Xrm.Utility.getGlobalContext().userSettings;

    var userLanguage = userSettings.languageId; // (1025 : ar-SA  && 1033: en-US)
    if (userLanguage != null && userLanguage == 1033) {
        toDateMsg = "To Date can not be less than from Date";
    }
    else if (userLanguage != null && userLanguage == 1025) {
        toDateMsg = "لا يمكن اختيار تاريخ لنهايةالحملة قبل تاريخ بدء الحملة";
    }

    fc = executionContext.getFormContext();
    var fromDate = fc.getAttribute('duc_fromdate').getValue();
    var toDate = fc.getAttribute('duc_todate').getValue();
    if (fromDate != null && toDate != null) {
        fromDate.setHours(0, 0, 0, 0);
        toDate.setHours(0, 0, 0, 0);
        if (fromDate > toDate) {
            alert(toDateMsg);
            fc.getAttribute('duc_todate').setValue();
        }
    }
}

function showCancelCampaign(formContext) {
    var userId = Xrm.Utility.getGlobalContext().getUserId();

    var OUid;
    if (formContext.getAttribute("duc_organizationalunitid").getValue() != null) {
        OUid = formContext.getAttribute("duc_organizationalunitid").getValue()[0].id;
    }
    var fetchXML = "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>" +
        "<entity name='systemuser'>" +
        "<attribute name='systemuserid' />" +
        "<order attribute='fullname' descending='false' />" +
        "<filter type='and'>" +
        "<condition attribute='systemuserid' operator='eq' value='" + userId + "'  />" +
        "</filter>" +
        "<link-entity name='duc_systemuser_approvalpositions' from='systemuserid' to='systemuserid' visible='false' intersect='true'>" +
        "<link-entity name='position' from='positionid' to='positionid' alias='ah'>" +
        "<filter type='and'>" +
        "<condition attribute='duc_cancancelcampaign' operator='eq' value='1' />" +
        "</filter>" +
        "</link-entity>" +
        "</link-entity>" +
        "<link-entity name='duc_systemuser_msdyn_approvalorgunit' from='systemuserid' to='systemuserid' visible='false' intersect='true'>" +
        "<link-entity name='msdyn_organizationalunit' from='msdyn_organizationalunitid' to='msdyn_organizationalunitid' alias='ai'>" +
        "<filter type='and'>" +
        "<condition attribute='msdyn_organizationalunitid' operator='eq' value='" + OUid + "' />" +
        "</filter>" +
        "</link-entity>" +
        "</link-entity>" +
        "</entity>" +
        "</fetch>";
    var result = GetFromRestAPI("systemusers", null, null, null, fetchXML);
    if (result != null && result.value != null && result.value.length > 0 && result.value[0].systemuserid != null) {
        return true;
    }
    return false;
}

function showModifyCampaign(formContext) {
    
    return showCancelCampaign(formContext);
}

function showRestartCampaign(formContext) {
    return showCancelCampaign(formContext);
}
//#endregion

//#region From duc_commoninspectioncampaignscirpts
function switchFormOnLoad(executionContext) {
    var formContext = executionContext.getFormContext();
    var campaignType = formContext.getAttribute("duc_deparmentsadvancedfind")?.getValue();

    if (campaignType == null) {
        console.log("Campaign Type is empty — skipping form switch.");
        return;
    }

    // Map option set values to form names
    var formMap = {
        100000000: "Sub Periodic Campaign Form", // form label in Dynamics
        100000001: "Sub Joint Campaign Form",
        100000002: "Main Periodic Campaign Form",
        100000003: "Main Joint Campaign Form"
    };

    var targetFormName = formMap[campaignType];
    if (!targetFormName) {
        console.log("No target form mapped for value: " + campaignType);
        return;
    }

    var currentForm = formContext.ui.formSelector.getCurrentItem();
    var currentFormName = currentForm.getLabel();

    // Avoid loops
    if (currentFormName === targetFormName) {
        console.log("Already on the correct form: " + targetFormName);
        return;
    }

    // Switch to the target form
    var allForms = formContext.ui.formSelector.items.get();
    for (var i = 0; i < allForms.length; i++) {
        if (allForms[i].getLabel() === targetFormName) {
            console.log("Switching from " + currentFormName + " to " + targetFormName);
            allForms[i].navigate();
            return;
        }
    }

    console.log("Target form not found: " + targetFormName);
}


function openDepartmentSelector(primaryControl) {
    var campaignId = primaryControl.data.entity.getId().replace("{", "").replace("}", "");
    var pageInput = {
        pageType: "webresource",
        webresourceName: "duc_deparmentsadvancedfind", // your new HTML page
        data: campaignId
    };

    var navigationOptions = {
        target: 2, // open as dialog
        width: { value: 70, unit: "%" },
        height: { value: 90, unit: "%" },
        position: 1
    };

    Xrm.Navigation.navigateTo(pageInput, navigationOptions);
}


function handleCampaignTypeVisibility(executionContext) {

    var formContext = executionContext.getFormContext();
    var campaignTypeAttr = formContext.getAttribute("duc_campaigninternaltype");

    if (!campaignTypeAttr) {
        console.error("Attribute 'duc_campaigninternaltype' not found on form.");
        return;
    }

    var campaignType = campaignTypeAttr.getValue();

    // Skip if value is null (e.g., new record not yet selected)
    if (campaignType === null) {
        console.warn("Campaign type is not selected yet.");
        return;
    }

    // Mapping for reference or future logic (optional)
    var formMap = {
        100000000: "Sub Periodic Campaign Form",
        100000001: "Sub Joint Campaign Form",
        100000002: "Main Periodic Campaign Form",
        100000003: "Main Joint Campaign Form",
        100000005: "PatrolCampaign Form"

    };
    
    // Centralized tab visibility rules
    var tabRules = {
        Department_Tab: campaignType === 100000003,              // Show only for Main Joint Campaign
        related_campaign_tab: campaignType === 100000002,        // Show only for Main Periodic Campaign
        account_tab: (campaignType !== 100000000 && campaignType !== 100000005 && campaignType !== 100000004),        // Hide for Sub Periodic Campaign
        inspectors_tab: !(campaignType === 100000000), // Hide for Sub Perodic types
        Patrol_Logs_Tab: (campaignType == 100000005 || campaignType == 100000004),
    };

    // Apply tab visibility based on rules
    for (var tabName in tabRules) {
        var tab = formContext.ui.tabs.get(tabName);
        if (tab) {
            tab.setVisible(tabRules[tabName]);
        } else {

            var tab2 = formContext.ui.tabs.get("{dfbcbc11-d392-495f-a96a-11c14d55af9e}");
            var section = tab2.sections.get(tabName);
            if (section) {

                section.setVisible(tabRules[tabName]);

            }
            else {
                console.warn("Tab or section of name '" + tabName + "' not found on this form.");
            }
        }
    }
}

function toggleProcessAutomationSection(executionContext) {
    return;//TBD TODO
    var formContext = executionContext.getFormContext();

    // Get the Option Set value
    var statusValue = formContext.getAttribute("duc_campaignstatus").getValue();

    // Reference the tab and section
    // Replace 'general' with the tab name where this section exists
    var tab = formContext.ui.tabs.get("general");
    if (!tab) {
        console.error("Tab containing 'process_automation_section' not found.");
        return;
    }

    var section = tab.sections.get("process_automation_section");
    if (!section) {
        console.error("Section 'process_automation_section' not found on this form.");
        return;
    }

    // Show or hide the section
    if (statusValue === 100000001) {
        section.setVisible(true);
    } else {
        section.setVisible(false);
    }
}

function setNextRunDate(executionContext) {
    var formContext = executionContext.getFormContext();

    var periodType = formContext.getAttribute("duc_campaignperiodtype")?.getValue();
    var periodValue = formContext.getAttribute("duc_campaignperiodvalue")?.getValue() || 1;
    var nextRunDate = formContext.getAttribute("duc_nextrundate")?.getValue();
    var startDateValue = formContext.getAttribute("duc_fromdate")?.getValue();
    if (!periodType) {
        console.log("Missing next run date or period type — cannot calculate.");
        return;
    }

    // Convert to Date object
    var startDate = new Date(startDateValue);
    var newDate = new Date(startDate);

    // switch (periodType) {
    // case 100000000: // Daily
    // newDate.setDate(periodValue);
    // break;

    // case 100000001: // Weekly
    // newDate.setDate(startDate.getDate() + (periodValue * 7));
    // break;

    // case 100000002: // Monthly
    // newDate.setMonth(startDate.getMonth() + periodValue);
    // break;

    // case 100000003: // Quarterly
    // newDate.setMonth(startDate.getMonth() + 4);
    // break;

    // case 100000004: // Semi-Annually
    // newDate.setMonth(startDate.getMonth() + 6);
    // break;

    // case 100000005: // Annually
    // newDate.setFullYear(startDate.getFullYear() + 1);
    // break;

    // default:
    // console.log("Unknown period type: " + periodType);
    // return;
    // }

    // Update the field value
    formContext.getAttribute("duc_nextrundate").setValue(newDate);
    formContext.getAttribute("duc_nextrundate").setSubmitMode("always");

    formContext.data.entity.save(1);

    console.log(`Next run date updated from ${startDate.toISOString()} to ${newDate.toISOString()}`);
}

function handleJointCampaignType(executionContext) {
    var formContext = executionContext.getFormContext();
    var campaignTypeAttr = formContext.getAttribute("duc_campaigntype");

    if (!campaignTypeAttr) {
        console.error("Attribute 'duc_campaigntype' not found on form.");
        return;
    }

    var campaignType = campaignTypeAttr.getValue();
    if (campaignType === null) {
        console.warn("Campaign type is not selected yet.");
        return;
    }

    if (campaignType == 100000003) { //Joint Campaign
        var campaignInternalTypeAttr = formContext.getAttribute("duc_campaigninternaltype");
        if (!campaignInternalTypeAttr) {
            console.error("Attribute 'duc_campaigninternaltype' not found on form.");
            return;
        }

        var internalTypeValue = campaignInternalTypeAttr.getValue();

        if (internalTypeValue === null) {
            campaignInternalTypeAttr.setValue(100000003); // Main Joint Campaign
        }
    }
}

function oncampaignStatusChange(executionContext) {
    var formContext = executionContext.getFormContext();
    handleFormLock(formContext, executionContext);
}

function handleFormLock(formContext, executionContext) {

    var campaignStatus = formContext.getAttribute("duc_campaignstatus").getValue();

    if (formContext.ui.getFormType() === 1) {
        disableEntireForm(executionContext, false);
        return;
    }

    if (campaignStatus === 4 || campaignStatus === 7 || campaignStatus === 100000001) {
        disableEntireForm(executionContext, false);
    }
    else {
        disableEntireForm(executionContext, true);
    }
}

function disableEntireForm(executionContext, disable) {
    var formContext = executionContext.getFormContext();

    var controls = formContext.ui.controls.get();

    controls.forEach(function (control) {
        try {
            if (control && control.setDisabled) {
                control.setDisabled(disable);
            }
        } catch (e) { }
    });
}
//#endregion

//#region From duc_ReadFromConfigurations
function customLoad(executionContext, ctrlNames) {
    var formContext = executionContext.getFormContext();
    if (ctrlNames != null && ctrlNames.length > 0) {
        var i = 0;
        for (i = 0; i < ctrlNames.length; i++) {
            var wrControl = formContext.getControl(ctrlNames[i]);
            if (wrControl) {
                wrControl.getContentWindow().then(
                    function (contentWindow) {
                        contentWindow.setClientApiContext(Xrm, formContext);
                    }
                )
            }
        }
    }
}

function getConfig(key) {
    var result = GetFromRestAPI("duc_configurationses", "duc_value", "duc_name", key);
    if (result != null && result.value != null && result.value.length > 0 && result.value[0].duc_value != null)
        return result.value[0].duc_value;
    else
        return null;
}

function GetFromRestAPI(entityName, selectColName, filterColName, value, fetchXml) {
    var globalContext = Xrm.Utility.getGlobalContext();
    var serverUrl = globalContext.getClientUrl();
    //serverUrl = _context.getClientUrl();
    var oDataSelect = serverUrl + "/api/data/v8.1/" + entityName;

    if (selectColName != null) {
        oDataSelect += "?$select=" + selectColName;
        if (filterColName != null) {
            if (value === true || value === false)
                oDataSelect += "&$filter=" + filterColName + " eq " + value;
            else
                oDataSelect += "&$filter=" + filterColName + " eq '" + value + "'";
        }
    }
    else if (filterColName != null) {
        if (value === true || value === false)
            oDataSelect += "?$filter=" + filterColName + " eq " + value;
        else
            oDataSelect += "?$filter=" + filterColName + " eq '" + value + "'";
    }
    else if (fetchXml != null)
        oDataSelect += "?fetchXml=" + fetchXml;

    var retrieveReq = new XMLHttpRequest();
    retrieveReq.open("GET", oDataSelect, false);

    retrieveReq.setRequestHeader("Accept", "application/json");
    retrieveReq.setRequestHeader("Content-Type", "application/json;charset=utf-8");

    retrieveReq.send();
    console.log(retrieveReq.readyState);
    if (retrieveReq.readyState == 4
        && retrieveReq.status == 200) {
        var result = JSON.parse(retrieveReq.responseText);
        return result;
    }
    console.log("return null");
    return null;
}

if (!String.prototype.format) {
    String.prototype.format = function () {
        var args = arguments;
        return this.replace(/{(\d+)}/g, function (match, number) {
            return typeof args[number] != 'undefined'
                ? args[number]
                : match
                ;
        });
    };
}
function GetDistance(originLat, originLong, destLat, destLong) {
    //var retrieveReq = new XMLHttpRequest();
    try {
        return Math.hypot(originLat - destLat, originLong - destLong) * 100000;


        //var from = new google.maps.LatLng(originLat, originLong);
        //var to = new google.maps.LatLng(destLat, destLong);
        //return google.maps.geometry.spherical.computeDistanceBetween(from, to);


        /*
        var serverUrl = _context.getClientUrl();

        var oDataSelect = getConfig("MapsURL");
        oDataSelect = oDataSelect.format(originLat, originLong, destLat, destLong, getConfig("MapsKey"));
        retrieveReq.open("GET", oDataSelect, false);

        retrieveReq.setRequestHeader("Accept", "application/json");
        retrieveReq.setRequestHeader("Content-Type", "application/json;charset=utf-8");

        retrieveReq.send();
        console.log(retrieveReq.readyState);
        if (retrieveReq.readyState == 4
            && retrieveReq.status == 200) {
            var result = JSON.parse(retrieveReq.responseText);
            return result.resourceSets[0].resources[0].results[0].travelDistance;
        }*/
        console.log("return null");
        return null;
    }
    catch (error) { console.error(error); }
}

//Fill Lookup based on value from another lookup
//all you have to do is specify the 
//form context, 
//srcFormField => name of source field on form, 
//destFormField => name of form field on form

//srcEntity => sorce entity name
//srcEntityIdField => id column name of source entity 
//srcEntitylookupFieldName => column name of dest field in source entity

//destEntity => destination entity name
//destEntityFieldId => id column name of destination entity
//destEntityFieldName => name field containing the value of dest entity
function fillLookup(executionContext, srcFormField, destFormField, srcEntity,
    srcEntityIdField, srcEntitylookupFieldName, destEntity, destEntityFieldId, destEntityFieldName) {
    var formContext = executionContext.getFormContext();
    var lookup = formContext.getAttribute(srcFormField).getValue();
    if (lookup != null) {
        var lookunfieldStateId = lookup[0].id;
        Xrm.WebApi.online.retrieveMultipleRecords(srcEntity, "?$select=" + srcEntitylookupFieldName
            + "&$expand=" + srcEntitylookupFieldName
            + "&$filter=" + srcEntityIdField + " eq " + lookunfieldStateId
            + "&$orderby=createdon desc&$top=1").then(
                function success(results) {
                    if (results != null && results.entities != null
                        && results.entities.length > 0 && results.entities[0] != null) {
                        var lookup = results.entities[0][srcEntitylookupFieldName];
                        if (lookup == null)
                            return;
                        var value = new Array(); //create a new object array
                        value[0] = new Object();
                        value[0].id = lookup[destEntityFieldId]; // set ID to ID
                        value[0].name = lookup[destEntityFieldName]; //set name to name
                        value[0].entityType = destEntity; //optional
                        formContext.getAttribute(destFormField).setValue(value);
                    }
                },
                function (error) {
                    Xrm.Utility.alertDialog(error.message);
                }
            );
    }
}

//Fill Lookup based on value from another lookup
//all you have to do is specify the 
//form context, 
//srcFormField => name of source field on form, 
//destFormField => name of form field on form

//srcEntity => sorce entity name
//srcEntityIdField => id column name of source entity 
//srcEntitydestFieldName => column name of dest field in source entity
function fillField(executionContext, srcFormField, destFormField, srcEntity,
    srcEntityIdField, srcEntitydestFieldName) {
    var formContext = executionContext.getFormContext();
    var lookup = formContext.getAttribute(srcFormField).getValue();
    if (lookup != null) {
        var lookunfieldStateId = lookup[0].id;
        Xrm.WebApi.online.retrieveMultipleRecords(srcEntity, "?$select=" + srcEntitydestFieldName
            + "&$filter=" + srcEntityIdField + " eq " + lookunfieldStateId
            + "&$orderby=createdon desc&$top=1").then(
                function success(results) {
                    if (results != null && results.entities != null
                        && results.entities.length > 0 && results.entities[0] != null) {
                        formContext.getAttribute(destFormField).setValue(results.entities[0][srcEntitydestFieldName]);
                    }
                },
                function (error) {
                    Xrm.Utility.alertDialog(error.message);
                }
            );
    }
}

function GetEntityTypeCode(entityLogicalName) {
    var result = GetFromRestAPI("EntityDefinitions", "ObjectTypeCode", "LogicalName", entityLogicalName);
    if (result != null && result.value != null && result.value.length > 0 && result.value[0] != null && result.value[0].ObjectTypeCode != null) {
        return result.value[0].ObjectTypeCode;
    } else {
        return null;
    }
}

function associateRequest(entity1Nameplural, entity1ID, entity2Nameplural, entity2ID, relName) {
    var associate = {
        "@odata.id": serverURL + "/api/data/v9.1/" + entity2Nameplural + "(" + entity2ID + ")"
    };
    var serverURL = Xrm.Page.context.getClientUrl();
    var req = new XMLHttpRequest();
    req.open("POST", serverURL + "/api/data/v9.1/" + entity1Nameplural + "(" + entity1ID + ")/" + relName + "/$ref", true);
    req.setRequestHeader("Accept", "application/json");
    req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    req.setRequestHeader("OData-MaxVersion", "4.0");
    req.setRequestHeader("OData-Version", "4.0");
    req.onreadystatechange = function () {
        if (this.readyState == 4 /* complete */) {
            req.onreadystatechange = null;
            if (this.status == 204) {
                console.log("Associated");
                return true;
            } else {
                var error = JSON.parse(this.response).error;
                console.error(error.message);
            }
        }
    };
    req.send(associate);
}

function disassociateRequest(entity1Nameplural, entity1ID, entity2Nameplural, entity2ID, relName) {

    var serverURL = Xrm.Page.context.getClientUrl();
    var req = new XMLHttpRequest();
    req.open("DELETE", serverURL + "/api/data/v9.1/" + entity1Nameplural + "(" + entity1ID + ")/" + relName + "/$ref?$id=" + serverURL + "/api/data/v9.1/" + entity2Nameplural + "(" + entity2ID + ")", true);
    req.setRequestHeader("Accept", "application/json");
    req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    req.setRequestHeader("OData-MaxVersion", "4.0");
    req.setRequestHeader("OData-Version", "4.0");
    req.onreadystatechange = function () {
        if (this.readyState == 4 /* complete */) {
            req.onreadystatechange = null;
            if (this.status == 204) {
                console.log('Record Disassociated');
            } else {
                var error = JSON.parse(this.response).error;
                console.error(error.message);
            }
        }
    };
    req.send();
}

function UpdateEntity(EntityName, id, object, callback) {
    Xrm.WebApi.updateRecord(EntityName, id, object).then(
        function success(result) {
            //alert("update success");
            console.log("updated");
            if (callback != null) {
                callback();
            } else {
                Xrm.Utility.closeProgressIndicator();
            }
            return true;
        },
        function (error) {
            Xrm.Utility.closeProgressIndicator();
            //alert("update failed");
            alert(error.message);
            console.log(error.message);
            return false;
        }
    );
}

function ReadOnly(executionContext, conditionFieldName, expectedVal, condition2FieldName, expectedVal2) {
    var formContext = executionContext.getFormContext();
    if (conditionFieldName != null && conditionFieldName != ''
        && (formContext.getAttribute(conditionFieldName) == null
            || formContext.getAttribute(conditionFieldName).getValue() != expectedVal))
        return;
    else if (condition2FieldName != null && condition2FieldName != ''
        && (formContext.getAttribute(condition2FieldName) == null
            || formContext.getAttribute(condition2FieldName).getValue() != expectedVal2))
        return;
    else {
        var cs = formContext.ui.controls.get();
        function DisableSubgrid() {
            var subGridCtrl = Xrm.Page.getControl("SubgridName");

            // If subgrid is not loaded yet, then call same function after some time.
            if (subGridCtrl == null) {
                setTimeout(DisableSubgrid, 1000);
                return;
            }

            // Disable the subgrid control
            subGridCtrl.setDisabled(true);
        }

        //
        //var tabs = formContext.ui.tabs.get();
        //for (var i in tabs) {
        //    var sections = tabs[i].sections.get();
        //    for (var j in sections) {
        //        cs.push(sections[j].controls.get());
        //    }
        //}


        for (var i in cs) {
            var c = cs[i];
            if (c.getName() != "" && c.getName() != null) {
                if (!c.getDisabled()) { c.setDisabled(true); }
            }
        }
    }
}

function updateRecordSync(id, entityObject, odataSetName) {

    var jsonEntity = window.JSON.stringify(entityObject);

    // Get Server URL

    var serverUrl = Xrm.Page.context.getServerUrl();

    //The OData end-point

    var ODATA_ENDPOINT = "/XRMServices/2011 / OrganizationData.svc";

    var updateRecordReq = new XMLHttpRequest();

    var ODataPath = serverUrl + ODATA_ENDPOINT;

    updateRecordReq.open('POST', ODataPath + "/" + odataSetName + "(guid'" + id + ")", false);

    updateRecordReq.setRequestHeader("Accept", "application / json");

    updateRecordReq.setRequestHeader("Content - Type", "application / json; charset = utf - 8");

    updateRecordReq.setRequestHeader("X - HTTP - Method", "MERGE");

    updateRecordReq.send(jsonEntity);

}

Date.prototype.addHours = function (h) {
    this.setTime(this.getTime() + (h * 60 * 60 * 1000));
    return this;
}

function IsFromMobile() {
    if (Xrm.Page.context.client.getClient() == "Mobile"
        && (Xrm.Page.context.client.getFormFactor() == 2 || Xrm.Page.context.client.getFormFactor() == 3)
    ) {
        return true;
    }
    return false;
}

function IsCreateMode() {
    var formType = Xrm.Page.ui.getFormType();
    if (formType == 1)
        return true;

    return false;
}

function IsDistanceAllowed(workorderid, shallOpenWO, formContext) {
    if (!IsFromMobile())
        return;

    if (shallOpenWO) {
        Xrm.Utility.showProgressIndicator();
    }
    var nearLocMsg = "";
    var cannotGetLocation = "";
    var noAddressFound = "";

    {
        var userSettings = Xrm.Utility.getGlobalContext().userSettings;

        var userLanguage = userSettings.languageId; // (1025 : ar-SA  && 1033: en-US)
        if (userLanguage != null && userLanguage == 1033) {
            nearLocMsg = "The current location is {0} meters away from the registered location and max allowed distance is {1}. Would you like to open Google maps to locate it?";
            cannotGetLocation = "Can't get current location. Please update your mobile location setting";
            noAddressFound = "Please specify account's address location to be able to start the inspection";
            NotAssignedInspector = "You are not assigned to this booking to start inspection";
        }

        else if (userLanguage != null && userLanguage == 1025) {
            nearLocMsg = "الموقع الحالي يبعد بمسافة {0} متر عن موقع التفتيش المسجل وأقصي مسافة مسموحة هي {1}. هل ترغب في فتح خريطة جوجل لتتمكن من تحديد الموقع؟";
            cannotGetLocation = "تعذر استرداد الموقع الحالي ،يرجى مراجعة إعدادت الموقع في هاتفك";
            noAddressFound = "يرجى تحديد العنوان الخاص بالحساب حتى تتمكن من بدء التفتيش";
            NotAssignedInspector = "يجب أن تكون المفتش المحدد في أمر التفتيش لتتمكن من بدء العمل";
        }
    }


    var incidentId;
    var isBSS = false;
    var islandmark = false;
    var allowedDist = -1;
    var accountid;
    //get channel
    // get bss long and lat
    var accLong = 0.00000;
    var accLat = 0.00000;

    var result = GetFromRestAPI("msdyn_workorders", "_msdyn_serviceaccount_value,_msdyn_primaryincidenttype_value,duc_details_latitude,duc_details_longitude,duc_channel,duc_details_islandmark,duc_details_alloweddistance", "msdyn_workorderid", workorderid);
    if (result != null && result.value != null && result.value.length > 0) {
        incidentId = result.value[0]._msdyn_primaryincidenttype_value;
        isBSS = result.value[0].duc_channel == 100000002;
        islandmark = result.value[0].duc_details_islandmark;
        if (islandmark) {
            allowedDist = result.value[0].duc_details_alloweddistance;
        }
        if (isBSS) {
            accLong = result.value[0].duc_details_longitude;//fill
            accLat = result.value[0].duc_details_latitude;//fill
        }
        accountid = result.value[0]._msdyn_serviceaccount_value;

        console.log(accountid);

    }

    var OUfieldName = isBSS ? "duc_bssdistanceinmeter" : "duc_distanceinmeter";

    if (allowedDist == -1 || allowedDist == null || allowedDist == 0) {
        //get OU from incident type
        if (incidentId != null)
            result = GetFromRestAPI("msdyn_incidenttypes", "_duc_organizationalunitid_value", "msdyn_incidenttypeid", incidentId);
        if (result != null && result.value != null && result.value.length > 0
            && result.value[0] != null && result.value[0]._duc_organizationalunitid_value != null) {
            var orgId = result.value[0]._duc_organizationalunitid_value;

            //Get Distance from OU
            if (orgId != null)//5399edb5-caa4-4a93-95ad-7f5c4625614d
                result = GetFromRestAPI("msdyn_organizationalunits", OUfieldName, "msdyn_organizationalunitid", orgId);
            if (result != null && result.value != null && result.value.length > 0
                && result.value[0] != null && result.value[0][OUfieldName] != null) {
                allowedDist = result.value[0][OUfieldName] - 100000000;

            }
        }
    }
    console.log("allowed Distance: " + allowedDist);

    if (allowedDist == null || allowedDist <= 0)
        allowedDist = getConfig("DistanceInMetre");

    if (allowedDist == null) {
        if (shallOpenWO) {
            Xrm.Utility.closeProgressIndicator();
            OpenWO(workorderid);
        }
    }
    else {
        //get account location
        if (!isBSS) {
            if (accountid != null)
                result = GetFromRestAPI("accounts", "address1_latitude,address1_longitude", "accountid", accountid);
            if (result != null && result.value != null && result.value.length > 0
                && result.value[0] != null && result.value[0].address1_latitude != null
                && result.value[0].address1_longitude != null) {
                accLat = result.value[0].address1_latitude;
                accLong = result.value[0].address1_longitude;
            }
        }

        var originlatitude = 0.00000;
        var originlongitude = 0.00000;

        Xrm.Device.getCurrentPosition().then(function (location) {
            originlatitude = location.coords.latitude;
            originlongitude = location.coords.longitude;
            console.log(originlatitude);
            console.log(originlongitude);

            if (accLat > 0 && accLong > 0) {
                console.log(accLat);
                console.log(accLong);

                //calculate distance if there is internet
                var distance = GetDistance(originlatitude, originlongitude, accLat, accLong);

                //alert("Origin: ({0},{1}), Dest: ({2},{3}), Dist: ({4}), allowed: ({5})".format(originlatitude, originlongitude, accLat, accLong, distance, allowedDist));

                if (distance <= allowedDist) {
                    logLocation(originlatitude, originlongitude, accLat, accLong, distance, true, allowedDist, workorderid);
                    if (shallOpenWO) {
                        Xrm.Utility.closeProgressIndicator();
                        OpenWO(workorderid);
                    }

                }
                else {
                    logLocation(originlatitude, originlongitude, accLat, accLong, distance, false, allowedDist, workorderid);
                    nearLocMsg = nearLocMsg.format(Math.round(distance), allowedDist);
                    End(nearLocMsg, !shallOpenWO, workorderid, "https://www.google.com/maps/dir/" + originlatitude + "," + originlongitude + "/" + accLat + "," + accLong);
                }
            }
            else {
                End(noAddressFound, !shallOpenWO, workorderid);
            }
        }, function () {
            End(cannotGetLocation, !shallOpenWO, workorderid);
        });
    }
}

function End(msg, redirect, workorderid, link) {
    if (link != null) {
        Xrm.Utility.confirmDialog(msg, function () {
            //CRM Crate - JavaScript Snippet
            //Creating An Object For Storing Window Height & Width
            var openUrlOptions = {
                height: 800,
                width: 800
            };
            //Using openURL Client API To Open The Website.
            Xrm.Navigation.openUrl(link, openUrlOptions);
            Xrm.Utility.closeProgressIndicator();

        }, function () { });
        Xrm.Utility.closeProgressIndicator();
    }
    else {
        Xrm.Utility.alertDialog(msg);
        Xrm.Utility.closeProgressIndicator();
    }
    if (redirect) {
        var entityFormOptions = {};
        entityFormOptions["entityName"] = "msdyn_workorder";
        entityFormOptions["entityId"] = workorderid;

        Xrm.Navigation.openForm(entityFormOptions).then(
            function (success) {
                console.log(success);
                Xrm.Utility.closeProgressIndicator();
            },
            function (error) {
                console.log(error);
            });
    }
    else {
    }
}

function logLocation(devicelat, devicelong, acclat, acclong, distance, isSuccess, alloweddist, woId) {
    var log = getConfig("LogLocation");
    //alert(log);
    if ((log == "1" && !isSuccess) // log failures
        || log == "2") {//log all 
        //alert('inside');

        var entityLogicalName = "duc_locationlog";
        var userId = Xrm.Utility.getGlobalContext().getUserId().toLowerCase();
        userId = userId.replace("{", '').replace("}", '');
        woId = woId.replace("{", '').replace("}", '');
        var data = {
            "duc_userid@odata.bind": "/systemusers(" + userId + ")",
            "duc_WorkOrder@odata.bind": "/msdyn_workorders(" + woId + ")",
            "duc_devicelat": devicelat.toString(),
            "duc_devicelong": devicelong.toString(),
            "duc_acclat": acclat.toString(),
            "duc_acclong": acclong.toString(),
            "duc_distance": distance.toString(),
            "duc_alloweddist": alloweddist.toString(),
            "duc_issuccess": isSuccess,
        };
        Xrm.WebApi.createRecord(entityLogicalName, data).then(
            function success(result) {
                //alert("success");
            }
            , function failed(error) {
                //alert(error.message);
            });
    }

}

function CreateEntity(EntityName, object, callback) {
    Xrm.WebApi.createRecord(EntityName, object).then(
        function success(result) {
            //alert("update success");
            console.log("updated");
            if (callback != null) {
                callback();
            } else {
                Xrm.Utility.closeProgressIndicator();
            }
            return true;
        },
        function (error) {
            Xrm.Utility.closeProgressIndicator();
            //alert("update failed");
            alert(error.message);
            console.log(error.message);
            return false;
        }
    );
}

function ExportPDF(entityTypeCode, templateId, recordId, filename, savetoNotes, successCallBack, errorCallBack) {
    var entityRecordId = "['" + recordId + "']";
    var selectedTemplateRef = {
        "id": templateId,
        "entityType": "documenttemplate"
    };

    exportPDFDocumentRequest = new CustomODataContract.ExportPDFDocumentRequest(entityTypeCode, selectedTemplateRef, entityRecordId);
    Xrm.Utility.showProgressIndicator();
    Xrm.WebApi.online.execute(exportPDFDocumentRequest).then(function (response) {
        if (response) {
            response.json().then(function (pdfFileContents) { successCallBack(pdfFileContents) }, function (error) { errorCallBack(error); })
        }
    }, function (error) {
        console.log(error.message);
        console.log(error.raw);
        ExportPDFCustom(entityTypeCode, templateId, recordId, filename, savetoNotes, successCallBack, errorCallBack);
    });
    return true;
}

var Base64ToBlob = function (fileContent, fileType) {
    if (!fileType)
        throw new Error("file Type cannot be empty");
    // convert base64 content to raw binary data held in a string
    var binary = base64DecToArr(fileContent);
    return new Blob([binary], { type: fileType });
};

var base64DecToArr = function (sBase64, nBlocksSize) {
    var sB64Enc = sBase64.replace(/[^A-Za-z0-9\+\/]/g, ""), nInLen = sB64Enc.length, nOutLen = nBlocksSize ? Math.ceil(((nInLen * 3 + 1) >> 2) / nBlocksSize) * nBlocksSize : (nInLen * 3 + 1) >> 2, taBytes = new Uint8Array(nOutLen);
    for (var nMod3 = void 0, nMod4 = void 0, nUint24 = 0, nOutIdx = 0, nInIdx = 0; nInIdx < nInLen; nInIdx++) {
        nMod4 = nInIdx & 3;
        nUint24 |= b64ToUint6(sB64Enc.charCodeAt(nInIdx)) << (18 - 6 * nMod4);
        if (nMod4 === 3 || nInLen - nInIdx === 1) {
            for (nMod3 = 0; nMod3 < 3 && nOutIdx < nOutLen; nMod3++, nOutIdx++) {
                taBytes[nOutIdx] = (nUint24 >>> ((16 >>> nMod3) & 24)) & 255;
            }
            nUint24 = 0;
        }
    }
    return taBytes;
};

var b64ToUint6 = function (nChr) {
    return nChr > 64 && nChr < 91
        ? nChr - 65
        : nChr > 96 && nChr < 123
            ? nChr - 71
            : nChr > 47 && nChr < 58
                ? nChr + 4
                : nChr === 43
                    ? 62
                    : nChr === 47
                        ? 63
                        : 0;
};

function checkPositionPermission(fieldName, showOnMobile) {
    if (!showOnMobile && IsFromMobile())
        return false;

    var userId = Xrm.Utility.getGlobalContext().getUserId();
    var fetchXML =
        "<fetch version='1.0' output-format='xml-platform' mapping='logical' no-lock='true' distinct='false'>"
        + "<entity name = 'systemuser' >"
        + "<attribute name='fullname' />"
        + "<attribute name='businessunitid' />"
        + "<attribute name='title' />"
        + "<attribute name='address1_telephone1' />"
        + "<attribute name='positionid' />"
        + "<attribute name='systemuserid' />"
        + "<order attribute='fullname' descending='false' />"
        + "<filter type='and'>"
        + "<condition attribute='systemuserid' operator='eq' uitype='systemuser' value='" + userId + "' />"
        + "</filter>"
        + "<link-entity name='position' from='positionid' to='positionid' link-type='inner' alias='ab'>"
        + "<filter type='and'>"
        + "<condition attribute='" + fieldName + "' operator='eq' value='1' />"
        + "</filter>"
        + "</link-entity>"
        + "</entity >"
        + "</fetch >"
    result = GetFromRestAPI("systemusers", null, null, null, fetchXML);
    if (result != null && result.value != null && result.value.length > 0 && result.value[0].systemuserid != null) {
        return true;

    }
    return false;
}
//#endregion

function onCampaignTypeChange(executionContext) {
    var formContext = executionContext.getFormContext();

    var campaignTypeAttr = formContext.getAttribute("duc_campaigntype");
    var internalTypeAttr = formContext.getAttribute("duc_campaigninternaltype");

    if (!campaignTypeAttr || !internalTypeAttr) {
        return;
    }

    var campaignTypeValue = campaignTypeAttr.getValue();

    if (campaignTypeValue === 100000004) {
        internalTypeAttr.setValue(100000005);
        internalTypeAttr.setSubmitMode("always");
    }
}

function onLoadSetWOSection(executionContext) {
    var formContext = executionContext.getFormContext();

    var CAMP_TYPE_ADHOC = 100000000;
    var CAMP_TYPE_PLANNED = 100000001;
    var CAMP_TYPE_PERIODIC = 100000002;
    var CAMP_TYPE_JOINT = 100000003;
    var CAMP_TYPE_PATROL = 100000004;

    var INT_SUB_PERIODIC = 100000000;
    var INT_SUB_JOINT = 100000001;
    var INT_MAIN_PERIODIC = 100000002;
    var INT_MAIN_JOINT = 100000003;
    var INT_SUB_PATROL = 100000004;
    var INT_MAIN_PATROL = 100000005;

    var campTypeAttr = formContext.getAttribute("duc_campaigntype");

    var intTypeAttr = formContext.getAttribute("duc_campaigninternaltype");

    var campType = campTypeAttr ? campTypeAttr.getValue() : null;

    var intType = intTypeAttr ? intTypeAttr.getValue() : null;

    var tab = formContext.ui.tabs.get("tab_7");

    var childSection = tab ? tab.sections.get("tab_7_section_1") : null;

    var parentSection = tab ? tab.sections.get("InspectionKPIs_Tab_section_10") : null;

    if (childSection) childSection.setVisible(false);

    if (parentSection) parentSection.setVisible(false);

    var isPeriodicJointPatrol =
        campType === CAMP_TYPE_PERIODIC ||
        campType === CAMP_TYPE_JOINT ||
        campType === CAMP_TYPE_PATROL;

    var isSubInternal =
        intType === INT_SUB_PERIODIC ||
        intType === INT_SUB_JOINT ||
        intType === INT_SUB_PATROL;

    var isMainInternal =
        intType === INT_MAIN_PERIODIC ||
        intType === INT_MAIN_JOINT ||
        intType === INT_MAIN_PATROL;

    if (isPeriodicJointPatrol) {
        if (isSubInternal && childSection) {
            childSection.setVisible(true);
        } else if (isMainInternal && parentSection) {
            parentSection.setVisible(true);
        }
    }
    else if ((campType === CAMP_TYPE_ADHOC || campType === CAMP_TYPE_PLANNED) && parentSection) {
        childSection.setVisible(true);
    }
}      