function onLastActionChange(executionContext) {
    var formContext = executionContext.getFormContext();

    // Get the field (replace with your actual field name)
    var lastActionAttr = formContext.getAttribute("duc_lastactiontaken");
    if (!lastActionAttr) return; // Field not on form

    var lastAction = lastActionAttr.getValue();

    // Skip if the field is empty (cleared)
    if (lastAction === null || lastAction === undefined || lastAction.length === 0) {
        console.log("Last Action cleared — no save triggered");
        return;
    }

    console.log("Last Action changed to:", lastAction);

    // Only save if there are unsaved changes
    if (formContext.data.entity.getIsDirty()) {
        formContext.data.entity.save();
    }
}


// ///////////////////////////////////////////////////////////////////////////
// Process Automation Offline


var DUC = DUC || {};
DUC.ProcessAutomation = DUC.ProcessAutomation || {};
DUC.ProcessAutomation.Offline = DUC.ProcessAutomation.Offline || {};

// ///////////////////////////////////////////////////////////////////////////
// MODULE 1: LOGGER
// ///////////////////////////////////////////////////////////////////////////

DUC.ProcessAutomation.Offline.Logger = (function () {
    "use strict";
    var _P = "[PA-Offline]";

    function trace(message, level) {
        level = level || "INF";
        var ts = new Date().toISOString();
        if (typeof console !== "undefined") {
            var f = _P + "[" + level + "][" + ts + "]: " + message;
            switch (level) {
                case "ERR": case "CRT": console.error(f); break;
                case "WRN": console.warn(f); break;
                default: console.log(f); break;
            }
        }
    }

    return {
        trace: trace,
        error: function (m) { trace(m, "ERR"); },
        warn: function (m) { trace(m, "WRN"); }
    };
})();

// ///////////////////////////////////////////////////////////////////////////
// MODULE 2: ENTITY SET NAME HELPER
// ///////////////////////////////////////////////////////////////////////////

DUC.ProcessAutomation.Offline._entitySetName = function (logicalName) {
    "use strict";
    var map = {
        "msdyn_workordersubstatus": "msdyn_workordersubstatuses",
        "msdyn_workorderstatus": "msdyn_workorderstatuses",
        "duc_processstatus": "duc_processstatuses",
        "duc_processsubstatus": "duc_processsubstatuses",
        "duc_processstage": "duc_processstages",
        "systemuser": "systemusers",
        "team": "teams",
        "account": "accounts",
        "contact": "contacts",
        "msdyn_workorder": "msdyn_workorders",
        "duc_stageaction": "duc_stageactions",
        "duc_actiontype": "duc_actiontypes",
        "duc_processdefinition": "duc_processdefinitions",
        "duc_processextension": "duc_processextensions",
        "duc_processactionlog": "duc_processactionlogs"
    };
    if (map[logicalName]) return map[logicalName];
    if (/(?:s|x|z|ch|sh)$/.test(logicalName)) return logicalName + "es";
    return logicalName + "s";
};

// ///////////////////////////////////////////////////////////////////////////
// MODULE 3: STEP ACTION INFO SERVICE
// Replaces 5-table FetchXML with sequential OData calls
// ///////////////////////////////////////////////////////////////////////////

DUC.ProcessAutomation.Offline.StepActionInfoService = (function () {
    "use strict";
    var _log = function () { return DUC.ProcessAutomation.Offline.Logger; };

    function getActionInfo(stageActionId) {
        var info = {};

        return Xrm.WebApi.offline.retrieveRecord("duc_stageaction", stageActionId,
            "?$select=duc_name,duc_usedefaultexecution,duc_generatecertificate," +
            "duc_generatedocument,duc_documenttemplatename,duc_usedefaultvalidation," +
            "_duc_actiontype_value,_duc_nextstage_value,_duc_relatedstage_value," +
            "_duc_executionpowerautomate_value,_duc_defaultstatus_value," +
            "_duc_defaultsubstatus_value,_duc_processaction_value," +
            "_duc_validationpowerautomate_value"
        )
            .then(function (a) {
                if (!a) return null;
                info.Action_Id = stageActionId;
                info.Action_Name = a.duc_name;
                info.Action_UseDefaultExecution = a.duc_usedefaultexecution || false;
                info.Action_GenerateCertificate = a.duc_generatecertificate || false;
                info.Action_GenerateDocument = a.duc_generatedocument || false;
                info.Action_DocumentTemplateName = a.duc_documenttemplatename;
                info.Action_UseDefaultValidation = a.duc_usedefaultvalidation || false;
                info.Action_ActionType_Id = a["_duc_actiontype_value"] || null;
                info.Action_NextStage_Id = a["_duc_nextstage_value"] || null;
                info.Action_RelatedStage_Id = a["_duc_relatedstage_value"] || null;
                info.Action_Flow_Id = a["_duc_executionpowerautomate_value"] || null;
                info.Action_Status_Id = a["_duc_defaultstatus_value"] || null;
                info.Action_SubStatus_Id = a["_duc_defaultsubstatus_value"] || null;

                var promises = [];

                if (info.Action_ActionType_Id) {
                    promises.push(
                        Xrm.WebApi.offline.retrieveRecord("duc_actiontype", info.Action_ActionType_Id,
                            "?$select=duc_sendtocustomer,duc_wfaction"
                        ).then(function (at) {
                            info.ActionType_SendToCustomer = at.duc_sendtocustomer || false;
                            info.Action_IsWorkflowAction = (at.duc_wfaction !== false);
                        }).catch(function () {
                            info.ActionType_SendToCustomer = false;
                            info.Action_IsWorkflowAction = true;
                        })
                    );
                } else { info.ActionType_SendToCustomer = false; info.Action_IsWorkflowAction = true; }

                if (info.Action_NextStage_Id) {
                    promises.push(
                        Xrm.WebApi.offline.retrieveRecord("duc_processstage", info.Action_NextStage_Id,
                            "?$select=duc_sequence,duc_usedefaultexecution,duc_executefinalpowerautomate," +
                            "duc_isautomaticcalculatestep,_duc_executionpowerautomate_value"
                        ).then(function (ns) {
                            info.NextStage_Sequence = ns.duc_sequence;
                            info.NextStage_IsAutoCalculateStep = ns.duc_isautomaticcalculatestep || false;
                        }).catch(function () {
                            info.NextStage_Sequence = null; info.NextStage_IsAutoCalculateStep = false;
                        })
                    );
                } else { info.NextStage_Sequence = null; info.NextStage_IsAutoCalculateStep = false; }

                if (info.Action_RelatedStage_Id) {
                    promises.push(
                        Xrm.WebApi.offline.retrieveRecord("duc_processstage", info.Action_RelatedStage_Id,
                            "?$select=duc_sequence,duc_usedefaultexecution,duc_isautomaticcalculatestep," +
                            "duc_name,_duc_relatedprocess_value,_duc_executionpowerautomate_value"
                        ).then(function (cs) {
                            info.RelatedStage_Sequence = cs.duc_sequence;
                            info.RelatedStage_Name = cs.duc_name;
                            info.RelatedStage_RelatedProcess_Id = cs["_duc_relatedprocess_value"] || null;
                        }).catch(function () {
                            info.RelatedStage_Sequence = null; info.RelatedStage_RelatedProcess_Id = null;
                        })
                    );
                }

                if (info.Action_Status_Id) {
                    promises.push(
                        Xrm.WebApi.offline.retrieveRecord("duc_processstatus", info.Action_Status_Id, "?$select=duc_value")
                            .then(function (s) { info.Status_Value = s.duc_value || null; })
                            .catch(function () { info.Status_Value = null; })
                    );
                } else { info.Status_Value = null; }

                if (info.Action_SubStatus_Id) {
                    promises.push(
                        Xrm.WebApi.offline.retrieveRecord("duc_processsubstatus", info.Action_SubStatus_Id, "?$select=duc_value")
                            .then(function (ss) { info.SubStatus_Value = ss.duc_value || null; })
                            .catch(function () { info.SubStatus_Value = null; })
                    );
                } else { info.SubStatus_Value = null; }

                return Promise.all(promises);
            })
            .then(function () {
                if (!info.RelatedStage_RelatedProcess_Id) {
                    info.ProcessDefinition_TargetStatusLookupName = null;
                    info.ProcessDefinition_TargetSubStatusLookupName = null;
                    info.ProcessDefinition_ExistingStatusEntity = null;
                    info.ProcessDefinition_ExistingSubStatusEntity = null;
                    info.ProcessDefinition_StatusLookupType = null;
                    info.ProcessDefinition_SubStatusLookupType = null;
                    return info;
                }
                return Xrm.WebApi.offline.retrieveRecord("duc_processdefinition", info.RelatedStage_RelatedProcess_Id,
                    "?$select=duc_usedefaultexecution,duc_targetentitystatuslookupname," +
                    "duc_targetentitysubstatuslookupname,duc_existingstatusentity," +
                    "duc_existingsubstatusentity,duc_statuslookuptype,duc_substatuslookuptype," +
                    "duc_targetentitycustomerlookupname,_duc_finalexecutionpowerautomate_value"
                ).then(function (pd) {
                    info.ProcessDefinition_UseDefaultExecution = pd.duc_usedefaultexecution || false;
                    info.ProcessDefinition_TargetStatusLookupName = pd.duc_targetentitystatuslookupname || null;
                    info.ProcessDefinition_TargetSubStatusLookupName = pd.duc_targetentitysubstatuslookupname || null;
                    info.ProcessDefinition_ExistingStatusEntity = pd.duc_existingstatusentity || null;
                    info.ProcessDefinition_ExistingSubStatusEntity = pd.duc_existingsubstatusentity || null;
                    info.ProcessDefinition_StatusLookupType = pd.duc_statuslookuptype || null;
                    info.ProcessDefinition_SubStatusLookupType = pd.duc_substatuslookuptype || null;
                    return info;
                }).catch(function () {
                    info.ProcessDefinition_TargetStatusLookupName = null;
                    info.ProcessDefinition_TargetSubStatusLookupName = null;
                    return info;
                });
            })
            .catch(function (err) {
                _log().error("getActionInfo failed: " + err.message);
                return null;
            });
    }

    return { getActionInfo: getActionInfo };
})();

// ///////////////////////////////////////////////////////////////////////////
// MODULE 4: ASSIGNMENT OFFLINE
// ///////////////////////////////////////////////////////////////////////////

DUC.ProcessAutomation.Offline.Assignment = (function () {
    "use strict";
    var _log = function () { return DUC.ProcessAutomation.Offline.Logger; };
    var _esn = function (n) { return DUC.ProcessAutomation.Offline._entitySetName(n); };
    var METHOD = { EqualDistribution: 771010000, PreferredResource: 771010001, ServiceHandler: 771010002 };

    function executeAssignmentLogic(processExtensionId, regarding) {
        return Xrm.WebApi.offline.retrieveRecord("duc_processextension", processExtensionId,
            "?$select=_duc_currentstage_value,_duc_nextassignee_value," +
            "_duc_lastactiontaken_value,_duc_customer_value," +
            "_duc_customer_value@Microsoft.Dynamics.CRM.lookuplogicalname,_ownerid_value"
        ).then(function (pe) {
            var stageId = pe["_duc_currentstage_value"];
            if (!stageId) return Promise.resolve();
            return Xrm.WebApi.offline.retrieveRecord("duc_processstage", stageId,
                "?$select=_duc_assignedteam_value,duc_autoassignment,duc_assignmentmethod,_duc_preferredresource_value"
            ).then(function (step) {
                var naId = pe["_duc_nextassignee_value"];
                if (naId) {
                    return _assignToUser(processExtensionId, naId).then(function () {
                        return Xrm.WebApi.offline.updateRecord("duc_processextension", processExtensionId, { "duc_nextassignee@odata.bind": null });
                    });
                }
                var lastActionId = pe["_duc_lastactiontaken_value"];
                if (lastActionId) {
                    return _getActionLog(regarding.id, stageId, lastActionId).then(function (owner) {
                        if (owner) {
                            if (owner.entityType === "team") return _assignToTeam(processExtensionId, owner.id, regarding);
                            return _checkLeave(owner.id).then(function (lv) {
                                if (lv.isOnLeave && lv.delegatedUserId) return _assignToUser(processExtensionId, lv.delegatedUserId);
                                if (!lv.isOnLeave) return _assignToUser(processExtensionId, owner.id);
                                return _autoAssign(pe, step, processExtensionId, regarding);
                            });
                        }
                        return _autoAssign(pe, step, processExtensionId, regarding);
                    });
                }
                return _autoAssign(pe, step, processExtensionId, regarding);
            });
        }).catch(function (err) { _log().error("Assignment failed: " + err.message); });
    }

    function _autoAssign(pe, step, peId, regarding) {
        if (step.duc_autoassignment) {
            var m = step.duc_assignmentmethod, t = step["_duc_assignedteam_value"];
            if (m === METHOD.EqualDistribution && t) return _equalDist(pe, peId, t, regarding).then(function () { return _markAuto(peId); });
            if (m === METHOD.PreferredResource) { var p = step["_duc_preferredresource_value"]; if (p) return _prefResource(peId, p, step, regarding).then(function () { return _markAuto(peId); }); }
            if (m === METHOD.ServiceHandler && t) return _equalDist(pe, peId, t, regarding).then(function () { return _markAuto(peId); });
            return _markAuto(peId);
        }
        var mt = step["_duc_assignedteam_value"];
        if (mt) return _assignToTeam(peId, mt, regarding);
        return Promise.resolve();
    }

    function _equalDist(pe, peId, teamId, regarding) {
        return _getTeamMembers(teamId, pe["_ownerid_value"]).then(function (members) {
            if (members.length === 0) return Promise.resolve();
            return _leastLoaded(members, teamId, regarding, peId);
        });
    }

    function _prefResource(peId, prefId, step, regarding) {
        return _checkLeave(prefId).then(function (lv) {
            if (lv.isOnLeave) {
                if (lv.delegatedUserId) {
                    return _checkLeave(lv.delegatedUserId).then(function (dlv) {
                        if (!dlv.isOnLeave) return _assignToUser(peId, lv.delegatedUserId);
                        var t = step["_duc_assignedteam_value"];
                        return t ? _equalDist({}, peId, t, regarding) : _assignToUser(peId, prefId);
                    });
                }
                var t = step["_duc_assignedteam_value"];
                return t ? _equalDist({}, peId, t, regarding) : _assignToUser(peId, prefId);
            }
            return _assignToUser(peId, prefId);
        });
    }

    function _getTeamMembers(teamId, excludeId) {
        return Xrm.WebApi.offline.retrieveMultipleRecords("teammembership",
            "?$select=systemuserid&$filter=_teamid_value eq '" + teamId + "'"
        ).then(function (r) {
            var m = [];
            for (var i = 0; i < r.entities.length; i++) { var u = r.entities[i].systemuserid; if (u && u !== excludeId) m.push({ id: u }); }
            return m;
        }).catch(function () { return []; });
    }

    function _leastLoaded(members, teamId, regarding, peId) {
        var today = new Date(); today.setHours(0, 0, 0, 0); var tISO = today.toISOString();
        var avail = [];
        var checks = members.map(function (m) {
            return _checkLeave(m.id).then(function (lv) {
                if (!lv.isOnLeave) return _countToday(m.id, tISO).then(function (c) { avail.push({ id: m.id, count: c }); });
            });
        });
        return Promise.all(checks).then(function () {
            if (avail.length === 0) return _assignToTeam(peId, teamId, regarding);
            avail.sort(function (a, b) { return a.count - b.count; });
            return _assignToUser(peId, avail[0].id);
        });
    }

    function _countToday(userId, todayISO) {
        return Xrm.WebApi.offline.retrieveMultipleRecords("duc_processextension",
            "?$select=activityid&$filter=_ownerid_value eq '" + userId + "' and modifiedon ge " + todayISO + "&$top=50"
        ).then(function (r) { return r.entities.length; }).catch(function () { return 0; });
    }

    function _checkLeave(userId) {
        var today = new Date().toISOString().split("T")[0];
        return Xrm.WebApi.offline.retrieveMultipleRecords("duc_employeeleave",
            "?$select=duc_employeeleaveid,_duc_delegatedemployee_value&$filter=_duc_employee_value eq '" + userId +
            "' and duc_startdate le " + today + " and duc_enddate ge " + today
        ).then(function (r) {
            if (r.entities.length > 0) return { isOnLeave: true, delegatedUserId: r.entities[0]["_duc_delegatedemployee_value"] || null };
            return { isOnLeave: false, delegatedUserId: null };
        }).catch(function () { return { isOnLeave: false, delegatedUserId: null }; });
    }

    function _getActionLog(regardingId, stageId, lastActionId) {
        return Xrm.WebApi.offline.retrieveMultipleRecords("duc_processactionlog",
            "?$select=_ownerid_value,_ownerid_value@Microsoft.Dynamics.CRM.lookuplogicalname" +
            "&$filter=_duc_processstage_value eq '" + stageId + "' and _duc_action_value ne '" + lastActionId +
            "' and _regardingobjectid_value eq '" + regardingId + "'&$orderby=createdon desc&$top=1"
        ).then(function (r) {
            if (r.entities.length > 0) return { id: r.entities[0]["_ownerid_value"], entityType: r.entities[0]["_ownerid_value@Microsoft.Dynamics.CRM.lookuplogicalname"] || "systemuser" };
            return null;
        }).catch(function () { return null; });
    }

    function _assignToUser(peId, userId) {
        _log().trace("Assignment: → user " + userId);
        return Xrm.WebApi.offline.updateRecord("duc_processextension", peId, { "ownerid@odata.bind": "/systemusers(" + userId + ")" });
    }

    function _assignToTeam(peId, teamId, regarding) {
        _log().trace("Assignment: → team " + teamId);
        if (regarding) return Xrm.WebApi.offline.updateRecord(regarding.entityType, regarding.id, { "ownerid@odata.bind": "/teams(" + teamId + ")" });
        return Promise.resolve();
    }

    function _markAuto(peId) {
        return Xrm.WebApi.offline.updateRecord("duc_processextension", peId, { "duc_assigntype": true });
    }

    return { executeAssignmentLogic: executeAssignmentLogic };
})();

// ///////////////////////////////////////////////////////////////////////////
// MODULE 5: MAIN HANDLER — OnChange duc_lastactiontaken
// ///////////////////////////////////////////////////////////////////////////

DUC.ProcessAutomation.Offline.onLastActionTakenChange = function (executionContext) {
    "use strict";

    var formContext = executionContext.getFormContext();
    var Logger = DUC.ProcessAutomation.Offline.Logger;
    var esn = DUC.ProcessAutomation.Offline._entitySetName;

    var lastActionAttr = formContext.getAttribute("duc_lastactiontaken");
    if (!lastActionAttr) return;
    var lastActionRef = lastActionAttr.getValue();
    if (!lastActionRef || lastActionRef.length === 0) return;

    var actionId = (lastActionRef[0].id || "").replace(/[{}]/g, "");
    var recordId = formContext.data.entity.getId().replace(/[{}]/g, "");

    Logger.trace("Action changed: " + (lastActionRef[0].name || "") + " | Record: " + recordId);

    // Set subject from Work Order name (mandatory on Activity entities)
    var subjectAttr = formContext.getAttribute("subject");
    if (subjectAttr && !subjectAttr.getValue()) {
        var regAttr = formContext.getAttribute("regardingobjectid");
        var regRef = regAttr ? regAttr.getValue() : null;
        subjectAttr.setValue(regRef && regRef.length > 0 ? (regRef[0].name || "Process Extension") : "Process Extension");
    }

    // Set simple fields via formContext (persisted on save at the end)
    var assignDateAttr = formContext.getAttribute("duc_assignmentdate");
    if (assignDateAttr) assignDateAttr.setValue(new Date());

    var commentAttr = formContext.getAttribute("duc_lastapprovercomment");
    if (commentAttr) commentAttr.setValue(null);

    Xrm.Utility.showProgressIndicator("Processing action...");

    var ActionService = DUC.ProcessAutomation.Offline.StepActionInfoService;
    var Assignment = DUC.ProcessAutomation.Offline.Assignment;
    var ctx = {};

    ActionService.getActionInfo(actionId)
        .then(function (actionInfo) {
            if (!actionInfo) throw new Error("Action info not found for " + actionId);
            ctx.actionInfo = actionInfo;
            Logger.trace("Action info: " + actionInfo.Action_Name);

            var authorAttr = formContext.getAttribute("duc_lastapprovedbyid");
            var authorRef = authorAttr ? authorAttr.getValue() : null;
            ctx.author = (authorRef && authorRef.length > 0)
                ? { entityType: authorRef[0].entityType, id: authorRef[0].id.replace(/[{}]/g, "") }
                : { entityType: "systemuser", id: Xrm.Utility.getGlobalContext().userSettings.userId.replace(/[{}]/g, "") };

            var regardingAttr = formContext.getAttribute("regardingobjectid");
            var regardingRef = regardingAttr ? regardingAttr.getValue() : null;
            ctx.regarding = (regardingRef && regardingRef.length > 0)
                ? { entityType: regardingRef[0].entityType, id: regardingRef[0].id.replace(/[{}]/g, "") }
                : null;

            if (actionInfo.RelatedStage_Sequence === 1) {
                var subDateAttr = formContext.getAttribute("duc_submissiondate");
                if (subDateAttr && !subDateAttr.getValue()) subDateAttr.setValue(new Date());
            }

            return DUC.ProcessAutomation.Offline._getParentRegarding(recordId);
        })
        .then(function (parentRegarding) {
            if (parentRegarding) ctx.regarding = parentRegarding;

            // Create action log
            return DUC.ProcessAutomation.Offline._createActionLog(formContext, ctx.author, ctx.regarding, ctx.actionInfo);
        })
        .then(function () {
            Logger.trace("Action log created");

            // Update extension lookup fields via WebApi
            var updateData = {};
            if (ctx.actionInfo.Action_NextStage_Id)
                updateData["duc_CurrentStage_duc_ProcessExtension@odata.bind"] = "/duc_processstages(" + ctx.actionInfo.Action_NextStage_Id + ")";
            if (ctx.author)
                updateData["duc_LastApprovedById_duc_ProcessExtension_" + ctx.author.entityType + "@odata.bind"] = "/" + esn(ctx.author.entityType) + "(" + ctx.author.id + ")";
            if (ctx.actionInfo.Action_Status_Id)
                updateData["duc_Status_duc_ProcessExtension@odata.bind"] = "/duc_processstatuses(" + ctx.actionInfo.Action_Status_Id + ")";
            if (ctx.actionInfo.Action_SubStatus_Id)
                updateData["duc_SubStatus_duc_ProcessExtension@odata.bind"] = "/duc_processsubstatuses(" + ctx.actionInfo.Action_SubStatus_Id + ")";
            updateData["duc_islastactiontakenoffline"] = true;

            return Xrm.WebApi.offline.updateRecord("duc_processextension", recordId, updateData);
        })
        .then(function () {
            Logger.trace("Extension fields updated");

            var naAttr = formContext.getAttribute("duc_nextassignee");
            var naRef = naAttr ? naAttr.getValue() : null;
            if (naRef && naRef.length > 0) {
                var naId = naRef[0].id.replace(/[{}]/g, "");
                return Xrm.WebApi.offline.updateRecord("duc_processextension", recordId, {
                    "ownerid@odata.bind": "/" + esn(naRef[0].entityType) + "(" + naId + ")"
                });
            }
            return Promise.resolve();
        })
        .then(function () {
            // Update target entity (Work Order status/substatus)
            if (!ctx.regarding) return Promise.resolve();
            return DUC.ProcessAutomation.Offline._updateTargetEntityFields(ctx.regarding, ctx.actionInfo);
        })
        .then(function () {
            // Assignment
            if (!ctx.regarding) return Promise.resolve();
            return Assignment.executeAssignmentLogic(recordId, ctx.regarding);
        })
        .then(function () {
            // Save form (persists formContext changes to local DB)
            if (formContext.data.entity.getIsDirty()) return formContext.data.entity.save();
            return Promise.resolve();
        })
        .then(function () {
            Logger.trace("All complete");
            Xrm.Utility.closeProgressIndicator();
        })
        .catch(function (error) {
            var detail = DUC.ProcessAutomation.Offline._extractErrorDetail(error);
            Logger.error("Action change failed: " + detail);
            Xrm.Utility.closeProgressIndicator();
            Xrm.Navigation.openAlertDialog({ text: "Processing error: " + detail, title: "Process Automation" });
        });
};

// ///////////////////////////////////////////////////////////////////////////
// SHARED HELPERS
// ///////////////////////////////////////////////////////////////////////////

DUC.ProcessAutomation.Offline._getParentRegarding = function (peId) {
    "use strict";
    return Xrm.WebApi.offline.retrieveMultipleRecords("duc_processextension",
        "?$select=duc_parentprocessextention&$filter=activityid eq '" + peId + "' and duc_parentprocessextention ne null"
    ).then(function (r) {
        if (r.entities.length === 0) return null;
        var parentRef = r.entities[0]["_duc_parentprocessextention_value"];
        if (!parentRef) return null;
        return Xrm.WebApi.offline.retrieveRecord("duc_processextension", parentRef,
            "?$select=_regardingobjectid_value,_regardingobjectid_value@Microsoft.Dynamics.CRM.lookuplogicalname");
    }).then(function (parent) {
        if (!parent || !parent["_regardingobjectid_value"]) return null;
        return { entityType: parent["_regardingobjectid_value@Microsoft.Dynamics.CRM.lookuplogicalname"] || "duc_processextension", id: parent["_regardingobjectid_value"] };
    }).catch(function () { return null; });
};

DUC.ProcessAutomation.Offline._createActionLog = function (formContext, author, regarding, actionInfo) {
    "use strict";
    var esn = DUC.ProcessAutomation.Offline._entitySetName;
    var userId = Xrm.Utility.getGlobalContext().userSettings.userId.replace(/[{}]/g, "");
    var logData = {};

    var subjectAttr = formContext.getAttribute("subject");
    if (subjectAttr && subjectAttr.getValue()) logData["subject"] = subjectAttr.getValue();

    var assignDateAttr = formContext.getAttribute("duc_assignmentdate");
    if (assignDateAttr && assignDateAttr.getValue()) logData["actualstart"] = assignDateAttr.getValue();

    var commentAttr = formContext.getAttribute("duc_lastapprovercomment");
    if (commentAttr && commentAttr.getValue()) logData["duc_comments"] = commentAttr.getValue();

    logData["duc_sendtocustomer"] = actionInfo.ActionType_SendToCustomer || false;
    logData["ownerid_duc_processactionlog@odata.bind"] = "/systemusers(" + userId + ")";

    if (author)
        logData["duc_AuthorId_duc_processActionLog_" + author.entityType + "@odata.bind"] = "/" + esn(author.entityType) + "(" + author.id + ")";
    if (regarding)
        logData["regardingobjectid_" + regarding.entityType + "_duc_processactionlog@odata.bind"] = "/" + esn(regarding.entityType) + "(" + regarding.id + ")";
    if (actionInfo.Action_Id)
        logData["duc_Action_duc_processActionLog@odata.bind"] = "/duc_stageactions(" + actionInfo.Action_Id + ")";
    if (actionInfo.Action_ActionType_Id)
        logData["duc_ActionType_duc_processActionLog@odata.bind"] = "/duc_actiontypes(" + actionInfo.Action_ActionType_Id + ")";
    if (actionInfo.Action_RelatedStage_Id)
        logData["duc_processStage_duc_processActionLog@odata.bind"] = "/duc_processstages(" + actionInfo.Action_RelatedStage_Id + ")";
    if (actionInfo.RelatedStage_RelatedProcess_Id)
        logData["duc_process_duc_processActionLog@odata.bind"] = "/duc_processdefinitions(" + actionInfo.RelatedStage_RelatedProcess_Id + ")";

    return Xrm.WebApi.offline.createRecord("duc_processactionlog", logData);
};

DUC.ProcessAutomation.Offline._updateTargetEntityFields = function (regarding, actionInfo) {
    "use strict";
    var esn = DUC.ProcessAutomation.Offline._entitySetName;
    var updateData = {};
    var hasUpdate = false;

    if (actionInfo.ProcessDefinition_TargetStatusLookupName && actionInfo.Status_Value) {
        var field = actionInfo.ProcessDefinition_TargetStatusLookupName;
        if (actionInfo.ProcessDefinition_StatusLookupType === 780500002) {
            updateData[field] = parseInt(actionInfo.Status_Value, 10);
        } else {
            updateData[field + "@odata.bind"] = "/" + esn(actionInfo.ProcessDefinition_ExistingStatusEntity) + "(" + actionInfo.Status_Value + ")";
        }
        hasUpdate = true;
    }

    if (actionInfo.ProcessDefinition_TargetSubStatusLookupName && actionInfo.SubStatus_Value) {
        var subField = actionInfo.ProcessDefinition_TargetSubStatusLookupName;
        if (actionInfo.ProcessDefinition_SubStatusLookupType === 780500002) {
            updateData[subField] = parseInt(actionInfo.SubStatus_Value, 10);
        } else {
            updateData[subField + "@odata.bind"] = "/" + esn(actionInfo.ProcessDefinition_ExistingSubStatusEntity) + "(" + actionInfo.SubStatus_Value + ")";
        }
        hasUpdate = true;
    }

    if (!hasUpdate) return Promise.resolve();
    return Xrm.WebApi.offline.updateRecord(regarding.entityType, regarding.id, updateData);
};

DUC.ProcessAutomation.Offline._extractErrorDetail = function (error) {
    "use strict";
    if (!error) return "Unknown error";
    if (typeof error === "string") return error;
    var parts = [];
    if (error.message) {
        try {
            var parsed = JSON.parse(error.message);
            if (parsed.error && parsed.error.message) parts.push(parsed.error.message);
            if (parsed.error && parsed.error.innererror && parsed.error.innererror.message) parts.push(parsed.error.innererror.message);
        } catch (e) { parts.push(error.message); }
    }
    if (error.errorCode) parts.push("Code: " + error.errorCode);
    if (error.innererror) {
        if (typeof error.innererror === "string") parts.push(error.innererror);
        else if (error.innererror.message) parts.push(error.innererror.message);
    }
    return parts.length > 0 ? parts.join(" | ") : String(error);
};
// ///////////////////////////////////////////////////////////////////////////
// MODULE 6: runActionLogic
// Direct-call version  no executionContext/formContext needed.
// Call this from button scripts AFTER duc_lastactiontaken is set via WebApi.
//
// Parameters:
//   actionId            duc_stageaction GUID that was set as last action
//   processExtensionId  duc_processextension record GUID
//   regarding           { entityType: "msdyn_workorder", id: "..." } or null
// ///////////////////////////////////////////////////////////////////////////

DUC.ProcessAutomation.Offline.runActionLogic = function (actionId, processExtensionId, regarding) {
    "use strict";

    var Logger = DUC.ProcessAutomation.Offline.Logger;
    var esn = DUC.ProcessAutomation.Offline._entitySetName;
    var ActionService = DUC.ProcessAutomation.Offline.StepActionInfoService;
    var Assignment = DUC.ProcessAutomation.Offline.Assignment;

    var userId = Xrm.Utility.getGlobalContext().userSettings.userId.replace(/[{}]/g, "");
    var ctx = {
        author: { entityType: "systemuser", id: userId },
        regarding: regarding || null
    };

    Logger.trace("runActionLogic: actionId=" + actionId + " peId=" + processExtensionId);

    return ActionService.getActionInfo(actionId)
        .then(function (actionInfo) {
            if (!actionInfo) throw new Error("Action info not found for " + actionId);
            ctx.actionInfo = actionInfo;
            Logger.trace("Action info loaded: " + actionInfo.Action_Name);
            return DUC.ProcessAutomation.Offline._getParentRegarding(processExtensionId)
                .then(function (parentRegarding) {
                    if (parentRegarding) ctx.regarding = parentRegarding;
                });
        })
        .then(function () {
            // Create action log (no form fields available  minimal data)
            var logData = {
                "subject": "Process Action",
                "actualstart": new Date(),
                "duc_sendtocustomer": ctx.actionInfo.ActionType_SendToCustomer || false,
                "duc_islastactiontakenoffline": true
            };
            logData["ownerid_duc_processactionlog@odata.bind"] = "/systemusers(" + userId + ")";
            logData["duc_AuthorId_duc_processActionLog_systemuser@odata.bind"] = "/systemusers(" + userId + ")";
            if (ctx.regarding)
                logData["regardingobjectid_" + ctx.regarding.entityType + "_duc_processactionlog@odata.bind"]
                    = "/" + esn(ctx.regarding.entityType) + "(" + ctx.regarding.id + ")";
            if (ctx.actionInfo.Action_Id)
                logData["duc_Action_duc_processActionLog@odata.bind"] = "/duc_stageactions(" + ctx.actionInfo.Action_Id + ")";
            if (ctx.actionInfo.Action_ActionType_Id)
                logData["duc_ActionType_duc_processActionLog@odata.bind"] = "/duc_actiontypes(" + ctx.actionInfo.Action_ActionType_Id + ")";
            if (ctx.actionInfo.Action_RelatedStage_Id)
                logData["duc_processStage_duc_processActionLog@odata.bind"] = "/duc_processstages(" + ctx.actionInfo.Action_RelatedStage_Id + ")";
            if (ctx.actionInfo.RelatedStage_RelatedProcess_Id)
                logData["duc_process_duc_processActionLog@odata.bind"] = "/duc_processdefinitions(" + ctx.actionInfo.RelatedStage_RelatedProcess_Id + ")";

            return Xrm.WebApi.offline.createRecord("duc_processactionlog", logData);
        })
        .then(function () {
            Logger.trace("Action log created");

            // Update process extension: next stage, approver, status, substatus
            var updateData = { "duc_islastactiontakenoffline": true };
            if (ctx.actionInfo.Action_NextStage_Id)
                updateData["duc_CurrentStage_duc_ProcessExtension@odata.bind"]
                    = "/duc_processstages(" + ctx.actionInfo.Action_NextStage_Id + ")";
            updateData["duc_LastApprovedById_duc_ProcessExtension_systemuser@odata.bind"]
                = "/systemusers(" + userId + ")";
            if (ctx.actionInfo.Action_Status_Id)
                updateData["duc_Status_duc_ProcessExtension@odata.bind"]
                    = "/duc_processstatuses(" + ctx.actionInfo.Action_Status_Id + ")";
            if (ctx.actionInfo.Action_SubStatus_Id)
                updateData["duc_SubStatus_duc_ProcessExtension@odata.bind"]
                    = "/duc_processsubstatuses(" + ctx.actionInfo.Action_SubStatus_Id + ")";

            return Xrm.WebApi.offline.updateRecord("duc_processextension", processExtensionId, updateData);
        })
        .then(function () {
            Logger.trace("Process extension updated");
            if (!ctx.regarding) return Promise.resolve();
            return DUC.ProcessAutomation.Offline._updateTargetEntityFields(ctx.regarding, ctx.actionInfo);
        })
        .then(function () {
            Logger.trace("Target entity fields updated  running assignment");
            if (!ctx.regarding) return Promise.resolve();
            return Assignment.executeAssignmentLogic(processExtensionId, ctx.regarding);
        })
        .then(function () {
            Logger.trace("runActionLogic complete for action: " + actionId);
        })
        .catch(function (error) {
            var detail = DUC.ProcessAutomation.Offline._extractErrorDetail(error);
            Logger.error("runActionLogic failed: " + detail);
            Xrm.Navigation.openAlertDialog({
                text: "Offline action processing error:\n" + detail,
                title: "Process Automation"
            });
        });
};
