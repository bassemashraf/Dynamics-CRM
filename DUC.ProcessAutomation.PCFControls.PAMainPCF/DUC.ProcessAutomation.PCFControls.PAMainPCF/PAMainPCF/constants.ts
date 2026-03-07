export const Constants = {
  isPhone: 3,
  isTablet: 2,
  // ==========================
  // 🔹 Entity Logical Names
  // ==========================
  ENTITY_NAME: "duc_stageaction",
  MAIN_ENTITY_NAME: "duc_processextension",
  STAGE_ENTITY_NAME: "duc_processstage",
  SERVICE_REQUEST_LOG_ENTITY: "duc_processactionlog",
  STATIC_REPLY_ENTITY: "duc_processstaticresponse",
  EMPLOYEE_LEAVE_ENTITY: "duc_employeeleave",
  SERVICE_REQUEST_TYPE_ENTITY: "duc_processdefinition",
  SERVICE_REQUEST_TYPE_CERTIFICATIONS_ENTITY:
    "duc_processdefinitioncertifications",
  REPORT_ENTITY: "duc_report",
  CUSTOMER_LICENSES_ENTITY: "duc_customerlicenses",

  // ==========================
  // 🔹 API Filter Strings
  // ==========================
  STAGE_ACTION_DEFAULT_FILTER:
    "?$filter=(duc_visible eq true and _duc_relatedstage_value eq {0})",

  STAGE_ACTION_MOBILE_FILTER:
    "?$filter=(duc_showonmobile eq true and _duc_relatedstage_value eq {0})",

  STAGE_ACTION_QUERY:
    "&$orderby=duc_sequence asc" +
    "&$expand=" +
    "duc_ActionType($select=duc_isassignaction,duc_mainpcfcontroltype,duc_wfaction,duc_actioncommand,duc_color,duc_icon,duc_sendtocustomer)," +
    "duc_RelatedStage($select=duc_sequence)," +
    "duc_NextStage($select=_duc_assignedteam_value)",

  STAGE_ACTION_MOBILE_VISIBLE_FIELD: "duc_showonmobile",

  GENERIC_CANCEL_FILTER:
    "?$filter=duc_stageactionid eq {0}" +
    "&$orderby=duc_sequence asc" +
    "&$expand=" +
    "duc_ActionType($select=duc_isassignaction,duc_mainpcfcontroltype,duc_wfaction,duc_actioncommand,duc_color,duc_icon,duc_sendtocustomer)",

  STAGE_VISIBLE_FIELD: "duc_visible",
  STAGE_MOBILE_VISIBLE_FIELD: "duc_visibleonmobile",
  STAGE_SELECT:
    "?$select=duc_visible,duc_visibleonmobile,duc_sequence,duc_processstageid,duc_name,duc_arabicname,duc_arabicdescription,duc_description,duc_sequenceoverride",
  STAGE_FILTER:
    "&$filter=_duc_relatedprocess_value eq {0} and statecode eq 0 and ({2} eq true {1})",
  STAGE_ORDER: "&$orderby=duc_sequence,duc_sequenceoverride asc",
  // ==========================
  // 🔹 Field Names
  // ==========================
  BUTTON_ICON: "duc_ActionType.duc_icon",
  REQUIRE_COMMENTS: "duc_requirescomments",
  IS_ASSIGN_ACTION_TYPE: "duc_ActionType.duc_isassignaction",
  ACTION_CODE: "duc_ActionType.duc_actioncommand",
  WFACTION: "duc_ActionType.duc_wfaction",
  BUTTON_COLOR: "duc_ActionType.duc_color",
  BUTTON_STATUS: "duc_enabled",
  BUTTON_ID: "duc_stageactionid",
  SEND_TO_CUSTOMER: "duc_ActionType.duc_sendtocustomer",
  NEXT_STEP_TEAM: "duc_NextStage._duc_assignedteam_value",
  STATIC_REPLY_TEMPLATE_ID: "_duc_relatedstaticresponsestemplate_value",
  ACTION_LOG_USER_COMMENTS: "duc_comments",

  // 🔹 ModalDialog specific fields
  SR_LOG_ID: "duc_processactionlogid",
  SR_LOG_CONTAINS_CUSTOMER_COMMENTS: "duc_containscustomercomments",
  SR_LOG_USER_COMMENTS: "duc_comments",
  STATIC_REPLY_ID: "duc_processstaticresponseid",
  STATIC_REPLY_NAME_AR: "duc_replybodyar",
  STATIC_REPLY_NAME_EN: "duc_replybodyen",
  STATIC_REPLY_TEMPLATE_LOOKUP: "_duc_template_value",
  STATIC_REPLY_IS_ACTIVE: "duc_isactive",
  CURRENT_STEP_SEQUENCE: "duc_RelatedStage.duc_sequence",
  REQUIRES_SURVEY: "duc_requiressurvey",
  ACTION_SURVEY: "_duc_actionsurvey_value",

  // ==========================
  // 🔹 Localized Resource Keys
  // ==========================
  DISPLAY_NAME: "ActionsDisplayName",
  NO_RECORDS: "NoRecords",
  NOT_AUTHORIZED_MSG: "NOT_AUTHORIZED_MSG",
  LOADING: "Loading",
  CONFIRM_ACTION_TITLE: "ConfirmActionTitle",
  WARNING_ALERT_TITLE: "WarningAlertTitle",
  CONFIRM_ACTION_MSG: "ConfirmActionMsg",
  SPECIFY_COMMENT_MSG: "MsgSpecifyComment",
  ENTER_COMMENTS_MSG: "EnterCommentsMessage",
  CLOSE_MODAL: "CloseModal",
  CANCEL_BUTTON_LABEL: "Cancel",
  NO_STATIC_REPLIES: "NoStaticReplies",
  LOADING_STATIC_REPLY_LABEL: "LoadingStaticReplies",
  ADD_PREVIOUS_COMMENTS_TITLE: "AddPreviousCommentsTitle",
  NO_CUSTOMER_COMMENTS: "NoCustomerComments",
  SELECT_ALL_LABEL: "SelectAllLabel",
  COMMENT_NUMBER_LABEL: "CommentNumberLabel",
  SELECTED_COUNT_LABEL: "SelectedCountLabel",
  ADD_SELECTED_COMMENTS_LABEL: "AddSelectedCommentsLabel",
  LOADING_COMMENTS_LABEL: "LoadingCommentsLabel",
  PROCEED_WITHOUT_COMMENTS_LABEL: "ProceedWithoutCommentsLabel",
  STATIC_REPLY_BODY_COLNAME: "staticReplyBodyColName",
  STATIC_REPLY_TITLE_COLNAME: "staticReplytitleColName",
  ATTACH_FILE_LABEL: "LBL_ATTACHMENT",
  DETAILS_BTN_TEXT: "DetailsBtnText",
  OWNING_TEAM_LABEL: "OwningTeamLabel",
  CHECKLIST_BTN_LABEL: "CheckListBtnLbl",
  ATTACHMENTS_BTN_LABEL: "AttachmentsBtnLbl",
  ASSIGNMENT_BTN_LABEL: "AssignmentBtnLbl",

  // ==========================
  // 🔹 Miscellaneous
  // ==========================
  NS: "NS",
  SURVEY_FORM_WR_NAME: "",
  SURVEY_RESPONSE_VIEW_WR_NAME: "SurveyResponseViewWRName",
  MSG_PREFIX: "Records retrieved with ID: ",
  VERIFY_ACTION_NAME: "duc_PAVerifyAction",
};

// ==========================
// 🔹 Utility Functions
// ==========================

export const getValue = (
  record: ComponentFramework.WebApi.Entity,
  fieldName: string | undefined,
): string => {
  let result = "";
  let color: ComponentFramework.WebApi.Entity = record;

  if (fieldName?.length) {
    const array = fieldName.split(".");

    array.forEach((element, i) => {
      if (color?.[element] == null) {
        result = "";
      } else if (i + 1 === array.length) {
        result = color[element]?.toString() ?? "";
      } else {
        color = color[element];
      }
    });
  }
  return result;
};

// ==========================
// 🔹 FetchXML Builders
// ==========================

export const getAssignFetchXML = (userId: string): string => `
  <fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false" aggregate="true">
    <entity name="${Constants.MAIN_ENTITY_NAME}">
      <attribute name="activityid" aggregate="count" alias="request_count" />
      <filter type="and">
        <condition attribute="ownerid" operator="eq" value="${userId}" />
        <condition attribute="statecode" operator="eq" value="0" />
      </filter>
    </entity>
  </fetch>
`;

export const getAssignedTodayFetchXML = (
  userId: string,
  todayStr: string,
): string => `
  <fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false" aggregate="true">
    <entity name="${Constants.MAIN_ENTITY_NAME}">
      <attribute name="activityid" aggregate="count" alias="today_count" />
      <filter type="and">
        <condition attribute="ownerid" operator="eq" value="${userId}" />
        <condition attribute="statecode" operator="eq" value="0" />
        <condition attribute="modifiedon" operator="on" value="${todayStr}" />
      </filter>
    </entity>
  </fetch>
`;

export const getCheckOnLeaveFetchXML = (
  userId: string,
  todayStr: string,
): string => `
  <fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false">
    <entity name="${Constants.EMPLOYEE_LEAVE_ENTITY}">
      <attribute name="duc_employeeleaveid" />
      <filter type="and">
        <condition attribute="duc_startdate" operator="on-or-before" value="${todayStr}" />
        <condition attribute="duc_enddate" operator="on-or-after" value="${todayStr}" />
        <condition attribute="duc_employee" operator="eq" uitype="systemuser" value="${userId}" />
      </filter>
    </entity>
  </fetch>
`;

export const getCertificateFetchXML = (srNumber: string): string => `
  <fetch top="1">
    <entity name="${Constants.MAIN_ENTITY_NAME}">
      <attribute name="activityid" />
      <filter>
        <condition attribute="subject" operator="eq" value="${srNumber}" />
      </filter>
      <link-entity name="${Constants.SERVICE_REQUEST_TYPE_ENTITY}" from="duc_processdefinitionid" to="${Constants.SERVICE_REQUEST_TYPE_ENTITY}" link-type="inner" alias="SRType">
        <attribute name="duc_processdefinitionid" />
        <link-entity name="${Constants.SERVICE_REQUEST_TYPE_CERTIFICATIONS_ENTITY}" from="duc_processdefinitions" to="duc_processdefinitionid" link-type="inner" alias="Certificate">
          <attribute name="duc_processdefinitions" />
          <link-entity name="${Constants.REPORT_ENTITY}" from="duc_reportid" to="duc_reporttemplate" link-type="inner" alias="Report">
            <attribute name="duc_id" />
            <attribute name="duc_name" />
          </link-entity>
        </link-entity>
      </link-entity>
      <link-entity name="${Constants.CUSTOMER_LICENSES_ENTITY}" from="${Constants.MAIN_ENTITY_NAME}" to="activityid" alias="License">
        <attribute name="duc_customerlicensesid" />
      </link-entity>
    </entity>
  </fetch>
`;
// FetchXML for service request logs with customer comments
export const getServiceRequestLogFetchXML = (primaryKey: string) => `
?$select=${Constants.SR_LOG_ID},${Constants.SR_LOG_CONTAINS_CUSTOMER_COMMENTS},${Constants.SR_LOG_USER_COMMENTS}
&$filter=(_duc_processextension_value eq ${primaryKey} and ${Constants.SR_LOG_CONTAINS_CUSTOMER_COMMENTS} eq true)
`;

// FetchXML for static replies by template
export const getStaticRepliesFetchXML = (templateId: string) => `
?$select=${Constants.STATIC_REPLY_ID},${Constants.STATIC_REPLY_NAME_AR},${Constants.STATIC_REPLY_NAME_EN}
&$filter=(${Constants.STATIC_REPLY_TEMPLATE_LOOKUP} eq ${templateId} and ${Constants.STATIC_REPLY_IS_ACTIVE} eq true)
`;

// ==========================
// 🔹 CRM Report Viewer URL
// ==========================
export const getReportViewerUrl = (ent: Record<string, unknown>): string => {
  const reportName = ent["Report.duc_name"];
  const reportId = ent["Report.duc_id"];
  const licenseId = ent["License.duc_customerlicensesid"];

  // Validate they are defined and of the correct type
  if (
    typeof reportName !== "string" ||
    typeof reportId !== "string" ||
    typeof licenseId !== "string"
  ) {
    console.warn("Missing or invalid report parameters:", {
      reportName,
      reportId,
      licenseId,
    });
    return "";
  }

  return `/crmreports/viewer/viewer.aspx?action=filter&helpID=${encodeURIComponent(reportName)}&id=${encodeURIComponent(reportId)}&p:CustomerLicenseId=${encodeURIComponent(licenseId)}`;
};
