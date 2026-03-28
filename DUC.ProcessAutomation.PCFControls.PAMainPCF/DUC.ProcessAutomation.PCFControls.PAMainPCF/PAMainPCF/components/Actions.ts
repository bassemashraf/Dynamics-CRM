import * as React from "react";
import { PrimaryButton } from "@fluentui/react";
import { useState, useEffect } from "react";
import ModalDialog from "./ModalDialog";
import AssignDialog from "./AssignDialog";
import { IInputs, IOutputs } from "../generated/ManifestTypes";
import {
  Constants,
  getCertificateFetchXML,
  getReportViewerUrl,
  getValue,
} from "../constants"; // Importing constants
import { FormContextHelper } from "./FormContextHelper";

interface RoleAssociation {
  name: string;
}

export interface IActionButtonProps {
  buttonId?: string;
  displayName?: string;
  buttonType?: string;
  buttonStyle?: string;
  buttonColor?: string;
  buttonIcon?: string;
  requireComments?: boolean;
  requireAssign?: boolean;
  buttonStatus?: boolean;
  requiresSurvey?: boolean;
  sendToCustomer?: boolean;
  entity?: ComponentFramework.WebApi.Entity;
  staticReplyTemplateId?: string;
  nextTeamId?: string;
}

export interface IActionsProps {
  _context: ComponentFramework.Context<IInputs>;
  onnotifyOutputChanged: () => void;
}

export const Actions: React.FC<IActionsProps> = ({
  _context,
  onnotifyOutputChanged,
}) => {
  const [isLoading, setIsLoading] = useState(false);
  const [isModalVisible, setIsModalVisible] = useState(false);
  const [isAssignModalVisible, setIsAssignModalVisible] = useState(false);
  const [modalContent, setModalContent] = useState<{
    action: IActionButtonProps;
  }>({ action: {} });
  const [dataSet, setDataSet] = useState<IActionButtonProps[]>([]);

  const [isAuthorized, setIsAuthorized] = useState(false);
  const [isPermissionLoading, setIsPermissionLoading] = useState(true);

  useEffect(() => {
    void evaluatePermissions();
  }, []);

  useEffect(() => {
    const fetchDataAsync = async () => {
      try {
        await fetchData();
      } catch (error: any) {
        console.error("Error fetching data:", error);
      }
    };

    fetchDataAsync().catch((error: any) => {
      console.error("Error in fetchDataAsync:", error);
    });
  }, []);

  // permission check function
  const evaluatePermissions = async () => {
    setIsPermissionLoading(true);
    try {
      // Offline: skip permission checks — system entities (teammembership, role, etc.) are not available
      const isOffline = _context.client.isOffline?.() === true;
      if (isOffline) {
        console.log("[evaluatePermissions] Offline mode — auto-authorizing");
        setIsAuthorized(true);
        setIsPermissionLoading(false);
        return;
      }

      const currentUserId = _context.userSettings.userId
        .replace(/[{}]/g, "")
        .toLowerCase();
      const ownerVal = FormContextHelper.getLookupValue('ownerid');
      const ownerId = ownerVal?.id?.toLowerCase().replace(/[{}]/g, "") ?? "";
      const ownerType = ownerVal?.entityType ?? "";

      // Case 1: Current user is the record owner
      const isUserOwner =
        ownerType === "systemuser" && currentUserId === ownerId;
      console.log(`Case 1 - Is current user the record owner? ${isUserOwner}`);

      // Case 2: Current user is System Administrator
      let isAdmin = false;
      let roles: string[] = [];

      try {
        const userRoles = await _context.webAPI.retrieveMultipleRecords(
          "systemuser",
          `?$filter=systemuserid eq ${currentUserId}&$expand=systemuserroles_association($select=name)`,
        );

        if (userRoles.entities.length > 0) {
          const roleAssociations: RoleAssociation[] =
            (userRoles.entities[0] as any).systemuserroles_association ?? [];
          roles = roleAssociations.map((r) => r.name).filter(Boolean);
        }

        isAdmin = roles.includes("System Administrator");
        console.log(
          `Case 2 - Is current user System Administrator? ${isAdmin}`,
        );
      } catch (e: any) {
        const errMsg = `[evaluatePermissions] Error fetching roles.\nError: ${e?.message || String(e)}`;
        console.error(errMsg, e);
        void _context.navigation.openAlertDialog({ text: errMsg });
      }

      // Case 3: Record is owned by a team that current user is a member of
      let isTeamMember = false;
      if (ownerType === "team" && ownerId) {
        try {
          const membershipResult =
            await _context.webAPI.retrieveMultipleRecords(
              "teammembership",
              `?$filter=systemuserid eq '${currentUserId}' and teamid eq '${ownerId}'`,
            );

          isTeamMember = membershipResult.entities.length > 0;
          console.log(
            `Case 3 - Is user member of owning team? ${isTeamMember}`,
          );
        } catch (e: any) {
          const errMsg = `[evaluatePermissions] Error checking team membership.\nError: ${e?.message || String(e)}`;
          console.error(errMsg, e);
          void _context.navigation.openAlertDialog({ text: errMsg });
        }
      }

      const authorized = isUserOwner || isAdmin || isTeamMember;
      setIsAuthorized(authorized);

      console.log(
        `User is ${authorized ? "authorized" : "not authorized"} for actions`,
      );
    } catch (error: any) {
      const errMsg = `[evaluatePermissions] Permission check failed.\nError: ${error?.message || String(error)}`;
      console.error(errMsg, error);
      void _context.navigation.openAlertDialog({ text: errMsg });
      setIsAuthorized(false);
    } finally {
      setIsPermissionLoading(false);
    }
  };

  const validateAction = async (
    actionId: string,
    processExtensionId: string,
    lang: string,
  ): Promise<boolean> => {
    let succeeded = false;
    const parameters = {
      processExtensionId: processExtensionId,
      actionId: actionId.replace("{", "").replace("}", ""),
      lang: lang,
    };

    try {
      setIsLoading(true);

      const isOffline = _context.client.isOffline?.() === true;
      if (isOffline) {
        console.log("[validateAction] Offline mode — skipping online validation");
        return true;
      }

      const result = await Xrm.WebApi.online.execute({
        processExtensionId: parameters.processExtensionId,
        actionId: parameters.actionId,
        lang: parameters.lang,

        getMetadata: function () {
          return {
            boundParameter: null,
            parameterTypes: {
              processExtensionId: {
                typeName: "Edm.String",
                structuralProperty: 1,
              },
              actionId: {
                typeName: "Edm.String",
                structuralProperty: 1,
              },
              lang: {
                typeName: "Edm.String",
                structuralProperty: 1,
              },
            },
            operationType: 0,
            operationName: Constants.VERIFY_ACTION_NAME.toString(),
          };
        },
      });
      // Parse the response
      const response = await result.json();
      succeeded = String(response.succeeded).toLowerCase() === "true";
      console.log(succeeded);

      if (succeeded) {
        if (response.warnMsg) {
          await _context.navigation.openAlertDialog({
            text: response.warnMsg,
          });
        }
        return true;
      } else {
        await _context.navigation.openAlertDialog({
          text: response.errorMsg,
        });
        return false;
      }
    } catch (error: unknown) {
      if (error instanceof Error) {
        console.error("Validation Error:", error.message);
        succeeded = false;
      } else {
        console.error("An unknown error occurred:", error);
        succeeded = false;
      }
    } finally {
      setIsLoading(false);
    }

    return succeeded;
  };

  const isMobile = function (
    _context: ComponentFramework.Context<IInputs, ComponentFramework.IEventBag>,
  ) {
    const formFactor = _context.client.getFormFactor();

    const isPhone = formFactor === Constants.isPhone;
    const isTablet = formFactor === Constants.isTablet;

    return isPhone || isTablet;
  };

  const fetchData = async () => {
    let currentQuery = "";
    try {
      const stepVal = FormContextHelper.getLookupValue('duc_currentstage');
      if (!stepVal || !stepVal.id)
        return;
      setIsLoading(true);
      const isMob = isMobile(_context);
      const stepId = stepVal.id;

      const Filter = isMob
        ? Constants.STAGE_ACTION_MOBILE_FILTER
        : Constants.STAGE_ACTION_DEFAULT_FILTER;

      currentQuery = Filter + Constants.STAGE_ACTION_QUERY;

      const response = await _context.webAPI.retrieveMultipleRecords(
        Constants.ENTITY_NAME,
        currentQuery,
      );

      // Local runtime filtering of the actions based on related stage ID
      const filteredEntities = response.entities.filter((entity) => {
        const relatedStageId = entity["_duc_relatedstage_value"] || entity["duc_relatedstage"];
        return relatedStageId && relatedStageId.toLowerCase() === stepId.toLowerCase();
      });

      console.log(Constants.MSG_PREFIX + filteredEntities.length);

      const tempDataSet: IActionButtonProps[] = [];

      // Collect all unique ActionType IDs to fetch metadata for buttons
      const actionTypeIds: string[] = [];
      filteredEntities.forEach((entity) => {
        const typeId = entity["duc_actiontype"] || entity["_duc_actiontype_value"];
        if (typeId && !actionTypeIds.includes(typeId)) {
          actionTypeIds.push(typeId);
        }
      });

      // Fetch ActionType details explicitly to bypass offline $expand limits
      const actionTypeMap: Record<string, any> = {};
      if (actionTypeIds.length > 0) {
        let typeFilter = "";
        actionTypeIds.forEach((id, index) => {
          if (index > 0) typeFilter += " or ";
          typeFilter += `duc_actiontypeid eq ${id}`;
        });

        const typeQuery = `?$filter=${typeFilter}&$select=duc_isassignaction,duc_mainpcfcontroltype,duc_wfaction,duc_actioncommand,duc_color,duc_icon,duc_sendtocustomer`;
        try {
          const typeResponse = await _context.webAPI.retrieveMultipleRecords(
            Constants.ACTION_TYPE_ENTITY_NAME,
            typeQuery
          );

          typeResponse.entities.forEach(typeEnt => {
            actionTypeMap[typeEnt.duc_actiontypeid] = typeEnt;
          });
        } catch (e: any) {
          console.error("Action Type Fetch Error", e);
        }
      }

      filteredEntities.forEach((entity) => {
        const typeId = entity["duc_actiontype"] || entity["_duc_actiontype_value"];
        const actionType = typeId ? actionTypeMap[typeId] : null;

        const prop: IActionButtonProps = {
          displayName:
            getValue(
              entity,
              _context.resources.getString(Constants.DISPLAY_NAME),
            ) ?? "",
          buttonIcon: actionType ? getValue(actionType, Constants.BUTTON_ICON) : "",
          buttonColor: actionType ? getValue(actionType, Constants.BUTTON_COLOR) : "",
          buttonStatus: getValue(entity, Constants.BUTTON_STATUS) === "true",
          buttonId: getValue(entity, Constants.BUTTON_ID) ?? "",
          requireComments:
            getValue(entity, Constants.REQUIRE_COMMENTS) === "true",
          requireAssign: actionType ?
            getValue(actionType, Constants.IS_ASSIGN_ACTION_TYPE) === "true" : false,
          requiresSurvey:
            getValue(entity, Constants.REQUIRES_SURVEY) === "true",
          sendToCustomer: actionType ?
            getValue(actionType, Constants.SEND_TO_CUSTOMER) === "true" : false,
          nextTeamId: entity["duc_nextstage"] ?? "",
          entity: { ...entity, duc_actiontype: actionType }, // Embed it for evaluating later code
          staticReplyTemplateId: entity["duc_relatedstaticresponsestemplate"] ?? "",
        };
        tempDataSet.push(prop);
      });
      setDataSet(tempDataSet);
    } catch (error: unknown) {
      const errMsg = `[fetchData] Error occurred in API call.\nError: ${error instanceof Error ? error.message : String(error)}`;
      console.error(errMsg, error);
    } finally {
      setIsLoading(false);
    }
  };

  function generateGuid(): string {
    // Generates a random UUID v4
    return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(
      /[xy]/g,
      function (c) {
        const r = (Math.random() * 16) | 0;
        const v = c === "x" ? r : (r & 0x3) | 0x8;
        return v.toString(16);
      },
    );
  }

  const uploadFilesToNotes = async (
    context: ComponentFramework.Context<IInputs>,
    parentEntityLogicalName: string,
    parentRecordId: string,
    files: File[],
    notetext: string,
  ) => {
    for (const file of files) {
      const base64 = await fileToBase64(file);
      const annotation = {
        subject: file.name,
        filename: file.name,
        notetext: notetext,
        documentbody: base64.split(",")[1],
        mimetype: file.type,
        [`objectid_${parentEntityLogicalName}@odata.bind`]: `/${parentEntityLogicalName}s(${parentRecordId})`,
      };

      await context.webAPI.createRecord("annotation", annotation);
    }
  };

  const fileToBase64 = (file: File): Promise<string> =>
    new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result as string);
      reader.onerror = reject;
      reader.readAsDataURL(file);
    });

  const updateFieldValues = (
    actionId: string,
    comment: string,
    notetext: string,
    assignee?: {
      id: string;
      name: string;
      type: "user" | "team";
      entityType: string;
    },
  ) => {
    if (actionId != null) {
      console.log(actionId);
      FormContextHelper.setLookupValue('duc_lastactiontaken', actionId, actionId, Constants.ENTITY_NAME.toString());
      if (_context.parameters.boundStringField) {
        _context.parameters.boundStringField.raw = actionId;
      }
      FormContextHelper.setStringValue('duc_lastapprovercomment', comment);
      FormContextHelper.setStringValue('duc_lastattachmentsguid', notetext);
      if (assignee != undefined) {
        FormContextHelper.setLookupValue('duc_nextassignee', assignee.id, assignee.name, assignee.entityType);
      }
    }
    onnotifyOutputChanged();
  };

  const openSurveyModal = async (
    actionId: string,
    actionSurvey: string,
    serviceRequestId: string,
  ) => {
    const Data = {
      ActionId: actionId,
      SurveyId: actionSurvey,
      ServiceRequestId: serviceRequestId,
    };

    const pageInput: Xrm.Navigation.PageInputHtmlWebResource = {
      pageType: "webresource",
      webresourceName:
        _context.parameters.SurveyFormWRName?.raw ??
        Constants.SURVEY_FORM_WR_NAME,
      data: JSON.stringify(Data),
    };

    const navigationOptions: Xrm.Navigation.NavigationOptions = {
      target: 2,
      width: { value: 700, unit: "px" },
      height: { value: 1000, unit: "px" },
      position: 1,
    };

    try {
      await Xrm.Navigation.navigateTo(pageInput, navigationOptions).then(
        () => {
          console.log("Survey form opened successfully.");
          return;
        },
        (error: any) => {
          console.error("Error opening survey form:", error);
          throw error;
        },
      );
    } catch (error: any) {
      console.error("Error in openSurveyModal:", error);
    }
  };

  const SetShowSurvey = () => {
    console.log("SetShowSurvey called");
    if (_context.parameters.ShowSurveyForm) {
      _context.parameters.ShowSurveyForm.raw = true;
    }

    console.log("SetShowSurvey updated");
    onnotifyOutputChanged();
  };

  const saveModal = async (
    action: IActionButtonProps,
    comment: string,
    files?: File[],
  ) => {
    if (comment == "") {
      const msg = _context.resources.getString(Constants.SPECIFY_COMMENT_MSG);
      return _context.navigation.openAlertDialog({ text: msg });
    }
    console.log("Saving data");
    await onClick(action, comment, files);
  };
  const saveAssignModal = async (
    action: IActionButtonProps,
    assignee: {
      id: string;
      name: string;
      type: "user" | "team";
      entityType: string;
    },
  ) => {
    console.log("Saving data");
    await onClick(action, "", undefined, assignee);
  };
  const ShowReport = async (srNumber: string) => {
    await Xrm.WebApi.retrieveMultipleRecords(
      Constants.MAIN_ENTITY_NAME,
      `?fetchXml=${getCertificateFetchXML(srNumber)}`,
    )
      .then(function success(result) {
        if (result.entities.length == 0) return;

        const ent = result.entities[0] as Record<string, unknown>;
        const url =
          Xrm.Utility.getGlobalContext().getClientUrl() +
          getReportViewerUrl(ent);

        const windowFeatures = "left=450,top=100,width=600,height=600";
        window.open(url, undefined, windowFeatures);
        return;
      })
      .catch((error) => {
        console.error("Error retrieving records:", error);
      });
  };

  const onClick = async (
    action: IActionButtonProps,
    comment: string,
    files?: File[],
    assignee?: {
      id: string;
      name: string;
      type: "user" | "team";
      entityType: string;
    },
  ) => {
    if (!action?.entity || !action?.buttonId) return;
    let isValid = true;

    if (!isModalVisible) {
      const recordId = (_context.mode as any).contextInfo.entityId;
      isValid = await validateAction(
        action.buttonId ?? "",
        recordId,
        _context.userSettings.languageId?.toString(),
      );
    }
    if (!isValid || !action.buttonStatus) {
      return;
    }

    if (!isModalVisible && action.requireComments) {
      setModalContent({ action: action });
      setIsModalVisible(true);
      return;
    }
    if (!isAssignModalVisible && action.requireAssign) {
      setModalContent({ action: action });
      setIsAssignModalVisible(true);
      return;
    } else if (action.requiresSurvey) {
      const surveyId = getValue(action.entity, Constants.ACTION_SURVEY);
      const serviceRequestId = FormContextHelper.getStringValue("activityid") || FormContextHelper.getFormContext()?.data?.entity?.getId()?.replace(/[{}]/g, "") || "";
      console.log("Before call SetShowSurvey");
      SetShowSurvey();

      return;
    } else {
      if (isModalVisible == true) {
        setIsModalVisible(false);
      }

      if (isAssignModalVisible == true) {
        setIsAssignModalVisible(false);
      }

      const isExecuteAction =
        getValue(action.entity, "duc_actiontype." + Constants.WFACTION) == "false";
      if (isExecuteAction || !isExecuteAction) {
        const codeFromField = action.entity?.duc_actiontype?.duc_actioncommand ?? "";

        if (codeFromField) {
          console.log(`[DEBUG] About to run JS code from duc_actioncommand:`, codeFromField);
          try {
            const updatedCode = codeFromField.replace(
              "#StageActionId#",
              action.buttonId,
            );

            const runCode = new Function(
              "languageId",
              "regardingField",
              "processExtension",
              "action",
              `"use strict"; return (async () => { ${updatedCode} })();`,
            );

            // Run the code dynamically
            await runCode(
              _context.userSettings.languageId,
              FormContextHelper.getLookupValue('regardingobjectid'),
              (_context.mode as any).contextInfo,
              action,
            );
            console.log("[DEBUG] JS code executed successfully.");
          } catch (error: any) {
            console.error(`[DEBUG] JS code execution FAILED:`, error);
          }
        } else {
          console.log("[DEBUG] No duc_actioncommand code found on this action.");
        }
      } else {
        try {
          const result = await _context.navigation.openConfirmDialog({
            title: _context.resources.getString(Constants.CONFIRM_ACTION_TITLE),
            text: _context.resources.getString(Constants.CONFIRM_ACTION_MSG),
          });

          if (result.confirmed) {
            const notetext = generateGuid();
            if (files && files.length > 0) {
              const primaryKey = FormContextHelper.getStringValue("activityid") || FormContextHelper.getFormContext()?.data?.entity?.getId()?.replace(/[{}]/g, "");

              if (primaryKey) {
                await uploadFilesToNotes(
                  _context,
                  Constants.MAIN_ENTITY_NAME,
                  primaryKey,
                  files,
                  notetext,
                );
              }
            }
            updateFieldValues(
              action.buttonId ?? "",
              comment,
              notetext,
              assignee,
            );
          }
        } catch (error) {
          console.error("Confirmation action failed:", error);
          throw error;
        }
      }
    }
  };

  const closeModal = () => {
    setIsModalVisible(false);
  };

  const closeAssignModal = () => {
    setIsAssignModalVisible(false);
  };

  return React.createElement(
    "div",
    { className: "buttonIcons" },
    isPermissionLoading
      ? React.createElement(
        "div",
        { className: "loading-spinner" },
        React.createElement("span", null, _context.resources.getString(Constants.LOADING))
      )
      : !isAuthorized
        ? React.createElement(
          "div",
          { className: "no-actions-message" },
          _context.resources.getString(Constants.NOT_AUTHORIZED_MSG)
        )
        : !isLoading && dataSet.length === 0
          ? React.createElement(
            "div",
            { className: "no-actions-message" },
            _context.resources.getString(Constants.NO_RECORDS)
          )
          : dataSet.map((item) =>
            React.createElement(
              PrimaryButton,
              {
                className: "btnaction",
                key: item.buttonId,
                onClick: () => void onClick(item, ""),
                disabled: !item.buttonStatus,
                style: {
                  backgroundColor: item.buttonColor,
                  borderColor: item.buttonColor,
                },
              },
              item.buttonIcon ? React.createElement("i", { className: item.buttonIcon, style: { marginRight: "8px" } }) : null,
              item.displayName
            )
          ),
    isLoading && React.createElement(
      "div",
      { className: "loading-spinner" },
      React.createElement("span", null, _context.resources.getString(Constants.LOADING))
    ),
    React.createElement(ModalDialog, {
      _context: _context,
      isVisible: isModalVisible,
      onClose: closeModal,
      onSave: saveModal,
      action: modalContent.action,
    }),
    React.createElement(AssignDialog, {
      _context: _context,
      isVisible: isAssignModalVisible,
      onClose: closeAssignModal,
      onSetNextAssigee: saveAssignModal,
      action: modalContent?.action ?? {},
      teamId: modalContent?.action.nextTeamId ?? undefined,
    })
  );
};
