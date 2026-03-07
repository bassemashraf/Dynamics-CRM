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
  }, [_context.parameters.stepLookup, _context.parameters.OwnerField]);

  useEffect(() => {
    const fetchDataAsync = async () => {
      try {
        await fetchData();
      } catch (error) {
        console.error("Error fetching data:", error);
        alert("Error fetching data: " + (error instanceof Error ? error.message : String(error)));
      }
    };

    fetchDataAsync().catch((error) => {
      console.error("Error in fetchDataAsync:", error);
      alert("Error in fetchDataAsync: " + (error instanceof Error ? error.message : String(error)));
    });
  }, [_context.parameters.stepLookup]);

  // permission check function
  const evaluatePermissions = async () => {
    setIsPermissionLoading(true);
    try {
      const currentUserId = _context.userSettings.userId
        .replace(/[{}]/g, "")
        .toLowerCase();
      const ownerId =
        _context.parameters.OwnerField.raw?.[0]?.id
          ?.toLowerCase()
          .replace(/[{}]/g, "") ?? "";
      const ownerType =
        _context.parameters.OwnerField.raw?.[0]?.entityType ?? "";

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
          const roleAssociations = userRoles.entities[0]
            .systemuserroles_association as RoleAssociation[] | undefined;
          roles = (roleAssociations ?? [])
            .map((r: RoleAssociation) => r.name)
            .filter(Boolean);
        }
        isAdmin = roles.includes("System Administrator");
        console.log(
          `Case 2 - Is current user System Administrator? ${isAdmin}`,
        );
      } catch (e) {
        console.error("Error fetching roles", e);
        alert("Error fetching roles: " + (e instanceof Error ? e.message : String(e)));
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
        } catch (e) {
          console.error("Error checking team membership", e);
          alert("Error checking team membership: " + (e instanceof Error ? e.message : String(e)));
        }
      }

      const authorized = isUserOwner || isAdmin || isTeamMember;
      setIsAuthorized(authorized);

      console.log(
        `User is ${authorized ? "authorized" : "not authorized"} for actions`,
      );
    } catch (error) {
      console.error("Permission check failed:", error);
      alert("Permission check failed: " + (error instanceof Error ? error.message : String(error)));
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
      /*if (succeeded) {
                return true; 
            } else {               
                await _context.navigation.openAlertDialog({text: response.errorMsg});
                return false;
            }*/
    } catch (error: unknown) {
      if (error instanceof Error) {
        await _context.navigation.openAlertDialog({ text: error.message });
        succeeded = false;
      } else {
        console.error("An unknown error occurred:", error);
        alert("An unknown error occurred in validateAction: " + String(error));
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
    try {
      if (
        _context.parameters.stepLookup.raw == null ||
        _context.parameters.stepLookup.raw.length == 0
      )
        return;
      setIsLoading(true);
      const Filter = (
        isMobile(_context)
          ? Constants.STAGE_ACTION_MOBILE_FILTER
          : Constants.STAGE_ACTION_DEFAULT_FILTER
      ).replace("{0}", _context.parameters.stepLookup.raw[0].id.toString());
      const query = Filter + Constants.STAGE_ACTION_QUERY;
      const response = await _context.webAPI.retrieveMultipleRecords(
        Constants.ENTITY_NAME,
        query,
      );

      console.log(Constants.MSG_PREFIX + response.entities.length);
      alert("Data retrieved (actions): " + response.entities.length + " records");

      const tempDataSet: IActionButtonProps[] = [];

      if (
        response.entities &&
        getValue(
          response.entities[0],
          _context.resources.getString(Constants.CURRENT_STEP_SEQUENCE),
        ) == "1" &&
        _context.parameters.GenericCancelAction.raw != null &&
        _context.parameters.GenericCancelAction.raw.length > 0
      ) {
        const genericfilter = Constants.GENERIC_CANCEL_FILTER.replace(
          "{0}",
          _context.parameters.GenericCancelAction.raw.toString(),
        );

        const genericcancelresponse =
          await _context.webAPI.retrieveMultipleRecords(
            Constants.ENTITY_NAME,
            genericfilter,
          );
        if (genericcancelresponse.entities)
          response.entities.push(genericcancelresponse.entities[0]);
      }

      response.entities.forEach((entity) => {
        const prop: IActionButtonProps = {
          displayName:
            getValue(
              entity,
              _context.resources.getString(Constants.DISPLAY_NAME),
            ) ?? "",
          buttonIcon: getValue(entity, Constants.BUTTON_ICON) ?? "",
          buttonColor: getValue(entity, Constants.BUTTON_COLOR) ?? "",
          buttonStatus: getValue(entity, Constants.BUTTON_STATUS) === "true",
          buttonId: getValue(entity, Constants.BUTTON_ID) ?? "",
          requireComments:
            getValue(entity, Constants.REQUIRE_COMMENTS) === "true",
          requireAssign:
            getValue(entity, Constants.IS_ASSIGN_ACTION_TYPE) === "true",
          requiresSurvey:
            getValue(entity, Constants.REQUIRES_SURVEY) === "true",
          sendToCustomer:
            getValue(entity, Constants.SEND_TO_CUSTOMER) === "true",
          nextTeamId: getValue(entity, Constants.NEXT_STEP_TEAM) ?? "",
          entity: entity,
          staticReplyTemplateId: getValue(
            entity,
            Constants.STATIC_REPLY_TEMPLATE_ID,
          ),
        };
        tempDataSet.push(prop);
      });
      setDataSet(tempDataSet);
    } catch (error: unknown) {
      if (error instanceof Error) {
        console.error("Error occurred in API call:", error.message);
        alert("Error occurred in API call: " + error.message);
      } else {
        console.error("An unknown error occurred");
        alert("An unknown error occurred in fetchData");
      }
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
      _context.parameters.lookupField.raw[0] = {
        entityType: Constants.ENTITY_NAME.toString(),
        id: actionId,
      };
      _context.parameters.commentField.raw = comment;
      _context.parameters.attachmentsGuid.raw = notetext;
      if (assignee != undefined) {
        _context.parameters.nextAssignee.raw[0] = {
          entityType: assignee.entityType,
          id: assignee.id,
        };
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
        _context.parameters.SurveyFormWRName.raw ??
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
        (error) => {
          console.error("Error opening survey form:", error);
          alert("Error opening survey form: " + (error instanceof Error ? error.message : String(error)));
          throw error;
        },
      );
    } catch (error) {
      console.error("Error in openSurveyModal:", error);
      alert("Error in openSurveyModal: " + (error instanceof Error ? error.message : String(error)));
    }
  };

  const SetShowSurvey = () => {
    console.log("SetShowSurvey called");
    _context.parameters.ShowSurveyForm.raw = true;

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
      //return;
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
        alert("Error retrieving records: " + (error instanceof Error ? error.message : String(error)));
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
      const serviceRequestId = _context.parameters.primaryKey?.formatted ?? "";
      console.log("Before call SetShowSurvey");
      SetShowSurvey();

      //await openSurveyModal(action.buttonId ?? '', surveyId, serviceRequestId);
      return;
    } else {
      if (isModalVisible == true) {
        setIsModalVisible(false);
      }

      if (isAssignModalVisible == true) {
        setIsAssignModalVisible(false);
      }

      const isExecuteAction =
        getValue(action.entity, Constants.WFACTION) == "false";
      //const code = getValue(action.entity, Constants.ACTION_CODE);
      //if (!iswfaction && code != "") {
      // Here, instead of eval, call the actual function directly
      //  await ShowReport(_context.parameters.reqNoField.raw!);
      //}
      //const varDef: string = "let srNumber = '" + _context.parameters.reqNoField.raw! + "'; ";
      if (isExecuteAction) {
        //await eval(varDef + code);
        // Example inside your PCF control
        const codeFromField = getValue(action.entity, Constants.ACTION_CODE); // e.g. "openScheduleBoard(formContext);"

        if (codeFromField) {
          try {
            //const recordId = (_context.mode as any).contextInfo.entityId;
            //const entityName = (_context.mode as any).contextInfo.entityTypeName;

            //openScheduleBoard(recordId, entityName, languageId);
            // Build a callable function that has access to context variables
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
              _context.parameters.regardingField.raw[0],
              (_context.mode as any).contextInfo,
              action,
            );
          } catch (error) {
            console.error("Error running dynamic code:", error);
            alert("Error running dynamic code: " + (error instanceof Error ? error.message : String(error)));
          }
        } else {
          console.log("No Code Specified");
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
              const primaryKey = _context.parameters.primaryKey?.raw;

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
          alert("Confirmation action failed: " + (error instanceof Error ? error.message : String(error)));
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

  return (
    <div className="buttonIcons">
      {isPermissionLoading ? (
        <div className="loading-spinner">
          <span>{_context.resources.getString(Constants.LOADING)}</span>
        </div>
      ) : !isAuthorized ? (
        <div className="no-actions-message">
          {_context.resources.getString(Constants.NOT_AUTHORIZED_MSG)}
        </div>
      ) : !isLoading && dataSet.length === 0 ? (
        <div className="no-actions-message">
          {_context.resources.getString(Constants.NO_RECORDS)}
        </div>
      ) : (
        dataSet.map((item) => (
          <PrimaryButton
            className="btnaction"
            key={item.buttonId}
            iconProps={{ iconName: item.buttonIcon }}
            text={item.displayName}
            onClick={() => void onClick(item, "")}
            disabled={!item.buttonStatus}
            style={{
              backgroundColor: item.buttonColor,
              border: item.buttonColor,
            }}
          />
        ))
      )}
      {isLoading && (
        <div className="loading-spinner">
          <span>{_context.resources.getString(Constants.LOADING)}</span>
        </div>
      )}

      <ModalDialog
        _context={_context}
        isVisible={isModalVisible}
        onClose={closeModal}
        onSave={saveModal}
        action={modalContent.action}
      />

      <AssignDialog
        _context={_context}
        isVisible={isAssignModalVisible}
        onClose={closeAssignModal}
        onSetNextAssigee={saveAssignModal}
        action={modalContent?.action ?? {}}
        teamId={modalContent?.action.nextTeamId ?? undefined}
      />
    </div>
  );
};
