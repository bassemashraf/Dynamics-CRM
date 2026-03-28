import * as React from "react";
import { useState, useEffect } from "react";
import { Actions } from "./Actions";
import { IInputs, IOutputs } from "../generated/ManifestTypes";
import {
  DefaultButton,
  PrimaryButton,
  Spinner,
  SpinnerSize,
} from "@fluentui/react";
import { Constants } from "../constants"; // Importing constants
import AssignDialog from "./AssignDialog";
import { FlowProgressBar } from "./FlowProgressBar";
import { FormContextHelper } from "./FormContextHelper";

export interface IMainProps {
  _context: ComponentFramework.Context<IInputs>;
  onnotifyOutputChanged: () => void;
}

interface TeamMembership {
  teamid: string;
}

interface RoleAssociation {
  name: string;
}

export const Main: React.FC<IMainProps> = ({
  _context,
  onnotifyOutputChanged,
}) => {
  const [isLoading, setIsLoading] = useState(false); // State to track loading status
  const [isInitializing, setIsInitializing] = useState(true); // Pre-fetch state
  const [isAssignModalVisible, setIsAssignModalVisible] = useState(false);

  const [canShowAssignmentButton, setCanShowAssignmentButton] = useState(false);

  const ShowBusinessFlow = _context.parameters.ShowBusinessFlow?.raw ?? true;
  const ShowMiddleSection = _context.parameters.ShowMiddleSection?.raw ?? true;
  const ShowAvailableActions =
    _context.parameters.ShowAvailableActions?.raw ?? true;
  let isMobileOrTablet = false;

  useEffect(() => {
    const initData = async () => {
      const xrmGlobal = (window.parent as any).Xrm || (window as any).Xrm;
      if (xrmGlobal) {
        await FormContextHelper.initAsync(xrmGlobal);
      }
      await evaluatePermissions();
      setIsInitializing(false);
    };

    void initData();

    const formFactor = _context.client.getFormFactor();
    const isPhone = formFactor === Constants.isPhone;
    const isTablet = formFactor === Constants.isTablet;
    isMobileOrTablet = isPhone || isTablet;
  }, []);

  const closeAssignModal = () => {
    setIsAssignModalVisible(false);
  };

  const saveAssignModal = (assignee: {
    id: string;
    name: string;
    type: "user" | "team";
    entityType: string;
  }) => {
    console.log("Saving data");

    // Close the modal if open
    if (isAssignModalVisible) {
      setIsAssignModalVisible(false);
    }

    // Only proceed if assignee is defined
    if (assignee) {
      FormContextHelper.setLookupValue('ownerid', assignee.id, assignee.name, assignee.entityType);
      onnotifyOutputChanged();
    }
  };

  //Check Permission
  const evaluatePermissions = async () => {
    try {
      // Offline: skip permission checks — system entities (teammembership, role, etc.) are not available
      const isOffline = _context.client.isOffline?.() === true;
      if (isOffline) {
        console.log("[evaluatePermissions] Offline mode — auto-granting permissions");
        setCanShowAssignmentButton(!!FormContextHelper.getLookupValue('duc_owningteam')?.id);
        return;
      }

      const currentUserId = _context.userSettings.userId
        .replace(/[{}]/g, "")
        .toLowerCase();
      
      const ownerVal = FormContextHelper.getLookupValue('ownerid');
      const ownerId = ownerVal?.id?.toLowerCase().replace(/[{}]/g, "") ?? "";
      const ownerType = ownerVal?.entityType ?? "";

      console.log("Current User ID:", currentUserId);
      console.log("Owner ID:", ownerId, "Owner Type:", ownerType);

      let currentUserTeams: string[] = [];
      try {
        const teamRes = await _context.webAPI.retrieveMultipleRecords(
          "teammembership",
          `?$filter=systemuserid eq ${currentUserId}`,
        );
        currentUserTeams = (teamRes.entities as TeamMembership[])
          .map((t) => t.teamid?.toLowerCase().replace(/[{}]/g, ""))
          .filter(Boolean);
      } catch (e: any) {
        console.error("Error fetching teams", e);
      }
      console.log("Current User Teams:", currentUserTeams);

      let roles: string[] = [];
      try {
        // Step 1: Query the many-to-many relationship entity to get role IDs
        const userRolesRes = await _context.webAPI.retrieveMultipleRecords(
          "systemuserrolescollection",
          `?$filter=systemuserid eq '${currentUserId}'`,
        );

        const roleIds = userRolesRes.entities
          .map((e: any) => e.roleid)
          .filter(Boolean);

        if (roleIds.length > 0) {
          // Step 2: Query the roles entity to get role names
          let roleFilter = "";
          roleIds.forEach((id: string, index: number) => {
            if (index > 0) roleFilter += " or ";
            roleFilter += `roleid eq ${id}`;
          });

          const rolesRes = await _context.webAPI.retrieveMultipleRecords(
            "role",
            `?$filter=${roleFilter}&$select=name`,
          );

          roles = rolesRes.entities.map((r: any) => r.name).filter(Boolean);
        }
      } catch (e: any) {
        console.error("Error fetching roles", e);
      }
      console.log("User Roles:", roles);

      // Case 1: Current user is the record owner
      const isUserOwner = ownerType === "systemuser" && currentUserId === ownerId;
      console.log(
        `Case 1 - Is current user the record owner? ${isUserOwner} (User ID: ${currentUserId} vs Owner ID: ${ownerId})`,
      );

      // Case 2: Current user is member of owning team
      const isTeamMember =
        ownerType === "team" && currentUserTeams.includes(ownerId);
      console.log(
        `Case 2 - Is current user member of owning team? ${isTeamMember} (Team ID: ${ownerId} in User Teams: [${currentUserTeams.join(", ")}])`,
      );

      // Case 3: Current user is System Administrator
      const isAdmin = roles.includes("System Administrator");
      console.log(
        `Case 3 - Is current user System Administrator? ${isAdmin} (Roles: [${roles.join(", ")}])`,
      );

      // Case 4: Check if current user is administrator of owning team
      let isTeamAdmin = false;
      if (ownerType === "team") {
        try {
          // Fetch team record with administrator information
          const team = await _context.webAPI.retrieveRecord(
            "team",
            ownerId,
            "?$select=administratorid",
          );

          const teamAdminId =
            team.administratorid?.replace(/[{}]/g, "").toLowerCase() ?? "";
          isTeamAdmin = teamAdminId === currentUserId;

          console.log(
            `Case 4 - Team Admin ID: ${teamAdminId}, Current User ID: ${currentUserId}`,
          );
          console.log(
            `Case 4 - Is current user team administrator? ${isTeamAdmin}`,
          );
        } catch (e: any) {
          console.error("Error fetching team administrator", e);
        }
      } else {
        console.log("Case 4 - Not applicable (owner is not a team)");
      }

      if (
        (isUserOwner || isTeamMember || isAdmin || isTeamAdmin) &&
        FormContextHelper.getLookupValue('duc_owningteam')?.id
      ) {
        console.log("Permission granted: Showing assignment button");
        setCanShowAssignmentButton(true);
      } else {
        console.log("Permission denied: Hiding assignment button");
        setCanShowAssignmentButton(false);
      }
    } catch (e: any) {
      console.error("[evaluatePermissions] Unexpected error — defaulting to hide assignment", e);
      setCanShowAssignmentButton(false);
    }
  };

  const onviewbtnClick = async () => {
    setIsLoading(true); // Show loading when async starts
    try {
      const regardingVal = FormContextHelper.getLookupValue('regardingobjectid');
      if (!regardingVal) return;

      const pageInput: Xrm.Navigation.PageInputEntityRecord = {
        pageType: "entityrecord",
        entityName: regardingVal.entityType,
        entityId: regardingVal.id,
      };

      const navigationOptions: Xrm.Navigation.NavigationOptions = {
        target: 2,
        width: { value: 70, unit: "%" },
        height: { value: 70, unit: "%" },
        position: 1, // center
      };

      await Xrm.Navigation.navigateTo(pageInput, navigationOptions);
    } catch (error: any) {
      console.error("Error in navigation:", error);
    } finally {
      setIsLoading(false); // Hide loading when async is done
    }
  };

  const openSurveyResponseModal = async (_serviceRequestId: string) => {
    //surveyResponseId
    setIsLoading(true); // Show loading when async starts
    try {
      const Data = {
        //ResponseId: surveyResponseId
        ServiceRequestId: _serviceRequestId,
      };

      const pageInput: Xrm.Navigation.PageInputHtmlWebResource = {
        pageType: "webresource",
        webresourceName: _context.parameters.SurveyResponseViewWRName?.raw ?? "",
        data: JSON.stringify(Data),
      };

      const navigationOptions: Xrm.Navigation.NavigationOptions = {
        target: 2, // Modal dialog
        width: { value: 700, unit: "px" },
        height: { value: 1000, unit: "px" },
        position: 1, // Center
      };

      await Xrm.Navigation.navigateTo(pageInput, navigationOptions);
    } catch (error: any) {
      console.error("Error opening survey response view:", error);
    } finally {
      setIsLoading(false); // Hide loading when async is done
    }
  };
  const lastSurveyResponseExists = !!FormContextHelper.getLookupValue('duc_surveyresponse')?.id;
  const onCheckListClick = async (): Promise<void> => {
    setIsLoading(true); // Show loading when async starts
    try {
      const serviceRequestId = FormContextHelper.getStringValue("activityid") || FormContextHelper.getFormContext()?.data?.entity?.getId()?.replace(/[{}]/g, "") || "";
      await openSurveyResponseModal(serviceRequestId);
      console.log("No survey response found for this service request");
    } catch (error: any) {
      console.error("Error in onCheckListClick:", error);
    } finally {
      setIsLoading(false); // Hide loading when async is done
    }
  };

  const onAttachmentsClick = async (): Promise<void> => {
    setIsLoading(true); // Show loading when async starts
    try {
      const serviceRequestId = FormContextHelper.getStringValue("activityid") || FormContextHelper.getFormContext()?.data?.entity?.getId()?.replace(/[{}]/g, "") || "";
      const pageInput: Xrm.Navigation.PageInputEntityRecord = {
        pageType: "entityrecord",
        entityName: Constants.MAIN_ENTITY_NAME,
        entityId: serviceRequestId,
        formId: "83F06A4B-5C6E-4A2D-B9D1-7C10E6E73709",
      };

      const navigationOptions: Xrm.Navigation.NavigationOptions = {
        target: 2,
        width: { value: 70, unit: "%" },
        height: { value: 70, unit: "%" },
        position: 1, // center,
      };

      await Xrm.Navigation.navigateTo(pageInput, navigationOptions);
    } catch (error: any) {
      console.error("Error in onAttachmentsClick:", error);
    } finally {
      setIsLoading(false); // Hide loading when async is done
    }
  };

  if (isInitializing) {
    return React.createElement(
      "div",
      { className: "loading-spinner", style: { padding: 20 } },
      React.createElement(Spinner, {
        size: SpinnerSize.large,
        label: _context.resources.getString(Constants.LOADING) || "Initializing..."
      })
    );
  }

  return React.createElement(
    "div",
    { className: "top fullWidth" },
    // ShowBusinessFlow section
    ShowBusinessFlow && React.createElement(
      "div",
      { className: isMobileOrTablet ? "" : "borders" },
      React.createElement(FlowProgressBar, {
        _context: _context,
        currentStepId: FormContextHelper.getLookupValue('duc_currentstage')?.id || "",
        serviceRequestTypeId: FormContextHelper.getLookupValue('duc_processdefinition')?.id || "",
      })
    ),
    // ShowMiddleSection
    ShowMiddleSection && React.createElement(
      "div",
      { className: isMobileOrTablet ? "" : "borders" },
      React.createElement(
        "div",
        { className: "col2", style: { width: "25%" } },
        React.createElement(
          "label",
          null,
          _context.resources.getString(Constants.OWNING_TEAM_LABEL)
        ),
        " ",
        React.createElement("input", {
          type: "text",
          className: "txt",
          disabled: true,
          value: FormContextHelper.getLookupValue('duc_owningteam')?.name ?? Constants.NS,
          style: { width: "80%" },
        })
      ),
      React.createElement(
        "div",
        { className: "col", style: { width: "75%" } },
        React.createElement(
          "div",
          { style: { display: "flex", gap: "8px" } },
          // Attachments button
          React.createElement(PrimaryButton, {
            className: "viewDetailsbtn",
            key: "attachmentsbtn",
            iconProps: { iconName: "FileImage" },
            text: _context.resources.getString(Constants.ATTACHMENTS_BTN_LABEL),
            onClick: () => void onAttachmentsClick(),
          }),
          // Assignment button (conditional)
          canShowAssignmentButton && React.createElement(
            React.Fragment,
            null,
            React.createElement(PrimaryButton, {
              className: "viewDetailsbtn",
              key: "assignmentbtn",
              iconProps: { iconName: "FollowUser" },
              text: _context.resources.getString(Constants.ASSIGNMENT_BTN_LABEL),
              onClick: () => setIsAssignModalVisible(true),
            }),
            React.createElement(AssignDialog, {
              _context: _context,
              isVisible: isAssignModalVisible,
              onClose: closeAssignModal,
              onAssign: saveAssignModal,
              teamId: FormContextHelper.getLookupValue('duc_owningteam')?.id,
            })
          )
        )
      )
    ),
    // Show loading spinner when isLoading is true
    isLoading && React.createElement(
      "div",
      { className: "loading-spinner" },
      React.createElement(Spinner, {
        size: SpinnerSize.large,
        label: _context.resources.getString(Constants.LOADING),
      }),
      " "
    ),
    // ShowAvailableActions section
    ShowAvailableActions && React.createElement(
      "div",
      { className: "buttonIcons" },
      React.createElement(Actions, {
        _context: _context,
        onnotifyOutputChanged: onnotifyOutputChanged,
      })
    )
  );
};
