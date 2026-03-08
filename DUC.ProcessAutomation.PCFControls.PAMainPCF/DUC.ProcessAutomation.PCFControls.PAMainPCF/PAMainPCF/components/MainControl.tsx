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
  const [isAssignModalVisible, setIsAssignModalVisible] = useState(false);

  const [canShowAssignmentButton, setCanShowAssignmentButton] = useState(false);

  const ShowBusinessFlow = _context.parameters.ShowBusinessFlow.raw ?? true;
  const ShowMiddleSection = _context.parameters.ShowMiddleSection.raw ?? true;
  const ShowAvailableActions =
    _context.parameters.ShowAvailableActions.raw ?? true;
  let isMobileOrTablet = false;

  useEffect(() => {
    void evaluatePermissions();
    const formFactor = _context.client.getFormFactor();
    const isPhone = formFactor === Constants.isPhone;
    const isTablet = formFactor === Constants.isTablet;
    isMobileOrTablet = isPhone || isTablet;
  }, [_context.parameters.stepLookup, _context.parameters.OwnerField]);

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
      _context.parameters.OwnerField.raw[0] = {
        entityType: assignee.entityType,
        id: assignee.id,
        name: assignee.name,
      };
      onnotifyOutputChanged();
    }
  };

  //Check Permission
  const evaluatePermissions = async () => {
    const currentUserId = _context.userSettings.userId
      .replace(/[{}]/g, "")
      .toLowerCase();
    const ownerId =
      _context.parameters.OwnerField.raw?.[0]?.id
        ?.toLowerCase()
        .replace(/[{}]/g, "") ?? "";
    const ownerType = _context.parameters.OwnerField.raw?.[0]?.entityType ?? "";

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
    } catch (e) {
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
    } catch (e) {
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
      } catch (e) {
        console.error("Error fetching team administrator", e);
      }
    } else {
      console.log("Case 4 - Not applicable (owner is not a team)");
    }

    if (
      (isUserOwner || isTeamMember || isAdmin || isTeamAdmin) &&
      _context.parameters.OwningTeam?.raw?.[0]?.id
    ) {
      console.log("Permission granted: Showing assignment button");
      setCanShowAssignmentButton(true);
    } else {
      console.log("Permission denied: Hiding assignment button");
      setCanShowAssignmentButton(false);
    }
  };

  const onviewbtnClick = async () => {
    setIsLoading(true); // Show loading when async starts
    try {
      const pageInput: Xrm.Navigation.PageInputEntityRecord = {
        pageType: "entityrecord",
        entityName: _context.parameters.regardingField.raw[0].entityType,
        entityId: _context.parameters.regardingField.raw[0].id,
      };

      const navigationOptions: Xrm.Navigation.NavigationOptions = {
        target: 2,
        width: { value: 70, unit: "%" },
        height: { value: 70, unit: "%" },
        position: 1, // center
      };

      await Xrm.Navigation.navigateTo(pageInput, navigationOptions);
    } catch (error) {
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
        webresourceName: _context.parameters.SurveyResponseViewWRName.raw ?? "",
        data: JSON.stringify(Data),
      };

      const navigationOptions: Xrm.Navigation.NavigationOptions = {
        target: 2, // Modal dialog
        width: { value: 700, unit: "px" },
        height: { value: 1000, unit: "px" },
        position: 1, // Center
      };

      await Xrm.Navigation.navigateTo(pageInput, navigationOptions);
    } catch (error) {
      console.error("Error opening survey response view:", error);
    } finally {
      setIsLoading(false); // Hide loading when async is done
    }
  };
  const lastSurveyResponseExists =
    !!_context.parameters.LastSurveyResponse?.raw;
  const onCheckListClick = async (): Promise<void> => {
    setIsLoading(true); // Show loading when async starts
    try {
      const serviceRequestId = _context.parameters.primaryKey?.formatted ?? "";
      await openSurveyResponseModal(serviceRequestId);
      console.log("No survey response found for this service request");
    } catch (error) {
      console.error("Error in onCheckListClick:", error);
    } finally {
      setIsLoading(false); // Hide loading when async is done
    }
  };

  const onAttachmentsClick = async (): Promise<void> => {
    setIsLoading(true); // Show loading when async starts
    try {
      const serviceRequestId = _context.parameters.primaryKey?.formatted ?? "";
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
    } catch (error) {
      console.error("Error in onAttachmentsClick:", error);
    } finally {
      setIsLoading(false); // Hide loading when async is done
    }
  };

  return (
    <div className="top fullWidth">
      {ShowBusinessFlow && (
        <div className={isMobileOrTablet ? "" : "borders"}>
          <FlowProgressBar
            _context={_context}
            currentStepId={_context.parameters.stepLookup.raw[0].id.toString()}
            serviceRequestTypeId={_context.parameters.processDefinition.raw[0].id.toString()}
          ></FlowProgressBar>
        </div>
      )}
      {ShowMiddleSection && (
        <div className={isMobileOrTablet ? "" : "borders"}>
          <div className="col2" style={{ width: "25%" }}>
            <label>
              {_context.resources.getString(Constants.OWNING_TEAM_LABEL)}
            </label>{" "}
            {/* Dynamically fetching the label for "OWNING TEAM" */}
            <input
              type="text"
              className="txt"
              disabled={true}
              value={
                _context.parameters.OwningTeam.raw?.[0]?.name?.toString() ??
                Constants.NS
              }
              style={{ width: "80%" }}
            />
          </div>
          <div className="col" style={{ width: "75%" }}>
            <div style={{ display: "flex", gap: "8px" }}>
              {/*
            <PrimaryButton
              className='viewDetailsbtn'
              key='viewbtn'
              iconProps={{ iconName: 'View' }}
              text={_context.resources.getString(Constants.DETAILS_BTN_TEXT)}
              onClick={() => void onviewbtnClick()} // Fixed handler
            />
            {lastSurveyResponseExists && (
              <PrimaryButton
                className='viewDetailsbtn'
                key='checklistbtn'
                iconProps={{ iconName: 'CheckList' }}
                text={_context.resources.getString(Constants.CHECKLIST_BTN_LABEL)}
                onClick={() => void onCheckListClick()}
              />
            )}*/}
              <PrimaryButton
                className="viewDetailsbtn"
                key="attachmentsbtn"
                iconProps={{ iconName: "FileImage" }}
                text={_context.resources.getString(
                  Constants.ATTACHMENTS_BTN_LABEL,
                )}
                onClick={() => void onAttachmentsClick()} // Fixed handler
              />
              {/* <PrimaryButton 
                className='viewDetailsbtn' 
                key='assignmentbtn' 
                iconProps={{ iconName: 'FollowUser' }} 
                text={_context.resources.getString(Constants.ASSIGNMENT_BTN_LABEL)} 
                onClick={() => void onAssignmentClick()}
              />*/}
              {canShowAssignmentButton && (
                <>
                  <PrimaryButton
                    className="viewDetailsbtn"
                    key="assignmentbtn"
                    iconProps={{ iconName: "FollowUser" }}
                    text={_context.resources.getString(
                      Constants.ASSIGNMENT_BTN_LABEL,
                    )}
                    onClick={() => setIsAssignModalVisible(true)}
                  />
                  <AssignDialog
                    _context={_context}
                    isVisible={isAssignModalVisible}
                    onClose={closeAssignModal}
                    onAssign={saveAssignModal}
                    teamId={_context.parameters.OwningTeam.raw?.[0]?.id}
                  />
                </>
              )}
            </div>
          </div>
        </div>
      )}

      {/* Show loading spinner when isLoading is true */}
      {isLoading && (
        <div className="loading-spinner">
          <Spinner
            size={SpinnerSize.large}
            label={_context.resources.getString(Constants.LOADING)}
          />{" "}
          {/* Dynamically displaying loading message */}
        </div>
      )}

      {ShowAvailableActions && (
        <div className="buttonIcons">
          <Actions
            _context={_context}
            onnotifyOutputChanged={onnotifyOutputChanged}
          />
        </div>
      )}
    </div>
  );
};
