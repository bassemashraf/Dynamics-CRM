import * as React from "react";
import { useState, useEffect } from "react";
import { IInputs } from "./generated/ManifestTypes";

// ---- Interfaces ----
export interface IActionButtonProps {
  buttonId?: string;
  displayName?: string;
  buttonColor?: string;
  buttonIcon?: string;
  requireComments?: boolean;
  requireAssign?: boolean;
  buttonStatus?: boolean;
  actionCommand?: string;
  entity?: any;
}

export interface IActionsProps {
  _context: ComponentFramework.Context<IInputs>;
  notifyOutputChanged: () => void;
  currentStepId: string;
  processExtId: string;
  xrmGlobal: any;
}

// ---- Component ----
export const Actions: React.FC<IActionsProps> = ({
  _context,
  notifyOutputChanged,
  currentStepId,
  processExtId,
  xrmGlobal
}) => {
  const [dataSet, setDataSet] = useState<IActionButtonProps[]>([]);
  const [isLoading, setIsLoading] = useState(true);

  useEffect(() => {
    fetchActions();
  }, [currentStepId]);

  const fetchActions = async () => {
    setIsLoading(true);
    try {
      const formFactor = _context.client.getFormFactor();
      const isMob = formFactor === 2 || formFactor === 3;

      const visibilityField = isMob ? "duc_showonmobile" : "duc_visible";

      const normalise = (g: string) => (g || "").replace(/[{}]/g, "").toLowerCase();
      const normalisedStepId = normalise(currentStepId);

      const actionQuery = `?$select=duc_stageactionid,duc_name,duc_arabicname,duc_enabled,duc_sequence,duc_requirescomments,duc_requiressurvey,_duc_actiontype_value,duc_actiontype,_duc_relatedstage_value,duc_relatedstage&$filter=${visibilityField} eq true&$orderby=duc_sequence asc`;

      const response = await xrmGlobal.WebApi.retrieveMultipleRecords(
        "duc_stageaction",
        actionQuery
      );

      // Local filter — braces stripped on both sides
      const filtered = normalisedStepId.length > 0
        ? response.entities.filter((e: any) => {
          const relatedStageId = e._duc_relatedstage_value || e.duc_relatedstage;
          return relatedStageId && normalise(relatedStageId) === normalisedStepId;
        })
        : response.entities; // If no step, show all

      // Collect unique action type IDs
      const actionTypeIds: string[] = [];
      filtered.forEach((e: any) => {
        const typeId = e.duc_actiontype || e._duc_actiontype_value;
        if (typeId && !actionTypeIds.includes(typeId)) {
          actionTypeIds.push(typeId);
        }
      });

      // Fetch ActionType metadata including duc_icon
      const actionTypeMap: Record<string, any> = {};
      if (actionTypeIds.length > 0) {
        const typeFilter = actionTypeIds.map((id) => `duc_actiontypeid eq ${id}`).join(" or ");
        const typeQuery = `?$filter=${typeFilter}&$select=duc_isassignaction,duc_wfaction,duc_actioncommand,duc_color,duc_icon,duc_sendtocustomer`;

        try {
          const typeResponse = await xrmGlobal.WebApi.retrieveMultipleRecords("duc_actiontype", typeQuery);
          typeResponse.entities.forEach((t: any) => {
            actionTypeMap[t.duc_actiontypeid] = t;
          });
        } catch (e: any) {
          console.warn("[Actions] Action Type Fetch Error:", e?.message || e);
        }
      }

      // Build button dataset
      const isLTR = _context.userSettings.languageId === 1033;
      const tempDataSet: IActionButtonProps[] = filtered.map((entity: any) => {
        const typeId = entity.duc_actiontype || entity._duc_actiontype_value;
        const actionType = typeId ? actionTypeMap[typeId] : null;
        return {
          displayName: isLTR ? (entity.duc_name || "") : (entity.duc_arabicname || entity.duc_name || ""),
          buttonIcon: actionType ? String(actionType.duc_icon || "") : "",
          buttonColor: actionType ? String(actionType.duc_color || "#0078d4") : "#0078d4",
          buttonStatus: entity.duc_enabled !== false,
          buttonId: entity.duc_stageactionid || "",
          requireComments: entity.duc_requirescomments === true,
          requireAssign: actionType ? actionType.duc_isassignaction === true : false,
          requiresSurvey: entity.duc_requiressurvey === true,
          actionCommand: actionType ? String(actionType.duc_actioncommand || "") : "",
          entity: { ...entity, duc_actiontype: actionType },
        };
      });

      setDataSet(tempDataSet);
    } catch (e: any) {
      console.error("[Actions] fetchActions error:", e?.message || e);
    } finally {
      setIsLoading(false);
    }
  };

  const handleClick = async (action: IActionButtonProps) => {
    if (!action.buttonStatus) return;

    if (action.actionCommand) {
      try {
        const updatedCode = action.actionCommand.replace("#StageActionId#", action.buttonId ?? "");
        const AsyncFunction = Object.getPrototypeOf(async function () {
          // empty
        }).constructor as any;
        const runCode = new AsyncFunction(
          "languageId",
          "processExtension",
          "action",
          '"use strict"; ' + updatedCode
        );
        await runCode(
          _context.userSettings.languageId,
          { entityId: processExtId, entityName: "duc_processextension" },
          action
        );
      } catch (err: any) {
        console.error("[Actions] Error executing action:", err?.message || err);
        await _context.navigation.openAlertDialog({ text: `Error executing action:\n${err?.message || err}` });
      }
    } else {
      const confirmed = await _context.navigation.openConfirmDialog({
        title: "Confirm Action",
        text: `Perform "${action.displayName}"?`,
      });
      if (confirmed.confirmed) {
        notifyOutputChanged();
      }
    }
  };

  if (isLoading || dataSet.length === 0) return null;

  return (
    <div style={{ display: "flex", flexWrap: "wrap", gap: "8px", padding: "10px 0" }}>
      {dataSet.map((item) => (
        <button
          key={item.buttonId}
          onClick={() => void handleClick(item)}
          disabled={!item.buttonStatus}
          style={{
            backgroundColor: item.buttonColor || "#0078d4",
            color: "#fff",
            border: "none",
            padding: "8px 16px",
            borderRadius: "4px",
            cursor: item.buttonStatus ? "pointer" : "not-allowed",
            opacity: item.buttonStatus ? 1 : 0.5,
            fontSize: "14px",
            fontWeight: "600",
            display: "flex",
            alignItems: "center",
            gap: "6px",
          }}
        >
          {item.buttonIcon && (
            <i
              className={item.buttonIcon}
              style={{ fontSize: "16px" }}
            />
          )}
          {item.displayName}
        </button>
      ))}
    </div>
  );
};
