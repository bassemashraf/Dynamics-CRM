import * as React from "react";
import { useState, useEffect, useRef } from "react";
import { IInputs } from "./generated/ManifestTypes";
import { FlowProgressBar, IStep } from "./FlowProgressBar";
import { Actions } from "./Actions";

export interface IMainProps {
  _context: ComponentFramework.Context<IInputs>;
  notifyOutputChanged: () => void;
}

export const MainControl: React.FC<IMainProps> = ({
  _context,
  notifyOutputChanged
}) => {
  const [steps, setSteps] = useState<IStep[]>([]);
  const [currentStepId, setCurrentStepId] = useState<string>("");
  const [processExtId, setProcessExtId] = useState<string>("");
  const [isReady, setIsReady] = useState(false);

  // Store xrm in a ref — NOT state — to avoid React calling it as an updater function
  const xrmRef = useRef<any>(null);

  const isLTR = _context.userSettings.languageId === 1033;
  const formFactor = _context.client.getFormFactor();
  const isMobileOrTablet = formFactor === 2 || formFactor === 3;

  useEffect(() => {
    fetchData();
  }, []);

  const getFormContext = (): any => {
    let parentXrm: any = null;
    try {
      parentXrm = (window.parent as any).Xrm;
    } catch (e) {
      console.warn("Offline isolation restricted parent XRM");
    }
    return (parentXrm || (window as any).Xrm) ?? null;
  };

  const getLookupId = (fieldName: string): string => {
    const xrm = getFormContext();
    const ctx = xrm ? xrm.Page : null;
    if (ctx && ctx.getAttribute) {
      const attr = ctx.getAttribute(fieldName);
      if (attr) {
        const val = attr.getValue();
        if (val && val.length > 0) return val[0].id.replace(/[{}]/g, "");
      }
    }
    return "";
  };

  const fetchData = async () => {
    try {
      const xrm = getFormContext();
      if (!xrm || !xrm.WebApi) {
        console.error("[MainControl] Xrm.WebApi is not available in this context!");
        return;
      }
      xrmRef.current = xrm;

      // Step 1: Read duc_processextension lookup from the Work Order form
      const extId = getLookupId('duc_processextension');
      if (!extId) {
        console.warn("[MainControl] duc_processextension lookup not found on the current form.");
        return;
      }
      setProcessExtId(extId);

      // Step 2: Fetch the Process Extension record
      const processExtRecord = await xrm.WebApi.retrieveRecord(
        "duc_processextension",
        extId,
        `?$select=_duc_processdefinition_value,_duc_currentstage_value`
      );

      const processDefId = (processExtRecord._duc_processdefinition_value || "").replace(/[{}]/g, "").toLowerCase();
      const stepId = (processExtRecord._duc_currentstage_value || "").replace(/[{}]/g, "").toLowerCase();
      setCurrentStepId(stepId);

      if (!processDefId) {
        console.warn("[MainControl] No Process Definition found on Process Extension record.");
        return;
      }

      // Step 3: Fetch Stages
      const visibleField = isMobileOrTablet ? "duc_visibleonmobile" : "duc_visible";
      const selectFields = "duc_visible,duc_visibleonmobile,duc_sequence,duc_processstageid,duc_name,duc_arabicname,duc_arabicdescription,duc_description,duc_sequenceoverride";

      const result = await xrm.WebApi.retrieveMultipleRecords(
        "duc_processstage",
        `?$select=${selectFields}` +
        `&$filter=duc_relatedprocess eq ${processDefId} and statecode eq 0 and (${visibleField} eq true)` +
        `&$orderby=duc_sequence,duc_sequenceoverride asc`
      );

      const fetchedSteps: IStep[] = result.entities.map((e: any) => ({
        id: (e.duc_processstageid || "").toLowerCase(),
        name: e.duc_name,
        displayNameEN: e.duc_name,
        displayNameAR: e.duc_arabicname,
        descriptionEN: e.duc_description,
        descriptionAR: e.duc_arabicdescription,
        sequence: e.duc_sequenceoverride ?? e.duc_sequence,
        visible: e.duc_visible ?? e.duc_visibleonmobile,
        done: false,
        active: false,
      })).sort((a: IStep, b: IStep) => a.sequence - b.sequence);

      // Deduplicate by name — keep only the first stage (lowest sequence) per unique name.
      // This ensures stages with the same name appear as a single step in the progress bar.
      const seenNames = new Set<string>();
      const uniqueSteps = fetchedSteps.filter((s) => {
        const key = (s.displayNameEN || s.name || "").trim().toLowerCase();
        if (seenNames.has(key)) return false;
        seenNames.add(key);
        return true;
      });

      const currentIndex = uniqueSteps.findIndex((s) => s.id === stepId);
      // If the active stage was deduplicated away, find it by name instead
      let resolvedCurrentIndex = currentIndex;
      if (currentIndex === -1 && stepId) {
        const activeStep = fetchedSteps.find((s) => s.id === stepId);
        if (activeStep) {
          const activeName = (activeStep.displayNameEN || activeStep.name || "").trim().toLowerCase();
          resolvedCurrentIndex = uniqueSteps.findIndex(
            (s) => (s.displayNameEN || s.name || "").trim().toLowerCase() === activeName
          );
        }
        if (resolvedCurrentIndex === -1) {
          console.warn("[MainControl] Current stage ID not found among deduplicated steps:", stepId);
        }
      }

      const lastIndex = uniqueSteps.length - 1;
      const processed = uniqueSteps.map((step, index) => ({
        ...step,
        done: resolvedCurrentIndex !== -1 && (index < resolvedCurrentIndex || (index === lastIndex && index === resolvedCurrentIndex)),
        active: resolvedCurrentIndex !== -1 && index === resolvedCurrentIndex,
      }));

      setSteps(processed);
      // IMPORTANT: set isReady LAST to trigger Actions render
      setIsReady(true);
    } catch (e: any) {
      console.error("[MainControl] fetchData error:", e?.message || e);
    }
  };

  return (
    <div style={{ fontFamily: "sans-serif" }}>
      {steps.length > 0 && (
        <FlowProgressBar
          steps={steps}
          isLTR={isLTR}
          isMobileOrTablet={isMobileOrTablet}
        />
      )}

      {/* Always render Actions once data is loaded — do NOT gate on currentStepId */}
      {isReady && xrmRef.current && (
        <Actions
          _context={_context}
          notifyOutputChanged={notifyOutputChanged}
          currentStepId={currentStepId}
          processExtId={processExtId}
          xrmGlobal={xrmRef.current}
        />
      )}
    </div>
  );
};
