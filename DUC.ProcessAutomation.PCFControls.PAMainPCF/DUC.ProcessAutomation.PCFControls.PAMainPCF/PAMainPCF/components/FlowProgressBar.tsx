import * as React from "react";
import { useState, useEffect, useRef } from "react";
import { IInputs } from "../generated/ManifestTypes";
import { Constants } from "../constants";
import "../css/FlowProgressBar.css";

interface IStep {
  id: string;
  name: string;
  displayNameEN: string;
  displayNameAR: string;
  descriptionEN?: string;
  descriptionAR?: string;
  sequence: number;
  visible: boolean;
  done: boolean;
  active: boolean;
}

interface IFlowProgressBarProps {
  _context: ComponentFramework.Context<IInputs>;
  currentStepId: string;
  serviceRequestTypeId: string;
}

export const FlowProgressBar: React.FC<IFlowProgressBarProps> = ({
  _context,
  currentStepId,
  serviceRequestTypeId,
}) => {
  const [steps, setSteps] = useState<IStep[]>([]);
  const [isLoading, setIsLoading] = useState<boolean>(true);
  const containerRef = useRef<HTMLDivElement>(null);
  let isMobileOrTablet = false;
  const isLTR = _context.userSettings.languageId === 1033;

  const doneColor = "#3BB674";
  const doneTextColor = "#fff";
  const activeColor = "#FFA800";
  const activeTextColor = "#fff";
  const activeShadow = "rgba(255, 193, 7, 0.8)";
  const pendingColor = "#95A0B6";
  const pendingTextColor = "#fff";
  const textColor = "#333";
  const connectorColor = "#3BB674";

  useEffect(() => {
    if (!serviceRequestTypeId) return;
    void fetchSteps();
  }, [serviceRequestTypeId, currentStepId]);

  const fetchSteps = async () => {
    try {
      setIsLoading(true);

      let currentStepFilter = "";

      const formFactor = _context.client.getFormFactor();
      const isPhone = formFactor === Constants.isPhone;
      const isTablet = formFactor === Constants.isTablet;
      isMobileOrTablet = isPhone || isTablet;

      let currentStepVisibleField = "";
      if (isMobileOrTablet) {
        currentStepVisibleField = Constants.STAGE_MOBILE_VISIBLE_FIELD;
        if (currentStepId != null) {
          currentStepFilter = `or (duc_processstageid eq '${currentStepId}' and duc_sequenceoverride ne null)`;
        }
      } else {
        currentStepVisibleField = Constants.STAGE_VISIBLE_FIELD;
        if (currentStepId != null) {
          currentStepFilter = `or (duc_processstageid eq '${currentStepId}' and duc_sequenceoverride ne null)`;
        }
      }

      const result = await _context.webAPI.retrieveMultipleRecords(
        Constants.STAGE_ENTITY_NAME,
        Constants.STAGE_SELECT +
          Constants.STAGE_FILTER.replace("{0}", serviceRequestTypeId)
            .replace("{1}", currentStepFilter)
            .replace("{2}", currentStepVisibleField) +
          Constants.STAGE_ORDER,
      );

      const fetchedSteps: IStep[] = result.entities
        .map((e: any) => ({
          id: e.duc_processstageid,
          name: e.duc_name,
          displayNameEN: e.duc_name ?? e.duc_name,
          displayNameAR: e.duc_arabicname,
          descriptionEN: e.duc_descriptionen,
          descriptionAR: e.duc_description,
          sequence: e.duc_sequenceoverride ?? e.duc_sequence,
          visible: e.duc_visible ?? e.duc_visibleonmobile,
          done: false,
          active: false,
        }))
        .sort((a, b) => a.sequence - b.sequence);

      let currentIndex = fetchedSteps.findIndex((s) => s.id === currentStepId);
      if (currentIndex === -1) currentIndex = 0;

      const lastIndex = fetchedSteps.length - 1;
      const processed = fetchedSteps.map((step, index) => ({
        ...step,
        done:
          index < currentIndex ||
          (index === lastIndex && index === currentIndex),
        active: index === currentIndex,
      }));

      setSteps(processed);
    } catch (error) {
      console.error("Error loading steps:", error);
    } finally {
      setIsLoading(false);
    }
  };

  const getConnectorWidth = () => {
    if (!containerRef.current || steps.length <= 1) return 0;
    const containerWidth = containerRef.current.offsetWidth;
    const totalStepWidth = steps.length * 32;
    const connectorsCount = steps.length - 1;
    return connectorsCount > 0
      ? (containerWidth - totalStepWidth) / connectorsCount
      : 0;
  };

  return (
    <div
      className={`main-pcf-flow-container ${isLTR ? "ltr" : "rtl"} ${
        isMobileOrTablet
          ? "main-pcf-flow-container-mobile-padding"
          : "main-pcf-flow-container-web-padding"
      }`}
      ref={containerRef}
    >
      {isLoading ? (
        <div>Loading steps...</div>
      ) : steps.length > 0 ? (
        steps.map((step, index) => {
          const isLast = index === steps.length - 1;
          const status = step.done ? "done" : step.active ? "active" : "";

          return (
            <div key={step.id} className={`step ${status}`}>
              {/* Step number / check */}
              <div
                className="step-number"
                title={
                  isLTR
                    ? (step.descriptionEN ?? step.displayNameEN)
                    : (step.descriptionAR ?? step.displayNameAR)
                }
              >
                {step.done ? (
                  <svg
                    style={{ width: "20px" }}
                    fill="#000000"
                    version="1.1"
                    xmlns="http://www.w3.org/2000/svg"
                    viewBox="0 0 335.765 335.765"
                  >
                    <g>
                      <g>
                        <polygon
                          fill="white"
                          points="311.757,41.803 107.573,245.96 23.986,162.364 0,186.393 107.573,293.962 335.765,65.795"
                        />
                      </g>
                    </g>
                  </svg>
                ) : (
                  (index + 1).toString()
                )}
              </div>

              {/* Step label */}
              <p
                title={
                  isLTR
                    ? (step.descriptionEN ?? step.displayNameEN)
                    : (step.descriptionAR ?? step.displayNameAR)
                }
              >
                {isLTR ? step.displayNameEN : step.displayNameAR}
              </p>

              {/* Connector */}
              {!isLast && <div className="connector"></div>}
            </div>
          );
        })
      ) : (
        <div>No steps found.</div>
      )}
    </div>
  );
};
