import * as React from "react";
import { useRef } from "react";

export interface IStep {
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

export interface IFlowProgressBarProps {
  steps: IStep[];
  isLTR: boolean;
  isMobileOrTablet: boolean;
}

export const FlowProgressBar: React.FC<IFlowProgressBarProps> = ({
  steps,
  isLTR,
  isMobileOrTablet
}) => {
  const containerRef = useRef<HTMLDivElement>(null);

  const renderCheckmark = () => (
    <svg style={{ width: "20px" }} fill="#000000" version="1.1" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 335.765 335.765">
      <g>
        <polygon fill="white" points="311.757,41.803 107.573,245.96 23.986,162.364 0,186.393 107.573,293.962 335.765,65.795" />
      </g>
    </svg>
  );

  return (
    <div
      className={`main-pcf-flow-container ${isLTR ? "ltr" : "rtl"} ${
        isMobileOrTablet ? "main-pcf-flow-container-mobile-padding" : "main-pcf-flow-container-web-padding"
      }`}
      ref={containerRef}
    >
      {steps.length > 0 ? (
        steps.map((step, index) => {
          const isLast = index === steps.length - 1;
          const status = step.done ? "done" : step.active ? "active" : "";

          return (
            <React.Fragment key={step.id}>
              <div className={`step ${status}`}>
                <div
                  className="step-number"
                  title={isLTR ? (step.descriptionEN ?? step.displayNameEN) : (step.descriptionAR ?? step.displayNameAR)}
                >
                  {step.done ? renderCheckmark() : (index + 1).toString()}
                </div>
                <p title={isLTR ? (step.descriptionEN ?? step.displayNameEN) : (step.descriptionAR ?? step.displayNameAR)}>
                  {isLTR ? step.displayNameEN : step.displayNameAR}
                </p>
              </div>
              {!isLast && <div className={`connector ${step.done ? "done" : ""}`}></div>}
            </React.Fragment>
          );
        })
      ) : (
        <div>No steps to display.</div>
      )}
    </div>
  );
};
