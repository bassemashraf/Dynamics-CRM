import * as React from "react";

interface CustomBadgeProps {
  label: string;
  value?: string | number;
  isLeave?: boolean;
  iconClass?: string;
}

export const CustomBadge: React.FC<CustomBadgeProps> = ({ label, value, isLeave, iconClass }) => {
  return (
    <div
      style={{
        backgroundColor: isLeave ? "#ffe5e5" : "#f0f0f0",
        borderRadius: "12px",
        padding: "4px 10px",
        fontSize: "12px",
        color: isLeave ? "red" : "#333",
        display: "flex",
        alignItems: "center",
        fontWeight: isLeave ? "bold" : "normal",
      }}
    >
      {iconClass && <i className={iconClass} style={{ marginRight: "5px" }}></i>}
      {label} {value !== undefined ? `: ${value}` : ""}
    </div>
  );
};
