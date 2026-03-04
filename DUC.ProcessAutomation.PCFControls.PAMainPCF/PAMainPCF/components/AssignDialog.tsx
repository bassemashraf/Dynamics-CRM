import * as React from "react";
import { Dialog, DialogType, DialogFooter, PrimaryButton, DefaultButton, TextField, Pivot, PivotItem, Persona, PersonaSize, MessageBar, MessageBarType, Spinner } from "@fluentui/react";
import { IActionButtonProps } from "./Actions";
import { Constants, getAssignedTodayFetchXML, getAssignFetchXML, getCheckOnLeaveFetchXML } from "../constants";
import { CustomBadge } from "./CustomBadge";

interface AssignDialogProps {
  isVisible: boolean;
  onClose: () => void;
  onAssign?: (assignee: { id: string; name: string; type: "user" | "team", entityType: string }) => void;
  onSetNextAssigee?: (action: IActionButtonProps, assignee: { id: string; name: string; type: "user" | "team", entityType: string }) => Promise<void>;
  action?: IActionButtonProps;
  _context: any; // Xrm context
  teamId?: string;
}

interface AssigneeItem {
  id: string;
  name: string;
  email?: string;
  type: "user" | "team";
  assignedToday?: number;
  assignedCount?: number;
  capacity?: number;
  onLeave?: boolean;
  entityType: string;
}

const AssignDialog: React.FC<AssignDialogProps> = ({ _context, isVisible, onClose, onAssign, onSetNextAssigee, action, teamId }) => {
  const [loading, setLoading] = React.useState(false);
  const [items, setItems] = React.useState<AssigneeItem[]>([]);
  const [search, setSearch] = React.useState("");
  const [tab, setTab] = React.useState<"user" | "team">("user");
  const [selected, setSelected] = React.useState<AssigneeItem | null>(null);
  const [message, setMessage] = React.useState<{ text: string; type: MessageBarType } | null>(null);

  // 🔄 Load assignees whenever tab changes
  React.useEffect(() => {
    if (!isVisible) return;
    setLoading(true);
    setSelected(null);

    const fetchAssignees = async () => {
      try {
        // 1️⃣ Fetch users/teams
        let filter = "";
        if (teamId) {
          filter = `&$filter=(teammembership_association/any(o1:(o1/teamid eq ${teamId}))) `;
        }
        const entityName = tab === "user" ? "systemuser" : "team";
        const columns = tab === "user" ? "fullname,internalemailaddress" : "name";

        const res = await _context.webAPI.retrieveMultipleRecords(entityName, `?$select=${columns}${filter}`);
        const usersLst: AssigneeItem[] = res.entities.map((e: any) => ({
          id: e.systemuserid ?? e.teamid,
          name: e.fullname ?? e.name,
          email: e.internalemailaddress,
          type: tab,
          entityType: tab === "user" ? "systemuser" : "team",
          capacity: 0
        }));

        setItems(usersLst); // Display users immediately

        // 2️⃣ Fetch counts & leave in parallel for each user
        const updateCounts = async (user: AssigneeItem) => {
          const todayStr = new Date().toISOString().split("T")[0];

          const fetchAssignedCount = async (): Promise<number> => {

            const result = await _context.webAPI.retrieveMultipleRecords(Constants.MAIN_ENTITY_NAME, "?fetchXml=" + encodeURIComponent(getAssignFetchXML(user.id)));
            return result.entities.length > 0 ? result.entities[0].request_count as number : 0;
          };

          const fetchAssignedToday = async (): Promise<number> => {

            const result = await _context.webAPI.retrieveMultipleRecords(Constants.MAIN_ENTITY_NAME, "?fetchXml=" + encodeURIComponent(getAssignedTodayFetchXML(user.id, todayStr)));
            return result.entities.length > 0 ? result.entities[0].today_count as number : 0;
          };

          const checkOnLeave = async (): Promise<boolean> => {

            const result = await _context.webAPI.retrieveMultipleRecords(Constants.EMPLOYEE_LEAVE_ENTITY, "?fetchXml=" + encodeURIComponent(getCheckOnLeaveFetchXML(user.id, todayStr)));
            return result.entities.length > 0;
          };

          try {
            const [assignedCount, assignedToday, onLeave] = await Promise.all([
              fetchAssignedCount(),
              fetchAssignedToday(),
              checkOnLeave()
            ]);
            user.assignedCount = assignedCount;
            user.assignedToday = assignedToday;
            user.onLeave = onLeave;
            setItems(prev => [...prev]); // Trigger re-render
          } catch (err) {
            console.error("Error updating counts for user", user.id, err);
          }
        };

        await Promise.all(usersLst.map(u => updateCounts(u))); // Run all users in parallel

      } catch (err) {
        setMessage({ text: "Failed to load assignees", type: MessageBarType.error });
      } finally {
        setLoading(false);
      }
    };

    void fetchAssignees();
  }, [tab, isVisible, teamId, _context]);

  const filteredItems = items.filter(i => i.name.toLowerCase().includes(search.toLowerCase()));

  const handleAssign = () => {
    if (!selected) {
      setMessage({ text: "Please select a user or team", type: MessageBarType.warning });
      return;
    }
    if (onAssign) void onAssign(selected);
    if (onSetNextAssigee && action) void onSetNextAssigee(action, selected);
    onClose();
  };

  return (
    <Dialog
      hidden={!isVisible}
      onDismiss={onClose}
      minWidth="40%"
      dialogContentProps={{
        type: DialogType.largeHeader,
        title: _context.resources.getString("Assigned_Record") ?? "Assign Record",
      }}
    >
      {message && (
        <MessageBar messageBarType={message.type} onDismiss={() => setMessage(null)}>
          {message.text}
        </MessageBar>
      )}

      <TextField
        placeholder={_context.resources.getString("Search") ?? "Search..."}
        value={search}
        onChange={(_, v) => setSearch(v ?? "")}
        iconProps={{ iconName: "Search" }}
        styles={{ root: { marginBottom: 10 } }}
      />

      <Pivot selectedKey={tab} onLinkClick={item => setTab(item?.props.itemKey as "user" | "team")}>
        <PivotItem headerText="Users" itemKey="user" />
        {/* <PivotItem headerText="Teams" itemKey="team" /> */}
      </Pivot>

      <div style={{ maxHeight: "300px", overflowY: "auto", border: "1px solid #ddd", borderRadius: 4, marginTop: 10 }}>
        {loading && items.length === 0 ? (
          <Spinner label={_context.resources.getString(Constants.LOADING) ?? "Loading..."} />
        ) : (
          filteredItems.map(item => (
            <div
              key={item.id}
              onClick={() => setSelected(item)}
              style={{
                display: "flex",
                alignItems: "center",
                padding: "8px",
                cursor: "pointer",
                backgroundColor: selected?.id === item.id ? "#e6f7ff" : "transparent",
              }}
            >
              <Persona
                text={item.name}
                secondaryText={item.email}
                size={PersonaSize.size40}
                initialsColor={"darkblue"} // Use same color as assign button
              />
              <div style={{ display: "flex", flexDirection: "row", gap: "8px", flexWrap: "wrap" }}>
                {item.assignedCount !== undefined && (
                  <CustomBadge
                    label={_context.resources.getString("Assigned_Label")}
                    value={item.assignedCount}
                  />
                )}

                {item.assignedToday !== undefined && (
                  <CustomBadge
                    label={_context.resources.getString("AssignedToday_Label")}
                    value={item.assignedToday}
                  />
                )}

                {item.capacity !== undefined && (
                  <CustomBadge
                    label={_context.resources.getString("Capacity_Label")}
                    value={item.capacity}
                  />
                )}

                {item.onLeave && (
                  <CustomBadge
                    label={_context.resources.getString("OnLeave_Label")}
                    isLeave={true}
                    iconClass="fas fa-calendar-minus"
                  />
                )}
              </div>

            </div>
          ))
        )}
      </div>

      <DialogFooter>
        {onSetNextAssigee && action && (
          <PrimaryButton
            text={action.displayName}
            style={{ backgroundColor: action.buttonColor }}
            onClick={handleAssign}
            disabled={!selected}
          />
        )}

        {onAssign && (
          <PrimaryButton
            text={_context.resources.getString(Constants.ASSIGNMENT_BTN_LABEL)}
            onClick={handleAssign}
            disabled={!selected}
            style={{ backgroundColor: "darkblue" }}
          />
        )}

        <DefaultButton onClick={onClose} text={_context.resources.getString(Constants.CLOSE_MODAL)} />
      </DialogFooter>
    </Dialog>
  );
};

export default AssignDialog;
