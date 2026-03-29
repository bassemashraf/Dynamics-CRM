/* eslint-disable */
/**
 * ProcessExtensionHelpers.ts
 * Client-side helpers that simulate the server-side C# plugins:
 *   - AutoCreateProcessExtension (process_extention_oncreate.cs)
 *   - Create_PostOperation_SetDefaults (Process_extention_set_defaults.cs)
 *
 * Instead of running on server create, the process definition is resolved
 * from the incident type selected by the user in the Multi Inspection popup.
 * Uses Xrm.WebApi for offline compatibility (no $expand).
 */

export interface ProcessDefinitionData {
  id: string;
  name?: string;
  /** duc_startstage lookup */
  startStageId?: string;
  startStageName?: string;
  /** duc_defaultstatus lookup */
  defaultStatusId?: string;
  defaultStatusName?: string;
  /** duc_defaultsubstatus lookup */
  defaultSubStatusId?: string;
  defaultSubStatusName?: string;
  /** duc_startaction lookup */
  startActionId?: string;
  startActionName?: string;
  /** duc_targetentitycustomerlookupname — field name on work order for customer */
  customerLookupName?: string;
  /** duc_targetentitysubjectname — field name on work order for subject */
  subjectFieldName?: string;
  /** duc_parentlookup — field name on work order for parent entity */
  parentLookupName?: string;
}

export class ProcessExtensionHelpers {
  private static get xrm(): Xrm.XrmStatic {
    return (window as any).Xrm;
  }

  // =====================================================================
  // GET PROCESS DEFINITION FROM INCIDENT TYPE
  // =====================================================================

  /**
   * Retrieve the process definition linked to the selected incident type.
   *
   * Step 1: Read the duc_processdefinition lookup from msdyn_incidenttype.
   * Step 2: Retrieve the full duc_processdefinition record by its ID
   *         to get all fields needed for creating the process extension.
   *
   * Uses Xrm.WebApi (no $expand) for offline compatibility.
   *
   * @param incidentTypeId  The ID of the incident type selected in the popup
   */
  static async getProcessDefinitionFromIncidentType(
    incidentTypeId: string,
  ): Promise<ProcessDefinitionData | null> {
    try {
      // Step 1: Get the process definition lookup from the incident type
      const incidentType = await this.xrm.WebApi.retrieveRecord(
        "msdyn_incidenttype",
        incidentTypeId,
        "?$select=_duc_processdefinition_value",
      );

      const processDefId = incidentType["_duc_processdefinition_value"];
      if (!processDefId) {
        const msg =
          "[ProcessExtension] Incident type has no process definition linked";
        console.warn(msg);
        return null;
      }

      const processDefName =
        incidentType[
        "_duc_processdefinition_value@OData.Community.Display.V1.FormattedValue"
        ];
      console.log(
        "[ProcessExtension] Incident type linked to process definition:",
        processDefName,
        processDefId,
      );

      // Step 2: Retrieve the full process definition record
      const def = await this.xrm.WebApi.retrieveRecord(
        "duc_processdefinition",
        processDefId,
        "?$select=duc_processdefinitionid,duc_name" +
        ",_duc_startstage_value" +
        ",_duc_defaultstatus_value" +
        ",_duc_defaultsubstatus_value" +
        ",_duc_startaction_value" +
        ",duc_targetentitycustomerlookupname" +
        ",duc_targetentitysubjectname" +
        ",duc_parentlookup",
      );

      const result: ProcessDefinitionData = {
        id: def.duc_processdefinitionid,
        name: def.duc_name,
        startStageId: def["_duc_startstage_value"],
        startStageName:
          def[
          "_duc_startstage_value@OData.Community.Display.V1.FormattedValue"
          ],
        defaultStatusId: def["_duc_defaultstatus_value"],
        defaultStatusName:
          def[
          "_duc_defaultstatus_value@OData.Community.Display.V1.FormattedValue"
          ],
        defaultSubStatusId: def["_duc_defaultsubstatus_value"],
        defaultSubStatusName:
          def[
          "_duc_defaultsubstatus_value@OData.Community.Display.V1.FormattedValue"
          ],
        startActionId: def["_duc_startaction_value"],
        startActionName:
          def[
          "_duc_startaction_value@OData.Community.Display.V1.FormattedValue"
          ],
        customerLookupName: def.duc_targetentitycustomerlookupname,
        subjectFieldName: def.duc_targetentitysubjectname,
        parentLookupName: def.duc_parentlookup,
      };


      return result;
    } catch (error: any) {
      const msg = `[ProcessExtension] Error retrieving process definition from incident type: ${error?.message || error}`;
      console.error(msg, error);
      return null;
    }
  }

  // =====================================================================
  // CREATE PROCESS EXTENSION RECORD
  // =====================================================================

  /**
   * Create a duc_processextension record linked to a work order,
   * setting default values from the process definition.
   *
   * Mirrors the C# plugins:
   *   - AutoCreateProcessExtension: creates record with RegardingObjectId
   *   - Create_PostOperation_SetDefaults: sets processDefinition, CurrentStage,
   *     Status, SubStatus, AssignmentDate
   *
   * @param workOrderId  The ID of the created work order
   * @param processDef   The process definition data to use for defaults
   * @param customerId   Optional customer (service account) to set on the PE
   * @returns The created process extension record ID, or null on failure
   */
  static async createProcessExtension(
    workOrderId: string,
    processDef: ProcessDefinitionData,
    customerId?: string,
  ): Promise<string | null> {
    try {
      // ================================================================
      // Build the full create payload with ALL lookups up front.
      // Uses the correct relationship schema names as navigation
      // property names for @odata.bind (from entities.cs N:1 defs).
      // ================================================================
      const createData: any = {
        // Regarding → Work Order (polymorphic activity lookup)
        "regardingobjectid_msdyn_workorder@odata.bind": `/msdyn_workorders(${workOrderId})`,
        // Assignment date
        duc_assignmentdate: new Date(),
      };

      // Process Definition
      // Nav prop: duc_processDefinition_duc_ProcessExtension
      if (processDef.id) {
        createData["duc_processDefinition_duc_ProcessExtension@odata.bind"] =
          `/duc_processdefinitions(${processDef.id})`;
      }

      // Current Stage (from Start Stage)
      // Nav prop: duc_CurrentStage_duc_ProcessExtension
      // if (processDef.startStageId) {
      createData["duc_CurrentStage_duc_ProcessExtension@odata.bind"] =
        `/duc_processstages(${processDef.startStageId})`;



      // Status
      // Nav prop: duc_Status_duc_ProcessExtension
      // if (processDef.defaultStatusId) {
      createData["duc_Status_duc_ProcessExtension@odata.bind"] =
        `/duc_processstatuses(${processDef.defaultStatusId})`;
      // }

      // SubStatus
      // Nav prop: duc_SubStatus_duc_ProcessExtension
      // if (processDef.defaultSubStatusId) {
      createData["duc_SubStatus_duc_ProcessExtension@odata.bind"] =
        `/duc_processsubstatuses(${processDef.defaultSubStatusId})`;
      // }

      // Customer (account — polymorphic lookup)
      // Nav prop: duc_CustomerId_duc_ProcessExtension_account
      // if (customerId) {
      createData["duc_CustomerId_duc_ProcessExtension_account@odata.bind"] =
        `/accounts(${customerId})`;
      // }

      createData["subject"] = "test";

      // Last Action Taken (from Start Action on process definition)
      // Nav prop: duc_LastActionTaken_duc_ProcessExtension
      // if (processDef.startActionId) {
      //   createData["duc_LastActionTaken_duc_ProcessExtension@odata.bind"] =
      //     `/duc_stageactions(01510048-cfff-4830-8e20-0552fe56a865)`;
      // }

      const result = await this.xrm.WebApi.createRecord(
        "duc_processextension",
        createData,
      );
      const peId = result.id;

      // ================================================================
      // Link PE back to work order
      // ================================================================
      try {
        await this.xrm.WebApi.updateRecord("msdyn_workorder", workOrderId, {
          "duc_processextension@odata.bind": `/duc_processextensions(${peId})`,
        });
      } catch (linkError: any) {
        console.error(`⚠️ Could not link PE to WO: ${linkError?.message || linkError}`);
      }

      return peId;
    } catch (error: any) {
      let errorDetails = error?.message || String(error);
      try {
        if (error?.raw) errorDetails += "\n\nRaw: " + JSON.stringify(error.raw);
        if (error?.innerError)
          errorDetails += "\n\nInnerError: " + JSON.stringify(error.innerError);
      } catch (e) {
        /* ignore */
      }

      console.error(`❌ Error creating PE:\n${errorDetails}`);
      return null;
    }
  }

  // =====================================================================
  // COMBINED HELPER — CONVENIENCE METHOD
  // =====================================================================

  /**
   * Full flow: retrieve the process definition from the incident type,
   * then create a process extension record linked to the given work order.
   *
   * This is the main entry point to call after creating a work order
   * in MultiTypeInspection.ts.
   *
   * @param workOrderId       The ID of the newly created work order
   * @param incidentTypeId    The ID of the selected incident type (has duc_processdefinition lookup)
   * @param serviceAccountId  Optional service account ID for customer field
   * @returns The created process extension ID, or null if no definition found or on failure
   */
  static async createProcessExtensionForWorkOrder(
    workOrderId: string,
    incidentTypeId: string,
    serviceAccountId?: string,
  ): Promise<{ peId: string | null; processDef: ProcessDefinitionData | null }> {
    // Step 1: Get the process definition from the incident type
    const processDef =
      await this.getProcessDefinitionFromIncidentType(incidentTypeId);

    if (!processDef) {
      const msg =
        "[ProcessExtension] No process definition found on incident type — skipping process extension creation";
      console.warn(msg);
      return { peId: null, processDef: null };
    }

    console.log(
      "[ProcessExtension] Found process definition:",
      processDef.name,
      processDef.id,
    );

    // Step 2: Create the process extension with defaults
    const peId = await this.createProcessExtension(
      workOrderId,
      processDef,
      serviceAccountId,
    );

    return { peId, processDef };
  }

  // =====================================================================
  // SUBSTATUS RESOLUTION
  // =====================================================================

  /**
   * Retrieves the raw GUID string (duc_value) from the process substatus record.
   * This is used to set the actual work order substatus (msdyn_substatus).
   *
   * @param defaultSubStatusId The ID of the duc_processsubstatus record
   */
  static async getSubStatusValueFromProcessDefinition(
    defaultSubStatusId: string,
  ): Promise<string | null> {
    try {
      const subStatusRecord = await this.xrm.WebApi.retrieveRecord(
        "duc_processsubstatus",
        defaultSubStatusId,
        "?$select=duc_value"
      );

      const subStatusGuid = subStatusRecord?.duc_value;
      if (!subStatusGuid) {
        console.warn("[ProcessExtension] duc_processsubstatus has no duc_value");
        return null;
      }
      return subStatusGuid;
    } catch (error: any) {
      console.error("[ProcessExtension] Error getting substatus value:", error);
      return null;
    }
  }
}
