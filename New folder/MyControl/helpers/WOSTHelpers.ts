/* eslint-disable */
/**
 * WOSTHelpers.ts
 * Client-side offline-safe equivalent of WOST_Create_PostOperation.cs
 *
 * Called after Work Order creation to:
 *   1. Create msdyn_workorderincident (OOB creates this server-side only)
 *   2. Create msdyn_workorderservicetask records from incident type templates
 *   3. Create duc_questionanswersconfiguration per inspection question
 *   4. Create duc_workorderpenalties per penalty linked to the incident type
 */

export interface WOSTCreationResult {
    wostId: string;
    questionsCreated: number;
    penaltiesCreated: number;
}

export class WOSTHelpers {

    private static get xrm(): Xrm.XrmStatic {
        return (window.parent as any).Xrm || (window as any).Xrm;
    }

    // =========================================================================
    // Create msdyn_workorderincident offline
    // OOB creates this server-side only — we replicate it here for offline support
    // @param incidentTypeName  from incidentTypeData.name — already in scope
    // =========================================================================

    static async createWorkOrderIncident(
        workOrderId: string,
        incidentTypeId: string,
        incidentTypeName: string
    ): Promise<string | null> {
        try {

            //alert(`[WOI] Creating — WO: ${workOrderId} | IncidentType: ${incidentTypeId} | Name: ${incidentTypeName}`);


            const result = await this.xrm.WebApi.createRecord(
                "msdyn_workorderincident",
                {
                    "msdyn_workorder@odata.bind":
                        `/msdyn_workorders(${workOrderId})`,
                    "msdyn_incidenttype@odata.bind":
                        `/msdyn_incidenttypes(${incidentTypeId})`,
                    "msdyn_name":incidentTypeName,
                    "msdyn_isprimary": true,
                    "msdyn_ismobile": true,
                }
            );

            //alert(`[WOI] Created successfully: ${result?.id}`);

            console.log("[WOSTHelpers] WorkOrderIncident created:", result?.id);
            return result?.id ?? null;
        } catch (error: any) {
            alert(`[WOI] ERROR: ${error?.message || ""} | Raw: ${JSON.stringify(error?.raw || error?.innererror || "")}`);
            console.error("[WOSTHelpers] Error creating WorkOrderIncident:", error);
            return null;
        }
    }

    // =========================================================================
    // MAIN ENTRY POINT
    // @param workOrderId         Newly created WO ID
    // @param incidentTypeId      Incident type selected in the popup
    // @param workOrderIncidentId Created by createWorkOrderIncident() above,
    //                            passed in from MultiTypeInspection.ts
    // =========================================================================

    static async createWOSTsForWorkOrder(
        workOrderId: string,
        incidentTypeId: string,
        workOrderIncidentId: string | null
    ): Promise<WOSTCreationResult[]> {

        const results: WOSTCreationResult[] = [];

        try {
            console.log("[WOSTHelpers] Starting WOST creation - WO:", workOrderId, "| IncidentType:", incidentTypeId, "| WOIncident:", workOrderIncidentId);

            const serviceTasks = await this.getIncidentTypeServiceTasks(incidentTypeId);

            //alert(`[WOST] getIncidentTypeServiceTasks returned: ${serviceTasks.length} | incidentTypeId: ${incidentTypeId}`);

            if (serviceTasks.length === 0) {
                console.warn("[WOSTHelpers] No service task templates found for incident type:", incidentTypeId);
                return results;
            }

            console.log(`[WOSTHelpers] Found ${serviceTasks.length} service task template(s)`);

            for (let i = 0; i < serviceTasks.length; i++) {
                const taskTemplate = serviceTasks[i];
                try {
                    const wostId = await this.createWorkOrderServiceTask(
                        workOrderId,
                        incidentTypeId,
                        taskTemplate,
                        workOrderIncidentId
                    );

                    if (!wostId) {
                        console.warn("[WOSTHelpers] WOST creation returned null - skipping questions/penalties for template:", taskTemplate?.msdyn_name);
                        continue;
                    }

                    console.log("[WOSTHelpers] WOST created:", wostId);

                    const questionsCreated = await this.createQuestionAnswerConfigurations(
                        wostId,
                        workOrderId,
                        incidentTypeId
                    );

                    const penaltiesCreated = await this.createWorkOrderPenalties(
                        wostId,
                        workOrderId,
                        incidentTypeId
                    );

                    results.push({ wostId, questionsCreated, penaltiesCreated });

                } catch (templateError: any) {
                    console.error("[WOSTHelpers] Error processing template index", i, ":", templateError);
                }
            }

            console.log(`[WOSTHelpers] Done. ${results.length} WOST(s) created.`);
            return results;

        } catch (error: any) {
            console.error("[WOSTHelpers] Fatal error in createWOSTsForWorkOrder:", error);
            return results;
        }
    }

    // =========================================================================
    // Fetch incident type service task templates
    // =========================================================================

    private static async getIncidentTypeServiceTasks(incidentTypeId: string): Promise<any[]> {
        try {
            const result = await this.xrm.WebApi.retrieveMultipleRecords(
                "msdyn_incidenttypeservicetask",
                `?$select=msdyn_incidenttypeservicetaskid,msdyn_name,msdyn_tasktype` +
                `&$filter=msdyn_incidenttype eq '${incidentTypeId}'`
            );
            return Array.from(result?.entities ?? []);
        } catch (error: any) {
            alert(`[WOST] getIncidentTypeServiceTasks ERROR: ${error?.message || JSON.stringify(error)}`);
            console.error("[WOSTHelpers] Error fetching incident type service tasks:", error);
            return [];
        }
    }

    // =========================================================================
    // Create msdyn_workorderservicetask
    // duc_iscreatedoffline = true — plugin will return early post-sync
    // =========================================================================

    private static async createWorkOrderServiceTask(
        workOrderId: string,
        incidentTypeId: string,
        taskTemplate: any,
        workOrderIncidentId: string | null
    ): Promise<string | null> {
        try {
           
            const wostData: any = {
                "msdyn_workorder@odata.bind":
                    `/msdyn_workorders(${workOrderId})`,
                "duc_IncidentType@odata.bind":
                    `/msdyn_incidenttypes(${incidentTypeId})`,
                "duc_iscreatedoffline": true,
                "msdyn_name": taskTemplate.msdyn_name || "Service Task",
            };

            const taskTypeId =
                taskTemplate["_msdyn_tasktype_value"] ??
                taskTemplate["msdyn_tasktype"];

            if (taskTypeId) {
                wostData["msdyn_tasktype@odata.bind"] =
                    `/msdyn_servicetasktypes(${taskTypeId})`;
            }

            if (workOrderIncidentId) {
                wostData["msdyn_workorderincident@odata.bind"] =
                    `/msdyn_workorderincidents(${workOrderIncidentId})`;
            }

            const result = await this.xrm.WebApi.createRecord(
                "msdyn_workorderservicetask",
                wostData
            );

        //alert(`[WOST] WOST create SUCCESS: ${result?.id}`);            

            return result?.id ?? null;

        } catch (error: any) {
             alert(`[WOST Create] ERROR: ${error?.message || ""} | Raw: ${JSON.stringify(error?.raw || error?.innererror || "")}`);
            console.error("[WOSTHelpers] Error creating WOST:", error);
            return null;
        }
    }

    // =========================================================================
    // Replicate Plugin PART 1 — Create duc_questionanswersconfiguration
    // Guard: skip if records already exist for this WOST
    // duc_image omitted — Primary Image field, not settable via createRecord
    // =========================================================================

    private static async createQuestionAnswerConfigurations(
        wostId: string,
        workOrderId: string,
        incidentTypeId: string
    ): Promise<number> {
        try {

            const existingCheck = await this.xrm.WebApi.retrieveMultipleRecords(
                "duc_questionanswersconfiguration",
                `?$select=duc_questionanswersconfigurationid` +
                `&$filter=duc_msdyn_workorderservicetask eq '${wostId}'` +
                `&$top=1`
            );

            if ((existingCheck?.entities?.length ?? 0) > 0) {
                console.log("[WOSTHelpers] Question answers already exist for WOST — skipping.");
                return 0;
            }
       
            const questionsResult = await this.xrm.WebApi.retrieveMultipleRecords(
                "duc_inspectionquestionconfiguration",
                `?$select=duc_inspectionquestionconfigurationid,duc_questionname,_duc_questioncategory_value` +
                `&$filter=_duc_primaryincidenttype_value eq '${incidentTypeId}'`
            );

            const questions = Array.from(questionsResult?.entities ?? [])
                .filter(q => {
                    // Client-side filter — offline engine ignores the $filter on lookup fields
                    const qIncidentTypeId =
                        q["_duc_primaryincidenttype_value"] ??
                        q["duc_primaryincidenttype"];
                    return qIncidentTypeId?.toLowerCase() === incidentTypeId?.toLowerCase();
                });


            //alert(`[Questions] found: ${questions.length} for incidentTypeId: ${incidentTypeId}`);

            if (questions.length === 0) {
                console.log("[WOSTHelpers] No inspection questions found for incident type:", incidentTypeId);
                return 0;
            }

            let created = 0;

            for (let i = 0; i < questions.length; i++) {
                const question = questions[i];
                try {
                    const configId = question.duc_inspectionquestionconfigurationid;
                    const questionName = question.duc_questionname;

                    const categoryId =
                        question["_duc_questioncategory_value"] ??
                        question["duc_questioncategory"];

                    if (!configId) {
                        console.warn("[WOSTHelpers] Skipping question — missing configId");
                        continue;
                    }

                    if (!categoryId) {
                        console.warn("[WOSTHelpers] Skipping question — missing categoryId for:", questionName);
                        continue;
                    }

                        await this.xrm.WebApi.createRecord(
                            "duc_questionanswersconfiguration",
                            {
                                "duc_questionname": questionName,
                                "duc_answercontent": questionName,
                                "duc_InspectionQuestionConfiguration@odata.bind":
                                    `/duc_inspectionquestionconfigurations(${configId})`,
                                "duc_QuestionCategory@odata.bind":
                                    `/duc_inspectionquestioncategorieses(${categoryId})`,
                                "duc_msdyn_workorder@odata.bind":
                                    `/msdyn_workorders(${workOrderId})`,
                                "duc_msdyn_workorderservicetask@odata.bind":
                                    `/msdyn_workorderservicetasks(${wostId})`,
                            }
);

                    created++;
                    console.log(`[WOSTHelpers] Created question answer config: ${questionName}`);

                } catch (qError: any) {
                    alert(`[Questions] CREATE ERROR at index ${i}: ${qError?.message}`);
                    console.error("[WOSTHelpers] Error creating question answer config at index", i, ":", qError);
                }
            }

            return created;

        } catch (error: any) {
            console.error("[WOSTHelpers] Error in createQuestionAnswerConfigurations:", error);
            return 0;
        }
    }

    // =========================================================================
    // Replicate Plugin PART 2 — Create duc_workorderpenalties
    // Guard: skip if records already exist for this WOST
    // Mirrors plugin tax type filter: INCOME=100000000, EXCISE=100000001
    // =========================================================================

    private static async createWorkOrderPenalties(
        wostId: string,
        workOrderId: string,
        incidentTypeId: string
    ): Promise<number> {
        try {
            const existingCheck = await this.xrm.WebApi.retrieveMultipleRecords(
                "duc_workorderpenalties",
                `?$select=duc_workorderpenaltiesid` +
                `&$filter=duc_workorderservicetask eq '${wostId}'` +
                `&$top=1`
            );

            if ((existingCheck?.entities?.length ?? 0) > 0) {
                console.log("[WOSTHelpers] Penalties already exist for WOST — skipping.");
                return 0;
            }

            // Get WO tax type to apply correct penalty filter
            let taxType: number | null = null;
            try {
                const woRecord = await this.xrm.WebApi.retrieveRecord(
                    "msdyn_workorder",
                    workOrderId,
                    "?$select=duc_taxtype"
                );
                if (woRecord?.duc_taxtype != null) {
                    taxType = woRecord.duc_taxtype;
                }
            } catch (taxError: any) {
                console.warn("[WOSTHelpers] Could not read duc_taxtype:", taxError?.message);
            }

            console.log(`[WOSTHelpers] WO tax type = ${taxType ?? "NULL — no tax filter"}`);

            const INCOME = 100000000;
            const EXCISE = 100000001;

            let penaltyFilter = `duc_incidenttype eq '${incidentTypeId}'`;
            if (taxType === INCOME) penaltyFilter += ` and duc_taxtype eq ${INCOME}`;
            if (taxType === EXCISE) penaltyFilter += ` and duc_taxtype eq ${EXCISE}`;

            const penaltiesResult = await this.xrm.WebApi.retrieveMultipleRecords(
                "duc_penalty",
                `?$select=duc_penaltyid,duc_maximumamount` +
                `&$filter=${penaltyFilter}`
            );

            const penalties = Array.from(penaltiesResult?.entities ?? []);

           //alert(`[Penalties] found: ${penalties.length} for incidentTypeId: ${incidentTypeId}`);


            if (penalties.length === 0) {
                console.log("[WOSTHelpers] No penalties found for incident type:", incidentTypeId);
                return 0;
            }

            let created = 0;

            for (let i = 0; i < penalties.length; i++) {
                const penalty = penalties[i];
                try {
                    const penaltyId = penalty.duc_penaltyid;

                    if (!penaltyId) {
                        console.warn("[WOSTHelpers] Skipping penalty at index", i, "— missing duc_penaltyid");
                        continue;
                    }

                    const penaltyData: any = {
                        "duc_Penalty@odata.bind":
                            `/duc_penalties(${penaltyId})`,
                        "duc_WorkOrderServiceTask@odata.bind":
                            `/msdyn_workorderservicetasks(${wostId})`,
                        "duc_WorkOrder@odata.bind":
                            `/msdyn_workorders(${workOrderId})`,
                    };

                    // duc_maximumamount → duc_penaltyamount (mirrors plugin mapping)
                    if (penalty.duc_maximumamount != null) {
                        penaltyData["duc_penaltyamount"] = penalty.duc_maximumamount;
                    }

                    await this.xrm.WebApi.createRecord(
                        "duc_workorderpenalties",
                        penaltyData
                    );

                    created++;
                    console.log(`[WOSTHelpers] Created penalty: ${penaltyId}`);

                } catch (pError: any) {
                    alert(`[Penalties] CREATE ERROR at index ${i}: ${pError?.message}`);
                    console.error("[WOSTHelpers] Error creating penalty at index", i, ":", pError);
                }
            }

            return created;

        } catch (error: any) {
            console.error("[WOSTHelpers] Error in createWorkOrderPenalties:", error);
            return 0;
        }
    }
}