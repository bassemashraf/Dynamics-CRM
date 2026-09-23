/* eslint-disable */
/**
 * ProjectHelpers.ts
 * Reusable helper functions for Project (duc_accounttype = 16) operations.
 * Handles Logic App API calls to find nearby projects and retrieve
 * project/account data from Dynamics CRM.
 *
 * Supports both online and offline modes using Xrm.WebApi.
 */

// =====================================================================
// CONFIGURATION
// =====================================================================

/**
 * Logic App API endpoint for searching nearby projects.
 * TODO: Replace with actual Logic App URL once provided.
 * When empty (""), the helper returns dummy data for development/testing.
 */
const PROJECT_SEARCH_API_URL = "";

/**
 * TODO: Get the actual radius from a CRM setting or user input.
 * For now, hardcoded to 5000 meters as a default search radius.
 */
const DEFAULT_SEARCH_RADIUS = 5000;

// =====================================================================
// INTERFACES
// =====================================================================

/** Request payload sent to the Logic App API */
export interface INearbyProjectRequest {
    radius: number;
    accountType: number;
    longitude: number;
    latitude: number;
    departmentId: string;
}

/** Single project returned from the Logic App API */
export interface INearbyProjectResponse {
    projectId: string;
    longitude: number;
    latitude: number;
}

/** Enriched project detail for display in the selection list */
export interface IProjectDetail {
    projectId: string;
    arabicName: string;
    projectNumber: string;   // duc_id
    accountName: string;
    accountId: string;
}

// =====================================================================
// DUMMY DATA (used when API URL is not configured)
// =====================================================================

const DUMMY_NEARBY_PROJECTS: INearbyProjectResponse[] = [
    { projectId: "aaa11111-1111-1111-1111-111111111111", longitude: 51.5310, latitude: 25.2860 },
    { projectId: "bbb22222-2222-2222-2222-222222222222", longitude: 51.5200, latitude: 25.2900 },
    { projectId: "ccc33333-3333-3333-3333-333333333333", longitude: 51.5400, latitude: 25.2800 },
    { projectId: "ddd44444-4444-4444-4444-444444444444", longitude: 51.5150, latitude: 25.2950 },
    { projectId: "eee55555-5555-5555-5555-555555555555", longitude: 51.5350, latitude: 25.2750 },
];

const DUMMY_PROJECT_DETAILS: IProjectDetail[] = [
    {
        projectId: "aaa11111-1111-1111-1111-111111111111",
        arabicName: "مشروع الواجهة البحرية",
        projectNumber: "PRJ-2025-001",
        accountName: "Al Wakra Waterfront LLC",
        accountId: "acc-11111-1111-1111-1111-111111111111",
    },
    {
        projectId: "bbb22222-2222-2222-2222-222222222222",
        arabicName: "مشروع الحديقة المركزية",
        projectNumber: "PRJ-2025-002",
        accountName: "Central Park Development Co.",
        accountId: "acc-22222-2222-2222-2222-222222222222",
    },
    {
        projectId: "ccc33333-3333-3333-3333-333333333333",
        arabicName: "مشروع المنطقة الصناعية",
        projectNumber: "PRJ-2025-003",
        accountName: "Industrial Zone Holdings",
        accountId: "acc-33333-3333-3333-3333-333333333333",
    },
    {
        projectId: "ddd44444-4444-4444-4444-444444444444",
        arabicName: "مشروع الطريق السريع",
        projectNumber: "PRJ-2025-004",
        accountName: "Highway Expansion Group",
        accountId: "acc-44444-4444-4444-4444-444444444444",
    },
    {
        projectId: "eee55555-5555-5555-5555-555555555555",
        arabicName: "مشروع المجمع السكني",
        projectNumber: "PRJ-2025-005",
        accountName: "Residential Complex WLL",
        accountId: "acc-55555-5555-5555-5555-555555555555",
    },
];

// =====================================================================
// HELPER CLASS
// =====================================================================

export class ProjectHelpers {
    private static xrm: Xrm.XrmStatic = (window.parent as any).Xrm || (window as any).Xrm;

    /** Expose the default radius so the component can use it */
    static get defaultSearchRadius(): number {
        return DEFAULT_SEARCH_RADIUS;
    }

    // =====================================================================
    // API: SEARCH NEARBY PROJECTS
    // =====================================================================

    /**
     * Call the Logic App API to find projects near the given coordinates.
     *
     * When `PROJECT_SEARCH_API_URL` is empty, returns dummy data.
     *
     * @param params - Search parameters (radius, accountType, lon, lat, departmentId)
     * @returns Array of nearby project references (projectId + coordinates)
     */
    static async searchNearbyProjects(
        params: INearbyProjectRequest
    ): Promise<INearbyProjectResponse[]> {
        // If no API URL configured, return dummy data for development
        if (!PROJECT_SEARCH_API_URL) {
            console.warn("[ProjectHelpers] No API URL configured — returning dummy data");
            return Promise.resolve(DUMMY_NEARBY_PROJECTS);
        }

        try {
            const response = await fetch(PROJECT_SEARCH_API_URL, {
                method: "POST",
                headers: {
                    "Content-Type": "application/json",
                },
                body: JSON.stringify({
                    radius: params.radius,
                    accountType: params.accountType,
                    longitude: params.longitude,
                    latitude: params.latitude,
                    departmentId: params.departmentId,
                }),
            });

            if (!response.ok) {
                throw new Error(`API returned status ${response.status}: ${response.statusText}`);
            }

            const data: INearbyProjectResponse[] = await response.json();
            console.log("[ProjectHelpers] Nearby projects from API:", data.length);
            return data;
        } catch (error: any) {
            console.error("[ProjectHelpers] Error searching nearby projects:", error);
            throw error;
        }
    }

    // =====================================================================
    // CRM: GET PROJECT DETAILS
    // =====================================================================

    /**
     * Retrieve detailed project information from CRM for a list of project IDs.
     * Fetches duc_arabicname, duc_id, and the related account (via account.duc_project lookup).
     *
     * When `PROJECT_SEARCH_API_URL` is empty, returns dummy project details.
     *
     * @param projectIds - Array of duc_project GUIDs returned by the API
     * @returns Array of enriched project details for display
     */
    static async getProjectDetails(
        projectIds: string[]
    ): Promise<IProjectDetail[]> {
        // If no API URL configured, return dummy details
        if (!PROJECT_SEARCH_API_URL) {
            console.warn("[ProjectHelpers] No API URL configured — returning dummy project details");
            return Promise.resolve(DUMMY_PROJECT_DETAILS);
        }

        try {
            const details: IProjectDetail[] = [];

            for (const projectId of projectIds) {
                try {
                    // Step 1: Get project record
                    const projectRecord = await this.xrm.WebApi.retrieveRecord(
                        "duc_project",
                        projectId,
                        "?$select=duc_arabicname,duc_id,duc_name"
                    );

                    // Step 2: Find account linked to this project
                    const accountResults = await this.xrm.WebApi.retrieveMultipleRecords(
                        "account",
                        `?$select=accountid,name&$filter=_duc_project_value eq '${projectId}'&$top=1`
                    );

                    const account = accountResults?.entities?.length > 0
                        ? accountResults.entities[0]
                        : null;

                    details.push({
                        projectId: projectId,
                        arabicName: projectRecord.duc_arabicname || projectRecord.duc_name || "",
                        projectNumber: projectRecord.duc_id || "",
                        accountName: account?.name || "",
                        accountId: account?.accountid || "",
                    });
                } catch (innerError: any) {
                    console.error(`[ProjectHelpers] Error fetching details for project ${projectId}:`, innerError);
                    // Skip this project but continue with others
                }
            }

            console.log("[ProjectHelpers] Project details retrieved:", details.length);
            return details;
        } catch (error: any) {
            console.error("[ProjectHelpers] Error getting project details:", error);
            throw error;
        }
    }

    // =====================================================================
    // CRM: GET ACCOUNT FOR PROJECT
    // =====================================================================

    /**
     * Retrieve the account record linked to a specific project.
     * Looks up the account entity where duc_project lookup = projectId.
     *
     * When `PROJECT_SEARCH_API_URL` is empty, returns dummy account data.
     *
     * @param projectId - The duc_project GUID
     * @returns Account info { accountId, accountName } or null
     */
    static async getAccountForProject(
        projectId: string
    ): Promise<{ accountId: string; accountName: string } | null> {
        // If no API URL configured, return dummy account
        if (!PROJECT_SEARCH_API_URL) {
            console.warn("[ProjectHelpers] No API URL configured — returning dummy account");
            const dummyProject = DUMMY_PROJECT_DETAILS.find(p => p.projectId === projectId);
            if (dummyProject) {
                return {
                    accountId: dummyProject.accountId,
                    accountName: dummyProject.accountName,
                };
            }
            return null;
        }

        try {
            const results = await this.xrm.WebApi.retrieveMultipleRecords(
                "account",
                `?$select=accountid,name&$filter=_duc_project_value eq '${projectId}'&$top=1`
            );

            if (results?.entities?.length > 0) {
                const account = results.entities[0];
                return {
                    accountId: account.accountid,
                    accountName: account.name || "",
                };
            }

            console.warn(`[ProjectHelpers] No account found for project ${projectId}`);
            return null;
        } catch (error: any) {
            console.error("[ProjectHelpers] Error getting account for project:", error);
            throw error;
        }
    }
}
