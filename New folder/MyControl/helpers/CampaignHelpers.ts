/* eslint-disable */
/**
 * CampaignHelpers.ts
 * Helper functions for Campaign operations
 */

export class CampaignHelpers {
    private static get xrm(): Xrm.XrmStatic {
        return (window as any).Xrm;
    }

    /**
     * Get campaigns filtered by organization unit and status
     */
    static async getCampaignsByOrgUnit(
        orgUnitId: string,
        options?: {
            status?: number; // duc_campaignstatus
            campaignType?: number; // duc_campaigntype
            internalType?: number; // duc_campaigninternaltype
            ownerId?: string; // filter by owner
            includeActivePatrols?: boolean; // for patrol campaigns
        }
    ): Promise<Array<{ id: string; name: string; status?: number; type?: number }>> {
        try {
            const today = new Date().toISOString().split('T')[0];
            let filter = `_duc_organizationalunitid_value eq '${orgUnitId}'`;

            if (options?.status !== undefined) {
                filter += ` and duc_campaignstatus eq ${options.status}`;
            }

            if (options?.campaignType !== undefined) {
                filter += ` and duc_campaigntype eq ${options.campaignType}`;
            }

            if (options?.internalType !== undefined) {
                filter += ` and duc_campaigninternaltype eq ${options.internalType}`;
            }

            if (options?.ownerId) {
                filter += ` and _ownerid_value eq '${options.ownerId}'`;
            }

            if (options?.includeActivePatrols) {
                filter += ` and duc_fromdate le ${today} and duc_todate ge ${today}`;
            }

            const results = await this.xrm.WebApi.retrieveMultipleRecords(
                'new_inspectioncampaign',
                `?$select=new_inspectioncampaignid,new_name,duc_campaignstatus,duc_campaigntype&$filter=${filter}&$orderby=new_name`
            );

            const entities = results?.entities || [];

            return entities.map((e: any) => ({
                id: e.new_inspectioncampaignid,
                name: e.new_name,
                status: e.duc_campaignstatus,
                type: e.duc_campaigntype
            }));
        } catch (error: any) {
            console.error('Error getting campaigns:', error);
            return [];
        }
    }

    /**
     * Get active inspection campaigns (status = 2, type = 100000000)
     */
    static async getActiveInspectionCampaigns(orgUnitId: string): Promise<Array<{ id: string; name: string }>> {
        return await this.getCampaignsByOrgUnit(orgUnitId, {
            status: 2, // In Progress
            campaignType: 100000000 // Inspection
        });
    }

    /**
     * Get patrol campaigns for user
     */
    static async getPatrolCampaigns(
        userId: string,
        orgUnitId: string,
        status?: number
    ): Promise<Array<{ id: string; name: string }>> {
        return await this.getCampaignsByOrgUnit(orgUnitId, {
            status: status,
            campaignType: 100000004, // Patrol type
            internalType: 100000004, // Patrol internal type
            ownerId: userId,
            includeActivePatrols: true
        });
    }

    /**
     * Check if campaign is for Natural Reserve section
     * Note: Uses separate retrieve (no $expand) for offline compatibility.
     */
    static async isNaturalReserveCampaign(campaignId: string): Promise<boolean> {
        try {
            const campaign = await this.xrm.WebApi.retrieveRecord(
                'new_inspectioncampaign',
                campaignId,
                '?$select=_duc_organizationalunitid_value'
            );

            const orgUnitId = campaign['_duc_organizationalunitid_value'];
            if (!orgUnitId) return false;

            // Fetch org unit name separately (no $expand for offline compat)
            const orgUnit = await this.xrm.WebApi.retrieveRecord(
                'msdyn_organizationalunit',
                orgUnitId,
                '?$select=duc_englishname'
            );

            const orgUnitName = orgUnit?.duc_englishname || '';
            return orgUnitName === 'Inspection Section – Natural Reserves';
        } catch (error: any) {
            console.error('Error checking if campaign is Natural Reserve:', error);
            return false;
        }
    }

    /**
     * Update campaign status
     */
    static async updateCampaignStatus(campaignId: string, status: number): Promise<boolean> {
        try {
            await this.xrm.WebApi.updateRecord(
                'new_inspectioncampaign',
                campaignId,
                { duc_campaignstatus: status }
            );
            return true;
        } catch (error: any) {
            console.error('Error updating campaign status:', error);
            return false;
        }
    }

    /**
     * Start patrol campaign (update status to 2 - In Progress)
     */
    static async startPatrol(campaignId: string): Promise<boolean> {
        return await this.updateCampaignStatus(campaignId, 2);
    }

    /**
     * End patrol campaign (update status to 100000004 - Completed)
     */
    static async endPatrol(campaignId: string): Promise<boolean> {
        return await this.updateCampaignStatus(campaignId, 100000004);
    }
}
