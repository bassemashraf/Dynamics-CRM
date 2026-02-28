/* eslint-disable */
/**
 * IncidentTypeHelpers.ts
 * Helper functions for Incident Type operations
 */

export class IncidentTypeHelpers {
    private static get xrm(): Xrm.XrmStatic {
        return (window as any).Xrm;
    }

    /**
     * Get incident types filtered by organization unit
     */
    static async getIncidentTypesByOrgUnit(orgUnitId: string): Promise<Array<{ id: string; name: string }>> {
        try {
            const results = await this.xrm.WebApi.retrieveMultipleRecords(
                'msdyn_incidenttype',
                `?$select=msdyn_incidenttypeid,msdyn_name&$filter=_duc_organizationalunitid_value eq '${orgUnitId}'&$orderby=msdyn_name`
            );

            const entities = results?.entities || [];

            return entities.map((e: any) => ({
                id: e.msdyn_incidenttypeid,
                name: e.msdyn_name
            }));
        } catch (error: any) {
            console.error('Error getting incident types:', error);
            return [];
        }
    }

    /**
     * Get default work order type from incident type
     */
    static async getDefaultWorkOrderType(incidentTypeId: string): Promise<{ id: string; name: string; entityType: string } | null> {
        try {
            const result = await this.xrm.WebApi.retrieveRecord(
                'msdyn_incidenttype',
                incidentTypeId,
                '?$select=_msdyn_defaultworkordertype_value'
            );

            const workOrderTypeId = result["_msdyn_defaultworkordertype_value"];
            if (!workOrderTypeId) return null;

            const workOrderTypeName = result["_msdyn_defaultworkordertype_value@OData.Community.Display.V1.FormattedValue"];
            const entityType = result["_msdyn_defaultworkordertype_value@Microsoft.Dynamics.CRM.lookuplogicalname"];


            return {
                id: workOrderTypeId,
                name: workOrderTypeName,
                entityType: entityType
            };
        } catch (error: any) {
            console.error('Error getting default work order type:', error);
            return null;
        }
    }

    /**
     * Get organization unit from incident type
     * Note: $expand not used — uses lookup formatted values for offline compatibility.
     */
    static async getOrgUnitFromIncidentType(incidentTypeId: string): Promise<{ id: string; name: string } | null> {
        try {
            const result = await this.xrm.WebApi.retrieveRecord(
                'msdyn_incidenttype',
                incidentTypeId,
                '?$select=_duc_organizationalunitid_value'
            );

            const orgUnitId = result["_duc_organizationalunitid_value"];
            const orgUnitName = result["_duc_organizationalunitid_value@OData.Community.Display.V1.FormattedValue"];


            if (orgUnitId) {
                return {
                    id: orgUnitId,
                    name: orgUnitName
                };
            }

            return null;
        } catch (error: any) {
            console.error('Error getting org unit from incident type:', error);
            return null;
        }
    }
}
