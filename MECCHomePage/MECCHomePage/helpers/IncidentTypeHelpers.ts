/* eslint-disable */
/**
 * IncidentTypeHelpers.ts
 * Helper functions for Incident Type operations
 */

export class IncidentTypeHelpers {
    private static xrm: Xrm.XrmStatic = (window.parent as any).Xrm || (window as any).Xrm;

    /**
     * Get incident types filtered by organization unit
     */
    static async getIncidentTypesByOrgUnit(orgUnitId: string): Promise<Array<{ id: string; name: string }>> {
        try {
            const results = await this.xrm.WebApi.retrieveMultipleRecords(
                'msdyn_incidenttype',
                `?$select=msdyn_incidenttypeid,msdyn_name&$filter=_duc_organizationalunitid_value eq '${orgUnitId}'&$orderby=msdyn_name`
            );

            return (results?.entities || []).map((e: any) => ({
                id: e.msdyn_incidenttypeid,
                name: e.msdyn_name
            }));
        } catch (error: any) {
            console.error('Error getting incident types:', error);
            // alert('Error getting incident types: ' + (error?.message || error));
            return [];
        }
    }

    /**
     * Get incident types by organization unit using M:M relationship
     */
    static async getIncidentTypesByOrgUnitManyToMany(userId: string): Promise<Array<{ id: string; name: string }>> {
        try {
            // This uses the many-to-many relationship between incident types and org units
            const fetchXml = `<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                <entity name='msdyn_incidenttype'>
                    <attribute name='msdyn_incidenttypeid' />
                    <attribute name='msdyn_name' />
                    <order attribute='msdyn_name' descending='false' />
                    <link-entity name='duc_msdyn_incidenttype_msdyn_organizational' from='msdyn_incidenttypeid' to='msdyn_incidenttypeid' visible='false' intersect='true'>
                        <link-entity name='msdyn_organizationalunit' from='msdyn_organizationalunitid' to='msdyn_organizationalunitid' alias='ae'>
                            <link-entity name='systemuser' from='duc_organizationalunitid' to='msdyn_organizationalunitid' link-type='inner' alias='af'>
                                <filter type='and'>
                                    <condition attribute='systemuserid' operator='eq' uitype='systemuser' value='${userId}' />
                                </filter>
                            </link-entity>
                        </link-entity>
                    </link-entity>
                </entity>
            </fetch>`;

            const globalContext = this.xrm.Utility.getGlobalContext();
            const clientUrl = globalContext.getClientUrl();
            const url = `${clientUrl}/api/data/v9.2/msdyn_incidenttypes?fetchXml=${encodeURIComponent(fetchXml)}`;

            const response = await fetch(url, {
                method: 'GET',
                headers: {
                    'Accept': 'application/json',
                    'Content-Type': 'application/json',
                }
            });

            const data = await response.json();

            return (data.value || []).map((e: any) => ({
                id: e.msdyn_incidenttypeid,
                name: e.msdyn_name
            }));
        } catch (error: any) {
            console.error('Error getting incident types (M:M):', error);
            // alert('Error getting incident types (M:M): ' + (error?.message || error));
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
            // alert('Error getting default work order type: ' + (error?.message || error));
            return null;
        }
    }

    /**
     * Get organization unit from incident type
     * @deprecated Use WorkOrderHelpers.getIncidentTypeData() for combined queries
     * that fetch both department and work order type in a single API call.
     */
    static async getOrgUnitFromIncidentType(incidentTypeId: string): Promise<{ id: string; name: string } | null> {
        try {
            const results = await this.xrm.WebApi.retrieveMultipleRecords(
                'msdyn_incidenttype',
                `?$select=_duc_organizationalunitid_value` +
                `&$expand=duc_organizationalunitid($select=msdyn_organizationalunitid,msdyn_name)` +
                `&$filter=msdyn_incidenttypeid eq ${incidentTypeId}`
            );

            if (results?.entities?.length > 0) {
                const incidentTypeRecord = results.entities[0];
                const orgUnitLookup = incidentTypeRecord["duc_organizationalunitid"];

                if (orgUnitLookup) {
                    return {
                        id: orgUnitLookup.msdyn_organizationalunitid,
                        name: orgUnitLookup.msdyn_name
                    };
                }
            }

            return null;
        } catch (error: any) {
            console.error('Error getting org unit from incident type:', error);
            // alert('Error getting org unit from incident type: ' + (error?.message || error));
            return null;
        }
    }
}
