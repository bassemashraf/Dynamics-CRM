/* eslint-disable */
/**
 * WorkOrderHelpers.ts
 * Reusable helper functions for Work Order operations
 * Supports both online and offline modes using Xrm.WebApi
 */

export class WorkOrderHelpers {
    private static get xrm(): Xrm.XrmStatic {
        return (window as any).Xrm;
    }

    // =====================================================================
    // DEPARTMENT OPERATIONS
    // =====================================================================

    /**
     * Get user's organization unit and set it on work order department field
     */
    static async setDepartmentFromUser(userId: string): Promise<{ id: string; name: string } | null> {
        try {
            const userResult = await this.xrm.WebApi.retrieveRecord(
                "systemuser",
                userId,
                "?$select=_duc_organizationalunitid_value"
            );

            const orgUnitId = userResult._duc_organizationalunitid_value;
            if (!orgUnitId) return null;

            const orgUnitName = userResult["_duc_organizationalunitid_value@OData.Community.Display.V1.FormattedValue"];

            alert(`User Organization Unit: ${orgUnitName}`);

            return {
                id: orgUnitId,
                name: orgUnitName
            };
        } catch (error: any) {
            console.error("Error getting department from user:", error);
            alert("Error getting department from user: " + (error?.message || error) + (error?.innerError ? "\nInner: " + JSON.stringify(error.innerError) : ""));
            return null;
        }
    }

    /**
     * Set department from incident type's organizational unit
     */
    static async setDepartmentFromIncidentType(incidentTypeId: string): Promise<{ id: string; name: string } | null> {
        try {
            const results = await this.xrm.WebApi.retrieveMultipleRecords(
                "msdyn_incidenttype",
                `?$select=_duc_organizationalunitid_value` +
                `&$filter=msdyn_incidenttypeid eq ${incidentTypeId}`
            );

            if (results?.entities?.length > 0) {
                const incidentTypeRecord = results.entities[0];
                const orgUnitId = incidentTypeRecord["_duc_organizationalunitid_value"];
                const orgUnitName = incidentTypeRecord["_duc_organizationalunitid_value@OData.Community.Display.V1.FormattedValue"];

                alert(`Department found for Incident Type: ${orgUnitName}`);

                if (orgUnitId) {
                    return {
                        id: orgUnitId,
                        name: orgUnitName
                    };
                }
            }

            return null;
        } catch (error: any) {
            console.error("Error setting department from incident type:", error);
            alert("Error setting department from incident type: " + (error?.message || error) + (error?.innerError ? "\nInner: " + JSON.stringify(error.innerError) : ""));
            return null;
        }
    }

    // =====================================================================
    // INCIDENT TYPE OPERATIONS
    // =====================================================================

    /**
     * Set work order type based on incident type's default work order type
     */
    static async setWorkOrderTypeFromIncidentType(incidentTypeId: string): Promise<{ id: string; name: string; entityType: string } | null> {
        try {
            const result = await this.xrm.WebApi.retrieveRecord(
                "msdyn_incidenttype",
                incidentTypeId,
                "?$select=_msdyn_defaultworkordertype_value"
            );

            const workOrderTypeId = result["_msdyn_defaultworkordertype_value"];
            if (!workOrderTypeId) return null;

            const workOrderTypeName = result["_msdyn_defaultworkordertype_value@OData.Community.Display.V1.FormattedValue"];
            const entityType = result["_msdyn_defaultworkordertype_value@Microsoft.Dynamics.CRM.lookuplogicalname"];

            alert(`Default Work Order Type: ${workOrderTypeName}`);

            return {
                id: workOrderTypeId,
                name: workOrderTypeName,
                entityType: entityType
            };
        } catch (error: any) {
            console.error("Error setting work order type from incident type:", error);
            alert("Error setting work order type from incident type: " + (error?.message || error) + (error?.innerError ? "\nInner: " + JSON.stringify(error.innerError) : ""));
            return null;
        }
    }

    /**
     * OPTIMIZATION: Get both department and work order type from incident type in a single call
     * Note: $expand is NOT used here for offline compatibility — splits into two safe selects.
     * department is retrieved via lookup formatted value from the incident type record.
     */
    static async getIncidentTypeData(incidentTypeId: string): Promise<{
        department?: { id: string; name: string };
        workOrderType?: { id: string; name: string; entityType: string };
    } | null> {
        try {
            const result = await this.xrm.WebApi.retrieveRecord(
                "msdyn_incidenttype",
                incidentTypeId,
                "?$select=_duc_organizationalunitid_value,_msdyn_defaultworkordertype_value"
            );

            const response: {
                department?: { id: string; name: string };
                workOrderType?: { id: string; name: string; entityType: string };
            } = {};

            // Extract department from organizational unit lookup values
            const orgUnitId = result["_duc_organizationalunitid_value"];
            if (orgUnitId) {
                response.department = {
                    id: orgUnitId,
                    name: result["_duc_organizationalunitid_value@OData.Community.Display.V1.FormattedValue"] || ""
                };
            }

            // Extract work order type
            const workOrderTypeId = result["_msdyn_defaultworkordertype_value"];
            if (workOrderTypeId) {
                const workOrderTypeName = result["_msdyn_defaultworkordertype_value@OData.Community.Display.V1.FormattedValue"];
                response.workOrderType = {
                    id: workOrderTypeId,
                    name: workOrderTypeName,
                    entityType: result["_msdyn_defaultworkordertype_value@Microsoft.Dynamics.CRM.lookuplogicalname"]
                };
                alert(`Retrieved Incident Data: Dept: ${response.department?.name}, WO Type: ${workOrderTypeName}`);
            } else if (response.department) {
                alert(`Retrieved Incident Data: Dept: ${response.department.name}, WO Type: Not Found`);
            }

            return Object.keys(response).length > 0 ? response : null;
        } catch (error: any) {
            console.error("Error getting incident type data:", error);
            alert("Error getting incident type data: " + (error?.message || error) + (error?.innerError ? "\nInner: " + JSON.stringify(error.innerError) : ""));
            return null;
        }
    }

    // =====================================================================
    // CAMPAIGN OPERATIONS
    // =====================================================================

    /**
     * Get incident type from campaign and return campaign data for parent campaign field
     */
    static async getCampaignData(campaignId: string): Promise<{
        incidentType?: { id: string; name: string; entityType: string };
        parentCampaign: { id: string; name: string; entityType: string };
    } | null> {
        try {
            const result = await this.xrm.WebApi.retrieveRecord(
                "new_inspectioncampaign",
                campaignId,
                "?$select=_duc_incidenttype_value,new_name"
            );

            const incidentTypeId = result["_duc_incidenttype_value"];
            const campaignName = result["new_name"];

            const campaignData: any = {
                parentCampaign: {
                    id: campaignId,
                    name: campaignName,
                    entityType: "new_inspectioncampaign"
                }
            };

            if (incidentTypeId) {
                const incidentTypeName = result["_duc_incidenttype_value@OData.Community.Display.V1.FormattedValue"];
                const entityType = result["_duc_incidenttype_value@Microsoft.Dynamics.CRM.lookuplogicalname"];

                campaignData.incidentType = {
                    id: incidentTypeId,
                    name: incidentTypeName,
                    entityType: entityType
                };
                alert(`Campaign Data: Campaign: ${campaignName}, Incident Type: ${incidentTypeName}`);
            } else {
                alert(`Campaign Data: Campaign: ${campaignName}, Incident Type: Not Found`);
            }

            return campaignData;
        } catch (error: any) {
            console.error("Error getting campaign data:", error);
            alert("Error getting campaign data: " + (error?.message || error) + (error?.innerError ? "\nInner: " + JSON.stringify(error.innerError) : ""));
            return null;
        }
    }

    // =====================================================================
    // ACCOUNT OPERATIONS
    // =====================================================================

    /**
     * Handle sub-account change - set service account and address information
     */
    static async handleSubAccountChange(subAccountId: string): Promise<{
        serviceAccount: { id: string; name: string; entityType: string };
        address?: { id: string; name: string; entityType: string };
        latitude?: number;
        longitude?: number;
    } | null> {
        try {
            const accountResult = await this.xrm.WebApi.retrieveRecord(
                "account",
                subAccountId,
                "?$select=name,_parentaccountid_value,_duc_address_value"
            );

            let serviceAccountId: string;
            let serviceAccountName: string;

            // Check if parent account exists
            const parentAccountId = accountResult["_parentaccountid_value"];
            if (parentAccountId) {
                serviceAccountId = parentAccountId;
                serviceAccountName = accountResult["_parentaccountid_value@OData.Community.Display.V1.FormattedValue"];
            } else {
                // Use sub-account as service account if no parent
                serviceAccountId = subAccountId;
                serviceAccountName = accountResult.name;
            }

            const result: any = {
                serviceAccount: {
                    id: serviceAccountId,
                    name: serviceAccountName,
                    entityType: "account"
                }
            };

            // Handle address if exists
            const addressId = accountResult["_duc_address_value"];
            if (addressId) {
                const addressName = accountResult["_duc_address_value@OData.Community.Display.V1.FormattedValue"];

                alert(`Service Account: ${serviceAccountName}, Address: ${addressName}`);

                result.address = {
                    id: addressId,
                    name: addressName,
                    entityType: "duc_addressinformation"
                };

                // Get coordinates from address (separate call — no $expand for offline compat)
                const addressData = await this.xrm.WebApi.retrieveRecord(
                    "duc_addressinformation",
                    addressId,
                    "?$select=duc_longitude,duc_latitude"
                );

                if (addressData.duc_longitude != null) {
                    result.longitude = addressData.duc_longitude;
                }

                if (addressData.duc_latitude != null) {
                    result.latitude = addressData.duc_latitude;
                }
            }

            return result;
        } catch (error: any) {
            console.error("Error handling sub-account change:", error);
            alert("Error handling sub-account change: " + (error?.message || error) + (error?.innerError ? "\nInner: " + JSON.stringify(error.innerError) : ""));
            return null;
        }
    }

    /**
     * Get coordinates from address lookup
     */
    static async getCoordinatesFromAddress(addressId: string): Promise<{ latitude: number; longitude: number } | null> {
        try {
            const addressRecord = await this.xrm.WebApi.retrieveRecord(
                "duc_addressinformation",
                addressId,
                "?$select=duc_latitude,duc_longitude"
            );

            if (addressRecord.duc_latitude != null && addressRecord.duc_longitude != null) {
                return {
                    latitude: addressRecord.duc_latitude,
                    longitude: addressRecord.duc_longitude
                };
            }

            return null;
        } catch (error: any) {
            console.error("Error getting coordinates from address:", error);
            alert("Error getting coordinates from address: " + (error?.message || error) + (error?.innerError ? "\nInner: " + JSON.stringify(error.innerError) : ""));
            return null;
        }
    }

    // =====================================================================
    // BOOKING OPERATIONS
    // =====================================================================

    /**
     * Create automatic booking for work order (for mobile creation).
     * Uses pre-cached bookableResourceId and bookingStatusId from InitCache.
     */
    static async createAutoBooking(
        workOrderId: string,
        userId: string,
        options: {
            bookableResourceId: string;
            bookingStatusId: string;
        }
    ): Promise<string | null> {
        try {
            const now = new Date();
            const end = new Date(Date.now() + 60 * 60 * 1000); // 1 hour later

            const bookingData = {
                "ownerid@odata.bind": `/systemusers(${userId})`,
                "starttime": now.toISOString(),
                "endtime": end.toISOString(),
                "duration": 1,
                "msdyn_workorder@odata.bind": `/msdyn_workorders(${workOrderId})`,
                "Resource@odata.bind": `/bookableresources(${options.bookableResourceId})`,
                "BookingStatus@odata.bind": `/bookingstatuses(${options.bookingStatusId})`
            };
            const result = await this.xrm.WebApi.createRecord("bookableresourcebooking", bookingData);
            console.log("Booking created successfully. Booking ID:", result.id);
            return result.id;
        } catch (error: any) {
            console.error("Error creating auto booking:", error);
            alert("Error creating auto booking: " + (error?.message || error) + (error?.innerError ? "\nInner: " + JSON.stringify(error.innerError) : ""));
            throw error;
        }
    }

    // =====================================================================
    // VALIDATION HELPERS
    // =====================================================================

    /**
     * Check if all required fields for work order creation are populated
     */
    static validateWorkOrderData(data: {
        serviceAccount?: any;
        incidentType?: any;
        department?: any;
    }): { isValid: boolean; missingFields: string[] } {
        const missingFields: string[] = [];

        if (!data.serviceAccount) {
            missingFields.push("Service Account");
        }

        if (!data.incidentType) {
            missingFields.push("Incident Type");
        }

        if (!data.department) {
            missingFields.push("Department");
        }

        return {
            isValid: missingFields.length === 0,
            missingFields
        };
    }

    // =====================================================================
    // MOBILE DETECTION
    // =====================================================================

    /**
     * Detect if work order is being created from mobile
     */
    static isMobileClient(): boolean {
        try {
            const context = this.xrm.Utility.getGlobalContext().client;
            const clientType = context.getClient();
            return clientType === "Mobile";
        } catch (error: any) {
            console.error("Error detecting mobile client:", error);
            alert("Error detecting mobile client: " + (error?.message || error) + (error?.innerError ? "\nInner: " + JSON.stringify(error.innerError) : ""));
            return false;
        }
    }

    // =====================================================================
    // CAMPAIGN FILTERS
    // =====================================================================

    /**
     * Get active patrol campaign for Natural Reserve org unit
     */
    static async getActivePatrolCampaign(userId: string, orgUnitId: string): Promise<{ id: string; name: string } | null> {
        try {
            const filter = `duc_campaignstatus eq 2 and duc_campaigninternaltype eq 100000004 ` +
                `and duc_campaigntype eq 100000004 and _ownerid_value eq ${userId} ` +
                `and _duc_organizationalunitid_value eq ${orgUnitId}`;

            const results = await this.xrm.WebApi.retrieveMultipleRecords(
                "new_inspectioncampaign",
                `?$select=new_inspectioncampaignid,new_name&$filter=${encodeURIComponent(filter)}`
            );

            if (results?.entities?.length > 0) {
                const campaign = results.entities[0];
                alert(`Active Patrol Campaign found: ${campaign.new_name}`);
                return {
                    id: campaign.new_inspectioncampaignid,
                    name: campaign.new_name
                };
            }

            alert("No active patrol campaign found.");
            return null;
        } catch (error: any) {
            console.error("Error getting active patrol campaign:", error);
            alert("Error getting active patrol campaign: " + (error?.message || error) + (error?.innerError ? "\nInner: " + JSON.stringify(error.innerError) : ""));
            return null;
        }
    }

    // =====================================================================
    // WORK ORDER CREATION
    // =====================================================================

    /**
     * Create work order with all required data
     */
    static async createWorkOrder(data: {
        subAccount: { id: string; name: string };
        serviceAccount: { id: string; name: string };
        incidentType: { id: string; name: string };
        department: { id: string; name: string };
        campaign?: { id: string; name: string };
        workOrderType?: { id: string; name: string };
        parentCampaign?: { id: string; name: string };
        address?: { id: string; name: string };
        latitude?: number;
        longitude?: number;
        anonymousCustomer?: boolean;
        accountInspectionType?: number;
        createdFromMobile?: boolean;
    }): Promise<string | null> {
        try {
            const workOrderData: any = {
                'duc_subaccount@odata.bind': `/accounts(${data.subAccount.id})`,
                'msdyn_serviceaccount@odata.bind': `/accounts(${data.serviceAccount.id})`,
                'msdyn_primaryincidenttype@odata.bind': `/msdyn_incidenttypes(${data.incidentType.id})`,
                'duc_Department@odata.bind': `/msdyn_organizationalunits(${data.department.id})`
            };

            // Optional fields
            if (data.campaign) {
                workOrderData['new_Campaign@odata.bind'] = `/new_inspectioncampaigns(${data.campaign.id})`;
            }

            if (data.workOrderType) {
                workOrderData['msdyn_workordertype@odata.bind'] = `/msdyn_workordertypes(${data.workOrderType.id})`;
            }

            if (data.parentCampaign) {
                workOrderData['duc_ParentCampaign@odata.bind'] = `/new_inspectioncampaigns(${data.parentCampaign.id})`;
            }

            if (data.address) {
                workOrderData['duc_Address@odata.bind'] = `/duc_addressinformations(${data.address.id})`;
            }

            if (data.latitude != null) {
                workOrderData.msdyn_latitude = data.latitude;
            }

            if (data.longitude != null) {
                workOrderData.msdyn_longitude = data.longitude;
            }

            if (data.anonymousCustomer !== undefined) {
                workOrderData.duc_anonymouscustomer = data.anonymousCustomer;
                workOrderData.duc_responsibleemployeeisnotavailable = data.anonymousCustomer;
            }

            if (data.accountInspectionType !== undefined) {
                workOrderData.duc_accountinspectiontype = data.accountInspectionType;
            }

            if (data.createdFromMobile !== undefined) {
                workOrderData.duc_createdfrommobile = data.createdFromMobile;
            }

            const result = await this.xrm.WebApi.createRecord('msdyn_workorder', workOrderData);
            console.log('Work order created successfully:', result.id);

            return result.id;
        } catch (error: any) {
            console.error('Error creating work order:', error);
            alert('Error creating work order: ' + (error?.message || error) + (error?.innerError ? '\nInner: ' + JSON.stringify(error.innerError) : ''));
            throw error;
        }
    }
}
