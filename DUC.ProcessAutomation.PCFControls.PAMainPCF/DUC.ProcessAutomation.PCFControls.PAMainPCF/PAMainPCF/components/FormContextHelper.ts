export class FormContextHelper {
    private static cachedProcessExtension: any = null;

    public static getFormContext(): any {
        let parentXrm: any = null;
        try {
            parentXrm = (window.parent as any).Xrm;
        } catch (e) {
            console.warn("Could not access window.parent.Xrm due to cross-origin or offline restrictions.");
        }
        const xrm = parentXrm || (window as any).Xrm;
        return xrm ? xrm.Page : null;
    }

    public static async initAsync(xrmGlobal: any): Promise<void> {
        try {
            const ctx = this.getFormContext();
            if (!ctx) return;
            
            // Check if we are on a work order form (or if doc_processextension is available)
            const processExtAttr = ctx.getAttribute("duc_processextension");
            if (!processExtAttr) return;

            const val = processExtAttr.getValue();
            if (!val || val.length === 0) return;

            const processExtId = val[0].id.replace(/[{}]/g, "");
            console.log(`[FormContextHelper] Pre-fetching Process Extension: ${processExtId}`);

            const selectQuery = "?$select=_duc_processdefinition_value,_duc_currentstage_value,_ownerid_value,_duc_owningteam_value,duc_stepname,_duc_surveyresponse_value,regardingobjectid,_regardingobjectid_value";
            
            const record = await xrmGlobal.WebApi.retrieveRecord("duc_processextension", processExtId, selectQuery);
            this.cachedProcessExtension = record;
            console.log("[FormContextHelper] Cached Process Extension", this.cachedProcessExtension);
        } catch (e) {
            console.error("[FormContextHelper] Failed to initAsync", e);
        }
    }

    public static getLookupValue(fieldName: string): { id: string, name: string, entityType: string } | null {
        // If cached extension is available and it's one of the known offline fields, serve from cache
        if (this.cachedProcessExtension && fieldName !== "duc_processextension") {
            const odataValueField = `_${fieldName}_value`;
            if (this.cachedProcessExtension[odataValueField]) {
                return {
                    id: this.cachedProcessExtension[odataValueField],
                    name: this.cachedProcessExtension[`${odataValueField}@OData.Community.Display.V1.FormattedValue`] || "",
                    entityType: this.cachedProcessExtension[`${odataValueField}@Microsoft.Dynamics.CRM.lookuplogicalname`] || ""
                };
            }
        }

        const ctx = this.getFormContext();
        if (ctx && ctx.getAttribute) {
            const attr = ctx.getAttribute(fieldName);
            if (attr) {
                const val = attr.getValue();
                if (val && val.length > 0) {
                    return {
                        id: val[0].id,
                        name: val[0].name,
                        entityType: val[0].entityType || val[0].typename
                    };
                }
            }
        }
        return null;
    }

    public static setLookupValue(fieldName: string, id: string, name: string, entityType: string): void {
        const ctx = this.getFormContext();
        if (ctx && ctx.getAttribute) {
            const attr = ctx.getAttribute(fieldName);
            if (attr) {
                attr.setValue([{
                    id: id,
                    name: name,
                    entityType: entityType
                }]);
            } else {
                console.warn(`[FormContextHelper] Field '${fieldName}' not found on the form to set.`);
            }
        }
    }

    public static getStringValue(fieldName: string): string | null {
        // Serve from cache if available
        if (this.cachedProcessExtension && fieldName !== "duc_processextension") {
            if (fieldName === "activityid") {
                return this.cachedProcessExtension["activityid"] || this.cachedProcessExtension["duc_processextensionid"] || null;
            }
            if (this.cachedProcessExtension[fieldName] !== undefined) {
                return this.cachedProcessExtension[fieldName];
            }
        }

        const ctx = this.getFormContext();
        if (ctx && ctx.getAttribute) {
            const attr = ctx.getAttribute(fieldName);
            if (attr) {
                return attr.getValue();
            }
        }
        return null;
    }

    public static setStringValue(fieldName: string, value: string): void {
        const ctx = this.getFormContext();
        if (ctx && ctx.getAttribute) {
            const attr = ctx.getAttribute(fieldName);
            if (attr) {
                attr.setValue(value);
            } else {
                console.warn(`[FormContextHelper] Field '${fieldName}' not found on the form to set.`);
            }
        }
    }
}
