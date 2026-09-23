/* eslint-disable */
/**
 * AccountProfileHelpers.ts
 * Cross-environment account resolution for the CR / CP flow.
 *
 * When a CR/CP account does not exist in this environment yet, MultiTypeInspection
 * creates a bare account and asks the MOCI API catalogue whether the same
 * identifier exists there. Every call goes through duc_CustomAction_APICatalogCA
 * and differs only by `key`:
 *
 *   1. loginAndStoreToken()  → key MOCI_LOGIN. The response carries
 *      { token }, which is written to the account's duc_token field so the
 *      catalogue plugin can read it back on the next call.
 *   2. lookupAccount()       → key MOCI_GET_BY_CR (CR) or MOCI_GET_BY_CP (CP),
 *      on the same account, i.e. with duc_token already set.
 *   3. If MOCI has the account, markAccountForSync() sets duc_syncaccount = true;
 *      a Power Automate flow then copies the MOCI data onto the account.
 *
 * Every response is the same doubly-encoded envelope:
 *   { succeeded, results: "{ Result: { response: "<SOAP envelope as JSON>" }, IsSuccess }" }
 * and the MOCI record sits at
 *   env:Envelope → env:Body → mecswcrallWSQueryResponse → ListOfMecSwCrAllData → Account
 * (ListOfMecSwCrAllData is null when MOCI has no match).
 */

const API_CATALOG_ACTION_NAME = "duc_CustomAction_APICatalogCA";

/** Authenticates against MOCI and returns { token } */
const LOGIN_KEY = "MOCI_LOGIN";

/** Identifier lookups — run after the token is stored on the account */
const LOOKUP_KEY_CR = "MOCI_GET_BY_CR";
const LOOKUP_KEY_CP = "MOCI_GET_BY_CP";

/** Account field the MOCI token is stored in before the lookup call */
const TOKEN_FIELD = "duc_token";

/** Setting this to true triggers the Power Automate flow that syncs the MOCI data */
const SYNC_FIELD = "duc_syncaccount";

/** duc_processautomationconfigurations switch: "1" calls the MOCI APIs, "0" skips them */
const API_CALLING_CONFIG_KEY = "enableAPIcallingHomepage";

/** Which identifier the user typed — picks the lookup key */
export type IdentifierKind = "cr" | "cp";

/** Normalized view of the API catalogue action's response */
export interface IApiCatalogResult {
    /** false when the nested payload could not be unwrapped (unexpected shape) */
    parsed: boolean;
    isSuccess: boolean;
    /** true when MOCI returned an account for the identifier */
    found: boolean;
    /** The MOCI account's Name, shown on the confirmation popup */
    accountName: string | null;
    message: string | null;
    /** The unwrapped `Result.response` body, for diagnostics */
    payload: any;
}

export class AccountProfileHelpers {
    private static xrm: Xrm.XrmStatic = (window.parent as any).Xrm || (window as any).Xrm;

    // =====================================================================
    // CONFIGURATION
    // =====================================================================

    /**
     * Whether the MOCI APIs should be called, from duc_processautomationconfigurations
     * (duc_key = "enableAPIcallingHomepage"). Only a value of "1" enables them;
     * a missing row, any other value or a read error keeps them off.
     */
    static async isApiCallingEnabled(): Promise<boolean> {
        try {
            const results = await this.xrm.WebApi.retrieveMultipleRecords(
                "duc_processautomationconfigurations",
                `?$select=duc_value&$filter=duc_key eq '${API_CALLING_CONFIG_KEY}'&$top=1`,
            );

            const value = results?.entities?.[0]?.duc_value;
            if (value === undefined) {
                console.warn(
                    `[AccountProfileHelpers] No '${API_CALLING_CONFIG_KEY}' configuration found; MOCI APIs are off`,
                );
            }
            return String(value ?? "").trim() === "1";
        } catch (error: any) {
            console.error(`[AccountProfileHelpers] Error reading '${API_CALLING_CONFIG_KEY}':`, error);
            return false;
        }
    }

    // =====================================================================
    // API CATALOGUE CUSTOM ACTION
    // =====================================================================

    /**
     * Call duc_CustomAction_APICatalogCA for the given account with `key`.
     * Returns the raw response, or null when the action could not be called.
     */
    static async callApiCatalogAction(accountId: string, key: string): Promise<any | null> {
        try {
            const request = {
                key: key,
                EntityReference_Id: accountId,
                EntityReference_EntityName: "account",
                getMetadata: () => ({
                    boundParameter: null,
                    parameterTypes: {
                        key: { typeName: "Edm.String", structuralProperty: 1 },
                        EntityReference_Id: { typeName: "Edm.String", structuralProperty: 1 },
                        EntityReference_EntityName: { typeName: "Edm.String", structuralProperty: 1 },
                    },
                    operationType: 0,
                    operationName: API_CATALOG_ACTION_NAME,
                }),
            };

            const result = await this.xrm.WebApi.online.execute(request as any);
            if (!result?.ok) {
                console.warn(`[AccountProfileHelpers] '${key}' returned non-OK status:`, result?.status);
                return null;
            }

            const response = await result.json().catch(() => null);
            console.log(`[AccountProfileHelpers] '${key}' response:`, response);
            return response;
        } catch (error: any) {
            console.error(`[AccountProfileHelpers] Error calling '${key}':`, error);
            return null;
        }
    }

    private static parseJson(value: any): any {
        if (typeof value !== "string") return value ?? null;
        try {
            return JSON.parse(value);
        } catch {
            return null;
        }
    }

    /** Unwrap results → Result.response, the body every key answers with */
    private static unwrapPayload(actionResponse: any): { outer: any; payload: any } {
        const outer = this.parseJson(actionResponse?.results);
        return { outer, payload: this.parseJson(outer?.Result?.response) };
    }

    // =====================================================================
    // STEP 1: LOGIN — fetch the MOCI token and store it on the account
    // =====================================================================

    /**
     * Call MOCI_LOGIN and write the returned token to the account's duc_token
     * field, so the lookup call that follows runs authenticated.
     * Returns the token, or null when login failed.
     */
    static async loginAndStoreToken(accountId: string): Promise<string | null> {
        const actionResponse = await this.callApiCatalogAction(accountId, LOGIN_KEY);
        const { payload } = this.unwrapPayload(actionResponse);
        const token = payload?.token || null;

        if (!token) {
            console.warn("[AccountProfileHelpers] MOCI_LOGIN returned no token:", actionResponse);
            return null;
        }

        await this.xrm.WebApi.updateRecord("account", accountId, { [TOKEN_FIELD]: token });
        return token;
    }

    // =====================================================================
    // STEP 2: LOOKUP — does this identifier exist on the MOCI side?
    // =====================================================================

    /** The MOCI Account record from the SOAP envelope, or null when there is none */
    private static getMociAccount(payload: any): any | null {
        const data = payload?.["env:Envelope"]?.["env:Body"]?.mecswcrallWSQueryResponse
            ?.ListOfMecSwCrAllData;
        const account = data?.Account;
        if (Array.isArray(account)) return account[0] ?? null;
        return account ?? null;
    }

    /** Unwrap a lookup response into the shape the CR/CP flow consumes. */
    static parseApiCatalogResponse(actionResponse: any): IApiCatalogResult {
        const empty: IApiCatalogResult = {
            parsed: false,
            isSuccess: actionResponse?.succeeded === true,
            found: false,
            accountName: null,
            message: null,
            payload: null,
        };

        if (!actionResponse) return empty;

        const { outer, payload } = this.unwrapPayload(actionResponse);

        // An empty `response` is a "not found" answer, not a parse failure —
        // only an unreadable envelope is.
        if (payload === null && outer?.Result?.response !== "") {
            console.warn("[AccountProfileHelpers] Could not unwrap the action payload:", actionResponse);
            return { ...empty, message: outer?.ErrorMessage || null };
        }

        const mociAccount = this.getMociAccount(payload);

        return {
            parsed: true,
            isSuccess: actionResponse.succeeded === true && outer?.IsSuccess === true,
            found: mociAccount !== null,
            accountName: mociAccount?.Name || mociAccount?.OrgNameEnu || null,
            message: outer?.ErrorMessage || null,
            payload: payload,
        };
    }

    /**
     * Full check for one account: log in, store the token, then ask the
     * catalogue whether the CR/CP identifier exists on its side.
     */
    static async lookupAccount(
        accountId: string,
        identifierKind: IdentifierKind,
    ): Promise<IApiCatalogResult> {
        const token = await this.loginAndStoreToken(accountId);
        if (!token) {
            return this.parseApiCatalogResponse(null);
        }

        const key = identifierKind === "cp" ? LOOKUP_KEY_CP : LOOKUP_KEY_CR;
        const actionResponse = await this.callApiCatalogAction(accountId, key);
        return this.parseApiCatalogResponse(actionResponse);
    }

    // =====================================================================
    // STEP 3: SYNC — let Power Automate copy the MOCI data onto the account
    // =====================================================================

    /** Set duc_syncaccount = true, which triggers the MOCI sync flow for the account. */
    static async markAccountForSync(accountId: string): Promise<void> {
        await this.xrm.WebApi.updateRecord("account", accountId, { [SYNC_FIELD]: true });
    }
}
