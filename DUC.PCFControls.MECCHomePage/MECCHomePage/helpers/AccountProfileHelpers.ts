/* eslint-disable */
/**
 * AccountProfileHelpers.ts
 * Cross-environment account resolution for the CR / CP flow.
 *
 * When a CR/CP account does not exist in this environment yet, MultiTypeInspection
 * creates a bare account and asks the MOCI API catalogue whether the same
 * identifier already exists there. Every call goes through
 * duc_CustomAction_APICatalogCA and differs only by `key`:
 *
 *   1. loginAndStoreToken()  → key MOCI_LOGIN. The response carries
 *      { token }, which is written to the account's duc_token field so the
 *      catalogue plugin can read it back on the next call.
 *   2. lookupAccount()       → key MOCI_GET_BY_CR (CR) or MOCI_GET_BY_CP (CP),
 *      on the same account, i.e. with duc_token already set.
 *   3. If the lookup found a match and the user confirms it, postAccountUpdate()
 *      reads the UpdateAccount endpoint from duc_processautomationconfigurations
 *      and POSTs the returned profiles to it.
 *
 * Every response is the same doubly-encoded envelope:
 *   { succeeded, results: "{ Result: { response: "<json or empty>" }, IsSuccess }" }
 */

const API_CATALOG_ACTION_NAME = "duc_CustomAction_APICatalogCA";

/** Authenticates against MOCI and returns { token } */
const LOGIN_KEY = "MOCI_LOGIN";

/** Identifier lookups — run after the token is stored on the account */
const LOOKUP_KEY_CR = "MOCI_GET_BY_CR";
const LOOKUP_KEY_CP = "MOCI_GET_BY_CP";

/** Account field the MOCI token is stored in before the lookup call */
const TOKEN_FIELD = "duc_token";

/** duc_processautomationconfigurations lookup for the UpdateAccount endpoint */
const UPDATE_ACCOUNT_CONFIG_KEY = "UpdateAccount";
const UPDATE_ACCOUNT_CONFIG_AREA = "API Catalogue";

/** Which identifier the user typed — picks the lookup key */
export type IdentifierKind = "cr" | "cp";

/** One profile entry — both the action's `data` items and the UpdateAccount request body */
export interface IAccountProfile {
    id: string;
    profileId: string | null;
    namear: string | null;
    nameen: string | null;
    email: string | null;
    phone: string | null;
    status: string | null;
    shortname: string | null;
}

/** Normalized view of the API catalogue action's response */
export interface IApiCatalogResult {
    /** false when the nested payload could not be unwrapped (unexpected shape) */
    parsed: boolean;
    isSuccess: boolean;
    /** true when the catalogue reported the identifier exists on its side */
    found: boolean;
    status: string | null;
    code: string | null;
    message: string | null;
    profiles: IAccountProfile[];
    /** The unwrapped `Result.response` body, for diagnostics */
    payload: any;
}

export class AccountProfileHelpers {
    private static xrm: Xrm.XrmStatic = (window.parent as any).Xrm || (window as any).Xrm;

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

    /**
     * Whether the catalogue payload represents a match.
     *
     * The not-found answer is an empty `response` string. The shape of the
     * found answer is not confirmed yet, so anything non-empty that is not an
     * explicit negative counts as a match — narrow this once the real payload
     * is known.
     */
    private static isFound(payload: any): boolean {
        if (payload === null || payload === undefined || payload === "") return false;
        if (Array.isArray(payload)) return payload.length > 0;
        if (Array.isArray(payload.data)) return payload.data.length > 0;
        if (typeof payload.found === "boolean") return payload.found;
        if (typeof payload.exists === "boolean") return payload.exists;
        return true;
    }

    private static normalizeProfile(item: any): IAccountProfile {
        return {
            id: item?.id ?? "",
            profileId: item?.profileId ?? null,
            namear: item?.namear ?? null,
            nameen: item?.nameen ?? null,
            email: item?.email ?? null,
            phone: item?.phone ?? null,
            status: item?.status ?? null,
            shortname: item?.shortname ?? null,
        };
    }

    /** Unwrap a lookup response into the shape the CR/CP flow consumes. */
    static parseApiCatalogResponse(actionResponse: any): IApiCatalogResult {
        const empty: IApiCatalogResult = {
            parsed: false,
            isSuccess: actionResponse?.succeeded === true,
            found: false,
            status: null,
            code: null,
            message: null,
            profiles: [],
            payload: null,
        };

        if (!actionResponse) return empty;

        const { outer, payload } = this.unwrapPayload(actionResponse);

        // An empty `response` is the catalogue's "not found" answer, not a
        // parse failure — only an unreadable envelope is.
        if (payload === null && outer?.Result?.response !== "") {
            console.warn("[AccountProfileHelpers] Could not unwrap the action payload:", actionResponse);
            return { ...empty, message: outer?.ErrorMessage || null };
        }

        const data = Array.isArray(payload?.data) ? payload.data : [];

        return {
            parsed: true,
            isSuccess: actionResponse.succeeded === true && outer?.IsSuccess === true,
            found: this.isFound(payload),
            status: payload?.status ?? null,
            code: payload?.code != null ? String(payload.code) : null,
            message: payload?.message ?? outer?.ErrorMessage ?? null,
            profiles: data.map((item: any) => this.normalizeProfile(item)),
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
    // STEP 3: UPDATE ACCOUNT API
    // =====================================================================

    /**
     * Read the UpdateAccount endpoint from duc_processautomationconfigurations
     * (duc_key = "UpdateAccount", duc_area = "API Catalogue").
     */
    static async getUpdateAccountApiUrl(): Promise<string | null> {
        try {
            const query =
                `?$select=duc_value&$filter=duc_key eq '${UPDATE_ACCOUNT_CONFIG_KEY}'` +
                ` and duc_area eq '${UPDATE_ACCOUNT_CONFIG_AREA}'&$top=1`;

            const results = await this.xrm.WebApi.retrieveMultipleRecords(
                "duc_processautomationconfigurations",
                query,
            );

            const url = results?.entities?.[0]?.duc_value || null;
            if (!url) {
                console.warn(
                    `[AccountProfileHelpers] No '${UPDATE_ACCOUNT_CONFIG_KEY}' configuration found in duc_processautomationconfigurations`,
                );
            }
            return url;
        } catch (error: any) {
            console.error("[AccountProfileHelpers] Error reading UpdateAccount configuration:", error);
            return null;
        }
    }

    /**
     * Import the account values from the other environment: read the endpoint
     * from configuration and POST the profiles the catalogue returned.
     */
    static async postAccountUpdate(profiles: IAccountProfile[]): Promise<any> {
        const url = await this.getUpdateAccountApiUrl();

        if (!url) {
            throw new Error(
                `No '${UPDATE_ACCOUNT_CONFIG_KEY}' URL configured in duc_processautomationconfigurations`,
            );
        }

        console.log("[AccountProfileHelpers] UpdateAccount payload:", profiles);

        const response = await fetch(url, {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
            },
            body: JSON.stringify(profiles),
        });

        if (!response.ok) {
            throw new Error(`UpdateAccount API returned status ${response.status}: ${response.statusText}`);
        }

        return response.json().catch(() => null);
    }
}
