/* eslint-disable */
import * as React from "react";
import { IInputs } from "../generated/ManifestTypes";
import { SearchResults } from "./SearchResults";
import { logoBase64 } from "../images/logo64";
import { startSvgContent } from "../images/start";
import { magnify } from "../images/magnify";
import { calender } from "../images/calender";
import { clock } from "../images/clock";
import { check } from "../images/check";
import { filters } from "../images/filters";
import { AdvancedSearch } from './AdvancedSearch';
import { close } from '../images/close';

// =============================================================================
// HELPER FUNCTIONS
// =============================================================================

/**
 * Centralized error handler
 * Always shows alert dialog for errors
 */
const handleError = async (
    context: any,
    error: unknown,
    title: string = "Error"
): Promise<string> => {
    let message: string;

    if (error instanceof Error) {
        // Show full error details: message, name, and stack if available
        message = `${error.name}: ${error.message}`;
        if (error.stack) {
            message += `\n\nStack Trace:\n${error.stack}`;
        }
    } else if (typeof error === "string") {
        message = error;
    } else {
        // Try to stringify the error object
        try {
            message = JSON.stringify(error, null, 2);
        } catch {
            message = String(error);
        }
    }

    console.error(`[${title}]`, error);

    // Always show error dialog
    if (context?.navigation?.openAlertDialog) {
        try {
            await context.navigation.openAlertDialog({
                title,
                text: message
            });
        } catch (dialogError) {
            console.error("Failed to show error dialog:", dialogError);
        }
    }

    return message;
};
/**
 * Check if the app is running in offline mode
 */
const isOffline = (context: any): boolean => {
    try {
        // Check multiple possible offline indicators
        if (context?.client?.isOffline) {
            return context.client.isOffline() === true;
        }
        if (context?.mode?.isInOfflineMode !== undefined) {
            return context.mode.isInOfflineMode === true;
        }
        // Fallback to network status
        return !navigator.onLine;
    } catch (error) {
        console.warn("Error checking offline status:", error);
        return !navigator.onLine;
    }
};

/**
 * Execute a Dataverse query with offline/online support
 */
const executeQuery = async (
    context: any,
    entityName: string,
    query: string,
    operationType: 'retrieve' | 'retrieveMultiple' = 'retrieveMultiple'
): Promise<any> => {
    const offline = isOffline(context);
    const xrm: Xrm.XrmStatic = (window.parent as any).Xrm || (window as any).Xrm;

    try {
        if (offline) {
            console.log("Executing in OFFLINE mode");

            // Xrm.WebApi works offline automatically in latest versions
            if (operationType === 'retrieveMultiple') {
                return await xrm.WebApi.retrieveMultipleRecords(entityName, query);
            } else {
                throw new Error("Single record retrieve not supported in this helper");
            }
        } else {
            console.log("Executing in ONLINE mode");

            // Same Xrm.WebApi call works online
            if (operationType === 'retrieveMultiple') {
                return await xrm.WebApi.retrieveMultipleRecords(entityName, query);
            } else {
                throw new Error("Single record retrieve not supported in this helper");
            }
        }
    } catch (error) {
        console.error(`Query failed for ${entityName} (${offline ? 'OFFLINE' : 'ONLINE'}):`, error);
        throw error;
    }
};

/**
 * Update a record with offline/online support
 */
const updateRecord = async (
    context: any,
    entityName: string,
    recordId: string,
    data: any
): Promise<void> => {
    const offline = isOffline(context);
    const xrm: Xrm.XrmStatic = (window.parent as any).Xrm || (window as any).Xrm;

    try {
        if (offline) {
            console.log("Updating record in OFFLINE mode");
            await xrm.WebApi.updateRecord(entityName, recordId, data);
        } else {
            console.log("Updating record in ONLINE mode");
            await context.webAPI.updateRecord(entityName, recordId, data);
        }
    } catch (error) {
        console.error(`Update failed for ${entityName}:`, error);
        throw error;
    }
};

/**
 * Navigate - works in both online and offline
 */
const navigateTo = async (
    context: any,
    navigationOptions: any
): Promise<boolean> => {
    try {
        await context.navigation.navigateTo(navigationOptions);
        return true;
    } catch (error) {
        await handleError(context, error, "Navigation Error");
        return false;
    }
};

/**
 * Open form - works in both online and offline
 */
const openForm = async (
    context: any,
    formOptions: any,
    defaultValues: any = {}
): Promise<boolean> => {
    try {
        await context.navigation.openForm(formOptions, defaultValues);
        return true;
    } catch (error) {
        await handleError(context, error, "Form Open Error");
        return false;
    }
};

/**
 * Cache helper functions
 */
const CacheManager = {
    set: (key: string, value: any): void => {
        try {
            localStorage.setItem(key, JSON.stringify(value));
        } catch (error) {
            console.warn("Failed to set cache:", error);
        }
    },

    get: <T>(key: string, defaultValue: T | null = null): T | null => {
        try {
            const cached = localStorage.getItem(key);
            return cached ? JSON.parse(cached) : defaultValue;
        } catch (error) {
            console.warn("Failed to get cache:", error);
            return defaultValue;
        }
    },

    remove: (key: string): void => {
        try {
            localStorage.removeItem(key);
        } catch (error) {
            console.warn("Failed to remove cache:", error);
        }
    }
};

// =============================================================================
// TYPES & INTERFACES
// =============================================================================

interface State {
    searchText: string;
    pendingTodayBookings: number | null;
    completedTodayWorkorders: number | null;
    TodayCampaigns: number | null;
    userName: string;
    message?: string;
    isLoading: boolean;
    showResults: boolean;
    searchResults: any[];
    showAdvancedSearch: boolean;
    patrolStatus: 'none' | 'start' | 'end';
    activePatrolId?: string;
    isNaturalReserve: boolean;
    unknownAccountId?: string;
    incidentTypeId?: string;
    orgUnitId?: string;
}

interface IProps {
    context: ComponentFramework.Context<IInputs>;
}

interface LocalizedStrings {
    WelcomeBack: string;
    FindFacility: string;
    RemainingInspections: string;
    PendingInspections: string;
    CompletedInspections: string;
    ScheduledInspections: string;
    TodaysPatrols: string;
    StartInspection: string;
    SearchPlaceholder: string;
    NoResults: string;
    Loading: string;
    SearchPrompt: string;
    StartPatrol: string;
    EndPatrol: string;
    StartAnonymousInspection: string;
}

// =============================================================================
// STYLES
// =============================================================================

const STYLES = {
    textBrown: { color: "#A89360" },
    textGreen: { color: "#29A283" },
    bgBrownLight: { backgroundColor: "#ECE7DA" },
    bgGreenLight: { backgroundColor: "#CCEEE9" },
    welcomeBanner: {
        fontSize: "18px",
        backgroundColor: "#CFE0E5",
        paddingTop: 8,
        paddingBottom: 8,
        textAlign: "center" as const,
        width: "110%",
        maxwidth: "110%",
        marginBottom: 0,
        fontWeight: "bold" as const
    },
    actions: {
        backgroundColor: "#F3F3F3",
        padding: 12,
        paddingTop: 16,
        paddingBottom: 16,
        display: "flex",
        width: "100%",
        maxWidth: "100%",
        justifyContent: "space-between",
        flexDirection: "column" as const,
        gap: 12,
        marginleft: "2%",
    },
    icon: {
        width: 56,
        height: 56,
        minWidth: 56,
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        borderRadius: 4,
    },
    btnBlueDark: {
        backgroundColor: "#113f61",
        border: "none",
    },
    btnRedDark: {
        backgroundColor: "#8A1538",
        border: "none",
    },
    btnGreenDark: {
        backgroundColor: "#8A1538",
        border: "none",
    },
    btnOrangeDark: {
        backgroundColor: "#113f61",
        border: "none",
    },
    flexGrow1: { flexGrow: 1 },
    dFlex: { display: "flex" },
    alignItemsCenter: { alignItems: "center" },
    justifyContentBetween: { justifyContent: "space-between" },
    justifyContentCenter: { justifyContent: "center" },
    gap3: { gap: 12 },
    p3: { padding: 12 },
    py4: { paddingTop: 16, paddingBottom: 16 },
    mb3: { marginBottom: 12 },
    h1: { fontSize: 32, fontWeight: "bold" as const, margin: 0 },
    h6: { fontSize: 16, margin: 0 },
    rounded4: { borderRadius: 16 },
    rounded3: { borderRadius: 4 },
    border: { border: "1px solid #ccc" },
    textWhite: { color: "white" },
    textCenter: { textAlign: "center" as const },
    width: { width: '90%' }
};

// =============================================================================
// MAIN COMPONENT
// =============================================================================

export const Main = (props: IProps) => {
    const { context } = props;

    // Load localized strings
    const strings: LocalizedStrings = React.useMemo(() => {
        return {
            WelcomeBack: context.resources.getString("WelcomeBack"),
            FindFacility: context.resources.getString("FindFacility"),
            RemainingInspections: context.resources.getString("RemainingInspections"),
            CompletedInspections: context.resources.getString("CompletedInspections"),
            ScheduledInspections: context.resources.getString("ScheduledInspections"),
            StartInspection: context.resources.getString("StartInspection"),
            PendingInspections: context.resources.getString("PendingInspections"),
            SearchPlaceholder: context.resources.getString("SearchPlaceholder"),
            NoResults: context.resources.getString("NoResults"),
            Loading: context.resources.getString("Loading"),
            TodaysPatrols: context.resources.getString("TodaysPatrols"),
            SearchPrompt: context.resources.getString("SearchPrompt"),
            StartPatrol: context.resources.getString("StartPatrol"),
            EndPatrol: context.resources.getString("EndPatrol"),
            StartAnonymousInspection: context.resources.getString("StartAnonymousInspection"),
        };
    }, [context]);

    // Detect RTL
    const isRTL = React.useMemo(() => {
        const rtlLanguages = [1025, 1037, 1054, 1056, 1065, 1068, 1069, 1101, 1114, 1119];
        return rtlLanguages.includes(context.userSettings.languageId);
    }, [context.userSettings.languageId]);

    // State
    const [state, setState] = React.useState<State>({
        searchText: "",
        pendingTodayBookings: null,
        completedTodayWorkorders: null,
        TodayCampaigns: null,
        userName: strings.Loading,
        message: undefined,
        isLoading: false,
        searchResults: [],
        showResults: false,
        showAdvancedSearch: false,
        patrolStatus: 'none',
        activePatrolId: undefined,
        isNaturalReserve: false,
        unknownAccountId: undefined,
        incidentTypeId: undefined,
        orgUnitId: undefined,
    });

    // Image data URIs
    const startDataUri = "data:image/svg+xml;base64," + btoa(startSvgContent);
    const magnifyDataUri = "data:image/svg+xml;base64," + btoa(magnify);
    const calenderDataUri = "data:image/svg+xml;base64," + btoa(calender);
    const checkDataUri = "data:image/svg+xml;base64," + btoa(check);
    const clockDataUri = "data:image/svg+xml;base64," + btoa(clock);
    const filtersDataUri = "data:image/svg+xml;base64," + btoa(filters);
    const closeDataUri = "data:image/svg+xml;base64," + btoa(close);

    // =============================================================================
    // DATA LOADING FUNCTIONS
    // =============================================================================

    /**
     * Load today's counts (completed, remaining, campaigns)
     */
    const loadTodaysCounts = React.useCallback(async (
        userId: string
    ): Promise<{ completedToday: number; remainingToday: number; campaignsToday: number }> => {
        const CACHE_KEY = `MOCI_User_ResourceID_${userId}`;
        let completedToday = 0;
        let remainingToday = 0;
        let campaignsToday = 0;
        const today = new Date();
        today.setHours(0, 0, 0, 0);
        const todayStart = today.toISOString();
        today.setHours(23, 59, 59, 999);
        const todayEnd = today.toISOString();
        try {
            let resourceId: string | null = CacheManager.get(CACHE_KEY);

            // STEP 1: Get Bookable Resource ID
            if (!resourceId) {
                try {
                    const userQuery = `?$select=_duc_bookableresourceid_value&$filter=systemuserid eq '${userId}'`;
                    const userResults = await executeQuery(context, "systemuser", userQuery);

                    const retrievedResourceId =
                        userResults?.entities?.length > 0
                            ? userResults.entities[0]["_duc_bookableresourceid_value"]
                            : null;

                    if (!retrievedResourceId) {
                        console.warn(`User ${userId} is not linked to a Bookable Resource.`);
                        return { completedToday, remainingToday, campaignsToday };
                    }

                    resourceId = retrievedResourceId;
                    CacheManager.set(CACHE_KEY, resourceId);
                } catch (error) {
                    await handleError(context, error, "Failed to Get Bookable Resource");
                    return { completedToday, remainingToday, campaignsToday };
                }
            }

            // STEP 2: Completed Work Orders (Today)
            try {
                const completedQuery =
                    `?$select=msdyn_workorderid` +
                    `&$filter=_duc_assignedresource_value eq '${resourceId}' `+
                    `and duc_completiondate ge ${todayStart} and duc_completiondate le ${todayEnd}`;

                //  +
                // `and Microsoft.Dynamics.CRM.Today(PropertyName='duc_completiondate')`;

                const completedResults = await executeQuery(context, "msdyn_workorder", completedQuery);
                completedToday = completedResults?.entities?.length ?? 0;
            } catch (error) {
                console.warn("Failed to load completed work orders:", error);
                await handleError(context, error, "Failed to Load Completed Work Orders");
            }

            // STEP 3: Remaining Bookings (Today)
            try {
                const bookingStatusGuid = "f16d80d1-fd07-4237-8b69-187a11eb75f9";
                const remainingQuery =
                    `?$select=bookableresourcebookingid` +
                    `&$filter=_resource_value eq '${resourceId}' ` +
                    `and _bookingstatus_value eq '${bookingStatusGuid}' ` +
                    `and starttime ge ${todayStart} and starttime le ${todayEnd}`;

                //  +
                // `and Microsoft.Dynamics.CRM.Today(PropertyName='starttime')`;

                const remainingResults = await executeQuery(context, "bookableresourcebooking", remainingQuery);
                remainingToday = remainingResults?.entities?.length ?? 0;
            } catch (error) {
                console.warn("Failed to load remaining bookings:", error);
                await handleError(context, error, "Failed to Load Remaining Bookings");
            }

            return { completedToday, remainingToday, campaignsToday };

        } catch (error) {
            await handleError(context, error, "Failed to Load Today's Counts");
            return { completedToday: 0, remainingToday: 0, campaignsToday: 0 };
        }
    }, [context]);

    /**
     * Check organization unit and get unknown account
     */
    const checkOrganizationUnit = React.useCallback(async (userId: string): Promise<void> => {
        try {
            // Get user's organizational unit
            const userQuery = `?$select=_duc_organizationalunitid_value&$filter=systemuserid eq '${userId}'`;
            const userResults = await executeQuery(context, "systemuser", userQuery);

            const userResult = userResults?.entities?.[0];
            if (!userResult?._duc_organizationalunitid_value) {
                console.log("No organizational unit found for the user.");
                return;
            }

            const orgUnitId = userResult._duc_organizationalunitid_value;

            // Retrieve the organizational unit details
            const orgUnitQuery = `?$select=_duc_unknownaccount_value,duc_englishname&$filter=msdyn_organizationalunitid eq '${orgUnitId}'`;
            const orgUnitResults = await executeQuery(context, "msdyn_organizationalunit", orgUnitQuery);

            const orgUnitResult = orgUnitResults?.entities?.[0];
            if (!orgUnitResult) {
                console.warn("Organization unit not found");
                return;
            }

            const orgUnitName = orgUnitResult.duc_englishname || "";
            const unknownAccountId = orgUnitResult._duc_unknownaccount_value || undefined;

            // Check if organization unit name contains "Natural Reserve"
            const isNaturalReserve = orgUnitName.includes("Inspection Section – Natural Reserves");

            let incidentTypeId: string | undefined = undefined;

            // If Natural Reserve, get the incident type for this org unit
            if (isNaturalReserve) {
                try {
                    const incidentTypeQuery = `?$filter=_duc_organizationalunitid_value eq '${orgUnitId}'&$top=1`;
                    const incidentTypeResults = await executeQuery(context, "msdyn_incidenttype", incidentTypeQuery);

                    if (incidentTypeResults?.entities?.length > 0) {
                        incidentTypeId = incidentTypeResults.entities[0].msdyn_incidenttypeid;
                        console.log("Incident Type ID found:", incidentTypeId);
                    }
                } catch (error) {
                    console.warn("Error fetching incident type:", error);
                    await handleError(context, error, "Failed to Fetch Incident Type");
                }
            }

            setState(prev => ({
                ...prev,
                isNaturalReserve: isNaturalReserve,
                unknownAccountId: unknownAccountId,
                incidentTypeId: incidentTypeId,
                orgUnitId: orgUnitId
            }));

            console.log("Organization Unit:", orgUnitName);
            console.log("Is Natural Reserve:", isNaturalReserve);

        } catch (error) {
            await handleError(context, error, "Failed to Check Organization Unit");
            setState(prev => ({
                ...prev,
                isNaturalReserve: false,
                unknownAccountId: undefined,
                incidentTypeId: undefined,
                orgUnitId: undefined
            }));
        }
    }, [context]);

    /**
     * Check patrol status
     */
    const checkPatrolStatus = React.useCallback(async (): Promise<void> => {
        try {
            const userSettings = context.userSettings;
            const userId = (userSettings.userId ?? "").replace(/[{}]/g, "");
            const today = new Date().toISOString().split('T')[0];

            // Check for active patrol (status = 2)
            const activePatrolQuery = `?$filter=_owninguser_value eq '${userId}' and duc_campaigninternaltype eq 100000004 and duc_campaignstatus eq 2 and duc_fromdate le ${today} and duc_todate ge ${today}&$top=1`;

            const activePatrolResults = await executeQuery(context, "new_inspectioncampaign", activePatrolQuery);

            if (activePatrolResults?.entities?.length > 0) {
                setState(prev => ({
                    ...prev,
                    patrolStatus: 'end',
                    activePatrolId: activePatrolResults.entities[0].new_inspectioncampaignid
                }));
                return;
            }

            // Check for available patrol to start (status = 1 or 100000004)
            const availablePatrolQuery = `?$filter=_owninguser_value eq '${userId}' and duc_campaigninternaltype eq 100000004 and (duc_campaignstatus eq 1 or duc_campaignstatus eq 100000004) and duc_fromdate le ${today} and duc_todate ge ${today}&$top=1`;

            const availablePatrolResults = await executeQuery(context, "new_inspectioncampaign", availablePatrolQuery);

            if (availablePatrolResults?.entities?.length > 0) {
                setState(prev => ({
                    ...prev,
                    patrolStatus: 'start',
                    activePatrolId: availablePatrolResults.entities[0].new_inspectioncampaignid
                }));
            } else {
                setState(prev => ({
                    ...prev,
                    patrolStatus: 'none',
                    activePatrolId: undefined
                }));
            }

        } catch (error) {
            await handleError(context, error, "Failed to Check Patrol Status");
            setState(prev => ({
                ...prev,
                patrolStatus: 'none',
                activePatrolId: undefined
            }));
        }
    }, [context]);

    /**
     * Load user data
     */
    const loadUserData = React.useCallback(async (): Promise<void> => {
        try {
            const userSettings = context.userSettings;
            const userId = (userSettings.userId ?? "").replace(/[{}]/g, "");

            // Load user name
            let username = strings.Loading;
            try {
                const userQuery = `?$filter=systemuserid eq '${userId}'&$select=duc_usernamearabic`;
                const results = await executeQuery(context, "systemuser", userQuery);
                const res = results?.entities?.[0];
                username = res?.duc_usernamearabic ?? userSettings?.userName ?? "Inspector";
            } catch (error) {
                console.warn("Failed to load username:", error);
                await handleError(context, error, "Failed to Load Username");
                username = userSettings?.userName ?? "Inspector";
            }

            // Load today's counts
            const { completedToday, remainingToday, campaignsToday } = await loadTodaysCounts(userId);

            setState(prev => ({
                ...prev,
                pendingTodayBookings: remainingToday,
                completedTodayWorkorders: completedToday,
                TodayCampaigns: campaignsToday,
                userName: username,
            }));

            // Cache the data
            CacheManager.set("MOCI_userCounts", {
                userName: username,
                remainingToday,
                completedToday,
                campaignsToday
            });

            // Check organization unit (works in both online and offline)
            await checkOrganizationUnit(userId);

        } catch (error) {
            await handleError(context, error, "Failed to Load User Data");

            // Try to restore from cache
            const cached = CacheManager.get<any>("MOCI_userCounts");
            if (cached) {
                setState(prev => ({
                    ...prev,
                    userName: cached.userName ?? "Inspector",
                    pendingTodayBookings: cached.remainingToday ?? null,
                    completedTodayWorkorders: cached.completedToday ?? null,
                    TodayCampaigns: cached.campaignsToday ?? null,
                }));
            }
        }
    }, [context, strings.Loading, loadTodaysCounts, checkOrganizationUnit]);

    // =============================================================================
    // EVENT HANDLERS
    // =============================================================================

    const handleAdvancedSearchResults = (results: any[]): void => {
        setState(prev => ({
            ...prev,
            searchResults: results,
            showResults: true
        }));
    };

    const handleOpenAdvancedSearch = (): void => {
        setState(prev => ({ ...prev, showAdvancedSearch: true }));
    };

    const handleCloseAdvancedSearch = (): void => {
        setState(prev => ({ ...prev, showAdvancedSearch: false }));
    };

    const onKeyUp = (e: React.KeyboardEvent<HTMLInputElement>): void => {
        if (e.key === "Enter") void onSearchClick();
    };

    const onSearchClick = async (): Promise<void> => {
        const text = state.searchText.trim();
        if (!text) {
            setState(prev => ({ ...prev, message: strings.SearchPrompt }));
            return;
        }

        setState(prev => ({ ...prev, isLoading: true, message: undefined }));
        const searchKey = `MOCI_lastSearch_${text.toLowerCase()}`;

        try {
            const query = `?$top=100`;
            const results = await executeQuery(context, "account", query);

            let entities = results?.entities ?? [];

            // Filter by search text
            if (text) {
                const lowerText = text.toLowerCase();
                entities = entities.filter((entity: any) => {
                    const entityName = entity.name as string | undefined;
                    return entityName?.toLowerCase().startsWith(lowerText) ?? false;
                });
            }

            if (entities.length === 0) {
                setState(prev => ({ ...prev, message: strings.NoResults, isLoading: false }));
            } else if (entities.length === 1) {
                // Navigate directly to single result (works in offline too)
                await navigateTo(context, {
                    pageType: "entityrecord",
                    entityName: "account",
                    entityId: entities[0].accountid
                });
                setState(prev => ({ ...prev, isLoading: false }));
            } else {
                setState(prev => ({
                    ...prev,
                    showResults: true,
                    searchResults: entities,
                    message: undefined,
                    isLoading: false
                }));
            }

            // Cache results
            CacheManager.set(searchKey, { entities });

        } catch (error) {
            await handleError(context, error, "Search Failed");

            // Try to load from cache
            const cached = CacheManager.get<any>(searchKey);
            if (cached?.entities) {
                setState(prev => ({
                    ...prev,
                    showResults: true,
                    searchResults: cached.entities,
                    isLoading: false
                }));
            } else {
                setState(prev => ({ ...prev, isLoading: false }));
            }
        }
    };

    const onRecordClick = (accountId: string): void => {
        setState(prev => ({ ...prev, showResults: false }));
        void navigateTo(context, {
            pageType: "entityrecord",
            entityName: "account",
            entityId: accountId
        });
    };

    const onCloseResults = (): void => {
        setState(prev => ({ ...prev, showResults: false }));
    };

    // =============================================================================
    // NAVIGATION HANDLERS
    // =============================================================================

    const openScheduled = (): void => {
        void navigateTo(context, {
            pageType: "entitylist",
            entityName: "bookableresourcebooking",
            viewId: "4073baca-cc5f-e611-8109-000d3a146973"
        });
    };

    const openPendingWorkorders = (): void => {
        void navigateTo(context, {
            pageType: "entitylist",
            entityName: "msdyn_workorder",
            viewId: "bee0efc7-40e4-f011-8406-6045bd9c224c"
        });
    };

    const openScheduledCampaigns = (): void => {
        void navigateTo(context, {
            pageType: "entitylist",
            entityName: "new_inspectioncampaign",
            viewId: "0628a5aa-88d7-f011-8406-7c1e524df347"
        });
    };

    const closedWorkorders = (): void => {
        void navigateTo(context, {
            pageType: "entitylist",
            entityName: "msdyn_workorder",
            viewId: "aad92cc9-00c6-f011-8544-000d3a2274a5"
        });
    };

    // =============================================================================
    // FORM HANDLERS
    // =============================================================================

    const startInspection = (): void => {
        void openForm(
            context,
            { entityName: "msdyn_workorder", useQuickCreateForm: true },
            {}
        );
    };

    const startAnonymousInspection = (): void => {
        const defaultValues: any = {
            duc_anonymouscustomer: true
        };

        if (state.unknownAccountId) {
            defaultValues.duc_subaccount = state.unknownAccountId;
            defaultValues.msdyn_serviceaccount = state.unknownAccountId;
        }

        if (state.incidentTypeId) {
            defaultValues.msdyn_primaryincidenttype = state.incidentTypeId;
        }

        void openForm(
            context,
            { entityName: "msdyn_workorder", useQuickCreateForm: true },
            defaultValues
        );
    };

    // =============================================================================
    // PATROL HANDLERS
    // =============================================================================

    const startPatrol = async (): Promise<void> => {
        if (!state.activePatrolId) {
            await handleError(context, "No patrol campaign found", "Patrol Error");
            return;
        }

        try {
            setState(prev => ({ ...prev, isLoading: true }));

            await updateRecord(context, "new_inspectioncampaign", state.activePatrolId, {
                duc_campaignstatus: 2
            });

            setState(prev => ({
                ...prev,
                patrolStatus: 'end',
                isLoading: false,
                message: "Patrol started successfully"
            }));

            setTimeout(() => {
                setState(prev => ({ ...prev, message: undefined }));
            }, 3000);

        } catch (error) {
            await handleError(context, error, "Failed to Start Patrol");
            setState(prev => ({ ...prev, isLoading: false }));
        }
    };

    const endPatrol = async (): Promise<void> => {
        try {
            setState(prev => ({ ...prev, isLoading: true }));

            const userSettings = context.userSettings;
            const userId = (userSettings.userId ?? "").replace(/[{}]/g, "");
            const today = new Date().toISOString().split('T')[0];

            const activePatrolQuery = `?$filter=_owninguser_value eq '${userId}' and duc_campaigninternaltype eq 100000004 and duc_campaignstatus eq 2 and duc_fromdate le ${today} and duc_todate ge ${today}&$top=1`;

            const activePatrolResults = await executeQuery(context, "new_inspectioncampaign", activePatrolQuery);

            if (!activePatrolResults?.entities?.length) {
                setState(prev => ({
                    ...prev,
                    isLoading: false
                }));
                await handleError(context, "No active patrol found", "Patrol Error");
                return;
            }

            const patrolId = activePatrolResults.entities[0].new_inspectioncampaignid;

            await updateRecord(context, "new_inspectioncampaign", patrolId, {
                duc_campaignstatus: 100000004
            });

            await checkPatrolStatus();

            setState(prev => ({
                ...prev,
                isLoading: false,
                message: "Patrol ended successfully"
            }));

            setTimeout(() => {
                setState(prev => ({ ...prev, message: undefined }));
            }, 3000);

        } catch (error) {
            await handleError(context, error, "Failed to End Patrol");
            setState(prev => ({ ...prev, isLoading: false }));
        }
    };

    // =============================================================================
    // EFFECTS
    // =============================================================================

    React.useEffect(() => {
        void loadUserData();
        void checkPatrolStatus();
    }, [loadUserData, checkPatrolStatus]);

    // =============================================================================
    // RENDER HELPERS
    // =============================================================================

    const {
        searchText,
        pendingTodayBookings,
        completedTodayWorkorders,
        TodayCampaigns,
        userName,
        message,
        isLoading,
        showResults,
        searchResults,
        patrolStatus,
        isNaturalReserve
    } = state;

    const isActionDisabled = isLoading;

    // Determine which buttons to show
    const showScheduledInspections = !isNaturalReserve || (isNaturalReserve && patrolStatus === 'none');
    const showStartInspection = !isNaturalReserve || (isNaturalReserve && patrolStatus === 'end') || (isNaturalReserve && patrolStatus === 'none');
    const showStartAnonymousInspection = isNaturalReserve && patrolStatus === 'end';
    const showStartPatrol = patrolStatus === 'start';
    const showEndPatrol = patrolStatus === 'end';

    const getButtonStyle = (baseStyle: React.CSSProperties) => ({
        ...baseStyle,
        opacity: isActionDisabled ? 0.6 : 1,
        cursor: isActionDisabled ? 'not-allowed' : 'pointer',
    });

    const containerStyle = {
        ...STYLES.dFlex,
        flexDirection: 'column' as const,
        width: '98%',
        maxWidth: '98%',
        height: '100%',
        maxHeight: '100%',
        direction: isRTL ? 'rtl' as const : 'ltr' as const,
        position: 'fixed' as const,
        top: 0,
        left: 0,
        right: 0,
        bottom: 0,
        display: 'flex',
        justifyContent: 'center',
        alignItems: 'center',
        zIndex: 1000,
        paddingtop: '0.5%',
        margin: 'auto',
    };

    // =============================================================================
    // RENDER
    // =============================================================================

    return React.createElement(
        React.Fragment,
        null,

        React.createElement(
            "div",
            { style: containerStyle },

            // Header
            React.createElement(
                "header",
                { style: { ...STYLES.dFlex, ...STYLES.alignItemsCenter, ...STYLES.justifyContentCenter, paddingTop: 16, paddingBottom: 16 } },
                React.createElement(
                    "a",
                    { href: "#", style: { display: 'block' } },
                    React.createElement("img", { alt: "", src: logoBase64 })
                )
            ),

            // Welcome Banner
            React.createElement(
                "div",
                { style: STYLES.welcomeBanner },
                `${strings.WelcomeBack}, ${userName}`
            ),

            // Main Content
            React.createElement(
                "div",
                { style: { ...STYLES.flexGrow1, ...STYLES.p3, ...STYLES.width, marginleft: '0.5%' } },

                // Search Bar
                React.createElement(
                    "div",
                    { style: { ...STYLES.dFlex, ...STYLES.gap3, ...STYLES.mb3, opacity: isLoading ? 0.6 : 1 } },
                    React.createElement(
                        "div",
                        { style: { ...STYLES.flexGrow1, ...STYLES.dFlex, ...STYLES.border, ...STYLES.rounded3, ...STYLES.alignItemsCenter, ...STYLES.gap3, ...STYLES.p3 } },
                        React.createElement("img", { src: magnifyDataUri }),
                        React.createElement(
                            "div",
                            { style: STYLES.flexGrow1 },
                            React.createElement("input", {
                                type: "text",
                                placeholder: strings.SearchPlaceholder,
                                value: searchText,
                                onChange: (e) => setState(prev => ({ ...prev, searchText: e.target.value })),
                                onKeyUp: onKeyUp,
                                disabled: isLoading,
                                style: {
                                    border: 'none',
                                    width: '90%',
                                    outline: 'none',
                                    direction: isRTL ? 'rtl' : 'ltr',
                                    background: 'none'
                                }
                            })
                        )
                    ),
                    React.createElement(
                        "div",
                        {
                            style: {
                                marginLeft: isRTL ? 0 : "auto",
                                marginRight: isRTL ? "auto" : 0,
                                ...STYLES.icon,
                                ...STYLES.rounded3,
                                ...STYLES.border,
                                cursor: 'pointer'
                            },
                            onClick: handleOpenAdvancedSearch
                        },
                        React.createElement("img", { src: filtersDataUri })
                    )
                ),

                // Inspection Cards
                React.createElement(
                    "div",
                    { style: { display: "flex", flexDirection: "column" as const, gap: 12 } },

                    // Pending Card
                    React.createElement(
                        "div",
                        {
                            onClick: openScheduled,
                            style: {
                                ...STYLES.border,
                                ...STYLES.rounded4,
                                ...STYLES.dFlex,
                                ...STYLES.alignItemsCenter,
                                ...STYLES.justifyContentBetween,
                                ...STYLES.p3,
                                ...STYLES.gap3,
                                cursor: isActionDisabled ? 'not-allowed' : 'pointer',
                                opacity: isActionDisabled ? 0.5 : 1,
                                pointerEvents: isActionDisabled ? 'none' : 'auto'
                            }
                        },
                        React.createElement(
                            "div",
                            { style: { ...STYLES.flexGrow1, ...STYLES.dFlex, ...STYLES.alignItemsCenter, ...STYLES.gap3 } },
                            React.createElement(
                                "div",
                                { style: { ...STYLES.icon, ...STYLES.rounded3, ...STYLES.bgBrownLight } },
                                React.createElement("img", { src: clockDataUri })
                            ),
                            React.createElement("h6", { style: STYLES.h6 }, strings.RemainingInspections)
                        ),
                        React.createElement("div", { style: { ...STYLES.textBrown, ...STYLES.h1 } }, pendingTodayBookings ?? "...")
                    ),

                    // Completed Card
                    React.createElement(
                        "div",
                        {
                            onClick: closedWorkorders,
                            style: {
                                ...STYLES.border,
                                ...STYLES.rounded4,
                                ...STYLES.dFlex,
                                ...STYLES.alignItemsCenter,
                                ...STYLES.justifyContentBetween,
                                ...STYLES.p3,
                                ...STYLES.gap3,
                                cursor: 'pointer'
                            }
                        },
                        React.createElement(
                            "div",
                            { style: { ...STYLES.flexGrow1, ...STYLES.dFlex, ...STYLES.alignItemsCenter, ...STYLES.gap3 } },
                            React.createElement(
                                "div",
                                { style: { ...STYLES.icon, ...STYLES.rounded3, ...STYLES.bgGreenLight } },
                                React.createElement("img", { src: checkDataUri })
                            ),
                            React.createElement("h6", { style: STYLES.h6 }, strings.CompletedInspections)
                        ),
                        React.createElement("div", { style: { ...STYLES.textGreen, ...STYLES.h1 } }, completedTodayWorkorders ?? "...")
                    ),

                    // Pending Workorders Card
                    React.createElement(
                        "div",
                        {
                            onClick: openPendingWorkorders,
                            style: {
                                ...STYLES.border,
                                ...STYLES.rounded4,
                                ...STYLES.dFlex,
                                ...STYLES.alignItemsCenter,
                                ...STYLES.justifyContentBetween,
                                ...STYLES.p3,
                                ...STYLES.gap3,
                                cursor: isActionDisabled ? 'not-allowed' : 'pointer',
                                opacity: isActionDisabled ? 0.5 : 1,
                                pointerEvents: isActionDisabled ? 'none' : 'auto'
                            }
                        },
                        React.createElement(
                            "div",
                            { style: { ...STYLES.flexGrow1, ...STYLES.dFlex, ...STYLES.alignItemsCenter, ...STYLES.gap3 } },
                            React.createElement(
                                "div",
                                { style: { ...STYLES.icon, ...STYLES.rounded3, ...STYLES.bgBrownLight } },
                                React.createElement("img", { src: clockDataUri })
                            ),
                            React.createElement("h6", { style: STYLES.h6 }, strings.PendingInspections)
                        ),
                        React.createElement("div", { style: { ...STYLES.textBrown, ...STYLES.h1 } }, pendingTodayBookings ?? "...")
                    )
                )
            ),

            // Action Buttons
            React.createElement(
                "div",
                { style: STYLES.actions },

                // Scheduled Inspections
                showScheduledInspections && React.createElement(
                    "button",
                    {
                        onClick: openScheduled,
                        disabled: isActionDisabled,
                        style: getButtonStyle({
                            ...STYLES.btnBlueDark,
                            ...STYLES.textWhite,
                            ...STYLES.rounded4,
                            ...STYLES.p3,
                            ...STYLES.dFlex,
                            ...STYLES.alignItemsCenter,
                            ...STYLES.justifyContentCenter,
                            ...STYLES.gap3
                        })
                    },
                    React.createElement("img", { src: calenderDataUri }),
                    React.createElement("span", null, strings.ScheduledInspections)
                ),

                // Start Inspection
                showStartInspection && React.createElement(
                    "button",
                    {
                        onClick: startInspection,
                        disabled: isActionDisabled,
                        style: getButtonStyle({
                            ...STYLES.btnRedDark,
                            ...STYLES.textWhite,
                            ...STYLES.rounded4,
                            ...STYLES.p3,
                            ...STYLES.dFlex,
                            ...STYLES.alignItemsCenter,
                            ...STYLES.justifyContentCenter,
                            ...STYLES.gap3
                        })
                    },
                    [
                        React.createElement("img", { src: startDataUri, key: "icon" }),
                        React.createElement("span", { key: "text" }, strings.StartInspection)
                    ]
                ),

                // Start Anonymous Inspection
                showStartAnonymousInspection && React.createElement(
                    "button",
                    {
                        onClick: startAnonymousInspection,
                        disabled: isActionDisabled,
                        style: getButtonStyle({
                            ...STYLES.btnOrangeDark,
                            ...STYLES.textWhite,
                            ...STYLES.rounded4,
                            ...STYLES.p3,
                            ...STYLES.dFlex,
                            ...STYLES.alignItemsCenter,
                            ...STYLES.justifyContentCenter,
                            ...STYLES.gap3
                        })
                    },
                    [
                        React.createElement("img", { src: startDataUri, key: "icon" }),
                        React.createElement("span", { key: "text" }, strings.StartAnonymousInspection)
                    ]
                ),

                // Start Patrol
                showStartPatrol && React.createElement(
                    "button",
                    {
                        onClick: startPatrol,
                        disabled: isActionDisabled,
                        style: getButtonStyle({
                            ...STYLES.btnBlueDark,
                            ...STYLES.textWhite,
                            ...STYLES.rounded4,
                            ...STYLES.p3,
                            ...STYLES.dFlex,
                            ...STYLES.alignItemsCenter,
                            ...STYLES.justifyContentCenter,
                            ...STYLES.gap3
                        })
                    },
                    [
                        React.createElement("img", { src: startDataUri, key: "icon" }),
                        React.createElement("span", { key: "text" }, strings.StartPatrol)
                    ]
                ),

                // End Patrol
                showEndPatrol && React.createElement(
                    "button",
                    {
                        onClick: endPatrol,
                        disabled: isActionDisabled,
                        style: getButtonStyle({
                            ...STYLES.btnGreenDark,
                            ...STYLES.textWhite,
                            ...STYLES.rounded4,
                            ...STYLES.p3,
                            ...STYLES.dFlex,
                            ...STYLES.alignItemsCenter,
                            ...STYLES.justifyContentCenter,
                            ...STYLES.gap3
                        })
                    },
                    [
                        React.createElement("img", { src: startDataUri, key: "icon" }),
                        React.createElement("span", { key: "text" }, strings.EndPatrol)
                    ]
                )
            )
        ),

        // Advanced Search Modal
        React.createElement(AdvancedSearch, {
            context: context,
            isOpen: state.showAdvancedSearch,
            onClose: handleCloseAdvancedSearch,
            onSearchResults: handleAdvancedSearchResults
        }),

        // Search Results Modal
        showResults && React.createElement(SearchResults, {
            results: searchResults,
            onRecordClick: onRecordClick,
            onClose: onCloseResults,
            context: context
        })
    );
};