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
import { MultiTypeInspection } from './MultiTypeInspection';
import { close } from '../images/close';

// Cache interface for organization unit data
interface OrgUnitCache {
    orgUnitId: string;
    orgUnitName: string;
    isNaturalReserve: boolean;
    isWildlifeSection: boolean;
    unknownAccountId?: string;
    unknownAccountName?: string;
    incidentTypeId?: string;
    incidentTypeName?: string;
    timestamp: number;
}

// Define the interface for the component's internal state.
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
    patrolStatus: 'none' | 'start' | 'end'; // Track patrol button state
    activePatrolId?: string; // Store active patrol campaign ID
    activePatrolName?: string; // Store active patrol campaign name
    isNaturalReserve: boolean; // Track if org unit is Natural Reserve
    isWildlifeSection: boolean; // Track if org unit is Wildlife section
    unknownAccountId?: string; // Store unknown account ID for anonymous inspections
    unknownAccountName?: string; // Store unknown account name
    incidentTypeName?: string; // Store incident type name for Natural Reserve
    incidentTypeId?: string; // Store incident type ID for Natural Reserve
    orgUnitId?: string; // Store organization unit ID
    showMultiTypeInspection: boolean; // Track if MultiTypeInspection modal is open
}

// Define the interface for the component's properties (props) coming from the Power Apps Component Framework (PCF).
interface IProps {
    context: ComponentFramework.Context<IInputs>;
}

// Localized strings interface
interface LocalizedStrings {
    WelcomeBack: string;
    FindFacility: string;
    RemainingInspections: string;
    CompletedInspections: string;
    ScheduledInspections: string;
    TodaysPatrols: string;
    StartInspection: string;
    SearchPlaceholder: string;
    NoResults: string;
    Loading: string;
    SearchPrompt: string;
    OfflineNavigationBlocked: string;
    OfflineQuickCreateBlocked: string;
    StartPatrol: string;
    EndPatrol: string;
    StartAnonymousInspection: string;
    PendingInspections: string;
    StartMultiTypeInspection: string;
}

// --- Custom Styles Derived from main.css & Bootstrap ---
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

// Cache constants
const ORG_UNIT_CACHE_KEY = 'MOCI_OrgUnit_Cache';
const CACHE_DURATION = 24 * 60 * 60 * 1000; // 24 hours

export const Main = (props: IProps) => {
    // Load localized strings
    const strings: LocalizedStrings = React.useMemo(() => {
        const ctx = props.context;
        return {
            WelcomeBack: ctx.resources.getString("WelcomeBack"),
            FindFacility: ctx.resources.getString("FindFacility"),
            RemainingInspections: ctx.resources.getString("RemainingInspections"),
            CompletedInspections: ctx.resources.getString("CompletedInspections"),
            ScheduledInspections: ctx.resources.getString("ScheduledInspections"),
            StartInspection: ctx.resources.getString("StartInspection"),
            SearchPlaceholder: ctx.resources.getString("SearchPlaceholder"),
            NoResults: ctx.resources.getString("NoResults"),
            Loading: ctx.resources.getString("Loading"),
            TodaysPatrols: ctx.resources.getString("TodaysPatrols"),
            SearchPrompt: ctx.resources.getString("SearchPrompt"),
            OfflineNavigationBlocked: ctx.resources.getString("OfflineNavigationBlocked"),
            OfflineQuickCreateBlocked: ctx.resources.getString("OfflineQuickCreateBlocked"),
            StartPatrol: ctx.resources.getString("StartPatrol"),
            EndPatrol: ctx.resources.getString("EndPatrol"),
            StartAnonymousInspection: ctx.resources.getString("StartAnonymousInspection"),
            PendingInspections: ctx.resources.getString("PendingInspections"),
            StartMultiTypeInspection: ctx.resources.getString("StartMultiTypeInspection"),
        };
    }, [props.context]);

    // Detect RTL
    const isRTL = React.useMemo(() => {
        const rtlLanguages = [1025, 1037, 1054, 1056, 1065, 1068, 1069, 1101, 1114, 1119];
        return rtlLanguages.includes(props.context.userSettings.languageId);
    }, [props.context.userSettings.languageId]);

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
        activePatrolName: undefined,
        isNaturalReserve: false,
        isWildlifeSection: false,
        unknownAccountId: undefined,
        incidentTypeId: undefined,
        orgUnitId: undefined,
        showMultiTypeInspection: false,
    });

    const startDataUri = "data:image/svg+xml;base64," + btoa(startSvgContent);
    const magnifyDataUri = "data:image/svg+xml;base64," + btoa(magnify);
    const calenderDataUri = "data:image/svg+xml;base64," + btoa(calender);
    const checkDataUri = "data:image/svg+xml;base64," + btoa(check);
    const clockDataUri = "data:image/svg+xml;base64," + btoa(clock);
    const filtersDataUri = "data:image/svg+xml;base64," + btoa(filters);
    const closeDataUri = "data:image/svg+xml;base64," + btoa(close);

    const isOffline = (): boolean => {
        const ctx: any = props.context;
        return ctx.mode?.isInOfflineMode === true;
    };

    const handleAdvancedSearchResults = (results: any[]): void => {
        setState(prev => ({
            ...prev,
            searchResults: results,
            showResults: true
        }));
    };

    // Get organization unit data from cache
    const getOrgUnitFromCache = (userId: string): OrgUnitCache | null => {
        try {
            const cacheKey = `${ORG_UNIT_CACHE_KEY}_${userId}`;
            const cached = localStorage.getItem(cacheKey);
            if (!cached) return null;

            const cacheData: OrgUnitCache = JSON.parse(cached);
            const now = Date.now();

            // Check if cache is still valid
            if (now - cacheData.timestamp > CACHE_DURATION) {
                localStorage.removeItem(cacheKey);
                return null;
            }

            return cacheData;
        } catch (error) {
            console.error("Error reading org unit cache:", error);
            return null;
        }
    };

    // Save organization unit data to cache
    const saveOrgUnitToCache = (userId: string, data: Omit<OrgUnitCache, 'timestamp'>): void => {
        try {
            const cacheKey = `${ORG_UNIT_CACHE_KEY}_${userId}`;
            const cacheData: OrgUnitCache = {
                ...data,
                timestamp: Date.now()
            };
            localStorage.setItem(cacheKey, JSON.stringify(cacheData));
            console.log("Organization unit data cached successfully");
        } catch (error) {
            console.error("Error saving org unit cache:", error);
        }
    };

    // Check organization unit and get unknown account
    const checkOrganizationUnit = React.useCallback(async (ctx: any, userId: string): Promise<void> => {
        try {
            // Try to get from cache first
            const cachedData = getOrgUnitFromCache(userId);
            if (cachedData) {
                console.log("Using cached organization unit data");
                setState(prev => ({
                    ...prev,
                    isNaturalReserve: cachedData.isNaturalReserve,
                    isWildlifeSection: cachedData.isWildlifeSection,
                    unknownAccountId: cachedData.unknownAccountId,
                    unknownAccountName: cachedData.unknownAccountName,
                    incidentTypeId: cachedData.incidentTypeId,
                    incidentTypeName: cachedData.incidentTypeName,
                    orgUnitId: cachedData.orgUnitId
                }));
                return;
            }

            // Get user's organizational unit
            const userResult = await ctx.webAPI.retrieveRecord(
                "systemuser",
                userId,
                "?$select=_duc_organizationalunitid_value"
            );

            if (!userResult._duc_organizationalunitid_value) {
                console.log("No organizational unit found for the user.");
                return;
            }

            const orgUnitId = userResult._duc_organizationalunitid_value;

            // Retrieve the organizational unit details
            const orgUnitResult = await ctx.webAPI.retrieveRecord(
                "msdyn_organizationalunit",
                orgUnitId,
                "?$select=_duc_unknownaccount_value,duc_englishname&$expand=duc_unknownaccount($select=name)"
            );

            const orgUnitName = orgUnitResult.duc_englishname || "";
            const unknownAccountId = orgUnitResult._duc_unknownaccount_value || undefined;
            const unknownAccountName = orgUnitResult.duc_unknownaccount?.name || undefined;

            // Check if organization unit name is "Natural Reserve" (case insensitive)
            const isNaturalReserve = orgUnitName.includes("Inspection Section – Natural Reserves");

            // Check if organization unit is Wildlife section
            const isWildlifePlantLife = orgUnitName.includes("Wildlife - Plant Life Section");
            const isWildlifeNaturalResources = orgUnitName.includes("Wildlife - Natural Resources Section");
            const isWildlifeSection = isWildlifePlantLife || isWildlifeNaturalResources;

            let incidentTypeId: string | undefined = undefined;
            let incidentTypeName: string | undefined = undefined;

            // If Natural Reserve, get the incident type for this org unit
            if (isNaturalReserve) {
                try {
                    const incidentTypeQuery = `?$filter=_duc_organizationalunitid_value eq '${orgUnitId}'&$select=msdyn_incidenttypeid,msdyn_name&$top=1`;
                    const incidentTypeResults = await ctx.webAPI.retrieveMultipleRecords(
                        "msdyn_incidenttype",
                        incidentTypeQuery
                    );

                    if (incidentTypeResults.entities.length > 0) {
                        incidentTypeId = incidentTypeResults.entities[0].msdyn_incidenttypeid;
                        incidentTypeName = incidentTypeResults.entities[0].msdyn_name;
                        console.log("Incident Type ID found:", incidentTypeId);
                        console.log("Incident Type Name found:", incidentTypeName);
                    } else {
                        console.warn("No incident type found for Natural Reserve org unit");
                    }
                } catch (error) {
                    console.error("Error fetching incident type:", error);
                }
            }

            // Save to cache
            saveOrgUnitToCache(userId, {
                orgUnitId,
                orgUnitName,
                isNaturalReserve,
                isWildlifeSection,
                unknownAccountId,
                unknownAccountName,
                incidentTypeId,
                incidentTypeName
            });

            setState(prev => ({
                ...prev,
                isNaturalReserve: isNaturalReserve,
                isWildlifeSection: isWildlifeSection,
                unknownAccountId: unknownAccountId,
                unknownAccountName: unknownAccountName,
                incidentTypeId: incidentTypeId,
                incidentTypeName: incidentTypeName,
                orgUnitId: orgUnitId
            }));

            console.log("Organization Unit:", orgUnitName);
            console.log("Organization Unit ID:", orgUnitId);
            console.log("Is Natural Reserve:", isNaturalReserve);
            console.log("Is Wildlife Section:", isWildlifeSection);
            console.log("Unknown Account ID:", unknownAccountId);
            console.log("Unknown Account Name:", unknownAccountName);
            console.log("Incident Type ID:", incidentTypeId);
            console.log("Incident Type Name:", incidentTypeName);

        } catch (error) {
            console.error("Error checking organization unit:", error);
            setState(prev => ({
                ...prev,
                isNaturalReserve: false,
                isWildlifeSection: false,
                unknownAccountId: undefined,
                unknownAccountName: undefined,
                incidentTypeId: undefined,
                incidentTypeName: undefined,
                orgUnitId: undefined
            }));
        }
    }, []);

    // Check patrol status on load
    const checkPatrolStatus = React.useCallback(async (): Promise<void> => {
        const ctx: any = props.context;
        try {
            const userSettings = ctx.userSettings;
            const userId = (userSettings.userId ?? "").replace("{", "").replace("}", "");

            // Get today's date in ISO format
            const today = new Date().toISOString().split('T')[0];

            // Check for active patrol (status = 2)
            const activePatrolQuery = `?$filter=_owninguser_value eq '${userId}' and duc_campaigninternaltype eq 100000004 and duc_campaignstatus eq 2 and duc_fromdate le ${today} and duc_todate ge ${today}&$select=new_inspectioncampaignid,new_name&$top=1`;

            const activePatrolResults = await ctx.webAPI.retrieveMultipleRecords(
                "new_inspectioncampaign",
                activePatrolQuery
            );

            if (activePatrolResults.entities.length > 0) {
                // Found active patrol, show End Patrol button
                setState(prev => ({
                    ...prev,
                    patrolStatus: 'end',
                    activePatrolId: activePatrolResults.entities[0].new_inspectioncampaignid,
                    activePatrolName: activePatrolResults.entities[0].new_name
                }));
                return;
            }

            // Check for available patrol to start (status = 1 or 100000004)
            const availablePatrolQuery = `?$filter=_owninguser_value eq '${userId}' and duc_campaigninternaltype eq 100000004 and (duc_campaignstatus eq 1 or duc_campaignstatus eq 100000004) and duc_fromdate le ${today} and duc_todate ge ${today}&$select=new_inspectioncampaignid,new_name&$top=1`;

            const availablePatrolResults = await ctx.webAPI.retrieveMultipleRecords(
                "new_inspectioncampaign",
                availablePatrolQuery
            );

            if (availablePatrolResults.entities.length > 0) {
                // Found available patrol, show Start Patrol button
                setState(prev => ({
                    ...prev,
                    patrolStatus: 'start',
                    activePatrolId: availablePatrolResults.entities[0].new_inspectioncampaignid,
                    activePatrolName: availablePatrolResults.entities[0].new_name
                }));
            } else {
                // No patrol available
                setState(prev => ({
                    ...prev,
                    patrolStatus: 'none',
                    activePatrolId: undefined,
                    activePatrolName: undefined
                }));
            }

        } catch (error) {
            console.error("Error checking patrol status:", error);
            setState(prev => ({
                ...prev,
                patrolStatus: 'none',
                activePatrolId: undefined,
                activePatrolName: undefined
            }));
        }
    }, [props.context]);

    const loadTodaysCounts = async (ctx: any, userId: string): Promise<{ completedToday: number; remainingToday: number; campaignsToday: number }> => {
        const CACHE_KEY = `MOCI_User_ResourceID_${userId}`;

        try {
            let completedToday = 0;
            let remainingToday = 0;
            let campaignsToday = 0;

            try {
                debugger;
                let resourceId: string | null = localStorage.getItem(CACHE_KEY);

                if (!resourceId) {
                    const userQuery = `?$select=_duc_bookableresourceid_value&$filter=systemuserid eq '${userId}'`;
                    const userResults = await ctx.webAPI.retrieveMultipleRecords("systemuser", userQuery);

                    const retrievedResourceId = userResults.entities.length > 0
                        ? userResults.entities[0]["_duc_bookableresourceid_value"]
                        : null;

                    if (retrievedResourceId) {
                        resourceId = retrievedResourceId;
                        localStorage.setItem(CACHE_KEY, resourceId?.toString() ?? "");
                        console.log("Resource ID fetched and cached:", resourceId);
                    } else {
                        console.warn(`User ${userId} is not linked to a Bookable Resource.`);
                        return { completedToday, remainingToday, campaignsToday };
                    }
                } else {
                    console.log("Resource ID retrieved from cache:", resourceId);
                }

                if (!resourceId) {
                    return { completedToday, remainingToday, campaignsToday };
                }

                const completedQuery = `?$select=msdyn_workorderid&$filter=_duc_assignedresource_value eq '${resourceId}' and Microsoft.Dynamics.CRM.Today(PropertyName='duc_completiondate')`;
                const completedResults = await ctx.webAPI.retrieveMultipleRecords("msdyn_workorder", completedQuery);
                completedToday = completedResults.entities.length;
                console.log("Completed Work Orders Count:", completedToday);

                const bookingStatusGuid = 'f16d80d1-fd07-4237-8b69-187a11eb75f9';
                const remainingQuery = `?$select=bookableresourcebookingid&$filter=_resource_value eq '${resourceId}' and _bookingstatus_value eq '${bookingStatusGuid}' and Microsoft.Dynamics.CRM.Today(PropertyName='starttime')`;
                const remainingResults = await ctx.webAPI.retrieveMultipleRecords("bookableresourcebooking", remainingQuery);
                remainingToday = remainingResults.entities.length;
                console.log("Remaining Bookings Count:", remainingToday);

                return { completedToday, remainingToday, campaignsToday };

            } catch (error) {
                console.error("Error retrieving counts:", error);
                return { completedToday: 0, remainingToday: 0, campaignsToday: 0 };
            }

        } catch (e) {
            console.log("Failed to load today's counts:", e);
            return { completedToday: 0, remainingToday: 0, campaignsToday: 0 };
        }
    };

    const loadUserData = React.useCallback(async (): Promise<void> => {
        const ctx: any = props.context;
        try {
            if (!ctx) return;
            const userSettings = ctx.userSettings;
            const userId = (userSettings.userId ?? "").replace("{", "").replace("}", "");
            let res: any = null;

            if (isOffline()) {
                const results = await ctx.utils.executeOffline(
                    "systemuser",
                    `?$filter=systemuserid eq ${userId}&$select=duc_usernamearabic`
                );
                res = results?.entities?.[0] ?? null;
            } else {
                res = await ctx.webAPI.retrieveRecord(
                    "systemuser",
                    userId,
                    "?$select=duc_usernamearabic"
                );
            }

            if (res) {
                const username = res.duc_usernamearabic ?? userSettings?.userName ?? "Inspector";
                const { completedToday, remainingToday, campaignsToday } = await loadTodaysCounts(ctx, userId);

                setState(prev => ({
                    ...prev,
                    pendingTodayBookings: remainingToday,
                    completedTodayWorkorders: completedToday,
                    TodayCampaigns: campaignsToday,
                    userName: username,
                }));

                localStorage.setItem(
                    "MOCI_userCounts",
                    JSON.stringify({ userName: username })
                );

                // Check organization unit after loading user data
                if (!isOffline()) {
                    await checkOrganizationUnit(ctx, userId);
                }
            }
        } catch (e) {
            console.warn("User data load failed, using cache.", e);
            restoreCache();
        }
    }, [props.context, checkOrganizationUnit]);

    const handleOpenAdvancedSearch = (): void => {
        setState(prev => ({ ...prev, showAdvancedSearch: true }));
    };

    const handleCloseAdvancedSearch = (): void => {
        setState(prev => ({ ...prev, showAdvancedSearch: false }));
    };

    const restoreCache = (): void => {
        const cached = localStorage.getItem("MOCI_userCounts");
        if (cached) {
            try {
                const obj = JSON.parse(cached);
                setState(prev => ({
                    ...prev,
                    remainingToday: obj.remainingToday,
                    completedTodayWorkorders: obj.completedTodayWorkorders,
                    TodayCampaigns: obj.TodayCampaigns,
                    userName: obj.userName ?? "Inspector"
                }));
            } catch {
                // ignore corrupted cache
            }
        }
    };

    React.useEffect(() => {
        void loadUserData();
        void checkPatrolStatus();
        restoreCache();
    }, [loadUserData, checkPatrolStatus]);

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
        const ctx: any = props.context;
        const searchKey = `MOCI_lastSearch_${text.toLowerCase()}`;

        try {
            let results: any = null;

            if (isOffline()) {
                results = await ctx.utils.executeOffline("account", `?$top=100`);
                if (text) {
                    const lowerText = text.toLowerCase();
                    results.entities = results.entities.filter((entity: any) => {
                        const entityName = entity.name as string | undefined;
                        return entityName?.toLowerCase().startsWith(lowerText) ?? false;
                    });
                }
            } else {
                results = await ctx.webAPI.retrieveMultipleRecords("account", `?$top=100`);
                if (text) {
                    const lowerText = text.toLowerCase();
                    results.entities = results.entities.filter((entity: any) => {
                        const entityName = entity.name as string | undefined;
                        return entityName?.toLowerCase().startsWith(lowerText) ?? false;
                    });
                }
            }

            const entities = results?.entities ?? [];

            if (entities.length === 0) {
                setState(prev => ({ ...prev, message: strings.NoResults }));
            } else if (entities.length === 1 && !isOffline()) {
                void ctx.navigation.navigateTo({
                    pageType: "entityrecord",
                    entityName: "account",
                    entityId: entities[0].accountid
                });
            } else {
                setState(prev => ({
                    ...prev,
                    showResults: true,
                    searchResults: entities,
                    message: undefined
                }));
            }

            localStorage.setItem(searchKey, JSON.stringify(results));
        } catch (err: any) {
            console.error(err);
            const cached = localStorage.getItem(searchKey);
            if (cached) {
                const parsed = JSON.parse(cached);
            } else {
                setState(prev => ({
                    ...prev,
                    message: `${err?.message ?? err}`
                }));
            }
        } finally {
            setState(prev => ({ ...prev, isLoading: false }));
        }
    };

    const onRecordClick = (accountId: string): void => {
        const ctx: any = props.context;
        setState(prev => ({ ...prev, showResults: false }));

        void ctx.navigation.navigateTo({
            pageType: "entityrecord",
            entityName: "account",
            entityId: accountId
        });
    };

    const onCloseResults = (): void => {
        setState(prev => ({ ...prev, showResults: false }));
    };

    const openScheduled = (): void => {
        const ctx: any = props.context;
        void ctx.navigation.navigateTo({
            pageType: "entitylist",
            entityName: "bookableresourcebooking",
            viewId: "4073baca-cc5f-e611-8109-000d3a146973"
        });
    };

    const openAllScheduled = (): void => {
        const ctx: any = props.context;
        void ctx.navigation.navigateTo({
            pageType: "entitylist",
            entityName: "bookableresourcebooking",
            viewId: "f52b7a9b-cae8-f011-8406-6045bd9c20cb"
        });
    };

    const openPendingWorkorders = (): void => {
        const ctx: any = props.context;
        if (isOffline()) {
            setState(prev => ({ ...prev, message: strings.OfflineNavigationBlocked }));
            return;
        }
        void ctx.navigation.navigateTo({
            pageType: "entitylist",
            entityName: "msdyn_workorder",
            viewId: "bee0efc7-40e4-f011-8406-6045bd9c224c"
        });
    };

    const closedWorkorders = (): void => {
        const ctx: any = props.context;
        if (isOffline()) {
            setState(prev => ({ ...prev, message: strings.OfflineNavigationBlocked }));
            return;
        }
        void ctx.navigation.navigateTo({
            pageType: "entitylist",
            entityName: "msdyn_workorder",
            viewId: "aad92cc9-00c6-f011-8544-000d3a2274a5"
        });
    };

    const startInspection = (): void => {
        const ctx: any = props.context;
        if (isOffline()) {
            setState(prev => ({ ...prev, message: strings.OfflineQuickCreateBlocked }));
            return;
        }

        // Prepare default values
        const defaultValues: any = {};

        // If we have an active patrol campaign, set it as default
        if (state.activePatrolId && state.activePatrolName) {
            defaultValues.new_campaign = [
                {
                    id: state.activePatrolId,
                    name: state.activePatrolName,
                    entityType: "new_inspectioncampaign"
                }
            ];
        }

        void ctx.navigation.openForm(
            { entityName: "msdyn_workorder", useQuickCreateForm: true },
            defaultValues
        );
    };

    const startPatrol = async (): Promise<void> => {
        const ctx: any = props.context;

        if (isOffline()) {
            setState(prev => ({ ...prev, message: strings.OfflineNavigationBlocked }));
            return;
        }

        if (!state.activePatrolId) {
            setState(prev => ({ ...prev, message: "No patrol campaign found" }));
            return;
        }

        try {
            setState(prev => ({ ...prev, isLoading: true }));

            // Update campaign status to 2 (In Progress)
            const updateData = {
                duc_campaignstatus: 2
            };

            await ctx.webAPI.updateRecord(
                "new_inspectioncampaign",
                state.activePatrolId,
                updateData
            );

            // Update state to show End Patrol button
            setState(prev => ({
                ...prev,
                patrolStatus: 'end',
                isLoading: false,
                message: "Patrol started successfully"
            }));

            // Clear message after 3 seconds
            setTimeout(() => {
                setState(prev => ({ ...prev, message: undefined }));
            }, 3000);

        } catch (error) {
            console.error("Error starting patrol:", error);
            setState(prev => ({
                ...prev,
                isLoading: false,
                message: "Failed to start patrol"
            }));
        }
    };

    const endPatrol = async (): Promise<void> => {
        const ctx: any = props.context;

        if (isOffline()) {
            setState(prev => ({ ...prev, message: strings.OfflineNavigationBlocked }));
            return;
        }

        try {
            setState(prev => ({ ...prev, isLoading: true }));

            const userSettings = ctx.userSettings;
            const userId = (userSettings.userId ?? "").replace("{", "").replace("}", "");
            const today = new Date().toISOString().split('T')[0];

            // Query for active patrol
            const activePatrolQuery = `?$filter=_owninguser_value eq '${userId}' and duc_campaigninternaltype eq 100000004 and duc_campaignstatus eq 2 and duc_fromdate le ${today} and duc_todate ge ${today}&$top=1`;

            const activePatrolResults = await ctx.webAPI.retrieveMultipleRecords(
                "new_inspectioncampaign",
                activePatrolQuery
            );

            if (activePatrolResults.entities.length === 0) {
                setState(prev => ({
                    ...prev,
                    isLoading: false,
                    message: "No active patrol found"
                }));
                return;
            }

            const patrolId = activePatrolResults.entities[0].new_inspectioncampaignid;

            // Update campaign status to 100000004 (Completed)
            const updateData = {
                duc_campaignstatus: 100000004
            };

            await ctx.webAPI.updateRecord(
                "new_inspectioncampaign",
                patrolId,
                updateData
            );

            // Check if there's another patrol available
            await checkPatrolStatus();

            setState(prev => ({
                ...prev,
                isLoading: false,
                message: "Patrol ended successfully"
            }));

            // Clear message after 3 seconds
            setTimeout(() => {
                setState(prev => ({ ...prev, message: undefined }));
            }, 3000);

        } catch (error) {
            console.error("Error ending patrol:", error);
            setState(prev => ({
                ...prev,
                isLoading: false,
                message: "Failed to end patrol"
            }));
        }
    };

    const { searchText, pendingTodayBookings, completedTodayWorkorders, TodayCampaigns, userName, message, isLoading, showResults, searchResults, patrolStatus, isNaturalReserve, isWildlifeSection } = state;
    const isActionDisabled = isLoading;

    // UPDATED BUTTON VISIBILITY LOGIC
    // Always show Scheduled Inspections
    const showScheduledInspections = true;

    // Start Inspection: Show for Wildlife sections ALWAYS, OR for non-Natural Reserve units
    // NEVER show for Natural Reserve
    const showStartInspection = isWildlifeSection || (!isNaturalReserve && !isWildlifeSection);

    // Multi Type Inspection: Show ONLY for Natural Reserve, with same patrol conditions
    // When patrol status is 'start': DON'T show (user needs to start patrol first)
    // When patrol status is 'end': SHOW (patrol is active)
    // When patrol status is 'none': SHOW (no patrol system active)
    const showMultiTypeInspection = isNaturalReserve && (patrolStatus === 'end' || patrolStatus === 'none');

    // Start/End Patrol buttons - show for all org units based on patrol status
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
        position: 'fixed',
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
                                style: { border: 'none', width: '90%', outline: 'none', direction: isRTL ? 'rtl' : 'ltr', background: 'none' }
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
                    )
                )
            ),
            // Action Buttons
            React.createElement(
                "div",
                { style: STYLES.actions },

                // Scheduled Inspections button - always show
                showScheduledInspections && React.createElement(
                    "button",
                    {
                        onClick: openAllScheduled,
                        disabled: isActionDisabled,
                        style: getButtonStyle({ ...STYLES.btnBlueDark, ...STYLES.textWhite, ...STYLES.rounded4, ...STYLES.p3, ...STYLES.dFlex, ...STYLES.alignItemsCenter, ...STYLES.justifyContentCenter, ...STYLES.gap3 })
                    },
                    React.createElement("img", { src: calenderDataUri }),
                    React.createElement("span", null, strings.ScheduledInspections)
                ),

                // Start Inspection button - show for Wildlife sections OR non-Natural Reserve, NEVER for Natural Reserve
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

                // Multi Type Inspection button - ONLY for Natural Reserve
                showMultiTypeInspection && React.createElement(
                    "button",
                    {
                        onClick: () => setState(prev => ({ ...prev, showMultiTypeInspection: true })),
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
                        React.createElement("span", { key: "text" }, strings.StartMultiTypeInspection)
                    ]
                ),

                // Start Patrol button - show for all org units when patrol is available
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

                // End Patrol button - show for all org units when patrol is active
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
        React.createElement(AdvancedSearch, {
            context: props.context,
            isOpen: state.showAdvancedSearch,
            onClose: handleCloseAdvancedSearch,
            onSearchResults: handleAdvancedSearchResults
        }),
        // MultiTypeInspection Modal
        state.showMultiTypeInspection && React.createElement(MultiTypeInspection, {
            context: props.context,
            isOpen: true,
            onClose: () => setState(prev => ({ ...prev, showMultiTypeInspection: false })),
            activePatrolId: state.activePatrolId,
            activePatrolName: state.activePatrolName,
            incidentTypeId: state.incidentTypeId,
            incidentTypeName: state.incidentTypeName,
            unknownAccountId: state.unknownAccountId,
            unknownAccountName: state.unknownAccountName
        } as any),
        // Search Results Modal
        showResults && React.createElement(SearchResults, {
            results: searchResults,
            onRecordClick: onRecordClick,
            onClose: onCloseResults,
            context: props.context
        })
    );
};