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
            OfflineQuickCreateBlocked: ctx.resources.getString("OfflineQuickCreateBlocked")
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
    });

    const startDataUri = "data:image/svg+xml;base64," + btoa(startSvgContent);
    const magnifyDataUri = "data:image/svg+xml;base64," + btoa(magnify);
    const calenderDataUri = "data:image/svg+xml;base64," + btoa(calender);
    const checkDataUri = "data:image/svg+xml;base64," + btoa(check);
    const clockDataUri = "data:image/svg+xml;base64," + btoa(clock);
    const filtersDataUri = "data:image/svg+xml;base64," + btoa(filters);

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
            }
        } catch (e) {
            console.warn("User data load failed, using cache.", e);
            restoreCache();
        }
    }, [props.context]);

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
        restoreCache();
    }, [loadUserData]);

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
        if (isOffline()) {
            setState(prev => ({ ...prev, message: strings.OfflineNavigationBlocked }));
            return;
        }
        void ctx.navigation.navigateTo({
            pageType: "entitylist",
            entityName: "bookableresourcebooking",
            viewId: "4073baca-cc5f-e611-8109-000d3a146973"
        });
    };
    const openScheduledCampaigns = (): void => {
        const ctx: any = props.context;
        if (isOffline()) {
            setState(prev => ({ ...prev, message: strings.OfflineNavigationBlocked }));
            return;
        }
        void ctx.navigation.navigateTo({
            pageType: "entitylist",
            entityName: "new_inspectioncampaign",
            viewId: "0628a5aa-88d7-f011-8406-7c1e524df347"
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
        void ctx.navigation.openForm({ entityName: "msdyn_workorder", useQuickCreateForm: true }, {});
    };

    const { searchText, pendingTodayBookings, completedTodayWorkorders, TodayCampaigns, userName, message, isLoading, showResults, searchResults } = state;
    const isActionDisabled = isLoading;

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
        // margin: '2px',
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
                // Messages
                message && React.createElement("div", { style: { color: "red", marginTop: 10, marginBottom: 10, background: 'none', padding: 8, borderRadius: 4 } }, message),
                isLoading && React.createElement("div", { style: { color: "blue", marginTop: 10, marginBottom: 10, background: 'none', padding: 8, borderRadius: 4 } }, strings.Loading),
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

                    //patrols card
                    React.createElement(
                        "div",
                        {
                            onClick: openScheduledCampaigns,
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
                            React.createElement("h6", { style: STYLES.h6 }, strings.TodaysPatrols)
                        ),
                        React.createElement("div", { style: { ...STYLES.textBrown, ...STYLES.h1 } }, TodayCampaigns ?? "...")
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
                React.createElement(
                    "button",
                    {
                        onClick: openScheduled,
                        disabled: isActionDisabled,
                        style: getButtonStyle({ ...STYLES.btnBlueDark, ...STYLES.textWhite, ...STYLES.rounded4, ...STYLES.p3, ...STYLES.dFlex, ...STYLES.alignItemsCenter, ...STYLES.justifyContentCenter, ...STYLES.gap3 })
                    },
                    React.createElement("img", { src: calenderDataUri }),
                    React.createElement("span", null, strings.ScheduledInspections)
                ),
                React.createElement(
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
                )
            )
        ),
        React.createElement(AdvancedSearch, {
            context: props.context,
            isOpen: state.showAdvancedSearch,
            onClose: handleCloseAdvancedSearch,
            onSearchResults: handleAdvancedSearchResults
        }),
        // Search Results Modal
        showResults && React.createElement(SearchResults, {
            results: searchResults,
            onRecordClick: onRecordClick,
            onClose: onCloseResults,
            context: props.context  // Add this line
        })
    );
};