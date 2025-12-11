/* eslint-disable */
import * as React from "react";
import { IInputs } from "../generated/ManifestTypes";
import { logoBase64 } from "../images/logo64";
import { startSvgContent } from "../images/start";
import { calender } from "../images/calender";
import { arrow } from "../images/arrow";
import { hr } from "../images/hr";
import { min } from "../images/min";
import { clock_bg } from "../images/clock_bg";



// Define the interface for the component's internal state.
interface State {
    searchText: string;
    pendingTodayBookings: number | null;
    completedTodayWorkorders: number | null;
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
    logout: string;
    loggingout: string;
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
    h5: { fontSize: 20, margin: 0, fontWeight: "500" as const }

};
const CLOCK_STYLES = {
    clockContainer: {
        width: 300,
        height: 300,
        position: "relative" as const,
    },
    clock: {
        width: 300,
        height: 300,
        position: "absolute" as const,
        top: "50%",
        left: "50%",
        transform: "translate(-50%, -50%)",
        borderRadius: "50%",
        display: "flex",
        justifyContent: "center",
        alignItems: "center",
        backgroundSize: "cover"
    },
    clockCenter: {
        content: "",
        width: 5,
        height: 5,
        backgroundColor: "#fff",
        borderRadius: "50%",
        border: "2px solid #fff",
        zIndex: 4,
        position: "absolute" as const
    },
    hr: {
        width: 150,
        height: 150,
        position: "absolute" as const,
        display: "flex",
        justifyContent: "center",
    },
    hrBefore: {
        content: "",
        position: "absolute" as const,
        width: 20,
        height: 80,
        zIndex: 1,
        backgroundSize: "cover"
    },
    min: {
        width: 200,
        height: 215,
        position: "absolute" as const,
        display: "flex",
        justifyContent: "center",
    },
    minBefore: {
        content: "",
        position: "absolute" as const,
        width: 20,
        height: 120,
        zIndex: 2,
        backgroundSize: "contain",
        backgroundRepeat: "no-repeat"
    },
    sec: {
        width: 220,
        height: 210,
        position: "absolute" as const,
        display: "flex",
        justifyContent: "center",
    },
    secBefore: {
        content: "",
        position: "absolute" as const,
        width: 3,
        height: 120,
        backgroundColor: "#8A1538",
        zIndex: 3,
        borderRadius: "50%"
    }
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
            logout: ctx.resources.getString("Logout"),
            loggingout: ctx.resources.getString("loggingout"),
            StartInspection: ctx.resources.getString("StartInspection"),
            SearchPlaceholder: ctx.resources.getString("SearchPlaceholder"),
            NoResults: ctx.resources.getString("NoResults"),
            Loading: ctx.resources.getString("Loading"),
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
        userName: strings.Loading,
        message: undefined,
        isLoading: false,
        searchResults: [],
        showResults: false,
        showAdvancedSearch: false,
    });

    const startDataUri = "data:image/svg+xml;base64," + btoa(startSvgContent);
    const calenderDataUri = "data:image/svg+xml;base64," + btoa(calender);
    const arrowSvg = "data:image/svg+xml;base64," + btoa(arrow);
    const hrSvg = "data:image/svg+xml;base64," + btoa(hr);
    const minSvg = "data:image/svg+xml;base64," + btoa(min);
    const clockBgSvg = "data:image/svg+xml;base64," + btoa(clock_bg);

    const hrRef = React.useRef<HTMLDivElement>(null);
    const minRef = React.useRef<HTMLDivElement>(null);
    const secRef = React.useRef<HTMLDivElement>(null);

    const isOffline = (): boolean => {
        const ctx: any = props.context;
        return ctx.mode?.isInOfflineMode === true;
    };

    // Clock update logic
    React.useEffect(() => {
        const updateClock = (): void => {
            const date = new Date();
            const hr = date.getHours() * 30;
            const min = date.getMinutes() * 6;
            const sec = date.getSeconds() * 6;

            if (hrRef.current) {
                hrRef.current.style.transform = `rotateZ(${hr + min / 12}deg)`;
            }
            if (minRef.current) {
                minRef.current.style.transform = `rotateZ(${min}deg)`;
            }
            if (secRef.current) {
                secRef.current.style.transform = `rotateZ(${sec}deg)`;
            }
        };

        updateClock();
        const interval = setInterval(updateClock, 1000);

        return () => clearInterval(interval);
    }, []);

    const handleBack = (): void => {
        // Implement back navigation
        console.log("Back clicked");
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

                setState(prev => ({
                    ...prev,
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



    const restoreCache = (): void => {
        const cached = localStorage.getItem("MOCI_userCounts");
        if (cached) {
            try {
                const obj = JSON.parse(cached);
                setState(prev => ({
                    ...prev,
                    remainingToday: obj.remainingToday,
                    completedTodayWorkorders: obj.completedTodayWorkorders,
                    userName: obj.userName ?? "Inspector"
                }));
            } catch {
                // ignore corrupted cache
            }
        }
    };

    React.useEffect(() => {
        void loadUserData();
    }, [loadUserData]);




    const earlyLogout = async (): Promise<void> => {
        debugger;

        const ctx: any = props.context;

        if (isOffline()) {
            setState(prev => ({ ...prev, message: strings.OfflineNavigationBlocked }));
            return;
        }

        // Use PCF context utils instead of Xrm.Utility
        Xrm.Utility.showProgressIndicator(strings.loggingout);
        await new Promise(resolve => setTimeout(resolve, 100));

        try {
            // Show loading indicator
            setState(prev => ({ ...prev, isLoading: true }));

            // Step 1: Get current logged-in user's ID
            const userId = ctx.userSettings.userId.replace(/[{}]/g, "");

            // Step 2: Get bookable resource for current user
            const resourceResult = await ctx.webAPI.retrieveMultipleRecords(
                "bookableresource",
                `?$select=bookableresourceid&$filter=_userid_value eq ${userId}`
            );

            if (!resourceResult.entities || resourceResult.entities.length === 0) {
                setState(prev => ({
                    ...prev,
                    message: "No bookable resource found for current user.",
                    isLoading: false
                }));
                ctx.utils.closeProgressIndicator();
                return;
            }

            const resourceId = resourceResult.entities[0].bookableresourceid;

            // Step 3: Get "Scheduled" booking status ID
            const bookingStatusGuid = 'f16d80d1-fd07-4237-8b69-187a11eb75f9';
            const closedStatusGuid = '0adbf4e6-86cc-4db0-9dbb-51b7d1ed4020';

            // Fixed query - use Today() function correctly
            const remainingQuery = `?$select=bookableresourcebookingid&$filter=_resource_value eq ${resourceId} and _bookingstatus_value eq ${bookingStatusGuid} and Microsoft.Dynamics.CRM.Today(PropertyName='starttime')`;

            const scheduledBookings = await ctx.webAPI.retrieveMultipleRecords(
                "bookableresourcebooking",
                remainingQuery
            );

            if (!scheduledBookings.entities || scheduledBookings.entities.length === 0) {
                setState(prev => ({
                    ...prev,
                    message: "No scheduled bookings found for today.",
                    isLoading: false
                }));
                // ctx.utils.closeProgressIndicator();
                return;
            }

            // Step 4: Update all today's scheduled bookings to canceled status
            const updatePromises = scheduledBookings.entities.map((booking: any) =>
                ctx.webAPI.updateRecord(
                    "bookableresourcebooking",
                    booking.bookableresourcebookingid,
                    {
                        "BookingStatus@odata.bind": `/bookingstatuses(${closedStatusGuid})`
                    }
                )
            );

            await Promise.all(updatePromises);

            // Step 5: Show success message
            setState(prev => ({
                ...prev,
                message: `Successfully closed ${scheduledBookings.entities.length} scheduled booking(s) for today.`,
                isLoading: false
            }));

        } catch (error: any) {
            console.error("Error closing today's scheduled bookings:", error);
            setState(prev => ({
                ...prev,
                message: `Error: ${error.message || error}`,
                isLoading: false
            }));
        } finally {
            // Always close progress indicator in finally block
            Xrm.Utility.closeProgressIndicator();
        }
    };




    const { userName, isLoading } = state;
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
            // Main Content with Clock
            React.createElement(
                "div",
                { style: { ...STYLES.flexGrow1, ...STYLES.p3 } },
                // Back Button
                React.createElement(
                    "a",
                    {
                        href: "#",
                        onClick: (e: React.MouseEvent) => { e.preventDefault(); handleBack(); },
                        style: {
                            ...STYLES.dFlex,
                            ...STYLES.gap3,
                            ...STYLES.mb3,
                            ...STYLES.alignItemsCenter,
                            textDecoration: "none",
                            color: "inherit"
                        }
                    },
                    React.createElement(
                        "div",
                        { style: { ...STYLES.border, ...STYLES.rounded3, ...STYLES.p3, } },
                        React.createElement("img", { src: arrowSvg, alt: "Back" })
                    ),
                ),
                // Clock Container
                React.createElement(
                    "div",
                    { style: { ...STYLES.dFlex, ...STYLES.alignItemsCenter, ...STYLES.justifyContentCenter, minHeight: 350 } },
                    React.createElement(
                        "div",
                        { style: CLOCK_STYLES.clockContainer },
                        React.createElement(
                            "div",
                            { style: { ...CLOCK_STYLES.clock, backgroundImage: `url(${clockBgSvg})` } },
                            // Center dot
                            React.createElement("div", { style: CLOCK_STYLES.clockCenter }),
                            // Hour hand
                            React.createElement(
                                "div",
                                { ref: hrRef, style: CLOCK_STYLES.hr },
                                React.createElement("img", { style: CLOCK_STYLES.hrBefore, src: hrSvg })
                            ),
                            // Minute hand
                            React.createElement(
                                "div",
                                { ref: minRef, style: CLOCK_STYLES.min },
                                React.createElement("img", { style: CLOCK_STYLES.minBefore, src: minSvg })
                            ),
                            // Second hand
                            React.createElement(
                                "div",
                                { ref: secRef, style: CLOCK_STYLES.sec },
                                React.createElement("div", { style: CLOCK_STYLES.secBefore })
                            )
                        )
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
                        onClick: earlyLogout,
                        disabled: isActionDisabled,
                        style: getButtonStyle({ ...STYLES.btnBlueDark, ...STYLES.textWhite, ...STYLES.rounded4, ...STYLES.p3, ...STYLES.dFlex, ...STYLES.alignItemsCenter, ...STYLES.justifyContentCenter, ...STYLES.gap3 })
                    },
                    React.createElement("img", { src: calenderDataUri }),
                    React.createElement("span", null, strings.logout)
                ),
            )
        ),
    );
};