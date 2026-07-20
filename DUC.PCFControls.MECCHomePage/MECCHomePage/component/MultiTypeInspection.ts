/* eslint-disable */
import * as React from "react";
import {
    WorkOrderHelpers,
    CampaignHelpers,
    IncidentTypeHelpers,
    InitCache,
    ProjectHelpers,
} from "../helpers";
import { IProjectDetail } from "../helpers/ProjectHelpers";

interface IMultiTypeInspectionProps {
    context: ComponentFramework.Context<any>;
    onClose?: () => void;
    isOpen?: boolean;
    activePatrolId?: string;
    activePatrolName?: string;
    incidentTypeId?: string;
    incidentTypeName?: string;
    unknownAccountId?: string;
    unknownAccountName?: string;
    siteAccountId?: string;
    siteAccountName?: string;
    organizationUnitId?: string;
    organizationUnitName?: string;
    defaultInspectionType?: number;
    lockInspectionType?: boolean;
}

interface IMultiTypeInspectionState {
    isRTL: boolean;
    inspectionTypes: Array<{
        value: number;
        label: string;
        accountTypeId: string;
        orgUnitAccountTypeId?: string;
    }>;
    selectedInspectionType: number | null;
    qataryId: string;
    name: string;
    crNumber: string;
    cpNumber: string;
    crCpToggle: 'cr' | 'cp';
    registrationNumber: string;
    id: string;
    carColor: string;
    vehicleBrand: number | null;
    vehicleBrands: Array<{ value: number; label: string }>;
    boatNumber: string;
    projectName: string;
    requestPermitNumber: string;
    locationDetails: string;
    loading: boolean;
    error: string | null;
    accountTypeRecord: any | null;
    showCampaignIncidentPopup: boolean;
    selectedCampaignId?: string;
    selectedCampaignName?: string;
    selectedIncidentTypeId?: string;
    selectedIncidentTypeName?: string;
    campaigns: Array<{ id: string; name: string }>;
    incidentTypes: Array<{ id: string; name: string }>;
    isAnonymous: boolean;
    popupShowCampaign: boolean;
    popupShowIncidentType: boolean;
    // NEW: incident type read-only when auto-filled from campaign
    incidentTypeReadOnly: boolean;
    // NEW: map campaign ID → incident type to avoid re-retrieves
    campaignIncidentTypeMap: Record<string, { id: string; name: string }>;
    // NEW: whether to show campaign field based on user's department setting
    shouldShowCampaignField: boolean;
    // Project selection (type 16)
    showProjectSelectionPopup: boolean;
    nearbyProjects: IProjectDetail[];
    selectedProjectId: string | null;
    projectSearchLoading: boolean;
}

interface LocalizedStrings {
    StartMultiTypeInspection: string;
    InspectionType: string;
    QataryID: string;
    Name: string;
    CRNumber: string;
    CPNumber: string;
    MonourNumber: string;
    ID: string;
    CarColor: string;
    VehicleBrand: string;
    Start: string;
    Close: string;
    Loading: string;
    PleaseSelectInspectionType: string;
    PleaseEnterRequiredFields: string;
    Error: string;
    chooseInspectionType: string;
    SelectCampaign: string;
    SelectIncidentType: string;
    Campaign: string;
    IncidentType: string;
    Continue: string;
    Anonymous: string;
    CreatingAccount: string;
    CreatingWorkOrder: string;
    CreatingBooking: string;
    ScanBarcode: string;
    Clear: string;
    PlateNumber: string;
    RegistrationNumber: string;
    ContactAdministrator: string;
    BoatNumber: string;
    QataryIDMustBe11Digits: string;
    ProjectName: string;
    RequestPermitNumber: string;
    LocationDetails: string;
    SelectProject: string;
    Back: string;
    Select: string;
    NoProjectsFound: string;
    SearchingProjects: string;
    ProjectNumber: string;
    AccountNameLabel: string;
}

// Cache constants
const INSPECTION_TYPES_CACHE_KEY = "MOCI_OrgUnit_InspectionTypes_Cache";
const VEHICLE_TYPES_CACHE_KEY = "MOCI_VehicleTypes_Cache";
const CAMPAIGNS_CACHE_KEY = "MOCI_Campaigns_Cache";
const INCIDENT_TYPES_CACHE_KEY = "MOCI_IncidentTypes_Cache";
const CACHE_DURATION = 60_000; // 1 minute

interface CacheData<T> {
    data: T;
    timestamp: number;
}

// =====================================================================
// FLUENT UI STYLE TOKENS
// =====================================================================

const FLUENT = {
    fontFamily:
        '"Segoe UI", "Segoe UI Web (West European)", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif',
    colorPrimary: "#0078d4",
    colorPrimaryHover: "#106ebe",
    colorNeutralDark: "#201f1e",
    colorNeutralPrimary: "#323130",
    colorNeutralSecondary: "#605e5c",
    colorNeutralLight: "#edebe9",
    colorNeutralLighter: "#f3f2f1",
    colorErrorPrimary: "#a4262c",
    colorErrorBackground: "#fde7e9",
    colorWhite: "#ffffff",
    borderRadius: 4,
    borderRadiusModal: 8,
    shadowModal: "0 8px 32px rgba(0, 0, 0, 0.14)",
    overlay: "rgba(0, 0, 0, 0.4)",
    focusOutline: "2px solid #0078d4",
    transitionFast: "0.1s ease",
} as const;

export class MultiTypeInspection extends React.Component<
    IMultiTypeInspectionProps,
    IMultiTypeInspectionState
> {
    private strings: LocalizedStrings;
    private xrm: Xrm.XrmStatic;
    // OPTIMIZATION: Cache account ID to avoid re-searching in handleContinueWithSelections
    private pendingAccountId: string | null = null;

    constructor(props: IMultiTypeInspectionProps) {
        super(props);

        const userSettings = (props.context as any).userSettings;
        const rtlLanguages = [
            1025, 1037, 1054, 1056, 1065, 1068, 1069, 1101, 1114, 1119,
        ];
        const isRTL = rtlLanguages.includes(userSettings?.languageId);

        this.xrm = (window.parent as any).Xrm || (window as any).Xrm;

        // Load localized strings
        this.strings = {
            StartMultiTypeInspection: props.context.resources.getString(
                "StartMultiTypeInspection",
            ),
            InspectionType: props.context.resources.getString("InspectionType"),
            QataryID: props.context.resources.getString("QataryID"),
            Name: props.context.resources.getString("Name"),
            CRNumber: props.context.resources.getString("CRNumber"),
            CPNumber: props.context.resources.getString("CPNumber") || "CP Number:",
            MonourNumber:
                props.context.resources.getString("MonourNumber") || "Monour Number",
            ID: props.context.resources.getString("ID"),
            CarColor: props.context.resources.getString("CarColor") || "Car Color",
            VehicleBrand:
                props.context.resources.getString("VehicleBrand") || "Vehicle Brand",
            Start: props.context.resources.getString("Start"),
            Close: props.context.resources.getString("Close"),
            Loading: props.context.resources.getString("Loading"),
            PleaseSelectInspectionType: props.context.resources.getString(
                "PleaseSelectInspectionType",
            ),
            PleaseEnterRequiredFields: props.context.resources.getString(
                "PleaseEnterRequiredFields",
            ),
            Error: props.context.resources.getString("Error"),
            chooseInspectionType: props.context.resources.getString(
                "chooseInspectionType",
            ),
            SelectCampaign:
                props.context.resources.getString("SelectCampaign") ||
                "Select Campaign",
            SelectIncidentType:
                props.context.resources.getString("SelectIncidentType") ||
                "Select Incident Type",
            Campaign: props.context.resources.getString("Campaign") || "Campaign",
            IncidentType:
                props.context.resources.getString("IncidentType") || "Incident Type",
            Continue: props.context.resources.getString("Continue") || "Continue",
            Anonymous: props.context.resources.getString("Anonymous") || "Anonymous",
            CreatingAccount:
                props.context.resources.getString("CreatingAccount") ||
                "Creating Account...",
            CreatingWorkOrder:
                props.context.resources.getString("CreatingWorkOrder") ||
                "Creating Work Order...",
            CreatingBooking:
                props.context.resources.getString("CreatingBooking") ||
                "Creating Booking...",
            ScanBarcode:
                props.context.resources.getString("ScanBarcode") || "Scan Barcode",
            Clear: props.context.resources.getString("Clear") || "Clear",
            PlateNumber:
                props.context.resources.getString("PlateNumber") || "Plate Number",
            RegistrationNumber: props.context.resources.getString("RegistrationNumber")
                || "Registration Number",
            ContactAdministrator:
                props.context.resources.getString("ContactAdministrator") ||
                "An error occurred. Please contact the administrator.",
            BoatNumber:
                props.context.resources.getString("BoatNumber") || "Boat Number",
            QataryIDMustBe11Digits:
                props.context.resources.getString("QataryIDMustBe11Digits") ||
                "Qatary ID must be exactly 11 digits",
            ProjectName:
                props.context.resources.getString("ProjectName") || "Project Name",
            RequestPermitNumber:
                props.context.resources.getString("RequestPermitNumber") || "Request/Permit Number",
            LocationDetails:
                props.context.resources.getString("LocationDetails") || "Location Details",
            SelectProject:
                props.context.resources.getString("SelectProject") || "Select Project",
            Back:
                props.context.resources.getString("Back") || "Back",
            Select:
                props.context.resources.getString("Select") || "Select",
            NoProjectsFound:
                props.context.resources.getString("NoProjectsFound") || "No projects found nearby",
            SearchingProjects:
                props.context.resources.getString("SearchingProjects") || "Searching nearby projects...",
            ProjectNumber:
                props.context.resources.getString("ProjectNumber") || "Project #",
            AccountNameLabel:
                props.context.resources.getString("AccountNameLabel") || "Account",
        };

        this.state = {
            isRTL: isRTL,
            inspectionTypes: [],
            selectedInspectionType: props.defaultInspectionType || null,
            qataryId: "",
            name: "",
            crNumber: "",
            cpNumber: "",
            crCpToggle: 'cr',
            id: "",
            carColor: "",
            vehicleBrand: null,
            vehicleBrands: [],
            boatNumber: "",
            projectName: "",
            requestPermitNumber: "",
            locationDetails: "",
            loading: false,
            error: null,
            accountTypeRecord: null,
            showCampaignIncidentPopup: false,
            selectedCampaignId: props.activePatrolId,
            selectedCampaignName: props.activePatrolName,
            selectedIncidentTypeId: undefined,
            selectedIncidentTypeName: undefined,
            campaigns: [],
            incidentTypes: [],
            isAnonymous: false,
            popupShowCampaign: false,
            popupShowIncidentType: false,
            incidentTypeReadOnly: false,
            campaignIncidentTypeMap: {},
            registrationNumber: "",
            shouldShowCampaignField: true, // Default to true, will be updated after checking user's department
            // Project selection (type 16)
            showProjectSelectionPopup: false,
            nearbyProjects: [],
            selectedProjectId: null,
            projectSearchLoading: false,
        };
    }

    async componentDidMount(): Promise<void> {
        try {
            const userId =
                this.xrm.Utility.getGlobalContext().userSettings.userId.replace(
                    /[{}]/g,
                    "",
                );

            // Parallel loading — includes InitCache
            await Promise.all([
                this.loadInspectionTypesFromOrgUnit(),
                this.loadVehicleTypes(),
                this.preloadCampaignsAndIncidentTypes(),
                InitCache.load(userId),
                this.checkCampaignVisibility(userId),
            ]);

            // If default inspection type is provided, set account type from loaded data
            if (this.props.defaultInspectionType) {
                const inspectionType = this.state.inspectionTypes.find(
                    (t) => t.value === this.props.defaultInspectionType,
                );
                if (inspectionType?.accountTypeId) {
                    const accountTypeRecord = {
                        duc_accounttypeid: inspectionType.accountTypeId,
                        duc_accounttype: this.props.defaultInspectionType,
                        duc_name: inspectionType.label,
                    };
                    this.setState({ accountTypeRecord });
                }
            }
        } catch (error: any) {
            console.error("Error in componentDidMount:", error);
            this.setState({ error: this.strings.ContactAdministrator });
        }
    }

    // =====================================================================
    // GENERIC CACHE UTILITIES
    // =====================================================================

    private getFromCache = <T>(key: string): T | null => {
        try {
            const cached = localStorage.getItem(key);
            if (!cached) return null;

            const cacheData: CacheData<T> = JSON.parse(cached);
            const now = Date.now();

            if (now - cacheData.timestamp > CACHE_DURATION) {
                localStorage.removeItem(key);
                return null;
            }

            return cacheData.data;
        } catch (error: any) {
            console.error(`Error reading cache for ${key}:`, error);
            return null;
        }
    };

    private saveToCache = <T>(key: string, data: T): void => {
        try {
            const cacheData: CacheData<T> = {
                data,
                timestamp: Date.now(),
            };
            localStorage.setItem(key, JSON.stringify(cacheData));
        } catch (error: any) {
            console.error(`Error saving cache for ${key}:`, error);
        }
    };

    // =====================================================================
    // INSPECTION TYPES FROM ORGANIZATION UNIT
    // =====================================================================

    private loadInspectionTypesFromOrgUnit = async (): Promise<void> => {
        if (!this.props.organizationUnitId) {
            console.warn("No organization unit ID provided");
            return;
        }

        try {
            const cacheKey = `${INSPECTION_TYPES_CACHE_KEY}_${this.props.organizationUnitId}`;
            const cachedTypes =
                this.getFromCache<
                    Array<{ value: number; label: string; accountTypeId: string }>
                >(cacheKey);

            if (cachedTypes) {
                this.setState({ inspectionTypes: cachedTypes });
                return;
            }

            // Fetch from junction entity — include duc_name & duc_namear for labels
            const query =
                `?$filter=_duc_organizationunit_value eq '${this.props.organizationUnitId}'` +
                `&$select=duc_organizationunitaccounttypesid,duc_name,duc_namear` +
                `&$expand=duc_AccountType($select=duc_accounttypeid,duc_name,duc_accounttype)`;

            const results = await this.xrm.WebApi.retrieveMultipleRecords(
                "duc_organizationunitaccounttypes",
                query,
            );

            if (!results?.entities || results.entities.length === 0) {
                console.warn("No account types found for organization unit");
                return;
            }

            // Map results to inspection types
            // Use duc_namear (Arabic) or duc_name (English) from the junction entity
            const types: Array<{
                value: number;
                label: string;
                accountTypeId: string;
                orgUnitAccountTypeId?: string;
            }> = [];

            for (const entity of results.entities) {
                if (entity.duc_AccountType?.duc_accounttype !== undefined) {
                    const optionValue = entity.duc_AccountType.duc_accounttype;
                    // Pick label based on language direction
                    const label = this.state.isRTL
                        ? (entity.duc_namear || entity.duc_name || `Type ${optionValue}`)
                        : (entity.duc_name || entity.duc_namear || `Type ${optionValue}`);

                    types.push({
                        value: optionValue,
                        label: label,
                        accountTypeId: entity.duc_AccountType.duc_accounttypeid,
                        orgUnitAccountTypeId: entity.duc_organizationunitaccounttypesid,
                    });
                }
            }

            // Sort by value
            types.sort((a, b) => a.value - b.value);

            this.saveToCache(cacheKey, types);
            this.setState({ inspectionTypes: types });
        } catch (error: any) {
            console.error("Error loading inspection types from org unit:", error);
            this.setState({ error: this.strings.ContactAdministrator });
        }
    };

    // =====================================================================
    // VEHICLE TYPES
    // =====================================================================

    private loadVehicleTypes = async (): Promise<void> => {
        try {
            const cachedTypes = this.getFromCache<
                Array<{ value: number; label: string }>
            >(VEHICLE_TYPES_CACHE_KEY);
            if (cachedTypes) {
                this.setState({ vehicleBrands: cachedTypes });
                return;
            }

            const entityMetadata = await this.xrm.Utility.getEntityMetadata(
                "account",
                ["duc_vehicletype"],
            );
            const attribute = (entityMetadata as any).Attributes.get(
                "duc_vehicletype",
            );

            if (attribute?.OptionSet) {
                const types = Object.values(attribute.OptionSet).map((opt: any) => ({
                    value: opt.value,
                    label: opt.text,
                }));

                this.saveToCache(VEHICLE_TYPES_CACHE_KEY, types);
                this.setState({ vehicleBrands: types });
            }
        } catch (error: any) {
            console.error("Error loading vehicle types:", error);
        }
    };

    // =====================================================================
    // CHECK CAMPAIGN VISIBILITY FROM USER'S DEPARTMENT
    // =====================================================================

    private checkCampaignVisibility = async (userId: string): Promise<void> => {
        try {
            // Fetch the current user's department
            const userQuery = `?$select=systemuserid&$expand=duc_department($select=duc_hidecampaignonhomepage)`;
            const userResult = await this.xrm.WebApi.retrieveRecord(
                "systemuser",
                userId,
                userQuery,
            );

            // Check if the department has the hide campaign flag set
            const hideCampaign = userResult?.duc_department?.duc_hidecampaignonhomepage;

            // Show campaign if the flag is null or false, hide if true
            this.setState({ shouldShowCampaignField: !hideCampaign });
        } catch (error: any) {
            console.error("Error checking campaign visibility:", error);
            // Default to showing campaign on error
            this.setState({ shouldShowCampaignField: true });
        }
    };

    // =====================================================================
    // CAMPAIGNS AND INCIDENT TYPES (PRELOAD)
    // =====================================================================

    private preloadCampaignsAndIncidentTypes = async (): Promise<void> => {
        try {
            let [campaigns, incidentTypes] = await Promise.all([
                this.loadCampaigns(),
                this.loadIncidentTypes(),
            ]);

            // INJECT ACTIVE PATROL: Ensure props.activePatrolId is in the list
            if (this.props.activePatrolId && this.props.activePatrolName) {
                const exists = campaigns.some(
                    (c) => c.id === this.props.activePatrolId,
                );
                if (!exists) {
                    campaigns.push({
                        id: this.props.activePatrolId,
                        name: this.props.activePatrolName,
                    });
                    // Sort again to keep list orderly
                    campaigns.sort((a, b) => a.name.localeCompare(b.name));
                }
            }

            // Build campaign → incident type map to avoid redundant retrieves
            const campaignIncidentTypeMap: Record<
                string,
                { id: string; name: string }
            > = {};
            for (const campaign of campaigns) {
                try {
                    const campaignData = await WorkOrderHelpers.getCampaignData(
                        campaign.id,
                    );
                    if (campaignData?.incidentType) {
                        campaignIncidentTypeMap[campaign.id] = {
                            id: campaignData.incidentType.id,
                            name: campaignData.incidentType.name,
                        };
                    }
                } catch {
                    // Skip individual campaign failures
                }
            }

            this.setState({ campaigns, incidentTypes, campaignIncidentTypeMap });
        } catch (error: any) {
            console.error("Error preloading campaigns/incident types:", error);
        }
    };

    private loadCampaigns = async (): Promise<
        Array<{ id: string; name: string }>
    > => {
        if (!this.props.organizationUnitId) return [];

        const cacheKey = `${CAMPAIGNS_CACHE_KEY}_${this.props.organizationUnitId}`;
        const cached =
            this.getFromCache<Array<{ id: string; name: string }>>(cacheKey);
        if (cached) return cached;

        try {

            // Campaign status: Active = 2, Campaign type: AdHoc = 100000000
            const query = `?$filter=_duc_organizationalunitid_value eq '${this.props.organizationUnitId}' and duc_campaignstatus eq 2 and  duc_campaigntype eq 100000000 and statecode eq 0 &$select=new_inspectioncampaignid,new_name&$orderby=new_name asc`;

            const results = await this.xrm.WebApi.retrieveMultipleRecords(
                "new_inspectioncampaign",
                query,
            );

            const campaigns = results.entities.map((entity: any) => ({
                id: entity.new_inspectioncampaignid,
                name: entity.new_name,
            }));

            this.saveToCache(cacheKey, campaigns);
            return campaigns;
        } catch (error: any) {
            console.error("Error loading campaigns:", error);
            return [];
        }
    };

    private loadIncidentTypes = async (): Promise<
        Array<{ id: string; name: string }>
    > => {
        if (!this.props.organizationUnitId) return [];

        try {
            const userId = this.xrm.Utility.getGlobalContext().userSettings.userId.replace(
                /[{}]/g,
                "",
            );
            const userRecord = await this.xrm.WebApi.retrieveRecord(
                "systemuser",
                userId,
                "?$select=duc_hasjudicialauth",
            );
            const hasJudicialAuth = userRecord?.duc_hasjudicialauth === true;

            const cacheKey = `${INCIDENT_TYPES_CACHE_KEY}_${this.props.organizationUnitId}_${hasJudicialAuth ? "1" : "0"}`;
            const cached =
                this.getFromCache<Array<{ id: string; name: string }>>(cacheKey);
            if (cached) return cached;

            const judicialFilter = hasJudicialAuth
                ? ""
                : `<filter type="or">
            <condition attribute="duc_judicialauthorityisrequired" operator="null" />
            <condition attribute="duc_judicialauthorityisrequired" operator="eq" value="0" />
        </filter>`;

            const fetchXml = `<fetch>
  <entity name="msdyn_incidenttype">
    <attribute name="msdyn_incidenttypeid" />
    <attribute name="msdyn_name" />
        <filter type="and">
            <filter type="or">
                <condition attribute="duc_organizationalunitid" operator="eq" value="${this.props.organizationUnitId}" />
                <link-entity name="duc_msdyn_incidenttype_msdyn_organizational" from="msdyn_incidenttypeid" to="msdyn_incidenttypeid" link-type="any" alias="DMIMO" intersect="true">
                    <filter>
                        <condition attribute="msdyn_organizationalunitid" operator="eq" value="${this.props.organizationUnitId}" />
                    </filter>
                </link-entity>
            </filter>
            ${judicialFilter}
        </filter>
    <order attribute="msdyn_name" />
  </entity>
</fetch>`;

            const results = await this.xrm.WebApi.retrieveMultipleRecords(
                "msdyn_incidenttype",
                `?fetchXml=${encodeURIComponent(fetchXml)}`
            );

            const incidentTypes = results.entities.map((entity: any) => ({
                id: entity.msdyn_incidenttypeid,
                name: entity.msdyn_name,
            }));

            this.saveToCache(cacheKey, incidentTypes);
            return incidentTypes;
        } catch (error: any) {
            console.error("Error loading incident types:", error);
            return [];
        }
    };

    // =====================================================================
    // HANDLERS
    // =====================================================================

    private handleInspectionTypeChange = (
        e: React.ChangeEvent<HTMLSelectElement>,
    ): void => {
        const selectedGuid = e.target.value || null;

        let accountTypeRecord: any = null;
        let optionValue: number | null = null;

        if (selectedGuid) {
            const inspectionType = this.state.inspectionTypes.find(
                (t) => t.orgUnitAccountTypeId === selectedGuid || t.accountTypeId === selectedGuid,
            );
            if (inspectionType) {
                optionValue = inspectionType.value;
                if (inspectionType.accountTypeId) {
                    accountTypeRecord = {
                        duc_accounttypeid: inspectionType.accountTypeId,
                        duc_accounttype: optionValue,
                        duc_name: inspectionType.label,
                        duc_organizationunitaccounttypesid: inspectionType.orgUnitAccountTypeId,
                    };
                }
            }
        }

        this.setState({
            selectedInspectionType: optionValue,
            qataryId: "",
            name: "",
            crNumber: "",
            cpNumber: "",
            crCpToggle: 'cr',
            id: "",
            carColor: "",
            vehicleBrand: null,
            boatNumber: "",
            projectName: "",
            requestPermitNumber: "",
            locationDetails: "",
            error: null,
            accountTypeRecord: accountTypeRecord,
            isAnonymous: false,
        });
    };

    private handleInputChange = (
        field: keyof IMultiTypeInspectionState,
        value: any,
    ): void => {
        this.setState({ [field]: value } as any);
    };

    /**
     * Handle numeric-only input fields (Qatari ID, CR Number, Registration Number)
     * Filters out any non-numeric characters
     */
    private handleNumericInputChange = (
        field: keyof IMultiTypeInspectionState,
        value: string,
    ): void => {
        // CR/CP Numbers: digits + at most one slash, not at the start.
        // Trailing slash is allowed while typing; stripped on submit validation.
        const numericValue = (field === 'crNumber' || field === 'cpNumber')
            ? (() => {
                const v = value.replace(/[^0-9/]/g, '');
                const parts = v.split('/');
                if (parts.length === 1) return parts[0];
                const before = parts[0];
                const after = parts.slice(1).join(''); // collapse extra slashes into one
                if (!before) return after; // strip leading slash
                return before + '/' + after; // allow trailing slash while user is still typing
            })()
            : value.replace(/[^0-9]/g, '');
        this.setState({ [field]: numericValue } as any);
    };

    // =====================================================================
    // BARCODE SCANNER
    // =====================================================================

    private handleScanBarcode = async (
        field: "qataryId" | "crNumber" | "cpNumber" | "registrationNumber",
    ): Promise<void> => {
        try {
            if (this.xrm?.Device?.getBarcodeValue) {
                const result: any = await this.xrm.Device.getBarcodeValue();
                if (result) {
                    this.setState({ [field]: result } as any);
                }
            } else {
                console.warn("Barcode scanner is not available on this device");
            }
        } catch (error: any) {
            console.error("Error scanning barcode:", error);
            this.setState({ error: this.strings.ContactAdministrator });
        }
    };

    // =====================================================================
    // CAMPAIGN ↔ INCIDENT TYPE POPUP HANDLERS
    // =====================================================================

    /**
     * When campaign is selected in the popup:
     * - Auto-fill incident type from cached map
     * - Make incident type read-only
     */
    private handleCampaignChange = (campaignId: string): void => {
        if (!campaignId) {
            // Campaign cleared
            this.setState({
                selectedCampaignId: undefined,
                selectedCampaignName: undefined,
                selectedIncidentTypeId: undefined,
                selectedIncidentTypeName: undefined,
                incidentTypeReadOnly: false,
            });
            return;
        }

        const campaignName = this.state.campaigns.find(
            (c) => c.id === campaignId,
        )?.name;
        const mappedIncidentType = this.state.campaignIncidentTypeMap[campaignId];

        this.setState({
            selectedCampaignId: campaignId,
            selectedCampaignName: campaignName,
            selectedIncidentTypeId: mappedIncidentType?.id,
            selectedIncidentTypeName: mappedIncidentType?.name,
            incidentTypeReadOnly: !!mappedIncidentType,
        });
    };

    /**
     * When incident type is selected manually in the popup:
     * - Campaign is NOT mandatory
     */
    private handleIncidentTypeChange = (incidentTypeId: string): void => {
        if (!incidentTypeId) {
            this.setState({
                selectedIncidentTypeId: undefined,
                selectedIncidentTypeName: undefined,
            });
            return;
        }

        const incidentTypeName = this.state.incidentTypes.find(
            (it) => it.id === incidentTypeId,
        )?.name;
        this.setState({
            selectedIncidentTypeId: incidentTypeId,
            selectedIncidentTypeName: incidentTypeName,
        });
    };

    // =====================================================================
    // VALIDATION
    // =====================================================================

    private getRequiredFields = (): Array<keyof IMultiTypeInspectionState> => {
        const { selectedInspectionType } = this.state;
        const requiredFields: Array<keyof IMultiTypeInspectionState> = [];

        if (!selectedInspectionType) return [];

        if (selectedInspectionType === 1) {
            requiredFields.push("id");
        } else if ([2, 3, 6, 15].includes(selectedInspectionType)) {
            requiredFields.push("qataryId", "name");
        } else if ([5, 7].includes(selectedInspectionType)) {
            requiredFields.push(this.state.crCpToggle === 'cr' ? "crNumber" : "cpNumber");
        } else if (selectedInspectionType === 14) {
            requiredFields.push("boatNumber");
        }
        // Type 16 (Project) — no fields needed on main form;
        // user picks from project list popup instead

        return requiredFields;
    };

    private validateFields = (): boolean => {
        const { selectedInspectionType, isAnonymous } = this.state;

        if (!selectedInspectionType) {
            this.setState({ error: this.strings.PleaseSelectInspectionType });
            return false;
        }

        // Anonymous (4), type 13 (Anonymous-like), and type 16 (Project) — no fields needed on main form
        if ([4, 13, 16].includes(selectedInspectionType)) {
            return true;
        }

        if (isAnonymous && [3, 6, 7].includes(selectedInspectionType)) {
            return true;
        }

        const requiredFields = this.getRequiredFields();
        for (const field of requiredFields) {
            const val: any = this.state[field];
            if (val === null || val === undefined || val === "") {
                this.setState({ error: this.strings.PleaseEnterRequiredFields });
                return false;
            }
        }

        // Validate Qatary ID must be exactly 11 digits
        if (requiredFields.includes("qataryId") && this.state.qataryId.length !== 11) {
            this.setState({ error: this.strings.QataryIDMustBe11Digits });
            return false;
        }

        // Validate CR/CP number: slash must be in the middle (not trailing)
        const crCpField = this.state.crCpToggle === 'cr' ? 'crNumber' : 'cpNumber';
        if (requiredFields.includes(crCpField)) {
            const crCpVal: string = this.state[crCpField] as string;
            if (crCpVal.endsWith('/')) {
                this.setState({ error: this.strings.PleaseEnterRequiredFields });
                return false;
            }
        }

        return true;
    };

    // =====================================================================
    // ACCOUNT SEARCH/CREATE
    // =====================================================================

    private getIdentifierValue = (): string => {
        const { selectedInspectionType, qataryId, crNumber, id, registrationNumber, boatNumber, requestPermitNumber } = this.state;

        if (selectedInspectionType === 1) return id;
        if ([2, 3, 6, 15].includes(selectedInspectionType!)) return qataryId;
        if ([5, 7].includes(selectedInspectionType!)) return this.state.crCpToggle === 'cr' ? crNumber : this.state.cpNumber;
        if ([10, 11].includes(selectedInspectionType!)) return registrationNumber;
        if (selectedInspectionType === 14) return boatNumber;

        return "";
    };

    private getAccountName = async (): Promise<string> => {
        const {
            selectedInspectionType,
            name,
            qataryId,
            crNumber,
            id,
            carColor,
            vehicleBrand,
            vehicleBrands,
            registrationNumber,
            boatNumber
        } = this.state;

        if (!selectedInspectionType)
            return "Account";

        if (name && name.trim() !== "") {
            return name.trim();
        }

        const typeInfo = this.state.inspectionTypes.find(t => t.value === selectedInspectionType);

        if (typeInfo && typeInfo.accountTypeId) {
            try {
                const accTypeResult = await Xrm.WebApi.retrieveRecord("duc_accounttype", typeInfo.accountTypeId, "?$select=duc_overrideaccountname");
                if (accTypeResult.duc_overrideaccountname) {
                    return accTypeResult.duc_overrideaccountname;
                }
            } catch (e) {
                console.warn("Failed to get override account name");
            }
        }

        switch (selectedInspectionType) {
            case 1: // Vehicle
                const brandLabel =
                    vehicleBrand !== null
                        ? vehicleBrands.find((v) => v.value === vehicleBrand)?.label ||
                        vehicleBrand
                        : "";
                return `${id} ${carColor} ${brandLabel}`.trim() || "Account";

            case 2: // Individual
            case 15: // Individual-like
            case 3: // Cabin
            case 6: // Wilderness Camp
                return qataryId || "Account";

            case 5: // Company
            case 7: // Manor
                return (this.state.crCpToggle === 'cr' ? crNumber : this.state.cpNumber) || "Account";

            case 10: // Establishment
            case 11: // Hospital
                return registrationNumber || "Account";

            case 14: // Boat
                return boatNumber || "Account";

            case 16: // Project
                return this.state.projectName || "Project";

            case 13: // Important maritime areas
                return "Important maritime areas";

            case 4: // Anonymous
            default:
                return "Account";
        }
    };

    private getCurrentLocation = async (): Promise<{
        latitude: number;
        longitude: number;
    } | null> => {
        try {
            if (this.xrm?.Device?.getCurrentPosition) {
                const location: any = await this.xrm.Device.getCurrentPosition();
                return {
                    latitude: location.coords.latitude,
                    longitude: location.coords.longitude,
                };
            }
        } catch (error: any) {
            console.error("Error getting location:", error);
        }
        return null;
    };

    private createAddressInformation = async (
        accountId: string,
        accountName: string,
    ): Promise<void> => {
        try {
            const location = await this.getCurrentLocation();
            if (!location) return;

            const today = new Date().toISOString().split("T")[0];
            const addressName = `${accountName} ${today}`;

            const addressData: any = {
                duc_name: addressName,
                duc_latitude: location.latitude,
                duc_longitude: location.longitude,
                "duc_Account@odata.bind": `/accounts(${accountId})`,
            };

            var results = await this.xrm.WebApi.createRecord("duc_addressinformation", addressData);

            // associate address to account
            await this.xrm.WebApi.updateRecord("account", accountId,
                {
                    'duc_Address@odata.bind': `/duc_addressinformations(${results.id})`
                });

        } catch (error: any) {
            console.error("Error creating address information:", error);
        }
    };

    private searchOrCreateAccount = async (): Promise<string | null> => {
        try {
            const {
                accountTypeRecord,
                selectedInspectionType,
                carColor,
                vehicleBrand,
                isAnonymous,
            } = this.state;
            const identifierValue = this.getIdentifierValue();

            // Anonymous type (4) uses unknown account
            if (selectedInspectionType === 4) {
                if (this.props.unknownAccountId) {
                    return this.props.unknownAccountId;
                } else {
                    throw new Error(this.strings.ContactAdministrator);
                }
            }

            // Type 13: Always create a new account with name "Important maritime areas", no identifier
            if (selectedInspectionType === 13) {
                const accountName = "Important maritime areas";
                const newAccount: any = {
                    name: accountName,
                    duc_accountinspectiontype: selectedInspectionType,
                };

                if (accountTypeRecord?.duc_accounttypeid) {
                    newAccount["duc_NewAccountType@odata.bind"] =
                        `/duc_accounttypes(${accountTypeRecord.duc_accounttypeid})`;
                }

                const createdAccount = await this.xrm.WebApi.createRecord(
                    "account",
                    newAccount,
                );
                const newAccountId = createdAccount?.id;

                await this.createAddressInformation(newAccountId, accountName);

                return newAccountId;
            }

            // Site
            else if (selectedInspectionType === 12) {
                if (this.props.siteAccountId) {
                    return this.props.siteAccountId;
                } else {
                    throw new Error(this.strings.ContactAdministrator);
                }
            }

            // If anonymous checkbox is checked for types 3, 6, or 7, use unknown account
            if (isAnonymous && [3, 6, 7].includes(selectedInspectionType!)) {
                if (this.props.unknownAccountId) {
                    return this.props.unknownAccountId;
                } else {
                    throw new Error(this.strings.ContactAdministrator);
                }
            }

            if (!identifierValue) {
                throw new Error(this.strings.PleaseEnterRequiredFields);
            }

            // Search for existing account
            var filterQuery;
            if ([10, 11].includes(selectedInspectionType!)) {
                filterQuery = `duc_moinumber eq '${identifierValue}'`;
            } else if ([5, 7].includes(selectedInspectionType!) && this.state.crCpToggle === 'cp') {
                filterQuery = `duc_cpnumber eq '${identifierValue}'`;
            } else {
                filterQuery = `duc_accountidentifier eq '${identifierValue}'`;
            }

            if (accountTypeRecord?.duc_accounttypeid) {
                filterQuery += ` and _duc_newaccounttype_value eq ${accountTypeRecord.duc_accounttypeid}`;
            }

            const searchResults = await this.xrm.WebApi.retrieveMultipleRecords(
                "account",
                `?$select=accountid,name,duc_vehicletype&$filter=${filterQuery}`,
            );

            // Update or create account
            if (searchResults?.entities?.length > 0) {
                const accountId = searchResults.entities[0].accountid;
                const accountName = searchResults.entities[0].name;

                // Update vehicle info if type is vehicle
                if (selectedInspectionType === 1) {
                    const updateData: any = {};
                    if (carColor) updateData.duc_vehiclecolor = carColor;
                    if (vehicleBrand !== null) updateData.duc_vehicletype = vehicleBrand;

                    if (Object.keys(updateData).length > 0) {
                        await this.xrm.WebApi.updateRecord(
                            "account",
                            accountId,
                            updateData,
                        );
                    }
                }

                // Existing account — skip address creation
                return accountId;
            }

            // Create new account
            const accountName = await this.getAccountName();
            const newAccount: any = {
                name: accountName,
                duc_accountinspectiontype: selectedInspectionType,
            };

            // Set identifier field based on type and CR/CP toggle
            if ([5, 7].includes(selectedInspectionType!) && this.state.crCpToggle === 'cp') {
                newAccount.duc_cpnumber = identifierValue;
            } else {
                newAccount.duc_accountidentifier = identifierValue;
            }

            // Establishment or Hospital
            if ([10, 11].includes(selectedInspectionType!)) {
                newAccount.duc_moinumber = identifierValue;
            }

            // Project (16): Set description with Request/Permit Number + Location Details
            if (selectedInspectionType === 16) {
                const descriptionParts = [];
                if (this.state.requestPermitNumber) {
                    descriptionParts.push(this.state.requestPermitNumber);
                }
                if (this.state.locationDetails) {
                    descriptionParts.push(this.state.locationDetails);
                }
                if (descriptionParts.length > 0) {
                    newAccount.description = descriptionParts.join(" - ");
                }
            }

            if (accountTypeRecord?.duc_accounttypeid) {
                newAccount["duc_NewAccountType@odata.bind"] =
                    `/duc_accounttypes(${accountTypeRecord.duc_accounttypeid})`;
            }

            if (selectedInspectionType === 1) {
                if (carColor) newAccount.duc_vehiclecolor = carColor;
                if (vehicleBrand !== null) newAccount.duc_vehicletype = vehicleBrand;
            }

            const createdAccount = await this.xrm.WebApi.createRecord(
                "account",
                newAccount,
            );
            const newAccountId = createdAccount?.id;

            await this.createAddressInformation(newAccountId, accountName);

            return newAccountId;
        } catch (error: any) {
            console.error("Error searching/creating account:", error);
            throw new Error(this.strings.ContactAdministrator);
        }
    };

    // =====================================================================
    // WORK ORDER CREATION WITH FULL BUSINESS LOGIC
    // =====================================================================

    private createWorkOrder = async (accountId: string): Promise<void> => {
        try {
            const userId =
                this.xrm.Utility.getGlobalContext().userSettings.userId.replace(
                    /[{}]/g,
                    "",
                );

            // Get account name and data
            const accountRecord = await this.xrm.WebApi.retrieveRecord(
                "account",
                accountId,
                "?$select=name",
            );
            const accountName = accountRecord?.name || "";

            // Prepare base data
            let serviceAccountData = {
                id: accountId,
                name: accountName,
                entityType: "account",
            };

            let addressData: any = null;
            let latitude: number | undefined = undefined;
            let longitude: number | undefined = undefined;

            // STEP 1: Handle sub-account change logic
            const subAccountResult =
                await WorkOrderHelpers.handleSubAccountChange(accountId);
            if (subAccountResult) {
                serviceAccountData = subAccountResult.serviceAccount;
                addressData = subAccountResult.address;
                latitude = subAccountResult.latitude;
                longitude = subAccountResult.longitude;
            }

            // STEP 2: Determine incident type
            let incidentTypeData:
                | { id: string; name: string; entityType: string }
                | undefined;

            if (
                this.state.selectedIncidentTypeId &&
                this.state.selectedIncidentTypeName
            ) {
                incidentTypeData = {
                    id: this.state.selectedIncidentTypeId,
                    name: this.state.selectedIncidentTypeName,
                    entityType: "msdyn_incidenttype",
                };
            } else if (this.state.selectedCampaignId) {
                // Get incident type from campaign using cached map (no API call)
                const mapped =
                    this.state.campaignIncidentTypeMap[this.state.selectedCampaignId];
                if (mapped) {
                    incidentTypeData = {
                        id: mapped.id,
                        name: mapped.name,
                        entityType: "msdyn_incidenttype",
                    };
                } else {
                    // Fallback: fetch from API (should rarely happen since we preloaded)
                    const campaignData = await WorkOrderHelpers.getCampaignData(
                        this.state.selectedCampaignId,
                    );
                    if (campaignData?.incidentType) {
                        incidentTypeData = campaignData.incidentType;
                    }
                }
            } else if (this.props.incidentTypeId && this.props.incidentTypeName) {
                incidentTypeData = {
                    id: this.props.incidentTypeId,
                    name: this.props.incidentTypeName,
                    entityType: "msdyn_incidenttype",
                };
            }

            if (!incidentTypeData) {
                throw new Error(this.strings.ContactAdministrator);
            }

            // STEP 3 & 4: Get work order type AND department in ONE API call
            const incidentData = await WorkOrderHelpers.getIncidentTypeData(
                incidentTypeData.id,
            );
            const workOrderTypeData = incidentData?.workOrderType;
            let departmentData = incidentData?.department || null;

            if (!departmentData) {
                departmentData = await WorkOrderHelpers.setDepartmentFromUser(userId);
            }

            if (
                !departmentData &&
                this.props.organizationUnitId &&
                this.props.organizationUnitName
            ) {
                departmentData = {
                    id: this.props.organizationUnitId,
                    name: this.props.organizationUnitName,
                };
            }

            if (!departmentData) {
                throw new Error(this.strings.ContactAdministrator);
            }

            // STEP 5: Prepare campaign data
            let campaignData: { id: string; name: string } | undefined;
            let parentCampaignData: { id: string; name: string } | undefined;

            // Helper function to retrieve parent campaign from campaign record
            const getParentCampaignFromCampaign = async (
                campaignId: string,
                campaignName: string
            ): Promise<{ id: string; name: string } | undefined> => {
                const query = "?$select=_duc_parentcampaign_value&$expand=duc_ParentCampaign($select=new_inspectioncampaignid,new_name)";
                console.group(`[getParentCampaignFromCampaign] campaignId=${campaignId} | campaignName=${campaignName}`);
                console.log("[QUERY] entity: new_inspectioncampaign");
                console.log("[QUERY] id:", campaignId);
                console.log("[QUERY] options:", query);
                try {
                    const campaignRecord = await this.xrm.WebApi.retrieveRecord(
                        "new_inspectioncampaign",
                        campaignId,
                        query
                    );

                    console.log("[RESPONSE] Full record:", JSON.stringify(campaignRecord, null, 2));
                    console.log("[RESPONSE] _duc_parentcampaign_value:", campaignRecord._duc_parentcampaign_value);
                    console.log("[RESPONSE] duc_ParentCampaign (expanded):", campaignRecord.duc_ParentCampaign);
                    console.log("[RESPONSE] duc_parentcampaign (lowercase):", campaignRecord.duc_parentcampaign);
                    console.log("[RESPONSE] All keys in record:", Object.keys(campaignRecord));

                    // If parent campaign exists, return it
                    const parentExpanded = campaignRecord.duc_ParentCampaign ?? campaignRecord.duc_parentcampaign;
                    if (campaignRecord._duc_parentcampaign_value && parentExpanded) {
                        const result = {
                            id: parentExpanded.new_inspectioncampaignid,
                            name: parentExpanded.new_name,
                        };
                        console.log("[RESULT] Parent campaign found:", result);
                        console.groupEnd();
                        return result;
                    }

                    console.warn("[RESULT] No parent campaign found — returning campaign itself as parent");
                    console.groupEnd();
                    return { id: campaignId, name: campaignName };
                } catch (error: any) {
                    console.error("[ERROR] retrieveRecord failed:", error);
                    console.error("[ERROR] message:", error?.message);
                    console.error("[ERROR] stack:", error?.stack);
                    console.groupEnd();
                    // Fallback: use the campaign itself as parent
                    return { id: campaignId, name: campaignName };
                }
            };

            if (this.state.selectedCampaignId && this.state.selectedCampaignName) {
                campaignData = {
                    id: this.state.selectedCampaignId,
                    name: this.state.selectedCampaignName,
                };
                // Retrieve parent campaign from campaign record
                parentCampaignData = await getParentCampaignFromCampaign(
                    this.state.selectedCampaignId,
                    this.state.selectedCampaignName
                );
            } else if (this.props.activePatrolId && this.props.activePatrolName) {
                campaignData = {
                    id: this.props.activePatrolId,
                    name: this.props.activePatrolName,
                };
                // Retrieve parent campaign from campaign record
                parentCampaignData = await getParentCampaignFromCampaign(
                    this.props.activePatrolId,
                    this.props.activePatrolName
                );
            }

            // STEP 6: Determine anonymous customer flag
            const anonymousCustomer =
                [4, 13].includes(this.state.selectedInspectionType!) ||
                (this.state.isAnonymous &&
                    [3, 6, 7].includes(this.state.selectedInspectionType!));

            // STEP 7: Detect if created from mobile
            const createdFromMobile = WorkOrderHelpers.isMobileClient();

            // STEP 8: Validate data
            const validation = WorkOrderHelpers.validateWorkOrderData({
                serviceAccount: serviceAccountData,
                incidentType: incidentTypeData,
                department: departmentData,
            });

            if (!validation.isValid) {
                throw new Error(this.strings.ContactAdministrator);
            }

            // STEP 9: Create work order — with progress indicator
            this.xrm.Utility.showProgressIndicator(this.strings.CreatingWorkOrder);
            const workOrderId = await WorkOrderHelpers.createWorkOrder({
                subAccount: {
                    id: accountId,
                    name: accountName,
                },
                serviceAccount: serviceAccountData,
                incidentType: incidentTypeData,
                department: departmentData,
                campaign: campaignData,
                workOrderType: workOrderTypeData || undefined,
                parentCampaign: parentCampaignData,
                address: addressData || undefined,
                latitude: latitude,
                longitude: longitude,
                anonymousCustomer: anonymousCustomer,
                accountInspectionType: this.state.selectedInspectionType || undefined,
                createdFromMobile: createdFromMobile,
            });
            this.xrm.Utility.closeProgressIndicator();

            if (!workOrderId) {
                throw new Error(this.strings.ContactAdministrator);
            }

            console.log("Work order created successfully:", workOrderId);

            // STEP 10: Create auto booking if from mobile — using cached values
            if (createdFromMobile && InitCache.hasBookableResource) {
                this.xrm.Utility.showProgressIndicator(this.strings.CreatingBooking);
                const bookingId = await WorkOrderHelpers.createAutoBooking(
                    workOrderId,
                    userId,
                    {
                        bookableResourceId: InitCache.bookableResourceId!,
                        bookingStatusId: InitCache.bookingStatusId!,
                    },
                );
                this.xrm.Utility.closeProgressIndicator();
                if (bookingId) {
                    console.log("Auto booking created:", bookingId);
                }
            }

            // STEP 11: Navigate to the created work order
            await this.xrm.Navigation.openForm({
                entityName: "msdyn_workorder",
                entityId: workOrderId,
                formId: "b7b3d199-8809-f111-8341-6045bd8e2841",
                openInNewWindow: true,
            });

            // Close the modal
            if (this.props.onClose) {
                this.props.onClose();
            }
        } catch (error: any) {
            this.xrm.Utility.closeProgressIndicator();
            console.error("Error creating work order:", error);
            throw new Error(this.strings.ContactAdministrator);
        }
    };

    // =====================================================================
    // MAIN START HANDLER
    // =====================================================================

    private handleStart = async (): Promise<void> => {
        if (!this.validateFields()) {
            return;
        }

        try {
            this.setState({ loading: true, error: null });

            // PROJECT TYPE (16): Intercept to show nearby projects selection first
            if (this.state.selectedInspectionType === 16) {
                this.xrm.Utility.showProgressIndicator(this.strings.SearchingProjects);

                // Get current location for nearby search
                const location = await this.getCurrentLocation();
                const latitude = location?.latitude || 0;
                const longitude = location?.longitude || 0;

                // Search nearby projects via Logic App API
                const nearbyResults = await ProjectHelpers.searchNearbyProjects({
                    radius: ProjectHelpers.defaultSearchRadius,
                    accountType: 16,
                    longitude: longitude,
                    latitude: latitude,
                    departmentId: this.props.organizationUnitId || "",
                });

                // Enrich with project details (names, account info)
                const projectIds = nearbyResults.map(r => r.projectId);
                const projectDetails = await ProjectHelpers.getProjectDetails(projectIds);

                this.xrm.Utility.closeProgressIndicator();

                // Show project selection popup
                this.setState({
                    showProjectSelectionPopup: true,
                    nearbyProjects: projectDetails,
                    selectedProjectId: null,
                    projectSearchLoading: false,
                    loading: false,
                    error: null,
                });
                return; // Stop here — user will pick a project, then handleProjectConfirm continues
            }

            // Get or create account — with progress indicator
            this.xrm.Utility.showProgressIndicator(this.strings.CreatingAccount);
            const accountId = await this.searchOrCreateAccount();
            this.xrm.Utility.closeProgressIndicator();

            if (!accountId) {
                throw new Error("Failed to get account");
            }

            // Determine Campaign and Incident Type values to allow pre-filling the popup
            let resolvedIncidentTypeId =
                this.state.selectedIncidentTypeId || this.props.incidentTypeId;
            let resolvedIncidentTypeName =
                this.state.selectedIncidentTypeName || this.props.incidentTypeName;

            const campaignId =
                this.state.selectedCampaignId || this.props.activePatrolId;
            let incidentTypeDerivedFromCampaign = false;

            // Auto-resolve incident type from campaign if available
            if (!resolvedIncidentTypeId && campaignId) {
                const mapped = this.state.campaignIncidentTypeMap[campaignId];
                if (mapped) {
                    resolvedIncidentTypeId = mapped.id;
                    resolvedIncidentTypeName = mapped.name;
                    incidentTypeDerivedFromCampaign = true;
                }
            }

            // ALWAYS show selection popup (no auto-creation)
            this.pendingAccountId = accountId;
            this.setState((prev) => ({
                ...prev,
                showCampaignIncidentPopup: true,
                popupShowCampaign: this.state.shouldShowCampaignField, // Show based on user's department setting
                popupShowIncidentType: true, // Always show incident type field
                loading: false,
                error: null,

                // Pre-fill popup with resolved values
                selectedCampaignId: campaignId,
                selectedIncidentTypeId: resolvedIncidentTypeId,
                selectedIncidentTypeName: resolvedIncidentTypeName,

                // If incident type is derived from campaign, make it read-only
                incidentTypeReadOnly: incidentTypeDerivedFromCampaign,
            }));
        } catch (error: any) {
            this.xrm.Utility.closeProgressIndicator();
            console.error("Error in handleStart:", error);
            this.setState({
                error: this.strings.ContactAdministrator,
                loading: false,
            });
        }
    };

    private handleContinueWithSelections = async (): Promise<void> => {
        try {
            // Validate that incident type is selected
            if (!this.state.selectedIncidentTypeId) {
                this.setState({
                    error:
                        this.strings.SelectIncidentType || "Please select an incident type",
                    loading: false,
                });
                return;
            }

            // Keep popup open while loading to prevent "Main Form" flash
            this.setState({ loading: true, error: null });

            // Reuse account ID from handleStart instead of re-searching
            const accountId =
                this.pendingAccountId || (await this.searchOrCreateAccount());
            this.pendingAccountId = null;
            if (!accountId) {
                throw new Error("Failed to get account");
            }

            await this.createWorkOrder(accountId);
            this.setState({ loading: false });
        } catch (error: any) {
            this.xrm.Utility.closeProgressIndicator();
            console.error("Error creating work order (handleContinueWithSelections):", error);
            this.setState({
                error: this.strings.ContactAdministrator,
                loading: false,
            });
        }
    };

    // =====================================================================
    // PROJECT SELECTION HANDLERS (TYPE 16)
    // =====================================================================

    /**
     * Handle user selecting a project from the nearby projects list.
     * Single-select: only one project can be selected at a time.
     */
    private handleProjectSelect = (projectId: string): void => {
        this.setState({ selectedProjectId: projectId });
    };

    /**
     * Handle user confirming the selected project.
     * Gets the account linked to the project, then continues the normal flow
     * (Campaign/Incident popup → work order creation).
     */
    private handleProjectConfirm = async (): Promise<void> => {
        const { selectedProjectId } = this.state;

        if (!selectedProjectId) {
            this.setState({ error: this.strings.SelectProject });
            return;
        }

        try {
            this.setState({ loading: true, error: null });

            // Get the account linked to the selected project
            this.xrm.Utility.showProgressIndicator(this.strings.Loading);
            const accountData = await ProjectHelpers.getAccountForProject(selectedProjectId);
            this.xrm.Utility.closeProgressIndicator();

            if (!accountData) {
                throw new Error("No account found for the selected project");
            }

            // Close project popup and continue to Campaign/Incident popup
            // (same flow as other types after searchOrCreateAccount)
            const accountId = accountData.accountId;

            // Determine Campaign and Incident Type values to allow pre-filling the popup
            let resolvedIncidentTypeId =
                this.state.selectedIncidentTypeId || this.props.incidentTypeId;
            let resolvedIncidentTypeName =
                this.state.selectedIncidentTypeName || this.props.incidentTypeName;

            const campaignId =
                this.state.selectedCampaignId || this.props.activePatrolId;
            let incidentTypeDerivedFromCampaign = false;

            // Auto-resolve incident type from campaign if available
            if (!resolvedIncidentTypeId && campaignId) {
                const mapped = this.state.campaignIncidentTypeMap[campaignId];
                if (mapped) {
                    resolvedIncidentTypeId = mapped.id;
                    resolvedIncidentTypeName = mapped.name;
                    incidentTypeDerivedFromCampaign = true;
                }
            }

            // Store account ID and show Campaign/Incident popup (normal flow)
            this.pendingAccountId = accountId;
            this.setState((prev) => ({
                ...prev,
                showProjectSelectionPopup: false,
                showCampaignIncidentPopup: true,
                popupShowCampaign: this.state.shouldShowCampaignField,
                popupShowIncidentType: true,
                loading: false,
                error: null,

                // Pre-fill popup with resolved values
                selectedCampaignId: campaignId,
                selectedIncidentTypeId: resolvedIncidentTypeId,
                selectedIncidentTypeName: resolvedIncidentTypeName,

                // If incident type is derived from campaign, make it read-only
                incidentTypeReadOnly: incidentTypeDerivedFromCampaign,
            }));
        } catch (error: any) {
            this.xrm.Utility.closeProgressIndicator();
            console.error("Error confirming project selection:", error);
            this.setState({
                error: this.strings.ContactAdministrator,
                loading: false,
            });
        }
    };

    /**
     * Handle user clicking "Back" on the project selection popup.
     * Returns to the main form.
     */
    private handleProjectBack = (): void => {
        this.setState({
            showProjectSelectionPopup: false,
            nearbyProjects: [],
            selectedProjectId: null,
            error: null,
        });
    };

    // =====================================================================
    // FIELD VISIBILITY
    // =====================================================================

    private shouldShowField = (
        field:
            | "qataryId"
            | "name"
            | "crNumber"
            | "cpNumber"
            | "registrationNumber"
            | "id"
            | "carColor"
            | "vehicleBrand"
            | "boatNumber"
            | "projectName"
            | "requestPermitNumber"
            | "locationDetails",
    ): boolean => {
        const { selectedInspectionType, isAnonymous } = this.state;
        if (!selectedInspectionType) return false;

        // Type 13 (Anonymous-like) — no fields
        if (selectedInspectionType === 13) return false;

        if (isAnonymous && [3, 6, 7].includes(selectedInspectionType)) {
            return false;
        }

        if (["id", "carColor", "vehicleBrand"].includes(field)) {
            return selectedInspectionType === 1;
        }

        if (["qataryId", "name"].includes(field)) {
            return [2, 3, 6, 15].includes(selectedInspectionType);
        }

        if (field === "crNumber") {
            return [5, 7].includes(selectedInspectionType) && this.state.crCpToggle === 'cr';
        }

        if (field === "cpNumber") {
            return [5, 7].includes(selectedInspectionType) && this.state.crCpToggle === 'cp';
        }

        if (field === "registrationNumber") {
            // 10 = Establishment, 11 = Hospital
            return [10, 11].includes(selectedInspectionType);
        }

        if (field === "boatNumber") {
            return selectedInspectionType === 14;
        }

        // Type 16 (Project) — no fields shown on main form;
        // user picks from project list popup instead
        if (["projectName", "requestPermitNumber", "locationDetails"].includes(field)) {
            return false;
        }

        return false;
    };

    private shouldShowAnonymousCheckbox = (): boolean => {
        const { selectedInspectionType } = this.state;
        return (
            selectedInspectionType !== null &&
            [3, 6, 7].includes(selectedInspectionType)
        );
    };

    // =====================================================================
    // STYLE HELPERS
    // =====================================================================

    private getStyles = () => {
        const { loading } = this.state;

        const containerStyle: React.CSSProperties = {
            position: "fixed",
            top: 0,
            left: 0,
            width: "100%",
            height: "100%",
            backgroundColor: FLUENT.overlay,
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            zIndex: 1000,
            direction: this.state.isRTL ? "rtl" : "ltr",
        };

        const modalStyle: React.CSSProperties = {
            backgroundColor: FLUENT.colorWhite,
            borderRadius: FLUENT.borderRadiusModal,
            padding: 24,
            maxWidth: 500,
            width: "90%",
            maxHeight: "90vh",
            overflowY: "auto",
            boxShadow: FLUENT.shadowModal,
        };

        const titleStyle: React.CSSProperties = {
            fontSize: 18,
            fontWeight: 600,
            marginBottom: 20,
            color: FLUENT.colorNeutralDark,
            fontFamily: FLUENT.fontFamily,
        };

        const fieldStyle: React.CSSProperties = {
            marginBottom: 16,
        };

        const labelStyle: React.CSSProperties = {
            display: "block",
            fontSize: 14,
            fontWeight: 600,
            marginBottom: 4,
            color: FLUENT.colorNeutralPrimary,
            fontFamily: FLUENT.fontFamily,
        };

        const inputStyle: React.CSSProperties = {
            width: "100%",
            padding: "6px 8px",
            border: `1px solid ${FLUENT.colorNeutralSecondary}`,
            borderRadius: FLUENT.borderRadius,
            fontSize: 14,
            fontFamily: FLUENT.fontFamily,
            boxSizing: "border-box",
            outline: "none",
            transition: `border-color ${FLUENT.transitionFast}`,
        };

        const selectWrapperStyle: React.CSSProperties = {
            position: "relative",
            display: "flex",
            alignItems: "center",
            gap: 6,
        };

        const selectInnerStyle: React.CSSProperties = {
            ...inputStyle,
            flex: 1,
            paddingRight: 28,
            appearance: "none" as any,
            backgroundImage: `url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 12 12'%3E%3Cpath fill='%23605e5c' d='M2.15 4.65a.5.5 0 01.7 0L6 7.79l3.15-3.14a.5.5 0 11.7.7l-3.5 3.5a.5.5 0 01-.7 0l-3.5-3.5a.5.5 0 010-.7z'/%3E%3C/svg%3E")`,
            backgroundRepeat: "no-repeat",
            backgroundPosition: this.state.isRTL
                ? "8px center"
                : "calc(100% - 8px) center",
        };

        const clearButtonStyle: React.CSSProperties = {
            padding: "6px 10px",
            border: "none",
            borderRadius: FLUENT.borderRadius,
            backgroundColor: FLUENT.colorErrorPrimary,
            color: FLUENT.colorWhite,
            fontSize: 12,
            fontWeight: 600,
            fontFamily: FLUENT.fontFamily,
            cursor: "pointer",
            transition: `background-color ${FLUENT.transitionFast}`,
            whiteSpace: "nowrap",
            flexShrink: 0,
        };

        const scanButtonStyle: React.CSSProperties = {
            padding: "6px",
            border: "none",
            borderRadius: FLUENT.borderRadius,
            backgroundColor: "transparent",
            color: FLUENT.colorNeutralSecondary, // Grey icon
            fontSize: 16,
            fontWeight: 600,
            fontFamily: FLUENT.fontFamily,
            cursor: loading ? "not-allowed" : "pointer",
            transition: `background-color ${FLUENT.transitionFast}, color ${FLUENT.transitionFast}`,
            whiteSpace: "nowrap",
            flexShrink: 0,
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
        };

        const buttonContainerStyle: React.CSSProperties = {
            display: "flex",
            gap: 8,
            marginTop: 24,
            justifyContent: "flex-end",
        };

        const buttonStyle: React.CSSProperties = {
            padding: "8px 16px",
            border: "1px solid transparent",
            borderRadius: FLUENT.borderRadius,
            fontSize: 14,
            fontWeight: 600,
            cursor: loading ? "not-allowed" : "pointer",
            fontFamily: FLUENT.fontFamily,
            transition: `background-color ${FLUENT.transitionFast}`,
            minWidth: 80,
        };

        const startButtonStyle: React.CSSProperties = {
            ...buttonStyle,
            backgroundColor: loading ? "#d3d3d3" : FLUENT.colorPrimary,
            color: FLUENT.colorWhite,
        };

        const closeButtonStyle: React.CSSProperties = {
            ...buttonStyle,
            backgroundColor: FLUENT.colorWhite,
            color: FLUENT.colorNeutralPrimary,
            border: `1px solid ${FLUENT.colorNeutralSecondary}`,
        };

        const errorStyle: React.CSSProperties = {
            color: FLUENT.colorErrorPrimary,
            fontSize: 12,
            marginBottom: 16,
            padding: 8,
            backgroundColor: FLUENT.colorErrorBackground,
            border: `1px solid ${FLUENT.colorErrorPrimary}`,
            borderRadius: FLUENT.borderRadius,
            fontFamily: FLUENT.fontFamily,
        };

        const checkboxContainerStyle: React.CSSProperties = {
            marginBottom: 16,
            display: "flex",
            alignItems: "center",
            gap: 8,
        };

        const checkboxStyle: React.CSSProperties = {
            width: 16,
            height: 16,
            cursor: loading ? "not-allowed" : "pointer",
        };

        const readOnlyDisplayStyle: React.CSSProperties = {
            padding: "6px 8px",
            backgroundColor: FLUENT.colorNeutralLighter,
            border: `1px solid ${FLUENT.colorNeutralLight}`,
            borderRadius: FLUENT.borderRadius,
            fontSize: 14,
            fontFamily: FLUENT.fontFamily,
            color: FLUENT.colorNeutralPrimary,
            flex: 1,
        };

        return {
            containerStyle,
            modalStyle,
            titleStyle,
            fieldStyle,
            labelStyle,
            inputStyle,
            selectWrapperStyle,
            selectInnerStyle,
            clearButtonStyle,
            scanButtonStyle,
            buttonContainerStyle,
            startButtonStyle,
            closeButtonStyle,
            errorStyle,
            checkboxContainerStyle,
            checkboxStyle,
            readOnlyDisplayStyle,
        };
    };

    // =====================================================================
    // RENDER HELPERS
    // =====================================================================

    /**
     * Render a select dropdown wrapped with a clear (X) button.
     */
    private renderSelectWithClear = (
        value: string,
        onChange: (val: string) => void,
        options: Array<{ key: string; value: string; label: string }>,
        placeholder: string,
        disabled: boolean,
        styles: ReturnType<typeof this.getStyles>,
        extraSelectStyle?: React.CSSProperties,
    ) => {
        return React.createElement(
            "div",
            { style: styles.selectWrapperStyle },
            React.createElement(
                "select",
                {
                    value: value || "",
                    onChange: (e: { target: { value: string } }) =>
                        onChange(e.target.value),
                    disabled: disabled,
                    style: {
                        ...styles.selectInnerStyle,
                        backgroundColor: disabled
                            ? FLUENT.colorNeutralLighter
                            : FLUENT.colorWhite,
                        cursor: disabled ? "not-allowed" : "pointer",
                        ...extraSelectStyle,
                    },
                },
                React.createElement("option", { value: "" }, placeholder),
                options.map((opt) =>
                    React.createElement(
                        "option",
                        { key: opt.key, value: opt.value },
                        opt.label,
                    ),
                ),
            ),
            // Clear button — colored button beside the field, shown when value is selected and not disabled
            value &&
            !disabled &&
            React.createElement(
                "button",
                {
                    onClick: () => onChange(""),
                    style: styles.clearButtonStyle,
                    title: this.strings.Clear,
                    type: "button",
                },
                this.strings.Clear,
            ),
        );
    };

    // =====================================================================
    // RENDER
    // =====================================================================

    render() {
        if (!this.props.isOpen) {
            return null;
        }

        const {
            inspectionTypes,
            selectedInspectionType,
            qataryId,
            name,
            crNumber,
            registrationNumber,
            id,
            carColor,
            vehicleBrand,
            vehicleBrands,
            loading,
            error,
            showCampaignIncidentPopup,
            campaigns,
            incidentTypes,
            selectedCampaignId,
            selectedIncidentTypeId,
            isAnonymous,
            incidentTypeReadOnly,
        } = this.state;

        const styles = this.getStyles();
        const isInspectionTypeDisabled = this.props.lockInspectionType || loading;

        // ------
        // Project Selection Popup (type 16)
        // ------
        if (this.state.showProjectSelectionPopup) {
            const { nearbyProjects, selectedProjectId } = this.state;

            // Styles specific to the project selection list
            const projectListStyle: React.CSSProperties = {
                maxHeight: "50vh",
                overflowY: "auto",
                marginBottom: 16,
                border: `1px solid ${FLUENT.colorNeutralLight}`,
                borderRadius: FLUENT.borderRadius,
            };

            const projectCardStyle = (isSelected: boolean): React.CSSProperties => ({
                display: "flex",
                alignItems: "center",
                gap: 12,
                padding: "12px 16px",
                borderBottom: `1px solid ${FLUENT.colorNeutralLighter}`,
                backgroundColor: isSelected ? "#e6f2ff" : FLUENT.colorWhite,
                cursor: loading ? "not-allowed" : "pointer",
                transition: `background-color ${FLUENT.transitionFast}`,
            });

            const radioStyle: React.CSSProperties = {
                width: 18,
                height: 18,
                accentColor: FLUENT.colorPrimary,
                cursor: loading ? "not-allowed" : "pointer",
                flexShrink: 0,
            };

            const projectInfoStyle: React.CSSProperties = {
                flex: 1,
                display: "flex",
                flexDirection: "column",
                gap: 2,
            };

            const projectNumberStyle: React.CSSProperties = {
                fontSize: 14,
                fontWeight: 600,
                color: FLUENT.colorNeutralDark,
                fontFamily: FLUENT.fontFamily,
            };

            const projectNameStyle: React.CSSProperties = {
                fontSize: 13,
                color: FLUENT.colorNeutralPrimary,
                fontFamily: FLUENT.fontFamily,
            };

            const projectAccountStyle: React.CSSProperties = {
                fontSize: 12,
                color: FLUENT.colorNeutralSecondary,
                fontFamily: FLUENT.fontFamily,
            };

            return React.createElement(
                "div",
                { style: styles.containerStyle },
                React.createElement(
                    "div",
                    { style: styles.modalStyle },
                    // Title
                    React.createElement(
                        "h2",
                        { style: styles.titleStyle },
                        this.strings.SelectProject,
                    ),

                    // Error message
                    error &&
                    React.createElement("div", { style: styles.errorStyle }, error),

                    // No projects found message
                    nearbyProjects.length === 0 &&
                    React.createElement(
                        "div",
                        {
                            style: {
                                padding: 24,
                                textAlign: "center",
                                color: FLUENT.colorNeutralSecondary,
                                fontFamily: FLUENT.fontFamily,
                                fontSize: 14,
                            },
                        },
                        this.strings.NoProjectsFound,
                    ),

                    // Project list
                    nearbyProjects.length > 0 &&
                    React.createElement(
                        "div",
                        { style: projectListStyle },
                        nearbyProjects.map((project) =>
                            React.createElement(
                                "div",
                                {
                                    key: project.projectId,
                                    style: projectCardStyle(selectedProjectId === project.projectId),
                                    onClick: () => !loading && this.handleProjectSelect(project.projectId),
                                },
                                // Radio button
                                React.createElement("input", {
                                    type: "radio",
                                    name: "projectSelection",
                                    checked: selectedProjectId === project.projectId,
                                    onChange: () => this.handleProjectSelect(project.projectId),
                                    disabled: loading,
                                    style: radioStyle,
                                }),
                                // Project info
                                React.createElement(
                                    "div",
                                    { style: projectInfoStyle },
                                    React.createElement(
                                        "span",
                                        { style: projectNumberStyle },
                                        this.strings.ProjectNumber + ": " + project.projectNumber,
                                    ),
                                    React.createElement(
                                        "span",
                                        { style: projectNameStyle },
                                        project.arabicName,
                                    ),
                                    React.createElement(
                                        "span",
                                        { style: projectAccountStyle },
                                        this.strings.AccountNameLabel + ": " + project.accountName,
                                    ),
                                ),
                            ),
                        ),
                    ),

                    // Buttons
                    React.createElement(
                        "div",
                        { style: styles.buttonContainerStyle },
                        // Select button (primary)
                        React.createElement(
                            "button",
                            {
                                onClick: this.handleProjectConfirm,
                                disabled: loading || !selectedProjectId,
                                style: {
                                    ...styles.startButtonStyle,
                                    opacity: !selectedProjectId ? 0.5 : 1,
                                },
                            },
                            loading ? this.strings.Loading : this.strings.Select,
                        ),
                        // Back button (secondary)
                        React.createElement(
                            "button",
                            {
                                onClick: this.handleProjectBack,
                                disabled: loading,
                                style: styles.closeButtonStyle,
                            },
                            this.strings.Back,
                        ),
                    ),
                ),
            );
        }

        // ------
        // Campaign/Incident Popup
        // ------
        if (showCampaignIncidentPopup) {
            const showCampaignField = this.state.popupShowCampaign;
            const showIncidentField = this.state.popupShowIncidentType;

            let popupTitle = "";
            if (showCampaignField && showIncidentField) {
                popupTitle =
                    this.strings.SelectCampaign + " / " + this.strings.SelectIncidentType;
            } else if (showCampaignField) {
                popupTitle = this.strings.SelectCampaign;
            } else if (showIncidentField) {
                popupTitle = this.strings.SelectIncidentType;
            }

            return React.createElement(
                "div",
                { style: styles.containerStyle },
                React.createElement(
                    "div",
                    { style: styles.modalStyle },
                    React.createElement("h2", { style: styles.titleStyle }, popupTitle),

                    error &&
                    React.createElement("div", { style: styles.errorStyle }, error),

                    // Campaign Selection
                    showCampaignField &&
                    React.createElement(
                        "div",
                        { style: styles.fieldStyle },
                        React.createElement(
                            "label",
                            { style: styles.labelStyle },
                            this.strings.Campaign,
                        ),
                        this.renderSelectWithClear(
                            selectedCampaignId || "",
                            (val) => this.handleCampaignChange(val),
                            campaigns.map((c) => ({
                                key: c.id,
                                value: c.id,
                                label: c.name,
                            })),
                            "--",
                            loading,
                            styles,
                        ),
                    ),

                    // Incident Type Selection
                    showIncidentField &&
                    React.createElement(
                        "div",
                        { style: styles.fieldStyle },
                        React.createElement(
                            "label",
                            { style: styles.labelStyle },
                            this.strings.IncidentType + " *",
                        ),
                        this.renderSelectWithClear(
                            selectedIncidentTypeId || "",
                            (val) => this.handleIncidentTypeChange(val),
                            incidentTypes.map((it) => ({
                                key: it.id,
                                value: it.id,
                                label: it.name,
                            })),
                            "--",
                            loading || incidentTypeReadOnly,
                            styles,
                        ),
                    ),

                    React.createElement(
                        "div",
                        { style: styles.buttonContainerStyle },
                        React.createElement(
                            "button",
                            {
                                onClick: this.handleContinueWithSelections,
                                disabled: loading,
                                style: styles.startButtonStyle,
                            },
                            loading ? this.strings.Loading : this.strings.Continue,
                        ),
                        React.createElement(
                            "button",
                            {
                                onClick: () =>
                                    this.setState({
                                        showCampaignIncidentPopup: false,
                                        incidentTypeReadOnly: false,
                                        // CLEAR selections on close so they don't show on main form
                                        selectedCampaignId: undefined,
                                        selectedCampaignName: undefined,
                                        selectedIncidentTypeId: undefined,
                                        selectedIncidentTypeName: undefined,
                                    }),
                                disabled: loading,
                                style: styles.closeButtonStyle,
                            },
                            this.strings.Close,
                        ),
                    ),
                ),
            );
        }

        // ------
        // Main Form
        // ------
        return React.createElement(
            "div",
            { style: styles.containerStyle },
            React.createElement(
                "div",
                { style: styles.modalStyle },
                React.createElement(
                    "h2",
                    { style: styles.titleStyle },
                    this.strings.StartMultiTypeInspection,
                ),

                error &&
                React.createElement("div", { style: styles.errorStyle }, error),


                // Inspection Type
                React.createElement(
                    "div",
                    { style: styles.fieldStyle },
                    React.createElement(
                        "label",
                        { style: styles.labelStyle },
                        this.strings.InspectionType + " *",
                    ),
                    this.renderSelectWithClear(
                        this.state.accountTypeRecord?.duc_organizationunitaccounttypesid || this.state.accountTypeRecord?.duc_accounttypeid || "",
                        (val) => {
                            // Build a synthetic event for the existing handler
                            const syntheticEvent = {
                                target: { value: val },
                            } as React.ChangeEvent<HTMLSelectElement>;
                            this.handleInspectionTypeChange(syntheticEvent);
                        },
                        inspectionTypes.map((t) => ({
                            key: t.orgUnitAccountTypeId || t.accountTypeId,
                            value: t.orgUnitAccountTypeId || t.accountTypeId || "",
                            label: t.label,
                        })),
                        this.strings.chooseInspectionType,
                        isInspectionTypeDisabled,
                        styles,
                    ),
                ),

                // Anonymous Checkbox (for types 3, 6, 7)
                this.shouldShowAnonymousCheckbox() &&
                React.createElement(
                    "div",
                    { style: styles.checkboxContainerStyle },
                    React.createElement("input", {
                        type: "checkbox",
                        checked: isAnonymous,
                        onChange: (e: React.ChangeEvent<HTMLInputElement>) => this.setState({ isAnonymous: e.target.checked }),
                        disabled: loading,
                        style: styles.checkboxStyle,
                    }),
                    React.createElement(
                        "label",
                        {
                            style: {
                                ...styles.labelStyle,
                                marginBottom: 0,
                                cursor: loading ? "not-allowed" : "pointer",
                            } as React.CSSProperties,
                            onClick: () =>
                                !loading && this.setState({ isAnonymous: !isAnonymous }),
                        },
                        this.strings.Anonymous,
                    ),
                ),

                // ID (Vehicle)
                this.shouldShowField("id") &&
                React.createElement(
                    "div",
                    { style: styles.fieldStyle },
                    React.createElement(
                        "label",
                        { style: styles.labelStyle },
                        this.strings.PlateNumber + " *",
                    ),
                    React.createElement("input", {
                        type: "text",
                        value: id,
                        onChange: (e: React.ChangeEvent<HTMLInputElement>) => this.handleInputChange("id", e.target.value),
                        disabled: loading,
                        style: styles.inputStyle,
                    }),
                ),

                // Car Color
                this.shouldShowField("carColor") &&
                React.createElement(
                    "div",
                    { style: styles.fieldStyle },
                    React.createElement(
                        "label",
                        { style: styles.labelStyle },
                        this.strings.CarColor,
                    ),
                    React.createElement("input", {
                        type: "text",
                        value: carColor,
                        onChange: (e: React.ChangeEvent<HTMLInputElement>) =>
                            this.handleInputChange("carColor", e.target.value),
                        disabled: loading,
                        style: styles.inputStyle,
                    }),
                ),

                // Vehicle Brand
                this.shouldShowField("vehicleBrand") &&
                React.createElement(
                    "div",
                    { style: styles.fieldStyle },
                    React.createElement(
                        "label",
                        { style: styles.labelStyle },
                        this.strings.VehicleBrand,
                    ),
                    this.renderSelectWithClear(
                        vehicleBrand !== null ? String(vehicleBrand) : "",
                        (val) => {
                            const parsed = val ? parseInt(val, 10) : null;
                            this.setState({ vehicleBrand: parsed });
                        },
                        vehicleBrands.map((vb) => ({
                            key: String(vb.value),
                            value: String(vb.value),
                            label: vb.label,
                        })),
                        "--",
                        loading,
                        styles,
                    ),
                ),

                // Qatary ID with Barcode Scanner
                this.shouldShowField("qataryId") &&
                React.createElement(
                    "div",
                    { style: styles.fieldStyle },
                    React.createElement(
                        "label",
                        { style: styles.labelStyle },
                        this.strings.QataryID + " *",
                    ),
                    React.createElement(
                        "div",
                        { style: { display: "flex", alignItems: "center", gap: 6 } },
                        React.createElement("input", {
                            type: "text",
                            value: qataryId,
                            onChange: (e: React.ChangeEvent<HTMLInputElement>) =>
                                this.handleNumericInputChange("qataryId", e.target.value),
                            disabled: loading,
                            style: { ...styles.inputStyle, flex: 1 },
                            placeholder: "01234567891",
                            maxLength: 11,
                            inputMode: "numeric" as any,
                        }),
                        React.createElement(
                            "button",
                            {
                                onClick: () => this.handleScanBarcode("qataryId"),
                                disabled: loading,
                                style: styles.scanButtonStyle,
                                type: "button",
                                title: this.strings.ScanBarcode,
                            },
                            // SVG Barcode Icon
                            React.createElement(
                                "svg",
                                {
                                    width: "20",
                                    height: "20",
                                    viewBox: "0 0 24 24",
                                    fill: "none",
                                    stroke: "currentColor",
                                    strokeWidth: "2",
                                    strokeLinecap: "round",
                                    strokeLinejoin: "round",
                                },
                                React.createElement("path", { d: "M3 5v14" }),
                                React.createElement("path", { d: "M8 5v14" }),
                                React.createElement("path", { d: "M12 5v14" }),
                                React.createElement("path", { d: "M17 5v14" }),
                                React.createElement("path", { d: "M21 5v14" }),
                            ),
                        ),
                    ),
                ),

                // Name
                this.shouldShowField("name") &&
                React.createElement(
                    "div",
                    { style: styles.fieldStyle },
                    React.createElement(
                        "label",
                        { style: styles.labelStyle },
                        this.strings.Name + " *",
                    ),
                    React.createElement("input", {
                        type: "text",
                        value: name,
                        onChange: (e: React.ChangeEvent<HTMLInputElement>) => this.handleInputChange("name", e.target.value),
                        disabled: loading,
                        style: styles.inputStyle,
                    }),
                ),

                // CR / CP Toggle (Company=5, Manor=7)
                [5, 7].includes(selectedInspectionType!) &&
                React.createElement(
                    "div",
                    { style: { ...styles.fieldStyle, marginBottom: 12 } },
                    React.createElement(
                        "div",
                        {
                            style: {
                                display: "flex",
                                borderRadius: FLUENT.borderRadius,
                                overflow: "hidden",
                                border: `1px solid ${FLUENT.colorNeutralSecondary}`,
                                width: "fit-content",
                            },
                        },
                        React.createElement(
                            "button",
                            {
                                type: "button",
                                disabled: loading,
                                onClick: () => this.setState({ crCpToggle: 'cr', cpNumber: "" }),
                                style: {
                                    padding: "6px 20px",
                                    border: "none",
                                    cursor: loading ? "not-allowed" : "pointer",
                                    fontFamily: FLUENT.fontFamily,
                                    fontSize: 14,
                                    fontWeight: 600,
                                    backgroundColor: this.state.crCpToggle === 'cr' ? FLUENT.colorPrimary : FLUENT.colorWhite,
                                    color: this.state.crCpToggle === 'cr' ? FLUENT.colorWhite : FLUENT.colorNeutralPrimary,
                                    transition: `background-color ${FLUENT.transitionFast}`,
                                },
                            },
                            "CR",
                        ),
                        React.createElement(
                            "button",
                            {
                                type: "button",
                                disabled: loading,
                                onClick: () => this.setState({ crCpToggle: 'cp', crNumber: "" }),
                                style: {
                                    padding: "6px 20px",
                                    border: "none",
                                    borderLeft: `1px solid ${FLUENT.colorNeutralSecondary}`,
                                    cursor: loading ? "not-allowed" : "pointer",
                                    fontFamily: FLUENT.fontFamily,
                                    fontSize: 14,
                                    fontWeight: 600,
                                    backgroundColor: this.state.crCpToggle === 'cp' ? FLUENT.colorPrimary : FLUENT.colorWhite,
                                    color: this.state.crCpToggle === 'cp' ? FLUENT.colorWhite : FLUENT.colorNeutralPrimary,
                                    transition: `background-color ${FLUENT.transitionFast}`,
                                },
                            },
                            "CP",
                        ),
                    ),
                ),

                // CR Number (shown when toggle = CR)
                this.shouldShowField("crNumber") &&
                React.createElement(
                    "div",
                    { style: styles.fieldStyle },
                    React.createElement(
                        "label",
                        { style: styles.labelStyle },
                        (selectedInspectionType === 7
                            ? this.strings.MonourNumber
                            : this.strings.CRNumber) + " *",
                    ),
                    React.createElement(
                        "div",
                        { style: { display: "flex", alignItems: "center", gap: 6 } },
                        React.createElement("input", {
                            type: "text",
                            value: crNumber,
                            onChange: (e: React.ChangeEvent<HTMLInputElement>) =>
                                this.handleNumericInputChange("crNumber", e.target.value),
                            disabled: loading,
                            style: { ...styles.inputStyle, flex: 1 },
                            placeholder: "0123456789",
                            inputMode: "url" as any,
                        }),
                        React.createElement(
                            "button",
                            {
                                onClick: () => this.handleScanBarcode("crNumber"),
                                disabled: loading,
                                style: styles.scanButtonStyle,
                                type: "button",
                                title: this.strings.ScanBarcode,
                            },
                            React.createElement(
                                "svg",
                                {
                                    width: "20",
                                    height: "20",
                                    viewBox: "0 0 24 24",
                                    fill: "none",
                                    stroke: "currentColor",
                                    strokeWidth: "2",
                                    strokeLinecap: "round",
                                    strokeLinejoin: "round",
                                },
                                React.createElement("path", { d: "M3 5v14" }),
                                React.createElement("path", { d: "M8 5v14" }),
                                React.createElement("path", { d: "M12 5v14" }),
                                React.createElement("path", { d: "M17 5v14" }),
                                React.createElement("path", { d: "M21 5v14" }),
                            ),
                        ),
                    ),
                ),

                // CP Number (shown when toggle = CP)
                this.shouldShowField("cpNumber") &&
                React.createElement(
                    "div",
                    { style: styles.fieldStyle },
                    React.createElement(
                        "label",
                        { style: styles.labelStyle },
                        this.strings.CPNumber + " *",
                    ),
                    React.createElement(
                        "div",
                        { style: { display: "flex", alignItems: "center", gap: 6 } },
                        React.createElement("input", {
                            type: "text",
                            value: this.state.cpNumber,
                            onChange: (e: React.ChangeEvent<HTMLInputElement>) =>
                                this.handleNumericInputChange("cpNumber", e.target.value),
                            disabled: loading,
                            style: { ...styles.inputStyle, flex: 1 },
                            placeholder: "0123456789",
                            inputMode: "url" as any,
                        }),
                        React.createElement(
                            "button",
                            {
                                onClick: () => this.handleScanBarcode("cpNumber"),
                                disabled: loading,
                                style: styles.scanButtonStyle,
                                type: "button",
                                title: this.strings.ScanBarcode,
                            },
                            React.createElement(
                                "svg",
                                {
                                    width: "20",
                                    height: "20",
                                    viewBox: "0 0 24 24",
                                    fill: "none",
                                    stroke: "currentColor",
                                    strokeWidth: "2",
                                    strokeLinecap: "round",
                                    strokeLinejoin: "round",
                                },
                                React.createElement("path", { d: "M3 5v14" }),
                                React.createElement("path", { d: "M8 5v14" }),
                                React.createElement("path", { d: "M12 5v14" }),
                                React.createElement("path", { d: "M17 5v14" }),
                                React.createElement("path", { d: "M21 5v14" }),
                            ),
                        ),
                    ),
                ),

                this.shouldShowField("registrationNumber") &&
                React.createElement(
                    "div",
                    { style: styles.fieldStyle },
                    React.createElement(
                        "label",
                        { style: styles.labelStyle },
                        this.strings.RegistrationNumber + " *",
                    ),
                    React.createElement(
                        "div",
                        { style: { display: "flex", alignItems: "center", gap: 6 } },
                        React.createElement("input", {
                            type: "text",
                            value: registrationNumber,
                            onChange: (e: React.ChangeEvent<HTMLInputElement>) =>
                                this.handleNumericInputChange("registrationNumber", e.target.value),
                            disabled: loading,
                            style: { ...styles.inputStyle, flex: 1 },
                            placeholder: "0123456789",
                            inputMode: "numeric" as any,
                        }),
                        React.createElement(
                            "button",
                            {
                                onClick: () => this.handleScanBarcode("registrationNumber"),
                                disabled: loading,
                                style: styles.scanButtonStyle,
                                type: "button",
                                title: this.strings.ScanBarcode,
                            },
                            // SVG Barcode Icon
                            React.createElement(
                                "svg",
                                {
                                    width: "20",
                                    height: "20",
                                    viewBox: "0 0 24 24",
                                    fill: "none",
                                    stroke: "currentColor",
                                    strokeWidth: "2",
                                    strokeLinecap: "round",
                                    strokeLinejoin: "round",
                                },
                                React.createElement("path", { d: "M3 5v14" }),
                                React.createElement("path", { d: "M8 5v14" }),
                                React.createElement("path", { d: "M12 5v14" }),
                                React.createElement("path", { d: "M17 5v14" }),
                                React.createElement("path", { d: "M21 5v14" }),
                            ),
                        ),
                    ),
                ),

                // Boat Number (type 14)
                this.shouldShowField("boatNumber") &&
                React.createElement(
                    "div",
                    { style: styles.fieldStyle },
                    React.createElement(
                        "label",
                        { style: styles.labelStyle },
                        this.strings.BoatNumber + " *",
                    ),
                    React.createElement("input", {
                        type: "text",
                        value: this.state.boatNumber,
                        onChange: (e: React.ChangeEvent<HTMLInputElement>) =>
                            this.handleInputChange("boatNumber", e.target.value),
                        disabled: loading,
                        style: styles.inputStyle,
                    }),
                ),

                // Project Name (type 16)
                this.shouldShowField("projectName") &&
                React.createElement(
                    "div",
                    { style: styles.fieldStyle },
                    React.createElement(
                        "label",
                        { style: styles.labelStyle },
                        this.strings.ProjectName,
                    ),
                    React.createElement("input", {
                        type: "text",
                        value: this.state.projectName,
                        onChange: (e: React.ChangeEvent<HTMLInputElement>) =>
                            this.handleInputChange("projectName", e.target.value),
                        disabled: loading,
                        style: styles.inputStyle,
                    }),
                ),

                // Request/Permit Number (type 16)
                this.shouldShowField("requestPermitNumber") &&
                React.createElement(
                    "div",
                    { style: styles.fieldStyle },
                    React.createElement(
                        "label",
                        { style: styles.labelStyle },
                        this.strings.RequestPermitNumber + " *",
                    ),
                    React.createElement("input", {
                        type: "text",
                        value: this.state.requestPermitNumber,
                        onChange: (e: React.ChangeEvent<HTMLInputElement>) =>
                            this.handleInputChange("requestPermitNumber", e.target.value),
                        disabled: loading,
                        style: styles.inputStyle,
                    }),
                ),

                // Location Details (type 16)
                this.shouldShowField("locationDetails") &&
                React.createElement(
                    "div",
                    { style: styles.fieldStyle },
                    React.createElement(
                        "label",
                        { style: styles.labelStyle },
                        this.strings.LocationDetails,
                    ),
                    React.createElement("textarea", {
                        value: this.state.locationDetails,
                        onChange: (e: React.ChangeEvent<HTMLTextAreaElement>) =>
                            this.handleInputChange("locationDetails", e.target.value),
                        disabled: loading,
                        style: { ...styles.inputStyle, minHeight: "60px", resize: "vertical" },
                    }),
                ),

                // Buttons
                React.createElement(
                    "div",
                    { style: styles.buttonContainerStyle },
                    React.createElement(
                        "button",
                        {
                            onClick: this.handleStart,
                            disabled: loading,
                            style: styles.startButtonStyle,
                        },
                        loading ? this.strings.Loading : this.strings.Start,
                    ),
                    React.createElement(
                        "button",
                        {
                            onClick: this.props.onClose,
                            disabled: loading,
                            style: styles.closeButtonStyle,
                        },
                        this.strings.Close,
                    ),
                ),
            ),
        );
    }
}
