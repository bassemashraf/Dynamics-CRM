/* eslint-disable */
import * as React from "react";
import {
  WorkOrderHelpers,
  CampaignHelpers,
  IncidentTypeHelpers,
  InitCache,
} from "../helpers";

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
  organizationUnitId?: string;
  organizationUnitName?: string;
  defaultInspectionType?: number;
  lockInspectionType?: boolean;
}

interface IMultiTypeInspectionState {
  isRTL: boolean;
  inspectionTypes: Array<{ value: number; label: string }>;
  selectedInspectionType: number | null;
  qataryId: string;
  name: string;
  crNumber: string;
  id: string;
  carColor: string;
  vehicleBrand: number | null;
  vehicleBrands: Array<{ value: number; label: string }>;
  loading: boolean;
  error: string | null;
  accountTypeRecord: any | null;
  // Campaign/Incident popup
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
  incidentTypeReadOnly: boolean;
  campaignIncidentTypeMap: Record<string, { id: string; name: string }>;
}

interface LocalizedStrings {
  StartMultiTypeInspection: string;
  InspectionType: string;
  QataryID: string;
  Name: string;
  CRNumber: string;
  MonourNumber: string;
  ID: string;
  CarColor: string;
  VehicleBrand: string;
  Start: string;
  Close: string;
  Loading: string;
  PleaseSelectInspectionType: string;
  PleaseEnterRequiredFields: string;
  SearchingAccount: string;
  CreatingAccount: string;
  CreatingWorkOrder: string;
  CreatingBooking: string;
  Error: string;
  AccountNotFound: string;
  AccountCreated: string;
  chooseInspectionType: string;
  SelectCampaign: string;
  SelectIncidentType: string;
  Campaign: string;
  IncidentType: string;
  Continue: string;
  Anonymous: string;
  ScanBarcode: string;
  Clear: string;
}

// Cache constants
const INSPECTION_TYPES_CACHE_KEY = "MOCI_InspectionTypes_Cache";
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
  // Cache account ID to avoid re-searching in handleContinueWithSelections
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
      SearchingAccount: props.context.resources.getString("SearchingAccount"),
      CreatingAccount:
        props.context.resources.getString("CreatingAccount") ||
        "Creating Account...",
      CreatingWorkOrder:
        props.context.resources.getString("CreatingWorkOrder") ||
        "Creating Work Order...",
      CreatingBooking:
        props.context.resources.getString("CreatingBooking") ||
        "Creating Booking...",
      Error: props.context.resources.getString("Error"),
      AccountNotFound: props.context.resources.getString("AccountNotFound"),
      AccountCreated: props.context.resources.getString("AccountCreated"),
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
      ScanBarcode:
        props.context.resources.getString("ScanBarcode") || "Scan Barcode",
      Clear: props.context.resources.getString("Clear") || "Clear",
    };

    this.state = {
      isRTL: isRTL,
      inspectionTypes: [],
      selectedInspectionType: props.defaultInspectionType || null,
      qataryId: "",
      name: "",
      crNumber: "",
      id: "",
      carColor: "",
      vehicleBrand: null,
      vehicleBrands: [],
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
    };
  }

  async componentDidMount(): Promise<void> {
    try {
      const userId =
        this.xrm.Utility.getGlobalContext().userSettings.userId.replace(
          /[{}]/g,
          "",
        );

      // Parallel loading
      await Promise.all([
        this.loadInspectionTypes(),
        this.loadVehicleTypes(),
        this.preloadCampaignsAndIncidentTypes(),
        InitCache.load(userId),
      ]);

      // If default inspection type is provided, retrieve account type
      if (this.props.defaultInspectionType) {
        const accountTypeRecord = await this.retrieveAccountTypeByOptionSet(
          this.props.defaultInspectionType,
        );
        this.setState({ accountTypeRecord });
      }
    } catch (error: any) {
      console.error("Error in componentDidMount:", error);
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
  // INSPECTION TYPES
  // =====================================================================

  private loadInspectionTypes = async (): Promise<void> => {
    try {
      const cachedTypes = this.getFromCache<
        Array<{ value: number; label: string }>
      >(INSPECTION_TYPES_CACHE_KEY);
      if (cachedTypes) {
        this.setState({ inspectionTypes: cachedTypes });
        return;
      }

      const entityMetadata = await this.xrm.Utility.getEntityMetadata(
        "account",
        ["duc_accountinspectiontype"],
      );
      const attribute = (entityMetadata as any).Attributes.get(
        "duc_accountinspectiontype",
      );

      if (attribute?.OptionSet) {
        const types = Object.values(attribute.OptionSet).map((opt: any) => ({
          value: opt.value,
          label: opt.text,
        }));

        this.saveToCache(INSPECTION_TYPES_CACHE_KEY, types);
        this.setState({ inspectionTypes: types });
      }
    } catch (error: any) {
      console.error("Error loading inspection types:", error);
      this.setState({ error: "Failed to load inspection types" });
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
  // ACCOUNT TYPE LOOKUP
  // =====================================================================

  private retrieveAccountTypeByOptionSet = async (
    optionSetValue: number,
  ): Promise<any | null> => {
    try {
      const result = await this.xrm.WebApi.retrieveMultipleRecords(
        "duc_accounttype",
        `?$select=duc_accounttypeid,duc_name,duc_accounttype&$filter=duc_accounttype eq ${optionSetValue}&$top=1`,
      );

      if (result?.entities?.length > 0) {
        return result.entities[0];
      }
      return null;
    } catch (error: any) {
      console.error("Error retrieving account type:", error);
      return null;
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

      // Ensure props.activePatrolId is in the list
      if (this.props.activePatrolId && this.props.activePatrolName) {
        const exists = campaigns.some(
          (c) => c.id === this.props.activePatrolId,
        );
        if (!exists) {
          campaigns.push({
            id: this.props.activePatrolId,
            name: this.props.activePatrolName,
          });
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
          // Skip individual campaign failures silently
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
      const query = `?$filter=_duc_organizationalunitid_value eq '${this.props.organizationUnitId}' and duc_campaigntype eq 100000000 and statecode eq 0&$select=new_inspectioncampaignid,new_name&$orderby=new_name asc`;
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

    const cacheKey = `${INCIDENT_TYPES_CACHE_KEY}_${this.props.organizationUnitId}`;
    const cached =
      this.getFromCache<Array<{ id: string; name: string }>>(cacheKey);
    if (cached) return cached;

    try {
      const query = `?$filter=_duc_organizationalunitid_value eq '${this.props.organizationUnitId}'&$select=msdyn_incidenttypeid,msdyn_name&$orderby=msdyn_name asc`;
      const results = await this.xrm.WebApi.retrieveMultipleRecords(
        "msdyn_incidenttype",
        query,
      );

      const incidentTypes = results.entities.map((entity: any) => ({
        id: entity.msdyn_incidenttypeid,
        name: entity.msdyn_name,
      }));

      this.saveToCache(cacheKey, incidentTypes);
      return incidentTypes;
    } catch (error: any) {
      console.error("Error loading incident types:", error);
      const message = `${error?.message || JSON.stringify(error)}`;
      alert(message);
      return [];
    }
  };

  // =====================================================================
  // HANDLERS
  // =====================================================================

  private handleInspectionTypeChange = async (
    e: React.ChangeEvent<HTMLSelectElement>,
  ): Promise<void> => {
    const value = e.target.value ? parseInt(e.target.value, 10) : null;

    this.setState({
      selectedInspectionType: value,
      qataryId: "",
      name: "",
      crNumber: "",
      id: "",
      carColor: "",
      vehicleBrand: null,
      error: null,
      accountTypeRecord: null,
      isAnonymous: false,
    });

    if (value !== null) {
      const accountTypeRecord =
        await this.retrieveAccountTypeByOptionSet(value);
      this.setState({ accountTypeRecord });
    }
  };

  private handleInputChange = (
    field: keyof IMultiTypeInspectionState,
    value: any,
  ): void => {
    this.setState({ [field]: value } as any);
  };

  // =====================================================================
  // BARCODE SCANNER
  // =====================================================================

  private handleScanBarcode = async (
    field: "qataryId" | "crNumber" | "id",
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
    }
  };

  // =====================================================================
  // CAMPAIGN ↔ INCIDENT TYPE POPUP HANDLERS
  // =====================================================================

  private handleCampaignChange = (campaignId: string): void => {
    if (!campaignId) {
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
      requiredFields.push("id", "carColor", "vehicleBrand");
    } else if (
      selectedInspectionType === 2 ||
      selectedInspectionType === 3 ||
      selectedInspectionType === 6
    ) {
      requiredFields.push("qataryId", "name");
    } else if (selectedInspectionType === 5 || selectedInspectionType === 7) {
      requiredFields.push("crNumber");
    }

    return requiredFields;
  };

  private validateFields = (): boolean => {
    const { selectedInspectionType, isAnonymous } = this.state;

    if (!selectedInspectionType) {
      this.setState({ error: this.strings.PleaseSelectInspectionType });
      return false;
    }

    // Type 4 (Anonymous) — no field validation needed
    if (selectedInspectionType === 4) {
      return true;
    }

    // Anonymous checkbox for types 3, 6, 7 — skip field validation
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

    return true;
  };

  // =====================================================================
  // ACCOUNT SEARCH / CREATE
  // =====================================================================

  private getIdentifierValue = (): string => {
    const { selectedInspectionType, qataryId, crNumber, id } = this.state;

    if (selectedInspectionType === 1) return id;
    if (
      selectedInspectionType === 2 ||
      selectedInspectionType === 3 ||
      selectedInspectionType === 6
    )
      return qataryId;
    if (selectedInspectionType === 5 || selectedInspectionType === 7)
      return crNumber;

    return "";
  };

  private getAccountName = (): string => {
    const {
      selectedInspectionType,
      name,
      qataryId,
      crNumber,
      id,
      carColor,
      vehicleBrand,
      vehicleBrands,
    } = this.state;

    if (selectedInspectionType === 1) {
      const brandLabel =
        vehicleBrand !== null
          ? vehicleBrands.find((v) => v.value === vehicleBrand)?.label ||
            String(vehicleBrand)
          : "";
      return `Vehicle ${id} ${carColor} ${brandLabel}`.trim();
    }
    if (selectedInspectionType === 2)
      return `Individual ${qataryId} ${name}`.trim();
    if (selectedInspectionType === 3) return name || `Cabin ${qataryId}`;
    if (selectedInspectionType === 6)
      return name || `Wilderness Camp ${qataryId}`;
    if (selectedInspectionType === 5) return `Company ${crNumber}`.trim();
    if (selectedInspectionType === 7) return `Manor ${crNumber}`.trim();
    if (selectedInspectionType === 4) return "Anonymous Account";

    return "Account";
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

      await this.xrm.WebApi.createRecord("duc_addressinformation", addressData);
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

      // Anonymous type (4) — use unknown account from props
      if (selectedInspectionType === 4) {
        if (this.props.unknownAccountId) {
          return this.props.unknownAccountId;
        } else {
          throw new Error("No anonymous account available");
        }
      }

      // Anonymous checkbox for types 3, 6, 7 — use unknown account
      if (isAnonymous && [3, 6, 7].includes(selectedInspectionType!)) {
        if (this.props.unknownAccountId) {
          return this.props.unknownAccountId;
        } else {
          throw new Error("No anonymous account available");
        }
      }

      if (!identifierValue) {
        throw new Error(this.strings.PleaseEnterRequiredFields);
      }

      // Search for existing account by identifier AND account type
      let filterQuery = `duc_accountidentifier eq '${identifierValue}'`;
      if (accountTypeRecord?.duc_accounttypeid) {
        filterQuery += ` and _duc_newaccounttype_value eq ${accountTypeRecord.duc_accounttypeid}`;
      }

      const searchResults = await this.xrm.WebApi.retrieveMultipleRecords(
        "account",
        `?$select=accountid,name,duc_vehicletype&$filter=${filterQuery}`,
      );

      if (searchResults?.entities?.length > 0) {
        const accountId = searchResults.entities[0].accountid;
        const accountName = searchResults.entities[0].name;
        const existingVehicleType = searchResults.entities[0].duc_vehicletype;

        if (existingVehicleType !== undefined && existingVehicleType !== null) {
          this.setState({ vehicleBrand: existingVehicleType });
        }

        // For vehicles (type 1), update color and brand on the existing record
        if (selectedInspectionType === 1) {
          const updateData: any = {};
          if (carColor) updateData.duc_vehiclecolor = carColor;
          if (vehicleBrand !== null && vehicleBrand !== undefined)
            updateData.duc_vehicletype = vehicleBrand;

          if (Object.keys(updateData).length > 0) {
            await this.xrm.WebApi.updateRecord(
              "account",
              accountId,
              updateData,
            );
          }
        }

        // Existing account — skip address creation (matches second component logic)
        return accountId;
      }

      // Create new account
      const accountName = this.getAccountName();
      const newAccount: any = {
        name: accountName,
        duc_accountidentifier: identifierValue,
        duc_accountinspectiontype: selectedInspectionType,
      };

      if (accountTypeRecord?.duc_accounttypeid) {
        newAccount["duc_NewAccountType@odata.bind"] =
          `/duc_accounttypes(${accountTypeRecord.duc_accounttypeid})`;
      }

      if (selectedInspectionType === 1) {
        if (carColor) newAccount.duc_vehiclecolor = carColor;
        if (vehicleBrand !== null && vehicleBrand !== undefined)
          newAccount.duc_vehicletype = vehicleBrand;
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
      throw error;
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

      // Get account name
      const accountRecord = await this.xrm.WebApi.retrieveRecord(
        "account",
        accountId,
        "?$select=name",
      );
      const accountName = accountRecord?.name || "";

      // Base service account data
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
        const mapped =
          this.state.campaignIncidentTypeMap[this.state.selectedCampaignId];
        if (mapped) {
          incidentTypeData = {
            id: mapped.id,
            name: mapped.name,
            entityType: "msdyn_incidenttype",
          };
        } else {
          // Fallback: fetch from API
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
        throw new Error("No incident type available");
      }

      // STEP 3 & 4: Get work order type AND department in one API call
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
        throw new Error("No department available");
      }

      // STEP 5: Prepare campaign data
      let campaignData: { id: string; name: string } | undefined;
      let parentCampaignData: { id: string; name: string } | undefined;

      if (this.state.selectedCampaignId && this.state.selectedCampaignName) {
        campaignData = {
          id: this.state.selectedCampaignId,
          name: this.state.selectedCampaignName,
        };
        parentCampaignData = campaignData;
      } else if (this.props.activePatrolId && this.props.activePatrolName) {
        campaignData = {
          id: this.props.activePatrolId,
          name: this.props.activePatrolName,
        };
        parentCampaignData = campaignData;
      }

      // STEP 6: Determine anonymous customer flag
      const anonymousCustomer =
        this.state.selectedInspectionType === 4 ||
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
        throw new Error(
          `Missing required fields: ${validation.missingFields.join(", ")}`,
        );
      }

      // STEP 9: Create work order with progress indicator
      this.xrm.Utility.showProgressIndicator(this.strings.CreatingWorkOrder);
      const workOrderId = await WorkOrderHelpers.createWorkOrder({
        subAccount: { id: accountId, name: accountName },
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
        throw new Error("Failed to create work order");
      }

      console.log("Work order created successfully:", workOrderId);

      // STEP 10: Create auto booking if from mobile
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
        formId: "eded7d77-6dc4-ed11-b596-6045bdf00fa1",
        openInNewWindow: true,
      });

      // Close modal
      if (this.props.onClose) {
        this.props.onClose();
      }
    } catch (error: any) {
      this.xrm.Utility.closeProgressIndicator();
      console.error("Error creating work order:", error);
      throw error;
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

      // Get or create account with progress indicator
      this.xrm.Utility.showProgressIndicator(this.strings.CreatingAccount);
      const accountId = await this.searchOrCreateAccount();
      this.xrm.Utility.closeProgressIndicator();

      if (!accountId) {
        throw new Error("Failed to get account");
      }

      // Determine Campaign and Incident Type values to pre-fill the popup
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

      // Always show campaign/incident selection popup before creating work order
      this.pendingAccountId = accountId;
      this.setState((prev) => ({
        ...prev,
        showCampaignIncidentPopup: true,
        popupShowCampaign: true,
        popupShowIncidentType: true,
        loading: false,
        error: null,
        selectedCampaignId: campaignId,
        selectedIncidentTypeId: resolvedIncidentTypeId,
        selectedIncidentTypeName: resolvedIncidentTypeName,
        incidentTypeReadOnly: incidentTypeDerivedFromCampaign,
      }));
    } catch (error: any) {
      this.xrm.Utility.closeProgressIndicator();
      console.error("Error in handleStart:", error);
      this.setState({
        error: error.message || "Error starting inspection",
        loading: false,
      });
    }
  };

  private handleContinueWithSelections = async (): Promise<void> => {
    try {
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

      // Reuse cached account ID from handleStart instead of re-searching
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
      console.error(
        "Error creating work order (handleContinueWithSelections):",
        error,
      );
      this.setState({
        error: error.message || "Error creating work order",
        loading: false,
      });
    }
  };

  // =====================================================================
  // FIELD VISIBILITY
  // =====================================================================

  private shouldShowField = (
    field:
      | "qataryId"
      | "name"
      | "crNumber"
      | "id"
      | "carColor"
      | "vehicleBrand",
  ): boolean => {
    const { selectedInspectionType, isAnonymous } = this.state;
    if (!selectedInspectionType) return false;

    // If anonymous checked for types 3, 6, 7, hide all fields
    if (isAnonymous && [3, 6, 7].includes(selectedInspectionType)) {
      return false;
    }

    if (field === "id" || field === "carColor" || field === "vehicleBrand") {
      return selectedInspectionType === 1;
    }

    if (field === "qataryId" || field === "name") {
      return (
        selectedInspectionType === 2 ||
        selectedInspectionType === 3 ||
        selectedInspectionType === 6
      );
    }

    if (field === "crNumber") {
      return selectedInspectionType === 5 || selectedInspectionType === 7;
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
    const { loading, isRTL } = this.state;

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
      direction: isRTL ? "rtl" : "ltr",
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

    const fieldStyle: React.CSSProperties = { marginBottom: 16 };

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
      backgroundPosition: isRTL ? "8px center" : "calc(100% - 8px) center",
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
      color: FLUENT.colorNeutralSecondary,
      fontSize: 16,
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
    };
  };

  // =====================================================================
  // RENDER HELPERS
  // =====================================================================

  private renderBarcodeIcon = () =>
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
    );

  private renderSelectWithClear = (
    value: string,
    onChange: (val: string) => void,
    options: Array<{ key: string; value: string; label: string }>,
    placeholder: string,
    disabled: boolean,
    styles: ReturnType<typeof this.getStyles>,
  ) => {
    return React.createElement(
      "div",
      { style: styles.selectWrapperStyle },
      React.createElement(
        "select",
        {
          value: value || "",
          onChange: (e: React.ChangeEvent<HTMLSelectElement>) =>
            onChange(e.target.value),
          disabled: disabled,
          style: {
            ...styles.selectInnerStyle,
            backgroundColor: disabled
              ? FLUENT.colorNeutralLighter
              : FLUENT.colorWhite,
            cursor: disabled ? "not-allowed" : "pointer",
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

    // ---- Campaign / Incident Popup ----
    if (showCampaignIncidentPopup) {
      const showCampaignField = this.state.popupShowCampaign;
      const showIncidentField = this.state.popupShowIncidentType;

      let popupTitle = "";
      if (showCampaignField && showIncidentField) {
        popupTitle = `${this.strings.SelectCampaign} / ${this.strings.SelectIncidentType}`;
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
                this.strings.IncidentType,
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

    // ---- Main Form ----
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
            this.strings.InspectionType,
          ),
          this.renderSelectWithClear(
            selectedInspectionType !== null
              ? String(selectedInspectionType)
              : "",
            (val) => {
              const syntheticEvent = {
                target: { value: val },
              } as React.ChangeEvent<HTMLSelectElement>;
              this.handleInspectionTypeChange(syntheticEvent);
            },
            inspectionTypes.map((t) => ({
              key: String(t.value),
              value: String(t.value),
              label: t.label,
            })),
            this.strings.chooseInspectionType,
            isInspectionTypeDisabled,
            styles,
          ),
        ),

        // Anonymous Checkbox (types 3, 6, 7)
        this.shouldShowAnonymousCheckbox() &&
          React.createElement(
            "div",
            { style: styles.checkboxContainerStyle },
            React.createElement("input", {
              type: "checkbox",
              checked: isAnonymous,
              onChange: (e: React.ChangeEvent<HTMLInputElement>) =>
                this.setState({ isAnonymous: e.target.checked }),
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
                },
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
              this.strings.ID,
            ),
            React.createElement(
              "div",
              { style: { display: "flex", alignItems: "center", gap: 6 } },
              React.createElement("input", {
                type: "text",
                value: id,
                onChange: (e: React.ChangeEvent<HTMLInputElement>) =>
                  this.handleInputChange("id", e.target.value),
                disabled: loading,
                style: { ...styles.inputStyle, flex: 1 },
              }),
              React.createElement(
                "button",
                {
                  onClick: () => this.handleScanBarcode("id"),
                  disabled: loading,
                  style: styles.scanButtonStyle,
                  type: "button",
                  title: this.strings.ScanBarcode,
                },
                this.renderBarcodeIcon(),
              ),
            ),
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
              this.strings.QataryID,
            ),
            React.createElement(
              "div",
              { style: { display: "flex", alignItems: "center", gap: 6 } },
              React.createElement("input", {
                type: "text",
                value: qataryId,
                onChange: (e: React.ChangeEvent<HTMLInputElement>) =>
                  this.handleInputChange("qataryId", e.target.value),
                disabled: loading,
                style: { ...styles.inputStyle, flex: 1 },
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
                this.renderBarcodeIcon(),
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
              this.strings.Name,
            ),
            React.createElement("input", {
              type: "text",
              value: name,
              onChange: (e: React.ChangeEvent<HTMLInputElement>) =>
                this.handleInputChange("name", e.target.value),
              disabled: loading,
              style: styles.inputStyle,
            }),
          ),

        // CR Number / Monour Number with Barcode Scanner
        this.shouldShowField("crNumber") &&
          React.createElement(
            "div",
            { style: styles.fieldStyle },
            React.createElement(
              "label",
              { style: styles.labelStyle },
              selectedInspectionType === 7
                ? this.strings.MonourNumber
                : this.strings.CRNumber,
            ),
            React.createElement(
              "div",
              { style: { display: "flex", alignItems: "center", gap: 6 } },
              React.createElement("input", {
                type: "text",
                value: crNumber,
                onChange: (e: React.ChangeEvent<HTMLInputElement>) =>
                  this.handleInputChange("crNumber", e.target.value),
                disabled: loading,
                style: { ...styles.inputStyle, flex: 1 },
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
                this.renderBarcodeIcon(),
              ),
            ),
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
