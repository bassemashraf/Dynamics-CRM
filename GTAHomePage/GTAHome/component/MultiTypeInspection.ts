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
}

interface IMultiTypeInspectionState {
  isRTL: boolean;
  // GTA: simplified to two checkbox types
  isTaxable: boolean;
  isNonTaxable: boolean;
  crNumber: string;
  tempCrNumber: string;
  loading: boolean;
  error: string | null;
  // Campaign / Incident Type popup (screen 2)
  showCampaignIncidentPopup: boolean;
  selectedCampaignId?: string;
  selectedCampaignName?: string;
  selectedIncidentTypeId?: string;
  selectedIncidentTypeName?: string;
  campaigns: Array<{ id: string; name: string }>;
  incidentTypes: Array<{ id: string; name: string }>;
  popupShowCampaign: boolean;
  popupShowIncidentType: boolean;
  incidentTypeReadOnly: boolean;
  campaignIncidentTypeMap: Record<string, { id: string; name: string }>;
  taxTypeOptions: Array<{ value: number; label: string }>;
  selectedTaxTypeId?: number;
}

interface LocalizedStrings {
  StartMultiTypeInspection: string;
  Taxable: string;
  NonTaxable: string;
  CRNumber: string;
  TempCRNumber: string;
  Start: string;
  Close: string;
  Loading: string;
  PleaseEnterRequiredFields: string;
  Error: string;
  SelectCampaign: string;
  SelectIncidentType: string;
  Campaign: string;
  IncidentType: string;
  Continue: string;
  CreatingAccount: string;
  CreatingWorkOrder: string;
  CreatingBooking: string;
  ScanBarcode: string;
  Clear: string;
  TaxType: string;
}

// Cache constants
const CAMPAIGNS_CACHE_KEY = "GTA_Campaigns_Cache";
const INCIDENT_TYPES_CACHE_KEY = "GTA_IncidentTypes_Cache";
const CACHE_DURATION = 60_000; // 1 minute

interface CacheData<T> {
  data: T;
  timestamp: number;
}

// =====================================================================
// GTA STYLE TOKENS – uses GTA brand colors
// =====================================================================

const GTA_STYLE = {
  fontFamily:
    '"Segoe UI", "Segoe UI Web (West European)", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif',
  colorPrimary: "#113f61",
  colorPrimaryHover: "#0d3350",
  colorAccent: "#8A1538",
  colorAccentHover: "#701030",
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
  focusOutline: "2px solid #113f61",
  transitionFast: "0.1s ease",
} as const;

export class MultiTypeInspection extends React.Component<
  IMultiTypeInspectionProps,
  IMultiTypeInspectionState
> {
  private strings: LocalizedStrings;
  private xrm: any;
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
      StartMultiTypeInspection:
        props.context.resources.getString("StartMultiTypeInspection") ||
        "Start Inspection",
      Taxable: props.context.resources.getString("Taxable") || "Taxable",
      NonTaxable:
        props.context.resources.getString("NonTaxable") || "Non-Taxable",
      CRNumber:
        props.context.resources.getString("CRNumber") || "CR Number:",
      TempCRNumber:
        props.context.resources.getString("TempCRNumber") ||
        "Temp CR Number:",
      Start: props.context.resources.getString("Start") || "Start",
      Close: props.context.resources.getString("Close") || "Close",
      Loading:
        props.context.resources.getString("Loading") || "Loading...",
      PleaseEnterRequiredFields:
        props.context.resources.getString("PleaseEnterRequiredFields") ||
        "Please enter required fields",
      Error: props.context.resources.getString("Error") || "Error",
      SelectCampaign:
        props.context.resources.getString("SelectCampaign") ||
        "Select Campaign",
      SelectIncidentType:
        props.context.resources.getString("SelectIncidentType") ||
        "Select Incident Type",
      Campaign:
        props.context.resources.getString("Campaign") || "Campaign",
      IncidentType:
        props.context.resources.getString("IncidentType") || "Incident Type",
      Continue:
        props.context.resources.getString("Continue") || "Continue",
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
      TaxType: props.context.resources.getString("TaxType") || "Tax Type",
    };

    this.state = {
      isRTL: isRTL,
      isTaxable: false,
      isNonTaxable: false,
      crNumber: "",
      tempCrNumber: "",
      loading: false,
      error: null,
      showCampaignIncidentPopup: false,
      selectedCampaignId: props.activePatrolId,
      selectedCampaignName: props.activePatrolName,
      selectedIncidentTypeId: undefined,
      selectedIncidentTypeName: undefined,
      campaigns: [],
      incidentTypes: [],
      popupShowCampaign: false,
      popupShowIncidentType: false,
      incidentTypeReadOnly: false,
      campaignIncidentTypeMap: {},
      taxTypeOptions: [],
      selectedTaxTypeId: undefined,
    };
  }

  async componentDidMount(): Promise<void> {
    const userId =
      this.xrm.Utility.getGlobalContext().userSettings.userId.replace(
        /[{}]/g,
        "",
      );

    // Parallel loading
    await Promise.all([
      this.preloadCampaignsAndIncidentTypes(),
      this.loadTaxTypeOptions(),
      InitCache.load(userId),
    ]);
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
    } catch (error) {
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
    } catch (error) {
      console.error(`Error saving cache for ${key}:`, error);
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

      // Build campaign → incident type map
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
    } catch (error) {
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
      const query = `?$filter=_duc_organizationalunitid_value eq '${this.props.organizationUnitId}' and duc_campaigntype eq 100000000 and statecode eq 0 &$select=new_inspectioncampaignid,new_name&$orderby=new_name asc`;

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
    } catch (error) {
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
    } catch (error) {
      console.error("Error loading incident types:", error);
      return [];
    }
  };
    
  private loadTaxTypeOptions = async (): Promise<void> => {
    try {
      const entityMetadata = await this.xrm.Utility.getEntityMetadata(
        "msdyn_workorder",
        ["duc_taxtype"],
      );
      const attribute = (entityMetadata as any).Attributes.get("duc_taxtype");
      const options: Array<{ value: number; label: string }> = [];

      if (attribute?.OptionSet) {
        Object.entries(attribute.OptionSet).forEach(([key, opt]: [string, any]) => {
          options.push({
            value: opt.value,
            label: opt.text,
          });
        });
      }
      
      this.setState({ taxTypeOptions: options });
    } catch (error) {
      console.error("Error loading tax type options:", error);
    }
  };

  // =====================================================================
  // HANDLERS
  // =====================================================================

  private handleTaxableChange = (checked: boolean): void => {
    this.setState({
      isTaxable: checked,
      isNonTaxable: checked ? false : this.state.isNonTaxable,
      crNumber: "",
      tempCrNumber: "",
      error: null,
    });
  };

  private handleNonTaxableChange = (checked: boolean): void => {
    this.setState({
      isNonTaxable: checked,
      isTaxable: checked ? false : this.state.isTaxable,
      crNumber: "",
      tempCrNumber: "",
      error: null,
    });
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
    field: "crNumber" | "tempCrNumber",
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
    } catch (error) {
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

  private handleTaxTypeChange = (taxTypeId: string): void => {
    this.setState({
      selectedTaxTypeId: taxTypeId ? parseInt(taxTypeId) : undefined,
    });
  };

  // =====================================================================
  // VALIDATION
  // =====================================================================

  private validateFields = (): boolean => {
    const { isTaxable, isNonTaxable, crNumber } = this.state;

    if (!isTaxable && !isNonTaxable) {
      this.setState({
        error: this.strings.PleaseEnterRequiredFields,
      });
      return false;
    }

    if (!crNumber.trim()) {
      this.setState({ error: this.strings.PleaseEnterRequiredFields });
      return false;
    }

    return true;
  };

  // =====================================================================
  // ACCOUNT SEARCH/CREATE
  // =====================================================================

  private getAccountName = (): string => {
    const { crNumber } = this.state;
    return `Company ${crNumber}`.trim();
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
    } catch (error) {
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
    } catch (error) {
      console.error("Error creating address information:", error);
    }
  };

  private searchOrCreateAccount = async (): Promise<string | null> => {
    try {
      const { crNumber } = this.state;

      if (!crNumber.trim()) {
        throw new Error(this.strings.PleaseEnterRequiredFields);
      }

      // Search for existing account by CR number
      const filterQuery = `duc_accountidentifier eq '${crNumber.trim()}'`;

      const searchResults = await this.xrm.WebApi.retrieveMultipleRecords(
        "account",
        `?$select=accountid,name&$filter=${filterQuery}`,
      );

      if (searchResults?.entities?.length > 0) {
        // Existing account found — return its ID
        return searchResults.entities[0].accountid;
      }

      // Create new account (type = Company, customertypecode = 2)
      const accountName = this.getAccountName();
      const newAccount: any = {
        name: accountName,
        duc_accountidentifier: crNumber.trim(),
        customertypecode: 2, // Company
      };

      const createdAccount = await this.xrm.WebApi.createRecord(
        "account",
        newAccount,
      );
      const newAccountId = createdAccount?.id;

      // Create address info for new accounts
      await this.createAddressInformation(newAccountId, accountName);

      return newAccountId;
    } catch (error: any) {
      console.error("Error searching/creating account:", error);
      throw error;
    }
  };

  // =====================================================================
  // WORK ORDER CREATION
  // =====================================================================

  private createWorkOrder = async (
    accountId: string,
    taxType?: number,
  ): Promise<void> => {
    try {
      const userId =
        this.xrm.Utility.getGlobalContext().userSettings.userId.replace(
          /[{}]/g,
          "",
        );

      // Get account data
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
        const mapped =
          this.state.campaignIncidentTypeMap[this.state.selectedCampaignId];
        if (mapped) {
          incidentTypeData = {
            id: mapped.id,
            name: mapped.name,
            entityType: "msdyn_incidenttype",
          };
        } else {
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

      // STEP 6: Detect if created from mobile
      const createdFromMobile = WorkOrderHelpers.isMobileClient();

      // STEP 7: Validate data
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

      // STEP 8: Create work order
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
        anonymousCustomer: false,
        createdFromMobile: createdFromMobile,
        taxType: taxType,
      });
      this.xrm.Utility.closeProgressIndicator();

      if (!workOrderId) {
        throw new Error("Failed to create work order");
      }

      console.log("Work order created successfully:", workOrderId);

      // STEP 9: Create auto booking if from mobile
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

      // STEP 10: Navigate to the created work order
      await this.xrm.Navigation.openForm({
        entityName: "msdyn_workorder",
        entityId: workOrderId,
        formId: "eded7d77-6dc4-ed11-b596-6045bdf00fa1",
        openInNewWindow: true,
      });

      // Close the modal
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
  // MAIN START HANDLER (Screen 1 → Screen 2)
  // =====================================================================

  private handleStart = async (): Promise<void> => {
    if (!this.validateFields()) {
      return;
    }

    try {
      this.setState({ loading: true, error: null });

      // Get or create account
      this.xrm.Utility.showProgressIndicator(this.strings.CreatingAccount);
      const accountId = await this.searchOrCreateAccount();
      this.xrm.Utility.closeProgressIndicator();

      if (!accountId) {
        throw new Error("Failed to get account");
      }

      // Resolve incident type
      let resolvedIncidentTypeId =
        this.state.selectedIncidentTypeId || this.props.incidentTypeId;
      let resolvedIncidentTypeName =
        this.state.selectedIncidentTypeName || this.props.incidentTypeName;

      const campaignId =
        this.state.selectedCampaignId || this.props.activePatrolId;
      let incidentTypeDerivedFromCampaign = false;

      if (!resolvedIncidentTypeId && campaignId) {
        const mapped = this.state.campaignIncidentTypeMap[campaignId];
        if (mapped) {
          resolvedIncidentTypeId = mapped.id;
          resolvedIncidentTypeName = mapped.name;
          incidentTypeDerivedFromCampaign = true;
        }
      }

      // Show campaign & incident type popup
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

      this.setState({ loading: true, error: null });

      const accountId =
        this.pendingAccountId || (await this.searchOrCreateAccount());
      this.pendingAccountId = null;
      if (!accountId) {
        throw new Error("Failed to get account");
      }

      await this.createWorkOrder(accountId, this.state.selectedTaxTypeId);
      this.setState({ loading: false });
    } catch (error: any) {
      this.xrm.Utility.closeProgressIndicator();
      console.error("Error creating work order:", error);
      this.setState({
        error: error.message || "Error creating work order",
        loading: false,
      });
    }
  };

  // =====================================================================
  // STYLE HELPERS — GTA branded
  // =====================================================================

  private getStyles = () => {
    const { loading } = this.state;

    const containerStyle: React.CSSProperties = {
      position: "fixed",
      top: 0,
      left: 0,
      width: "100%",
      height: "100%",
      backgroundColor: GTA_STYLE.overlay,
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      zIndex: 1000,
      direction: this.state.isRTL ? "rtl" : "ltr",
    };

    const modalStyle: React.CSSProperties = {
      backgroundColor: GTA_STYLE.colorWhite,
      borderRadius: GTA_STYLE.borderRadiusModal,
      padding: 24,
      maxWidth: 500,
      width: "90%",
      maxHeight: "90vh",
      overflowY: "auto",
      boxShadow: GTA_STYLE.shadowModal,
    };

    const titleStyle: React.CSSProperties = {
      fontSize: 18,
      fontWeight: 600,
      marginBottom: 20,
      color: GTA_STYLE.colorPrimary,
      fontFamily: GTA_STYLE.fontFamily,
    };

    const fieldStyle: React.CSSProperties = {
      marginBottom: 16,
    };

    const labelStyle: React.CSSProperties = {
      display: "block",
      fontSize: 14,
      fontWeight: 600,
      marginBottom: 4,
      color: GTA_STYLE.colorNeutralPrimary,
      fontFamily: GTA_STYLE.fontFamily,
    };

    const inputStyle: React.CSSProperties = {
      width: "100%",
      padding: "6px 8px",
      border: `1px solid ${GTA_STYLE.colorNeutralSecondary}`,
      borderRadius: GTA_STYLE.borderRadius,
      fontSize: 14,
      fontFamily: GTA_STYLE.fontFamily,
      boxSizing: "border-box",
      outline: "none",
      transition: `border-color ${GTA_STYLE.transitionFast}`,
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
      borderRadius: GTA_STYLE.borderRadius,
      backgroundColor: GTA_STYLE.colorAccent,
      color: GTA_STYLE.colorWhite,
      fontSize: 12,
      fontWeight: 600,
      fontFamily: GTA_STYLE.fontFamily,
      cursor: "pointer",
      transition: `background-color ${GTA_STYLE.transitionFast}`,
      whiteSpace: "nowrap",
      flexShrink: 0,
    };

    const scanButtonStyle: React.CSSProperties = {
      padding: "6px",
      border: "none",
      borderRadius: GTA_STYLE.borderRadius,
      backgroundColor: "transparent",
      color: GTA_STYLE.colorNeutralSecondary,
      fontSize: 16,
      fontWeight: 600,
      fontFamily: GTA_STYLE.fontFamily,
      cursor: loading ? "not-allowed" : "pointer",
      transition: `background-color ${GTA_STYLE.transitionFast}, color ${GTA_STYLE.transitionFast}`,
      whiteSpace: "nowrap",
      flexShrink: 0,
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
    };

    const checkboxContainerStyle: React.CSSProperties = {
      display: "flex",
      alignItems: "center",
      gap: 8,
      marginBottom: 16,
      padding: "10px 12px",
      border: `1px solid ${GTA_STYLE.colorNeutralLight}`,
      borderRadius: GTA_STYLE.borderRadius,
      cursor: loading ? "not-allowed" : "pointer",
      transition: `background-color ${GTA_STYLE.transitionFast}`,
    };

    const checkboxStyle: React.CSSProperties = {
      width: 18,
      height: 18,
      accentColor: GTA_STYLE.colorPrimary,
      cursor: loading ? "not-allowed" : "pointer",
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
      borderRadius: GTA_STYLE.borderRadius,
      fontSize: 14,
      fontWeight: 600,
      cursor: loading ? "not-allowed" : "pointer",
      fontFamily: GTA_STYLE.fontFamily,
      transition: `background-color ${GTA_STYLE.transitionFast}`,
      minWidth: 80,
    };

    const startButtonStyle: React.CSSProperties = {
      ...buttonStyle,
      backgroundColor: loading ? "#d3d3d3" : GTA_STYLE.colorPrimary,
      color: GTA_STYLE.colorWhite,
    };

    const closeButtonStyle: React.CSSProperties = {
      ...buttonStyle,
      backgroundColor: GTA_STYLE.colorNeutralLighter,
      color: GTA_STYLE.colorNeutralDark,
      border: `1px solid ${GTA_STYLE.colorNeutralLight}`,
    };

    const errorStyle: React.CSSProperties = {
      backgroundColor: GTA_STYLE.colorErrorBackground,
      color: GTA_STYLE.colorErrorPrimary,
      padding: "8px 12px",
      borderRadius: GTA_STYLE.borderRadius,
      fontSize: 13,
      fontFamily: GTA_STYLE.fontFamily,
      marginBottom: 16,
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
      checkboxContainerStyle,
      checkboxStyle,
      buttonContainerStyle,
      startButtonStyle,
      closeButtonStyle,
      errorStyle,
    };
  };

  // =====================================================================
  // RENDER HELPERS
  // =====================================================================

  private renderSelectWithClear = (
    value: string,
    onChange: (val: string) => void,
    options: Array<{ key: string; value: string; label: string }>,
    placeholder: string,
    disabled: boolean,
    styles: ReturnType<typeof this.getStyles>,
  ): React.ReactElement => {
    const hasValue = value !== "";

    return React.createElement(
      "div",
      { style: styles.selectWrapperStyle },
      React.createElement(
        "select",
        {
          value: value,
          onChange: (e: React.ChangeEvent<HTMLSelectElement>) =>
            onChange(e.target.value),
          disabled: disabled,
          style: styles.selectInnerStyle,
        },
        React.createElement("option", { value: "" }, placeholder),
        ...options.map((opt) =>
          React.createElement(
            "option",
            { key: opt.key, value: opt.value },
            opt.label,
          ),
        ),
      ),
      hasValue &&
        React.createElement(
          "button",
          {
            type: "button",
            onClick: () => onChange(""),
            disabled: disabled,
            style: styles.clearButtonStyle,
          },
          this.strings.Clear,
        ),
    );
  };

  private renderBarcodeButton = (
    field: "crNumber" | "tempCrNumber",
    styles: ReturnType<typeof this.getStyles>,
  ): React.ReactElement => {
    return React.createElement(
      "button",
      {
        onClick: () => this.handleScanBarcode(field),
        disabled: this.state.loading,
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
    );
  };

  // =====================================================================
  // RENDER
  // =====================================================================

  render(): React.ReactElement | null {
    if (!this.props.isOpen) return null;

    const {
      isTaxable,
      isNonTaxable,
      crNumber,
      tempCrNumber,
      loading,
      error,
      showCampaignIncidentPopup,
      campaigns,
      incidentTypes,
      selectedCampaignId,
      selectedIncidentTypeId,
      incidentTypeReadOnly,
      popupShowCampaign,
      popupShowIncidentType,
    } = this.state;

    const styles = this.getStyles();

    const showCampaignField = popupShowCampaign;
    const showIncidentField = popupShowIncidentType;

    // ------
    // Screen 2 — Campaign & Incident Type Selection
    // ------
    if (showCampaignIncidentPopup) {
      return React.createElement(
        "div",
        { style: styles.containerStyle },
        React.createElement(
          "div",
          { style: styles.modalStyle },
          React.createElement(
            "h2",
            { style: styles.titleStyle },
            this.strings.SelectCampaign,
          ),

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

          // Tax Type Selection
          React.createElement(
            "div",
            { style: styles.fieldStyle },
            React.createElement(
              "label",
              { style: styles.labelStyle },
              this.strings.TaxType,
            ),
            this.renderSelectWithClear(
              this.state.selectedTaxTypeId?.toString() || "",
              (val) => this.handleTaxTypeChange(val),
              this.state.taxTypeOptions.map((opt) => ({
                key: opt.value.toString(),
                value: opt.value.toString(),
                label: opt.label,
              })),
              "--",
              loading,
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

    // ------
    // Screen 1 — Tax Registration Type & CR Number
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

        // Taxable Checkbox
        React.createElement(
          "div",
          {
            style: {
              ...styles.checkboxContainerStyle,
              backgroundColor: isTaxable
                ? `${GTA_STYLE.colorPrimary}0D`
                : "transparent",
              borderColor: isTaxable
                ? GTA_STYLE.colorPrimary
                : GTA_STYLE.colorNeutralLight,
            },
            onClick: () => !loading && this.handleTaxableChange(!isTaxable),
          },
          React.createElement("input", {
            type: "checkbox",
            checked: isTaxable,
            onChange: (e: React.ChangeEvent<HTMLInputElement>) =>
              this.handleTaxableChange(e.target.checked),
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
                color: isTaxable
                  ? GTA_STYLE.colorPrimary
                  : GTA_STYLE.colorNeutralPrimary,
                fontWeight: isTaxable ? 700 : 600,
              } as React.CSSProperties,
            },
            this.strings.Taxable,
          ),
        ),

        // Non-Taxable Checkbox
        React.createElement(
          "div",
          {
            style: {
              ...styles.checkboxContainerStyle,
              backgroundColor: isNonTaxable
                ? `${GTA_STYLE.colorAccent}0D`
                : "transparent",
              borderColor: isNonTaxable
                ? GTA_STYLE.colorAccent
                : GTA_STYLE.colorNeutralLight,
            },
            onClick: () =>
              !loading && this.handleNonTaxableChange(!isNonTaxable),
          },
          React.createElement("input", {
            type: "checkbox",
            checked: isNonTaxable,
            onChange: (e: React.ChangeEvent<HTMLInputElement>) =>
              this.handleNonTaxableChange(e.target.checked),
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
                color: isNonTaxable
                  ? GTA_STYLE.colorAccent
                  : GTA_STYLE.colorNeutralPrimary,
                fontWeight: isNonTaxable ? 700 : 600,
              } as React.CSSProperties,
            },
            this.strings.NonTaxable,
          ),
        ),

        // CR Number (shown for both Taxable and Non-Taxable)
        (isTaxable || isNonTaxable) &&
          React.createElement(
            "div",
            { style: styles.fieldStyle },
            React.createElement(
              "label",
              { style: styles.labelStyle },
              this.strings.CRNumber,
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
              this.renderBarcodeButton("crNumber", styles),
            ),
          ),

        // Temp CR Number (shown only for Taxable)
        isTaxable &&
          React.createElement(
            "div",
            { style: styles.fieldStyle },
            React.createElement(
              "label",
              { style: styles.labelStyle },
              this.strings.TempCRNumber,
            ),
            React.createElement(
              "div",
              { style: { display: "flex", alignItems: "center", gap: 6 } },
              React.createElement("input", {
                type: "text",
                value: tempCrNumber,
                onChange: (e: React.ChangeEvent<HTMLInputElement>) =>
                  this.handleInputChange("tempCrNumber", e.target.value),
                disabled: loading,
                style: { ...styles.inputStyle, flex: 1 },
              }),
              this.renderBarcodeButton("tempCrNumber", styles),
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
