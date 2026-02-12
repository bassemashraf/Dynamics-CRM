/* eslint-disable */
    import * as React from 'react';
    import { WorkOrderHelpers, CampaignHelpers, IncidentTypeHelpers } from '../helpers';

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
        inspectionTypes: Array<{ value: number; label: string; accountTypeId: string }>;
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
        Error: string;
        chooseInspectionType: string;
        SelectCampaign: string;
        SelectIncidentType: string;
        Campaign: string;
        IncidentType: string;
        Continue: string;
        Anonymous: string;
    }

    // Cache constants
    const INSPECTION_TYPES_CACHE_KEY = 'MOCI_OrgUnit_InspectionTypes_Cache';
    const VEHICLE_TYPES_CACHE_KEY = 'MOCI_VehicleTypes_Cache';
    const CAMPAIGNS_CACHE_KEY = 'MOCI_Campaigns_Cache';
    const INCIDENT_TYPES_CACHE_KEY = 'MOCI_IncidentTypes_Cache';
    const CACHE_DURATION = 60_000; // 1 minute

    interface CacheData<T> {
        data: T;
        timestamp: number;
    }

    export class MultiTypeInspection extends React.Component<IMultiTypeInspectionProps, IMultiTypeInspectionState> {
        private strings: LocalizedStrings;
        private xrm: Xrm.XrmStatic;
        // OPTIMIZATION: Cache account ID to avoid re-searching in handleContinueWithSelections
        private pendingAccountId: string | null = null;

        constructor(props: IMultiTypeInspectionProps) {
            super(props);

            const userSettings = (props.context as any).userSettings;
            const rtlLanguages = [1025, 1037, 1054, 1056, 1065, 1068, 1069, 1101, 1114, 1119];
            const isRTL = rtlLanguages.includes(userSettings?.languageId);

            this.xrm = (window.parent as any).Xrm || (window as any).Xrm;

            // Load localized strings
            this.strings = {
                StartMultiTypeInspection: props.context.resources.getString("StartMultiTypeInspection"),
                InspectionType: props.context.resources.getString("InspectionType"),
                QataryID: props.context.resources.getString("QataryID"),
                Name: props.context.resources.getString("Name"),
                CRNumber: props.context.resources.getString("CRNumber"),
                MonourNumber: props.context.resources.getString("MonourNumber") || "Monour Number",
                ID: props.context.resources.getString("ID"),
                CarColor: props.context.resources.getString("CarColor") || "Car Color",
                VehicleBrand: props.context.resources.getString("VehicleBrand") || "Vehicle Brand",
                Start: props.context.resources.getString("Start"),
                Close: props.context.resources.getString("Close"),
                Loading: props.context.resources.getString("Loading"),
                PleaseSelectInspectionType: props.context.resources.getString("PleaseSelectInspectionType"),
                PleaseEnterRequiredFields: props.context.resources.getString("PleaseEnterRequiredFields"),
                Error: props.context.resources.getString("Error"),
                chooseInspectionType: props.context.resources.getString("chooseInspectionType"),
                SelectCampaign: props.context.resources.getString("SelectCampaign") || "Select Campaign",
                SelectIncidentType: props.context.resources.getString("SelectIncidentType") || "Select Incident Type",
                Campaign: props.context.resources.getString("Campaign") || "Campaign",
                IncidentType: props.context.resources.getString("IncidentType") || "Incident Type",
                Continue: props.context.resources.getString("Continue") || "Continue",
                Anonymous: props.context.resources.getString("Anonymous") || "Anonymous",
            };

            this.state = {
                isRTL: isRTL,
                inspectionTypes: [],
                selectedInspectionType: props.defaultInspectionType || null,
                qataryId: '',
                name: '',
                crNumber: '',
                id: '',
                carColor: '',
                vehicleBrand: null,
                vehicleBrands: [],
                loading: false,
                error: null,
                accountTypeRecord: null,
                showCampaignIncidentPopup: false,
                selectedCampaignId: props.activePatrolId,
                selectedCampaignName: props.activePatrolName,
                selectedIncidentTypeId: props.incidentTypeId,
                selectedIncidentTypeName: props.incidentTypeName,
                campaigns: [],
                incidentTypes: [],
                isAnonymous: false,
                // Track which fields to show in popup (set when popup opens)
                popupShowCampaign: false,
                popupShowIncidentType: false,
            };
        }

        async componentDidMount(): Promise<void> {
            // Parallel loading for better performance
            await Promise.all([
                this.loadInspectionTypesFromOrgUnit(),
                this.loadVehicleTypes(),
                this.preloadCampaignsAndIncidentTypes()
            ]);

            // If default inspection type is provided, set account type from loaded data
            if (this.props.defaultInspectionType) {
                const inspectionType = this.state.inspectionTypes.find(t => t.value === this.props.defaultInspectionType);
                if (inspectionType?.accountTypeId) {
                    const accountTypeRecord = {
                        duc_accounttypeid: inspectionType.accountTypeId,
                        duc_accounttype: this.props.defaultInspectionType,
                        duc_name: inspectionType.label
                    };
                    this.setState({ accountTypeRecord });
                }
            }
        }

        // =====================================================================
        // GENERIC CACHE UTILITIES
        // =====================================================================

        private getFromCache = <T,>(key: string): T | null => {
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

        private saveToCache = <T,>(key: string, data: T): void => {
            try {
                const cacheData: CacheData<T> = {
                    data,
                    timestamp: Date.now()
                };
                localStorage.setItem(key, JSON.stringify(cacheData));
            } catch (error) {
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
                const cachedTypes = this.getFromCache<Array<{ value: number; label: string; accountTypeId: string }>>(cacheKey);

                if (cachedTypes) {
                    this.setState({ inspectionTypes: cachedTypes });
                    return;
                }

                // Fetch from junction entity
                const query = `?$filter=_duc_organizationunit_value eq '${this.props.organizationUnitId}'` +
                    `&$select=duc_organizationunitaccounttypesid` +
                    `&$expand=duc_AccountType($select=duc_accounttypeid,duc_name,duc_accounttype)`;

                const results = await this.xrm.WebApi.retrieveMultipleRecords(
                    'duc_organizationunitaccounttypes',
                    query
                );

                if (!results?.entities || results.entities.length === 0) {
                    console.warn("No account types found for organization unit");
                    return;
                }

                // Get option set metadata for labels
                const entityMetadata = await this.xrm.Utility.getEntityMetadata('account', ['duc_accountinspectiontype']);
                const attribute = (entityMetadata as any).Attributes.get('duc_accountinspectiontype');
                const optionSetMap = new Map<number, string>();

                if (attribute?.OptionSet) {
                    Object.values(attribute.OptionSet).forEach((opt: any) => {
                        optionSetMap.set(opt.value, opt.text);
                    });
                }

                // Map results to inspection types
                const types: Array<{ value: number; label: string; accountTypeId: string }> = [];
                
                for (const entity of results.entities) {
                    if (entity.duc_AccountType?.duc_accounttype !== undefined) {
                        const optionValue = entity.duc_AccountType.duc_accounttype;
                        const label = optionSetMap.get(optionValue) || `Type ${optionValue}`;
                        
                        types.push({
                            value: optionValue,
                            label: label,
                            accountTypeId: entity.duc_AccountType.duc_accounttypeid
                        });
                    }
                }

                // Sort by value
                types.sort((a, b) => a.value - b.value);

                this.saveToCache(cacheKey, types);
                this.setState({ inspectionTypes: types });

            } catch (error) {
                console.error('Error loading inspection types from org unit:', error);
                this.setState({ error: 'Failed to load inspection types' });
            }
        };

        // =====================================================================
        // VEHICLE TYPES
        // =====================================================================

        private loadVehicleTypes = async (): Promise<void> => {
            try {
                const cachedTypes = this.getFromCache<Array<{ value: number; label: string }>>(VEHICLE_TYPES_CACHE_KEY);
                if (cachedTypes) {
                    this.setState({ vehicleBrands: cachedTypes });
                    return;
                }

                const entityMetadata = await this.xrm.Utility.getEntityMetadata('account', ['duc_vehicletype']);
                const attribute = (entityMetadata as any).Attributes.get('duc_vehicletype');

                if (attribute?.OptionSet) {
                    const types = Object.values(attribute.OptionSet).map((opt: any) => ({
                        value: opt.value,
                        label: opt.text,
                    }));

                    this.saveToCache(VEHICLE_TYPES_CACHE_KEY, types);
                    this.setState({ vehicleBrands: types });
                }
            } catch (error) {
                console.error('Error loading vehicle types:', error);
            }
        };

        // =====================================================================
        // CAMPAIGNS AND INCIDENT TYPES (PRELOAD)
        // =====================================================================

        private preloadCampaignsAndIncidentTypes = async (): Promise<void> => {
            // ALWAYS preload both campaigns and incident types
            // This ensures they're ready when the popup is shown
            try {
                const [campaigns, incidentTypes] = await Promise.all([
                    this.loadCampaigns(),
                    this.loadIncidentTypes()
                ]);

                this.setState({ campaigns, incidentTypes });
            } catch (error) {
                console.error('Error preloading campaigns/incident types:', error);
            }
        };

        private loadCampaigns = async (): Promise<Array<{ id: string; name: string }>> => {
            if (!this.props.organizationUnitId) return [];

            const cacheKey = `${CAMPAIGNS_CACHE_KEY}_${this.props.organizationUnitId}`;
            const cached = this.getFromCache<Array<{ id: string; name: string }>>(cacheKey);
            if (cached) return cached;

            try {
                // Get inspection campaigns with duc_campaigntype = 100000000 (Inspection Campaign)
                const query = `?$filter=_duc_organizationalunitid_value eq '${this.props.organizationUnitId}' and duc_campaigntype eq 100000000 and statecode eq 0 &$select=new_inspectioncampaignid,new_name&$orderby=new_name asc`;
                
                const results = await this.xrm.WebApi.retrieveMultipleRecords(
                    'new_inspectioncampaign',
                    query
                );

                const campaigns = results.entities.map((entity: any) => ({
                    id: entity.new_inspectioncampaignid,
                    name: entity.new_name
                }));

                this.saveToCache(cacheKey, campaigns);
                return campaigns;
            } catch (error) {
                console.error('Error loading campaigns:', error);
                return [];
            }
        };

        private loadIncidentTypes = async (): Promise<Array<{ id: string; name: string }>> => {
            if (!this.props.organizationUnitId) return [];

            const cacheKey = `${INCIDENT_TYPES_CACHE_KEY}_${this.props.organizationUnitId}`;
            const cached = this.getFromCache<Array<{ id: string; name: string }>>(cacheKey);
            if (cached) return cached;

            try {
                // Get all incident types related to organization unit
                const query = `?$filter=_duc_organizationalunitid_value eq '${this.props.organizationUnitId}'&$select=msdyn_incidenttypeid,msdyn_name&$orderby=msdyn_name asc`;
                
                const results = await this.xrm.WebApi.retrieveMultipleRecords(
                    'msdyn_incidenttype',
                    query
                );

                const incidentTypes = results.entities.map((entity: any) => ({
                    id: entity.msdyn_incidenttypeid,
                    name: entity.msdyn_name
                }));

                this.saveToCache(cacheKey, incidentTypes);
                return incidentTypes;
            } catch (error) {
                console.error('Error loading incident types:', error);
                return [];
            }
        };

        // =====================================================================
        // HANDLERS
        // =====================================================================

        // OPTIMIZED: No API call needed - use already loaded data!
        private handleInspectionTypeChange = (e: React.ChangeEvent<HTMLSelectElement>): void => {
            const value = e.target.value ? parseInt(e.target.value, 10) : null;

            let accountTypeRecord: any = null;

            // Get account type from already loaded inspection types (instant, no waiting!)
            if (value !== null) {
                const inspectionType = this.state.inspectionTypes.find(t => t.value === value);
                if (inspectionType?.accountTypeId) {
                    accountTypeRecord = {
                        duc_accounttypeid: inspectionType.accountTypeId,
                        duc_accounttype: value,
                        duc_name: inspectionType.label
                    };
                }
            }

            this.setState({
                selectedInspectionType: value,
                qataryId: '',
                name: '',
                crNumber: '',
                id: '',
                carColor: '',
                vehicleBrand: null,
                error: null,
                accountTypeRecord: accountTypeRecord,
                isAnonymous: false,
            });
        };

        private handleInputChange = (field: keyof IMultiTypeInspectionState, value: any): void => {
            this.setState({ [field]: value } as any);
        };

        // =====================================================================
        // VALIDATION
        // =====================================================================

        private getRequiredFields = (): Array<keyof IMultiTypeInspectionState> => {
            const { selectedInspectionType } = this.state;
            const requiredFields: Array<keyof IMultiTypeInspectionState> = [];

            if (!selectedInspectionType) return [];

            // Type 1 (Vehicle) => ID, car color, and vehicle brand
            if (selectedInspectionType === 1) {
                requiredFields.push('id', 'carColor', 'vehicleBrand');
            }
            // Type 2 (Individual), 3 (Cabin), or 6 (Wilderness camps) => Qatary ID and Name
            else if ([2, 3, 6].includes(selectedInspectionType)) {
                requiredFields.push('qataryId', 'name');
            }
            // Type 5 (Company) or 7 (Manor) => CR Number
            else if ([5, 7].includes(selectedInspectionType)) {
                requiredFields.push('crNumber');
            }

            return requiredFields;
        };

        private validateFields = (): boolean => {
            const { selectedInspectionType, isAnonymous } = this.state;

            if (!selectedInspectionType) {
                this.setState({ error: this.strings.PleaseSelectInspectionType });
                return false;
            }

            // Type 4 (Anonymous) doesn't need field validation
            if (selectedInspectionType === 4) {
                return true;
            }

            // If anonymous checkbox is checked for types 3, 6, or 7, skip field validation
            if (isAnonymous && [3, 6, 7].includes(selectedInspectionType)) {
                return true;
            }

            const requiredFields = this.getRequiredFields();
            for (const field of requiredFields) {
                const val: any = this.state[field];
                if (val === null || val === undefined || val === '') {
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
            const { selectedInspectionType, qataryId, crNumber, id } = this.state;

            if (selectedInspectionType === 1) return id;
            if ([2, 3, 6].includes(selectedInspectionType!)) return qataryId;
            if ([5, 7].includes(selectedInspectionType!)) return crNumber;
            
            return '';
        };

        private getAccountName = (): string => {
            const { selectedInspectionType, name, qataryId, crNumber, id, carColor, vehicleBrand, vehicleBrands } = this.state;

            if (selectedInspectionType === 1) {
                const brandLabel = vehicleBrand !== null
                    ? (vehicleBrands.find(v => v.value === vehicleBrand)?.label || vehicleBrand)
                    : '';
                return `Vehicle ${id} ${carColor} ${brandLabel}`.trim();
            }
            if (selectedInspectionType === 2) return `Individual ${qataryId} ${name}`.trim();
            if (selectedInspectionType === 3) return name || `Cabin ${qataryId}`;
            if (selectedInspectionType === 6) return name || `Wilderness Camp ${qataryId}`;
            if (selectedInspectionType === 5) return `Company ${crNumber}`.trim();
            if (selectedInspectionType === 7) return `Manor ${crNumber}`.trim();
            if (selectedInspectionType === 4) return 'Anonymous Account';

            return 'Account';
        };

        private getCurrentLocation = async (): Promise<{ latitude: number; longitude: number } | null> => {
            try {
                if (this.xrm?.Device?.getCurrentPosition) {
                    const location: any = await this.xrm.Device.getCurrentPosition();
                    return {
                        latitude: location.coords.latitude,
                        longitude: location.coords.longitude
                    };
                }
            } catch (error) {
                console.error('Error getting location:', error);
            }
            return null;
        };

        private createAddressInformation = async (accountId: string, accountName: string): Promise<void> => {
            try {
                const location = await this.getCurrentLocation();
                if (!location) return;

                const today = new Date().toISOString().split('T')[0];
                const addressName = `${accountName} ${today}`;

                const addressData: any = {
                    duc_name: addressName,
                    duc_latitude: location.latitude,
                    duc_longitude: location.longitude,
                    'duc_Account@odata.bind': `/accounts(${accountId})`
                };

                await this.xrm.WebApi.createRecord('duc_addressinformation', addressData);
            } catch (error) {
                console.error('Error creating address information:', error);
            }
        };

        private searchOrCreateAccount = async (): Promise<string | null> => {
            try {
                const { accountTypeRecord, selectedInspectionType, carColor, vehicleBrand, isAnonymous } = this.state;
                const identifierValue = this.getIdentifierValue();

                // Anonymous type (4) uses unknown account
                if (selectedInspectionType === 4) {
                    if (this.props.unknownAccountId) {
                        await this.createAddressInformation(
                            this.props.unknownAccountId,
                            this.props.unknownAccountName || 'Unknown Account'
                        );
                        return this.props.unknownAccountId;
                    } else {
                        throw new Error('No anonymous account available');
                    }
                }

                // If anonymous checkbox is checked for types 3, 6, or 7, use unknown account
                if (isAnonymous && [3, 6, 7].includes(selectedInspectionType!)) {
                    if (this.props.unknownAccountId) {
                        await this.createAddressInformation(
                            this.props.unknownAccountId,
                            this.props.unknownAccountName || 'Unknown Account'
                        );
                        return this.props.unknownAccountId;
                    } else {
                        throw new Error('No anonymous account available');
                    }
                }

                if (!identifierValue) {
                    throw new Error(this.strings.PleaseEnterRequiredFields);
                }

                // Search for existing account
                let filterQuery = `duc_accountidentifier eq '${identifierValue}'`;
                if (accountTypeRecord?.duc_accounttypeid) {
                    filterQuery += ` and _duc_newaccounttype_value eq ${accountTypeRecord.duc_accounttypeid}`;
                }

                const searchResults = await this.xrm.WebApi.retrieveMultipleRecords(
                    'account',
                    `?$select=accountid,name,duc_vehicletype&$filter=${filterQuery}`
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
                            await this.xrm.WebApi.updateRecord('account', accountId, updateData);
                        }
                    }

                    await this.createAddressInformation(accountId, accountName);
                    return accountId;
                }

                // Create new account
                const accountName = this.getAccountName();
                const newAccount: any = {
                    name: accountName,
                    duc_accountidentifier: identifierValue,
                    duc_accountinspectiontype: selectedInspectionType
                };

                if (accountTypeRecord?.duc_accounttypeid) {
                    newAccount['duc_NewAccountType@odata.bind'] = `/duc_accounttypes(${accountTypeRecord.duc_accounttypeid})`;
                }

                if (selectedInspectionType === 1) {
                    if (carColor) newAccount.duc_vehiclecolor = carColor;
                    if (vehicleBrand !== null) newAccount.duc_vehicletype = vehicleBrand;
                }

                const createdAccount = await this.xrm.WebApi.createRecord('account', newAccount);
                const newAccountId = createdAccount?.id;

                await this.createAddressInformation(newAccountId, accountName);

                return newAccountId;
            } catch (error: any) {
                console.error('Error searching/creating account:', error);
                throw error;
            }
        };

        // =====================================================================
        // WORK ORDER CREATION WITH FULL BUSINESS LOGIC
        // =====================================================================

        private createWorkOrder = async (accountId: string): Promise<void> => {
            try {
                const userId = this.xrm.Utility.getGlobalContext().userSettings.userId.replace(/[{}]/g, "");

                // Get account name and data
                const accountRecord = await this.xrm.WebApi.retrieveRecord('account', accountId, '?$select=name');
                const accountName = accountRecord?.name || '';

                // Prepare base data
                let serviceAccountData = {
                    id: accountId,
                    name: accountName,
                    entityType: 'account'
                };

                let addressData: any = null;
                let latitude: number | undefined = undefined;
                let longitude: number | undefined = undefined;

                // STEP 1: Handle sub-account change logic (onSubaccountChange)
                // Get service account, address, and coordinates
                const subAccountResult = await WorkOrderHelpers.handleSubAccountChange(accountId);
                if (subAccountResult) {
                    serviceAccountData = subAccountResult.serviceAccount;
                    addressData = subAccountResult.address;
                    latitude = subAccountResult.latitude;
                    longitude = subAccountResult.longitude;
                }

                // STEP 2: Determine incident type
                let incidentTypeData: { id: string; name: string; entityType: string } | undefined;
                
                if (this.state.selectedIncidentTypeId && this.state.selectedIncidentTypeName) {
                    // User selected from popup
                    incidentTypeData = {
                        id: this.state.selectedIncidentTypeId,
                        name: this.state.selectedIncidentTypeName,
                        entityType: 'msdyn_incidenttype'
                    };
                } else if (this.state.selectedCampaignId) {
                    // Get incident type from campaign (SetParentCampaign logic)
                    const campaignData = await WorkOrderHelpers.getCampaignData(this.state.selectedCampaignId);
                    if (campaignData?.incidentType) {
                        incidentTypeData = campaignData.incidentType;
                    }
                } else if (this.props.incidentTypeId && this.props.incidentTypeName) {
                    // Use from props
                    incidentTypeData = {
                        id: this.props.incidentTypeId,
                        name: this.props.incidentTypeName,
                        entityType: 'msdyn_incidenttype'
                    };
                }

                if (!incidentTypeData) {
                    throw new Error('No incident type available');
                }

                // STEP 3 & 4: OPTIMIZATION - Get work order type AND department in ONE API call
                const incidentData = await WorkOrderHelpers.getIncidentTypeData(incidentTypeData.id);
                const workOrderTypeData = incidentData?.workOrderType;
                let departmentData = incidentData?.department || null;

                // If not found, get from user (SetDepartment fallback)
                if (!departmentData) {
                    departmentData = await WorkOrderHelpers.setDepartmentFromUser(userId);
                }

                // Fallback to props
                if (!departmentData && this.props.organizationUnitId && this.props.organizationUnitName) {
                    departmentData = {
                        id: this.props.organizationUnitId,
                        name: this.props.organizationUnitName
                    };
                }

                if (!departmentData) {
                    throw new Error('No department available');
                }

                // STEP 5: Prepare campaign data
                let campaignData: { id: string; name: string } | undefined;
                let parentCampaignData: { id: string; name: string } | undefined;

                if (this.state.selectedCampaignId && this.state.selectedCampaignName) {
                    campaignData = {
                        id: this.state.selectedCampaignId,
                        name: this.state.selectedCampaignName
                    };

                    // Set parent campaign (SetParentCampaign logic)
                    parentCampaignData = campaignData;
                } else if (this.props.activePatrolId && this.props.activePatrolName) {
                    campaignData = {
                        id: this.props.activePatrolId,
                        name: this.props.activePatrolName
                    };
                    parentCampaignData = campaignData;
                }

                // STEP 6: Determine anonymous customer flag
                const anonymousCustomer = 
                    this.state.selectedInspectionType === 4 || 
                    (this.state.isAnonymous && [3, 6, 7].includes(this.state.selectedInspectionType!));

                // STEP 7: Detect if created from mobile (DetectMObileCreation)
                const createdFromMobile = WorkOrderHelpers.isMobileClient();

                // STEP 8: Validate data
                const validation = WorkOrderHelpers.validateWorkOrderData({
                    serviceAccount: serviceAccountData,
                    incidentType: incidentTypeData,
                    department: departmentData
                });

                if (!validation.isValid) {
                    throw new Error(`Missing required fields: ${validation.missingFields.join(', ')}`);
                }

                // STEP 9: Create work order using helper
                const workOrderId = await WorkOrderHelpers.createWorkOrder({
                    subAccount: {
                        id: accountId,
                        name: accountName
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
                    createdFromMobile: createdFromMobile
                });

                if (!workOrderId) {
                    throw new Error('Failed to create work order');
                }

                console.log('Work order created successfully:', workOrderId);

                // STEP 10: Create auto booking if from mobile (createAutoBookingOnWorkOrderCreate)
                // OPTIMIZATION: createAutoBooking now parallelizes resource/status fetch internally
                if (createdFromMobile) {
                    const bookingId = await WorkOrderHelpers.createAutoBooking(workOrderId, userId);
                    if (bookingId) {
                        console.log('Auto booking created:', bookingId);
                    }
                }

                // STEP 11: Navigate to the created work order
                await this.xrm.Navigation.openForm({
                    entityName: "msdyn_workorder",
                    entityId: workOrderId,
                    openInNewWindow: true
                });

                // Close the modal
                if (this.props.onClose) {
                    this.props.onClose();
                }

            } catch (error: any) {
                console.error('Error creating work order:', error);
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

                // Get or create account
                const accountId = await this.searchOrCreateAccount();
                if (!accountId) {
                    throw new Error('Failed to get account');
                }

                // Check if campaign or incident type is missing
                const needsCampaign = !this.state.selectedCampaignId && !this.props.activePatrolId;
                const needsIncidentType = !this.state.selectedIncidentTypeId && !this.props.incidentTypeId;

                if (needsCampaign || needsIncidentType) {
                    // Show selection popup with only the missing fields
                    // OPTIMIZATION: Store accountId for reuse in handleContinueWithSelections
                    this.pendingAccountId = accountId;
                    this.setState({
                        showCampaignIncidentPopup: true,
                        popupShowCampaign: needsCampaign,
                        popupShowIncidentType: needsIncidentType,
                        loading: false
                    });
                } else {
                    // Has both campaign and incident type, can create work order directly
                    await this.createWorkOrder(accountId);
                    this.setState({ loading: false });
                }

            } catch (error: any) {
                console.error('Error in handleStart:', error);
                this.setState({
                    error: error.message || 'Error starting inspection',
                    loading: false
                });
            }
        };

        private handleContinueWithSelections = async (): Promise<void> => {
            try {
                // Validate that incident type is selected
                if (!this.state.selectedIncidentTypeId) {
                    this.setState({ 
                        error: this.strings.SelectIncidentType || 'Please select an incident type',
                        loading: false 
                    });
                    return;
                }

                this.setState({ loading: true, error: null, showCampaignIncidentPopup: false });

                // OPTIMIZATION: Reuse account ID from handleStart instead of re-searching
                const accountId = this.pendingAccountId || await this.searchOrCreateAccount();
                this.pendingAccountId = null; // Clear after use
                if (!accountId) {
                    throw new Error('Failed to get account');
                }

                await this.createWorkOrder(accountId);
                this.setState({ loading: false });

            } catch (error: any) {
                console.error('Error creating work order:', error);
                this.setState({
                    error: error.message || 'Error creating work order',
                    loading: false
                });
            }
        };

        // =====================================================================
        // FIELD VISIBILITY
        // =====================================================================

        private shouldShowField = (field: 'qataryId' | 'name' | 'crNumber' | 'id' | 'carColor' | 'vehicleBrand'): boolean => {
            const { selectedInspectionType, isAnonymous } = this.state;
            if (!selectedInspectionType) return false;

            // If anonymous is checked for types 3, 6, or 7, hide all fields
            if (isAnonymous && [3, 6, 7].includes(selectedInspectionType)) {
                return false;
            }

            if (['id', 'carColor', 'vehicleBrand'].includes(field)) {
                return selectedInspectionType === 1;
            }

            if (['qataryId', 'name'].includes(field)) {
                return [2, 3, 6].includes(selectedInspectionType);
            }

            if (field === 'crNumber') {
                return [5, 7].includes(selectedInspectionType);
            }

            return false;
        };

        private shouldShowAnonymousCheckbox = (): boolean => {
            const { selectedInspectionType } = this.state;
            return selectedInspectionType !== null && [3, 6, 7].includes(selectedInspectionType);
        };

        // =====================================================================
        // RENDER
        // =====================================================================

        render() {
            if (!this.props.isOpen) {
                return null;
            }

            const {
                isRTL, inspectionTypes, selectedInspectionType, qataryId, name, crNumber, id,
                carColor, vehicleBrand, vehicleBrands, loading, error, showCampaignIncidentPopup,
                campaigns, incidentTypes, selectedCampaignId, selectedIncidentTypeId, isAnonymous
            } = this.state;

            // Styles
            const containerStyle: React.CSSProperties = {
                position: 'fixed',
                top: 0,
                left: 0,
                width: '100%',
                height: '100%',
                backgroundColor: 'rgba(0, 0, 0, 0.5)',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                zIndex: 1000,
                direction: isRTL ? 'rtl' : 'ltr',
            };

            const modalStyle: React.CSSProperties = {
                backgroundColor: 'white',
                borderRadius: 2,
                padding: 24,
                maxWidth: 500,
                width: '90%',
                maxHeight: '90vh',
                overflowY: 'auto',
                boxShadow: '0 0 10px rgba(0, 0, 0, 0.3)',
            };

            const titleStyle: React.CSSProperties = {
                fontSize: 18,
                fontWeight: 600,
                marginBottom: 20,
                color: '#333333',
                fontFamily: '"Segoe UI", sans-serif',
            };

            const fieldStyle: React.CSSProperties = {
                marginBottom: 16,
            };

            const labelStyle: React.CSSProperties = {
                display: 'block',
                fontSize: 14,
                fontWeight: 600,
                marginBottom: 4,
                color: '#323130',
                fontFamily: '"Segoe UI", sans-serif',
            };

            const inputStyle: React.CSSProperties = {
                width: '100%',
                padding: '6px 8px',
                border: '1px solid #605e5c',
                borderRadius: 2,
                fontSize: 14,
                fontFamily: '"Segoe UI", sans-serif',
                boxSizing: 'border-box',
                outline: 'none',
            };

            const buttonContainerStyle: React.CSSProperties = {
                display: 'flex',
                gap: 8,
                marginTop: 24,
                justifyContent: 'flex-end',
            };

            const buttonStyle: React.CSSProperties = {
                padding: '8px 16px',
                border: '1px solid transparent',
                borderRadius: 2,
                fontSize: 14,
                fontWeight: 600,
                cursor: loading ? 'not-allowed' : 'pointer',
                fontFamily: '"Segoe UI", sans-serif',
                transition: 'background-color 0.1s ease',
                minWidth: 80,
            };

            const startButtonStyle: React.CSSProperties = {
                ...buttonStyle,
                backgroundColor: loading ? '#d3d3d3' : '#0078d4',
                color: 'white',
            };

            const closeButtonStyle: React.CSSProperties = {
                ...buttonStyle,
                backgroundColor: 'white',
                color: '#323130',
                border: '1px solid #8a8886',
            };

            const errorStyle: React.CSSProperties = {
                color: '#a4262c',
                fontSize: 12,
                marginBottom: 16,
                padding: 8,
                backgroundColor: '#fde7e9',
                border: '1px solid #a4262c',
                borderRadius: 2,
                fontFamily: '"Segoe UI", sans-serif',
            };

            const checkboxContainerStyle: React.CSSProperties = {
                marginBottom: 16,
                display: 'flex',
                alignItems: 'center',
                gap: 8,
            };

            const checkboxStyle: React.CSSProperties = {
                width: 16,
                height: 16,
                cursor: loading ? 'not-allowed' : 'pointer',
            };

            const isInspectionTypeDisabled = this.props.lockInspectionType || loading;

            // Campaign/Incident Popup - show only missing fields
            if (showCampaignIncidentPopup) {
                // Determine which fields need to be shown based on initial check when popup opened
                const showCampaignField = this.state.popupShowCampaign;
                const showIncidentField = this.state.popupShowIncidentType;

                // Dynamic title based on what's missing
                let popupTitle = '';
                if (showCampaignField && showIncidentField) {
                    popupTitle = this.strings.SelectCampaign + ' / ' + this.strings.SelectIncidentType;
                } else if (showCampaignField) {
                    popupTitle = this.strings.SelectCampaign;
                } else if (showIncidentField) {
                    popupTitle = this.strings.SelectIncidentType;
                }

                return React.createElement(
                    'div',
                    { style: containerStyle },
                    React.createElement(
                        'div',
                        { style: modalStyle },
                        React.createElement('h2', { style: titleStyle }, popupTitle),

                        error && React.createElement('div', { style: errorStyle }, error),

                        // Campaign Selection - show only if missing
                        showCampaignField && React.createElement(
                            'div',
                            { style: fieldStyle },
                            React.createElement('label', { style: labelStyle }, this.strings.Campaign),
                            React.createElement(
                                'select',
                                {
                                    value: selectedCampaignId || '',
                                    onChange: (e: { target: { value: string; }; }) => this.setState({
                                        selectedCampaignId: e.target.value,
                                        selectedCampaignName: campaigns.find(c => c.id === e.target.value)?.name
                                    }),
                                    disabled: loading,
                                    style: inputStyle,
                                },
                                React.createElement('option', { value: '' }, '--'),
                                campaigns.map((c) =>
                                    React.createElement('option', { key: c.id, value: c.id }, c.name)
                                )
                            )
                        ),

                        // Incident Type Selection - show only if missing
                        showIncidentField && React.createElement(
                            'div',
                            { style: fieldStyle },
                            React.createElement('label', { style: labelStyle }, this.strings.IncidentType),
                            React.createElement(
                                'select',
                                {
                                    value: selectedIncidentTypeId || '',
                                    onChange: (e: { target: { value: string; }; }) => this.setState({
                                        selectedIncidentTypeId: e.target.value,
                                        selectedIncidentTypeName: incidentTypes.find(it => it.id === e.target.value)?.name
                                    }),
                                    disabled: loading,
                                    style: inputStyle,
                                },
                                React.createElement('option', { value: '' }, '--'),
                                incidentTypes.map((it) =>
                                    React.createElement('option', { key: it.id, value: it.id }, it.name)
                                )
                            )
                        ),

                        React.createElement(
                            'div',
                            { style: buttonContainerStyle },
                            React.createElement(
                                'button',
                                {
                                    onClick: this.handleContinueWithSelections,
                                    disabled: loading,
                                    style: startButtonStyle,
                                },
                                loading ? this.strings.Loading : this.strings.Continue
                            ),
                            React.createElement(
                                'button',
                                {
                                    onClick: () => this.setState({ showCampaignIncidentPopup: false }),
                                    disabled: loading,
                                    style: closeButtonStyle,
                                },
                                this.strings.Close
                            )
                        )
                    )
                );
            }

            // Main Form
            return React.createElement(
                'div',
                { style: containerStyle },
                React.createElement(
                    'div',
                    { style: modalStyle },
                    React.createElement('h2', { style: titleStyle }, this.strings.StartMultiTypeInspection),

                    error && React.createElement('div', { style: errorStyle }, error),

                    // Inspection Type
                    React.createElement(
                        'div',
                        { style: fieldStyle },
                        React.createElement('label', { style: labelStyle }, this.strings.InspectionType),
                        React.createElement(
                            'select',
                            {
                                value: selectedInspectionType || '',
                                onChange: this.handleInspectionTypeChange,
                                disabled: isInspectionTypeDisabled,
                                style: {
                                    ...inputStyle,
                                    backgroundColor: isInspectionTypeDisabled ? '#f3f2f1' : 'white',
                                    cursor: isInspectionTypeDisabled ? 'not-allowed' : 'pointer'
                                },
                            },
                            React.createElement('option', { value: '' }, this.strings.chooseInspectionType),
                            inspectionTypes.map((type) =>
                                React.createElement('option', { key: type.value, value: type.value }, type.label)
                            )
                        )
                    ),

                    // Anonymous Checkbox (for types 3, 6, 7)
                    this.shouldShowAnonymousCheckbox() && React.createElement(
                        'div',
                        { style: checkboxContainerStyle },
                        React.createElement('input', {
                            type: 'checkbox',
                            checked: isAnonymous,
                            onChange: (e) => this.setState({ isAnonymous: e.target.checked }),
                            disabled: loading,
                            style: checkboxStyle,
                        }),
                        React.createElement('label', { 
                            style: { ...labelStyle, marginBottom: 0, cursor: loading ? 'not-allowed' : 'pointer' },
                            onClick: () => !loading && this.setState({ isAnonymous: !isAnonymous })
                        }, this.strings.Anonymous)
                    ),

                    // ID (Vehicle)
                    this.shouldShowField('id') && React.createElement(
                        'div',
                        { style: fieldStyle },
                        React.createElement('label', { style: labelStyle }, this.strings.ID),
                        React.createElement('input', {
                            type: 'text',
                            value: id,
                            onChange: (e) => this.handleInputChange('id', e.target.value),
                            disabled: loading,
                            style: inputStyle,
                        })
                    ),

                    // Car Color
                    this.shouldShowField('carColor') && React.createElement(
                        'div',
                        { style: fieldStyle },
                        React.createElement('label', { style: labelStyle }, this.strings.CarColor),
                        React.createElement('input', {
                            type: 'text',
                            value: carColor,
                            onChange: (e) => this.handleInputChange('carColor', e.target.value),
                            disabled: loading,
                            style: inputStyle,
                        })
                    ),

                    // Vehicle Brand
                    this.shouldShowField('vehicleBrand') && React.createElement(
                        'div',
                        { style: fieldStyle },
                        React.createElement('label', { style: labelStyle }, this.strings.VehicleBrand),
                        React.createElement(
                            'select',
                            {
                                value: vehicleBrand ?? '',
                                onChange: (e: { target: { value: string; }; }) => {
                                    const val = e.target.value ? parseInt(e.target.value, 10) : null;
                                    this.setState({ vehicleBrand: val });
                                },
                                disabled: loading,
                                style: inputStyle,
                            },
                            React.createElement('option', { value: '' }, '--'),
                            vehicleBrands.map((vb) =>
                                React.createElement('option', { key: vb.value, value: vb.value }, vb.label)
                            )
                        )
                    ),

                    // Qatary ID
                    this.shouldShowField('qataryId') && React.createElement(
                        'div',
                        { style: fieldStyle },
                        React.createElement('label', { style: labelStyle }, this.strings.QataryID),
                        React.createElement('input', {
                            type: 'text',
                            value: qataryId,
                            onChange: (e) => this.handleInputChange('qataryId', e.target.value),
                            disabled: loading,
                            style: inputStyle,
                        })
                    ),

                    // Name
                    this.shouldShowField('name') && React.createElement(
                        'div',
                        { style: fieldStyle },
                        React.createElement('label', { style: labelStyle }, this.strings.Name),
                        React.createElement('input', {
                            type: 'text',
                            value: name,
                            onChange: (e) => this.handleInputChange('name', e.target.value),
                            disabled: loading,
                            style: inputStyle,
                        })
                    ),

                    // CR Number
                    this.shouldShowField('crNumber') && React.createElement(
                        'div',
                        { style: fieldStyle },
                        React.createElement('label', { style: labelStyle },
                            selectedInspectionType === 7 ? this.strings.MonourNumber : this.strings.CRNumber
                        ),
                        React.createElement('input', {
                            type: 'text',
                            value: crNumber,
                            onChange: (e) => this.handleInputChange('crNumber', e.target.value),
                            disabled: loading,
                            style: inputStyle,
                        })
                    ),

                    // Buttons
                    React.createElement(
                        'div',
                        { style: buttonContainerStyle },
                        React.createElement(
                            'button',
                            {
                                onClick: this.handleStart,
                                disabled: loading,
                                style: startButtonStyle,
                            },
                            loading ? this.strings.Loading : this.strings.Start
                        ),
                        React.createElement(
                            'button',
                            {
                                onClick: this.props.onClose,
                                disabled: loading,
                                style: closeButtonStyle,
                            },
                            this.strings.Close
                        )
                    )
                )
            );
        }
    }