/* eslint-disable */
import * as React from 'react';

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
    defaultInspectionType?: number; // NEW: Default inspection type to pre-select
    lockInspectionType?: boolean; // NEW: Whether to lock the inspection type field
}

interface IMultiTypeInspectionState {
    isRTL: boolean;
    inspectionTypes: Array<{ value: number; label: string }>;
    selectedInspectionType: number | null;
    qataryId: string;
    name: string;
    crNumber: string;
    id: string;
    carColor: string; // NEW: Car color field for vehicles
    vehicleBrand: string; // NEW: Vehicle brand field for vehicles
    loading: boolean;
    error: string | null;
    accountTypeRecord: any | null; // Store the retrieved account type record
}

interface LocalizedStrings {
    StartMultiTypeInspection: string;
    InspectionType: string;
    QataryID: string;
    Name: string;
    CRNumber: string;
    ID: string;
    CarColor: string; // NEW
    VehicleBrand: string; // NEW
    Start: string;
    Close: string;
    Loading: string;
    PleaseSelectInspectionType: string;
    PleaseEnterRequiredFields: string;
    SearchingAccount: string;
    CreatingAccount: string;
    Error: string;
    AccountNotFound: string;
    AccountCreated: string;
    chooseInspectionType: string;
}

// Cache constants
const INSPECTION_TYPES_CACHE_KEY = 'MOCI_InspectionTypes_Cache';
const CACHE_DURATION = 0; // 1min

interface InspectionTypesCache {
    types: Array<{ value: number; label: string }>;
    timestamp: number;
}

export class MultiTypeInspection extends React.Component<IMultiTypeInspectionProps, IMultiTypeInspectionState> {
    private strings: LocalizedStrings;

    constructor(props: IMultiTypeInspectionProps) {
        super(props);

        const userSettings = (props.context as any).userSettings;
        const rtlLanguages = [1025, 1037, 1054, 1056, 1065, 1068, 1069, 1101, 1114, 1119];
        const isRTL = rtlLanguages.includes(userSettings?.languageId);

        // Load localized strings
        this.strings = {
            StartMultiTypeInspection: props.context.resources.getString("StartMultiTypeInspection"),
            InspectionType: props.context.resources.getString("InspectionType"),
            QataryID: props.context.resources.getString("QataryID"),
            Name: props.context.resources.getString("Name"),
            CRNumber: props.context.resources.getString("CRNumber"),
            ID: props.context.resources.getString("ID"),
            CarColor: props.context.resources.getString("CarColor") || "Car Color", // NEW
            VehicleBrand: props.context.resources.getString("VehicleBrand") || "Vehicle Brand", // NEW
            Start: props.context.resources.getString("Start"),
            Close: props.context.resources.getString("Close"),
            Loading: props.context.resources.getString("Loading"),
            PleaseSelectInspectionType: props.context.resources.getString("PleaseSelectInspectionType"),
            PleaseEnterRequiredFields: props.context.resources.getString("PleaseEnterRequiredFields"),
            SearchingAccount: props.context.resources.getString("SearchingAccount"),
            CreatingAccount: props.context.resources.getString("CreatingAccount"),
            Error: props.context.resources.getString("Error"),
            AccountNotFound: props.context.resources.getString("AccountNotFound"),
            AccountCreated: props.context.resources.getString("AccountCreated"),
            chooseInspectionType: props.context.resources.getString("chooseInspectionType"),
        };

        this.state = {
            isRTL: isRTL,
            inspectionTypes: [],
            selectedInspectionType: props.defaultInspectionType || null, // NEW: Use default if provided
            qataryId: '',
            name: '',
            crNumber: '',
            id: '',
            carColor: '', // NEW
            vehicleBrand: '', // NEW
            loading: false,
            error: null,
            accountTypeRecord: null,
        };
    }

    async componentDidMount(): Promise<void> {
        await this.loadInspectionTypes();

        // NEW: If default inspection type is provided, retrieve account type
        if (this.props.defaultInspectionType) {
            const accountTypeRecord = await this.retrieveAccountTypeByOptionSet(this.props.defaultInspectionType);
            this.setState({ accountTypeRecord });
        }
    }

    // Get inspection types from cache
    private getInspectionTypesFromCache = (): Array<{ value: number; label: string }> | null => {
        try {
            const cached = localStorage.getItem(INSPECTION_TYPES_CACHE_KEY);
            if (!cached) return null;

            const cacheData: InspectionTypesCache = JSON.parse(cached);
            const now = Date.now();

            // Check if cache is still valid
            if (now - cacheData.timestamp > CACHE_DURATION) {
                localStorage.removeItem(INSPECTION_TYPES_CACHE_KEY);
                return null;
            }

            console.log("Using cached inspection types");
            return cacheData.types;
        } catch (error) {
            console.error("Error reading inspection types cache:", error);
            return null;
        }
    };

    // Save inspection types to cache
    private saveInspectionTypesToCache = (types: Array<{ value: number; label: string }>): void => {
        try {
            const cacheData: InspectionTypesCache = {
                types,
                timestamp: Date.now()
            };
            localStorage.setItem(INSPECTION_TYPES_CACHE_KEY, JSON.stringify(cacheData));
            console.log("Inspection types cached successfully");
        } catch (error) {
            console.error("Error saving inspection types cache:", error);
        }
    };

    private loadInspectionTypes = async (): Promise<void> => {
        try {
            // Try to get from cache first
            const cachedTypes = this.getInspectionTypesFromCache();
            if (cachedTypes) {
                this.setState({ inspectionTypes: cachedTypes });
                return;
            }

            // If not in cache, fetch from metadata
            const xrm: Xrm.XrmStatic = (window.parent as any).Xrm || (window as any).Xrm;

            const entityMetadata = await xrm.Utility.getEntityMetadata('account', ['duc_accountinspectiontype']);
            const attribute = (entityMetadata as any).Attributes.get('duc_accountinspectiontype');

            if (attribute && attribute.OptionSet) {
                const types = Object.values(attribute.OptionSet).map((opt: any) => ({
                    value: opt.value,
                    label: opt.text,
                }));

                // Save to cache
                this.saveInspectionTypesToCache(types);

                this.setState({ inspectionTypes: types });
            }
        } catch (error) {
            console.error('Error loading inspection types:', error);
            this.setState({ error: 'Failed to load inspection types' });
        }
    };

    /**
     * Retrieve account type record based on the selected inspection type option set value
     */
    private retrieveAccountTypeByOptionSet = async (optionSetValue: number): Promise<any | null> => {
        try {
            const xrm: Xrm.XrmStatic = (window.parent as any).Xrm || (window as any).Xrm;

            // Query duc_AccountType entity where duc_accounttype equals the option set value
            const result = await xrm.WebApi.retrieveMultipleRecords(
                'duc_accounttype',
                `?$select=duc_accounttypeid,duc_name,duc_accounttype&$filter=duc_accounttype eq ${optionSetValue}&$top=1`
            );

            if (result?.entities && result.entities.length > 0) {
                console.log('Account type record found:', result.entities[0]);
                return result.entities[0];
            } else {
                console.warn('No account type record found for option set value:', optionSetValue);
                return null;
            }
        } catch (error) {
            console.error('Error retrieving account type:', error);
            return null;
        }
    };

    private handleInspectionTypeChange = async (e: React.ChangeEvent<HTMLSelectElement>): Promise<void> => {
        const value = e.target.value ? parseInt(e.target.value, 10) : null;
        
        this.setState({
            selectedInspectionType: value,
            qataryId: '',
            name: '',
            crNumber: '',
            id: '',
            carColor: '', // NEW: Reset car color
            vehicleBrand: '', // NEW: Reset vehicle brand
            error: null,
            accountTypeRecord: null,
        });

        // Retrieve account type record when inspection type is selected
        if (value !== null) {
            const accountTypeRecord = await this.retrieveAccountTypeByOptionSet(value);
            this.setState({ accountTypeRecord });
        }
    };

    private handleInputChange = (field: keyof IMultiTypeInspectionState, value: string): void => {
        this.setState({ [field]: value } as any);
    };

    private getRequiredFields = (): Array<keyof IMultiTypeInspectionState> => {
        const { selectedInspectionType } = this.state;
        const requiredFields: Array<keyof IMultiTypeInspectionState> = [];

        if (!selectedInspectionType) return [];

        // Type 1 (Vehicle) => ID, car color, and vehicle brand
        if (selectedInspectionType === 1) {
            requiredFields.push('id', 'carColor', 'vehicleBrand');
        }
        // Type 2 (Individual), 3 (Cabin), or 6 (Wilderness camps) => Qatary ID and Name
        else if (selectedInspectionType === 2 || selectedInspectionType === 3 || selectedInspectionType === 6) {
            requiredFields.push('qataryId', 'name');
        }
        // Type 5 (Company) => CR Number
        else if (selectedInspectionType === 5) {
            requiredFields.push('crNumber');
        }
        // Type 4 (Anonymous) => No required fields (uses unknown account from props)

        return requiredFields;
    };

    private validateFields = (): boolean => {
        const { selectedInspectionType } = this.state;

        if (!selectedInspectionType) {
            this.setState({ error: this.strings.PleaseSelectInspectionType });
            return false;
        }

        // Type 4 (Anonymous) doesn't need field validation (uses unknown account from props)
        // The account availability check happens in searchOrCreateAccount
        if (selectedInspectionType === 4) {
            return true;
        }

        // All other types require identifier fields
        const requiredFields = this.getRequiredFields();
        for (const field of requiredFields) {
            if (!this.state[field]) {
                this.setState({ error: this.strings.PleaseEnterRequiredFields });
                return false;
            }
        }

        return true;
    };

    private getIdentifierValue = (): string => {
        const { selectedInspectionType, qataryId, crNumber, id } = this.state;

        // Type 1 (Vehicle) => ID
        if (selectedInspectionType === 1) {
            return id;
        }
        // Type 2 (Individual), 3 (Cabin), or 6 (Wilderness camps) => Qatary ID
        else if (selectedInspectionType === 2 || selectedInspectionType === 3 || selectedInspectionType === 6) {
            return qataryId;
        }
        // Type 5 (Company) => CR Number
        else if (selectedInspectionType === 5) {
            return crNumber;
        }
        // Type 4 (Anonymous) => Empty (uses unknown account from props)
        
        return '';
    };

    private getAccountName = (): string => {
        const { selectedInspectionType, name, qataryId, crNumber, id, carColor, vehicleBrand } = this.state;

        // Type 1 (Vehicle) => "Vehicle + ID + Color + Brand"
        if (selectedInspectionType === 1) {
            return `Vehicle ${id} ${carColor} ${vehicleBrand}`;
        }
        // Type 2 (Individual) => "Individual + Qatary ID + Name"
        else if (selectedInspectionType === 2) {
            return `Individual ${qataryId} ${name}`;
        }
        // Type 3 (Cabin) => Use provided name
        else if (selectedInspectionType === 3) {
            return name || `Cabin ${qataryId}`;
        }
        // Type 6 (Wilderness camps) => Use provided name (similar to cabin)
        else if (selectedInspectionType === 6) {
            return name || `Wilderness Camp ${qataryId}`;
        }
        // Type 5 (Company) => "Company + CR Number"
        else if (selectedInspectionType === 5) {
            return `Company ${crNumber}`;
        }
        // Type 4 (Anonymous) => "Anonymous Account" (uses unknown account from props)
        else if (selectedInspectionType === 4) {
            return 'Anonymous Account';
        }

        return 'Account';
    };

    private getCurrentLocation = (): Promise<{ latitude: number; longitude: number }> => {
        return new Promise((resolve, reject) => {
            const xrm: Xrm.XrmStatic = (window.parent as any).Xrm || (window as any).Xrm;

            // Try to use Xrm.Device API first (works in Field Service Mobile app)
            if (xrm && xrm.Device && xrm.Device.getCurrentPosition) {
                console.log('Using Xrm.Device.getCurrentPosition (Field Service Mobile)');
                xrm.Device.getCurrentPosition().then(
                    (location: any) => {
                        resolve({
                            latitude: location.coords.latitude,
                            longitude: location.coords.longitude
                        });
                    },
                    (error: any) => {
                        console.error('Xrm.Device.getCurrentPosition failed, falling back to navigator.geolocation:', error);
                        return;
                        // Fallback to navigator.geolocation
                        // this.getLocationViaNavigator(resolve, reject);
                    }
                );
            } else {
                console.log('Using navigator.geolocation (Web/PWA)');
                return;
            }
        });
    };


    private createAddressInformation = async (accountId: string, accountName: string): Promise<void> => {
        try {
            const xrm: Xrm.XrmStatic = (window.parent as any).Xrm || (window as any).Xrm;

            // Get current location
            const location = await this.getCurrentLocation();
            if (location) {
                // Format today's date (YYYY-MM-DD)
                const today = new Date();
                const formattedDate = today.toISOString().split('T')[0].replace(/\//g, '-');

                // Create address name: AccountName + Date
                const addressName = `${accountName} ${formattedDate}`;

                // Create address information record
                const addressData: any = {
                    duc_name: addressName,
                    duc_latitude: location.latitude,
                    duc_longitude: location.longitude,
                    'duc_Account@odata.bind': `/accounts(${accountId})`
                };

                const createdAddress = await xrm.WebApi.createRecord('duc_addressinformation', addressData);
                // console.log('Address information created:', createdAddress.id);
                // console.log('Address details:', {
                //     name: addressName,
                //     latitude: location.latitude,
                //     longitude: location.longitude,
                //     accountId: accountId
                // });
            }

        } catch (error: any) {
            console.error('Error creating address information:', error);
            // alert('Failed to create address information: ' + (error.message || error));
            return;

            // Don't throw error - we don't want to block the inspection if address creation fails
            // Just log it for troubleshooting
        }
    };

    private searchOrCreateAccount = async (): Promise<string | null> => {
        try {
            this.setState({ loading: true, error: null });

            const xrm: Xrm.XrmStatic = (window.parent as any).Xrm || (window as any).Xrm;
            const identifierValue = this.getIdentifierValue();
            const { accountTypeRecord, selectedInspectionType, carColor, vehicleBrand } = this.state;

            // For Anonymous type (4), ONLY use the unknown account from props
            if (selectedInspectionType === 4) {
                if (this.props.unknownAccountId) {
                    console.log('Using unknown account for anonymous inspection:', this.props.unknownAccountId);
                    // Create address information for the unknown account
                    await this.createAddressInformation(this.props.unknownAccountId, this.props.unknownAccountName || 'Unknown Account');
                    return this.props.unknownAccountId;
                } else {
                    // No unknown account available - show error
                    this.setState({ 
                        error: 'No anonymous account',
                        loading: false 
                    });
                    return null;
                }
            }

            // Identifier is mandatory for all other types
            if (!identifierValue) {
                this.setState({ error: this.strings.PleaseEnterRequiredFields, loading: false });
                return null;
            }

            // Search for existing account by duc_accountidentifier AND account type
            let filterQuery = `duc_accountidentifier eq '${identifierValue}'`;
            
            // Add account type filter if we have the account type record
            if (accountTypeRecord?.duc_accounttypeid) {
                filterQuery += ` and _duc_newaccounttype_value eq ${accountTypeRecord.duc_accounttypeid}`;
            }

            const searchResults = await xrm.WebApi.retrieveMultipleRecords(
                'account',
                `?$select=accountid,name&$filter=${filterQuery}`
            );

            if (searchResults?.entities && searchResults.entities.length > 0) {
                const accountId = searchResults.entities[0].accountid;
                const accountName = searchResults.entities[0].name;
                console.log('Account found:', accountId);

                // For vehicles (type 1), update the existing account with color and brand
                if (selectedInspectionType === 1) {
                    const updateData: any = {};
                    
                    if (carColor) {
                        updateData.duc_VehicleColor = carColor;
                    }
                    if (vehicleBrand) {
                        updateData.duc_VehicleType = vehicleBrand;
                    }

                    if (Object.keys(updateData).length > 0) {
                        await xrm.WebApi.updateRecord('account', accountId, updateData);
                        console.log('Vehicle account updated with color and brand');
                    }
                }

                // Create address information for the existing account
                await this.createAddressInformation(accountId, accountName);

                return accountId;
            }

            // Account doesn't exist, create new one
            const accountName = this.getAccountName();
            const newAccount: any = {
                name: accountName,
                duc_accountidentifier: identifierValue,
                duc_accountinspectiontype: selectedInspectionType
            };

            // Set account type lookup if found
            if (accountTypeRecord?.duc_accounttypeid) {
                newAccount['duc_NewAccountType@odata.bind'] = `/duc_accounttypes(${accountTypeRecord.duc_accounttypeid})`;
            }

            // NEW: For vehicles (type 1), add color and brand to the account
            if (selectedInspectionType === 1) {
                if (carColor) {
                    newAccount.duc_VehicleColor = carColor;
                }
                if (vehicleBrand) {
                    newAccount.duc_VehicleType = vehicleBrand;
                }
            }

            const createdAccount = await xrm.WebApi.createRecord('account', newAccount);
            const newAccountId = createdAccount?.id;

            console.log('Account created:', newAccountId);

            // Create address information for the new account
            await this.createAddressInformation(newAccountId, accountName);

            return newAccountId;
        } catch (error: any) {
            console.error('Error searching/creating account:', error);
            this.setState({ error: error.message || 'Error processing account', loading: false });
            return null;
        }
    };

    private handleStart = async (): Promise<void> => {
        if (!this.validateFields()) {
            return;
        }

        try {
            this.setState({ loading: true, error: null });

            const accountId = await this.searchOrCreateAccount();

            if (accountId) {
                console.log('Processing with account ID:', accountId);
                // Open work order quick create form with the account and inspection type as default values
                await this.openWorkOrderForm(accountId);

                // Close the modal after successfully opening the work order form
                if (this.props.onClose) {
                    this.props.onClose();
                }
            }
        } catch (error: any) {
            console.error('Error in handleStart:', error);
            this.setState({ error: error.message || 'Error starting inspection', loading: false });
        } finally {
            this.setState({ loading: false });
        }
    };

    private openWorkOrderForm = async (accountId: string): Promise<void> => {
        try {
            const xrm: Xrm.XrmStatic = (window.parent as any).Xrm || (window as any).Xrm;

            // Get account name
            const accountRecord = await xrm.WebApi.retrieveRecord('account', accountId, '?$select=name');
            const accountName = accountRecord?.name || '';

            // Prepare default values for the work order quick create form
            const defaultValues: any = {
                // Set the sub-account (duc_subaccount)
                duc_subaccount: [
                    {
                        id: accountId,
                        name: accountName,
                        entityType: "account"
                    }
                ],
                // Set the service account (msdyn_serviceaccount)
                msdyn_serviceaccount: [
                    {
                        id: accountId,
                        name: accountName,
                        entityType: "account"
                    }
                ],
                // Set the inspection type (duc_accountinspectiontype) - this is an option set value
                duc_accountinspectiontype: this.state.selectedInspectionType
            };

            // For Anonymous type (4), set the anonymous customer field
            if (this.state.selectedInspectionType === 4) {
                defaultValues.duc_anonymouscustomer = true;
            }

            // If we have an active patrol campaign, set it as default
            if (this.props.activePatrolId && this.props.activePatrolName) {
                defaultValues.new_campaign = [
                    {
                        id: this.props.activePatrolId,
                        name: this.props.activePatrolName,
                        entityType: "new_inspectioncampaign"
                    }
                ];
            }

            // If we have incident type ID (from organization unit like Natural Reserve), set it as default
            if (this.props.incidentTypeId && this.props.incidentTypeName) {
                defaultValues.msdyn_primaryincidenttype = [
                    {
                        id: this.props.incidentTypeId,
                        name: this.props.incidentTypeName,
                        entityType: "msdyn_incidenttype"
                    }
                ];
            }

            // If we have organization unit (department), set it as default
            if (this.props.organizationUnitId && this.props.organizationUnitName) {
                defaultValues.duc_department = [
                    {
                        id: this.props.organizationUnitId,
                        name: this.props.organizationUnitName,
                        entityType: "msdyn_organizationalunit"
                    }
                ];
            }

            console.log('Opening work order form with default values:', defaultValues);

            // Open the work order quick create form
            await xrm.Navigation.openForm(
                { entityName: "msdyn_workorder", useQuickCreateForm: true },
                defaultValues
            );

        } catch (error: any) {
            console.error('Error opening work order form:', error);
            throw error; // Re-throw to be caught by handleStart
        }
    };

    private shouldShowField = (field: 'qataryId' | 'name' | 'crNumber' | 'id' | 'carColor' | 'vehicleBrand'): boolean => {
        const { selectedInspectionType } = this.state;

        if (!selectedInspectionType) return false;

        // Type 1 (Vehicle) => Show ID, car color, and vehicle brand
        if (field === 'id' || field === 'carColor' || field === 'vehicleBrand') {
            return selectedInspectionType === 1;
        }

        // Type 2 (Individual), 3 (Cabin), or 6 (Wilderness camps) => Show Qatary ID and Name
        if (field === 'qataryId' || field === 'name') {
            return selectedInspectionType === 2 || selectedInspectionType === 3 || selectedInspectionType === 6;
        }

        // Type 5 (Company) => Show CR Number
        if (field === 'crNumber') {
            return selectedInspectionType === 5;
        }

        // Type 4 (Anonymous) => Show nothing (uses unknown account from props)
        return false;
    };

    render() {
        if (!this.props.isOpen) {
            return null;
        }

        const {
            isRTL, inspectionTypes, selectedInspectionType, qataryId, name, crNumber, id, carColor, vehicleBrand, loading, error,
        } = this.state;

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
            fontFamily: '"Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif',
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
            fontFamily: '"Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif',
        };

        const inputStyle: React.CSSProperties = {
            width: '100%',
            padding: '6px 8px',
            border: '1px solid #605e5c',
            borderRadius: 2,
            fontSize: 14,
            fontFamily: '"Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif',
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
            fontFamily: '"Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif',
            transition: 'background-color 0.1s ease',
            minWidth: 80,
        };

        const startButtonStyle: React.CSSProperties = {
            ...buttonStyle,
            backgroundColor: loading ? '#d3d3d3' : '#0078d4',
            color: 'white',
            border: '1px solid transparent',
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
            fontFamily: '"Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif',
        };

        // NEW: Determine if inspection type field should be disabled
        const isInspectionTypeDisabled = this.props.lockInspectionType || loading;

        return React.createElement(
            'div',
            { style: containerStyle },
            React.createElement(
                'div',
                { style: modalStyle },
                React.createElement('h2', { style: titleStyle }, this.strings.StartMultiTypeInspection),

                error && React.createElement('div', { style: errorStyle }, error),

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
                            React.createElement(
                                'option',
                                { key: type.value, value: type.value },
                                type.label
                            )
                        )
                    )
                ),

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

                this.shouldShowField('vehicleBrand') && React.createElement(
                    'div',
                    { style: fieldStyle },
                    React.createElement('label', { style: labelStyle }, this.strings.VehicleBrand),
                    React.createElement('input', {
                        type: 'text',
                        value: vehicleBrand,
                        onChange: (e) => this.handleInputChange('vehicleBrand', e.target.value),
                        disabled: loading,
                        style: inputStyle,
                    })
                ),

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

                this.shouldShowField('crNumber') && React.createElement(
                    'div',
                    { style: fieldStyle },
                    React.createElement('label', { style: labelStyle }, this.strings.CRNumber),
                    React.createElement('input', {
                        type: 'text',
                        value: crNumber,
                        onChange: (e) => this.handleInputChange('crNumber', e.target.value),
                        disabled: loading,
                        style: inputStyle,
                    })
                ),

                React.createElement(
                    'div',
                    { style: buttonContainerStyle },
                    React.createElement(
                        'button',
                        {
                            onClick: this.handleStart,
                            disabled: loading,
                            style: startButtonStyle,
                            onMouseEnter: (e: { target: HTMLButtonElement; }) => {
                                if (!loading) {
                                    (e.target as HTMLButtonElement).style.backgroundColor = '#106ebe';
                                }
                            },
                            onMouseLeave: (e: { target: HTMLButtonElement; }) => {
                                if (!loading) {
                                    (e.target as HTMLButtonElement).style.backgroundColor = '#0078d4';
                                }
                            },
                        },
                        loading ? this.strings.Loading : this.strings.Start
                    ),
                    React.createElement(
                        'button',
                        {
                            onClick: this.props.onClose,
                            disabled: loading,
                            style: closeButtonStyle,
                            onMouseEnter: (e: { target: HTMLButtonElement; }) => {
                                if (!loading) {
                                    (e.target as HTMLButtonElement).style.backgroundColor = '#f3f2f1';
                                }
                            },
                            onMouseLeave: (e: { target: HTMLButtonElement; }) => {
                                if (!loading) {
                                    (e.target as HTMLButtonElement).style.backgroundColor = 'white';
                                }
                            },
                        },
                        this.strings.Close
                    )
                )
            )
        );
    }
}