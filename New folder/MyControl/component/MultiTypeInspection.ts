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
}

interface IMultiTypeInspectionState {
    isRTL: boolean;
    inspectionTypes: Array<{ value: number; label: string }>;
    selectedInspectionType: number | null;
    qataryId: string;
    name: string;
    crNumber: string;
    id: string;
    loading: boolean;
    error: string | null;
}

interface LocalizedStrings {
    StartMultiTypeInspection: string;
    InspectionType: string;
    QataryID: string;
    Name: string;
    CRNumber: string;
    ID: string;
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
        };

        this.state = {
            isRTL: isRTL,
            inspectionTypes: [],
            selectedInspectionType: null,
            qataryId: '',
            name: '',
            crNumber: '',
            id: '',
            loading: false,
            error: null,
        };
    }

    async componentDidMount(): Promise<void> {
        await this.loadInspectionTypes();
    }

    private loadInspectionTypes = async (): Promise<void> => {
        try {
            const xrm: Xrm.XrmStatic = (window.parent as any).Xrm || (window as any).Xrm;

            const entityMetadata = await xrm.Utility.getEntityMetadata('account', ['duc_accountinspectiontype']);
            const attribute = (entityMetadata as any).Attributes.get('duc_accountinspectiontype');

            if (attribute && attribute.OptionSet) {
                const types = Object.values(attribute.OptionSet).map((opt: any) => ({
                    value: opt.value,
                    label: opt.text,
                }));
                this.setState({ inspectionTypes: types });
            }
        } catch (error) {
            console.error('Error loading inspection types:', error);
            this.setState({ error: 'Failed to load inspection types' });
        }
    };

    private handleInspectionTypeChange = (e: React.ChangeEvent<HTMLSelectElement>): void => {
        const value = e.target.value ? parseInt(e.target.value, 10) : null;
        this.setState({
            selectedInspectionType: value,
            qataryId: '',
            name: '',
            crNumber: '',
            id: '',
            error: null,
        });
    };

    private handleInputChange = (field: keyof IMultiTypeInspectionState, value: string): void => {
        this.setState({ [field]: value } as any);
    };

    private getRequiredFields = (): Array<keyof IMultiTypeInspectionState> => {
        const { selectedInspectionType } = this.state;
        const requiredFields: Array<keyof IMultiTypeInspectionState> = [];

        if (!selectedInspectionType) return [];

        // Type 1 (Vehicle) => ID only
        if (selectedInspectionType === 1) {
            requiredFields.push('id');
        }
        // Type 2 (Individual) or 3 (Cabin) => Qatary ID and Name
        else if (selectedInspectionType === 2 || selectedInspectionType === 3) {
            requiredFields.push('qataryId', 'name');
        }
        // Type 5 (Company) => CR Number
        else if (selectedInspectionType === 5) {
            requiredFields.push('crNumber');
        }
        // Type 4 (Anonymous) => No required fields

        return requiredFields;
    };

    private validateFields = (): boolean => {
        const { selectedInspectionType } = this.state;

        if (!selectedInspectionType) {
            this.setState({ error: this.strings.PleaseSelectInspectionType });
            return false;
        }

        // Type 4 (Anonymous) doesn't need any fields
        if (selectedInspectionType === 4) {
            return true;
        }

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
        // Type 2 (Individual) or 3 (Cabin) => Qatary ID
        else if (selectedInspectionType === 2 || selectedInspectionType === 3) {
            return qataryId;
        }
        // Type 5 (Company) => CR Number
        else if (selectedInspectionType === 5) {
            return crNumber;
        }
        // Type 4 (Anonymous) => Empty
        return '';
    };

    private getAccountName = (): string => {
        const { selectedInspectionType, name, qataryId, crNumber, id } = this.state;

        // Type 1 (Vehicle) => "Vehicle + ID"
        if (selectedInspectionType === 1) {
            return `Vehicle ${id}`;
        }
        // Type 2 (Individual) => "Individual + Qatary ID + Name"
        else if (selectedInspectionType === 2) {
            return `Individual ${qataryId} ${name}`;
        }
        // Type 3 (Cabin) => Use provided name
        else if (selectedInspectionType === 3) {
            return name || `Cabin ${qataryId}`;
        }
        // Type 5 (Company) => "Company + CR Number"
        else if (selectedInspectionType === 5) {
            return `Company ${crNumber}`;
        }
        // Type 4 (Anonymous) => "Anonymous Account"
        else if (selectedInspectionType === 4) {
            return 'Anonymous Account';
        }

        return 'Account';
    };

    private searchOrCreateAccount = async (): Promise<string | null> => {
        try {
            this.setState({ loading: true, error: null });

            const xrm: Xrm.XrmStatic = (window.parent as any).Xrm || (window as any).Xrm;
            const identifierValue = this.getIdentifierValue();

            // For Anonymous type or if no identifier, create account directly
            if (!identifierValue || this.state.selectedInspectionType === 4) {
                const newAccount: any = {
                    name: this.getAccountName(),
                    duc_accountidentifier: identifierValue || '',
                    duc_accountinspectiontype: this.state.selectedInspectionType
                };

                const createdAccount = await xrm.WebApi.createRecord('account', newAccount);
                const newAccountId = createdAccount?.id;
                console.log('Anonymous account created:', newAccountId);
                return newAccountId;
            }

            // Search for existing account by duc_accountidentifier
            const searchResults = await xrm.WebApi.retrieveMultipleRecords(
                'account',
                `?$select=accountid,name&$filter=duc_accountidentifier eq '${identifierValue}'`
            );

            if (searchResults?.entities && searchResults.entities.length > 0) {
                const accountId = searchResults.entities[0].accountid;
                console.log('Account found:', accountId);
                return accountId;
            }

            // Account doesn't exist, create new one
            const newAccount: any = {
                name: this.getAccountName(),
                duc_accountidentifier: identifierValue,
                duc_accountinspectiontype: this.state.selectedInspectionType
            };

            const createdAccount = await xrm.WebApi.createRecord('account', newAccount);
            const newAccountId = createdAccount?.id;

            console.log('Account created:', newAccountId);
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

    private shouldShowField = (field: 'qataryId' | 'name' | 'crNumber' | 'id'): boolean => {
        const { selectedInspectionType } = this.state;

        if (!selectedInspectionType) return false;

        // Type 1 (Vehicle) => Show ID only
        if (field === 'id') {
            return selectedInspectionType === 1;
        }

        // Type 2 (Individual) or 3 (Cabin) => Show Qatary ID and Name
        if (field === 'qataryId' || field === 'name') {
            return selectedInspectionType === 2 || selectedInspectionType === 3;
        }

        // Type 5 (Company) => Show CR Number
        if (field === 'crNumber') {
            return selectedInspectionType === 5;
        }

        // Type 4 (Anonymous) => Show nothing
        return false;
    };

    render() {
        if (!this.props.isOpen) {
            return null;
        }

        const {
            isRTL, inspectionTypes, selectedInspectionType, qataryId, name, crNumber, id, loading, error,
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
                            disabled: loading,
                            style: inputStyle,
                        },
                        React.createElement('option', { value: '' }, '-- Select --'),
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