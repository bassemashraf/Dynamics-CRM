/* eslint-disable */
import * as React from 'react';

interface ISearchResult {
    accountid: string;
    name: string;
    emailaddress1: string;
    telephone1: string;
    new_licensenumber?: string;
    new_businessregistrationnumber?: string;
}

interface IAdvancedSearchProps {
    context: ComponentFramework.Context<any>;
    onClose?: () => void;
    isOpen?: boolean;
    onSearchResults?: (results: ISearchResult[]) => void;
}

interface IAdvancedSearchState {
    valAccName: string;
    valRegNo: string;
    valBuildingNo: string;
    valPinNo: string;
    valLicenseNo: string;
    isRTL: boolean;
    showResults: boolean;
    results: ISearchResult[];
}

interface LocalizedStrings {
    AdvancedSearch: string;
    FacilityName: string;
    LicenseNumber: string;
    BuildingNumber: string;
    PINNumber: string;
    Search: string;
    SearchResults: string;
    AccountName: string;
    Email: string;
    Phone: string;
    Close: string;
    PleaseEnterSearchText: string;
    NoResultsFound: string;
    SearchError: string;
}

export class AdvancedSearch extends React.Component<IAdvancedSearchProps, IAdvancedSearchState> {
    private strings: LocalizedStrings;

    constructor(props: IAdvancedSearchProps) {
        super(props);
        const userSettings = (props.context as any).userSettings;
        const rtlLanguages = [1025, 1037, 1054, 1056, 1065, 1068, 1069, 1101, 1114, 1119];
        const isRTL = rtlLanguages.includes(userSettings?.languageId);

        // Load localized strings
        this.strings = {
            AdvancedSearch: props.context.resources.getString("AdvancedSearch"),
            FacilityName: props.context.resources.getString("FacilityName"),
            LicenseNumber: props.context.resources.getString("LicenseNumber"),
            BuildingNumber: props.context.resources.getString("BuildingNumber"),
            PINNumber: props.context.resources.getString("PINNumber"),
            Search: props.context.resources.getString("Search"),
            SearchResults: props.context.resources.getString("SearchResults"),
            AccountName: props.context.resources.getString("AccountName"),
            Email: props.context.resources.getString("Email"),
            Phone: props.context.resources.getString("Phone"),
            Close: props.context.resources.getString("Close"),
            PleaseEnterSearchText: props.context.resources.getString("PleaseEnterSearchText"),
            NoResultsFound: props.context.resources.getString("NoResultsFound"),
            SearchError: props.context.resources.getString("SearchError")
        };

        this.state = {
            valAccName: '',
            valRegNo: '',
            valBuildingNo: '',
            valPinNo: '',
            valLicenseNo: '',
            isRTL: isRTL,
            showResults: false,
            results: []
        };
    }

    private handleInputChange = (field: keyof IAdvancedSearchState, value: string): void => {
        this.setState({ [field]: value } as any);
    };

    private submit = async (): Promise<void> => {
        const {
            valAccName, valRegNo, valBuildingNo, valPinNo, valLicenseNo
        } = this.state;

        if (!valAccName && !valRegNo && !valBuildingNo && !valPinNo && !valLicenseNo) {
            return;
        }

        let fetch = "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>";
        fetch += "<entity name='account'>";
        fetch += "<attribute name='name' />";
        fetch += "<attribute name='accountid' />";
        fetch += "<attribute name='new_businessregistrationnumber' />";
        fetch += "<order attribute='name' descending='false' />";
        fetch += "<filter type='and'>";

        if (valAccName) {
            fetch += "<filter type='or'>";
            fetch += `<condition attribute='name' operator='like' value='%${valAccName}%' />`;
            fetch += "</filter>";
        }
        if (valBuildingNo) {
            fetch += `<condition attribute='address1_line1' operator='eq' value='${valBuildingNo}' />`;
        }
        if (valPinNo) {
            fetch += `<condition attribute='duc_address1_pinno' operator='eq' value='${valPinNo}' />`;
        }

        fetch += "</filter>";

        if (valLicenseNo) {
            fetch += "<link-entity name='new_license' from='new_licenseid' to='new_licenseid' link-type='inner' alias='lic'>";
            fetch += "<filter type='and'>";
            fetch += `<condition attribute='new_name' operator='like' value='%${valLicenseNo}%' />`;
            fetch += "</filter>";
            fetch += "</link-entity>";
        }

        fetch += "</entity>";
        fetch += "</fetch>";

        try {
            const result = await (this.props.context.webAPI as any).retrieveMultipleRecords(
                'account',
                `?fetchXml=${encodeURIComponent(fetch)}`
            );

            if (result.entities.length === 0) {
            } else if (result.entities.length === 1) {
                this.openRecord(result.entities[0].accountid);
            } else {
                this.setState({ results: result.entities, showResults: true });
            }
        } catch (error: any) {
            console.error('Search error:', error);
        }
    };

    private openRecord = (accountId: string): void => {
        const navigationOptions = {
            pageType: 'entityrecord' as any,
            entityName: 'account',
            entityId: accountId,
            target: 1
        };
        (this.props.context.navigation as any).navigateTo(navigationOptions);
    };

    private handleRecordClick = (accountId: string): void => {
        this.openRecord(accountId);
    };

    private handleCloseResults = (): void => {
        this.setState({ showResults: false, results: [] });
    };

    public render(): React.ReactNode {
        const { isRTL, showResults, results } = this.state;
        const { isOpen } = this.props;

        if (!isOpen) {
            return null;
        }

        if (showResults) {
            return this.renderSearchResults();
        }

        return React.createElement(
            'div',
            {
                style: {
                    position: 'fixed',
                    top: 0,
                    left: 0,
                    right: 0,
                    bottom: 0,
                    backgroundColor: 'rgba(0, 0, 0, 0.4)',
                    display: 'flex',
                    justifyContent: 'center',
                    alignItems: 'center',
                    zIndex: 2000,
                    direction: isRTL ? 'rtl' : 'ltr'
                }
            },
            React.createElement(
                'div',
                {
                    style: {
                        backgroundColor: 'white',
                        borderRadius: '0',
                        boxShadow: '0 2px 8px rgba(0, 0, 0, 0.15)',
                        width: '100%',
                        maxWidth: '100%',
                        height: '100%',
                        maxHeight: '100%',
                        display: 'flex',
                        flexDirection: 'column',
                        overflow: 'hidden'
                    }
                },
                // Header with close button
                React.createElement(
                    'div',
                    {
                        style: {
                            display: 'flex',
                            justifyContent: 'space-between',
                            alignItems: 'center',
                            padding: '16px 20px',
                            borderBottom: '1px solid #edebe9',
                        }
                    },
                    React.createElement(
                        'label',
                        {
                            style: {
                                fontSize: '20px',
                                fontWeight: 600,
                                color: '#323130',
                                margin: 0
                            }
                        },
                        this.strings.AdvancedSearch
                    ),
                    React.createElement(
                        'button',
                        {
                            style: {
                                background: 'none',
                                border: 'none',
                                fontSize: '28px',
                                color: '#605e5c',
                                cursor: 'pointer',
                                padding: '0 8px',
                                lineHeight: 1,
                            },
                            onClick: this.props.onClose,
                            'aria-label': this.strings.Close
                        },
                        '×'
                    )
                ),
                // Scrollable content
                React.createElement(
                    'div',
                    {
                        style: {
                            flex: 1,
                            overflowY: 'auto',
                            padding: '20px'
                        }
                    },
                    // Row 1
                    this.renderRow(
                        this.strings.FacilityName,
                        'valAccName',
                        this.strings.LicenseNumber,
                        'valLicenseNo'
                    ),
                    // Row 2
                    this.renderRow(
                        this.strings.BuildingNumber,
                        'valBuildingNo',
                        this.strings.PINNumber,
                        'valPinNo'
                    ),
                    // Submit Button
                    React.createElement(
                        'div',
                        {
                            style: {
                                display: 'flex',
                                marginTop: '32px',
                                gap: '12px',
                                direction: isRTL ? 'rtl' : 'ltr'
                            }
                        },
                        React.createElement(
                            'button',
                            {
                                style: {
                                    width: '200px',
                                    padding: '10px 20px',
                                    backgroundColor: '#0078d4',
                                    color: 'white',
                                    border: 'none',
                                    borderRadius: '2px',
                                    fontSize: '14px',
                                    fontWeight: 600,
                                    cursor: 'pointer'
                                },
                                onClick: this.submit
                            },
                            this.strings.Search
                        )
                    )
                )
            )
        );
    }

    private renderRow(
        label1: string, 
        field1: keyof IAdvancedSearchState,
        label2?: string, 
        field2?: keyof IAdvancedSearchState
    ): React.ReactElement {
        const { isRTL } = this.state;
        const inputs = [this.renderInput(label1, field1)];
        
        if (label2 && field2) {
            inputs.push(this.renderInput(label2, field2));
        }
        
        return React.createElement(
            'div',
            {
                style: {
                    display: 'flex',
                    gap: '16px',
                    marginBottom: '16px',
                    flexWrap: 'wrap',
                    direction: isRTL ? 'rtl' : 'ltr'
                }
            },
            ...inputs
        );
    }

    private renderInput(label: string, field: keyof IAdvancedSearchState): React.ReactElement {
        const { isRTL } = this.state;
        return React.createElement(
            'div',
            {
                style: {
                    flex: '1',
                    minWidth: '250px',
                    textAlign: isRTL ? 'right' : 'left'
                }
            },
            React.createElement(
                'label',
                {
                    style: {
                        display: 'block',
                        marginBottom: '4px',
                        fontSize: '14px',
                        fontWeight: 600,
                        color: '#323130'
                    }
                },
                label
            ),
            React.createElement('input', {
                type: 'text',
                className: 'form-control',
                style: {
                    width: '100%',
                    padding: '8px 12px',
                    fontSize: '14px',
                    border: '1px solid #8a8886',
                    borderRadius: '2px',
                    boxSizing: 'border-box',
                    direction: isRTL ? 'rtl' : 'ltr'
                },
                value: this.state[field] as string,
                onChange: (e: React.ChangeEvent<HTMLInputElement>) =>
                    this.handleInputChange(field, e.target.value)
            })
        );
    }

    private renderSearchResults(): React.ReactElement {
        const { results, isRTL } = this.state;
        const isLandscape = window.innerWidth > window.innerHeight;
        const title = this.strings.SearchResults.replace("{0}", results.length.toString());

        return React.createElement(
            'div',
            {
                style: {
                    position: 'fixed',
                    top: 0,
                    left: 0,
                    right: 0,
                    bottom: 0,
                    backgroundColor: 'rgba(0, 0, 0, 0.4)',
                    display: 'flex',
                    justifyContent: 'center',
                    alignItems: 'center',
                    zIndex: 1000,
                    direction: isRTL ? 'rtl' : 'ltr'
                }
            },
            React.createElement(
                'div',
                {
                    style: {
                        backgroundColor: 'white',
                        borderRadius: isLandscape ? '2px' : '0',
                        boxShadow: '0 2px 8px rgba(0, 0, 0, 0.15)',
                        width: '100%',
                        maxWidth: '100%',
                        height: '100%',
                        maxHeight: '100%',
                        display: 'flex',
                        flexDirection: 'column',
                    }
                },
                // Header
                React.createElement(
                    'div',
                    {
                        style: {
                            display: 'flex',
                            justifyContent: 'space-between',
                            alignItems: 'center',
                            padding: '16px 20px',
                            borderBottom: '1px solid #edebe9',
                        }
                    },
                    React.createElement(
                        'h2',
                        {
                            style: {
                                margin: 0,
                                fontSize: '20px',
                                fontWeight: 600,
                                color: '#323130',
                            }
                        },
                        title
                    ),
                    React.createElement(
                        'button',
                        {
                            style: {
                                background: 'none',
                                border: 'none',
                                fontSize: '28px',
                                color: '#605e5c',
                                cursor: 'pointer',
                                padding: '0 8px',
                                lineHeight: 1,
                            },
                            onClick: this.handleCloseResults,
                            'aria-label': this.strings.Close
                        },
                        '×'
                    )
                ),
                // Content
                React.createElement(
                    'div',
                    {
                        style: {
                            flex: 1,
                            overflow: 'hidden',
                            display: 'flex',
                            flexDirection: 'column',
                        }
                    },
                    // Column Headers with scroll wrapper
                    React.createElement(
                        'div',
                        {
                            style: {
                                overflowX: 'auto',
                                borderBottom: '1px solid #edebe9',
                            }
                        },
                        React.createElement(
                            'div',
                            {
                                style: {
                                    display: 'flex',
                                    padding: '12px 20px',
                                    backgroundColor: '#faf9f8',
                                    fontWeight: 600,
                                    fontSize: '14px',
                                    color: '#323130',
                                    minWidth: '600px',
                                }
                            },
                            React.createElement(
                                'div',
                                { style: { flex: '1 1 0', minWidth: '120px' } },
                                this.strings.AccountName
                            ),
                            React.createElement(
                                'div',
                                { style: { flex: '1 1 0', minWidth: '150px' } },
                                this.strings.Email
                            ),
                            React.createElement(
                                'div',
                                { style: { flex: '1 1 0', minWidth: '120px' } },
                                this.strings.Phone
                            )
                        )
                    ),
                    // Results
                    React.createElement(
                        'div',
                        {
                            style: {
                                flex: 1,
                                overflowY: 'auto',
                                overflowX: 'auto',
                            }
                        },
                        results.map((record: ISearchResult) =>
                            React.createElement(
                                'div',
                                {
                                    key: record.accountid,
                                    style: {
                                        display: 'flex',
                                        padding: '12px 20px',
                                        borderBottom: '1px solid #edebe9',
                                        cursor: 'pointer',
                                        backgroundColor: 'white',
                                        minWidth: '600px',
                                    },
                                    onClick: () => this.handleRecordClick(record.accountid),
                                    onMouseEnter: (e: React.MouseEvent<HTMLDivElement>) => {
                                        e.currentTarget.style.backgroundColor = '#f3f2f1';
                                    },
                                    onMouseLeave: (e: React.MouseEvent<HTMLDivElement>) => {
                                        e.currentTarget.style.backgroundColor = 'white';
                                    }
                                },
                                React.createElement(
                                    'div',
                                    {
                                        style: {
                                            flex: '1 1 0',
                                            minWidth: '120px',
                                            fontSize: '14px',
                                            color: '#0078d4',
                                            overflow: 'hidden',
                                            textOverflow: 'ellipsis',
                                            whiteSpace: 'nowrap',
                                        }
                                    },
                                    record.name
                                ),
                                React.createElement(
                                    'div',
                                    {
                                        style: {
                                            flex: '1 1 0',
                                            minWidth: '150px',
                                            fontSize: '14px',
                                            color: '#605e5c',
                                            overflow: 'hidden',
                                            textOverflow: 'ellipsis',
                                            whiteSpace: 'nowrap',
                                        }
                                    },
                                    record.emailaddress1
                                ),
                                React.createElement(
                                    'div',
                                    {
                                        style: {
                                            flex: '1 1 0',
                                            minWidth: '120px',
                                            fontSize: '14px',
                                            color: '#605e5c',
                                            overflow: 'hidden',
                                            textOverflow: 'ellipsis',
                                            whiteSpace: 'nowrap',
                                        }
                                    },
                                    record.telephone1
                                )
                            )
                        )
                    )
                )
            )
        );
    }
}