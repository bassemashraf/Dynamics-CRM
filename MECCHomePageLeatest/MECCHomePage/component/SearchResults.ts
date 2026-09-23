/* eslint-disable */
import * as React from 'react';
import { IInputs } from "../generated/ManifestTypes";

interface ISearchResult {
    accountid: string;
    name: string;
    emailaddress1: string;
    telephone1: string;
}

interface ISearchResultsProps {
    results: ISearchResult[];
    onRecordClick: (accountId: string) => void;
    onClose: () => void;
    context: ComponentFramework.Context<IInputs>;
}

interface LocalizedStrings {
    SearchResults: string;
    AccountName: string;
    Email: string;
    Phone: string;
    Close: string;
}

export const SearchResults = (props: ISearchResultsProps) => {
    // Load localized strings
    const strings: LocalizedStrings = React.useMemo(() => {
        const ctx = props.context;
        return {
            SearchResults: ctx.resources.getString("SearchResults"),
            AccountName: ctx.resources.getString("AccountName"),
            Email: ctx.resources.getString("Email"),
            Phone: ctx.resources.getString("Phone"),
            Close: ctx.resources.getString("Close")
        };
    }, [props.context]);

    // Detect RTL
    const isRTL = React.useMemo(() => {
        const rtlLanguages = [1025, 1037, 1054, 1056, 1065, 1068, 1069, 1101, 1114, 1119];
        return rtlLanguages.includes(props.context.userSettings.languageId);
    }, [props.context.userSettings.languageId]);

    const handleMouseEnter = (e: React.MouseEvent<HTMLDivElement>): void => {
        e.currentTarget.style.backgroundColor = '#f3f2f1';
    };

    const handleMouseLeave = (e: React.MouseEvent<HTMLDivElement>): void => {
        e.currentTarget.style.backgroundColor = 'white';
    };

    const isLandscape = (): boolean => {
        return window.innerWidth > window.innerHeight;
    };

    const { results, onRecordClick, onClose } = props;
    const landscape = isLandscape();

    // Format title with count
    const title = strings.SearchResults.replace("{0}", results.length.toString());

    return React.createElement(
        'div',
        {
            style: {
                position: 'fixed',
                top: 0,
                left: 0,
                right: 0,
                bottom: 0,
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
                    borderRadius: landscape ? '2px' : '0',
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
                        onClick: onClose,
                        'aria-label': strings.Close
                    },
                    '×'
                )
            ),
            // Content wrapper
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
                // Column Headers
                React.createElement(
                    'div',
                    {
                        style: {
                            display: 'flex',
                            padding: '12px 20px',
                            backgroundColor: '#faf9f8',
                            borderBottom: '1px solid #edebe9',
                            fontWeight: 600,
                            fontSize: '14px',
                            color: '#323130',
                            minWidth: '600px',
                        }
                    },
                    React.createElement(
                        'div',
                        { style: { flex: '1 1 0', minWidth: '120px' } },
                        strings.AccountName
                    ),
                    React.createElement(
                        'div',
                        { style: { flex: '1 1 0', minWidth: '150px' } },
                        strings.Email
                    ),
                    React.createElement(
                        'div',
                        { style: { flex: '1 1 0', minWidth: '120px' } },
                        strings.Phone
                    )
                ),
                // Scrollable Results
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
                                    transition: 'background-color 0.1s',
                                    backgroundColor: 'white',
                                    minWidth: '600px',
                                },
                                onClick: () => onRecordClick(record.accountid),
                                onMouseEnter: handleMouseEnter,
                                onMouseLeave: handleMouseLeave
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
};