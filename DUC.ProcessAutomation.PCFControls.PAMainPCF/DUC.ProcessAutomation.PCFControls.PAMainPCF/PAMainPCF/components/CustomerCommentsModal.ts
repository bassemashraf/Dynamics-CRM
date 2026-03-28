import * as React from "react";
import ModalDialog from "./ModalDialog";
import { 
    Modal, 
    PrimaryButton, 
    DefaultButton, 
    Checkbox, 
    Stack, 
    Text, 
    Separator,
    IStackTokens
} from '@fluentui/react';
import { IActionButtonProps } from "./Actions";
import { IInputs } from "../generated/ManifestTypes";
import { Constants } from "../constants";

const stackTokens: IStackTokens = { childrenGap: 10 };

export interface CustomerComment {
    id: string;
    comment: string;
    isSelected: boolean;
}

export interface StaticReply {
    id: string;
    title: string;
    text: string;
    isSelected: boolean;
}

export interface ICustomerCommentsModalProps {
    _context: ComponentFramework.Context<IInputs>;
    action: IActionButtonProps;
    isVisible: boolean;
    onClose: () => void;
    onProceed: (chosenComment: string) => void;
    staticReplies?: StaticReply[];
    previousComments?: CustomerComment[];
}

export const CustomerCommentsModal: React.FC<ICustomerCommentsModalProps> = ({
    _context,
    action,
    isVisible,
    onClose,
    onProceed,
    staticReplies = [],
    previousComments = []
}) => {

    const [selectedCount, setSelectedCount] = React.useState(0);
    const [selectedRepliesCount, setSelectedRepliesCount] = React.useState(0);

    // ✅ Skip CustomerCommentsModal if nothing to show
    if (!isVisible) return null;
    
    // Select all / individual selection handlers
    const handleSelectAll = () => {
        const allSelected = selectedCount === previousComments.length && previousComments.length > 0;
        previousComments.forEach(c => c.isSelected = !allSelected);
        setSelectedCount(!allSelected ? previousComments.length : 0);
    };

    const handleSelectAllReplies = () => {
        const allSelected = selectedRepliesCount === staticReplies.length && staticReplies.length > 0;
        staticReplies.forEach(r => r.isSelected = !allSelected);
        setSelectedRepliesCount(!allSelected ? staticReplies.length : 0);
    };

    const handleCommentSelection = (id: string, isChecked: boolean) => {
        previousComments.forEach(c => {
            if (c.id === id) c.isSelected = isChecked;
        });
        setSelectedCount(previousComments.filter(c => c.isSelected).length);
    };

    const handleStaticReplySelection = (id: string, isChecked: boolean) => {
        staticReplies.forEach(r => {
            if (r.id === id) r.isSelected = isChecked;
        });
        setSelectedRepliesCount(staticReplies.filter(r => r.isSelected).length);
    };

    const handleConfirmSelection = (proceedwithout=false) => {
        let chosen = '';
        if(!proceedwithout){
            const selectedCommentsArray = previousComments.filter(c => c.isSelected);
            const selectedRepliesArray = staticReplies.filter(r => r.isSelected);

            chosen = [
                ...selectedCommentsArray.map(c => c.comment),
                ...selectedRepliesArray.map(r => r.text)
            ].join('\n\n');
        }
        handleonClose();
        onProceed(chosen);
    };
    const handleonClose = (forceClose=false)=>{
        // Reset selections
        setSelectedCount(0);
        setSelectedRepliesCount(0);
        if(forceClose)
            onClose();
    }

    // --- UI ---
    return React.createElement(
        Modal,
        {
            isOpen: isVisible,
            isBlocking: true,
            containerClassName: "customer-comments-modal",
        },
        React.createElement(
            "div",
            { style: { padding: 24, width: 600, maxHeight: '70vh', backgroundColor: 'white', display: 'flex', flexDirection: 'column' } },
            React.createElement(
                Stack,
                { tokens: stackTokens },
                // Title
                React.createElement(
                    Text,
                    { variant: "xLarge", block: true, style: { fontWeight: 600, marginBottom: 2 } } as any,
                    _context.resources.getString(Constants.ADD_PREVIOUS_COMMENTS_TITLE)
                ),
                React.createElement(Separator, null),

                // Previous comments section
                previousComments.length > 0 && React.createElement(
                    React.Fragment,
                    null,
                    // Select all header
                    React.createElement(
                        "div",
                        { style: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '8px 0', borderBottom: '1px solid #edebe9' } },
                        React.createElement(Checkbox, {
                            label: _context.resources.getString(Constants.SELECT_ALL_LABEL).replace('{0}', previousComments.length.toString()),
                            checked: selectedCount === previousComments.length && previousComments.length > 0,
                            indeterminate: selectedCount > 0 && selectedCount < previousComments.length,
                            onChange: handleSelectAll
                        }),
                        React.createElement(
                            Text,
                            { variant: "small", style: { color: '#605e5c' } } as any,
                            _context.resources.getString(Constants.SELECTED_COUNT_LABEL)
                                .replace('{0}', selectedCount.toString())
                                .replace('{1}', previousComments.length.toString())
                        )
                    ),
                    // Comments list
                    React.createElement(
                        "div",
                        { style: { maxHeight: 300, overflowY: 'auto', border: '1px solid #edebe9', borderRadius: 4, padding: 8 } },
                        previousComments.map((comment, index) =>
                            React.createElement(
                                "div",
                                { key: comment.id, style: { padding: 12, marginBottom: 8, border: '1px solid #edebe9', borderRadius: 4, backgroundColor: comment.isSelected ? '#f3f2f1' : 'white' } },
                                React.createElement(
                                    "div",
                                    { style: { display: 'flex', alignItems: 'flex-start' } },
                                    React.createElement(Checkbox, {
                                        checked: comment.isSelected,
                                        onChange: (_: any, checked: any) => handleCommentSelection(comment.id, checked ?? false),
                                        styles: { root: { marginTop: 2, marginRight: 8 } }
                                    }),
                                    React.createElement(
                                        "div",
                                        { style: { flex: 1, cursor: 'pointer' }, onClick: () => handleCommentSelection(comment.id, !comment.isSelected) },
                                        React.createElement(
                                            Text,
                                            { variant: "small", style: { color: '#605e5c', fontWeight: 600, marginBottom: 4 } } as any,
                                            _context.resources.getString(Constants.COMMENT_NUMBER_LABEL).replace('{0}', (index + 1).toString())
                                        ),
                                        React.createElement(
                                            Text,
                                            { variant: "small", block: true, style: { lineHeight: 1.4, wordBreak: 'break-word' } } as any,
                                            comment.comment
                                        )
                                    )
                                )
                            )
                        )
                    )
                ),

                // Static replies section
                staticReplies.length > 0 && React.createElement(
                    React.Fragment,
                    null,
                    React.createElement(Separator, null),
                    // Select all header for replies
                    React.createElement(
                        "div",
                        { style: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '8px 0', borderBottom: '1px solid #edebe9' } },
                        React.createElement(Checkbox, {
                            label: _context.resources.getString(Constants.SELECT_ALL_LABEL).replace('{0}', staticReplies.length.toString()),
                            checked: selectedRepliesCount === staticReplies.length && staticReplies.length > 0,
                            indeterminate: selectedRepliesCount > 0 && selectedRepliesCount < staticReplies.length,
                            onChange: handleSelectAllReplies
                        }),
                        React.createElement(
                            Text,
                            { variant: "small", style: { color: '#605e5c' } } as any,
                            _context.resources.getString(Constants.SELECTED_COUNT_LABEL)
                                .replace('{0}', selectedRepliesCount.toString())
                                .replace('{1}', staticReplies.length.toString())
                        )
                    ),
                    // Replies list
                    React.createElement(
                        "div",
                        { style: { maxHeight: 300, overflowY: 'auto', border: '1px solid #edebe9', borderRadius: 4, padding: 8 } },
                        staticReplies.map((reply, index) =>
                            React.createElement(
                                "div",
                                { key: reply.id, style: { padding: 12, marginBottom: 8, border: '1px solid #edebe9', borderRadius: 4, backgroundColor: reply.isSelected ? '#f3f2f1' : 'white' } },
                                React.createElement(
                                    "div",
                                    { style: { display: 'flex', alignItems: 'flex-start' } },
                                    React.createElement(Checkbox, {
                                        checked: reply.isSelected,
                                        onChange: (_: any, checked: any) => handleStaticReplySelection(reply.id, checked ?? false),
                                        styles: { root: { marginTop: 2, marginRight: 8 } }
                                    }),
                                    React.createElement(
                                        "div",
                                        { style: { flex: 1, cursor: 'pointer' }, onClick: () => handleStaticReplySelection(reply.id, !reply.isSelected) },
                                        React.createElement(
                                            Text,
                                            { variant: "small", style: { color: '#605e5c', fontWeight: 600, marginBottom: 4 } } as any,
                                            (index + 1).toString()
                                        ),
                                        React.createElement(
                                            Text,
                                            { variant: "small", block: true, style: { lineHeight: 1.4, wordBreak: 'break-word' } } as any,
                                            reply.title
                                        )
                                    )
                                )
                            )
                        )
                    )
                ),

                React.createElement(Separator, null),
                // Footer buttons
                React.createElement(
                    "div",
                    { style: { display: 'flex', justifyContent: 'flex-end', gap: 8, marginTop: 16 } },
                    React.createElement(DefaultButton, {
                        text: _context.resources.getString(Constants.CANCEL_BUTTON_LABEL),
                        onClick: () => handleonClose(true)
                    }),
                    React.createElement(DefaultButton, {
                        text: _context.resources.getString(Constants.PROCEED_WITHOUT_COMMENTS_LABEL),
                        onClick: () => handleConfirmSelection(true)
                    }),
                    React.createElement(PrimaryButton, {
                        text: _context.resources.getString(Constants.ADD_SELECTED_COMMENTS_LABEL).replace('{0}', (selectedCount + selectedRepliesCount).toString()),
                        onClick: () => handleConfirmSelection(false),
                        disabled: selectedCount + selectedRepliesCount === 0
                    })
                )
            )
        )
    );
};

export default CustomerCommentsModal;
