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
    return (
        <Modal
            isOpen={isVisible}
            isBlocking={true}
            containerClassName="customer-comments-modal"
        >
            <div style={{ padding: 24, width: 600, maxHeight: '70vh', backgroundColor: 'white', display: 'flex', flexDirection: 'column' }}>
                <Stack tokens={stackTokens}>
                    <Text variant="xLarge" block style={{ fontWeight: 600, marginBottom: 2 }}>
                        {_context.resources.getString(Constants.ADD_PREVIOUS_COMMENTS_TITLE)}
                    </Text>
                    <Separator />

                    {previousComments.length > 0 && (
                        <>
                            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '8px 0', borderBottom: '1px solid #edebe9' }}>
                                <Checkbox
                                    label={_context.resources.getString(Constants.SELECT_ALL_LABEL).replace('{0}', previousComments.length.toString())}
                                    checked={selectedCount === previousComments.length && previousComments.length > 0}
                                    indeterminate={selectedCount > 0 && selectedCount < previousComments.length}
                                    onChange={handleSelectAll}
                                />
                                <Text variant="small" style={{ color: '#605e5c' }}>
                                    {_context.resources.getString(Constants.SELECTED_COUNT_LABEL)
                                        .replace('{0}', selectedCount.toString())
                                        .replace('{1}', previousComments.length.toString())}
                                </Text>
                            </div>

                            <div style={{ maxHeight: 300, overflowY: 'auto', border: '1px solid #edebe9', borderRadius: 4, padding: 8 }}>
                                {previousComments.map((comment, index) => (
                                    <div key={comment.id} style={{ padding: 12, marginBottom: 8, border: '1px solid #edebe9', borderRadius: 4, backgroundColor: comment.isSelected ? '#f3f2f1' : 'white' }}>
                                        <div style={{ display: 'flex', alignItems: 'flex-start' }}>
                                            <Checkbox
                                                checked={comment.isSelected}
                                                onChange={(_, checked) => handleCommentSelection(comment.id, checked ?? false)}
                                                styles={{ root: { marginTop: 2, marginRight: 8 } }}
                                            />
                                            <div style={{ flex: 1, cursor: 'pointer' }} onClick={() => handleCommentSelection(comment.id, !comment.isSelected)}>
                                                <Text variant="small" style={{ color: '#605e5c', fontWeight: 600, marginBottom: 4 }}>
                                                    {_context.resources.getString(Constants.COMMENT_NUMBER_LABEL).replace('{0}', (index + 1).toString())}
                                                </Text>
                                                <Text variant="small" block style={{ lineHeight: 1.4, wordBreak: 'break-word' }}>
                                                    {comment.comment}
                                                </Text>
                                            </div>
                                        </div>
                                    </div>
                                ))}
                            </div>
                        </>
                    )}

                    {staticReplies.length > 0 && (
                        <>
                            <Separator />
                            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '8px 0', borderBottom: '1px solid #edebe9' }}>
                                <Checkbox
                                    label={_context.resources.getString(Constants.SELECT_ALL_LABEL).replace('{0}', staticReplies.length.toString())}
                                    checked={selectedRepliesCount === staticReplies.length && staticReplies.length > 0}
                                    indeterminate={selectedRepliesCount > 0 && selectedRepliesCount < staticReplies.length}
                                    onChange={handleSelectAllReplies}
                                />
                                <Text variant="small" style={{ color: '#605e5c' }}>
                                    {_context.resources.getString(Constants.SELECTED_COUNT_LABEL)
                                        .replace('{0}', selectedRepliesCount.toString())
                                        .replace('{1}', staticReplies.length.toString())}
                                </Text>
                            </div>

                            <div style={{ maxHeight: 300, overflowY: 'auto', border: '1px solid #edebe9', borderRadius: 4, padding: 8 }}>
                                {staticReplies.map((reply, index) => (
                                    <div key={reply.id} style={{ padding: 12, marginBottom: 8, border: '1px solid #edebe9', borderRadius: 4, backgroundColor: reply.isSelected ? '#f3f2f1' : 'white' }}>
                                        <div style={{ display: 'flex', alignItems: 'flex-start' }}>
                                            <Checkbox
                                                checked={reply.isSelected}
                                                onChange={(_, checked) => handleStaticReplySelection(reply.id, checked ?? false)}
                                                styles={{ root: { marginTop: 2, marginRight: 8 } }}
                                            />
                                            <div style={{ flex: 1, cursor: 'pointer' }} onClick={() => handleStaticReplySelection(reply.id, !reply.isSelected)}>
                                                <Text variant="small" style={{ color: '#605e5c', fontWeight: 600, marginBottom: 4 }}>
                                                    {(index + 1).toString()}
                                                </Text>
                                                <Text variant="small" block style={{ lineHeight: 1.4, wordBreak: 'break-word' }}>
                                                    {reply.title}
                                                </Text>
                                            </div>
                                        </div>
                                    </div>
                                ))}
                            </div>
                        </>
                    )}

                    <Separator />
                    <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8, marginTop: 16 }}>
                        <DefaultButton text={_context.resources.getString(Constants.CANCEL_BUTTON_LABEL)} onClick={()=>handleonClose(true)} />
                        <DefaultButton text={_context.resources.getString(Constants.PROCEED_WITHOUT_COMMENTS_LABEL)} onClick={() => handleConfirmSelection(true)} />
                        <PrimaryButton
                            text={_context.resources.getString(Constants.ADD_SELECTED_COMMENTS_LABEL).replace('{0}', (selectedCount + selectedRepliesCount).toString())}
                            onClick={()=>handleConfirmSelection(false)}
                            disabled={selectedCount + selectedRepliesCount === 0}
                        />
                    </div>
                </Stack>
            </div>
        </Modal>
    );
};

export default CustomerCommentsModal;
