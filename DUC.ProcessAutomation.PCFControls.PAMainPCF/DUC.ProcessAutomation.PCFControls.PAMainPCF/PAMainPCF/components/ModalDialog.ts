import * as React from 'react';
import { useEffect } from 'react';
import { IInputs } from "../generated/ManifestTypes";
import { IActionButtonProps } from './Actions';
import { CustomerComment, StaticReply, CustomerCommentsModal } from './CustomerCommentsModal';
import { CommentsModalDialog } from './CommentsModalDialog';
import { Constants, getServiceRequestLogFetchXML, getStaticRepliesFetchXML, getValue } from "../constants";
import { FormContextHelper } from "./FormContextHelper";

interface IModalDialogProps {
    _context: ComponentFramework.Context<IInputs>;
    isVisible: boolean;
    onClose: () => void;
    onSave: (action: IActionButtonProps, comment: string, files?: File[]) => Promise<void>;
    action: IActionButtonProps;
    prefilledComments?: string;
}

const ModalDialog: React.FC<IModalDialogProps> = ({
    _context,
    action,
    isVisible,
    onClose,
    onSave
}) => {
    const [showCustomerComments, setShowCustomerComments] = React.useState<boolean>(false);
    const [showComments, setShowComments] = React.useState<boolean>(false);
    const [isLoading, setIsLoading] = React.useState(false);
    const [comments, setComments] = React.useState<CustomerComment[]>([]);
    const [staticReplies, setStaticReplies] = React.useState<StaticReply[]>([]);
    const [chosenComment, setChosenComment] = React.useState<string>("");

    useEffect(() => {
        if (!isVisible) return;

        const fetchData = async (): Promise<void> => {
            let replies: StaticReply[] = [];
            let fetchedComments: CustomerComment[] = [];

            try {
                setIsLoading(true);

                // 🔹 Fetch customer comments
                if (action.sendToCustomer) {
                    const primaryKey = FormContextHelper.getStringValue("activityid") || FormContextHelper.getFormContext()?.data?.entity?.getId()?.replace(/[{}]/g, "") || "";
                    const query = getServiceRequestLogFetchXML(primaryKey);
                    const response = await _context.webAPI.retrieveMultipleRecords(Constants.SERVICE_REQUEST_LOG_ENTITY, query);

                    fetchedComments = response.entities
                        .filter(e => e[Constants.ACTION_LOG_USER_COMMENTS]?.trim())
                        .map(e => ({
                            id: e.id,
                            comment: e[Constants.ACTION_LOG_USER_COMMENTS],
                            isSelected: false,
                        }));

                    setComments(fetchedComments);
                }

                // 🔹 Fetch static replies
                if (action?.staticReplyTemplateId) {
                    const query = getStaticRepliesFetchXML(action.staticReplyTemplateId);
                    const response = await _context.webAPI.retrieveMultipleRecords(Constants.STATIC_REPLY_ENTITY, query);

                    replies = response.entities.map(e => ({
                        id: e.id,
                        title: getValue(e, _context.resources.getString(Constants.STATIC_REPLY_TITLE_COLNAME)) ?? '',
                        text: getValue(e, _context.resources.getString(Constants.STATIC_REPLY_BODY_COLNAME)) ?? '',
                        isSelected: false,
                    }));

                    setStaticReplies(replies);
                }

                // 🔹 Decide which modal to show
                if (replies.length > 0 || fetchedComments.length > 0) {
                    setShowCustomerComments(true);
                    setShowComments(false);
                } else {
                    setShowCustomerComments(false);
                    setShowComments(true);
                }

            } catch (error) {
                console.error("Error during fetch:", error);
            } finally {
                setIsLoading(false);
            }
        };

        fetchData().catch(console.error);
    }, [isVisible, action]);

    const onSelect = (selectedComment: string) => {
        setIsLoading(true);
        setChosenComment(selectedComment);
        setShowCustomerComments(false);
        setShowComments(true);
        setIsLoading(false);
    };

    const handleClose = () => {
        setShowCustomerComments(false);
        setShowComments(false);
        setChosenComment("");
        onClose();
    };

    return React.createElement(
        "div",
        null,
        isLoading && React.createElement(
            "div",
            { className: "loading-spinner" },
            React.createElement("span", null, _context.resources.getString(Constants.LOADING))
        ),
        React.createElement(CustomerCommentsModal, {
            _context: _context,
            isVisible: isVisible && showCustomerComments,
            action: action,
            staticReplies: staticReplies,
            previousComments: comments,
            onClose: handleClose,
            onProceed: onSelect,
        }),
        React.createElement(CommentsModalDialog, {
            _context: _context,
            isVisible: isVisible && showComments,
            onClose: handleClose,
            onSave: onSave,
            action: action,
            prefilledComments: chosenComment,
        })
    );
};

export default ModalDialog;
