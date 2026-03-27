import * as React from 'react';
import { useEffect } from 'react';
import { IInputs } from "../generated/ManifestTypes";
import { Dialog, DialogType, DialogFooter, PrimaryButton, DefaultButton, TextField, Icon } from '@fluentui/react';
import { Constants } from "../constants";
import { IActionButtonProps } from './Actions';

interface ICommentsModalDialogProps {
  _context: ComponentFramework.Context<IInputs>;
  isVisible: boolean;
  onClose: () => void;
  onSave: (action: IActionButtonProps, comment: string, files?: File[]) => Promise<void>;
  action: IActionButtonProps;
  prefilledComments?: string | undefined;
}

export const CommentsModalDialog: React.FC<ICommentsModalDialogProps> = ({
  _context,
  isVisible,
  onClose,
  onSave,
  action,
  prefilledComments = ''
}) => {
  const [comment, setComment] = React.useState<string>("");
  const [errorMessage, setErrorMessage] = React.useState<string | null>(null);
  const [selectedFiles, setSelectedFiles] = React.useState<File[]>([]);
  const [isDragOver, setIsDragOver] = React.useState<boolean>(false);

  // Handle prefilled comments
  useEffect(() => {
    console.log("=== ModalDialog useEffect triggered ===");
    console.log("isVisible:", isVisible);
    console.log("prefilledComments:", prefilledComments);
    console.log("prefilledComments length:", prefilledComments?.length);
    console.log("Current comment state:", comment);

    if (isVisible && prefilledComments) {
      setComment(prefilledComments);
    }
  }, [isVisible, prefilledComments]);
  // Clear form when modal closes

  useEffect(() => {
    if (!isVisible) {
      setComment("");
      setSelectedFiles([]);
      setErrorMessage(null);
      setIsDragOver(false);
    }
  }, [isVisible]);

  // Handle change in the text box input
  const handleCommentChange = (
    event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ): void => {
    setComment(newValue ?? "");
    if (errorMessage) setErrorMessage(null);
  };

  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    if (event.target.files) {
      handleFiles(event.target.files);
    }
  };

  const handleFiles = (files: FileList) => {
    if (files) setSelectedFiles(Array.from(files));
  };

  if (!isVisible) return null;

  return React.createElement(
    Dialog,
    {
      hidden: !isVisible,
      dialogContentProps: {
        type: DialogType.largeHeader,
        title: action.displayName,
      },
      styles: {
        main: {
          maxWidth: 'none !important',
          minWidth: '30% !important',
          bordercolor: `${action.buttonColor} !important`
        },
        root: {
          bordercolor: `${action.buttonColor} !important`
        }
      } as any,
    },
    React.createElement(
      "div",
      { style: { marginBottom: '16px' } },
      // TextField
      React.createElement(TextField, {
        multiline: true,
        rows: 16,
        value: comment,
        onChange: handleCommentChange,
        placeholder: _context.resources.getString(Constants.ENTER_COMMENTS_MSG),
        styles: {
          fieldGroup: {
            height: '200px',
          },
          field: {
            overflowY: 'auto',
            height: '100%',
          }
        }
      }),
      // File Upload Section
      React.createElement(
        "div",
        { className: "file-upload-section" },
        React.createElement(
          "label",
          { className: "file-upload-label" },
          _context.resources.getString(Constants.ATTACH_FILE_LABEL) || "Attach File (optional):"
        ),
        // Drop Zone
        React.createElement(
          "div",
          {
            className: `file-dropzone ${isDragOver ? "drag-over" : ""}`,
            onDragOver: (e: any) => {
              e.preventDefault();
              setIsDragOver(true);
            },
            onDragLeave: (e: any) => {
              e.preventDefault();
              setIsDragOver(false);
            },
            onDrop: (e: any) => {
              e.preventDefault();
              setIsDragOver(false);
              handleFiles(e.dataTransfer.files);
            },
            onClick: () => document.getElementById("fileInput")?.click(),
          },
          React.createElement(Icon, { iconName: "CloudUpload", className: "upload-icon" }),
          React.createElement(
            "div",
            { className: "upload-text" },
            "Drag and drop files here or ",
            React.createElement("span", null, "browse")
          ),
          React.createElement("input", {
            id: "fileInput",
            type: "file",
            accept: "*/*",
            multiple: true,
            onChange: handleFileChange,
            style: { display: "none" },
          })
        ),
        // File Preview List
        selectedFiles.length > 0 && React.createElement(
          "div",
          { className: "file-list" },
          selectedFiles.map((file, idx) =>
            React.createElement(
              "div",
              { key: idx, className: "file-item" },
              React.createElement(Icon, { iconName: "Attach", className: "file-icon" }),
              React.createElement("span", { className: "file-name" }, file.name),
              React.createElement("span", { className: "file-size" }, (file.size / 1024).toFixed(1) + " KB"),
              React.createElement(Icon, {
                iconName: "Cancel",
                title: "Remove",
                className: "file-remove",
                onClick: (e: any) => {
                  e.stopPropagation();
                  const updatedFiles = selectedFiles.filter((_, i) => i !== idx);
                  setSelectedFiles(updatedFiles);
                }
              })
            )
          )
        )
      ),
      // Error Message
      errorMessage && React.createElement(
        "div",
        { style: { color: 'red', marginTop: '8px' } },
        errorMessage
      )
    ),
    // Modal Footer Buttons
    React.createElement("div", { style: { height: '0', borderTop: `4px solid ${action.buttonColor} !important` } }),
    React.createElement(
      DialogFooter,
      null,
      React.createElement(PrimaryButton, {
        onClick: () => {
          if (!comment.trim()) {
            setErrorMessage(_context.resources.getString(Constants.SPECIFY_COMMENT_MSG));
            return;
          }
          setErrorMessage(null);
          void onSave(action, comment, selectedFiles);
        },
        text: action.displayName,
        style: {
          backgroundColor: action.buttonColor
        }
      }),
      React.createElement(DefaultButton, {
        onClick: onClose,
        text: _context.resources.getString(Constants.CLOSE_MODAL)
      })
    )
  );
};

export default CommentsModalDialog;
