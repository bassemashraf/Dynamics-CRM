import * as React from 'react';
import { useEffect } from 'react';
import { IInputs } from "../generated/ManifestTypes";
import { Dialog, DialogType, DialogFooter, PrimaryButton, DefaultButton, TextField, Icon } from '@fluentui/react';
import { Constants } from "../constants";
import { IActionButtonProps } from './Actions';
import { DialogTitle } from '@fluentui/react-components';

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

  return (
    <Dialog
      hidden={!isVisible}
      dialogContentProps={{
        type: DialogType.largeHeader,
        title: action.displayName,
      }}
      styles={{
        main: {
          maxWidth: 'none !important',  // Force removing max-width
          minWidth: '30% !important',
          bordercolor: `${action.buttonColor} !important`
        },
        root: {
          bordercolor: `${action.buttonColor} !important`
        }
      }}
    >
      <div style={{ marginBottom: '16px' }}>
        <TextField
          multiline
          rows={16}
          value={comment}
          onChange={handleCommentChange}
          placeholder={_context.resources.getString(Constants.ENTER_COMMENTS_MSG)}
          styles={{
            fieldGroup: {
              height: '200px', // You can adjust the height as needed
            },
            field: {
              overflowY: 'auto',
              height: '100%',
            }
          }}
        />

        {/* === File Upload Section === */}
        <div className="file-upload-section">
          <label className="file-upload-label">
            {_context.resources.getString(Constants.ATTACH_FILE_LABEL) ||
              "Attach File (optional):"}
          </label>

          {/* Drop Zone */}
          <div
            className={`file-dropzone ${isDragOver ? "drag-over" : ""}`}
            onDragOver={(e) => {
              e.preventDefault();
              setIsDragOver(true);
            }}
            onDragLeave={(e) => {
              e.preventDefault();
              setIsDragOver(false);
            }}
            onDrop={(e) => {
              e.preventDefault();
              setIsDragOver(false);
              handleFiles(e.dataTransfer.files); // ✅ safe, no 'any'
            }}
            onClick={() => document.getElementById("fileInput")?.click()}
          >
            <Icon iconName="CloudUpload" className="upload-icon" />
            <div className="upload-text">
              Drag and drop files here or <span>browse</span>
            </div>
            <input
              id="fileInput"
              type="file"
              accept="*/*"
              multiple
              onChange={handleFileChange}
              style={{ display: "none" }}
            />
          </div>
          {/* File Preview List */}
          {selectedFiles.length > 0 && (
            <div className="file-list">
              {selectedFiles.map((file, idx) => (
                <div key={idx} className="file-item">
                  <Icon iconName="Attach" className="file-icon" />
                  <span className="file-name">{file.name}</span>
                  <span className="file-size">{(file.size / 1024).toFixed(1)} KB</span>
                  <Icon
                    iconName="Cancel"
                    title="Remove"
                    className="file-remove"
                    onClick={(e) => {
                      e.stopPropagation();
                      const updatedFiles = selectedFiles.filter((_, i) => i !== idx);
                      setSelectedFiles(updatedFiles);
                    }}
                  />
                </div>
              ))}
            </div>
          )}
        </div>

        {/* Error Message */}
        {errorMessage && (
          <div style={{ color: 'red', marginTop: '8px' }}>
            {errorMessage}
          </div>
        )}

      </div>

      {/* === Modal Footer Buttons === */}
      <DialogTitle style={{ borderTop: `4px solid ${action.buttonColor} !important` }}></DialogTitle>
      <DialogFooter>
        <PrimaryButton
          onClick={() => {
            if (!comment.trim()) {
              setErrorMessage(_context.resources.getString(Constants.SPECIFY_COMMENT_MSG));
              return;
            }
            setErrorMessage(null); // Clear previous error
            void onSave(action, comment, selectedFiles);// 👈 explicitly discard the returned Promise
          }}                      // Use `void` to ignore the Promise returned
          text={action.displayName}
          style={{
            backgroundColor: action.buttonColor
          }}
        />
        <DefaultButton onClick={onClose} text={_context.resources.getString(Constants.CLOSE_MODAL)} />
      </DialogFooter>
    </Dialog>
  );
};

export default CommentsModalDialog;
