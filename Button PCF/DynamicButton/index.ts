/* eslint-disable */

import * as React from "react";
import * as ReactDOM from "react-dom";
import { PrimaryButton, IconButton } from "@fluentui/react";
import { IInputs, IOutputs } from "./generated/ManifestTypes";

interface IButtonState {
    isLoading: boolean;
}

interface IActionTypeData {
    code: string | null;
    color: string | null;
    icon: string | null;
}

class ButtonComponent extends React.Component<{
    buttonText: string;
    buttonColor: string;
    buttonIcon: string | null;
    isDisabled: boolean;
    languageId: number;
    onClick: () => Promise<void>;
}, IButtonState> {

    constructor(props: any) {
        super(props);
        this.state = {
            isLoading: false
        };
    }

    private getLoadingMessage(): string {
        const languageId = this.props.languageId;
        const translations: { [key: number]: string } = {
            1025: "جاري العمل ...", 
            1033: "Processing...",     
        };

        return translations[languageId] || translations[1033];
    }

    private handleClick = async (): Promise<void> => {
        // Prevent double clicks
        if (this.state.isLoading) {
            console.log("Button already processing, ignoring click");
            return;
        }

        console.log("Button clicked, setting loading state");
        this.setState({ isLoading: true });

        try {
            await this.props.onClick();
            console.log("Button action completed");
        } catch (error) {
            console.error("Button action failed:", error);
        } finally {
            this.setState({ isLoading: false });
            console.log("Loading state cleared");
        }
    };

    render() {
        const { buttonText, buttonColor, buttonIcon, isDisabled } = this.props;
        const { isLoading } = this.state;

        console.log("Rendering button, isLoading:", isLoading);

        // If no text provided, use IconButton instead
        if (!buttonText || buttonText.trim() === "") {
            const iconName = buttonIcon || "ButtonControl";

            return React.createElement(
                IconButton,
                {
                    className: "Jsaction",
                    iconProps: { iconName: iconName },
                    onClick: this.handleClick,
                    title: isLoading ? this.getLoadingMessage() : iconName,
                    ariaLabel: iconName,
                    disabled: isDisabled || isLoading,
                    style: {
                        width: "40px",
                        height: "40px",
                        opacity: isLoading ? 0.6 : 1
                    }
                }
            );
        } else {
            // Use PrimaryButton with text - Add loading message when loading
            const displayText = isLoading ? this.getLoadingMessage() : buttonText;

            return React.createElement(
                PrimaryButton,
                {
                    className: "Jsaction",
                    text: displayText,
                    iconProps: buttonIcon ? { iconName: buttonIcon } : undefined,
                    onClick: this.handleClick,
                    disabled: isDisabled || isLoading,
                    style: {
                        backgroundColor: buttonColor,
                        border: `1px solid ${buttonColor}`,
                        borderColor: buttonColor,
                        opacity: isLoading ? 0.8 : 1,
                        cursor: isLoading ? "wait" : "pointer"
                    }
                }
            );
        }
    }
}

export class DynamicButton implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    private _context: ComponentFramework.Context<IInputs>;
    private _container: HTMLDivElement;
    private _buttonComponentRef: ButtonComponent | null = null;
    private _actionTypeData: IActionTypeData | null = null;

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        this._context = context;
        this._container = container;

        // Render the React component
        this.renderButton();
    }

    private getJavaScriptCodeFromField(): string | null {
        const field = this._context.parameters.jsCodeField;
        if (field && field.raw) {
            return field.raw as string;
        }
        return null;
    }

    private getActionTypeName(): string | null {
        const field = this._context.parameters.actionTypeId;
        if (field && field.raw) {
            return field.raw as string;
        }
        return null;
    }

    private async getActionTypeData(): Promise<IActionTypeData | null> {
        const actionTypeName = this.getActionTypeName();

        if (!actionTypeName || actionTypeName.trim() === "") {
            console.log("No action type name provided");
            return null;
        }

        try {
            console.log("Retrieving action type data:", actionTypeName);

            // Use OData filter query to get all needed fields
            const filter = `?$select=duc_actioncommand,duc_color,duc_icon&$filter=duc_name eq '${actionTypeName}'&$top=1`;

            const result = await this._context.webAPI.retrieveMultipleRecords(
                "duc_actiontype",
                filter
            );

            if (result && result.entities && result.entities.length > 0) {
                const entity = result.entities[0];
                
                const actionTypeData: IActionTypeData = {
                    code: entity.duc_actioncommand || null,
                    color: entity.duc_color || null,
                    icon: entity.duc_icon || null
                };

                console.log("Action type data retrieved successfully:", actionTypeData);
                return actionTypeData;
            } else {
                console.log("No action type found with name:", actionTypeName);
                return null;
            }

        } catch (error) {
            console.error("Error retrieving action type record:", error);
            return null;
        }
    }

    private async getJavaScriptCode(): Promise<string | null> {
        // Priority 1: Check jsCodeField
        const fieldCode = this.getJavaScriptCodeFromField();
        if (fieldCode && fieldCode.trim() !== "") {
            console.log("Using JavaScript code from field");
            return fieldCode;
        }

        // Priority 2: Check action type lookup
        if (!this._actionTypeData) {
            this._actionTypeData = await this.getActionTypeData();
        }

        if (this._actionTypeData && this._actionTypeData.code && this._actionTypeData.code.trim() !== "") {
            console.log("Using JavaScript code from action type");
            return this._actionTypeData.code;
        }

        console.log("No JavaScript code found in either field or action type");
        return null;
    }

    private getButtonTextEnglish(): string {
        const field = this._context.parameters.buttonText;
        if (field && field.raw) {
            return field.raw as string;
        }
        return "";
    }

    private getButtonTextArabic(): string {
        const field = this._context.parameters.buttonTextArabic;
        if (field && field.raw) {
            return field.raw as string;
        }
        return "";
    }

    private getButtonText(): string {
        const languageId = this.getLanguageId();
        const arabicLanguageId = 1025;

        const isArabic = languageId === arabicLanguageId ||
            languageId === 1026 ||
            languageId === 2049 ||
            languageId === 3073 ||
            languageId === 4097 ||
            languageId === 5121;

        const arabicText = this.getButtonTextArabic();
        const englishText = this.getButtonTextEnglish();

        // Return Arabic text if language is Arabic and Arabic text is available
        if (isArabic && arabicText && arabicText.trim() !== "") {
            return arabicText;
        }

        // Otherwise return English text (or Arabic as fallback if English is empty)
        return englishText || arabicText || "";
    }


    private async getButtonColor(): Promise<string> {
        // Priority 1: Check parameter field
        const field = this._context.parameters.buttonColor;
        if (field && field.raw) {
            return field.raw as string;
        }

        // Priority 2: Check action type
        if (!this._actionTypeData) {
            this._actionTypeData = await this.getActionTypeData();
        }

        if (this._actionTypeData && this._actionTypeData.color) {
            console.log("Using button color from action type:", this._actionTypeData.color);
            return this._actionTypeData.color;
        }

        // Default color
        return "#0078d4"; // Default Fluent UI blue
    }

    private async getButtonIcon(): Promise<string | null> {
        // Priority 1: Check parameter field
        const field = this._context.parameters.buttonIcon;
        if (field && field.raw) {
            return field.raw as string;
        }

        // Priority 2: Check action type
        if (!this._actionTypeData) {
            this._actionTypeData = await this.getActionTypeData();
        }

        if (this._actionTypeData && this._actionTypeData.icon) {
            console.log("Using button icon from action type:", this._actionTypeData.icon);
            return this._actionTypeData.icon;
        }

        return null;
    }

    private getLanguageId(): number {
        return this._context.userSettings.languageId;
    }

    private getRelatedRecordId(): string {
        const pageContext = (this._context as any).page;
        if (pageContext && pageContext.entityId) {
            return pageContext.entityId;
        }
        return "00000000-0000-0000-0000-000000000000";
    }

    private async handleButtonClick(): Promise<void> {
        const jsCode = await this.getJavaScriptCode();

        if (!jsCode || jsCode.trim() === "") {
            console.error("No JavaScript code found in field or action type");
            await this._context.navigation.openAlertDialog({
                text: "No JavaScript code configured for this button.",
                confirmButtonLabel: "OK"
            });
            return;
        }

        try {
            const languageId = this.getLanguageId();
            const relatedRecord = this.getRelatedRecordId();

            console.log("Executing JavaScript with parameters:", {
                languageId,
                relatedRecord
            });

            // Use string concatenation to avoid template literal conflicts
            const promiseWrapper =
                "return (async () => {" +
                "    try {" +
                "        " + jsCode +
                "    } catch (innerError) {" +
                "        console.error('Script execution error:', innerError);" +
                "        throw innerError;" +
                "    }" +
                "})();";

            const runCode = new Function(
                "pcfContext",
                "languageId",
                "relatedRecord",
                promiseWrapper
            );

            // Execute and wait for completion
            const result = await runCode(this._context, languageId, relatedRecord);

            console.log("JavaScript code executed successfully", result);

        } catch (error: any) {
            console.error("Error executing JavaScript code:", error);

            const errorMessage = error?.message || String(error);
            const stackTrace = error?.stack ? "\n\nStack: " + error.stack : "";

            await this._context.navigation.openAlertDialog({
                text: "Error executing code: " + errorMessage + stackTrace,
                confirmButtonLabel: "OK"
            });
        }
    }

    private async renderButton(): Promise<void> {
        const buttonText = this.getButtonText();
        const buttonColor = await this.getButtonColor();
        const buttonIcon = await this.getButtonIcon();
        const languageId = this.getLanguageId();

        // Check if we have any JavaScript code available
        const jsCode = await this.getJavaScriptCode();
        const isDisabled = !jsCode || jsCode.trim() === "";

        const element = React.createElement(
            ButtonComponent,
            {
                ref: (ref: ButtonComponent | null) => {
                    this._buttonComponentRef = ref;
                },
                buttonText: buttonText,
                buttonColor: buttonColor,
                buttonIcon: buttonIcon,
                isDisabled: isDisabled,
                languageId: languageId,
                onClick: () => this.handleButtonClick()
            }
        );

        ReactDOM.render(element, this._container);
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this._context = context;
        // Reset cached action type data on update
        this._actionTypeData = null;
        this.renderButton();
    }

    public getOutputs(): IOutputs {
        return {};
    }

    public destroy(): void {
        ReactDOM.unmountComponentAtNode(this._container);
    }
}