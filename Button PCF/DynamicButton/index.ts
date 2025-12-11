/* eslint-disable */

import * as React from "react";
import * as ReactDOM from "react-dom";
import { PrimaryButton, IconButton } from "@fluentui/react";
import { IInputs, IOutputs } from "./generated/ManifestTypes";

interface IButtonState {
    isLoading: boolean;
}

class ButtonComponent extends React.Component<{
    buttonText: string;
    buttonColor: string;
    buttonIcon: string | null;
    isDisabled: boolean;
    onClick: () => Promise<void>;
}, IButtonState> {

    constructor(props: any) {
        super(props);
        this.state = {
            isLoading: false
        };
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
                    title: isLoading ? "Processing..." : iconName,
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
            // Use PrimaryButton with text - Add "..." when loading
            const displayText = isLoading ? `${buttonText}...` : buttonText;

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

    private getJavaScriptCode(): string | null {
        const field = this._context.parameters.jsCodeField;
        if (field && field.raw) {
            return field.raw as string;
        }
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

    private getButtonColor(): string {
        const field = this._context.parameters.buttonColor;
        if (field && field.raw) {
            return field.raw as string;
        }
        return "#307eafff"; // Default color
    }

    private getButtonIcon(): string | null {
        const field = this._context.parameters.buttonIcon;
        if (field && field.raw) {
            return field.raw as string;
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
    const jsCode = this.getJavaScriptCode();

    if (!jsCode || jsCode.trim() === "") {
        console.error("No JavaScript code found in the field");
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

    private renderButton(): void {
        const jsCode = this.getJavaScriptCode();
        const buttonText = this.getButtonText();
        const buttonColor = this.getButtonColor();
        const buttonIcon = this.getButtonIcon();
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
                onClick: () => this.handleButtonClick()
            }
        );

        ReactDOM.render(element, this._container);
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this._context = context;
        this.renderButton();
    }

    public getOutputs(): IOutputs {
        return {};
    }

    public destroy(): void {
        ReactDOM.unmountComponentAtNode(this._container);
    }
}