/* eslint-disable */

import * as React from "react";
import * as ReactDOM from "react-dom";
import { PrimaryButton, IconButton } from "@fluentui/react";
import { IInputs, IOutputs } from "./generated/ManifestTypes";

// ─── Interfaces ───────────────────────────────────────────────

interface IButtonState {
    isLoading: boolean;
}

interface IActionTypeData {
    code: string | null;
    color: string | null;
    icon: string | null;
}

interface ISingleButtonProps {
    buttonText: string;
    buttonColor: string;
    buttonIcon: string | null;
    isDisabled: boolean;
    languageId: number;
    onClick: () => Promise<void>;
}

// ─── Single Button Component (same as DynamicButton) ──────────

class ButtonComponent extends React.Component<ISingleButtonProps, IButtonState> {

    constructor(props: ISingleButtonProps) {
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

// ─── Multi-Button Wrapper Component ───────────────────────────

interface IMultiButtonProps {
    buttons: ISingleButtonProps[];
}

class MultiButtonComponent extends React.Component<IMultiButtonProps> {
    render() {
        return React.createElement(
            "div",
            { className: "multi-button-container" },
            ...this.props.buttons.map((btnProps, index) =>
                React.createElement(ButtonComponent, {
                    key: index,
                    ...btnProps
                })
            )
        );
    }
}

// ─── Button Configuration ─────────────────────────────────────

interface IButtonConfig {
    actionTypeIdParam: keyof IInputs;
    textParam: keyof IInputs;
    textArabicParam: keyof IInputs;
}

const BUTTON_CONFIGS: IButtonConfig[] = [
    {
        actionTypeIdParam: "button1ActionTypeId",
        textParam: "button1Text",
        textArabicParam: "button1TextArabic",
    },
    {
        actionTypeIdParam: "button2ActionTypeId",
        textParam: "button2Text",
        textArabicParam: "button2TextArabic",
    },
    {
        actionTypeIdParam: "button3ActionTypeId",
        textParam: "button3Text",
        textArabicParam: "button3TextArabic",
    },
];

// ─── PCF Control ──────────────────────────────────────────────

export class MultiDynamicButton implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    private _context: ComponentFramework.Context<IInputs>;
    private _container: HTMLDivElement;
    private _actionTypeCache: Map<string, IActionTypeData> = new Map();

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        this._context = context;
        this._container = container;
        this.renderButtons();
    }

    // ── Helpers ────────────────────────────────────────────────

    private getParamValue(paramName: keyof IInputs): string | null {
        const field = this._context.parameters[paramName];
        if (field && field.raw) {
            return field.raw as string;
        }
        return null;
    }

    private getLanguageId(): number {
        return this._context.userSettings.languageId;
    }

    private isArabicLanguage(): boolean {
        const lang = this.getLanguageId();
        return lang === 1025 || lang === 1026 || lang === 2049 || lang === 3073 || lang === 4097 || lang === 5121;
    }

    private getRelatedRecordId(): string {
        const pageContext = (this._context as any).page;
        if (pageContext && pageContext.entityId) {
            return pageContext.entityId;
        }
        return "00000000-0000-0000-0000-000000000000";
    }

    // ── Action Type retrieval (with caching) ──────────────────

    private async getActionTypeData(actionTypeName: string): Promise<IActionTypeData | null> {
        if (!actionTypeName || actionTypeName.trim() === "") {
            return null;
        }

        // Check cache first
        if (this._actionTypeCache.has(actionTypeName)) {
            return this._actionTypeCache.get(actionTypeName)!;
        }

        try {
            console.log("MultiDynamicButton: Retrieving action type data:", actionTypeName);

            const filter = `?$select=duc_actioncommand,duc_color,duc_icon&$filter=duc_name eq '${actionTypeName}'&$top=1`;
            const result = await this._context.webAPI.retrieveMultipleRecords("duc_actiontype", filter);

            if (result && result.entities && result.entities.length > 0) {
                const entity = result.entities[0];
                const data: IActionTypeData = {
                    code: entity.duc_actioncommand || null,
                    color: entity.duc_color || null,
                    icon: entity.duc_icon || null,
                };
                console.log("MultiDynamicButton: Action type data retrieved:", data);
                this._actionTypeCache.set(actionTypeName, data);
                return data;
            } else {
                console.log("MultiDynamicButton: No action type found:", actionTypeName);
                return null;
            }
        } catch (error) {
            console.error("MultiDynamicButton: Error retrieving action type:", error);
            return null;
        }
    }

    // ── Build props for a single button ───────────────────────

    private async buildButtonProps(config: IButtonConfig): Promise<ISingleButtonProps> {
        const actionTypeName = this.getParamValue(config.actionTypeIdParam);
        const englishText = this.getParamValue(config.textParam) || "";
        const arabicText = this.getParamValue(config.textArabicParam) || "";

        // Determine display text based on language
        let buttonText: string;
        if (this.isArabicLanguage() && arabicText.trim() !== "") {
            buttonText = arabicText;
        } else {
            buttonText = englishText || arabicText || "";
        }

        // Fetch action type data (color, icon, code)
        let actionTypeData: IActionTypeData | null = null;
        if (actionTypeName) {
            actionTypeData = await this.getActionTypeData(actionTypeName);
        }

        // Resolve color: action type color → default blue
        const buttonColor = (actionTypeData && actionTypeData.color) ? actionTypeData.color : "#0078d4";

        // Resolve icon: action type icon → null
        const buttonIcon = (actionTypeData && actionTypeData.icon) ? actionTypeData.icon : null;

        // Resolve JS code
        const jsCode = (actionTypeData && actionTypeData.code && actionTypeData.code.trim() !== "") ? actionTypeData.code : null;
        const isDisabled = !jsCode;

        const languageId = this.getLanguageId();

        return {
            buttonText,
            buttonColor,
            buttonIcon,
            isDisabled,
            languageId,
            onClick: () => this.handleButtonClick(config),
        };
    }

    // ── Click handler ─────────────────────────────────────────

    private async handleButtonClick(config: IButtonConfig): Promise<void> {
        const actionTypeName = this.getParamValue(config.actionTypeIdParam);
        let jsCode: string | null = null;

        if (actionTypeName) {
            const actionTypeData = await this.getActionTypeData(actionTypeName);
            if (actionTypeData && actionTypeData.code && actionTypeData.code.trim() !== "") {
                jsCode = actionTypeData.code;
            }
        }

        if (!jsCode || jsCode.trim() === "") {
            console.error("MultiDynamicButton: No JavaScript code found for button");
            await this._context.navigation.openAlertDialog({
                text: "No JavaScript code configured for this button.",
                confirmButtonLabel: "OK"
            });
            return;
        }

        try {
            const languageId = this.getLanguageId();
            const relatedRecord = this.getRelatedRecordId();

            console.log("MultiDynamicButton: Executing JavaScript with parameters:", {
                languageId,
                relatedRecord
            });

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

            const result = await runCode(this._context, languageId, relatedRecord);
            console.log("MultiDynamicButton: JavaScript code executed successfully", result);

        } catch (error: any) {
            console.error("MultiDynamicButton: Error executing JavaScript code:", error);

            const errorMessage = error?.message || String(error);
            const stackTrace = error?.stack ? "\n\nStack: " + error.stack : "";

            await this._context.navigation.openAlertDialog({
                text: "Error executing code: " + errorMessage + stackTrace,
                confirmButtonLabel: "OK"
            });
        }
    }

    // ── Render ─────────────────────────────────────────────────

    private async renderButtons(): Promise<void> {
        // Build props for all 3 buttons in parallel
        const buttonPropsArray = await Promise.all(
            BUTTON_CONFIGS.map(config => this.buildButtonProps(config))
        );

        // Only include buttons that have an action type configured
        const activeButtons = buttonPropsArray.filter((_, index) => {
            const actionTypeName = this.getParamValue(BUTTON_CONFIGS[index].actionTypeIdParam);
            return actionTypeName && actionTypeName.trim() !== "";
        });

        if (activeButtons.length === 0) {
            // Nothing to render
            ReactDOM.unmountComponentAtNode(this._container);
            return;
        }

        const element = React.createElement(MultiButtonComponent, {
            buttons: activeButtons,
        });

        ReactDOM.render(element, this._container);
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this._context = context;
        // Reset cache on update so fresh data is fetched
        this._actionTypeCache.clear();
        this.renderButtons();
    }

    public getOutputs(): IOutputs {
        return {};
    }

    public destroy(): void {
        ReactDOM.unmountComponentAtNode(this._container);
    }
}
