import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import { Main, IMainProps } from "./components/MainControl";
import * as ReactDOM from 'react-dom';
import { initializeIcons } from '@fluentui/react/lib/Icons';

// Disable default CDN font loading to prevent offline crashes
initializeIcons('', { disableWarnings: true });

export class PAMainOfflinePCF implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private notifyOutputChanged: () => void;
    private container: HTMLDivElement;
    private currentProps: IInputs;
    private _context: ComponentFramework.Context<IInputs>;
    /**
     * Empty constructor.
     */
    constructor() {
        // Empty
    }

    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     */
    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        try {
            this.container = container;
            this.notifyOutputChanged = notifyOutputChanged;
            this.currentProps = context.parameters;
            this._context = context;
        } catch (e: any) {
            const errMsg = `[PAMainOfflinePCF.init] Error: ${e?.message || String(e)}`;
            console.error(errMsg, e);
            if (typeof alert === "function") alert(errMsg);
        }
    }

    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     * @returns ReactElement root react element for the control
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void {
        try {
            this.currentProps = context.parameters;
            this._context = context;
            // Get the dataset from the context
            const props: IMainProps = {
                _context: context,
                onnotifyOutputChanged: this.notifyOutputChanged
            };
            ReactDOM.render(React.createElement(Main, props), this.container);
        } catch (e: any) {
            const errMsg = `[PAMainOfflinePCF.updateView] Error: ${e?.message || String(e)}`;
            console.error(errMsg, e);
            if (typeof alert === "function") alert(errMsg);
        }
    }

    public getOutputs(): IOutputs {
        try {
            return {};
        } catch (e: any) {
            const errMsg = `[PAMainOfflinePCF.getOutputs] Error: ${e?.message || String(e)}`;
            console.error(errMsg, e);
            if (typeof alert === "function") alert(errMsg);
            return {};
        }
    }


    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void {
        ReactDOM.unmountComponentAtNode(this.container);
    }
}
