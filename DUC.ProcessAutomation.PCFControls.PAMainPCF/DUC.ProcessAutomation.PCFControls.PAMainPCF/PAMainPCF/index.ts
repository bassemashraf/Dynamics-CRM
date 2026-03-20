import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { HelloWorld, IHelloWorldProps } from "./HelloWorld";
import * as React from "react";
import { Main, IMainProps } from "./components/MainControl";
import * as ReactDOM from 'react-dom';

export class PAMainPCFOffline implements ComponentFramework.StandardControl<IInputs, IOutputs> {
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
            const errMsg = `[PAMainPCFOffline.init] Error: ${e?.message || JSON.stringify(e)}`;
            console.error(errMsg, e);
            alert(errMsg);
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
            const errMsg = `[PAMainPCFOffline.updateView] Error: ${e?.message || JSON.stringify(e)}`;
            console.error(errMsg, e);
            alert(errMsg);
        }
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as "bound" or "output"
     */
    public getOutputs(): IOutputs {
        try {
            return {
                lookupField: this._context?.parameters?.lookupField?.raw,
                commentField: this._context?.parameters?.commentField?.raw ?? "",
                ShowSurveyForm: this._context?.parameters?.ShowSurveyForm?.raw ?? false,
                attachmentsGuid: this._context?.parameters?.attachmentsGuid?.raw ?? "",
                nextAssignee: this._context?.parameters?.nextAssignee?.raw ?? "",
                OwnerField: this._context?.parameters?.OwnerField?.raw
            };
        } catch (e: any) {
            const errMsg = `[PAMainPCFOffline.getOutputs] Error: ${e?.message || JSON.stringify(e)}`;
            console.error(errMsg, e);
            alert(errMsg);
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
