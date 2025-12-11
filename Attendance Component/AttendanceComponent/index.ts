import * as React from "react";
import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as ReactDOM from "react-dom";
import { Main } from "./Component/Main";

export class AttendanceComponent implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    /**
     * Empty constructor.
*/  private container: HTMLDivElement;
    private _context: ComponentFramework.Context<IInputs>;
    private currentProps: IInputs;
    constructor() {
        // Empty
    }

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ) {
        const sectionElement = document.querySelector('[aria-label="General"]') as HTMLElement;

        if (sectionElement) {
            sectionElement.style.boxShadow = 'none';
            sectionElement.style.background = 'none';
        }
        this.container = container;

    }


    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this._context = context;
        this.currentProps = context.parameters;

        const element = React.createElement(Main, {
            context: this._context,
        });
        const sectionElement = document.querySelector('[aria-label="General"]') as HTMLElement;

        if (sectionElement) {
            sectionElement.style.boxShadow = 'none';
            sectionElement.style.background = 'none';
        }
        ReactDOM.render(element, this.container);
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as "bound" or "output"
     */
    public getOutputs(): IOutputs {
        return {};
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void {
        ReactDOM.unmountComponentAtNode(this.container);

        // Add code to cleanup control if necessary
    }
}
