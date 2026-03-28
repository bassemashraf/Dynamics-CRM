import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import * as ReactDOM from 'react-dom';
import { MainControl, IMainProps } from "./MainControl";

export class NewPAOfflinePCF implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private notifyOutputChanged: () => void;
    private container: HTMLDivElement;
    private currentProps: IInputs;
    private _context: ComponentFramework.Context<IInputs>;

    constructor() {}

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        this.container = container;
        this.notifyOutputChanged = notifyOutputChanged;
        this.currentProps = context.parameters;
        this._context = context;
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this.currentProps = context.parameters;
        this._context = context;
        
        const props: IMainProps = {
            _context: context,
            notifyOutputChanged: this.notifyOutputChanged
        };
        
        ReactDOM.render(React.createElement(MainControl, props), this.container);
    }

    public getOutputs(): IOutputs {
        return {};
    }

    public destroy(): void {
        ReactDOM.unmountComponentAtNode(this.container);
    }
}
