/* eslint-disable */

import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { Main } from "./component/Main";
export class MyControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    private container: HTMLDivElement;
    private _context: ComponentFramework.Context<IInputs>;
    private currentProps: IInputs;

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

    public updateView(context: ComponentFramework.Context<IInputs>): void {

        this._context = context;
        this.currentProps = context.parameters;
        const element = React.createElement(Main, {
            context: this._context,
        });
        debugger
        const sectionElement = document.querySelector('[aria-label="General"]') as HTMLElement;

        if (sectionElement) {
            sectionElement.style.boxShadow = 'none';
            sectionElement.style.background = 'none';
        }


        ReactDOM.render(element, this.container);
    }

    public getOutputs(): IOutputs { return {}; }

    public destroy(): void {
        ReactDOM.unmountComponentAtNode(this.container);
    }
}
