import * as React from 'react';
import { Callout, DirectionalHint } from 'office-ui-fabric-react/lib/Callout';
import { ICustomCalloutProps } from "./ICustomCalloutProps";
import { ICustomCalloutState } from "./ICustomCalloutState";
import "./styles.css";

export default class CustomCallout extends React.Component<ICustomCalloutProps, ICustomCalloutState> {
    public constructor(props) {
        super(props);
        let index: number = this.getItemIndex(this.props);
        this.state = {
            isShowed: index === -1
        };
        this.handleClickCloseButton = this.handleClickCloseButton.bind(this);
    }

    private getItemIndex(props: any): number {
        let index: number = -1;
        let storage: Array<any> = JSON.parse(localStorage.getItem("closedPromts")) || [];
        for (let i: number = 0; i < storage.length; i++) {
            if (storage[i].selector === props.selector && storage[i].user === props.user &&
                storage[i].instanceID === props.instanceID) {
                index = i;
            }
        }
        return index;
    }

    public handleClickCloseButton() {
        this.setState({ isShowed: false });
        let storage: Array<any> = JSON.parse(localStorage.getItem("closedPromts")) || [];
        storage.push({
            instanceID: this.props.instanceID,
            selector: this.props.selector,
            user: this.props.user
        });
        localStorage.setItem("closedPromts", JSON.stringify(storage));
    }

    public render(): JSX.Element {
        return (
            <div>
                {
                    this.state.isShowed && <Callout
                        target={document.body.querySelector(this.props.selector)}
                        directionalHint={DirectionalHint.bottomLeftEdge}
                        className="callout"
                    >
                        <h3>{this.props.title}</h3>
                        <p>{this.props.message}</p>
                        <p>{this.props.link}</p>
                        <button
                            onClick={this.handleClickCloseButton}
                        >
                            X
                        </button>
                    </Callout>
                }
            </div>

        );
    }
}