import { override } from '@microsoft/decorators';
import * as React from 'react';
import {
    BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import CustomCallout from "../components/CustomCallout";
import * as ReactDOM from 'react-dom';
import { ICustomCalloutProps } from '../components/ICustomCalloutProps';

export interface IPromtsApplicationCustomizerProperties {
    data: Array<any>;
}

export default class PromtsApplicationCustomizer
    extends BaseApplicationCustomizer<IPromtsApplicationCustomizerProperties> {

    @override
    public onInit(): Promise<void> {
        debugger;
        const data: Array<any> = [
            {
                title: "Options",
                message: "Call the \"Options \" menu to access personal and application settings",
                link: "link1",
                selector: "#O365_MainLink_Settings",
                isShowed: false
            },
            {
                title: "Search",
                message: "Search this site",
                link: "link2",
                selector: "#spPageChromeAppDiv > div > div > div:nth-child(2) > div > div.ms-compositeHeader-siteAndActionsContainer > div.ms-compositeHeader-actionsContainer > div.ms-compositeHeader-searchBoxContainer > div.searchBoxContainer_b51fc60b > div > div > form > input",
                isShowed: false
            },
            {
                title: "Application Launcher",
                message: "Opening the launcher to access Office 365 applications",
                link: "link3",
                selector: "#O3165_MainLink_NavMenu",
                isShowed: false
            }
        ];
        const instanceId: string = this.componentId;
        const user: string = this.context.pageContext.user.loginName;
        this.tryToRenderPromts(data, instanceId, user);
        return Promise.resolve();
    }

    private handleClickOnDocument(event: any, promtsContainers: Array<any>): void {
        if (!this.isPromtElement(event.target)) {
            this.removePromts(promtsContainers);
        }
    }

    private removePromts(containers: Array<any>): void {
        containers.forEach(container => {
            ReactDOM.unmountComponentAtNode(container);
        });
    }

    private isPromtElement(element: any): boolean {
        let currentElement: any = element;
        while (currentElement !== document.body && !currentElement.classList.contains("ms-Layer")) {
            currentElement = currentElement.parentNode;
        }
        return currentElement !== document.body;
    }

    private renderPromt(item: any, instanceId: string, user: string): any {
        let element: any = document.body.querySelector(item.selector);
        const reactElement: React.ReactElement<ICustomCalloutProps> = React.createElement(
            CustomCallout,
            {
                title: item.title,
                message: item.message,
                link: item.link,
                selector: item.selector,
                instanceID: instanceId,
                user: user
            }
        );
        element.appendChild(document.createElement("div"));
        ReactDOM.render(reactElement, element.lastChild);
        return element.lastChild;
    }

    private tryToRenderPromts(data: Array<any>, instanceId: string, user: string): void {
        const that: any = this;
        let promtsContainers: Array<any> = [];
        let tryToRenderPromtResult: [boolean, any];
        let numberOfShowedPromts: number = 0;
        let numberOfCycles: number = 0;
        const intervalID: number = setInterval(() => {
            if (data !== undefined) {
                data.forEach(item => {
                    if (!item.isShowed) {
                        tryToRenderPromtResult = that.tryToRenderPromt(item, instanceId, user);
                        item.isShowed = tryToRenderPromtResult[0];
                        if (tryToRenderPromtResult[1] !== undefined) {
                            promtsContainers.push(tryToRenderPromtResult[1]);
                            numberOfShowedPromts++;
                        }
                    }
                }
                );
                if (numberOfShowedPromts === data.length || numberOfCycles > 200) {
                    clearInterval(intervalID);                   
                }               
            } else {
                clearInterval(intervalID);
            }
            numberOfCycles++;
        }, 250);
        document.body.addEventListener("click", event => {
            that.handleClickOnDocument(event, promtsContainers);
        });
    }

    private tryToRenderPromt(item: any, instanceId: string, user: string): [boolean, any] {
        if (document.body.querySelector(item.selector) !== null) {
            return [true, this.renderPromt(item, instanceId, user)];
        } else {
            return [false, undefined];
        }
    }

    private checkSelectors(data: Array<any>): boolean {
        let result: boolean = true;
        if (data !== undefined) {
            data.forEach(item => {
                if (document.body.querySelector(item.selector) === null) {
                    result = false;
                }
            });
        }
        return result;
    }
}
