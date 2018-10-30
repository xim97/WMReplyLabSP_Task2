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
        const data: Array<any> = [
            {
                title: "Options",
                message: "Call the \"Options \" menu to access personal and application settings",
                link: "link1",
                selector: "#O365_MainLink_Settings"
            },
            {
                title: "Search",
                message: "Search this site",
                link: "link2",
                selector: "#spPageChromeAppDiv > div > div > div:nth-child(2) > div > div.ms-compositeHeader-siteAndActionsContainer > div.ms-compositeHeader-actionsContainer > div.ms-compositeHeader-searchBoxContainer > div.searchBoxContainer_b51fc60b > div > div > form > input"
            },
            {
                title: "Application Launcher",
                message: "Opening the launcher to access Office 365 applications",
                link: "link3",
                selector: "#O365_MainLink_NavMenu"
            }
        ];
        const instanceId: string = this.componentId;
        const user: string = this.context.pageContext.user.loginName;        
        this.tryToRenderPromts(data, instanceId, user);
        return Promise.resolve();
    }

    private handleClickOnDocument(event): void {
        if (!this.isPromtElement(event.target)) {
            this.removePromts();
        }
    }

    private removePromts(): void {
        let promts: NodeListOf<Element> = document.body.querySelectorAll(".ms-Layer.ms-Layer--fixed");
        for (let i: number = 0; i < promts.length; i++) {
            this.removeChild(promts[i]);
        }
    }

    private removeChild(rootNode: Element): void {
        let element: any = rootNode;
        while (element.lastChild) {
            element.removeChild(element.lastChild);
        }
    }

    private isPromtElement(element: any): boolean {
        let currentElement: any = element;
        while (currentElement !== document.body && !currentElement.classList.contains("ms-Layer")) {
            currentElement = currentElement.parentNode;
        }
        return currentElement !== document.body;
    }

    private renderPromt(item: any, instanceId: string, user: string): void {
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
    }

    private tryToRenderPromts(data: Array<any>, instanceId: string, user: string): void {
        const that: any = this;
        const intervalID: number = setInterval(() => {
            if (data !== undefined && that.checkSelectors(data)) {
                data.forEach(item => {
                    that.renderPromt(item, instanceId, user);
                }
                );
                document.body.addEventListener("click", event => {
                    that.handleClickOnDocument(event);
                });
                clearInterval(intervalID);
            }
        }, 250);
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
