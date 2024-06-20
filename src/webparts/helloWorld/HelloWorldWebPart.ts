import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
 IPropertyPaneConfiguration,
 PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { sp } from "@pnp/sp";
import { Logger, LogLevel } from "@pnp/logging";

import * as strings from 'HelloWorldWebPartStrings';
import HelloWorld from './components/HelloWorld';
import { IHelloWorldProps } from './components/IHelloWorldProps';

export interface IHelloWorldWebPartProps {
 description: string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

 public render(): void {
    const element: React.ReactElement<IHelloWorldProps> = React.createElement(
      HelloWorld,
      {
        siteurl: this.context.pageContext.site.absoluteUrl,
        UserName: "Divya"
      }
    );

    ReactDom.render(element, this.domElement);
 }

 protected onInit(): Promise<void> {
    // Setup PnPjs
    sp.setup({
      sp: {
        baseUrl: "https://3c3tsp.sharepoint.com/sites/demosite/siteone/training11",
        headers: {
          Accept: "application/json;odata=verbose",
          // "X-RequestDigest": "your_request_digest_value", // This is typically handled by PnPjs
        },
      },
    });

    // Enable logging
    // Logger.subscribe(new ConsoleListener());
    Logger.activeLogLevel = LogLevel.Verbose;

    return this._getEnvironmentMessage().then(message => {
      //this._environmentMessage = message;
    });
 }

 private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
 }

 protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
 }

 protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
 }

 protected get dataVersion(): Version {
    return Version.parse('1.0');
 }

 protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                 label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
 }
}
