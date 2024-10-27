import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import { PropertyFieldPassword } from '@pnp/spfx-property-controls/lib/PropertyFieldPassword';

import * as strings from 'ChatStreamingWebPartStrings';

import { Providers, customElementHelper } from "@microsoft/mgt-element";
import { SharePointProvider } from "@microsoft/mgt-sharepoint-provider";
import { lazyLoadComponent } from "@microsoft/mgt-spfx-utils";
import {IChatStreamingWebPartProps} from './IChatStreamingWebPartProps'

// Async import of component that imports the React Components
const MgtReact = React.lazy(
  () =>
    import(/* webpackChunkName: 'mgt-react-component' */ "./components/LazyLoad")
);

// set the disambiguation before initializing any web part
customElementHelper.withDisambiguation("react-azure-openai-stream");

export default class ChatStreamingWebPart extends BaseClientSideWebPart<IChatStreamingWebPartProps> {

  private _environmentMessage: string = '';

    protected async onInit(): Promise<void> {
    if (!Providers.globalProvider) {
      Providers.globalProvider = new SharePointProvider(this.context);
      console.log(Providers.globalProvider);
    }
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
      console.log(this._environmentMessage);
    });
  }

  public render(): void {
    const element = lazyLoadComponent(MgtReact, {
      lazyLoadArray: [{webPartContext: this.context,
      webPartProperties: this.properties}]
    });
    ReactDom.render(element, this.domElement);
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
                PropertyPaneTextField('openAiApiEndpoint', {
                  label: strings.OpenAiApiEndpointFieldLabel
                }),
                PropertyPaneTextField('openAiApiDeploymentName', {
                  label: strings.OpenAiApiDeploymentFieldLabel
                }),
                PropertyFieldPassword("openAiApiKey", {
                  key: "openAiApiKey",
                  label: strings.OpenAiApiKeyFieldLabel,
                  value: this.properties.openAiApiKey
                }),
                PropertyPaneTextField('openAiApiVersion', {
                  label: strings.OpenAiApiVersionFieldLabel
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
