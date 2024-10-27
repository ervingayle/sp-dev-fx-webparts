import * as React from "react";
import ChatStreaming from './ChatStreaming';
import {IChatStreamingProps} from '../components/IChatStreamingProps'
import {IChatStreamingWebPartProps} from '../IChatStreamingWebPartProps'

export default class LazyLoad extends React.Component<IChatStreamingWebPartProps> {
    private _isDarkTheme: boolean = false;
    private _environmentMessage: string = '';
    webPartContext: any;

    constructor(wpContext, props: IChatStreamingWebPartProps) {
      super(props);
      this.webPartContext = wpContext.lazyLoadArray[0];
    }


    public render(): React.ReactElement {
      const element: React.ReactElement<IChatStreamingProps> = React.createElement(
          ChatStreaming,
          {
            openApiOptions: {
              apiKey: this.webPartContext.webPartProperties.openAiApiKey,
              endpoint: this.webPartContext.webPartProperties.openAiApiEndpoint,
              deploymentName: this.webPartContext.webPartProperties.openAiApiDeploymentName,
              version: this.webPartContext.webPartProperties.openAiApiVersion
            },
            isDarkTheme: this._isDarkTheme,
            environmentMessage: this._environmentMessage,
            hasTeamsContext: false,
            userDisplayName: this.webPartContext.webPartContext.pageContext.user.displayName,
            httpClient: this.webPartContext.webPartContext.httpClient
          }
        );
        return (
          element
        );
      }

}
