import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'MsTeamsHandlerWebPartStrings';
import MsTeamsHandler from './components/MsTeamsHandler';
import { IMsTeamsHandlerProps } from './components/IMsTeamsHandlerProps';
import { MSGraphClient } from '@microsoft/sp-http';

export interface IMsTeamsHandlerWebPartProps {
  TeamTitle1: string;
  client:MSGraphClient;
}

export default class MsTeamsHandlerWebPart extends BaseClientSideWebPart<IMsTeamsHandlerWebPartProps> {

  public render(): void {

    this.context.msGraphClientFactory.getClient()
    .then((grphclient: MSGraphClient): void => {
      const element: React.ReactElement<IMsTeamsHandlerProps > = React.createElement(
        MsTeamsHandler,
        {
          TeamTitle: this.properties.TeamTitle1,
        client:grphclient
        });
        ReactDom.render(element, this.domElement);
      });
      
   
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
            description: strings.TeamTitle
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('TeamTitle1', {
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
