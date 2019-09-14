import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { MSGraphClient } from '@microsoft/sp-http';
import * as strings from 'CustomSchemaEditorWebPartStrings';
import CustomSchemaEditor from './components/CustomSchemaEditor';
import { ICustomSchemaEditorProps } from './components/ICustomSchemaEditorProps';
import { IGroupService, GroupService } from '../../shared/services/';

import {
  Logger,
  ConsoleListener,
  LogLevel
} from "@pnp/logging";
const LOG_SOURCE: string = 'CustomSchemaEditorWebPart';

export interface ICustomSchemaEditorWebPartProps {
  description: string;
}

export default class CustomSchemaEditorWebPart extends BaseClientSideWebPart<ICustomSchemaEditorWebPartProps> {

  private graphClient: MSGraphClient;
  private groupService: IGroupService;

  public async onInit(): Promise<void> {

    Logger.subscribe(new ConsoleListener());
    Logger.activeLogLevel = LogLevel.Info;
    Logger.write(`[${LOG_SOURCE}] onInit();`);

    try {
      Logger.write(`[${LOG_SOURCE}] Retrieving of Graph Client`);
      this.graphClient = await this.context.msGraphClientFactory.getClient();
      this.groupService = new GroupService(this.graphClient);
    } catch (error) {
      Logger.writeJSON(error,LogLevel.Error);
    }
  }

  public render(): void {
    const element: React.ReactElement<ICustomSchemaEditorProps > = React.createElement(
      CustomSchemaEditor,
      {
        description: this.properties.description,
        groupService: this.groupService
      }
    );

    ReactDom.render(element, this.domElement);
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
