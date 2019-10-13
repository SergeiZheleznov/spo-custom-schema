import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, PropertyPaneCheckbox } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { MSGraphClient } from '@microsoft/sp-http';
import * as strings from 'CustomSchemaEditorWebPartStrings';
import CustomSchemaEditor from './components/CustomSchemaEditor';
import { ICustomSchemaEditorProps } from './components/ICustomSchemaEditorProps';
import { IGroupService, GroupService, ICustomSchemaService, CustomSchemaService } from '../../shared/services/';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

import {
  Logger,
  ConsoleListener,
  LogLevel
} from "@pnp/logging";

const LOG_SOURCE: string = 'CustomSchemaEditorWebPart';

export interface ICustomSchemaEditorWebPartProps {
  customSchemaId: string;
  lockCustomSchema: boolean;
}

export default class CustomSchemaEditorWebPart extends BaseClientSideWebPart<ICustomSchemaEditorWebPartProps> {

  private graphClient: MSGraphClient;
  private groupService: IGroupService;
  private customSchemaService: ICustomSchemaService;

  private customSchemaErrorMessage: string = null;
  private customSchema: MicrosoftGraph.SchemaExtension;

  public async onInit(): Promise<void> {

    Logger.subscribe(new ConsoleListener());
    Logger.activeLogLevel = LogLevel.Info;
    Logger.write(`[${LOG_SOURCE}] onInit();`);

    try {
      Logger.write(`[${LOG_SOURCE}] Retrieving of Graph Client`);
      this.graphClient = await this.context.msGraphClientFactory.getClient();
      this.groupService = new GroupService(this.graphClient);
      this.customSchemaService = new CustomSchemaService(this.graphClient);

      if (this.properties.customSchemaId) {
        Logger.write(`[${LOG_SOURCE}] trying to get custom schema`);
        this.customSchema = await this.customSchemaService.get(this.properties.customSchemaId);
        if (this.customSchema) {
          this.properties.lockCustomSchema = true;
        } else {
          this.properties.lockCustomSchema = false;
        }
        Logger.writeJSON(this.customSchema);
      } else {
        this.properties.lockCustomSchema = false;
      }

    } catch (error) {
      Logger.writeJSON(error,LogLevel.Error);
    }
  }

  public render(): void {
    const element: React.ReactElement<ICustomSchemaEditorProps> = React.createElement(
      CustomSchemaEditor,
      {
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

  protected onAfterPropertyPaneChangesApplied(){
    Logger.write(`[${LOG_SOURCE}] onAfterPropertyPaneChangesApplied`);
  }

  protected async onPropertyPaneConfigurationComplete(){
    Logger.write(`[${LOG_SOURCE}] onPropertyPaneConfigurationComplete`);
  }

  protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any) {
    Logger.write(`[${LOG_SOURCE}] onPropertyPaneFieldChanged(${propertyPath}, ${oldValue}, ${newValue})`);
    switch (propertyPath) {
      case "lockCustomSchema":
        await this.lockCustomSchemaPropertyHanfler(oldValue, newValue);
        break;
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupName: "Custom Schema",
              groupFields: [
                PropertyPaneTextField('customSchemaId', {
                  label: "ID",
                  disabled: this.customSchema ? true : false,
                  errorMessage: this.customSchemaErrorMessage,
                }),
                PropertyPaneCheckbox('lockCustomSchema',{
                  text: "Lock schema",
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private async lockCustomSchemaPropertyHanfler(oldValue: any, newValue: any) {
    Logger.write(`[${LOG_SOURCE}] lockCustomSchemaPropertyHanfler()`);
    switch (newValue) {
      case true:
        this.customSchema = await this.customSchemaService.get(this.properties.customSchemaId);
        break;
      case false:
        this.customSchema = null;
        break;
    }
  }
}
