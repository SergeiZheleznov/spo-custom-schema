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
import { IGroupService, GroupService } from '../../shared/services/';

import {
  Logger,
  ConsoleListener,
  LogLevel
} from "@pnp/logging";
import { ICustomSchema } from '../../shared/interfaces';
const LOG_SOURCE: string = 'CustomSchemaEditorWebPart';

export interface ICustomSchemaEditorWebPartProps {
  customSchemaId: string;
  lockCustomSchema: boolean;
}

export default class CustomSchemaEditorWebPart extends BaseClientSideWebPart<ICustomSchemaEditorWebPartProps> {

  private graphClient: MSGraphClient;
  private groupService: IGroupService;

  private customSchemaErrorMessage: string = null;
  private customSchema: ICustomSchema = null;

  public async onInit(): Promise<void> {

    Logger.subscribe(new ConsoleListener());
    Logger.activeLogLevel = LogLevel.Info;
    Logger.write(`[${LOG_SOURCE}] onInit();`);

    try {
      Logger.write(`[${LOG_SOURCE}] Retrieving of Graph Client`);
      this.graphClient = await this.context.msGraphClientFactory.getClient();
      this.groupService = new GroupService(this.graphClient);

      if (this.properties.customSchemaId) {
        this.customSchema = await this.getCustomSchema(this.properties.customSchemaId);
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

  private async getCustomSchema(customSchemaId: string): Promise<ICustomSchema> {
    Logger.write(`[${LOG_SOURCE}] getCustomSchema('${this.properties.customSchemaId}');`);
    this.customSchemaErrorMessage = null;

    try {

      const response = await this.graphClient
        .api('/schemaExtensions')
        .version("v1.0")
        .filter(`id eq '${customSchemaId}'`)
        .get();

      if (response.value && response.value.length > 0) {
        const customSchema = response.value[0] as microsoftgraph.SchemaExtension;

        return {
          id: customSchema.id
        } as ICustomSchema;

      } else {
        this.customSchemaErrorMessage = `Custom Schema not exists ${this.properties.customSchemaId}`;
        return null;
      }

    } catch (error) {
      Logger.writeJSON(error,LogLevel.Error);
    }
  }

  protected onAfterPropertyPaneChangesApplied(){
    Logger.write(`[${LOG_SOURCE}] onAfterPropertyPaneChangesApplied`);
  }

  protected async onPropertyPaneConfigurationComplete(){
    Logger.write(`[${LOG_SOURCE}] onPropertyPaneConfigurationComplete`);
    if (this.properties.lockCustomSchema) {
      this.customSchema = await this.getCustomSchema(this.properties.customSchemaId);

    } else {
      this.customSchema = null;
    }


  }
  protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): Promise<void> {
    //super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
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
}
