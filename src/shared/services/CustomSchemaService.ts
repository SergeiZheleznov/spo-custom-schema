import { MSGraphClient } from "@microsoft/sp-http";
import { ICustomSchemaService } from "./ICustomSchemaService";
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import {
  Logger,
  LogLevel
} from "@pnp/logging";

const LOG_SOURCE: string = 'CustomSchemaService';

export class CustomSchemaService implements ICustomSchemaService {
  private graphClient: MSGraphClient;

  constructor(graphClient: MSGraphClient){
    this.graphClient = graphClient;
  }

  public async create(){
    Logger.write(`[${LOG_SOURCE}] create();`);
    const request: MicrosoftGraph.SchemaExtension = {
      "id":"zx0test",
      "description": "Graph Learn training courses extensions",
      "targetTypes": [
        "Group"
      ],
      "properties": [
        {
          "name": "courseId",
          "type": "Integer"
        },
        {
          "name": "courseName",
          "type": "String"
        },
        {
          "name": "courseType",
          "type": "String"
        }
      ]
    };

    try {
      Logger.write(`[${LOG_SOURCE}] trying to create Schema Extension`);
      const response = await this.graphClient.api('/schemaExtensions').post(request);
      console.log(response);
    } catch (error) {
      Logger.writeJSON(error,LogLevel.Error);
    }


  }
}
