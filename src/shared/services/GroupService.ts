import { IGroupService } from "./";
import { MSGraphClient } from "@microsoft/sp-http";
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import {
  Logger,
  LogLevel
} from "@pnp/logging";
import { IGroup } from "../interfaces";
const LOG_SOURCE: string = 'GroupService';

export class GroupService implements IGroupService {

  private graphClient: MSGraphClient;
  private customSchema: MicrosoftGraph.SchemaExtension;
  private errorMessage: string;

  constructor(graphClient: MSGraphClient){
    this.graphClient = graphClient;
  }

  private get propsToSelect() : string[] {
    let groupProperties = ['id','displayName','mailNickname'];
    if (this.customSchema) {

    }
    return groupProperties;
  }

  public atatchCustomSchema(customSchema: MicrosoftGraph.SchemaExtension): GroupService {
    this.customSchema = customSchema;
    return this;
  }

  public async getGroups(searchStr: string = ""): Promise<MicrosoftGraph.Group[]> {
    Logger.write(`[${LOG_SOURCE}] getGroups();`);
    try {
      let request = await this.graphClient
      .api('/groups')
      .version('v1.0')
      .select(this.propsToSelect);

      if (searchStr){
        request.filter(`startswith(mailNickname,'${searchStr}')`);
      }

      const response = await request.get();

      let groups = new Array<IGroup>();
      response.value.forEach( (o365Group : MicrosoftGraph.Group) => {
        groups.push(this.convertO365GroupToGroup(o365Group));
      });
      return groups;

    } catch (error) {
      Logger.writeJSON(error,LogLevel.Error);
    }
  }

  private convertO365GroupToGroup(o365Group: MicrosoftGraph.Group) : IGroup {
    return {
      id: o365Group.id,
      displayName: o365Group.displayName,
      mailNickname: o365Group.mailNickname
    } as IGroup;
  }
}
