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

  constructor(graphClient: MSGraphClient){
    this.graphClient = graphClient;
  }

  public async getGroupsByName(searchStr: string = ""): Promise<MicrosoftGraph.Group[]> {
    Logger.write(`[${LOG_SOURCE}] getGroups();`);
    try {
      const response = await this.graphClient
      .api('/groups')
      .version('v1.0')
      .filter(searchStr ? `startswith(mailNickname,'${searchStr}')` : null)
      .get();

      let groups = new Array<IGroup>();
      response.value.forEach( (o365Group : MicrosoftGraph.Group) => {
        groups.push(this.toGroup(o365Group));
      });

      Logger.writeJSON(groups);

      return groups;
    } catch (error) {
      Logger.writeJSON(error,LogLevel.Error);
    }
  }

  private toGroup(o365Group: MicrosoftGraph.Group) : IGroup {
    return {
      id: o365Group.id,
      displayName: o365Group.displayName,
      mailNickname: o365Group.mailNickname
    } as IGroup;
  }
}
