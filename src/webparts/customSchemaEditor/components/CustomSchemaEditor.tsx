import * as React from 'react';
import styles from './CustomSchemaEditor.module.scss';
import { ICustomSchemaEditorProps } from './ICustomSchemaEditorProps';
import { escape } from '@microsoft/sp-lodash-subset';
import {
  Logger,
  LogLevel
} from "@pnp/logging";
import { IGroup } from '../../../shared/interfaces';

import {
  SearchBox,
  DetailsList,
  IColumn
} from 'office-ui-fabric-react';

const LOG_SOURCE: string = 'CustomSchemaEditor';

export interface ICustomSchemaEditorState {
  groups: IGroup[];
}

export default class CustomSchemaEditor extends React.Component<ICustomSchemaEditorProps,ICustomSchemaEditorState> {

  public constructor(props: ICustomSchemaEditorProps){
    super(props);
    this.state = {
      groups : new Array<IGroup>()
    };
  }

  public async componentDidMount(){
    Logger.write(`[${LOG_SOURCE}] componentDidMount();`);
    try {
      let groups = await this.props.groupService.getGroupsByName("spo");
      this.setState({
        groups: groups
      });
    } catch (error) {
      Logger.writeJSON(error,LogLevel.Error);
    }
  }

  public render(): React.ReactElement<ICustomSchemaEditorProps> {
    Logger.write(`[${LOG_SOURCE}] render();`);

    const columns : IColumn[] = [
      {
        name: "Name",
        fieldName: "displayName",
        key: "displayName",
        minWidth: 300
      }
    ];

    return (
      <div className={ styles.customSchemaEditor }>
        <SearchBox
          placeholder="Find Group"
          onSearch={searchString => {
            this.getGroupsByName(searchString);
          }}
        />

        <DetailsList
          items={this.state.groups}
          columns={columns}
        />
      </div>
    );
  }

  private async getGroupsByName(searchString: string) : Promise<void> {
    Logger.write(`[${LOG_SOURCE}] getGroupsByName(${searchString});`);

    this.setState({
      groups: await this.props.groupService.getGroupsByName(searchString)
    })

  }


}
