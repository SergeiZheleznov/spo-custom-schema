import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

export interface ICustomSchemaService {
  create();
  get(customSchemaId: string): Promise<MicrosoftGraph.SchemaExtension>;
}
