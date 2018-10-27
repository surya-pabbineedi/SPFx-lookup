import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ILookupProps {
  listName: string;
  listUrl: string;
  context: WebPartContext;
}
