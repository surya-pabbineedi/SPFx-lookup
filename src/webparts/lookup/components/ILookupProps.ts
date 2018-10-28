import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IFieldConfiguration } from '../../../../lib/webparts/lookup/components/IFieldConfiguration';

export interface ILookupProps {
  listUrl: string;
  title: string;
  fields?: Array<IFieldConfiguration>;
  description: string;
  context: WebPartContext;
}
