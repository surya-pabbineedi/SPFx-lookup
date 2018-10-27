import { IFieldConfiguration } from '../../../../../lib/webparts/lookup/components/IFieldConfiguration';

export interface ICheckboxListState {
  loading: boolean;
  options: IFieldConfiguration[];
  error: string;
}