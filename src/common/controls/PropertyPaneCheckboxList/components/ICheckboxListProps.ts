import { IFieldConfiguration } from '../../../../../lib/webparts/lookup/components/IFieldConfiguration';

export interface ICheckboxListProps {
  label: string;
  loadOptions: () => Promise<IFieldConfiguration[]>;
  onChanged: (option: IFieldConfiguration, checked: boolean) => void;
  stateKey: string;
}