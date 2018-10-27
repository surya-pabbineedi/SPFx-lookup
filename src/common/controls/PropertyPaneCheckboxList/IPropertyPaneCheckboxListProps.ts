import { IFieldConfiguration } from '../../../../lib/webparts/lookup/components/IFieldConfiguration';

export interface IPropertyPaneCheckboxListProps {
  label: string;
  loadOptions: () => Promise<IFieldConfiguration[]>;
  onChanged: (field: IFieldConfiguration, checked: boolean) => void;
}
