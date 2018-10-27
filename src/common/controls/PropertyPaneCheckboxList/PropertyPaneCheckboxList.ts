import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  IPropertyPaneField,
  PropertyPaneFieldType
} from '@microsoft/sp-webpart-base';
import { IPropertyPaneCheckboxListProps } from './IPropertyPaneCheckboxListProps';
import { IPropertyPaneCheckboxListInternalProps } from './IPropertyPaneCheckboxListInternalProps';
import { ICheckboxListProps } from './components/ICheckboxListProps';
import CheckboxList from './components/CheckboxList';
import { IFieldConfiguration } from '../../../../lib/webparts/lookup/components/IFieldConfiguration';

export class PropertyPaneCheckboxList implements IPropertyPaneField<IPropertyPaneCheckboxListProps> {
  public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: IPropertyPaneCheckboxListInternalProps;
  private elem: HTMLElement;

  constructor(targetProperty: string, properties: IPropertyPaneCheckboxListProps) {
    this.targetProperty = targetProperty;
    this.properties = {
      key: properties.label,
      label: properties.label,
      loadOptions: properties.loadOptions,
      onChanged: properties.onChanged,
      onRender: this.onRender.bind(this)
    };
  }

  public render(): void {
    if (!this.elem) {
      return;
    }

    this.onRender(this.elem);
  }

  private onRender(elem: HTMLElement): void {
    if (!this.elem) {
      this.elem = elem;
    }

    const element: React.ReactElement<ICheckboxListProps> = React.createElement(CheckboxList, {
      label: this.properties.label,
      loadOptions: this.properties.loadOptions,
      onChanged: this.onChanged.bind(this),
      // required to allow the component to be re-rendered by calling this.render() externally
      stateKey: new Date().toString()
    });
    ReactDom.render(element, elem);
  }

  private onChanged(option: IFieldConfiguration, checked: boolean): void {
     this.properties.onChanged(option, checked);
  }
}
