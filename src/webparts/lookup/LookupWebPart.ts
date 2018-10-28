import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, Environment, EnvironmentType, DisplayMode } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneToggle,
  PropertyPaneChoiceGroup
} from '@microsoft/sp-webpart-base';

import * as strings from 'LookupWebPartStrings';
import Lookup from './components/Lookup';
import { ILookupProps } from './components/ILookupProps';
import { ListService } from '../../common/services/ListService';
import { IDropdownOption, MessageBar, MessageBarType, ICheckbox } from 'office-ui-fabric-react';
import { update, get } from '@microsoft/sp-lodash-subset';
import { PropertyPaneAsyncDropdown } from '../../common/controls/PropertyPaneAsyncDropdown/PropertyPaneAsyncDropdown';
import { ControlMode } from '../../common/datatypes/ControlMode';
import { Web } from 'sp-pnp-js/lib/pnp';
import { IFieldConfiguration } from '../../../lib/webparts/lookup/components/IFieldConfiguration';
import { PropertyPaneCheckboxList } from '../../common/controls/PropertyPaneCheckboxList/PropertyPaneCheckboxList';
import ConfigureWebPart from '../../common/components/ConfigureWebpart/ConfigureWebPart';


export interface ILookupWebPartProps {
  title: string;
  description: string;
  listUrl: string;
  formType: ControlMode;
  itemId?: string;
  showUnsupportedFields: boolean;
  redirectUrl?: string;
  fields?: IFieldConfiguration[];
}

export default class LookupWebPart extends BaseClientSideWebPart<
  ILookupWebPartProps
> {
  public dynamicProps: any;
  private listService: ListService;
  private cachedLists = null;

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  protected onInit(): Promise<void> {
    return super.onInit().then( _ => {
      const web = new Web(this.context.pageContext.web.absoluteUrl);
      this.listService = new ListService(this.context.spHttpClient, web);
    });
  }

  public render(): void {
    // const element: React.ReactElement<ILookupProps> = React.createElement(
    //   Lookup,
    //   {
    //     listName: this.properties.title,
    //     context: this.context
    //   }
    // );

    // ReactDom.render(element, this.domElement);

    let itemId;
    if (this.properties.itemId) {
      itemId = Number(this.properties.itemId);
      if (isNaN(itemId)) {
        // if item Id is not a number we assume it is a query string parameter
        const urlParams = new URLSearchParams(window.location.search);
        itemId = Number(urlParams.get(this.properties.itemId));
      }
    }

    let element;
    if (Environment.type === EnvironmentType.Local) {
      // show message that local workbench is not supported
      element = React.createElement(
        MessageBar,
        {messageBarType: MessageBarType.blocked},
        strings.LocalWorkbenchUnsupported
      );
    } else if (this.properties.listUrl) {
      // show actual list form react component
      element = React.createElement(
        Lookup,
        {
          context: this.context,
          title: this.properties.title,
          description: this.properties.description,
          listUrl: this.properties.listUrl,
          fields: this.properties.fields,
          onSubmitSucceeded: (id: number) => this.formSubmitted(id),
          onUpdateFields: (fields: IFieldConfiguration[]) => this.updateField(fields),
        }
      );
    } else {
      // show configure web part react component
      element = React.createElement(
        ConfigureWebPart,
        {
          webPartContext: this.context,
          title: this.properties.title,
          description: strings.MissingListConfiguration,
          buttonText: strings.ConfigureWebpartButtonText
        }
      );
    }

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const mainGroup = {
      groupName: strings.BasicGroupName,
      groupFields: [
        PropertyPaneTextField('title', {
          label: strings.TitleFieldLabel
        }),
        PropertyPaneTextField('description', {
          label: strings.DescriptionFieldLabel,
          multiline: true
        }),
        new PropertyPaneAsyncDropdown('listUrl', {
          label: strings.ListFieldLabel,
          loadOptions: this.loadLists.bind(this),
          onPropertyChange: this.onListChange.bind(this),
          selectedKey: this.properties.listUrl
        })
        // PropertyPaneDropdown('formType', {
        //   label: strings.FormTypeFieldLabel,
        //   options: Object.keys(ControlMode)
        //                    .map( (k) => ControlMode[k]).filter( (v) => typeof v === 'string' )
        //                      .map( (n) => ({key: ControlMode[n], text: n}) ),
        //   disabled: !this.properties.listUrl
        // })        
      ]
    };

    if (this.properties.listUrl) {
      mainGroup.groupFields.push(
        new PropertyPaneCheckboxList( 'fields', {
          label: strings.FieldsLabel,
          loadOptions: this.loadFields.bind(this),
          onChanged: this.onFieldChanged.bind(this)
        }));
    }

    // mainGroup.groupFields.push(
    //   PropertyPaneToggle('showUnsupportedFields', {
    //     label: strings.ShowUnsupportedFieldsLabel,
    //     disabled: !this.properties.listUrl
    //   })
    // );
    // mainGroup.groupFields.push(
    //   PropertyPaneTextField('redirectUrl', {
    //     label: strings.RedirectUrlFieldLabel,
    //     description: strings.RedirectUrlFieldDescription,
    //     disabled: !this.properties.listUrl
    //   })
    // );
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [mainGroup]
        }
      ]
    };
  }

  private onFieldChanged(field: IFieldConfiguration, checked: boolean){
    console.log(field, checked);
  }

  private onListChange(propertyPath: string, newValue: any): void {
    const oldValue: any = get(this.properties, propertyPath);
    if (oldValue !== newValue) {
      this.properties.fields = null;
    }
    // store new value in web part properties
    update( this.properties, propertyPath, (): any => newValue );
    // refresh property Pane
    this.context.propertyPane.refresh();
    // refresh web part
    this.render();
  }


  private updateField(fields: IFieldConfiguration[]): any {
    console.log(fields);
    this.properties.fields = fields;
    // render web part again so that React List Form component is rerendered with changed fields
    this.render();
  }


  private formSubmitted(id: number) {
    if (this.properties.redirectUrl) {
      // redirect to configured URL after successfully submitting form
      window.location.href = this.properties.redirectUrl.replace('[ID]', id.toString() );
    }
  }

  private loadFields(): Promise<IFieldConfiguration[]> {
    return new Promise<IFieldConfiguration[]>((resolve: (options: IFieldConfiguration[]) => void, reject: (error: any) => void) => {
      if (Environment.type === EnvironmentType.Local) {
        resolve( [{
            key: 'fieldA',
            fieldName: 'fieldAs',
          },
          {
            key: 'fieldB',
            fieldName: 'fieldB',
          }] );
      } else if (Environment.type === EnvironmentType.SharePoint ||
                Environment.type === EnvironmentType.ClassicSharePoint) {
        try {
            return this.listService.getFields(this.properties.listUrl)
              .then( (data: any[]) => {
                const fields = data.map((f: any) => { return { key: f.InternalName, fieldName: f.Title }; });
                resolve( fields );
              } );
         
        } catch (error) {
          alert( error );
        }
      }
    });
  }

  private loadLists(): Promise<IDropdownOption[]> {
    return new Promise<IDropdownOption[]>((resolve: (options: IDropdownOption[]) => void, reject: (error: any) => void) => {
      if (Environment.type === EnvironmentType.Local) {
        resolve( [{
            key: 'sharedDocuments',
            text: 'Shared Documents',
          },
          {
            key: 'someList',
            text: 'Some List',
          }] );
      } else if (Environment.type === EnvironmentType.SharePoint ||
                Environment.type === EnvironmentType.ClassicSharePoint) {
        try {
          if (!this.cachedLists) {
            return this.listService.getListsFromWeb(this.context.pageContext.web.absoluteUrl)
              .then( (lists) => {
                this.cachedLists = lists.map( (l) => ({ key: l.url, text: l.title } as IDropdownOption) );
                resolve( this.cachedLists );
              } );
          } else {
            // using cached lists if available to avoid loading spinner every time property pane is refreshed
            return resolve( this.cachedLists );
          }
        } catch (error) {
          alert( error );
        }
      }
    });
  }
}
