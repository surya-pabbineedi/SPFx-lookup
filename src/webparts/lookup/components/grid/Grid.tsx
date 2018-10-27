import * as React from 'react';
import { IGridProps } from './IGridProps';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import {
  Button,
  ButtonType,
  Spinner,
  SpinnerSize
} from 'office-ui-fabric-react';
import {
  DetailsList,
  DetailsListLayoutMode,
  Selection,
  SelectionMode,
  IColumn
} from 'office-ui-fabric-react/lib/DetailsList';

export interface IDetailsListDocumentsExampleState {
  columns: IColumn[];
  items: IDocument[];
  selectionDetails: string;
  isModalSelection: boolean;
  isCompactMode: boolean;
}

export interface IDocument {
  [key: string]: any;
  name: string;
  value: string;
  iconName: string;
  modifiedBy: string;
  dateModified: string;
  dateModifiedValue: number;
  fileSize: string;
  fileSizeRaw: number;
}

export default class Grid extends React.Component<IGridProps, {}> {
  public state = { lists: [], isLoading: false };
  constructor() {
    super();
  }

  public componentDidMount() {
    //if (Environment.type === EnvironmentType.SharePoint) this.getLists();
  }

  private getColumns() {
    const columns = [
      {
        key: 'column1',
        name: 'File Type',
        headerClassName: 'DetailsListExample-header--FileIcon',
        className: 'DetailsListExample-cell--FileIcon',
        iconClassName: 'DetailsListExample-Header-FileTypeIcon',
        ariaLabel: 'Column operations for File type',
        iconName: 'Page',
        isIconOnly: true,
        fieldName: 'name',
        minWidth: 16,
        maxWidth: 16,
        // onColumnClick: this._onColumnClick,
        onRender: (item: IDocument) => {
          return (
            <img
              src={item.iconName}
              className={'DetailsListExample-documentIconImage'}
            />
          );
        }
      }
    ];
  }

  private getLists(): Promise<any> {
    this.setState({ isLoading: true });
    return this.props.context.spHttpClient
      .get(
        `${this.props.context.pageContext.web.absoluteUrl}/_api/lists/getbytitle('Kaizen Report')/items`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        response.json().then((responseJSON: any) => {
          this.setState({ lists: responseJSON.value, isLoading: false });
        });
      });
  }

  public render() {
    return (
      <div>
        {this.state.isLoading && <Spinner size={SpinnerSize.large} />}
        <ul>
          {this.state.lists.map(list => (
            <li>{list.Item} - {list.Subject} - {list.Area_x0020__x002f__x0020_Locatio}</li>
          ))}
        </ul>
      </div>
    );
  }
}
