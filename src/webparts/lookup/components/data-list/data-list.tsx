import * as React from 'react';
import {
  DetailsList,
  DetailsListLayoutMode,
  Selection,
  SelectionMode,
  IColumn
} from 'office-ui-fabric-react/lib/DetailsList';
import { MarqueeSelection } from 'office-ui-fabric-react/lib/MarqueeSelection';
import { Web } from 'sp-pnp-js/lib/pnp';
import { spODataEntityArray } from 'sp-pnp-js';
import { KaizenModel } from '../../models/kaizen.model';
import {
  ActionButton,
  Dialog,
  DialogType,
  DialogFooter,
  PrimaryButton,
  DefaultButton,
  Shimmer,
  TeachingBubble,
  Spinner,
  SpinnerSize,
  ChoiceGroup,
  Checkbox,
  Panel,
  PanelType,
  Persona,
  PersonaSize,
  PersonaPresence
} from 'office-ui-fabric-react';
import styled from 'styled-components';
import { assign, css } from '@uifabric/utilities';

const listName = 'Kaizen ModernWeb';
const displayColumns = [
  'Reference_x0023_',
  'Subject',
  'Area_x0020__x002f__x0020_Locatio',
  'Initial_x0020_Condition',
  'Before',
  'Approved_x0020_By',
  'Benefits_x0020_Category',
  'Process_x0020__x002f__x0020_Proj',
  'Department_x0020__x002f__x0020_D',
  'Solution_x0020_Description',
  'After',
  'Benefits_x0020_Description',
  'Validated_x0020_By',
  'Contact_x0020_Details',
  'Team_x0020_Members',
  'Implementation_x0020_Date',
  'Date_x0020_of_x0020_Completion',
  'Title',
  'Standardization_x0020_Remarks',
  'Author'
];
export { displayColumns };

export interface IKaizenListState {
  allColumns: Array<IColumn>;
  columns: Array<IColumn>;
  items: Array<KaizenModel>;
  selectionDetails: string;
  isModalSelection: boolean;
  isCompactMode: boolean;
  deleteDialog: boolean;
  isLoading?: boolean;
}

const Action = styled(ActionButton)`
  background: transparent;
  min-width: 10px;
  .ms-Button-icon {
    color: black;

    :hover {
      background: transparent;
      color: #1A70BB;
    }
  }
`;

const FieldContainer = styled('div')`
  button {
    padding: 3px;
    min-width: 0 !important;
  }
`;

const GridContainer = styled('div')`
  width: 100%;
  max-width: 100%;
`;

class GridActions extends React.Component<any, any> {
  private _menuButtonElement: HTMLElement;
  public state = {
    showTeachingBubble: false,
    showPanel: false,
    selectedColumns: {
      Reference_x0023_: true,
      Subject: true,
      Area_x0020__x002f__x0020_Locatio: true,
      Initial_x0020_Condition: true,
      Before: true,
      Approved_x0020_By: true,
      Benefits_x0020_Category: true,
      Process_x0020__x002f__x0020_Proj: true,
      Department_x0020__x002f__x0020_D: true,
      Solution_x0020_Description: true,
      After: true,
      Benefits_x0020_Description: true,
      Validated_x0020_By: true,
      Contact_x0020_Details: true,
      Team_x0020_Members: true,
      Implementation_x0020_Date: true,
      Date_x0020_of_x0020_Completion: true,
      Title: true,
      Standardization_x0020_Remarks: true,
      Author: true
    }
  };

  private _onColumnSelected = (checked: boolean, column: IColumn) => {
    this.setState(prev => {
      const columns = prev.selectedColumns;
      columns[column.fieldName] = checked;
      return {
        selectedColumns: columns
      };
    });
  }

  private _onClosePanel = (): void => {
    this.setState({ showPanel: false });
  }

  private _onRenderFooterContent = (): JSX.Element => {
    return (
      <div>
        <PrimaryButton
          onClick={() => {
            this.props.onColumnsSelected(this.state.selectedColumns);
            this._onClosePanel();
          }}
          style={{ marginRight: '8px' }}
        >
          Apply
        </PrimaryButton>
        <DefaultButton onClick={this._onClosePanel}>Cancel</DefaultButton>
      </div>
    );
  }

  public render() {
    return (
      <div className="ms-Grid ms-u-sm12">
        <div className="ms-Grid-col ms-u-sm6">
          {this.props.selectionDetails}
        </div>
        <div className="ms-Grid-col ms-u-sm6">
          <div style={{ float: 'right' }}>
            {this.props.selectionCount === 1 && (
              <Action
                onClick={() => {
                  if (this.props.selectionCount === 0)
                    this.setState({ showTeachingBubble: true });

                  this.props.onEdit();
                }}
                iconProps={{ iconName: 'edit' }}
              />
            )}
            <Action
              innerRef={node => (this._menuButtonElement = node)}
              onClick={() => {
                if (this.props.selectionCount === 0)
                  this.setState({ showTeachingBubble: true });
                this.props.onDelete();
              }}
              iconProps={{ iconName: 'delete' }}
            />
            <Action
              // innerRef={node => (this._menuButtonElement = node)}
              onClick={() => {
                this.setState({ showPanel: true });
              }}
              iconProps={{ iconName: 'settings' }}
            />

            {this.props.selectionCount === 0 &&
              this.state.showTeachingBubble && (
                <TeachingBubble
                  isWide={true}
                  targetElement={this._menuButtonElement}
                  hasCondensedHeadline={true}
                  onDismiss={() => this.setState({ showTeachingBubble: false })}
                  hasCloseIcon={true}
                >
                  Select at least one item for edit/delete actions
                </TeachingBubble>
              )}

            <Panel
              isOpen={this.state.showPanel}
              type={PanelType.smallFixedFar}
              onDismiss={this._onClosePanel}
              headerText="Select desired columns to show in the grid"
              closeButtonAriaLabel="Close"
              onRenderFooterContent={this._onRenderFooterContent}
            >
            <div style={{display: 'flex', justifyContent:'flex-start', flexDirection: 'column'}}>
              {this.props.columns.map((column: IColumn, i) => (
                <FieldContainer>
                  <Checkbox
                    disabled={column.key == '0'}
                    defaultChecked={
                      this.state.selectedColumns[column.fieldName]
                    }
                    label={column.name}
                    onChange={(event, checked) => {
                      this._onColumnSelected(checked, column);
                    }}
                  />
                </FieldContainer>
              ))}
              </div>
            </Panel>
          </div>
        </div>
      </div>
    );
  }
}

export class KaizenList extends React.Component<any, IKaizenListState> {
  private _selection: Selection;
  private _web: Web;
  constructor(props: any) {
    super(props);
    this._web = new Web(this.props.context.pageContext.web.absoluteUrl);
    this._selection = new Selection({
      onSelectionChanged: () => {
        this.setState({
          selectionDetails: this._getSelectionDetails(),
          isModalSelection: this._selection.isModal()
        });
      }
    });

    this.state = {
      items: [],
      allColumns: [],
      columns: [],
      selectionDetails: this._getSelectionDetails(),
      isModalSelection: this._selection.isModal(),
      isCompactMode: false,
      deleteDialog: true,
      isLoading: true
    };
  }

  public componentDidMount() {
    this.getColumns();
    this.getLists();
  }

  public reloadData = () => {
    this.getLists();
    this._selection.setItems([]);
  }

  private getColumns(): Promise<any> {
    return this._web.lists
      .getByTitle(listName)
      .fields.orderBy('InternalName', true)
      .get()
      .then((response: any[]) => {
        const columns: IColumn[] = [];

        displayColumns.forEach((col, i) => {
          if (response.filter(f => f.InternalName === col).length > 0) {
            const field = response.filter(f => f.InternalName === col)[0];
            columns.push({
              key: i.toString(),
              name: field.Title,
              className: `${field.InternalName}-field`,
              fieldName: field.InternalName,
              onColumnClick: this._onColumnClick,
              minWidth: 30,
              maxWidth: 200,
              onRender: (item, index, column: IColumn) => {
                const isPersona = typeof item[column.fieldName] === 'object';
                const imageURL = this.props.context.pageContext.web.absoluteUrl + "/_layouts/15/userphoto.aspx?size=M&accountname=" + item[column.fieldName].EMail;
                if (isPersona) {
                  return <Persona text={item[column.fieldName].Title} 
                  imageUrl={imageURL} 
                  // imageInitials={item[column.fieldName].Title ? item[column.fieldName].Title.match(/\b(\w)/g).join('') : ''}
                  size={PersonaSize.size32} 
                  presence={PersonaPresence.none} />;
                }

                return <div dangerouslySetInnerHTML={{ __html: item[column.fieldName] }} />;
              }
            });
          }
        });

        this.setState({
          allColumns: columns,
          columns
        });
      });
  }

  private getLists(): Promise<any> {
    this.setState({ isLoading: true });
    return this._web.lists
      .getByTitle(listName)
      .items.orderBy('Modified', false)
      .select('Id',
        'Reference_x0023_',
        'Subject',
        'Area_x0020__x002f__x0020_Locatio',
        'Initial_x0020_Condition',
        'Before',
        'Benefits_x0020_Category',
        'Process_x0020__x002f__x0020_Proj',
        'Department_x0020__x002f__x0020_D',
        'Solution_x0020_Description',
        'After',
        'Benefits_x0020_Description',
        'Contact_x0020_Details',
        'Team_x0020_Members',
        'Implementation_x0020_Date',
        'Date_x0020_of_x0020_Completion',
        'Title',
        'Standardization_x0020_Remarks',
        'Author', 'Author/Id', 'Author/Name', 'Author/Title', 'Author/EMail',
        'Validated_x0020_By', 'Validated_x0020_By/Id', 'Validated_x0020_By/Name', 'Validated_x0020_By/Title', 'Validated_x0020_By/EMail',
        'Approved_x0020_By', 'Approved_x0020_By/Id', 'Approved_x0020_By/Name', 'Approved_x0020_By/Title', 'Approved_x0020_By/EMail',)
      .expand('Author', 'Validated_x0020_By', 'Approved_x0020_By')
      .getAs(spODataEntityArray(KaizenModel))
      .then((items: Array<KaizenModel>) => {
        this.setState({ items, isLoading: false });
      })
      .catch(error => this.props.onError(error));
  }

  private onDelete = () => {
    const selectionCount = this._selection.getSelectedCount();
    if (selectionCount === 0) return;

    this._showDialog();
  }

  private onEdit = () => {
    const selectedItem = this._selection.getSelection()[0] as KaizenModel;
    this.props.onEdit(selectedItem);
  }

  private deleteItem = () => {
    const selectionCount = this._selection.getSelectedCount();
    if (selectionCount === 0) return;

    const selectedItems = this._selection.getSelection() as Array<KaizenModel>;

    this._closeDialog();
    const batch = this._web.createBatch();

    selectedItems.forEach(item =>
      this._web.lists
        .inBatch(batch)
        .getByTitle(listName)
        .items.getById(item.Id)
        .delete()
        .then(x => this.getLists())
    );

    // batch.execute().then(() => {
    //   this.getLists();
    // });
  }

  private onColumnsSelected = columns => {
    const newColumns = [];
    this.state.allColumns.forEach(col => {
      if (columns[col.fieldName]) {
        newColumns.push(col);
      }
    });

    this.setState({ columns: newColumns });
  }

  private getDeleteDialogHeader = () => {
    const selectionCount = this._selection.getSelectedCount();
    if (selectionCount === 0) return '';

    const selection = this._selection.getSelection();
    return `Delete Kaizen items - ${
      selectionCount === 1 ? (selection[0] as any).Title : selectionCount
    }`;
  }

  public render() {
    const {
      allColumns,
      columns,
      isCompactMode,
      items,
      selectionDetails
    } = this.state;
    return (
      <div>
        <GridActions
          selectionDetails={selectionDetails}
          selectionCount={this._selection.getSelectedCount()}
          onDelete={this.onDelete}
          onEdit={this.onEdit}
          columns={allColumns}
          onColumnsSelected={this.onColumnsSelected}
        />

        <Spinner
          style={{ marginTop: 10 }}
          hidden={this.state.isLoading === false}
          size={SpinnerSize.large}
          label="loading..."
          ariaLive="assertive"
        />

        <GridContainer style={{maxWidth: (window.innerWidth - 300)}}>
          <MarqueeSelection selection={this._selection}>
            <DetailsList
              enableShimmer={true}
              items={items}
              compact={isCompactMode}
              columns={columns}
              selectionMode={SelectionMode.multiple}
              setKey="set"
              layoutMode={DetailsListLayoutMode.fixedColumns}
              isHeaderVisible={true}
              selection={this._selection}
              selectionPreservedOnEmptyClick={true}
              // onItemInvoked={this._onItemInvoked}
              enterModalSelectionOnTouch={true}
            />
          </MarqueeSelection>
        </GridContainer>

        <Dialog
          hidden={this.state.deleteDialog}
          onDismiss={this._closeDialog}
          dialogContentProps={{
            type: DialogType.normal,
            title: this.getDeleteDialogHeader(),
            subText: `Are you sure deleting item(s)?`
          }}
          modalProps={{
            titleAriaId: 'myLabelId',
            subtitleAriaId: 'mySubTextId',
            isBlocking: false,
            containerClassName: 'ms-dialogMainOverride'
          }}
        >
          <DialogFooter>
            <PrimaryButton onClick={this.deleteItem} text="Delete" />
            <DefaultButton onClick={this._closeDialog} text="Cancel" />
          </DialogFooter>
        </Dialog>
      </div>
    );
  }

  private _showDialog = (): void => {
    this.setState({ deleteDialog: false });
  }

  private _closeDialog = (): void => {
    this.setState({ deleteDialog: true });
  }

  public componentDidUpdate(
    previousProps: any,
    previousState: IKaizenListState
  ) {
    if (previousState.isModalSelection !== this.state.isModalSelection) {
      this._selection.setModal(this.state.isModalSelection);
    }
  }

  private _getSelectionDetails(): string {
    const selectionCount = this._selection.getSelectedCount();

    switch (selectionCount) {
      case 0:
        return '';
      case 1:
        return (
          '1 item selected: ' +
          (this._selection.getSelection()[0] as any).Subject
        );
      default:
        return `${selectionCount} items selected`;
    }
  }

  private _onColumnClick = (
    ev: React.MouseEvent<HTMLElement>,
    column: IColumn
  ): void => {
    const { columns } = this.state;
    const newColumns: IColumn[] = columns.slice();
    const currColumn: IColumn = newColumns.filter((currCol: IColumn) => {
      return column.key === currCol.key;
    })[0];
    newColumns.forEach((newCol: IColumn) => {
      if (newCol === currColumn) {
        currColumn.isSortedDescending = !currColumn.isSortedDescending;
        currColumn.isSorted = true;
      } else {
        newCol.isSorted = false;
        newCol.isSortedDescending = true;
      }
    });
  }
}
