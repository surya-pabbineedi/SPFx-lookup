import * as React from 'react';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/components/Dropdown';
import { Spinner } from 'office-ui-fabric-react/lib/components/Spinner';
import { ICheckboxListProps } from './ICheckboxListProps';
import { ICheckboxListState } from './ICheckboxListState';
import { Checkbox } from 'office-ui-fabric-react';
import { IFieldConfiguration } from '../../../../../lib/webparts/lookup/components/IFieldConfiguration';
import styled from 'styled-components';

const FieldContainer = styled('div')`
  button {
    padding: 3px;
    min-width: 0 !important;
  }
`;

export default class CheckboxList extends React.Component<ICheckboxListProps, ICheckboxListState> {
  private selectedKey: React.ReactText;

  constructor(props: ICheckboxListProps, state: ICheckboxListState) {
    super(props);
    // this.selectedKey = props.selectedKey;

    this.state = {
      loading: false,
      options: undefined,
      error: undefined
    };
  }

  public componentDidMount(): void {
    this.loadOptions();
  }

  public componentDidUpdate(prevProps: ICheckboxListProps, prevState: ICheckboxListState): void {
    if (this.props.stateKey !== prevProps.stateKey) {
      this.loadOptions();
    }
  }

  public render(): JSX.Element {
    const loading = this.state.loading;
    const error: JSX.Element = this.state.error !== undefined ? <div className={'ms-TextField-errorMessage ms-u-slideDownIn20'}>Error while loading items: {this.state.error}</div> : <div />;

    return (
      <div style={{ padding: 10 }}>
        {loading && <Spinner />}
        {this.state.options && <div><label>{this.props.label}</label><hr /></div>}
        {this.state.options && this.state.options.map((column: IFieldConfiguration, i) => (
          <FieldContainer>
            <Checkbox
              defaultChecked={true}
              label={column.fieldName}
              onChange={(event, checked) =>
                this.onChanged(column, checked)
              }
            />
          </FieldContainer>
        ))}
        {error}
      </div>
    );
  }

  private loadOptions(): void {
    this.setState({
      loading: true,
      error: undefined,
      options: undefined
    });

    this.props.loadOptions()
      .then((options: IFieldConfiguration[]): void => {
        console.log(options);
        this.setState({
          loading: false,
          error: undefined,
          options: options
        });
      }, (error: any): void => {
        this.setState((prevState: ICheckboxListState, props: ICheckboxListProps): ICheckboxListState => {
          prevState.loading = false;
          prevState.error = error;
          return prevState;
        });
      });
  }

  private onChanged(option: IFieldConfiguration, checked: boolean): void {
    if (this.props.onChanged) {
      this.props.onChanged(option, checked);
    }
  }
}
