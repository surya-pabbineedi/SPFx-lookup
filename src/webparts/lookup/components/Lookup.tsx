import * as React from 'react';
import styles from './Lookup.module.scss';
import { ILookupProps } from './ILookupProps';
import { KaizenList } from './data-list/data-list';
import CreateKaizen from './kaizen/create';
import { loadTheme, getTheme } from 'office-ui-fabric-react/lib/Styling';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react';
import { KaizenModel } from '../models/kaizen.model';

loadTheme({
  palette: {
    themePrimary: '#FBB543'
  }
});

export default class Lookup extends React.Component<ILookupProps, {}> {
  private kaizenList: KaizenList;
  private _createKaizen: CreateKaizen;
  public state = {
    error: '',
    model: null
  };

  constructor(props) {
    super();

    this.state = {
      error: '',
      model: null //new KaizenModel(props.context.pageContext.web.absoluteUrl)
    };
  }

  private onError = error => {
    this.setState({
      error: error.data.responseBody['odata.error'].message.value
    });
  }

  public _renderError() {
    return (
      this.state.error && (
        <div style={{ backgroundColor: '#FFF' }}>
          <MessageBar
            messageBarType={MessageBarType.blocked}
            isMultiline={false}
            dismissButtonAriaLabel="Close"
            truncated={true}
            overflowButtonAriaLabel="..."
            onDismiss={() => this.setState({ error: '' })}
          >
            {this.state.error}
          </MessageBar>
        </div>
      )
    );
  }

  public reloadData = () => {
    this.kaizenList.reloadData();
  };

  public onEdit = (model: KaizenModel) => {
    this._createKaizen.editItem(model);
  };

  public render(): React.ReactElement<ILookupProps> {
    return (
      <div className="ms-Grid">
        <div className={styles.lookup}>
          <div className={styles.container}>
            <div className={styles.row}>
              {this._renderError()}
              <h1>{this.props.listUrl}</h1>
            </div>
          </div>
        </div>
      </div>
    );
  }

  public renderOld(): React.ReactElement<ILookupProps> {
    return (
      <div className="ms-Grid">
        <div className={styles.lookup}>
          <div className={styles.container}>
            <div className={styles.row}>
              {this._renderError()}
              <CreateKaizen
                ref={node => (this._createKaizen = node)}
                model={this.state.model}
                context={this.props.context}
                onError={this.onError}
                reloadData={this.reloadData}
              />
              <div>
                <KaizenList
                  ref={node => (this.kaizenList = node)}
                  listName={this.props.listName}
                  context={this.props.context}
                  onError={this.onError}
                  onEdit={this.onEdit}
                />
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
