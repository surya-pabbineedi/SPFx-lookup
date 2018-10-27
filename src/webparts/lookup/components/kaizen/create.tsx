import * as React from 'react';
import {
  TextField,
  TooltipHost,
  DirectionalHint,
  Dropdown,
  DatePicker,
  DayOfWeek,
  Pivot,
  PivotLinkSize,
  PivotLinkFormat,
  PivotItem,
  DefaultButton,
  PrimaryButton,
  MessageBar,
  MessageBarType,
  MessageBarButton
} from 'office-ui-fabric-react';
import styled from 'styled-components';
import DayPickerStrings from '../shared/DatePickerData';
import { KaizenModel } from '../../models/kaizen.model';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { assign } from '@uifabric/utilities';
import { PeoplePicker } from '../shared/PeoplePicker';
import { Web, ItemAddResult, ItemUpdateResult } from 'sp-pnp-js';

export interface ICreateKaizenProps {
  context: WebPartContext;
  onError: Function;
  reloadData: Function;
  model?: KaizenModel;
}

export interface ICreateKaizenState {
  isNew: boolean;
  selectedSection: number;
  model: KaizenModel;
  web: Web;
  error?: string;
}

const AlignRightContainer = styled('div')`
  padding: 10px;
  display: flex;
  justify-content: flex-end;
`;

const toolTipCloseDelay = 500;

const BeforeDataSource = [
  {
    key: 'Cost',
    text: 'Cost'
  },
  {
    key: 'Customer responsiveness',
    text: 'Customer responsiveness'
  },
  {
    key: 'Delivery',
    text: 'Delivery'
  },
  {
    key: 'Efficiency',
    text: 'Efficiency'
  },
  {
    key: 'Environment / Energy',
    text: 'Environment / Energy'
  },
  {
    key: 'Morale',
    text: 'Morale'
  },
  {
    key: 'Quality',
    text: 'Quality'
  },
  {
    key: 'Safety / Health',
    text: 'Safety / Health'
  },
  {
    key: 'Waste',
    text: 'Waste'
  },
  {
    key: 'Other',
    text: 'Other'
  }
];

export default class CreateKaizen extends React.Component<
  ICreateKaizenProps,
  ICreateKaizenState
> {
  constructor(props) {
    super();
    this.state = {
      isNew: false,
      selectedSection: 0,
      model: new KaizenModel(props.context.pageContext.web.absoluteUrl),
      web: new Web(props.context.pageContext.web.absoluteUrl)
    };
  }

  public newItem = () => {
    this.setState((prevState: any) => {
      return {
        isNew: !prevState.isNew,
        model: new KaizenModel(this.props.context.pageContext.web.absoluteUrl),
        selectedSection: 0
      };
    });
  }

  public next = () => {
    this.setState((prev: any) => {
      return {
        selectedSection:
          prev.selectedSection === 2 ? 0 : prev.selectedSection + 1
      };
    });
  }

  public addItem = () => {
    if (!this.state.model) return;

    const keys = Object.keys(this.state.model);
    const values = {};
    keys.forEach((key: string) => {
      if (key.indexOf('_') != 0) values[key] = this.state.model[key];
    });

    const listName = 'Kaizen ModernWeb';

    if (this.state.model.Id) {
      this.state.web.lists
        .getByTitle(listName)
        .items.getById(this.state.model.Id)
        .update(values)
        .then(() => {
            this.props.reloadData();
            this.setState({ model: null, isNew: false });
          })
        .catch(error => this.props.onError(error));
    } else {
      this.state.web.lists
        .getByTitle(listName)
        .items.add(values)
        .then(() => {
            this.props.reloadData();
            this.setState({ model: null, isNew: false });
          })
        .catch(error => this.props.onError(error));
    }
  }

  public editItem(model: KaizenModel) {
    this.setState({ model, isNew: true });
  }

  public stripHTML(input) {
    if (input) {
      return input.replace(/<\/?[^>]+(>|$)/g, '');
    }

    return '';
  }

  public render() {
    return (
      <div style={{ backgroundColor: '#FFFFFF' }}>
        <MessageBar
          truncated={true}
          overflowButtonAriaLabel="..."
          actions={
            <MessageBarButton
              primary={true}
              onClick={this.newItem}
              iconProps={{ iconName: this.state.isNew ? 'Cancel' : 'Add' }}
            >
              {this.state.isNew ? 'Cancel' : 'New Kaizen'}
            </MessageBarButton>
          }
          messageBarType={MessageBarType.info}
          isMultiline={false}
        >
          {`Kaizen ${this.state.model && this.state.model.Subject
              ? ' - ' + this.state.model.Subject
              : ''}`}
        </MessageBar>
        {this.state.isNew &&
          this.state.model && (
            <div>
              <Pivot
                linkSize={PivotLinkSize.large}
                linkFormat={PivotLinkFormat.links}
                selectedKey={`${this.state.selectedSection}`}
                onLinkClick={(event: any) =>
                  this.setState({ selectedSection: +event.props.itemKey })
                }
              >
                <PivotItem linkText="Reference" itemKey="0">
                  <div className="ms-Grid" style={{ marginTop: 10 }}>
                    <div className="ms-Grid-col ms-u-sm6">
                      <TextField
                        label="Reference "
                        required={true}
                        value={this.state.model.Reference_x0023_}
                        onBeforeChange={(newValue: string) => {
                          this.setState(prevState => {
                            return {
                              model: assign(prevState.model, {
                                Reference_x0023_: newValue
                              })
                            };
                          });
                        }}
                      />
                      <TextField
                        label="Subject "
                        value={this.state.model.Subject}
                        onBeforeChange={newValue => {
                          this.setState(prevState => {
                            return {
                              model: assign(prevState.model, {
                                Subject: newValue
                              })
                            };
                          });
                        }}
                      />
                      <TextField
                        label="Area / Location "
                        value={
                          this.state.model.Area_x0020__x002f__x0020_Locatio
                        }
                        onBeforeChange={newValue => {
                          this.setState(prevState => {
                            return {
                              model: assign(prevState.model, {
                                Area_x0020__x002f__x0020_Locatio: newValue
                              })
                            };
                          });
                        }}
                      />
                      <TooltipHost
                        content="What was the problem?"
                        closeDelay={toolTipCloseDelay}
                        directionalHint={DirectionalHint.rightCenter}
                      >
                        <TextField
                          label="Initial Condition "
                          multiline
                          value={this.stripHTML(this.state.model.Initial_x0020_Condition)}
                          onBeforeChange={newValue => {
                            this.setState(prevState => {
                              return {
                                model: assign(prevState.model, {
                                  Initial_x0020_Condition: newValue
                                })
                              };
                            });
                          }}
                        />
                      </TooltipHost>
                    </div>
                    <div className="ms-Grid-col ms-u-sm6">
                      <Dropdown
                        placeHolder="Select an Option"
                        label="Benefits Category "
                        ariaLabel="Benefits Category"
                        options={BeforeDataSource}
                        defaultSelectedKey={
                          this.state.model.Benefits_x0020_Category
                        }
                        onSelect={newValue => {
                          this.setState(prevState => {
                            return {
                              model: assign(prevState.model, {
                                Benefits_x0020_Category: newValue
                              })
                            };
                          });
                        }}
                      />
                      <PeoplePicker
                        context={this.props.context}
                        label="Approved By "
                        values={this.state.model.Approved_x0020_By ? [this.state.model.Approved_x0020_By.Id] : []}
                        onChange={newValue => {
                          this.setState(prevState => {
                            return {
                              model: assign(prevState.model, {
                                AuthorId: newValue
                              })
                            };
                          });
                        }}
                      />
                      <TooltipHost
                        content="Include pictures, diagrams, etc."
                        closeDelay={toolTipCloseDelay}
                        directionalHint={DirectionalHint.rightCenter}
                      >
                        <TextField
                          multiline
                          style={{ height: 115 }}
                          label="Before "
                          value={this.stripHTML(this.state.model.Before)}
                          onBeforeChange={newValue => {
                            this.setState(prevState => {
                              return {
                                model: assign(prevState.model, {
                                  Before: newValue
                                })
                              };
                            });
                          }}
                        />
                      </TooltipHost>
                    </div>
                  </div>
                </PivotItem>
                <PivotItem linkText="Process/Project" itemKey="1">
                  <div className="ms-Grid" style={{ marginTop: 10 }}>
                    <div className="ms-Grid-col ms-u-sm6">
                      <TextField
                        label="Process / Project "
                        value={
                          this.state.model.Process_x0020__x002f__x0020_Proj
                        }
                        onBeforeChange={newValue => {
                          this.setState(prevState => {
                            return {
                              model: assign(prevState.model, {
                                Process_x0020__x002f__x0020_Proj: newValue
                              })
                            };
                          });
                        }}
                      />
                      <TextField
                        label="Department / Division "
                        value={
                          this.state.model.Department_x0020__x002f__x0020_D
                        }
                        onBeforeChange={newValue => {
                          this.setState(prevState => {
                            return {
                              model: assign(prevState.model, {
                                Department_x0020__x002f__x0020_D: newValue
                              })
                            };
                          });
                        }}
                      />
                      <TooltipHost
                        content="What was implemented, changed or improved?"
                        closeDelay={toolTipCloseDelay}
                        directionalHint={DirectionalHint.rightCenter}
                      >
                        <TextField
                          multiline
                          label="Solution Description "
                          value={this.stripHTML(this.state.model.Solution_x0020_Description)}
                          onBeforeChange={newValue => {
                            this.setState(prevState => {
                              return {
                                model: assign(prevState.model, {
                                  Solution_x0020_Description: newValue
                                })
                              };
                            });
                          }}
                        />
                      </TooltipHost>
                      <TooltipHost
                        content="Include pictures, diagrams, etc."
                        closeDelay={toolTipCloseDelay}
                        directionalHint={DirectionalHint.rightCenter}
                      >
                        <TextField
                          multiline
                          label="After "
                          value={this.stripHTML(this.state.model.After)}
                          onBeforeChange={newValue => {
                            this.setState(prevState => {
                              return assign(prevState, {
                                model: assign(prevState.model, {
                                  After: newValue
                                })
                              });
                            });
                          }}
                        />
                      </TooltipHost>
                    </div>
                    <div className="ms-Grid-col ms-u-sm6">
                      <PeoplePicker
                        context={this.props.context}
                        label="Validated By "
                        values={this.state.model.Validated_x0020_By ? [this.state.model.Validated_x0020_By.Id] : []}
                        onChange={newValue => {
                          this.setState(prevState => {
                            return {
                              model: assign(prevState.model, {
                                Validated_x0020_ById: newValue
                              })
                            };
                          });
                        }}
                      />
                      <TooltipHost
                        content="e.g. results, cost benefit analysis, cost savings etc."
                        closeDelay={toolTipCloseDelay}
                        directionalHint={DirectionalHint.rightCenter}
                      >
                        <TextField
                          multiline
                          label="Benefits Description "
                          value={this.stripHTML(this.state.model.Benefits_x0020_Description)}
                          onBeforeChange={newValue => {
                            this.setState(prevState => {
                              return {
                                model: assign(prevState.model, {
                                  Benefits_x0020_Description: newValue
                                })
                              };
                            });
                          }}
                        />
                      </TooltipHost>
                      <TextField
                        multiline
                        label="Contact Details "
                        value={this.stripHTML(this.state.model.Contact_x0020_Details)}
                        onBeforeChange={newValue => {
                          this.setState(prevState => {
                            return {
                              model: assign(prevState.model, {
                                Contact_x0020_Details: newValue
                              })
                            };
                          });
                        }}
                      />
                    </div>
                  </div>
                </PivotItem>
                <PivotItem linkText="Who was involved?" itemKey="2">
                  <div className="ms-Grid" style={{ marginTop: 10 }}>
                    <div className="ms-Grid-col ms-u-sm12">
                      <TooltipHost
                        content="Who was involved?"
                        closeDelay={toolTipCloseDelay}
                        directionalHint={DirectionalHint.rightCenter}
                      >
                        <TextField
                          multiline
                          label="Team Members "
                          value={this.stripHTML(this.state.model.Team_x0020_Members)}
                          onBeforeChange={newValue => {
                            this.setState(prevState => {
                              return {
                                model: assign(prevState.model, {
                                  Team_x0020_Members: newValue
                                })
                              };
                            });
                          }}
                        />
                      </TooltipHost>
                    </div>
                    <div className="ms-Grid-col ms-u-sm6">
                      <DatePicker
                        firstDayOfWeek={DayOfWeek.Sunday}
                        strings={DayPickerStrings}
                        label="Implementation Date "
                        placeholder="Select a date..."
                        value={this.state.model.Implementation_x0020_Date ? this.state.model.Implementation_x0020_Date: null}
                        onSelectDate={newValue => {
                          this.setState(prevState => {
                            return {
                              model: assign(prevState.model, {
                                Implementation_x0020_Date: newValue
                              })
                            };
                          });
                        }}
                      />
                    </div>
                    <div className="ms-Grid-col ms-u-sm6">
                      <DatePicker
                        firstDayOfWeek={DayOfWeek.Sunday}
                        strings={DayPickerStrings}
                        label="Date of Completion"
                        placeholder="Select a date..."
                        value={this.state.model.Date_x0020_of_x0020_Completion ? this.state.model.Date_x0020_of_x0020_Completion : null}
                        onSelectDate={newValue => {
                          this.setState(prevState => {
                            return {
                              model: assign(prevState.model, {
                                Date_x0020_of_x0020_Completion: newValue
                              })
                            };
                          });
                        }}
                      />
                    </div>
                  </div>
                  <div className="ms-Grid-col ms-u-sm12">
                    <TextField
                      multiline
                      label="Standardization Remarks "
                      value={this.stripHTML(this.state.model.Standardization_x0020_Remarks)}
                      onBeforeChange={newValue => {
                        this.setState(prevState => {
                          return {
                            model: assign(prevState.model, {
                              Standardization_x0020_Remarks: newValue
                            })
                          };
                        });
                      }}
                    />
                  </div>
                </PivotItem>
              {/* <PivotItem linkText="Sample editor" itemKey="3">
                <div className="ms-Grid" style={{ marginTop: 10 }}>
                  <div className="ms-Grid-col ms-u-sm12">
                    <Editor
                      init={{
                        plugins: ['paste', 'link']
                      }}
                      // initialValue={this.state.content}
                      // onChange={(event) => { this.handleChange(event.target.getContent()); }}
                    />
                  </div>
                </div>
              </PivotItem> */}
              </Pivot>
              <AlignRightContainer>
                <DefaultButton onClick={this.next} text="Next" />
                <PrimaryButton
                  iconProps={{ iconName: 'Add' }}
                  text={this.state.model.Id ? 'Update' : 'Create'}
                  style={{ marginLeft: '3px' }}
                  onClick={this.addItem}
                />
              </AlignRightContainer>
            </div>
          )}
      </div>
    );
  }
}
