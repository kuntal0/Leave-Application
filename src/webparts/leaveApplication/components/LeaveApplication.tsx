import * as React from 'react';
import styles from './LeaveApplication.module.scss';
import { ILeaveApplicationProps } from './ILeaveApplicationProps';
//import { escape } from '@microsoft/sp-lodash-subset';
import * as moment from "moment";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
//import { Stack, IStackTokens } from '@fluentui/react';
//import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
//import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';

// import PnPTelemetry from "@pnp/telemetry-js";
// const telemetry = PnPTelemetry.getInstance();
// telemetry.optOut();



interface IListItem {
  ID: number;
  Title: string;
  ManagerEmail: string;
  StartDate: any;
  EndDate: any;
  Reason: string;
  Status: any;
}

interface IListItems {
  AllItems: IListItem[];

  //
  listTitle: string;
  listManagerEmail: string;
  listStartDate: any;
  listEndDate: any;
  listReason: string;
  liststatus: any;
  listSelectedID: number;
}
export default class LeaveApplication extends React.Component<ILeaveApplicationProps, IListItems>{


  constructor(props: ILeaveApplicationProps, state: IListItems) {
    super(props);
    this.state = {
      AllItems: [],
      listSelectedID: 0,
      listTitle: undefined,
      listManagerEmail: undefined,
      listStartDate: 0,
      listEndDate: 0,
      listReason: undefined,
      liststatus: undefined,
    };
  }

  componentDidMount(): void {
    this.getListItems();
  }

  // Get items
  public getListItems = () => {
    let listName = `Leave App Form`;

    let requestURL = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items`;
    this.props.context.spHttpClient
      .get(requestURL, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          return response.json();
        }
      })
      .then((i) => {
        if (i == undefined) {
        } else {
          this.setState({
            AllItems: i.value,
          });
          console.log(this.state.AllItems);
        }
      });
  };

  // Delete item
  public deleteItem = (itemID: number) => {
    // alert("this is delete");
    let listName = `Leave App Form`;

    let requestURL = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items(${itemID})`;

    this.props.context.spHttpClient
      .post(requestURL, SPHttpClient.configurations.v1, {
        headers: {
          Accept: "application/json;odata=nometadata",
          "Content-type": "application/json;odata=verbose",
          "odata-version": "",
          "IF-MATCH": "*",
          "X-HTTP-Method": "DELETE",
        },
      })
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          alert(`Item ID: ${itemID} deleted successfully!`);
          this.getListItems();
        } else {
          alert(`Something went wrong!`);
          console.log(response.json());
        }
      });
  };

  // Add item
  public addItemInList = () => {
    // alert("this is delete");
    let listName = `Leave App Form`;

    let requestURL = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items`;

    const body: string = JSON.stringify({
      Title: this.state.listTitle,
      ManagerEmail: this.state.listManagerEmail,
      StartDate: this.state.listStartDate,
      EndDate: this.state.listEndDate,
      Reason: this.state.listReason,
      Status: this.state.liststatus

    });

    this.props.context.spHttpClient
      .post(requestURL, SPHttpClient.configurations.v1, {
        headers: {
          Accept: "application/json;odata=nometadata",
          "Content-type": "application/json;odata=nometadata",
          "odata-version": "",
        },
        body: body,
      })
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          alert(`Item added successfully!`);
          this.getListItems();
        } else {
          alert(`Something went wrong!`);
          console.log(response.json());
        }
      });
  };

  //Update item
  public updateItemInList = (itemID: number) => {
    // alert("this is delete");
    let listName = `Leave App Form`;

    let requestURL = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items(${itemID})`;

    const body: string = JSON.stringify({
      Title: this.state.listTitle,
      ManagerEmail: this.state.listManagerEmail,
      StartDate: this.state.listStartDate,
      EndDate: this.state.listEndDate,
      Reason: this.state.listReason,
      Status: this.state.liststatus
    });

    this.props.context.spHttpClient
      .post(requestURL, SPHttpClient.configurations.v1, {
        headers: {
          Accept: "application/json;odata=nometadata",
          "Content-type": "application/json;odata=nometadata",
          "odata-version": "",
          "IF-MATCH": "*",
          "X-HTTP-Method": "MERGE",
        },
        body: body,
      })
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          alert(`Item updated successfully!`);
          this.getListItems();
        } else {
          alert(`Something went wrong!`);
          console.log(response.json());
        }
      });
  };

  public render(): React.ReactElement<ILeaveApplicationProps> {


    return (
      <><h1>Leave Application</h1>
      <div className={styles.leaveApplication}>
        <h3>Employee Name  <input
          value={this.state.listTitle}
          type="text"
          name=""
          id="lsTitle"
          placeholder="Employee Name"
          onChange={(e) => {
            this.setState({
              listTitle: e.currentTarget.value,
            });
            // console.log(this.state.listTitle);
          } } />
          </h3>
        <h3>
          Manager Email  <input
          value={this.state.listManagerEmail}
          type="text"
          name=""
          id="lsManagerEmail"
          placeholder="ManagerEmail"
          onChange={(e) => {
            this.setState({
              listManagerEmail: e.currentTarget.value,
            });
          } } />
          </h3>

        <h3>Start Date  <input
          value={this.state.listStartDate}
          type="date"
          name=""
          id="lsStartdate"
          placeholder="StartDate"
          onChange={(e) => {
            this.setState({
              listStartDate: e.currentTarget.value as any,
            });
          } } />
          </h3>

          <h3>End Date  <input
          value={this.state.listEndDate}
          type="date"
          name=""
          id="lsEnddate"
          placeholder="EndDate"
          onChange={(e) => {
            this.setState({
              listEndDate: e.currentTarget.value as any,
            });
          } } />
          </h3>

          <h3>Reason  <input
          value={this.state.listReason}
          type="text"
          name=""
          id="lsReason"
          placeholder="Reason"
          onChange={(e) => {
            this.setState({
              listReason: e.currentTarget.value,
            });
          } } />
          </h3>

          <h3>Status  <select 
          value={this.state.liststatus}
          placeholder='Status'
          id="Isstatus" name="status"
             onChange={(e) => {
            this.setState({
              liststatus: e.currentTarget.value as any,
            });
            
          }}>

          <option value="Pending">Pending</option>
          <option value="Approved">Approved</option>
          <option value="Rejected">Rejected</option>
        </select>
        </h3>


        {/* <h3>Status  <select
          value={this.state.liststatus}
          name=""
          id="lsStatus"
          placeholder="Status"
          onChange={(e) => {
            this.setState({
              liststatus: e.currentTarget.value as any,
            });
          } } />
          
          <option value="Married">Married</option>
          <option value="Unmarried">Unmarried</option>
          <option value="Widow">Widow</option>
          
          </h3> */}
          

        <br /><br />

        <button
          onClick={() => {
            this.addItemInList();
          } }
        >
          Submit
        </button>
        
        <button
          onClick={() => {
            this.updateItemInList(this.state.listSelectedID);
          } }
        >
          Update
        </button>
        <hr />
        <table>
          <th>Employee Name</th>
          <th>Manager Email</th>
          <th>Start Date</th>
          <th>End Date</th>
          <th>Reason</th>
          <th>Status</th>
          {this.state.AllItems.map((emp) => {
            return (
              <tr>
                <td>{emp.Title}</td>
                <td>{emp.ManagerEmail}</td>
                <td>{moment(emp.StartDate).format("LL")}</td>
                <td>{moment(emp.EndDate).format("LL")}</td>
                <td>{emp.Reason}</td>
                <td>{emp.Status}</td>
                <td>
                  <button
                    onClick={() => {
                      this.setState({
                        listTitle: emp.Title,
                        listManagerEmail: emp.ManagerEmail,
                        listSelectedID: emp.ID,
                        listStartDate: emp.StartDate,
                        listEndDate: emp.EndDate,
                        listReason: emp.Reason,
                        liststatus: emp.Status,
                      });
                    } }
                  >
                    Edit
                  </button>
                </td>
                <td>
                  <button
                    onClick={() => {
                      this.deleteItem(emp.ID);
                    } }
                  >
                    Delete
                  </button>
                </td>
              </tr>
            );
          })}
        </table>
      </div></>
    );
  }
}
