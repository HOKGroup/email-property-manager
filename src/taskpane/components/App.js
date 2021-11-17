import * as React from "react";
import PropTypes from "prop-types";
import axios from 'axios';

import { DefaultButton } from "@fluentui/react";
import Header from "./Header";
import HeroList from "./HeroList";
import SwitchLabels from "./SwitchLabels";
import Progress from "./Progress";
import InputForm from "./InputForm";
import CustomizedTables from "./PropTable";
import Button from '@mui/material/Button';

import "../styles/table.css";

// images references in the manifest
import "../../../assets/icon-16.png";
import "../../../assets/icon-32.png";
import "../../../assets/icon-80.png";


/* global require */

export default class App extends React.Component {
  constructor(props, context) {
    super(props, context);
      this.state = {
        isReadMode: false,
        keys: [],
        customProps: null,
        itemId: null,
        authToken: null,
        uploadCustomProps: null,
    };
  }

  componentDidMount() {
     
  }

  componentDidUpdate() {
 
  }

  saveMailItem = () => {
      Office.context.mailbox.item.saveAsync(
          (result) => {
              let id = result.value;
              console.log(id);
              this.setState({ itemId: id });
          });
  }

  uploadCustomProps = () => {
      console.log("start upload");
      this.saveMailItem();
      setTimeout(this.clickGetMailInfo, 500);
      setTimeout(this.upload, 1000);
  }

  upload = () => {
      const { itemId, authToken } = this.state;
      console.log(this.state.keys);
      let keyValPair = this.state.keys.map((pair) => pair.name + "#" + pair.value);
      console.log(keyValPair);
      let rest_url = "https://outlook.office365.com/api/v2.0/me/messages('" + itemId + "')";
      axios.patch(rest_url,
          {
              "MultiValueExtendedProperties": [
                  {
                      "PropertyId": "StringArray {00020329-0000-0000-C000-000000000046} Name myTest",
                      "Value": keyValPair
                  }
              ]
          },
          {
              dataType: 'json',
              headers: {
                  "Authorization": "Bearer " + authToken
              }
          }
      ).then(
          item => {
              console.log(item);
              this.setState({ status: item.status });
          }
      ).catch(
          err => { console.log(JSON.stringify(err.message)) }
      );
  }

  clickGetMailInfo = async () => {
      Office.context.mailbox.getCallbackTokenAsync(
          {
              isRest: true
          },
          (asyncResult) => {
              console.log(Office.context.mailbox.item);
              let token = asyncResult.value;
              this.setState({ authToken: token });
          }
      );

      Office.context.mailbox.item.getItemIdAsync((result) => {
          if (result.status !== Office.AsyncResultStatus.Succeeded) {
              console.error(`getItemIdAsync failed with message: ${result.error.message}`);
          } else {
              console.log(result.value);
              let id = result.value;
              console.log(id.replace("-", "/"));
              this.setState({ itemId: id });
          }
      });
  }

  getCustomProps = async () => {
      this.setState({item : Office.context.mailbox.item});
      Office.context.mailbox.getCallbackTokenAsync(
          {
              isRest: true
          },
          (asyncResult) => {
              console.log(asyncResult.value);
              if (asyncResult.status === Office.AsyncResultStatus.Succeeded
                  && asyncResult.value !== "") {
                  let item_rest_id = Office.context.mailbox.convertToRestId(
                      Office.context.mailbox.item.itemId,
                      Office.MailboxEnums.RestVersion.v2_0);
                  console.log(item_rest_id);
                  let rest_url = Office.context.mailbox.restUrl +
                      "/v2.0/me/messages('" +
                      item_rest_id +
                      "')";
                  rest_url += "?expand=MultiValueExtendedProperties(filter=PropertyId eq 'StringArray {00020329-0000-0000-C000-000000000046} Name myTest')";

                  let auth_token = asyncResult.value;
                  console.log(auth_token);
                  axios.get(rest_url,
                      {
                          dataType: 'json',
                          headers: {
                              "Authorization": "Bearer " + auth_token
                          }
                      }
                  ).then(
                      item => {
                          console.log(item.data.MultiValueExtendedProperties[0].Value);
                          let value = item.data.MultiValueExtendedProperties[0].Value;
                          value.map((pair) => {
                              let idx = pair.indexOf("#");
                              let name = pair.substring(0, idx);
                              let value = pair.substring(idx + 1);
                              let itemAdded = { "name": name, "value": value };
                              this.setState({ keys: [...this.state.keys, itemAdded] });
                          });
                          console.log(this.state.keys);
                      }
                      
                  ).catch(
                      err => { console.log(JSON.stringify(err.message)) }
                  );

              } else {
                  console.log(JSON.stringify(asyncResult));
              }
          }
      );
      
  }

  addKeyValue = (pair) => {
      this.setState({ keys: [...this.state.keys, pair] });
      console.log(this.state.keys);
  }

  switchMode = () => { this.setState({ isReadMode: !this.state.isReadMode }) }

  render() {
      const { title, isOfficeInitialized } = this.props;
    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo={require("./../../../assets/logo-filled.png")}
          message="Please sideload your addin to see app body."
        />
      );
    }

    return (
      <div className="ms-welcome">
      <Header logo={require("./../../../assets/logo-filled.png")} title={this.props.title} message="Welcome" />
            <div className="box">
                <SwitchLabels switchMode={this.switchMode} />
                {
                    this.state.isReadMode
                        ?
                    <div>
                            <Button variant="contained" onClick={this.getCustomProps}>Get Custom Properties</Button>
                        <div className="table">
                            <CustomizedTables keys={this.state.keys} />
                        </div>
                    </div>
                        :
                    <div>
                        <InputForm addKeyValue={this.addKeyValue} />
                        <div className="table">
                            <CustomizedTables keys={this.state.keys} />
                        </div>
                        <Button size="small" variant="contained" onClick={this.uploadCustomProps}>Set Custom Properties</Button>
                    </div>
                }
            </div>
      </div>
    );
  }
}

App.propTypes = {
  title: PropTypes.string,
  isOfficeInitialized: PropTypes.bool,
};
