import * as React from "react";
import PropTypes from "prop-types";
import axios from 'axios';

import Header from "./Header";
import SwitchLabels from "./SwitchLabels";
import Progress from "./Progress";
import Button from '@mui/material/Button';
import ProjectInfoCard from "./ProjectInfoCard";
import Typography from "@mui/material/Typography";
import Box from "@mui/material/Box";
import ArrowDropDownIcon from '@mui/icons-material/ArrowDropDown';
import ArrowDropUpIcon from '@mui/icons-material/ArrowDropUp';
import ProjectInfoSettings from './ProjectInfoSettings';
import EventInfo from './EventInfo';

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
        projectList: [],
        eventList: [],
        customProps: null,
        draftFloderId: null,
        itemId: null,
        authToken: null,
        uploadCustomProps: null,
        val: null,
        isFavoriteMenueExpaned: false,
        isReportingMenueExpaned: false,
    };
  }

  componentDidMount() {
      this.fetchDraftFolderId();
      setTimeout(this.fetchProjectList, 500);
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
      //this.saveMailItem();
      //setTimeout(this.clickGetMailInfo, 500);
      //setTimeout(this.upload, 1000);
  }

  //upload = () => {
  //    const { itemId, authToken } = this.state;
  //    console.log(this.state.keys);
  //    let keyValPair = this.state.keys.map((pair) => pair.name + "#" + pair.value);
  //    console.log(keyValPair);
  //    let rest_url = "https://outlook.office365.com/api/v2.0/me/messages('" + itemId + "')";
  //    axios.patch(rest_url,
  //        {
  //            "MultiValueExtendedProperties": [
  //                {
  //                    "PropertyId": "StringArray {00020329-0000-0000-C000-000000000046} Name myTest",
  //                    "Value": keyValPair
  //                }
  //            ]
  //        },
  //        {
  //            dataType: 'json',
  //            headers: {
  //                "Authorization": "Bearer " + authToken
  //            }
  //        }
  //    ).then(
  //        item => {
  //            console.log(item);
  //            this.setState({ status: item.status });
  //        }
  //    ).catch(
  //        err => { console.log(JSON.stringify(err.message)) }
  //    );
  //}

  upload = (data) => {
      const { itemId, authToken } = this.state;
      let rest_url = "https://outlook.office365.com/api/v2.0/me/messages('" + itemId + "')";
      axios.patch(rest_url,
          {
              "SingleValueExtendedProperties": [
                  {
                      "PropertyId": "String {00020329-0000-0000-C000-000000000046} Name myTest",
                      "Value": data
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

  //clickGetMailInfo = async () => {
  //    Office.context.mailbox.getCallbackTokenAsync(
  //        {
  //            isRest: true
  //        },
  //        (asyncResult) => {
  //            let id = Office.context.mailbox.item.itemId;
  //            this.setState({ itemId: id });
  //            console.log(id);
  //            let token = asyncResult.value;
  //            this.setState({ authToken: token });
  //        }
  //    );

  //    Office.context.mailbox.item.getItemIdAsync((result) => {
  //        if (result.status !== Office.AsyncResultStatus.Succeeded) {
  //            console.error(`getItemIdAsync failed with message: ${result.error.message}`);
  //        } else {
  //            console.log(result.value);
  //            let id = result.value;
  //            console.log(id.replace("-", "/"));
  //            this.setState({ itemId: id });
  //        }
  //    });
  //}

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
                  axios.patch(rest_url,
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

  projectListAdd = (key, val, color) => {
      this.setState({ projectList: [...this.state.projectList, key + '-' + val + '_' + color] })
      this.updateProjectList();
      this.fetchProjectList();
  }

  switchMode = () => { this.setState({ isReadMode: !this.state.isReadMode }) }

  getProps = () => {
      Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
              console.log("Loaded following custom properties:");
              this.setState({ customProps: result.value})
              let customProps = result.value;
              var dataKey = Object.keys(customProps)[0];
              var data = customProps[dataKey];
              for (var propertyName in data) {
                  var propertyValue = data[propertyName];
                  console.log(`${propertyName}: ${propertyValue}`);
                  this.setState({ val: propertyValue });
              }
          } else {
              console.error(`loadCustomPropertiesAsync failed with message ${result.error.message}`);
          }
      });
    }

  createNewMessage = async (projectInfo) => {
      Office.context.mailbox.getCallbackTokenAsync(
          {
              isRest: true
          },
          (asyncResult) => {
              console.log(asyncResult.value);
              if (asyncResult.status === Office.AsyncResultStatus.Succeeded
                  && asyncResult.value !== "") {
                  let auth_token = asyncResult.value;
                  this.setState({ authToken: auth_token });
                  let rest_url = Office.context.mailbox.restUrl + "/v2.0/me/messages";

                  axios.post(rest_url,
                      {},
                      {
                          dataType: 'json',
                          headers: {
                              "Authorization": "Bearer " + auth_token
                          }
                      }
                  ).then(
                      item => {
                          this.setState({ itemId: item.data.Id})
                          this.openMessageByID();
                          this.upload(projectInfo);
                          this.upload(projectInfo);
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

  openMessageByID = () => {
    Office.context.mailbox.displayMessageForm(this.state.itemId);
  }

  createAppointment = (projectInfo) => {
    var start = new Date();
    var end = new Date();
    end.setHours(start.getHours() + 1);
    
    Office.context.mailbox.displayNewAppointmentForm({
        requiredAttendees: [],
        optionalAttendees: [],
        start: start,
        end: end,
        location: "TBD",
        subject: projectInfo,
        body: ""
    });
  }


  switchIsFavoriteManueExpanded = () => {
      console.log(this.state.isFavoriteMenueExpaned);
      this.setState({ isFavoriteMenueExpaned: !this.state.isFavoriteMenueExpaned });
  }

  switchIsReportingManueExpanded = () => {
      console.log(this.state.isReportingMenueExpaned);
      this.setState({ isReportingMenueExpaned: !this.state.isReportingMenueExpaned });
  }

  fetchDraftFolderId = () => {
      Office.context.mailbox.getCallbackTokenAsync(
        {
            isRest: true
        },
        (asyncResult) => {
            console.log(asyncResult.value);
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded
                && asyncResult.value !== "") {
                let auth_token = asyncResult.value;
                this.setState({ authToken: auth_token });
                console.log(auth_token);
                let rest_url = Office.context.mailbox.restUrl + "/beta/me/mailfolders/drafts";

                axios.get(rest_url,
                    {
                        dataType: 'json',
                        headers: {
                            "Authorization": "Bearer " + auth_token
                        }
                    }
                ).then(
                    item => {
                        console.log(item.id);
                        this.setState({ draftFloderId: item.data.Id });
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

  fetchProjectList = async () => {
      Office.context.mailbox.getCallbackTokenAsync(
          {
              isRest: true
          },
          (asyncResult) => {
              console.log(asyncResult.value);
              if (asyncResult.status === Office.AsyncResultStatus.Succeeded
                  && asyncResult.value !== "") {
                  let auth_token = asyncResult.value;
                  this.setState({ authToken: auth_token });

                  let rest_url = Office.context.mailbox.restUrl +
                      "/v2.0/me/mailFolders('" +
                      this.state.draftFloderId +
                      "')";
                  rest_url += "?expand=MultiValueExtendedProperties(filter=PropertyId eq 'StringArray {00020329-0000-0000-C000-000000000046} Name myTest')";

                  axios.get(rest_url,
                      {
                          dataType: 'json',
                          headers: {
                              "Authorization": "Bearer " + auth_token
                          }
                      }
                  ).then(
                      item => {
                          console.log(item.data);
                          if (item.data.MultiValueExtendedProperties != null) {
                              let events = item.data.MultiValueExtendedProperties[0].Value;
                              let eventsListUpdate = events.map(
                                  (event) => {
                                      let eventArr = event.split("_");
                                      return { 'name': eventArr[0], 'value': 0 };
                                  }
                              );
                              eventsListUpdate = [...eventsListUpdate, { 'name': 'others', 'value': 0}];
                              this.setState({ projectList: events });
                              this.setState({ eventList: eventsListUpdate });
                              console.log(this.state.eventList);
                          }
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

  updateProjectList = async () => {
    Office.context.mailbox.getCallbackTokenAsync(
        {
            isRest: true
        },
        (asyncResult) => {
            console.log(asyncResult.value);
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded
                && asyncResult.value !== "") {
                let auth_token = asyncResult.value;
                this.setState({ authToken: auth_token });

                let rest_url = Office.context.mailbox.restUrl +
                    "/v2.0/me/mailFolders('" +
                    this.state.draftFloderId +
                    "')";
                rest_url += "?expand=MultiValueExtendedProperties(filter=PropertyId eq 'StringArray {00020329-0000-0000-C000-000000000046} Name myTest')";

                axios.patch(rest_url,
                    {
                        "MultiValueExtendedProperties": [
                            {
                                "PropertyId": "StringArray {00020329-0000-0000-C000-000000000046} Name myTest",
                                "Value": this.state.projectList
                            }
                        ]
                    },
                    {
                        dataType: 'json',
                        headers: {
                            "Authorization": "Bearer " + auth_token
                        }
                    }
                ).then(
                    item => {
                        console.log(item.data);
                        this.setState({ projectList: item.data.MultiValueExtendedProperties[0].Value });
                    }

                ).catch(
                    err => { console.log(JSON.stringify(err.message)) }
                );

            } else {
                console.log(JSON.stringify(asyncResult));
            }
        }
    );
  };

  deleteProject = (projectInfo) => {
    console.log(projectInfo);
    let newProjectList = this.state.projectList.filter(curProject => curProject !== projectInfo);
    console.log(newProjectList);
    this.setState({ projectList: newProjectList });
    this.updateProjectList();
    setTimeout(this.fetchProjectList, 1000);
  }

  clearEventList = () => {
      let eventListUpdate = [];
      this.state.eventList.forEach(event => {
          eventListUpdate.push({ 'name':event.name, 'value':0});
      })
      this.setState({ eventList: eventListUpdate});
  }


  findAllEvents = (startDate, endDate) => {
    this.clearEventList();
    console.log(startDate);
    console.log(endDate);
    let restUrl = Office.context.mailbox.restUrl + "/v1.0/me/calendarview?startdatetime=" + startDate + "&enddatetime=" + endDate;
    axios.get(
        restUrl,
        {
            dataType: 'json',
            headers: {
                "Authorization": "Bearer " + this.state.authToken
            }
        }
    ).then(
        item => {
            console.log(item);
            let eventListUpdate = [];
            this.state.eventList.forEach(event => { eventListUpdate.push(event) });
            item.data.value.forEach(
                (meeting) => {
                    let startArr = meeting.Start.split('T')[1].substring(0, 8).split(':');
                    let endArr = meeting.End.split('T')[1].substring(0, 8).split(':');
                    let duration = (parseInt(endArr[0]) - parseInt(startArr[0])) * 60 + (parseInt(endArr[1]) - parseInt(startArr[1]));
                    let meetingSubjectArr = meeting.Subject.split('_');
                    let isFound = false;
                    if (meetingSubjectArr.length > 0) {
                        let projectName = meetingSubjectArr[0];
                        let tmp = [];
                        for (let i = 0; i < eventListUpdate.length; i++) {
                            if (eventListUpdate[i].name === projectName) {
                                tmp.push({ 'name': eventListUpdate[i].name, 'value': eventListUpdate[i].value + duration });
                                isFound = true;
                            } else if (!isFound && eventListUpdate[i].name === 'others') {
                                tmp.push({ 'name': eventListUpdate[i].name, 'value': eventListUpdate[i].value + duration });
                            } else {
                                tmp.push(eventListUpdate[i]);
                            }
                        }
                        eventListUpdate = tmp;
                    }
                }
            );
            this.setState({ eventList: eventListUpdate });
        }
    ).catch(
        err => { console.log(JSON.stringify(err.message)) }
    );
  }

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
            <div className="box" background="lightgrey">
                <Box sx={{position: 'relative', width: 300}}>
                    <Box sx={{ backgroundColor: "black", display: 'flex', flexDirection: 'row', position: 'relative'}}>
                        <Typography sx={{ variant: "body1", color: "white", pl: 2}}>
                                FAVORITES
                        </Typography>
                        <ProjectInfoSettings addProjectInfo={this.projectListAdd }/>
                        {
                            this.state.isFavoriteMenueExpaned 
                                    ?
                            <ArrowDropUpIcon sx={{ color: 'white', position: 'absolute', right: 40 }}
                                        onClick={this.switchIsFavoriteManueExpanded} />
                                    :
                            <ArrowDropDownIcon sx={{ color: 'white', position: 'absolute', right: 40 }}
                                        onClick={this.switchIsFavoriteManueExpanded} />

                        }
                        
                    </Box>

                        {
                            this.state.isFavoriteMenueExpaned
                                ?
                                this.state.projectList.map(
                                    (project) => {
                                        let projectArr = project.split("_");
                                        return <ProjectInfoCard sx={{ mt: 30 }}
                                            key={projectArr[0]}
                                            createNewMessage={this.createNewMessage}
                                            createAppointment={this.createAppointment}
                                            projectInfo={projectArr[0]}
                                            deleteProject={this.deleteProject}
                                            color={projectArr[1]} />
                                    }
                                )
                                :
                               null
                        }

                    <Box sx={{ mt:3, backgroundColor: "black", display: 'flex', flexDirection: 'row', position: 'relative' }}>
                        <Typography sx={{ variant: "body1", color: "white", pl: 2 }}>
                                REPORTING   
                        </Typography>
                        {
                          this.state.isReportingMenueExpaned
                                ?
                                <ArrowDropUpIcon sx={{ color: 'white', position: 'absolute', right: 40 }}
                                    onClick={this.switchIsReportingManueExpanded} />
                                :
                                <ArrowDropDownIcon sx={{ color: 'white', position: 'absolute', right: 40 }}
                                    onClick={this.switchIsReportingManueExpanded} />
                        }
                    </Box>
                        {
                            this.state.isReportingMenueExpaned
                                ?
                            <Box sx={{ mt: 3 }}>
                                <EventInfo findAllEvents={this.findAllEvents} events={this.state.eventList} />
                            </Box>
                            :
                            null

                    }
                </Box>
            </div>
      </div>
    );
  }
}

App.propTypes = {
  title: PropTypes.string,
  isOfficeInitialized: PropTypes.bool,
};
