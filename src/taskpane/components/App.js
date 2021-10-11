import * as React from "react";
import PropTypes from "prop-types";
import { DefaultButton } from "@fluentui/react";
import Header from "./Header";
import HeroList from "./HeroList";
import Progress from "./Progress";
// images references in the manifest
import "../../../assets/icon-16.png";
import "../../../assets/icon-32.png";
import "../../../assets/icon-80.png";

/* global require */

export default class App extends React.Component {
  constructor(props, context) {
    super(props, context);
    this.state = {
        listItems: [],
        customProps: null,
        item: null,
        item_holder: null
    };
  }

  componentDidMount() {
    this.setState({
      listItems: [
        {
          icon: "Ribbon",
          primaryText: "Achieve more with Office integration",
        },
        {
          icon: "Unlock",
          primaryText: "Unlock features and functionality",
        },
        {
          icon: "Design",
          primaryText: "Create and visualize like a pro",
        },
      ],
    });

  }

    componentDidUpdate() {
  
    }

  click = async () => {
    /**
     * Insert your Outlook code here
     */
      this.setState({item : Office.context.mailbox.item});
    /*console.log(this.state.item);*/
      if (this.state.item != null) {
          console.log(this.state.item);
      }
      Office.context.mailbox.item.loadCustomPropertiesAsync(
          (asyncResult) => {
              if (asyncResult.status == "failed") {
                  console.log("Failed to load custom property");
              }
              else {
                  var customProps = asyncResult.value;
                  var myProp = customProps.get("RFI Number");
                  if (myProp != null) {
                      console.log(myProp);
                      this.setState({ customProps: "RFI Number: " + myProp });
                  } else {
                      this.setState({ customProps: null});
                  }
              }
          }
      );
      
  };

  getProps = async () => {

  }

  setProps = async () => {
      Office.context.mailbox.item.loadCustomPropertiesAsync(
          function customPropsCallback(asyncResult) {
              if (asyncResult.status == "failed") {
                  console.log("Failed to load custom property");
              }
              else {
                  var customProps = asyncResult.value;
                  customProps.set("RFI Number", "123456");
                  customProps.saveAsync(
                      function (asyncResult) {
                          if (asyncResult.status == "failed") {
                              console.log("Failed to save custom property");
                          }
                          else {
                              console.log("Saved custom property");
                          }
                      }
                  );
              }
          }
      );
  }

  deleteProps = async () => {
      Office.context.mailbox.item.loadCustomPropertiesAsync(
          function customPropsCallback(asyncResult) {
              if (asyncResult.status == "failed") {
                  console.log("Failed to load custom property");
              }
              else {
                  var customProps = asyncResult.value;
                  customProps.set("RFI Number", null);
                  customProps.saveAsync(
                      function (asyncResult) {
                          if (asyncResult.status == "failed") {
                              console.log("Failed to save custom property");
                          }
                          else {
                              console.log("Saved custom property");
                          }
                      }
                  );
              }
          }
      );   
  }

  checkAllProps = async () => {
      Office.context.mailbox.item.loadCustomPropertiesAsync(
          (asyncResult) => {
              if (asyncResult.status == "failed") {
                  console.log("Failed to load custom property");
              }
              else {
                  var customProps = asyncResult.value;
                  console.log(customProps.getAll());
              }
          }
      );
  }



  render() {
      const { title, isOfficeInitialized } = this.props;
      const { item, customProps } = this.state;

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
            <HeroList message="Discover what Office Add-ins can do for you today!" items={this.state.listItems}>

            <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={this.click}>
                    Get Email Information
            </DefaultButton>
                <br/>
                <DefaultButton onClick={this.setProps}>Set Custom Information</DefaultButton>
                <br/>
                <DefaultButton onClick={this.deleteProps}>Delete Custom Information</DefaultButton>
                <br/>
                <DefaultButton onClick={this.checkAllProps}>Check All Custom Information</DefaultButton>

                <p>{item === null
                    ?
                    "email subject is null"
                    :
                    item.subject}
                </p>

                <p>{customProps === null
                    ?
                    "customProps is null"
                    :
                    customProps}
                </p>
            
           
        </HeroList>
      </div>
    );
  }
}

App.propTypes = {
  title: PropTypes.string,
  isOfficeInitialized: PropTypes.bool,
};
