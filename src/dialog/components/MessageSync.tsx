import * as React from "react";
import { Button, ButtonType } from "office-ui-fabric-react";
import { DialogPopUp } from "../../dialogpopup/dialogpopup";

export interface MessageSyncProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface MessageSyncState {
  listItems: any[];
}

export default class MessageSync extends React.Component<MessageSyncProps, MessageSyncState> {
  private dialog: DialogPopUp;

  constructor(props, context) {
    super(props, context);
    this.dialog = new DialogPopUp();

    this.state = {
      listItems: []
    };
  }

  componentDidMount() {
    this.setState({
      listItems: [
        {
          icon: "Ribbon",
          primaryText: "Achieve more with Office integration"
        },
        {
          icon: "Unlock",
          primaryText: "Unlock features and functionality"
        },
        {
          icon: "Design",
          primaryText: "Create and visualize like a pro"
        }
      ]
    });  
  }

  click = async () => {
    console.log("Message Parent has been clicked")
    // sets an item in local storage
    localStorage.setItem("CLOSE ME", "Will the parent get this close value?");

    this.dialog.messageParent("Close#");
  };

  render() {
    return (
      <div className="ms-welcome">
          <p className="ms-font-l">
            Run this button to message the parent. Click <b>Message Parent</b>.
          </p>   
          <Button
            className="ms-welcome__action"
            buttonType={ButtonType.hero}
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.click}>
            Message Parent
          </Button>
      </div>
    );
  }
}
