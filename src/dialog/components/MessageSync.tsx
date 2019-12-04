import * as React from "react";
import { Button, ButtonType } from "office-ui-fabric-react";
import { DialogPopUp } from "../../dialogpopup/dialogpopup";

export interface MessageSyncProps {
  title: string;
  isOfficeInitialized: boolean;
}

export default class MessageSync extends React.Component<MessageSyncProps, null> {
  private dialog: DialogPopUp;

  constructor(props, context) {
    super(props, context);
    this.dialog = new DialogPopUp();
  }

  componentDidMount() { }

  closeClick = async () => {
    console.log("Clear the storage value")
    // sets an item in local storage
    localStorage.setItem("CLOSE ME", "");
  };

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
            className="ms-message-parent__action"
            buttonType={ButtonType.hero}
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.click}>
            Message Parent
          </Button>
          <Button
            className="ms-clear-storage__action"
            buttonType={ButtonType.hero}
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.closeClick}>
            Clear Storage
          </Button>
      </div>
    );
  }
}
