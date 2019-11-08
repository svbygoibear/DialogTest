import * as React from "react";
/* global Button, Header, HeroList, HeroListItem, Progress */

export interface MessageSyncProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface MessageSyncState {
  listItems: any[];
}

export default class MessageSync extends React.Component<MessageSyncProps, MessageSyncState> {
  constructor(props, context) {
    super(props, context);
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
    localStorage.setItem("UWU", "WHAT is UWU anyway");
    Office.context.ui.displayDialogAsync("https://localhost:3000/",
    {
        height: 50,
        width: 50,
        displayInIframe: false
    }, (asyncResult: Office.AsyncResult<any>) => { 
        const dialog = asyncResult.value;
        console.log(dialog);
    });
  };

  render() {
    return (
      <div className="ms-welcome">
        TEST
      </div>
    );
  }
}
