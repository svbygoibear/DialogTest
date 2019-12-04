import * as React from "react";
import { Button, ButtonType } from "office-ui-fabric-react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";
import { DialogPopUp } from "../../dialogpopup/dialogpopup";
/* global Button, Header, HeroList, HeroListItem, Progress */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  listItems: HeroListItem[];
}

export default class App extends React.Component<AppProps, AppState> {
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
          primaryText: "Test why the heck we can't share local storage!"
        },
      ]
    });
    // workaround to see if we can access local storage
    window.addEventListener("storage", e => {
      console.log("key:"+e.key+", value:"+e.newValue);
    }, false);
  }

  click = async () => {
    const dialogOptions = {
      inFrame: false,
      heightPercent: 85,
      widthPercent: 93,
      route: `${window.location.origin + "/dialog.html"}`
    };

    // sets an item in local storage
    localStorage.setItem("This is a test key", "This is a test value");

    this.dialog.showDialog(
      dialogOptions,
      (dialog: any, message: string) => {
          if (message.indexOf("Close#") === 0) {
              dialog.close();
          }
      }
    );
  };

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress title={title} logo="assets/logo-filled.png" message="Please sideload your addin to see app body." />
      );
    }

    return (
      <div className="ms-welcome">
        <Header logo="assets/logo-filled.png" title={this.props.title} message="Welcome" />
        <HeroList message="Discover what Office Add-ins can do for you today!" items={this.state.listItems}>
          <p className="ms-font-l">
            Modify the source files, then click <b>Run</b>.
          </p>
          <Button
            className="ms-welcome__action"
            buttonType={ButtonType.hero}
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.click}
          >
            Run
          </Button>
        </HeroList>
      </div>
    );
  }
}
