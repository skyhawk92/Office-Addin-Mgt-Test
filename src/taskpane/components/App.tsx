import '../../../node_modules/@webcomponents/webcomponentsjs/webcomponents-bundle.js';

import {Providers, MsalProvider, LoginType} from '@microsoft/mgt';
import { PeoplePicker, Login } from '@microsoft/mgt-react';

import 'react-app-polyfill/ie11';
import 'react-app-polyfill/stable';

import * as React from "react";
import { Button, ButtonType } from "office-ui-fabric-react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";
import * as excel from "./Excel.App";
import * as onenote from "./OneNote.App";
import * as outlook from "./Outlook.App";
import * as powerpoint from "./PowerPoint.App";
import * as project from "./Project.App";
import * as word from "./Word.App";
import axios from 'axios';
import qs = require('qs');
/* global Button, Header, HeroList, HeroListItem, Office */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  listItems: HeroListItem[];
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: []
    };
    Providers.globalProvider = new MsalProvider({ 
      clientId: '9b154016-8b25-4923-b167-a2ac52c68017',
      loginType: LoginType.Popup,
      authority: 'https://login.microsoftonline.com/6785298f-e857-464b-9b4b-807178402632',
      redirectUri: "https://localhost:3000"
    });
  }

  async componentDidMount() {

    const url = 'https://login.microsoftonline.com/6785298f-e857-464b-9b4b-807178402632/oauth2/v2.0/token';
    const data = qs.stringify({
      grant_type: 'client_credentials',
      client_id: '9b154016-8b25-4923-b167-a2ac52c68017',
      client_secret: '',
      scope: 'https://graph.microsoft.com/.default'
    });
     const [response] = await Promise.all([
      axios.post(url,
        data,
        {
          headers: {
            'content-type': 'application/x-www-form-urlencoded'
          },
      })
    ]);
    console.log("RESPONSE:", response);

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
    switch (Office.context.host) {
      case Office.HostType.Excel: {
        const excelApp = new excel.default(this.props, this.context);
        return excelApp.click();
      }
      case Office.HostType.OneNote: {
        const onenoteApp = new onenote.default(this.props, this.context);
        return onenoteApp.click();
      }
      case Office.HostType.Outlook: {
        const outlookApp = new outlook.default(this.props, this.context);
        return outlookApp.click();
      }
      case Office.HostType.PowerPoint: {
        const powerpointApp = new powerpoint.default(this.props, this.context);
        return powerpointApp.click();
      }
      case Office.HostType.Project: {
        const projectApp = new project.default(this.props, this.context);
        return projectApp.click();
      }
      case Office.HostType.Word: {
        const wordApp = new word.default(this.props, this.context);
        return wordApp.click();
      }
    }
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
          <div className="login">
              <Login loginCompleted={(_e) => console.log('Logged in')} />
            </div>
            <PeoplePicker />
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
