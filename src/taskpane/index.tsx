import "office-ui-fabric-react/dist/css/fabric.min.css";
import App from "./components/App";
//import Login from "./components/Login";
//import UsersList from "./components/UsersList";
import { initializeIcons } from "office-ui-fabric-react/lib/Icons";
import * as React from "react";
import * as ReactDOM from "react-dom";
//import { Redirect, Route, Switch } from "react-router";
//import { HashRouter as Router,  Switch, Route } from "react-router-dom";
//import UsersList from './components/UsersList';
/* global AppCpntainer, Component, document, Office, module, require */
{/* <Router>
  <Switch>
    <Route exact path="/UsersList" component={UsersList} >

    </Route>
  </Switch>
</Router> */}
initializeIcons();

let isOfficeInitialized = false;

const title = "Contoso Task Pane Add-in";

const render = Component => {
  ReactDOM.render(
    <Component title={title} isOfficeInitialized={isOfficeInitialized} />,
    document.getElementById("container")
  );
};

/* Render application after Office initializes */
Office.initialize = () => {
  isOfficeInitialized = true;
  render(App);
};

/* Initial render showing a progress bar */
render(App);

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    render(NextApp);
  });
}
