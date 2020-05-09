import * as React from "react";
import Login from "./Login"; 
 
export default class App extends React.Component<any, any> {   
  render() {    
    return (
      <div className="ms-welcome">
        <Login />        
      </div>
    );
  }
}

