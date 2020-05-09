import * as React from "react";
import {
    TextField, PrimaryButton
} from "office-ui-fabric-react";
import { Link } from "valuelink/lib";
import { HttpService } from "../services/httpservice";
import { mergeStyleSets } from 'office-ui-fabric-react/lib/Styling';

//import HistoryRouter from 'history-router'


const classNames = mergeStyleSets({
    error: {
        color: "Red",
    },
});



export interface LoginProps {
    title: string;
    logo: string;
    message: string;
}
//const router = new HistoryRouter()
export interface State {
    email: any,
    password: any,
    errors: IErrors,
}

export interface IErrors {
    password: string;
    username: string;
    loginSuccess: string;
}

export default class Login extends React.Component<LoginProps, State> {
    state: State = {
        email: null,
        password: null,
        errors: {
            username: "",
            password: "",
            loginSuccess: "",

        },
    };
    valueLink: Link<any>;
    serv: HttpService;

    constructor(props: LoginProps) {
        super(props);
        this.serv = new HttpService();
    }

    validateForm = () => {
        let isDataValid = true;
        let errorsMessage = this.state.errors;
        console.log(this.state.email, this.state.password);
        if (!this.state.email) {
            errorsMessage.username = "User Name is required ";
            isDataValid = false;
        } else {
            errorsMessage.username = null;
        }

        if (!this.state.password) {
            errorsMessage.password = "Password is required ";
            isDataValid = false;
        } else {
            errorsMessage.password = null;
        }
        this.setState({
            errors: errorsMessage,
        });
        return isDataValid;
    }

    handleLogin = async () => {
        if (this.validateForm()) {
            this.serv.loginService({ email: this.state.email, password: this.state.password }).then(resp => resp.data).then(data => {
                console.log("Login :- ", data);
                //  history.push("/UsersList");
            }).catch(error => console.log("error :- ", error));

            return Word.run(async context => {
                /**
                 * Insert your Word code here
                 */

                // insert a paragraph at the end of the document.
                const paragraph = context.document.body.insertParagraph("Welcome Ankush", Word.InsertLocation.end);

                // change the paragraph color to blue.
                paragraph.font.color = "red";
                paragraph.font.bold = true;

                await context.sync();
            });
        }
    };

    handleEmailId = (evt: React.ChangeEvent<HTMLInputElement>) => {
        this.setState({ email: evt.target.value });
    };

    handlePassword = (evt: React.ChangeEvent<HTMLInputElement>) => {
        this.setState({ password: evt.target.value });
    }

    render() {
        return (
            <div className="ms-Grid" dir="ltr">
                <div className="ms-Grid-row">
                    <h3>User Login </h3>
                    <div className="ms-Grid-col ms-sm6 ms-md12 ms-lg12">
                        <TextField type='email' label='User Name' value={this.state.email} onChange={this.handleEmailId} />
                        <span className={classNames.error}>{this.state.errors.username}</span>
                        <TextField type='password' label='Password' value={this.state.password} onChange={this.handlePassword} />
                        <span className={classNames.error}>{this.state.errors.password}</span>
                        <hr />
                        <PrimaryButton text="Login" allowDisabledFocus onClick={this.handleLogin} />

                    </div>
                </div>
            </div>
        );
    }
}
