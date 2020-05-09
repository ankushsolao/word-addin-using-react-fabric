import axios from 'axios';

export class HttpService {
    private url: string;

    constructor() {
        this.url = "https://reqres.in/api/";

    }
    getUsersList() {
        let resp = axios.get(this.url + "users?page=2");
        return resp;
    }

    loginService(loginData: any) {
        let resp = axios.post(this.url + "login", loginData, {
            headers: {
                'Content-Type': 'application/json'
            }
        }
        );
        return resp;
    }

}