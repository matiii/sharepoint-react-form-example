/// <reference path="../../node_modules/@types/sharepoint/index.d.ts" />

import { IPersonaProps } from 'office-ui-fabric-react/lib/Persona';
import { Promise } from 'es6-promise';
import axios from 'axios';
import * as Response from '../model/Response';
import { User, PersonaUser } from '../model/User';
import * as Endpoints from './Endpoints';

const getPeoplesByStringEnvelope = (text: string, maxResults: number) => {
    return (
        `<?xml version="1.0" encoding="utf-8"?>
         <soap12:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap12="http://www.w3.org/2003/05/soap-envelope">
         <soap12:Body>
         <SearchPrincipals xmlns="http://schemas.microsoft.com/sharepoint/soap/">
            <searchText>${text}</searchText>
            <maxResults>${maxResults}</maxResults>
            <principalType>User</principalType>
         </SearchPrincipals>
         </soap12:Body>
         </soap12:Envelope>`
    );
};

export default class PeopleService {

    getPeoplesByString(text: string, maxResults: number): Promise<(IPersonaProps & { key: string | number })[]> {
        
        return axios.post(Endpoints.PeopleEndpoint,
            getPeoplesByStringEnvelope(text, maxResults),
            {
                headers: { 'Content-Type': 'application/soap+xml; charset=utf-8'}
            })
            .then(value => {

            let xml = value.data as string;
            let parser = new DOMParser();

            let xmlDoc = parser.parseFromString(xml, 'text/xml');
            let result: (IPersonaProps & { key: string | number })[] = [];
            let principals = xmlDoc.getElementsByTagName('PrincipalInfo');

            for (var i = 0; i < principals.length; i++) {
                let principal = principals[i];

                let accountName = principal.getElementsByTagName('AccountName')[0].textContent;
                let displayName = principal.getElementsByTagName('DisplayName')[0].textContent;

                let title = '';

                let titleArray = principal.getElementsByTagName('Title');

                if (titleArray.length > 0) {
                    title = titleArray[0].textContent;
                }

                let names = displayName.split(' ');

                let firstLetter = names.length > 0 && names[0] != undefined ? names[0][0] : '';
                let secondLetter = names.length > 1 && names[1] != undefined ? names[1][0] : '';

                result.push({
                    key: accountName,
                    imageUrl: `/_layouts/14/userphoto.aspx?size=L&accountname=${accountName}`,
                    imageInitials: firstLetter + secondLetter,
                    primaryText: displayName,
                    secondaryText: title
                });
            }

            return result;
        });
    }

    getCurrentUser(success: (user: User) => void, error: (msg: string) => void): void {

        let callback = () => {

            let clientContext = SP.ClientContext.get_current();
            let web = clientContext.get_web();
            let currentUser = web.get_currentUser();
            clientContext.load(currentUser);

            clientContext.executeQueryAsync(successHandler, errorHandler);

            function successHandler(sender: any, args: SP.ClientRequestSucceededEventArgs) {

                let user: User = {
                    userId: currentUser.get_id(),
                    email: currentUser.get_email(),
                    title: currentUser.get_title(),
                    loginName: currentUser.get_loginName()
                };

                success(user);
            }

            function errorHandler(sender: any, args: SP.ClientRequestFailedEventArgs) {
                console.log(args);
                console.log(args.get_message());
                console.log(args.get_errorDetails());
                console.log(args.get_stackTrace());
                error(args.get_message());
            }
        }

        ExecuteOrDelayUntilScriptLoaded(callback, "sp.js");
    }

    getCurrentPersonaUser(success: (user: PersonaUser) => void, error: (msg: string) => void) {

        this.getCurrentUser(user => {

            this.getPeoplesByString(user.email, 1)
                .then(result => {

                    if (result.length === 0) {
                        error(`Cannot find user by email: ${user.email}`);
                    } else {

                        let up = result[0];

                        let personaUser: PersonaUser = {
                            userId: user.userId,
                            email: user.email,
                            title: user.title,
                            loginName: user.loginName,
                            key: up.key,
                            imageUrl: up.imageUrl,
                            imageInitials: up.imageInitials,
                            primaryText: up.primaryText,
                            secondaryText: up.title
                        };

                        success(personaUser);
                    }
                })
                .catch(reason => {
                    console.log(reason);
                    error(JSON.stringify(reason));
                });

        }, error);

    }
}