import * as React from 'react';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { DatePicker, DayOfWeek } from 'office-ui-fabric-react/lib/DatePicker';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import * as update from 'immutability-helper';
import { getListItemService, getPeopleService } from '../api/bootstrap';
import { PeoplePicker } from '../components/PeoplePicker';
import { FileUploader } from '../components/FileUploader';
import { IPersonaProps } from 'office-ui-fabric-react/lib/Persona';
import { BaseComponent, autobind } from 'office-ui-fabric-react/lib/Utilities';
import {
    Spinner,
    SpinnerSize
    } from 'office-ui-fabric-react/lib/Spinner';

export interface AppState {
    name: string;
    date: Date;
    users: IPersonaProps[];
    defaultUsers: IPersonaProps[];
    appIsLoading: boolean;
}

const ListItemService = getListItemService();
const PeopleService = getPeopleService();

export class App extends React.Component<{}, AppState>{

    private readonly listApi = new ListItemService();
    private readonly peopleService = new PeopleService();

    constructor(props, context) {
        super(props, context);
        this.state = { name: '', date: new Date(), users: [], defaultUsers: [], appIsLoading: true };
    }

    componentDidMount() {
        this.peopleService.getCurrentPersonaUser(user => {
            this.setState(prev => update(prev, { defaultUsers: { $set: [user] } }));

            this.setState(prev => update(prev, { appIsLoading: { $set: false }}));
        }, msg => {
            
        });
    }

    setName = (name: any) => {
        let newState = update(this.state, { name: { $set: name } });
        this.setState(newState);
    }

    setDate = (date: Date) => {
        let newState = update(this.state, { date: { $set: date } });
        this.setState(newState);
    }

    submit = () => {
        this.listApi.setItem(
            'Data',
            [
            { property: 'Title', value: this.state.name },
            { property: 'Date', value: this.state.date },
            { property: 'Person', value: SP.FieldUserValue.fromUser('arkadiusz.czaplicki') }
            ],
            (listItemId: number) => { },
            (error: string) => { });
    }

    @autobind
    private setUsers(users: IPersonaProps[]) {
        this.setState(prev => update(prev, { users: { $set: users }}));
    }

    render() {

        if (this.state.appIsLoading) {
            return (
                <Spinner size={SpinnerSize.large} label='Form is loading, please wait...' />
            );
        }

        return (
            <div>
                <TextField label='Name' onChanged={this.setName} />
                <DatePicker onSelectDate={this.setDate} placeholder='Select a date...' />
                <PeoplePicker
                    isRequired={true}
                    label='Person'
                    selected={this.setUsers}
                    defaultSelected={this.state.defaultUsers} />
                <FileUploader onUploadedFilesChange={(files) => console.log(files)} />
                <DefaultButton onClick={e => { e.preventDefault(); this.submit() }} iconProps={{ iconName: 'Add' }} text='Save' />
            </div>
        );
    }

}