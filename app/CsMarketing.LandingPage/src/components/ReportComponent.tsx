//import * as React from 'react';
//import { TextField } from 'office-ui-fabric-react/lib/TextField';
//import { DatePicker, DayOfWeek } from 'office-ui-fabric-react/lib/DatePicker';
//import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
//import * as update from 'immutability-helper';
//import { getListItemService, getPeopleService } from '../api/bootstrap';
//import { PeoplePicker } from '../components/PeoplePicker';
//import { FileUploader } from '../components/FileUploader';
//import { IPersonaProps } from 'office-ui-fabric-react/lib/Persona';
//import { BaseComponent, autobind } from 'office-ui-fabric-react/lib/Utilities';
//import {
//    Spinner,
//    SpinnerSize
//    } from 'office-ui-fabric-react/lib/Spinner';
//import strings from '../strings';

//export interface ReportComponentState {
//    appIsLoading: boolean
//}

//const ListItemService = getListItemService();
//const PeopleService = getPeopleService();

//export abstract class ReportComponent<P, S extends ReportComponentState> extends BaseComponent<P, S>{

//    protected listApi: ListIte
//    protected readonly peopleService = new PeopleService();

//    constructor(props, context) {
//        super(props, context);

//        const ListItemService = getListItemService();
//    }

//    abstract render(): JSX.Element;
//}