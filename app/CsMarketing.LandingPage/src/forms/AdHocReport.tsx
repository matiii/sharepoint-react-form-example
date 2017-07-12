import * as React from 'react';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { DatePicker, DayOfWeek } from 'office-ui-fabric-react/lib/DatePicker';
import * as update from 'immutability-helper';
import { getListItemService, getPeopleService } from '../api/bootstrap';
import { PeoplePicker } from '../components/PeoplePicker';
import { FileUploader } from '../components/FileUploader';
import { IPersonaProps } from 'office-ui-fabric-react/lib/Persona';
import { BaseComponent, autobind } from 'office-ui-fabric-react/lib/Utilities';
import strings from '../strings';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import {
    Checkbox
} from 'office-ui-fabric-react/lib/Checkbox';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { DefaultButton, PrimaryButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import CsSpinner from '../components/CsSpinner';
import { RadioChoice } from '../components/RadioChoice';
import { CheckboxChoice } from '../components/CheckboxChoice';
import { PersonaUser } from '../model/User';

export interface AdHocReportState {
    appIsLoading: boolean;
    formIsSaving: boolean;

    errors: string[];

    requestor: IPersonaProps & { key: string | number };
    projectLeader: IPersonaProps & { key: string | number };
    eventName: string;
    mapId: string;
    typeOfReport: string;
    typeOfReportIsOther: boolean;
    inputTypes: string[];
    reportingSolutions: string[];
    requestedDeliveryDate?: Date;
    kpi: string;
    targets: string;
    remarks: string;
    attachments: { file: File, base64: string }[];
    isConfirmed: boolean;
}

export interface AdHocReportProps {
    currentUser: PersonaUser
}

const ListItemService = getListItemService();
const PeopleService = getPeopleService();

const containerStyle: React.CSSProperties = {
    margin: '25px auto',
    maxWidth: '1200px',
    border: 'solid 1px #eaeaea'
}


export class AdHocReport extends BaseComponent<AdHocReportProps, AdHocReportState>{

    private readonly listApi = new ListItemService();
    private readonly peopleService = new PeopleService();

    constructor(props, context) {
        super(props, context);
        this.state = {
            errors: [],
            appIsLoading: true,
            formIsSaving: false,
            requestor: this.props.currentUser,
            projectLeader: undefined,
            eventName: '',
            mapId: '',
            typeOfReport: 'Feedback Report (Event/Activity Evaluation, Management/Staff Feedback)',
            typeOfReportIsOther: false,
            inputTypes: [],
            reportingSolutions: [],
            requestedDeliveryDate: null,
            kpi: '',
            targets: '',
            remarks: '',
            attachments: [],
            isConfirmed: false
        };
    }

    componentDidMount(): void {
        this.setState(prev => update(prev,
            {
                appIsLoading: { $set: false }
            }));

        parent.location.hash = 'app/forms/ad-hoc-report';
    }

    @autobind
    private save() {

        let forceError = false;

        if (this.state.eventName.length === 0) {
            this.setEventActivityNameError();
            forceError = true;
        }

        if (!this.state.projectLeader) {
            this.setProjectLeaderError();
            forceError = true;
        }

        if (!this.state.requestor) {
            this.setProjectRequestorError();
            forceError = true;
        }

        if (!this.state.isConfirmed) {
            this.setConfirmationDataError();
            forceError = true;
        }

        if (forceError) {
            document.getElementById('ad-hoc-errors').scrollIntoView();
            return;
        }

        this.setState(prev => update(prev, { formIsSaving: { $set: true } }), () => { console.log(this.state)});
    }

    @autobind
    private setInputTypes(inputs: string[]) {

        this.setState(prev => update(prev, { inputTypes: { $set: inputs } }));

    }

    @autobind
    private setTypeOfReport(choice: string | { isOptional: boolean, value: string }) {
        let optional = choice as { isOptional: boolean, value: string };

        if (optional.isOptional) {
            this.setState(prev => update(prev,
                { typeOfReport: { $set: optional.value }, typeOfReportIsOther: { $set: true }, inputTypes: { $set: [] } }));
        } else {
            this.setState(prev => update(prev,
                { typeOfReport: { $set: choice }, typeOfReportIsOther: { $set: false }, inputTypes: { $set: [] } }));
        }
    }

    @autobind
    private setEventActivityName(value: string) {

        this.setState(prev => update(prev, { eventName: { $set: value } }));

        if (value.length === 0) {

            this.setEventActivityNameError();
            
        } else {
            this.setState(prev => update(prev, { errors: { $set: this.state.errors.filter(x => x !== strings.eventActivityNameErrorMsg) } }));
        }
    }

    @autobind
    private setEventActivityNameError() {
        if (this.state.errors.filter(x => x === strings.eventActivityNameErrorMsg).length === 0) {
            this.setState(prev => update(prev, { errors: { $push: [strings.eventActivityNameErrorMsg] } }));
        }
    }

    @autobind
    private setProjectLeader(value: IPersonaProps) {

        this.setState(prev => update(prev, { projectLeader: { $set: value } }));

        if (!value) {

            this.setProjectLeaderError();
            
        } else {
            this.setState(prev => update(prev, { errors: { $set: this.state.errors.filter(x => x !== strings.projectLeaderErrorMsg) } }));
        }
    }

    @autobind
    private setProjectLeaderError() {
        if (this.state.errors.filter(x => x === strings.projectLeaderErrorMsg).length === 0) {
            this.setState(prev => update(prev, { errors: { $push: [strings.projectLeaderErrorMsg] } }));
        }
    }

    @autobind
    private setProjectRequestor(value: IPersonaProps) {

        this.setState(prev => update(prev, { requestor: { $set: value } }));

        if (!value) {
            this.setProjectRequestorError();
        } else {
            this.setState(prev => update(prev, { errors: { $set: this.state.errors.filter(x => x !== strings.reportRequestorErrorMsg) } }));
        }
    }

    @autobind
    private setProjectRequestorError() {
        if (this.state.errors.filter(x => x === strings.reportRequestorErrorMsg).length === 0) {
            this.setState(prev => update(prev, { errors: { $push: [strings.reportRequestorErrorMsg] } }));
        }
    }

    @autobind
    private setConfirmationData(checked: boolean) {

        this.setState(prev => update(prev, { isConfirmed: { $set: checked }}));

        if (checked) {
            this.setState(prev => update(prev, { errors: { $set: this.state.errors.filter(x => x !== strings.doNotIncludeAnyClientDataErrorMsg) } }));
        } else {
            this.setConfirmationDataError();
        }
    }

    @autobind
    private setConfirmationDataError() {
        if (this.state.errors.filter(x => x === strings.doNotIncludeAnyClientDataErrorMsg).length === 0) {
            this.setState(prev => update(prev, { errors: { $push: [strings.doNotIncludeAnyClientDataErrorMsg] } }));
        }
    }

    @autobind
    private setProjectMapId(value: string) {

        this.setState(prev => update(prev, { mapId: { $set: value }}));

    }

    @autobind
    private setReportSolution(value: string, checked: boolean) {

        if (checked) {
            let solutions = [...this.state.reportingSolutions];
            solutions.push(value);
            this.setReportingSolutions(solutions);
        } else {
            this.setReportingSolutions(this.state.reportingSolutions.filter(x => x !== value));
        }
    }

    @autobind
    private setReportingSolutions(values: string[]) {

        this.setState(prev => update(prev, { reportingSolutions: { $set: values }}));
    }

    @autobind
    private setRequestDeliveryDate(value: Date) {

        this.setState(prev => update(prev, { requestedDeliveryDate: { $set: value }}));
    }

    @autobind
    private setKpis(kpis: string) {

        this.setState(prev => update(prev, { kpi: { $set: kpis }}));
    }

    @autobind
    private setTargets(target: string) {

        this.setState(prev => update(prev, { targets: { $set: target } }));

    }

    @autobind
    private setRemarks(remarks: string) {

        this.setState(prev => update(prev, { remarks: { $set: remarks }}));

    }

    @autobind
    private setFiles(files: { file: File, base64: string }[]) {

        this.setState(prev => update(prev, { attachments: { $set: files }}));

    }

    render() {

        if (this.state.appIsLoading) {
            return (
                <CsSpinner text={strings.formIsLoadingMessage} isVisible={true} />);
        }

        if (this.state.formIsSaving) {
            return (
                <CsSpinner text={strings.formIsSavingMessage} isVisible={true} />);
        }

        return (
            <div className="ms-Grid" style={containerStyle}>

                <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-u-sm1 ms-u-md1 ms-u-lg1"></div>
                    <div className="ms-Grid-col ms-u-sm3 ms-u-md3 ms-u-lg3">
                        <img style={{ marginTop: '35px' }} src='src/content/cs-logo.gif' />
                    </div>

                    <div className="ms-Grid-col ms-u-sm8 ms-u-md8 ms-u-lg8">
                        <span style={{ display: 'block', marginTop: '15px' }} className="ms-font-xxl">Ad Hoc Report</span>
                        <span className="ms-font-xl">Order form</span>
                    </div>
                </div>

                <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-u-sm1 ms-u-md1 ms-u-lg1"></div>

                    <div className="ms-Grid-col ms-u-sm10 ms-u-md10 ms-u-lg10">
                        <div id='ad-hoc-errors' style={{ width: '750px', marginTop: '25px' }}>
                            <MessageBar
                                messageBarType={MessageBarType.warning}
                                ariaLabel='Aria help text here'>
                                It is prohibited to enter data that directly or indirectly infer the identity of a client or that would expose the existence of an existing, terminated or future client relationship.
                            </MessageBar>
                        </div>

                        {this.state.errors.length > 0 &&
                            <div style={{ marginTop: '25px' }}>

                                {this.state.errors.map((error, key) => (

                                    <MessageBar
                                        key={`${key}-${error}`}

                                        messageBarType={MessageBarType.error}
                                        ariaLabel='Aria help text here'>
                                        {error}
                                    </MessageBar>

                                ))}

                            </div>}

                    </div>

                    <div className="ms-Grid-col ms-u-sm1 ms-u-md1 ms-u-lg1"></div>
                </div>

                <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-u-sm1 ms-u-md1 ms-u-lg1">
                    </div>

                    <div className="ms-Grid-col ms-u-sm10 ms-u-md10 ms-u-lg10">
                        <div style={{ marginTop: '25px', marginBottom: '15px', backgroundColor: '#f4f4f4', padding: '10px' }}>
                            <i className="ms-Icon ms-Icon--Info" style={{ color: '#005a9e' }} aria-hidden="true"></i>
                            <span className="ms-fontWeight-light" style={{ marginLeft: '10px' }}>General information</span>
                        </div>
                    </div>

                    <div className="ms-Grid-col ms-u-sm1 ms-u-md1 ms-u-lg1">
                    </div>

                </div>

                <div className="ms-Grid-row">

                    <div className="ms-Grid-col ms-u-sm1 ms-u-md1 ms-u-lg1">
                    </div>

                    <div className="ms-Grid-col ms-u-sm5 ms-u-md5 ms-u-lg5">
                        <TextField label='Event/Activity Name' required={true} onChanged={value => this.setEventActivityName(value as string)} />
                    </div>

                    <div className="ms-Grid-col ms-u-sm5 ms-u-md5 ms-u-lg5">
                        <TextField label='Project MAP ID' onChanged={value => this.setProjectMapId(value as string)} />
                    </div>

                    <div className="ms-Grid-col ms-u-sm1 ms-u-md1 ms-u-lg1">
                    </div>

                </div>


                <div className="ms-Grid-row">

                    <div className="ms-Grid-col ms-u-sm1 ms-u-md1 ms-u-lg1">
                    </div>

                    <div className="ms-Grid-col ms-u-sm5 ms-u-md5 ms-u-lg5">
                        <PeoplePicker
                            label='Project Leader'
                            isRequired={true}
                            defaultSelected={[]}
                            selected={(users: IPersonaProps[]) => { if(users.length > 0) this.setProjectLeader(users[0]); else this.setProjectLeader(undefined); }}
                        />
                    </div>

                    <div className="ms-Grid-col ms-u-sm5 ms-u-md5 ms-u-lg5">
                        <PeoplePicker
                            label='Report Requestor'
                            isRequired={true}
                            defaultSelected={this.state.requestor ? [this.state.requestor] : []}
                            selected={(users: IPersonaProps[]) => { if (users.length > 0) this.setProjectRequestor(users[0]); else this.setProjectRequestor(undefined); }}
                        />
                    </div>

                    <div className="ms-Grid-col ms-u-sm1 ms-u-md1 ms-u-lg1">
                    </div>

                </div>

                <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-u-sm1 ms-u-md1 ms-u-lg1">
                    </div>

                    <div className="ms-Grid-col ms-u-sm10 ms-u-md10 ms-u-lg10">
                        <div style={{ marginTop: '25px', marginBottom: '15px', backgroundColor: '#f4f4f4', padding: '10px' }}>
                            <i className="ms-Icon ms-Icon--Info" style={{ color: '#005a9e' }} aria-hidden="true"></i>
                            <span className="ms-fontWeight-light" style={{ marginLeft: '10px' }}>Type of Report</span>
                        </div>
                    </div>

                    <div className="ms-Grid-col ms-u-sm1 ms-u-md1 ms-u-lg1">
                    </div>

                </div>

                <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-u-sm2 ms-u-md2 ms-u-lg2">
                    </div>

                    <div className="ms-Grid-col ms-u-sm9 ms-u-md9 ms-u-lg9">
                        <RadioChoice
                            onChoice={this.setTypeOfReport}
                            list={['Feedback Report (Event/Activity Evaluation, Management/Staff Feedback)', 'Campaign Results Analysis (Online Campaign, Mailing, Banner, Postering, etc.)', 'Survey Results Consolidation (handwritten responses, onsite surveys)']}
                            isOpenChoice={true} />
                    </div>

                    <div className="ms-Grid-col ms-u-sm1 ms-u-md1 ms-u-lg1">
                    </div>

                </div>

                <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-u-sm1 ms-u-md1 ms-u-lg1">
                    </div>

                    <div className="ms-Grid-col ms-u-sm10 ms-u-md10 ms-u-lg10">
                        <div style={{ marginTop: '25px', marginBottom: '15px', backgroundColor: '#f4f4f4', padding: '10px' }}>
                            <i className="ms-Icon ms-Icon--Info" style={{ color: '#005a9e' }} aria-hidden="true"></i>
                            <span className="ms-fontWeight-light" style={{ marginLeft: '10px' }}>Input type (Data source)</span>
                        </div>
                    </div>

                    <div className="ms-Grid-col ms-u-sm1 ms-u-md1 ms-u-lg1">
                    </div>

                </div>


                <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-u-sm2 ms-u-md2 ms-u-lg2">
                    </div>

                    <div className="ms-Grid-col ms-u-sm9 ms-u-md9 ms-u-lg9">

                        {this.state.typeOfReport === 'Feedback Report (Event/Activity Evaluation, Management/Staff Feedback)' &&
                            <CheckboxChoice
                                id='Feedback Report (Event/Activity Evaluation, Management/Staff Feedback)'
                                list={['Easy Form Survey', 'Excel Results', 'Excel List']}
                                isOpenChoice={true}
                                onChoice={this.setInputTypes}
                            />}

                        {this.state.typeOfReport === 'Campaign Results Analysis (Online Campaign, Mailing, Banner, Postering, etc.)' &&
                            <CheckboxChoice
                                id='Campaign Results Analysis (Online Campaign, Mailing, Banner, Postering, etc.)'
                                list={['Excel Results', 'External/Internal Statistic (Agency, Sales Support, OBS, TELAG)', 'LBM/BI Results', 'Online Statistics', 'Database']}
                                isOpenChoice={true}
                                onChoice={this.setInputTypes}
                            />}

                        {this.state.typeOfReport === 'Survey Results Consolidation (handwritten responses, onsite surveys)' &&
                            <CheckboxChoice
                                id='Survey Results Consolidation (handwritten responses, onsite surveys)'
                                list={['Excel Results', 'PDF Scans (no CID)']}
                                isOpenChoice={true}
                                onChoice={this.setInputTypes}
                            />}

                        {this.state.typeOfReportIsOther &&
                            <CheckboxChoice
                                id='Type of report is other'
                                list={['Easy Form Survey', 'Excel Results', 'Excel List', 'External/Internal Statistics (Agency, Sales Support, OBS, TELAG)', 'LBM/BI Results', 'Online Statistics', 'PDF Scans (no CID)', 'Database']}
                                isOpenChoice={true}
                                onChoice={this.setInputTypes}
                            />}

                    </div>

                    <div className="ms-Grid-col ms-u-sm1 ms-u-md1 ms-u-lg1">
                    </div>

                </div>


                <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-u-sm1 ms-u-md1 ms-u-lg1">
                    </div>

                    <div className="ms-Grid-col ms-u-sm10 ms-u-md10 ms-u-lg10">
                        <div style={{ marginTop: '25px', marginBottom: '15px', backgroundColor: '#f4f4f4', padding: '10px' }}>
                            <i className="ms-Icon ms-Icon--Info" style={{ color: '#005a9e' }} aria-hidden="true"></i>
                            <span className="ms-fontWeight-light" style={{ marginLeft: '10px' }}>Reporting Solution (Choose appropiate solution. Time delivery can vary depending on the output type.)</span>
                        </div>
                    </div>

                    <div className="ms-Grid-col ms-u-sm1 ms-u-md1 ms-u-lg1">
                    </div>

                </div>

                <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-u-sm1 ms-u-md1 ms-u-lg1">
                    </div>

                    <div className="ms-Grid-col ms-u-sm10 ms-u-md10 ms-u-lg10">
                        <div>
                            <div style={{ display: 'flex', justifyContent: 'flex-start', alignItems: 'center', marginLeft: '35px' }}>
                                <Checkbox
                                    defaultChecked={false}
                                    onChange={(e, checked) => { this.setReportSolution('Excel Overview (3-4 business days)', checked); }} />
                                <span style={{ marginLeft: '15px' }}>
                                    <span style={{ display: 'block' }} className='ms-font-m-plus double-asterix'>
                                        Excel Overview (3-4 business days)
                                </span>
                                    <span className='ms-font-s'>
                                        short report summarizing the average results and all comments
                                </span>
                                </span>
                                <a style={{ marginLeft: '204px' }} href="www.onet.pl"><img src="src/content/adhocreport/f_image.gif" /></a>
                            </div>

                            <div style={{ display: 'flex', justifyContent: 'flex-start', alignItems: 'center', marginLeft: '35px' }}>
                                <Checkbox
                                    defaultChecked={false}
                                    onChange={(e, checked) => { this.setReportSolution('PowerPoint Report - basic version (3-4 business days)', checked); }} />
                                <span style={{ marginLeft: '15px' }}>
                                    <span style={{ display: 'block' }} className='ms-font-m-plus double-asterix'>
                                        PowerPoint Report - basic version (3-4 business days)
                                </span>
                                    <span className='ms-font-s'>
                                        1-2 slides report presenting average results and all comments
                                </span>
                                </span>
                                <a style={{ marginLeft: '170px' }} href="www.onet.pl"><img src="src/content/adhocreport/s_image.gif" /></a>
                            </div>

                            <div style={{ display: 'flex', justifyContent: 'flex-start', alignItems: 'center', marginLeft: '35px' }}>
                                <Checkbox
                                    defaultChecked={false}
                                    onChange={(e, checked) => { this.setReportSolution('PowerPoint Report - extended version (6-10 business days)', checked); }} />
                                <span style={{ marginLeft: '15px' }}>
                                    <span style={{ display: 'block' }} className='ms-font-m-plus double-asterix'>
                                        PowerPoint Report - extended version (6-10 business days)
                                </span>
                                    <span className='ms-font-s'>
                                        a report presenting results in graphs/charts/diagrams with comments and recommendations
                                </span>
                                </span>
                                <a style={{ marginLeft: '50px' }} href="www.onet.pl"><img src="src/content/adhocreport/t_image.gif" /></a>
                            </div>

                        </div>
                    </div>

                    <div className="ms-Grid-col ms-u-sm1 ms-u-md1 ms-u-lg1">
                    </div>

                </div>

                <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-u-sm1 ms-u-md1 ms-u-lg1"></div>

                    <div className="ms-Grid-col ms-u-sm10 ms-u-md10 ms-u-lg10">
                        <div style={{ width: '750px', marginTop: '25px' }}>
                            <MessageBar
                                messageBarType={MessageBarType.warning}
                                ariaLabel='Aria help text here'>
                                <span style={{ color: '#a80000' }}>** </span> Time delivery can vary depending on the output type, given time scope is an estimation. For written feedback reports/consolidated reports the indicated preparation time is defined after the data delivery.
                        </MessageBar>
                        </div>
                    </div>

                    <div className="ms-Grid-col ms-u-sm1 ms-u-md1 ms-u-lg1"></div>

                </div>


                <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-u-sm1 ms-u-md1 ms-u-lg1"></div>

                    <div className="ms-Grid-col ms-u-sm10 ms-u-md10 ms-u-lg10" style={{ width: '750px', paddingTop: '25px' }}>
                        <DatePicker label='Requested Delivery Date' placeholder='To be confirmed individually' onSelectDate={this.setRequestDeliveryDate} />
                    </div>

                    <div className="ms-Grid-col ms-u-sm1 ms-u-md1 ms-u-lg1"></div>

                </div>


                <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-u-sm1 ms-u-md1 ms-u-lg1">
                    </div>

                    <div className="ms-Grid-col ms-u-sm10 ms-u-md10 ms-u-lg10">
                        <div style={{ marginTop: '25px', marginBottom: '15px', backgroundColor: '#f4f4f4', padding: '10px' }}>
                            <i className="ms-Icon ms-Icon--Info" style={{ color: '#005a9e' }} aria-hidden="true"></i>
                            <span className="ms-fontWeight-light" style={{ marginLeft: '10px' }}>KPIs and Targets (Please define main KPIs and relevant target values)</span>
                        </div>
                    </div>

                    <div className="ms-Grid-col ms-u-sm1 ms-u-md1 ms-u-lg1">
                    </div>

                </div>

                <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-u-sm1 ms-u-md1 ms-u-lg1"></div>

                    <div className="ms-Grid-col ms-u-sm5 ms-u-md5 ms-u-lg5">
                        <TextField label='KPIs' onChanged={value => this.setKpis(value as string)} />
                    </div>

                    <div className="ms-Grid-col ms-u-sm5 ms-u-md5 ms-u-lg5">
                        <TextField label='Targets' onChanged={value => this.setTargets(value as string)} />
                    </div>

                    <div className="ms-Grid-col ms-u-sm1 ms-u-md1 ms-u-lg1"></div>

                </div>

                <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-u-sm1 ms-u-md1 ms-u-lg1"></div>

                    <div className="ms-Grid-col ms-u-sm10 ms-u-md10 ms-u-lg10">
                        <TextField label='Remarks and further information' multiline autoAdjustHeight onChanged={value => this.setRemarks(value as string)} />
                    </div>

                    <div className="ms-Grid-col ms-u-sm1 ms-u-md1 ms-u-lg1"></div>

                </div>


                <div className="ms-Grid-row" style={{ marginTop: '35px' }}>
                    <div className="ms-Grid-col ms-u-sm1 ms-u-md1 ms-u-lg1"></div>

                    <div className="ms-Grid-col ms-u-sm10 ms-u-md10 ms-u-lg10">
                        <FileUploader onUploadedFilesChange={this.setFiles} />
                    </div>

                    <div className="ms-Grid-col ms-u-sm1 ms-u-md1 ms-u-lg1"></div>

                </div>

                <div className="ms-Grid-row" style={{ marginTop: '35px' }}>
                    <div className="ms-Grid-col ms-u-sm1 ms-u-md1 ms-u-lg1"></div>

                    <div className="ms-Grid-col ms-u-sm10 ms-u-md10 ms-u-lg10 asterix-toggle">
                        <Toggle
                            onChanged={this.setConfirmationData}
                            defaultChecked={false}
                            label=''
                            onText='You confirmed, that your request and documents attached do not include any client data'
                            offText='Please confirm, that your request and documents attached do not include any client data' />

                        <div style={{ width: '750px', marginTop: '25px' }}>
                            <MessageBar
                                messageBarType={MessageBarType.warning}
                                ariaLabel='Aria help text here'>
                                <span style={{ color: '#a80000' }}>* </span> mandatory field
                        </MessageBar>
                        </div>
                    </div>

                    <div className="ms-Grid-col ms-u-sm1 ms-u-md1 ms-u-lg1"></div>

                </div>


                <div className="ms-Grid-row" style={{ marginTop: '35px', marginBottom: '35px' }}>
                    <div className="ms-Grid-col ms-u-sm1 ms-u-md1 ms-u-lg1"></div>

                    <div className="ms-Grid-col ms-u-sm5 ms-u-md5 ms-u-lg5">
                    </div>

                    <div className="ms-Grid-col ms-u-sm5 ms-u-md5 ms-u-lg5">
                        <DefaultButton
                            style={{ marginLeft: '50%' }}
                            iconProps={{ iconName: 'Cancel' }}
                            description='Cancel'
                            text='Cancel'
                        />

                        <PrimaryButton
                            onClick={this.save}
                            style={{ marginLeft: '25px' }}
                            iconProps={{ iconName: 'Add' }}
                            description='Submit'
                            text='Submit'
                        />
                    </div>

                    <div className="ms-Grid-col ms-u-sm1 ms-u-md1 ms-u-lg1"></div>

                </div>

            </div>
        );
    }
}