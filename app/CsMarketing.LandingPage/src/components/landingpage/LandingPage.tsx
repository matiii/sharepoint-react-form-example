import * as React from 'react';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { DatePicker, DayOfWeek } from 'office-ui-fabric-react/lib/DatePicker';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import * as update from 'immutability-helper';
import { getListItemService, getPeopleService } from '../../api/bootstrap';
import { IPersonaProps } from 'office-ui-fabric-react/lib/Persona';
import { BaseComponent, autobind } from 'office-ui-fabric-react/lib/Utilities';
import strings from '../../strings';
import CsSpinner, { CsSpinnerLight } from '../CsSpinner';
import { PersonaUser } from '../../model/User';
import {
    CompoundButton,
    IButtonProps
} from 'office-ui-fabric-react/lib/Button';
import { AdHocReport } from '../../forms/AdHocReport';
import * as routing from './Routing';

export interface LandingPageState {
    appIsLoading: boolean;
    currentUser: PersonaUser;

    renderAdHocReport: boolean;
}


const ListItemService = getListItemService();
const PeopleService = getPeopleService();

export class LandingPage extends React.Component<{}, LandingPageState>{

    private readonly listApi = new ListItemService();
    private readonly peopleService = new PeopleService();

    constructor(props, context) {
        super(props, context);

        this.state = { appIsLoading: true, currentUser: undefined, renderAdHocReport: false };
    }

    componentDidMount(): void {

        this.peopleService.getCurrentPersonaUser(user => {

            this.setState(prev => update(prev, { appIsLoading: { $set: false }, currentUser: { $set: user } }));

            console.log(user);

            if (routing.isAdHocReportFormNew()) {

                parent.location.hash = routing.adHocReportFormNew;
                this.goToAdHocReport();

            } else {
                parent.location.hash = routing.landingPage;
            }

        }, console.log);
    }

    @autobind
    private goToAdHocReport() {

        this.setState(prev => update(prev, { renderAdHocReport: { $set: true } }));

    }

    render() {

        if (this.state.appIsLoading) {
            return (<CsSpinnerLight text={strings.landingPageIsLoadingMessage} isVisible={true} />);
        }

        if (this.state.renderAdHocReport) {
            return (<AdHocReport
                currentUser={this.state.currentUser}
            />);
        }

        return (<div>

            <CompoundButton
                onClick={() => this.goToAdHocReport()}
            >
                <span className='ms-fontSize-sPlus ms-fontWeight-semibold'>Ad Hoc Report </span>
                <span className='ms-font-xs'>Ad Hoc Campaign Evaluation, Client Feedback, Event Report</span>
            </CompoundButton>

        </div>);
    }
}
