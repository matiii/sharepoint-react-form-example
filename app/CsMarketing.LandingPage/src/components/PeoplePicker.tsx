import * as React from 'react';
import { BaseComponent, autobind } from 'office-ui-fabric-react/lib/Utilities';
import { NormalPeoplePicker } from 'office-ui-fabric-react/lib/Pickers';
import { IPersonaProps } from 'office-ui-fabric-react/lib/Persona';
import { getPeopleService } from '../api/bootstrap';
import { Label } from 'office-ui-fabric-react/lib/Label';

export interface PeoplePickerProps {
    defaultSelected: IPersonaProps[];
    label: string;
    isRequired: boolean;

    selected: (users: IPersonaProps[]) => void;
}

export interface PeoplePickerState {

}

const PeopleService = getPeopleService();

const suggestionProps = {
    suggestionsHeaderText: 'Suggested People',
    mostRecentlyUsedHeaderText: 'Suggested Contacts',
    noResultsFoundText: 'No results found',
    loadingText: 'Loading',
    showRemoveButtons: true
};

export class PeoplePicker extends BaseComponent<PeoplePickerProps, PeoplePickerState> {

    private readonly peopleService = new PeopleService();

    constructor(props: PeoplePickerProps, context: any) {
        super(props, context);
    }

    @autobind
    private onFilterChanged(filterText: string, currentPersonas: IPersonaProps[], limitResults?: number) {
        if (filterText && currentPersonas.length === 0) {
            return this.peopleService.getPeoplesByString(filterText, 5);
        } else {
            return [];
        }
    }

    @autobind
    private onChangeHandler(items?: IPersonaProps[]) {
        if (items && items.length) {
            this.props.selected(items);
        } else {
            this.props.selected([]);
        }
    }

    render() {
        return (
            <div style={{ position: 'relative'}}>
                <Label required={this.props.isRequired}>{this.props.label}</Label> 
                <NormalPeoplePicker
                    onChange={this.onChangeHandler}
                    defaultSelectedItems={this.props.defaultSelected}
                    onResolveSuggestions={this.onFilterChanged}
                    getTextFromItem={(persona: IPersonaProps) => persona.primaryText}
                    pickerSuggestionsProps={suggestionProps}
                    className={'ms-PeoplePicker'}
                    key={'normal'}
                />
                <i style={{ position: 'absolute', top:'38px', right: '5px'}} className="ms-Icon ms-Icon--AddFriend" aria-hidden="true"></i>
            </div>);
    }
}