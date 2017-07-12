import * as React from 'react';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import * as update from 'immutability-helper';

export interface CheckboxChoiceProps {
    id: string;
    list: string[];
    isOpenChoice: boolean;
    onChoice: (choices: string[]) => void;
}

export interface CheckboxChoiceState {
    choices: string[];
    openChoice: string;
    enableOpenChoice: boolean;
}

const openChoice = 'Other (please specify)';

export class CheckboxChoice extends React.Component<CheckboxChoiceProps, CheckboxChoiceState> {

    constructor(props, context) {
        super(props, context);

        if (this.props.isOpenChoice) {
            this.props.list.push(openChoice);
        }

        this.state = { choices: [], enableOpenChoice: false, openChoice: '' };
    }

    componentWillReceiveProps(nextProps: {readonly [P in "list" | "isOpenChoice" | "onChoice" | "id"]: CheckboxChoiceProps[P]},
        nextContext): void {
        if (nextProps.isOpenChoice) {
            nextProps.list.push(openChoice);
        }
    }

    @autobind
    private handleChoice(choice: string, checked: boolean) {

        if (choice === openChoice) {
            this.setState(prev => update(prev, { enableOpenChoice: { $set: checked } }), this.propagateChanges);
        } else {

            if (checked) {
                this.setState(prev => update(prev, { choices: { $push: [choice] } }), this.propagateChanges);
            } else {
                this.setState(prev => update(prev, { choices: { $set: prev.choices.filter(x => x !== choice) } }), this.propagateChanges);
            }
            

        }

    }

    @autobind
    private propagateChanges() {
        let choices = [...this.state.choices];

        if (this.state.enableOpenChoice) {
            choices.push(this.state.openChoice);
        }

        this.props.onChoice(choices);
    }

    @autobind
    private handleOpenChoice(choice: string) {
        this.setState(prev => update(prev, { openChoice: { $set: choice }}), this.propagateChanges);
    }

    render() {

        return (
            <div>
                {
                    this.props.list.map(x =>
                        <Checkbox
                            key={this.props.id+x}
                            label={x}
                            checked={this.state.choices.filter(c => c === x).length > 0 || x === openChoice && this.state.enableOpenChoice}
                            onChange={(ev, checked) => this.handleChoice(x, checked)}
                        />)
                    }
                {this.props.isOpenChoice && <TextField onChanged={(value) => { this.handleOpenChoice(value as string); }} disabled={!this.state.enableOpenChoice} />}
            </div>
        );
    }


}