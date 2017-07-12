import * as React from 'react';
import { ChoiceGroup, IChoiceGroupOption } from 'office-ui-fabric-react/lib/ChoiceGroup';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';


export interface RadioChoiceProps {
    list: string[];
    isOpenChoice: boolean;
    onChoice: (choice: string | { isOptional: boolean, value: string }) => void;
}

export interface RadioChoiceState {
    choice: string;
}

const openChoice = 'Other (please specify)';

export class RadioChoice extends React.Component<RadioChoiceProps, RadioChoiceState> {

    private openChoice: string = '';

    constructor(props, context) {
        super(props, context);

        if (this.props.isOpenChoice) {
            this.props.list.push(openChoice);
        }

        this.state = { choice: this.props.list[0] };
    }

    componentWillReceiveProps(nextProps: {readonly [P in "list" | "isOpenChoice" | "onChoice"]: RadioChoiceProps[P]},
        nextContext): void {

        if (nextProps.isOpenChoice) {
            nextProps.list.push(openChoice);
        }
    }

    @autobind
    private handleChoice(choice: string) {

        if (choice !== openChoice) {
            this.props.onChoice(choice);
        } else {
            this.props.onChoice({ isOptional: true, value: this.openChoice });
        }

        this.setState({ choice });
    }

    @autobind
    private handleOpenChoice(value: string) {
        this.openChoice = value;

        this.props.onChoice({ isOptional: true, value: this.openChoice });
    }

    render() {

        return (
            <div>
                <ChoiceGroup
                    selectedKey={this.state.choice}
                    options={this.props.list.map(x => ({ key: x, text: x }))}
                    onChange={(ev?: React.FormEvent<HTMLElement | HTMLInputElement>, option?: IChoiceGroupOption) => { if (option) this.handleChoice(option.key); }}
                />

                {this.props.isOpenChoice && <TextField onChanged={(value) => { this.handleOpenChoice(value as string); }} disabled={this.state.choice !== openChoice} />}
            </div>
        );
    }


}