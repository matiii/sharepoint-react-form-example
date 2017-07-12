import * as React from 'react';
import { Overlay } from 'office-ui-fabric-react/lib/Overlay';
import {
    Spinner,
    SpinnerSize
    } from 'office-ui-fabric-react/lib/Spinner';

export interface CsSpinnerProps {
    text: string;
    isVisible: boolean;
}

const CsSpinner: React.StatelessComponent<CsSpinnerProps> = (props: CsSpinnerProps) => {

    return (
        <div>
            {props.isVisible && <Overlay style={{ display: 'flex', justifyContent: 'center', flexDirection: 'column'}}> <Spinner size={SpinnerSize.large} label={props.text} /> </Overlay>}
        </div>
        );

}

export const CsSpinnerLight: React.StatelessComponent<CsSpinnerProps> = (props: CsSpinnerProps) => {
    return (<div style={{ marginTop: '150px', marginBottom: '150px'}}>
                {props.isVisible &&<Spinner size={SpinnerSize.large} label={props.text} />}
    </div>);
}

export default CsSpinner;