import * as React from 'react';
import { Placeholder } from '@pnp/spfx-controls-react/lib/Placeholder';


export interface IErrorProps{
    ErrorText : string;
}


const error = (props : IErrorProps) => <Placeholder
    iconName="Error"
    iconText="Oh no! Something has gone wrong!!"
    description={props.ErrorText}
/>;

export default error;