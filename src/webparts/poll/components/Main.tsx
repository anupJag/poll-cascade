import * as React from 'react';
import { Placeholder } from '@pnp/spfx-controls-react/lib/Placeholder';
import Poll from './Poll';
import { IMainProps } from './IPollProps';

export default class Main extends React.Component<IMainProps, {}>{

    public render(): React.ReactElement<IMainProps> {

        const { pollTitle, pollGUID, pollSetupCompleted } = this.props;
        const renderPlaceHolder: JSX.Element = (pollSetupCompleted === undefined || pollSetupCompleted === false) ?
            <Placeholder
                iconName='CheckboxComposite'
                iconText='Poll'
                description='Find out what others think'
                buttonLabel='Configure'
                onConfigure={this.props._onConfigure}
            />
            :
            <Poll
                pollGUID={pollGUID}
                pollListGUID={this.props.pollListGUID}
                pollTitle={pollTitle}
                webURL={this.props.webURL}
            />;

        return (
            <div>
                {renderPlaceHolder}
            </div>
        );
    }
}


