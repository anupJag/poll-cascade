import * as React from 'react';
import { Placeholder } from '@pnp/spfx-controls-react/lib/Placeholder';
import Poll from './Poll';
import { IMainProps } from './IPollProps';

export default class Main extends React.Component<IMainProps, {}>{

    state = {
        showPlaceHolder: false
    }

    render(): React.ReactElement<IMainProps> {

        const { pollTitle, list, pollResult, pollOption } = this.props;
        const renderPlaceHolder: JSX.Element = !(pollTitle && list && pollResult && pollOption) ?
            <Placeholder
                iconName='CheckboxComposite'
                iconText='Poll'
                description='Find out what others think'
                buttonLabel='Configure'
                onConfigure={this.props._onConfigure}
            />
            :
            <Poll
                list={list}
                pollOption={pollOption}
                pollResult={pollResult}
                pollTitle={pollTitle}
                webURL={this.props.webURL}
            />

        return (
            <div>
                {renderPlaceHolder}
            </div>
        );
    }
}


