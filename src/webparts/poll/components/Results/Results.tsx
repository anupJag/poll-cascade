import * as React from 'react';
import Result from './Result/Result';
import { IResultProps } from './Result/Result';
import styles from './Results.module.scss';
import { IconButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';

export interface IResultsProps {
    pollTitle: string;
    totalVotes: number;
    votingOptions: IResultProps[];
    backButtonClicked : () => void;
}

const results = (props: IResultsProps) => {
    return (
        <div className={styles.resultsMain}>
            <div className={styles.header}>
                <div>
                    <IconButton 
                        iconProps={{iconName : 'Back'}}
                        ariaLabel="Back"
                        title="Back"
                        onClick={props.backButtonClicked}
                    />
                </div>
                <div className={styles.title}>{props.pollTitle}</div>
            </div>
            <div>
                {
                    props.votingOptions.map((el : IResultProps) => {
                        return <Result title={el.title} votes={el.votes} percentage={el.percentage}/>;
                    })
                }
            </div>
            <div className={styles.footer}>
                <p>Total Votes: {props.totalVotes}</p>
            </div>
        </div>
    );
};

export default results;