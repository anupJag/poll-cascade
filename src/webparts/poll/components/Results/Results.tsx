import * as React from 'react';
import Result from './Result/Result';
import { IResultProps } from './Result/Result';
import styles from './Results.module.scss';

export interface IResultsProps {
    pollTitle: string;
    totalVotes: number;
    votingOptions: IResultProps[];
}

const results = (props: IResultsProps) => {
    return (
        <div className={styles.resultsMain}>
            <div className={styles.header}>
                <p>{props.pollTitle}</p>
            </div>
            <div>
                {
                    props.votingOptions.map((el : IResultProps) => {
                        return <Result title={el.title} votes={el.votes} percentage={el.percentage}/>
                    })
                }
            </div>
            <div className={styles.footer}>
                <p>Total Votes: {props.totalVotes}</p>
            </div>
        </div>
    );
}

export default results;