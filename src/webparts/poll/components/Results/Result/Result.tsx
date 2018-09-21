import * as React from 'react';
import styles from './Result.module.scss';

export interface IResultProps{
    title : string;
    votes : number;
    percentage : number;
}

const result = (props : IResultProps) => {

    const configurableWidth : React.CSSProperties = {};

    if(props.percentage || props.percentage === 0){
        configurableWidth.width = `${props.percentage}%`;
    }

    return (
        <div className={styles.resultHolder}>
            <div className={styles.resultLabel}>
                <div>{props.title}</div>
                <div>{props.percentage}% ({props.votes} votes)</div>
            </div>
            <div className={styles.resultBarHolder}>
                <div className={styles.resultProgress} style={configurableWidth}></div>
            </div>
        </div>
    );
}

export default result;