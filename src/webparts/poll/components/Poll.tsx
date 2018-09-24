import * as React from 'react';
import styles from './Poll.module.scss';
import { IPollProps, IFieldTypeKind, FieldNames } from './IPollProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Web } from 'sp-pnp-js';
import { ChoiceGroup, IChoiceGroupOption } from 'office-ui-fabric-react/lib/ChoiceGroup';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import Results from './Results/Results';
import { IResultProps } from './Results/Result/Result';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';

export interface IPollState {
  listGUID: string;
  pollGUID : string;
  pollData: any[];
  selectedVote: string;
  errorOccured: boolean;
  errorMessage: string;
  renderResult: boolean;
  resultSet: IResultProps[];
  showSpinner: boolean;
  shouldSubmitButtonDisabled: boolean;
  shouldResultsBeDisabled: boolean;
}


export default class Poll extends React.Component<IPollProps, IPollState> {

  constructor(props: IPollProps) {
    super(props);
    this.state = {
      listGUID: props.pollListGUID,
      pollGUID : props.pollGUID,
      pollData: [],
      selectedVote: undefined,
      errorOccured: false,
      errorMessage: undefined,
      renderResult: false,
      resultSet: undefined,
      showSpinner: false,
      shouldSubmitButtonDisabled: true,
      shouldResultsBeDisabled: false,
    };
  }

  // tslint:disable-next-line:member-access
  componentDidMount() {
    this.listData()
      .then((listData: any[]) => {
        this.createDateForState(listData);
      }).catch((error: any) => {
        this.setState({
          errorOccured: true,
          errorMessage: error
        });
      });
  }


  protected createDateForState = async (listData: any[]) => {
    await this.setState({
      pollData: listData,
    });
  }

  protected createChartData = (): IResultProps[] => {
    let tempResult: IResultProps[] = [];
    let currentPollData = [...this.state.pollData];

    if (!(currentPollData && currentPollData.length > 0)) {
      return tempResult;
    }

    const totalVotes: number = this.getTotalVotes();

    currentPollData.forEach((element) => {
      let votes: number = isNaN(parseInt(element[FieldNames.Votes], 0)) ? 0 : parseInt(element[FieldNames.Votes], 0);
      let percentage: number = 0;
      if (votes === 0 || totalVotes === 0) {
        percentage = 0;
      }
      else {
        percentage = (votes / totalVotes) * 100;
        percentage = parseFloat(percentage.toFixed(2));
      }


      tempResult.push({
        title: element[FieldNames.Title],
        votes: isNaN(element[FieldNames.Votes]) || element[FieldNames.Votes] === null || element[FieldNames.Votes] === undefined ? 0 : element[FieldNames.Votes],
        percentage: percentage
      });
    });

    return tempResult;
  }

  protected listData = async () => {
    const web = new Web(this.props.webURL);
    const selectParams = [FieldNames.Id, FieldNames.Title, FieldNames.Votes, FieldNames.PollID];
    const data = await web.lists.getById(this.state.listGUID).items.configure({
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'odata-version': ''
      }
    }).select(...selectParams).filter("PollID eq '" + this.state.pollGUID + "'").get()
      .then(p => p).catch((reject: any) => reject);

    return data;
  }

  protected createChoiceOptions = (): IChoiceGroupOption[] => {
    let choiceGroupToBeReturned: IChoiceGroupOption[] = [];
    if (this.state.pollData && this.state.pollData.length > 0) {
      const data = [...this.state.pollData];
      data.forEach((element: IChoiceGroupOption) => {
        choiceGroupToBeReturned.push({
          key: element[FieldNames.Id].toString(),
          text: element[FieldNames.Title].toString()
        });
      });
    }
    return choiceGroupToBeReturned;
  }

  protected _onChange = (option: IChoiceGroupOption, ev: React.FormEvent<HTMLInputElement>): void => {
    this.setState((prevState: IPollState) => {
      return {
        selectedVote: option.key,
        shouldSubmitButtonDisabled: !prevState.shouldSubmitButtonDisabled
      };
    });
  }

  protected _voteClickedHandler = async () => {
    this.setState((prevState: IPollState) => {
      return {
        shouldSubmitButtonDisabled: !prevState.shouldSubmitButtonDisabled,
        shouldResultsBeDisabled: !prevState.shouldResultsBeDisabled,
        showSpinner: !prevState.showSpinner
      };
    });
    let pollData: any[] = [...this.state.pollData];
    let dataSetToBeModified = pollData.filter(el => el[FieldNames.Id] === parseInt(this.state.selectedVote, 0));

    if (dataSetToBeModified.length > 0) {
      const web = new Web(this.props.webURL);
      let idToBeModified = dataSetToBeModified[0][FieldNames.Id];
      let currentValue: any = dataSetToBeModified[0][FieldNames.Votes];
      if (isNaN(parseInt(currentValue, 0))) {
        currentValue = 1;
      }
      else {
        currentValue = parseInt(currentValue, 0) + 1;
      }

      let valueToBeUpdated = {};
      valueToBeUpdated[FieldNames.Votes] = currentValue;

      await web.lists.getById(this.state.listGUID).items.getById(parseInt(idToBeModified, 0)).update(valueToBeUpdated).then(i => { console.log(i); }).catch(error => { console.log(error); });

      this.listData().then((listData: any[]) => {
        this.createDateForState(listData);
      }).then(() => {
        this.setState((prevState: IPollState) => {
          return {
            showSpinner: !prevState.showSpinner,
            renderResult: !prevState.renderResult
          };
        });
      });
    }
    else {
      //Cause Error
    }
  }

  protected getTotalVotes = (): number => {
    let currentPollData = [...this.state.pollData];

    if (!(currentPollData && currentPollData.length > 0)) {
      return 0;
    }

    let sum = 0;

    currentPollData.forEach((element) => {
      var individualVoteCount = element[FieldNames.Votes];

      if (isNaN(parseInt(individualVoteCount, 0))) {
        individualVoteCount = 0;
      }
      else {
        individualVoteCount = parseInt(individualVoteCount, 0);
      }

      sum = sum + individualVoteCount;

    });

    return sum;
  }

  protected _showResultsHandler = () => {
    this.setState((prevState: IPollState) => {
      return {
        renderResult: !prevState.renderResult,
        shouldResultsBeDisabled: !prevState.shouldResultsBeDisabled,
      };
    });
  }

  public render(): React.ReactElement<IPollProps> {
    const showResults: JSX.Element = this.state.renderResult ?
      <Results
        pollTitle={this.props.pollTitle}
        votingOptions={this.createChartData()}
        totalVotes={this.getTotalVotes()}
        backButtonClicked={this._showResultsHandler.bind(this)}
      /> :
      <div>
        <header className={styles.title}>{this.props.pollTitle}</header>
        <ChoiceGroup
          options={this.createChoiceOptions()}
          onChanged={this._onChange}
        >
        </ChoiceGroup>
        <div className={styles.buttonControls}>
          <DefaultButton
            primary={true}
            data-automation-id="test"
            disabled={this.state.shouldSubmitButtonDisabled}
            text="Vote"
            onClick={this._voteClickedHandler}
          />
          <DefaultButton
            style={{ marginLeft: "2em" }}
            primary={true}
            data-automation-id="test"
            text="Poll Results"
            disabled={this.state.shouldResultsBeDisabled}
            onClick={this._showResultsHandler}
          />
        </div>
      </div>;

    const showSpinner: JSX.Element = this.state.showSpinner ?
      <div className={styles.spinnerMainHolder}>
        <div className={styles.spinnerHolder}>
          <Spinner
            size={SpinnerSize.large}
            label="Thank You For Your Valuable Contribution"
            style={{ zIndex: 100 }}
          />
        </div>
      </div> : null;

    return (
      <div className={styles.poll}>
        {showSpinner}
        {showResults}
      </div>
    );
  }
}
