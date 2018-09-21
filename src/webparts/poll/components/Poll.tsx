import * as React from 'react';
import styles from './Poll.module.scss';
import { IPollProps, IFieldTypeKind, FieldType } from './IPollProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Web } from 'sp-pnp-js';
import { ChoiceGroup, IChoiceGroupOption } from 'office-ui-fabric-react/lib/ChoiceGroup';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import Results from './Results/Results';
import { IResultProps } from './Results/Result/Result';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';

export interface IPollState {
  list: string;
  option: string;
  votes: string;
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
      list: props.list,
      option: props.pollOption,
      votes: props.pollResult,
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

    if (!(this.state.list && this.state.option && this.state.votes)) {
      return;
    }

    this.getExisitingListData()
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

  // tslint:disable-next-line:member-access
  componentWillReceiveProps(nextProps: IPollProps) {
    if (nextProps.list !== this.props.list || nextProps.pollResult !== this.props.pollResult || nextProps.pollOption !== this.props.pollOption) {
      if (nextProps.list && nextProps.pollResult && nextProps.pollOption) {
        this.setState({
          list: nextProps.list,
          option: nextProps.pollOption,
          votes: nextProps.pollResult,
        }, () => {
          this.getExisitingListData()
            .then((listData: any[]) => {
              this.createDateForState(listData);
            });
        });
      }
    }
  }

  protected createChartData = (): IResultProps[] => {
    let tempResult: IResultProps[] = [];
    let currentPollData = [...this.state.pollData];

    if (!(currentPollData && currentPollData.length > 0)) {
      return tempResult;
    }

    const totalVotes: number = this.getTotalVotes();

    currentPollData.forEach((element) => {
      let votes: number = isNaN(parseInt(element[this.state.votes], 0)) ? 0 : parseInt(element[this.state.votes], 0);
      let percentage: number = 0;
      if (votes === 0 || totalVotes === 0) {
        percentage = 0;
      }
      else {
        percentage = (votes / totalVotes) * 100;
        percentage = parseFloat(percentage.toFixed(2));
      }


      tempResult.push({
        title: element[this.state.option],
        votes: isNaN(element[this.state.votes]) || element[this.state.votes] === null || element[this.state.votes] === undefined ? 0 : element[this.state.votes],
        percentage: percentage
      });
    });

    return tempResult;
  }

  protected getExisitingListData = async () => {
    const web = new Web(this.props.webURL);
    const selectParams = ["Id", this.state.option, this.state.votes];
    const data = await web.lists.getById(this.state.list).items.configure({
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'odata-version': ''
      }
    }).select(...selectParams).get()
      .then(p => p).catch((reject: any) => reject);

    return data;
  }

  protected createChoiceOptions = (): IChoiceGroupOption[] => {
    let choiceGroupToBeReturned: IChoiceGroupOption[] = [];
    if (this.state.pollData && this.state.pollData.length > 0) {
      const data = [...this.state.pollData];
      data.forEach((element: IChoiceGroupOption) => {
        choiceGroupToBeReturned.push({
          key: element["ID"].toString(),
          text: element[this.props.pollOption].toString()
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
      }
    });
    let pollData: any[] = [...this.state.pollData];
    let dataSetToBeModified = pollData.filter(el => el["Id"] === parseInt(this.state.selectedVote, 0));

    if (dataSetToBeModified.length > 0) {
      const web = new Web(this.props.webURL);
      let idToBeModified = dataSetToBeModified[0]["Id"];
      let currentValue: any = dataSetToBeModified[0][this.state.votes];
      if (isNaN(parseInt(currentValue, 0))) {
        currentValue = 1;
      }
      else {
        currentValue = parseInt(currentValue, 0) + 1;
      }

      //Handle Condition if the Field is of type Text or Number
      const dataType: IFieldTypeKind = await web.lists.getById(this.state.list).fields.getByInternalNameOrTitle(this.state.votes).select("FieldTypeKind").configure({
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'odata-version': ''
        }
      }).get().then(p => p);

      switch (dataType.FieldTypeKind) {
        case FieldType.Text:
          currentValue = currentValue.toString();
          break;

        default:
          currentValue = currentValue;
          break;
      }


      let valueToBeUpdated = {};
      valueToBeUpdated[this.state.votes] = currentValue;

      await web.lists.getById(this.state.list).items.getById(parseInt(idToBeModified, 0)).update(valueToBeUpdated).then(i => { console.log(i); }).catch(error => { console.log(error); });

      this.getExisitingListData().then((listData: any[]) => {
        this.createDateForState(listData);
      }).then(() => {
        this.setState((prevState: IPollState) => {
          return {
            showSpinner: !prevState.showSpinner,
            renderResult: !prevState.renderResult
          }
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
      var individualVoteCount = element[this.state.votes];

      if (isNaN(parseInt(individualVoteCount, 0))) {
        individualVoteCount = 0
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
        renderResult: !prevState.renderResult
      }
    })
  }

  public render(): React.ReactElement<IPollProps> {
    const showResults: JSX.Element = this.state.renderResult ?
      <Results
        pollTitle={this.props.pollTitle}
        votingOptions={this.createChartData()}
        totalVotes={this.getTotalVotes()}
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
            text="See what's trending"
            disabled={this.state.shouldResultsBeDisabled}
            onClick={this._showResultsHandler}
          />
        </div>
      </div>

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
