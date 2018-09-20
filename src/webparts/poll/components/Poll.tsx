import * as React from 'react';
import styles from './Poll.module.scss';
import { IPollProps } from './IPollProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Web } from 'sp-pnp-js';
import { ChoiceGroup, IChoiceGroupOption } from 'office-ui-fabric-react/lib/ChoiceGroup';

export interface IPollState {
  list: string;
  option: string;
  votes: string;
  pollData: any[];
  selectedVote: string;
  errorOccured: boolean;
  errorMessage: string;
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
      errorMessage: undefined
    }
  }

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
      pollData: listData
    });
  }

  componentWillReceiveProps(nextProps: IPollProps) {
    if (nextProps.list !== this.props.list || nextProps.pollResult !== this.props.pollResult || nextProps.pollOption !== this.props.pollOption) {
      if (nextProps.list && nextProps.pollResult && nextProps.pollOption) {
        this.setState({
          list: nextProps.list,
          option: nextProps.pollOption,
          votes: nextProps.pollResult
        }, () => {
          this.getExisitingListData()
            .then((listData: any[]) => {
              this.createDateForState(listData);
            });
        });
      }
    }
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
      .then(p => p).catch((reject: any) => reject)

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
    this.setState({
      selectedVote: option.key
    });
  }


  public render(): React.ReactElement<IPollProps> {
    return (
      <div className={styles.poll}>
        <ChoiceGroup options={this.createChoiceOptions()} onChanged={this._onChange}></ChoiceGroup>
      </div>
    );
  }
}
