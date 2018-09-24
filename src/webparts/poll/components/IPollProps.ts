export interface IPollProps {
  pollTitle: string;
  pollGUID: string;
  pollListGUID: string;
  webURL: string;
}


export interface IPollOption {
  option: string;
}

export interface IPollData {
  Title: string;
  Votes: number;
  PollID: string;
}


export interface IMainProps {
  pollTitle: string;
  pollGUID: string;
  pollListGUID: string;
  webURL: string;
  pollSetupCompleted: boolean;
  _onConfigure: () => void;
}

export interface IFieldTypeKind {
  FieldTypeKind: number;
}

export enum FieldNames {
  Id = "Id",
  Title = "Title",
  Votes = "Votes",
  PollID = "PollID"
}