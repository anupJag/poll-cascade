import { IColumnDataStructure } from '../PollWebPart';

export interface IPollProps {
  pollTitle: string;
  list: string;
  pollOption: string;
  pollResult: string;
  webURL: string;
  columnDataStructure : IColumnDataStructure[];
}


export interface IFieldTypeKind{
  FieldTypeKind: number;
}

export enum FieldType{
  Integer = 1,
  Text = 2,
  Number = 9
}