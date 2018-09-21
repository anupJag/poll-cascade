import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import * as strings from 'PollWebPartStrings';

import Main from './components/Main';
import { IMainProps } from './components/IPollProps';
import { IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';

export interface IPollWebPartProps {
  pollTitle: string;
  list: string;
  pollOption: string;
  pollResult: string;
}

export interface IListColumnOptions {
  key: string;
  text: string;
}

export interface IColumnDataStructure {
  InternalName: string;
  FieldTypeKind: number;
}

export default class PollWebPart extends BaseClientSideWebPart<IPollWebPartProps> {

  private _ListColumns: any[];
  private _ResultColumns: any[];
  private _PollOptionSelection: boolean = true;
  private _PollResultSelection: boolean = true;

  public render(): void {
    const element: React.ReactElement<IMainProps> = React.createElement(
      Main,
      {
        pollTitle: this.properties.pollTitle,
        list: this.properties.list,
        pollOption: this.properties.pollOption,
        pollResult: this.properties.pollResult,
        webURL: this.context.pageContext.web.absoluteUrl,
        _onConfigure : this._onConfigure.bind(this)
      }
    );

    ReactDom.render(element, this.domElement);
  }
 

  private _onConfigure() {
    // Context of the web part
    this.context.propertyPane.open();
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  protected getColumnsForPropertyPane = (): Promise<any[]> => {

    if (!this.properties.list) {
      // resolve to empty options since no list has been selected
      return Promise.resolve();
    }

    return new Promise<any[]>((resolve: (columns: any[]) => void, reject: (error: any) => void) => {
      this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbyid('${this.properties.list}')/fields/?$filter=((Hidden eq false) and (ReadOnlyField eq false) and (FieldTypeKind ne 19) and (FieldTypeKind ne 12) and ((FieldTypeKind eq 2) or (FieldTypeKind eq 9) or (FieldTypeKind eq 1)))`, SPHttpClient.configurations.v1, {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'odata-version': ''
        }
      }).then((response: SPHttpClientResponse) => {
        return response.json();
      }).then((data: any) => {
        let _tempData: any[] = [];
        (data.value).forEach((element: any) => {
          _tempData.push(element);
        });
        resolve(_tempData);
      }).catch((error: Error) => {
        reject(error);
      });
    });
  }

  protected onPropertyPaneConfigurationStart(): void {

    this._PollOptionSelection = !this.properties.list || !this._ListColumns;
    this._PollResultSelection = !this.properties.list || !this._ListColumns || !this._ResultColumns;

    if (!this.properties.list) {
      return;
    }

    this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'pollOption');
    this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'pollResult');

    this.getColumnsForPropertyPane()
      .then((columns: any[]): void => {
        var columnsRequired: IDropdownOption[] = [];
        var columnDataStructureTemp : IColumnDataStructure[] = [];
        columnsRequired.push({ key: null, text: null, selected: true });
        columns.forEach((element) => {
          columnDataStructureTemp.push({
            InternalName: element.InternalName,
            FieldTypeKind: element.FieldTypeKind
          });
          columnsRequired.push({
            key: element.InternalName,
            text: element.Title,
            selected: false
          });
        });
        this._ListColumns = columnsRequired;
        this._PollOptionSelection = !this.properties.list;
        this.context.propertyPane.refresh();
        this.context.statusRenderer.clearLoadingIndicator(this.domElement);
      }).then(() => {
        let listColumnsTemp: IDropdownOption[] = [];
        listColumnsTemp.push({ key: null, text: null, selected: true });
        this._ListColumns.forEach((element: IDropdownOption) => {
          if (element.key !== this.properties.pollOption && element.key !== null) {
            listColumnsTemp.push(element);
          }
        });
        this._ResultColumns = listColumnsTemp;
        this._PollResultSelection = !this.properties.list;
        this.context.propertyPane.refresh();
        this.context.statusRenderer.clearLoadingIndicator(this.domElement);
        this.render();
      });
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if (propertyPath == "list" && (newValue || newValue == "")) {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      // get previously selected item
      const previousPollOptionColumn: string = this.properties.pollOption;
      this.properties.pollOption = undefined;
      this._ListColumns = undefined;
      this._ResultColumns = undefined;
      const previousPollResultColumn: string = this.properties.pollResult;
      this.properties.pollResult = undefined;
      this.onPropertyPaneFieldChanged('pollOption', previousPollOptionColumn, this.properties.pollOption);
      this.onPropertyPaneFieldChanged('pollResult', previousPollResultColumn, this.properties.pollResult);
      this._PollOptionSelection = true;
      this._PollResultSelection = true;
      // refresh the item selector control by repainting the property pane
      this.context.propertyPane.refresh();
      // communicate loading items
      this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'pollOption');

      if (!this.properties.list) {
        return;
      }

      this.getColumnsForPropertyPane().then((columns: any[]): void => {
        var columnsRequired: IDropdownOption[] = [];
        var columnDataStructureTemp : IColumnDataStructure[] = [];
        columnsRequired.push({ key: null, text: null, selected: true });
        columns.forEach((element) => {
          columnDataStructureTemp.push({
            InternalName: element.InternalName,
            FieldTypeKind: element.FieldTypeKind
          });
          columnsRequired.push({
            key: element.InternalName,
            text: element.Title,
            selected: false
          });
        });
        this._ListColumns = columnsRequired;
        this._PollOptionSelection = false;
        this.context.statusRenderer.clearLoadingIndicator(this.domElement);
        this.render();
        this.context.propertyPane.refresh();
      });
    }
    else if (propertyPath == "pollOption" && newValue) {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      const previousPollResultColumn: string = this.properties.pollResult;
      this.properties.pollResult = undefined;
      this.onPropertyPaneFieldChanged('pollResult', previousPollResultColumn, this.properties.pollResult);
      this._PollResultSelection = true;
      this.context.propertyPane.refresh();
      // communicate loading items
      this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'pollResult');
      let listColumnsTemp: IDropdownOption[] = [];
      listColumnsTemp.push({ key: null, text: null, selected: true });
      this._ListColumns.forEach((element: IDropdownOption) => {
        if (element.key !== this.properties.pollOption && element.key !== null) {
          listColumnsTemp.push({
            key: element.key,
            text: element.text,
            selected: false
          });
        }
      });
      this._ResultColumns = listColumnsTemp;
      this._PollResultSelection = false;
      this.context.statusRenderer.clearLoadingIndicator(this.domElement);
      this.context.propertyPane.refresh();
      this.render();
    }
    else {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    }
  }

  protected onPollTitleHandler = (value: string): string => {
    if (value.trim().length > 0) {
      return '';
    }

    return 'Your Poll Should Have a Title';
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "Configure your Poll Title, and select the List where you wish to record the Poll."
          },
          groups: [
            {
              groupFields: [
                PropertyPaneTextField('pollTitle', {
                  label: "Poll Title",
                  onGetErrorMessage: this.onPollTitleHandler.bind(this)
                }),
                PropertyFieldListPicker('list', {
                  label: 'Select a list',
                  selectedList: this.properties.list,
                  includeHidden: false,
                  baseTemplate: 100,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'listPickerFieldId'
                }),
                PropertyPaneDropdown('pollOption', {
                  label: "Select the field for your Poll Options",
                  options: this._ListColumns,
                  disabled: this._PollOptionSelection,
                  selectedKey: null
                }),
                PropertyPaneDropdown('pollResult', {
                  label: "Select the field to store the votes",
                  options: this._ResultColumns,
                  disabled: this._PollResultSelection,
                  selectedKey: null
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
