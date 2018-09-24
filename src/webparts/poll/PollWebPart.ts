import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneLabel,
  PropertyPaneToggle
} from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import * as strings from 'PollWebPartStrings';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import Main from './components/Main';
import { Web, ItemAddResult } from 'sp-pnp-js';
import { IMainProps, IPollOption, IPollData } from './components/IPollProps';

export interface IPollWebPartProps {
  pollTitle: string;
  guid: string;
  pollListGUID: string;
  pollDataCollection: any[];
  pollSetupCompleted : boolean;
}

export default class PollWebPart extends BaseClientSideWebPart<IPollWebPartProps> {

  private _IspollOptionsSetupCompleted: boolean = false;

  public render(): void {
    const element: React.ReactElement<IMainProps> = React.createElement(
      Main,
      {
        pollTitle: this.properties.pollTitle,
        pollGUID: this.properties.guid,
        pollListGUID : this.properties.pollListGUID,
        pollSetupCompleted : this.properties.pollSetupCompleted,
        webURL : this.context.pageContext.web.absoluteUrl,
        _onConfigure: this._onConfigure.bind(this)
      }
    );

    ReactDom.render(element, this.domElement);
  }


  private _onConfigure() {
    // Context of the web part
    this.context.propertyPane.open();
  }

  protected onAfterPropertyPaneChangesApplied() {
    //Poll Option setup
    if (!this._IspollOptionsSetupCompleted) {
      if (this.properties.pollDataCollection && this.properties.pollDataCollection.length > 0) {
        //Reach out to the web and create the list items
        let tempPollOption: IPollData[] = [];
        const pollID = this.properties.guid
        this.properties.pollDataCollection.forEach((poll: IPollOption) => {
          tempPollOption.push({
            Title: poll.option,
            PollID: pollID,
            Votes: 0
          });
        });

        this.context.statusRenderer.displayLoadingIndicator(this.domElement, "Please wait while we create Poll Options..");

        this.createPollOptions(tempPollOption).then(() => {
          this._IspollOptionsSetupCompleted = true;
          this.properties["pollSetupCompleted"] = true;
          this.context.propertyPane.refresh();
          this.context.statusRenderer.clearLoadingIndicator(this.domElement);
          this.render();
        }).catch((error: any) => {
          this.context.statusRenderer.renderError(this.domElement, error);
        });
      }
    }
  }

  protected createPollOptions = async (pollData: IPollData[]) => {
    const web = new Web(this.context.pageContext.web.absoluteUrl);
    const listGUID = this.properties.pollListGUID;

    for (var i = 0; i < pollData.length; i++) {
      await web.lists.getById(listGUID).items.add(pollData[i]);
      console.log("Poll Item Added : " + pollData[i].Title);
    }

    console.log("Poll Items Created");
  }


  protected guidGenerator = (): string => {
    return (this.S4() + this.S4() + "-" + this.S4() + "-4" + this.S4().substr(0, 3) + "-" + this.S4() + "-" + this.S4() + this.S4() + this.S4()).toLowerCase();
  }

  protected S4 = (): string => {
    return (((1 + Math.random()) * 0x10000) | 0).toString(16).substring(1);
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

  protected getPollListGUID = (): Promise<string> => {
    return new Promise<string>((resolve: (listGUID: string) => void, reject: (error: any) => void) => {
      this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$filter=(EntityTypeName eq 'QuickPollList')&$select=Id`, SPHttpClient.configurations.v1, {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'odata-version': ''
        }
      }).then((response: SPHttpClientResponse) => {
        return response.json();
      }).then((data: any) => {
        let dataArray: any[] = data.value;
        if (dataArray.length > 0) {
          resolve(dataArray[0].Id);
        }
        else {
          reject("Error while retrieveing Poll List ID");
        }
      }).catch((error: any) => {
        reject(error);
      });
    });
  }

  protected onPropertyPaneConfigurationStart() {
    if (!this.properties["guid"]) {
      this.properties["guid"] = this.guidGenerator();
      this.context.propertyPane.refresh();
    }

    if (!this.properties["pollSetupCompleted"]) {
      this.properties["pollSetupCompleted"] = false;
      this.context.propertyPane.refresh();
    }

    if (this.properties.pollDataCollection && this.properties.pollDataCollection.length > 0) {
      this._IspollOptionsSetupCompleted = true;
      this.context.propertyPane.refresh();
    }

    if (!this.properties.pollListGUID) {
      this.context.statusRenderer.displayLoadingIndicator(this.domElement, "Loading Configuration");

      this.getPollListGUID().then((listID: string) => {
        this.properties["pollListGUID"] = listID;
        console.log("Configuration Setup Completed - Poll GUID : " + this.properties.guid);
        this.context.statusRenderer.clearLoadingIndicator(this.domElement);
      }).catch((error: any) => {
        this.context.statusRenderer.renderError(this.domElement, "There was error in Loading Configurations, please contact admin");
        console.dir(error);
      });

      this.context.propertyPane.refresh();
    }

    this.render();
  }

  protected pollTitleValidator = (value: string): string => {
    if (value.trim().length > 0) {
      return '';
    }

    return "Poll title cannot be left blank";
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          displayGroupsAsAccordion: true,
          groups: [
            {
              groupName: strings.BasicGroupName,
              isCollapsed: false,
              groupFields: [
                PropertyPaneTextField('pollTitle', {
                  label: "Poll Title",
                  onGetErrorMessage: this.pollTitleValidator.bind(this)
                }),
                PropertyFieldCollectionData('pollDataCollection', {
                  key: "pollDataCollection",
                  label: "Poll Options",
                  manageBtnLabel: "Manage Poll Option",
                  panelHeader: "Add your poll options here",
                  fields: [
                    {
                      id: "option",
                      title: "Option",
                      placeholder: "Enter your poll option here",
                      type: CustomCollectionFieldType.string,
                      required: true
                    }
                  ],
                  value: this.properties.pollDataCollection,
                  disabled: this._IspollOptionsSetupCompleted
                }),
                PropertyPaneLabel('', {
                  text: "*Poll Options Setup can be done only once",
                })
              ]
            },
            {
              groupName: "Internal Use",
              isCollapsed: false,
              groupFields: [
                PropertyPaneLabel('', {
                  text: "This area is strictly used for Internal Processing of the webpart"
                }),
                PropertyPaneTextField('guid', {
                  label: "Poll GUID (internal property)",
                  disabled: true
                }),
                PropertyPaneTextField('pollListGUID', {
                  label: "Poll List GUID (internal property)",
                  disabled: true
                }),
                PropertyPaneToggle('pollSetupCompleted', {
                  label : "Poll Setup Completed?",
                  disabled : true,
                  offText : "No",
                  onText : "YES"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
