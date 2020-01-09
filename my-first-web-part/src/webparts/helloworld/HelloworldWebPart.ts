import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import {
  PropertyPaneCheckbox,
  PropertyPaneLabel,
  PropertyPaneLink,
  PropertyPaneSlider,
  PropertyPaneToggle,
} from '@microsoft/sp-webpart-base';
// import styles from './ListItemsForm.module.scss';
import $ from "jquery";
import pnp from "sp-pnp-js";
import * as strings from 'HelloworldWebPartStrings';
import Helloworld from './components/Helloworld';
import { IHelloworldProps } from './components/IHelloworldProps';
import { IStackTokens, Stack } from 'office-ui-fabric-react/lib/Stack';
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import {
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';
import {
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';

export interface IHelloworldWebPartProps {
  description: string;
  name: string;
  state: string;
}

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Id: string;
  Title: string;
}

export default class HelloworldWebPart extends BaseClientSideWebPart<IHelloworldWebPartProps> {

  private _getListData(): Promise<ISPLists> {
    console.log(this.context.pageContext.web.absoluteUrl);
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/lists/GetByTitle('kendoGrid')/Items", SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        console.log(response);
        return response.json();
      });
  }

  private _dropdownOptions: IPropertyPaneDropdownOption[] = [];
  public onInit<T>(): Promise<T> {
    this.test();
    this._getListData()
      .then((response) => {
        this._dropdownOptions = response.value.map((list: ISPList) => {
          return {
            key: list.Title,
            text: list.Title
          };
        });
      });
    return Promise.resolve();
  }

  public render(): void {
    const element: React.ReactElement<IHelloworldProps> = React.createElement(
      Helloworld,
      {
        description: this.properties.description,
        name: this.properties.name,
        state: this.properties.state,
        context: this.context,
      }
    );
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneTextField('name', {
                  label: strings.nameFieldLabel
                }),
                PropertyPaneDropdown('state', {
                  label: strings.stateFieldLabel,
                  options: this._dropdownOptions
                  // options: [
                  //   { key: 'Javascript', text: 'Javascript'},
                  //   { key: 'Angular Js', text: 'Angular JS' },
                  //   { key: 'React Js', text: 'React Js' },
                  //   { key: 'Node Js', text: 'Node Js'}
                  // ],
                })
              ]
            }
          ]
        }
      ]
    };
  }
  private test() {
    pnp.sp.web.lists.getByTitle("kendoGrid").get().then(function (response) {
      console.log(response);
    })
  }
}