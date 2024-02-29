import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'ManagementControlWebPartStrings';
import ManagementControl from './components/ManagementControl';
import { IManagementControlProps } from './components/IManagementControlProps';
import {
  PropertyFieldListPicker,
  PropertyFieldListPickerOrderBy,
} from "@pnp/spfx-property-controls/lib/PropertyFieldListPicker";

const { solution } = require("../../../config/package-solution.json");

export interface IManagementControlWebPartProps {
  Title: string;
  ProductsListId: string;
  Cell1: string;
  width1: number;
  Cell2: string;
  width2: number;
  Cell3: string;
  width3: number;
  Cell4: string;
  width4: number;
  Cell5: string;
  width5: number;
  Cell6: string;
  width6: number;
  Cell7: string;
  width7: number;
  Cell8: string;
  width8: number;
  Cell9: string;
  width9: number;
  Cell10: string;
  width10: number;
  Cell11: string;
  width11: number;
  Cell12: string;
  width12: number;
  Cell13: string;
  width13: number;

}

export default class ManagementControlWebPart extends BaseClientSideWebPart<IManagementControlWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IManagementControlProps> = React.createElement(
      ManagementControl,
      {
        Title: this.properties.Title,
        ProductsListId: this.properties.ProductsListId,
        context: this.context,
        Cell1: this.properties.Cell1,
        width1: this.properties.width1,
        Cell2: this.properties.Cell2,
        width2: this.properties.width2,
        Cell3: this.properties.Cell3,
        width3: this.properties.width3,
        Cell4: this.properties.Cell4,
        width4: this.properties.width4,
        Cell5: this.properties.Cell5,
        width5: this.properties.width5,
        Cell6: this.properties.Cell6,
        width6: this.properties.width6,
        Cell7: this.properties.Cell7,
        width7: this.properties.width7,
        Cell8: this.properties.Cell8,
        width8: this.properties.width8,
        Cell9: this.properties.Cell9,
        width9: this.properties.width9,
        Cell10: this.properties.Cell10,
        width10: this.properties.width10,
        Cell11: this.properties.Cell11,
        width11: this.properties.width11,
        Cell12: this.properties.Cell12,
        width12: this.properties.width12,
        Cell13: this.properties.Cell13,
        width13: this.properties.width13,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse(solution.version);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupFields: [
                PropertyPaneTextField("Title", {
                  label: "Title",
                }),
                PropertyFieldListPicker("ProductsListId", {
                  label: "Select Products list",
                  selectedList: this.properties.ProductsListId,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context as any,
                  deferredValidationTime: 0,
                  key: "listPickerFieldId",
                }),
              ]
            },
            {
              groupFields: [
                PropertyPaneTextField("Cell1", {
                  label: "Cell1 Name",
                }),
                PropertyPaneTextField("Cell2", {
                  label: "Cell2 Name",
                }),
                PropertyPaneTextField("Cell3", {
                  label: "Cell3 Name",
                }),
                PropertyPaneTextField("Cell4", {
                  label: "Cell4 Name",
                }),
                PropertyPaneTextField("Cell5", {
                  label: "Cell5 Name",
                }),
                PropertyPaneTextField("Cell6", {
                  label: "Cell6 Name",
                }),
                PropertyPaneTextField("Cell7", {
                  label: "Cell7 Name",
                }),
                PropertyPaneTextField("Cell8", {
                  label: "Cell8 Name",
                }),
                PropertyPaneTextField("Cell9", {
                  label: "Cell9 Name",
                }),
                PropertyPaneTextField("Cell10", {
                  label: "Cell10 Name",
                }),
                PropertyPaneTextField("Cell11", {
                  label: "Cell11 Name",
                }),
                PropertyPaneTextField("width1", {
                  label: "width1 Name",
                }),
                PropertyPaneTextField("width2", {
                  label: "width2 Name",
                }),
                PropertyPaneTextField("width3", {
                  label: "width3 Name",
                }),
                PropertyPaneTextField("width4", {
                  label: "width4 Name",
                }),
                PropertyPaneTextField("width5", {
                  label: "width5 Name",
                }),
                PropertyPaneTextField("width6", {
                  label: "width6 Name",
                }),
                PropertyPaneTextField("width7", {
                  label: "width7 Name",
                }),
                PropertyPaneTextField("width8", {
                  label: "width8 Name",
                }),
                PropertyPaneTextField("width9", {
                  label: "width9 Name",
                }),
                PropertyPaneTextField("width10", {
                  label: "width10 Name",
                }),
                PropertyPaneTextField("width11", {
                  label: "width11 Name",
                }),
                PropertyPaneTextField("width12", {
                  label: "width12 Name",
                }),
                PropertyPaneTextField("width13", {
                  label: "width13 Name",
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
