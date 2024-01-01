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

  //* Dev
  DevAntibodyLabellingId: string;
  DevAntibodyPurificationId: string;
  DevBlockingPeptidepreparationId: string;
  DevColumnPreparationId: string;
  DevColumnPreparationForFusionPeptideId: string;
  DevFusionBlockingPeptidePreparationId: string;
  //* QC
  QcDirectFlowCytometryId: string;
  QcBlockingPeptideWesternBlotQcId: string;
  QCAntibodyWesternBlotQcId: string;
  QcImmunohistochemistryId: string;
  //* Application
  AppIndirectFlowCytometryId: string;
  AppImmunohistochemistryId: string;  
}

export default class ManagementControlWebPart extends BaseClientSideWebPart<IManagementControlWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IManagementControlProps> = React.createElement(
      ManagementControl,
      {
        Title: this.properties.Title,
        ProductsListId: this.properties.ProductsListId,
        context: this.context,

        //* Dev
        DevAntibodyLabellingId: this.properties.DevAntibodyLabellingId,
        DevAntibodyPurificationId: this.properties.DevAntibodyPurificationId,
        DevBlockingPeptidepreparationId:
          this.properties.DevBlockingPeptidepreparationId,
        DevColumnPreparationId: this.properties.DevColumnPreparationId,
        DevColumnPreparationForFusionPeptideId:
          this.properties.DevColumnPreparationForFusionPeptideId,
        DevFusionBlockingPeptidePreparationId:
          this.properties.DevFusionBlockingPeptidePreparationId,
        //* QC
        QcDirectFlowCytometryId: this.properties.QcDirectFlowCytometryId,
        QcBlockingPeptideWesternBlotQcId:
          this.properties.QcBlockingPeptideWesternBlotQcId,
        QCAntibodyWesternBlotQcId: this.properties.QCAntibodyWesternBlotQcId,
        QcImmunohistochemistryId: this.properties.QcImmunohistochemistryId,
        //* App
        AppImmunohistochemistryId: this.properties.AppImmunohistochemistryId,
        AppIndirectFlowCytometryId: this.properties.AppIndirectFlowCytometryId,        
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
              groupName: "Development Lists",
              groupFields: [
                PropertyFieldListPicker("DevAntibodyPurificationId", {
                  label: "AntibodyPurification list",
                  selectedList: this.properties.DevAntibodyPurificationId,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context as any,
                  deferredValidationTime: 0,
                  key: "listPickerFieldId",
                }),
                PropertyFieldListPicker("DevAntibodyLabellingId", {
                  label: "AntibodyLabelling list",
                  selectedList: this.properties.DevAntibodyLabellingId,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context as any,
                  deferredValidationTime: 0,
                  key: "listPickerFieldId",
                }),
                PropertyFieldListPicker("DevBlockingPeptidepreparationId", {
                  label: "BlockingPeptidePreparation list",
                  selectedList: this.properties.DevBlockingPeptidepreparationId,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context as any,
                  deferredValidationTime: 0,
                  key: "listPickerFieldId",
                }),
                PropertyFieldListPicker("DevColumnPreparationId", {
                  label: "ColumnPreparation list",
                  selectedList: this.properties.DevColumnPreparationId,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context as any,
                  deferredValidationTime: 0,
                  key: "listPickerFieldId",
                }),
                PropertyFieldListPicker(
                  "DevColumnPreparationForFusionPeptideId",
                  {
                    label: "ColumnPreparationForFusionPeptide list",
                    selectedList:
                      this.properties.DevColumnPreparationForFusionPeptideId,
                    includeHidden: false,
                    orderBy: PropertyFieldListPickerOrderBy.Title,
                    disabled: false,
                    onPropertyChange:
                      this.onPropertyPaneFieldChanged.bind(this),
                    properties: this.properties,
                    context: this.context as any,
                    deferredValidationTime: 0,
                    key: "listPickerFieldId",
                  }
                ),
                PropertyFieldListPicker(
                  "DevFusionBlockingPeptidePreparationId",
                  {
                    label: "FusionBlockingPeptidePreparation list",
                    selectedList:
                      this.properties.DevFusionBlockingPeptidePreparationId,
                    includeHidden: false,
                    orderBy: PropertyFieldListPickerOrderBy.Title,
                    disabled: false,
                    onPropertyChange:
                      this.onPropertyPaneFieldChanged.bind(this),
                    properties: this.properties,
                    context: this.context as any,
                    deferredValidationTime: 0,
                    key: "listPickerFieldId",
                  }
                ),
              ],
            },
            {
              groupName: "QC Lists",
              groupFields: [
                PropertyFieldListPicker("QcDirectFlowCytometryId", {
                  label: "DirectFlowCytometry list",
                  selectedList: this.properties.QcDirectFlowCytometryId,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context as any,
                  deferredValidationTime: 0,
                  key: "listPickerFieldId",
                }),
                PropertyFieldListPicker("QcBlockingPeptideWesternBlotQcId", {
                  label: "BlockingPeptideWesternBlotQc list",
                  selectedList:
                    this.properties.QcBlockingPeptideWesternBlotQcId,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context as any,
                  deferredValidationTime: 0,
                  key: "listPickerFieldId",
                }),
                PropertyFieldListPicker("QCAntibodyWesternBlotQcId", {
                  label: "AntibodyWesternBlotQc list",
                  selectedList: this.properties.QCAntibodyWesternBlotQcId,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context as any,
                  deferredValidationTime: 0,
                  key: "listPickerFieldId",
                }),
                PropertyFieldListPicker("QcImmunohistochemistryId", {
                  label: "Immunohistochemistry list",
                  selectedList: this.properties.QcImmunohistochemistryId,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context as any,
                  deferredValidationTime: 0,
                  key: "listPickerFieldId",
                }),
              ],
            },
            {
              groupName: "Applications Lists",
              groupFields: [
                PropertyFieldListPicker("AppIndirectFlowCytometryId", {
                  label: "IndirectFlowCytometry list",
                  selectedList: this.properties.AppIndirectFlowCytometryId,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context as any,
                  deferredValidationTime: 0,
                  key: "listPickerFieldId",
                }),

                PropertyFieldListPicker("AppImmunohistochemistryId", {
                  label: "Immunohistochemistry list",
                  selectedList: this.properties.AppImmunohistochemistryId,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context as any,
                  deferredValidationTime: 0,
                  key: "listPickerFieldId",
                }),
              ],
            },
            
          ]
        }
      ]
    };
  }
}
