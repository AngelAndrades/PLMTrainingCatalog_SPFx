import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneHorizontalRule,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
//import { escape } from '@microsoft/sp-lodash-subset';

//import styles from './PlmTrainingCatalogWebPart.module.scss';
//import * as strings from 'PlmTrainingCatalogWebPartStrings';

// Kendo UI styles
import '@progress/kendo-ui/css/web/kendo.common-material.min.css';
import '@progress/kendo-ui/css/web/kendo.material.min.css';
import '@progress/kendo-ui/css/web/kendo.material.mobile.min.css';

//import * as $ from 'jquery';
//import '@progress/kendo-ui';
import * as pnp from '@pnp/sp/presets/all';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { SPA } from './app/spa';

export interface IPlmTrainingCatalogWebPartProps {
  catalogList: string;
  rolesList: string;
  maturityList: string;
  lnkSafe: string;
  lnkWAM: string;
  enableMaturity: boolean;
}

export default class PlmTrainingCatalogWebPart extends BaseClientSideWebPart<IPlmTrainingCatalogWebPartProps> {
  protected onInit(): Promise < void > {
    return super.onInit().then(_ => {
      pnp.sp.setup({
        spfxContext: this.context,
        sp: {
          headers: {
            Accept: 'application/json;odata=nometadata'
          }
        }
      });
    });
  }

  public render(): void {
    this.domElement.innerHTML = `<style>.k-tabstrip>.k-tabstrip-items>.k-item { text-transform: none; }</style>` +
                                `<div id="tabstrip"><ul><li class="k-state-active">By Role</li><li>DevSecOps</li><li>Recommended Reading</li><li>SAFe</li><li>Weekly Agile Meetings</li></ul>` +
                                // Roles tab
                                `<div><br /><div id="filters" style="margin: auto; padding: 1em; border: 1px solid lightgrey">` +
                                `<p style="font-weight: bold; margin-top: 0;">Filter Training Catalog By:</p>` +
                                `<input id="organization" style="width: 25vw; margin-bottom: 5px;" />` +
                                `<br/>` +
                                `<input id="team" disabled="disabled" style="width: 25vw; margin-bottom: 5px;" />` +
                                `<br/>` +
                                `<input id="role" disabled="disabled" style="width: 25vw; margin-bottom: 5px;" />` +
                                `<br/>` +
                                `<input id="byRoleMaturity" style="width: 25vw;" />` +
                                `</div>` +
                                `<br/>` +
                                `<div id="grid"></div></div>` + 
                                // DevSecOps tab
                                `<div><br /><div id="dsoFilters" style="margin: auto; padding: 1em; border: 1px solid lightgrey">` +
                                `<p style="font-weight: bold; margin-top: 0;">Filter Training Catalog By:</p>` +
                                `<input id="dsoMaturity" style="width: 25vw;" />` +
                                `</div>` +
                                `<br/>` +
                                `<div id="grid2"></div></div>` + 
                                // Recommended Reading tab
                                `<div><div id="grid3"></div></div>` + 
                                // SAFe tab
                                `<div></div>` +
                                // WAM tab
                                `<div></div></div>` +
                                `<div id="dialog"></div>`;
    
    const app = SPA.getInstance({
      catalogGuid: this.properties.catalogList,
      rolesGuid: this.properties.rolesList,
      maturityGuid: this.properties.maturityList,
      safeLink: this.properties.lnkSafe,
      wamLink: this.properties.lnkWAM,
      enableMaturity: this.properties.enableMaturity
    });
  }

  //@ts-expect-error
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  //@ts-expect-error
  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'Provide the necessary values below to drive this web part'
          },
          groups: [
            {
              groupFields: [
                PropertyFieldListPicker('catalogList', {
                  label: 'Select the training catalog list',
                  selectedList: this.properties.catalogList,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'listPickerFieldId'
                }),
                PropertyFieldListPicker('rolesList', {
                  label: 'Select the roles list',
                  selectedList: this.properties.rolesList,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'listPickerFieldId'
                }),
                PropertyFieldListPicker('maturityList', {
                  label: 'Select the maturity list',
                  selectedList: this.properties.maturityList,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'listPickerFieldId'
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneTextField('lnkSafe', {
                  label: 'Paste the SAFe web page URL'
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneTextField('lnkWAM', {
                  label: 'Paste the Weekly Agile Meeting web page URL'
                }),
                PropertyPaneToggle('enableMaturity', {
                  label: 'Enable Maturity Level Filter',
                  checked: false
                }
                )
              ]
            }
          ]
        }
      ]
    };
  }
}
