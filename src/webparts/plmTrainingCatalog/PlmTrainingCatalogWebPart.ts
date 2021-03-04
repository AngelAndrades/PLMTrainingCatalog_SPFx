import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneHorizontalRule,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

//import styles from './PlmTrainingCatalogWebPart.module.scss';
import * as strings from 'PlmTrainingCatalogWebPartStrings';

import * as $ from 'jquery';
import '@progress/kendo-ui';
import { SPComponentLoader } from '@microsoft/sp-loader';
import * as pnp from '@pnp/sp/presets/all';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { SPA } from './app/spa';

export interface IPlmTrainingCatalogWebPartProps {
  catalog: string;
  roles: string;
  lnkSafe: string;
  lnkWAM: string;
  maturity: boolean;
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
    SPComponentLoader.loadCss('https://kendo.cdn.telerik.com/2020.3.1118/styles/kendo.common-material.min.css');
    SPComponentLoader.loadCss('https://kendo.cdn.telerik.com/2020.3.1118/styles/kendo.material.min.css');
    //SPComponentLoader.loadCss('https://kendo.cdn.telerik.com/2020.3.1118/styles/kendo.bootstrap-v4.min.css');

    SPComponentLoader.loadScript('https://kendo.cdn.telerik.com/2020.3.1118/js/jszip.min.js');
    //SPComponentLoader.loadScript('https://kendo.cdn.telerik.com/2020.3.1118/js/kendo.all.min.js');

    this.domElement.innerHTML = `<style>.k-tabstrip>.k-tabstrip-items>.k-item { text-transform: none; }</style>` +
                                `<div id="tabstrip"><ul><li class="k-state-active">By Role</li><li>DevSecOps</li><li>SAFe</li><li>Weekly Agile Meetings</li></ul>` +
                                // Roles tab
                                `<div><br /><div id="filters" style="margin: auto; padding: 1em; border: 1px solid lightgrey">` +
                                `<p style="font-weight: bold; margin-top: 0;">Filter Training Catalog By:</p>` +
                                `<input id="organization" style="width: 25vw; margin-bottom: 5px;" />` +
                                `<br/>` +
                                `<input id="team" disabled="disabled" style="width: 25vw; margin-bottom: 5px;" />` +
                                `<br/>` +
                                `<input id="role" disabled="disabled" style="width: 25vw;" />` +
                                `</div>` +
                                `<br/>` +
                                `<div id="grid"></div></div>` + 
                                // DevSecOps tab
                                `<div><div id="grid2"></div></div>` + 
                                // SAFe tab
                                `<div></div>` +
                                // WAM tab
                                `<div></div></div>` +
                                `<div id="dialog"></div>`;
    
    const app = SPA.getInstance({
      catalogGuid: this.properties.catalog,
      rolesGuid: this.properties.roles,
      safeLink: this.properties.lnkSafe,
      wamLink: this.properties.lnkWAM,
      maturity: this.properties.maturity
    });
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
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
                PropertyFieldListPicker('catalog', {
                  label: 'Select the training catalog list',
                  selectedList: this.properties.catalog,
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
                PropertyFieldListPicker('roles', {
                  label: 'Select the roles list',
                  selectedList: this.properties.roles,
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
                PropertyPaneToggle('maturity', {
                  label: 'Hide Maturity Level'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
