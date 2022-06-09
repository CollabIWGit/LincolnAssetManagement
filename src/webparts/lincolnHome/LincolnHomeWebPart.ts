import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { Navigation } from 'spfx-navigation';
import * as strings from 'LincolnHomeWebPartStrings';
import * as $ from "jquery";

import { navUtils } from '../../utils/navUtils';
let NavUtils = new navUtils();

import { navbar } from '../../utils/navbar';
let Navbar = new navbar();

require('../../../node_modules/bootstrap/dist/css/bootstrap.min.css');
require('../../../node_modules/bootstrap/js/src/collapse.js');
require('../../../node_modules/@fortawesome/fontawesome-free/css/all.min.css');
require('../../styles/spcommon.css');
require('../../styles/style.css');
require('../../styles/test.css');

import * as commonConfig from "../../utils/commonConfig.json";

export interface ILincolnHomeWebPartProps {
  description: string;
}

export default class LincolnHomeWebPart extends BaseClientSideWebPart<ILincolnHomeWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
    <div id="wrapper" class="">
      ${Navbar.navbar}
      <!-- Page Content  -->
      <div id="page-content-wrapper">
        <div class="container-fluid">
          <div class="row">
            <div id="content" class="p-4 p-md-5 pt-5">
              <h2 class="mb-4">Home</h2>
              <div class="mdc-layout-grid">
                <div class="mdc-layout-grid__inner">
                  <div class="mdc-layout-grid__cell stretch-card mdc-layout-grid__cell--span-8">
                    <div class="mdc-layout-grid__inner">
                      <div id="caseMgt" class="mdc-layout-grid__cell stretch-card mdc-layout-grid__cell--span-6">
                        <div class="mdc-card py-3 pl-2 d-flex flex-row align-item-center dashboardCard">
                          <div class="mdc--tile mdc--tile-primary rounded">
                            <i class="fas fa-file-contract fa-fw text-white icon-md"></i>
                          </div>
                          <div class="text-wrapper pl-1">
                            <h6 class="mdc-typography--display1 font-weight-bold mb-1">Case Management</h6>
                          </div>
                        </div>
                      </div>
                      <div id="assetMgt" class="mdc-layout-grid__cell stretch-card mdc-layout-grid__cell--span-6">
                        <div class="mdc-card py-3 pl-2 d-flex flex-row align-item-center dashboardCard">
                          <div class="mdc--tile mdc--tile-warning rounded">
                            <i class="fas fa-folder-open fa-fw text-white icon-md"></i>
                          </div>
                          <div class="text-wrapper pl-1">
                            <h6 class="mdc-typography--display1 font-weight-bold mb-1">Asset Management</h6>
                          </div>
                        </div>
                      </div>
                      <div id="userMgt" class="mdc-layout-grid__cell stretch-card mdc-layout-grid__cell--span-6">
                        <div class="mdc-card py-3 pl-2 d-flex flex-row align-item-center dashboardCard">
                          <div class="mdc--tile mdc--tile-success rounded">
                            <i class="fa fa-user fa-fw text-white icon-md"></i>
                          </div>
                          <div class="text-wrapper pl-1">
                            <h6 class="mdc-typography--display1 font-weight-bold mb-1">User Management</h6>
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>`;
    
    this.setButtonsEventHandlers();
    NavUtils.collapse();
    NavUtils.navTriggers();
    NavUtils.cover();
  }

  private setButtonsEventHandlers(): void {
		document.getElementById('caseMgt').addEventListener('click', () => this._navigateToCaseMgt());
    document.getElementById('assetMgt').addEventListener('click', () => this._navigateToAssetMgt());
    document.getElementById('userMgt').addEventListener('click', () => this._navigateToUserMgt());
	}

  private _navigateToCaseMgt() {
    Navigation.navigate(`${this.context.pageContext.web.absoluteUrl}/SitePages/${commonConfig.Page.CaseList}`, true);
  }

  private _navigateToAssetMgt() {
    Navigation.navigate(`${this.context.pageContext.web.absoluteUrl}/SitePages/${commonConfig.Page.AssetList}`, true);
  }

  private _navigateToUserMgt() {
    Navigation.navigate(`${this.context.pageContext.web.absoluteUrl}/SitePages/${commonConfig.Page.UsersList}`, true);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
