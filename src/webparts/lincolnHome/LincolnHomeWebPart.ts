import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { Navigation } from 'spfx-navigation';
import styles from './LincolnHomeWebPart.module.scss';
import * as strings from 'LincolnHomeWebPartStrings';
import * as $ from "jquery";

require('../../../node_modules/bootstrap/dist/css/bootstrap.min.css');
require('../../../node_modules/@fortawesome/fontawesome-free/css/all.min.css');
require('../../styles/spcommon.css');
require('../../styles/navbar.css');
require('../../styles/style.css');

import * as commonConfig from "../../utils/commonConfig.json";
import { sidebarDetails } from "../../utils/sidebarDetails";
let SidebarDetails = new sidebarDetails();

export interface ILincolnHomeWebPartProps {
  description: string;
}

export default class LincolnHomeWebPart extends BaseClientSideWebPart<ILincolnHomeWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
    <div class="wrapper d-flex align-items-stretch">
      <div class="nav-placeholder">
        <nav id="sidebar">
          <div class="custom-menu">
            <button type="button" id="sidebarCollapse" class="btn btn-primary">
              <i class="fa fa-bars"></i>
              <span class="sr-only">Toggle Menu</span>
          </div>
          <img id="imgLogo" src="${this.context.pageContext.web.absoluteUrl}/SiteAssets/Lincoln-Realty-Logo-orange.png"
            alternate="lincoln-logo">
          <ul class="list-unstyled components mb-5">
            <li class="active">
              <a id="home">
                <span class="fa fa-home mr-3"> </span>Home
              </a>
            </li>
            <li id="adminMgtComponent">
              <a id="adminMgt">
                <span class="fa fa-sliders-h mr-3"> </span>Admin Management
              </a>
            </li>
            <li>
              <a id="CaseMgt">
                <span class="fas fa-file-contract mr-3"> </span>Case Management
              </a>
              <div class="collapse1 collapse">
                <ul style="list-style-type:none;" id="caseManagementUl">
                  <li>
                    <a id="caseList">
                      <span class="fa fa-list"> </span>  List of Case
                    </a>
                  </li>
                  <li>
                    <a id="addCase">
                      <span class="fa fa-plus"> </span>  Add new Case
                    </a>
                  </li>
                </ul>
              </div>
            </li>  
            <li>
              <a id="AssetMgt">
                <span class="fas fa-folder-open mr-3"></span>Asset Management
              </a>
              <div class="collapse2 collapse">
                <ul style="list-style-type:none;" id="assetManagementUl">
                  <li>
                    <a id="assetList">
                      <span class="fa fa-list"> </span>  List of Assets
                    </a>
                  </li>
                  <li>
                    <a id="addAsset">
                      <span class="fa fa-plus"> </span>  Add new Asset
                    </a>
                  </li>
                </ul>
              </div>
            </li>
          </ul>
        </nav>
      </div>
      <!-- Page Content  -->
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
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>`;
    // SidebarDetails.sidebarMenu(this.context.pageContext.web.absoluteUrl);
    this.setButtonsEventHandlers();
    this.collapse();
    this.navTriggers();
  }

  private setButtonsEventHandlers(): void {
		document.getElementById('caseMgt').addEventListener('click', () => this._navigateToCaseMgt());
    document.getElementById('assetMgt').addEventListener('click', () => this._navigateToAssetMgt());
	}

  private _navigateToCaseMgt() {
    Navigation.navigate(`${this.context.pageContext.web.absoluteUrl}/SitePages/${commonConfig.Page.CaseDashboard}`, true);
  }

  private _navigateToAssetMgt() {
    Navigation.navigate(`${this.context.pageContext.web.absoluteUrl}/SitePages/${commonConfig.Page.AssetDashboard}`, true);
  }

  public navTriggers() {
    $('#caseList').on('click', () => {
      Navigation.navigate(`${this.context.pageContext.web.absoluteUrl}/SitePages/${commonConfig.Page.CaseDashboard}`, true);
    });

    $('#addCase').on('click', () => {
      Navigation.navigate(`${this.context.pageContext.web.absoluteUrl}/SitePages/${commonConfig.Page.AddCase}`, true);
    });

    $('#home').on('click', () => {
      Navigation.navigate(`${this.context.pageContext.web.absoluteUrl}/SitePages/${commonConfig.Page.HomePage}`, true);
    });

    $('#addAsset').on('click', () => {
      Navigation.navigate(`${this.context.pageContext.web.absoluteUrl}/SitePages/${commonConfig.Page.AddAssets}`, true);
    });

    $('#assetList').on('click', () => {
      Navigation.navigate(`${this.context.pageContext.web.absoluteUrl}/SitePages/${commonConfig.Page.AssetDashboard}`, true);
    });
  }

  private collapse() {
    $("#CaseMgt").hover(
      () => {
        (<any>$(".collapse1")).show();
      },
      () => {
        (<any>$(".collapse1")).hide();
      }
    );

    $("#AssetMgt").hover(
      () => {
        (<any>$(".collapse2")).show();
      },
      () => {
        (<any>$(".collapse2")).hide();
      }
    );

    $(".collapse2").hover(
      () => {
        (<any>$(".collapse2")).show();
      },
      () => {
        (<any>$(".collapse2")).hide();
      }
    );

    $(".collapse1").hover(
      () => {
        (<any>$(".collapse1")).show();
      },
      () => {
        (<any>$(".collapse1")).hide();
      }
    );

    $("#btnfd").click(() => {
      (<any>$(".collapsecard")).slideToggle(500);
    });

    $("#sidebarCollapse").click(() => {
      (<any>$("#sidebar")).slideToggle(500);
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
