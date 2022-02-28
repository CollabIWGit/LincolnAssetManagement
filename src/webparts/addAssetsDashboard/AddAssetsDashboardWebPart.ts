import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { Navigation } from 'spfx-navigation';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import * as strings from 'AddAssetsDashboardWebPartStrings';
import * as $ from "jquery";
import 'datatables.net';
import 'datatables.net-dt/css/jquery.dataTables.css';

require('../../../node_modules/bootstrap/dist/css/bootstrap.min.css');
require('../../../node_modules/@fortawesome/fontawesome-free/css/all.min.css');
require('../../styles/dashboardcss.css');
require('../../styles/spcommon.css');
require('../../styles/navbar.css');

import * as commonConfig from "../../utils/commonConfig.json";
import { sidebarDetails } from "../../utils/sidebarDetails";
let SidebarDetails = new sidebarDetails();

//#region Interfaces
export interface IAddAssetsDashboardWebPartProps {
  description: string;
}

export interface IApplicationDetailsList {
  Name: string;
  ReferenceNumber: string;
  BuildingName: string;
  FloorNo: string;
  Ownership: string;
  TypeOfAsset: string;
  ServicingRequired: boolean;
  LastServicingDate: string;
  ServicingPeriod: string;
  Comment: string;
  AttachmentFileName: string;
  AttachmentFileContent: string;
}

export interface IDynamicField extends IApplicationDetailsList {
  [key: string]: any;
}

export interface ITypeOfAssetLists {
  value: ITypeOfAssetList[];
}

export interface ITypeOfAssetList {
  Title: string;
}

export interface IDropdownLists {
  value: IDropdownList[];
}

export interface IDropdownList {
  Title: string;
}

export interface IFieldsRequiredLists {
  value: IFieldsRequiredList[];
}

export interface IFieldsRequiredList {
  Title: string;
  TypeOfAssets: { Title: string, Description: string };
  FieldType: string;
  DropdownListName: string;
  Required: boolean;
}

export interface IBuildings {
  ID: number;
  Title: string;
  Location: string;
  ShortForm: string;
}

export interface IOffices {
  Title: string;
  FloorNumber: number;
  BuildingIDId: number;
  ID: number;
  ShortForm: string;
}
//#endregion

export default class AddAssetsDashboardWebPart extends BaseClientSideWebPart<IAddAssetsDashboardWebPartProps> {
  private static accessToken: string = "";
  private ListOfBuildings: IBuildings[];
  private ListOfOffices: IOffices[];

  public render(): void {
    this.domElement.innerHTML = `
    <div id="cover"> <span class="glyphicon glyphicon-refresh w3-spin preloader-Icon"></span> loading...</div>
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
      <div class="container">
        <div class="inner-container">
          <div class="form-row">
            <div class="col-md-12">
              <h3>Asset Management Dashboard</h3>
            </div>
          </div>
          <div class="filters">
            <div class="form-row">
              <div class="col-md-6">
                <div class="input-group input-group-sm mb-3" id="AssetName">
                  <div class="input-group-prepend">
                    <span class="input-group-text" id="inputGroup-sizing-sm">Asset Name</span>
                  </div>
                  <input type="text" class="form-control" id="idAssetName" autocomplete="off"/>
                </div>
              </div>
              <div class="col-md-6">
                <div class="input-group input-group-sm mb-3" id="AssetRefNo">
                  <div class="input-group-prepend">
                    <span class="input-group-text" id="inputGroup-sizing-sm">Asset Ref No</span>
                  </div>
                  <input type="text" class="form-control" id="idAssetRefNo" autocomplete="off"/>
                </div>
              </div>
            </div>
            <div class="form-row">
              <div class="col-md-6">
                <div class="input-group input-group-sm mb-3" id="TypeOfAssets">
                  <div class="input-group-prepend">
                    <span class="input-group-text" id="inputGroup-sizing-sm">Type of Assets</span>
                  </div>
                  <input type="text" class="form-control" id="idTypeOfAssets" autocomplete="off"/>
                </div>
              </div>
              <div class="col-md-6">
                <div class="input-group input-group-sm mb-3" id="Office">
                  <div class="input-group-prepend">
                    <span class="input-group-text" id="inputGroup-sizing-sm">Office</span>
                  </div>
                  <input type="text" class="form-control" id="idOffice" autocomplete="off"/>
                </div>
              </div>
            </div>
          </div>
          <div id="divContainer">
          </div>
        </div>
      </div>
    </div>`;
    $("#cover").fadeOut(1750);
    // SidebarDetails.sidebarMenu(this.context.pageContext.web.absoluteUrl);
    this._getAssetsAsync();
    this._getBuildingsList();
    this._getOfficesList();
    this.collapse();
    this.navTriggers();
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

  private _getAccessToken() {
    var body = {
      grant_type: 'password',
      client_id: 'myClientId',
      client_secret: 'myClientSecret',
      username: "roukaiyan@frci.net",
      password: "Pa$$w0rd"
    };

    return $.ajax({
      type: 'POST',
      url: commonConfig.baseUrl + '/token',
      dataType: 'json',
      data: body,
      contentType: 'application/x-www-form-urlencoded'
    }).then((response) => {
      AddAssetsDashboardWebPart.accessToken = response["access_token"];
      return AddAssetsDashboardWebPart.accessToken;
    });
  }

  private async _getAssets() {
    let token = await this._getAccessToken();
    return $.ajax({
      type: 'GET',
      url: commonConfig.baseUrl + '/api/Asset/GetAssets',
      headers: {
        Authorization: 'Bearer ' + token
      }
    }).then((response) => {
      return response;
    });
  }

  private async _getAssetsAsync() {
    var assets = await this._getAssets();
    this._renderTable(assets);
    this._renderTableAsync();
  }

  private _getBuildingsList() {
    this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${commonConfig.List.BuildingList}')/items`, SPHttpClient.configurations.v1)
      .then(response => {
        return response.json()
          .then((items: any): void => {
            this.ListOfBuildings = items.value;
          });
      });
  }

  private _getOfficesList() {
    let html: string = '';
    this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${commonConfig.List.OfficeList}')/items`, SPHttpClient.configurations.v1)
      .then(response => {
        return response.json()
          .then((items: any): void => {
            this.ListOfOffices = items.value;
          });
      });
  }

  private _deleteAsset(refNo: string): void {
    $.ajax({
      type: 'DELETE',
      url: commonConfig.baseUrl + `/api/Asset/delete/${refNo}`,
      headers: {
        Authorization: 'Bearer ' + AddAssetsDashboardWebPart.accessToken
      },
      dataType: 'json',
      contentType: 'application/json',
      success: (result) => {
        console.log("success " + result);
      },
      error: (result) => {
        console.log("error " + result);
      }
    });
  }

  private _renderTable(listOfAssets: IDynamicField[]) {
    var officeName: string = "";
    let html: string = `<table id="tbl_asset_list" class="table table-striped">
      <thead>
        <tr>
          <th class="text-left">Asset Name</th>
          <th class="text-left">Asset Reference No</th>
          <th class="text-left">Type of Assets</th>
          <th class="text-left">Office</th>
          <th class="text-center">View</th>
          <th class="text-center">Delete</th>
        </tr>
      </thead>
      <tbody id="tb_asset_list">`;
    listOfAssets.forEach((item: IDynamicField) => {
      this.ListOfBuildings.forEach((buildingItem: IBuildings) => {
        if (item.BuildingName == buildingItem.Title) {
          this.ListOfOffices.forEach((officeItem: IOffices) => {
            if (officeItem.BuildingIDId == buildingItem.ID && item.FloorNo == officeItem.FloorNumber.toString()) {
              officeName = officeItem.Title;
            }
          });
        }
      });
      html += `
        <tr>
          <td class="text-left">${item.Name}</td>
          <td class="text-left">${item.ReferenceNumber}</td>
          <td class="text-left">${item.TypeOfAsset}</td>
          <td class="text-left">${officeName}</td>
          <td class="text-center view">                
            <button class="btn btn-sm rounded-circle" id="btn_${item.ReferenceNumber}_View" type="button"><i class="fa fa-eye"></i></button>
          </td>
          <td class="text-center delete">                
            <button class="btn btn-sm rounded-circle" id="btn_${item.ReferenceNumber}_Delete" type="button"><i class="fa fa-trash"></i></button>
          </td>
        </tr>`;
    });
    html += `</tbody>
    </table>`;

    const listContainer: Element = this.domElement.querySelector('#divContainer');
    listContainer.innerHTML = html;
  }

  private _renderTableAsync() {
    try {
      var table = $('#tbl_asset_list').DataTable({
        paging: true,
        info: true,
        language: {
          searchPlaceholder: "Search assets",
          search: "",
        },
        responsive: true,
        columnDefs: [
          { orderable: false, targets: [4, 5] }
        ],
        order: [[0, "asc"]]
      });

      $('#AssetName').on('keyup', 'input', function () {
        table
          .columns(0)
          .search(this.value)
          .draw();
      });
      $('#AssetRefNo').on('keyup', 'input', function () {
        table
          .columns(1)
          .search(this.value)
          .draw();
      });
      $('#TypeOfAssets').on('keyup', 'input', function () {
        table
          .columns(2)
          .search(this.value)
          .draw();
      });
      $('#Office').on('keyup', 'input', function () {
        table
          .columns(3)
          .search(this.value)
          .draw();
      });

      //Click view btn
      $(".view").on('click', 'button', function (){
        var data = table.row($(this).parents('tr')).data();
        var refNo = data[1];
        var url = new URL(`https://frcidevtest.sharepoint.com/sites/Lincoln/SitePages/${commonConfig.Page.AddAssets}`);
        url.searchParams.append('refNo',refNo);
        Navigation.navigate(url.toString(), true);
      });

      //Click delete btn
      $(".delete").on('click', 'button', function (){
        var data = table.row($(this).parents('tr')).data();
        var refNo = data[1];
        $.ajax({
          type: 'DELETE',
          data: {'action': 'delete'},
          url: commonConfig.baseUrl + '/api/Asset/delete/' + data[1],
          headers: {
            Authorization: 'Bearer ' + AddAssetsDashboardWebPart.accessToken
          },
          dataType: 'json',
          contentType: 'application/json',
          success: (result) => {
            console.log("success " + result);
            var url = new URL("https://frcidevtest.sharepoint.com/sites/Lincoln/SitePages/Asset-Mngt-Dashboard.aspx");
            Navigation.navigate(url.toString(), true);
          },
          error: (result) => {
            console.log("error " + result);
          }
        });
        // this._deleteAsset(refNo);
      });
    }
    catch (error) {
      console.log(error);
    }
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
