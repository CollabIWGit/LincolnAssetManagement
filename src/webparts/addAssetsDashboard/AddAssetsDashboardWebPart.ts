import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { Navigation } from 'spfx-navigation';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import * as strings from 'AddAssetsDashboardWebPartStrings';
import * as $ from "jquery";
import 'datatables.net';
import 'datatables.net-dt/css/jquery.dataTables.css';

import { navUtils } from '../../utils/navUtils';
let NavUtils = new navUtils();

import { navbar } from '../../utils/navbar';
let Navbar = new navbar();

require('../../../node_modules/bootstrap/dist/css/bootstrap.min.css');
require('../../../node_modules/@fortawesome/fontawesome-free/css/all.min.css');
require('../../styles/dashboardcss.css');
require('../../styles/spcommon.css');
require('../../styles/test.css');

import * as commonConfig from "../../utils/commonConfig.json";

var selectedLocationArr: any = [];
var selectedBuildingArr: any = [];
var selectedOfficeArr: any = [];
var filteredOfficeName: any = [];
var filteredTypeOfAsset: any = [];
var filteredAssetRefNo: any = [];

//#region Interfaces
export interface IAddAssetsDashboardWebPartProps {
  description: string;
}

export interface IApplicationDetailsList {
  Name: string;
  ReferenceNumber: string;
  BuildingName: string;
  OfficeName: string;
  BuildingLocation: string;
  FloorNo: string;
  Ownership: string;
  TypeOfAsset: string;
  ServicingRequired: boolean;
  LastServicingDate: string;
  ServicingPeriod: string;
  Comment: string;
  AssetAttachments: IAttachmentDetails[];
}

export interface IAttachmentDetails {
  AttachmentGUID: string;
  AttachmentFileName: string;
  AttachmentFileContent: any[];
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
  private accessTokenValue: string = "";
  private ListOfAssets: ITypeOfAssetList[];
  private assetList: IDynamicField[];
  private assetByFilterList: IDynamicField[];
  private ListOfBuildings: IBuildings[];
  private ListOfOffices: IOffices[];
  private ListOfOfficeFiltered: IOffices[];
  private ListOfBuildingsFiltered: IBuildings[];

  private LocationsFilterFromLocalStorage = [];
  private BuildingsFilterFromLocalStorage = [];
  private OfficesFilterFromLocalStorage = [];
  private TypeOfAssetFilterFromLocalStorage = "";
  private AssetRefNoFilterFromLocalStorage = "";

  public render(): void {
    SPComponentLoader.loadCss("https://cdn.datatables.net/1.10.19/css/jquery.dataTables.min.css");
    SPComponentLoader.loadScript("https://code.jquery.com/jquery-3.5.1.js");
    SPComponentLoader.loadScript("https://cdn.datatables.net/1.11.4/js/jquery.dataTables.min.js");
    SPComponentLoader.loadScript("https://cdn.datatables.net/fixedheader/3.2.1/js/dataTables.fixedHeader.min.js");
    SPComponentLoader.loadScript("https://cdn.datatables.net/responsive/2.2.3/js/dataTables.responsive.js");
    SPComponentLoader.loadCss("https://cdn.datatables.net/responsive/2.2.3/css/responsive.bootstrap.css");

    this.domElement.innerHTML = `<div id="loader"></div>
    <div id="wrapper" class="">
      <!-- Sidebar -->
      ${Navbar.navbar}
      <!-- /#sidebar-wrapper -->
      <!-- Page Content -->
      <div id="page-content-wrapper">
        <div class="container-fluid">
          <div class="row">
            <div class="col-lg-12">
              <div class="navnav">
                <a href="#menu-toggle" class="btn btn-default" id="menu-toggle"><i class="fas fa-align-justify"></i></a>
              </div>
              <nav class="navbar navbar-expand-lg navbar-dark bg-dark" id="navnavr">
                <div class="container-fluid">
                  <div class="col-lg-12" id="title">
                    <h3>Asset List</h3>
                  </div>
                </div>
              </nav>
              <div id="content2">
                <div class="w3-container" id="form">
                  <div id="content3">
                    <div class="filters">
                      <div class="form-row">
                        <div class="col-md-6" style="display:none;">
                          <div id="locationFilter">
                            <div>
                              <h7>Location</h7>
                            </div>
                            <div class="card" id="card">
                              <div class="card-body" id="card">
                                <form>
                                  <div class="inner-form">
                                    <div class="advance-search">
                                      <div class="form-row" id="locationFilters">
                                      </div>
                                    </div>
                                  </div>
                                </form>
                              </div>
                            </div>
                          </div>
                        </div>
                        <div class="col-md-6">
                          <div id="buildingFilter">
                            <div>
                              <h7>Building</h7>
                            </div>
                            <div class="card" id="card">
                              <div class="card-body" id="card">
                                <form>
                                  <div class="inner-form">
                                    <div class="advance-search">
                                      <div class="form-row" id="buildingFilters">
                                      </div>
                                    </div>
                                  </div>
                                </form>
                              </div>
                            </div>
                          </div>
                        </div>
                        <div class="col-md-6">
                          <div id="officeFilter">
                            <div>
                              <h7>Office</h7>
                            </div>
                            <div class="card" id="card">
                              <div class="card-body" id="card">
                                <form>
                                  <div class="inner-form">
                                    <div class="advance-search">
                                      <div class="form-row" id="officeFilters">
                                      </div>
                                    </div>
                                  </div>
                                </form>
                              </div>
                            </div>
                          </div>
                        </div>
                      </div>
                      <hr class="lineBreak">
                      <div class="form-row">
                        <div class="col-md-6">
                          <div>
                            <h7>Type Of Asset</h7>
                          </div>
                          <div class="input-group">
                            <input list="idTypeOfAsset" id="myListTypeOfAsset" name="myBrowserTypeOfAsset" autocomplete="off" />
                            <datalist id="idTypeOfAsset">
                            </datalist>
                          </div>
                        </div>
                        <div class="col-md-6">
                          <div>
                            <h7>Asset Reference No</h7>
                          </div>
                          <div class="input-group">
                            <input list="idAssetReferenceNo" id="myListAssetReferenceNo" name="myBrowserAssetReferenceNo" autocomplete="off" />
                            <datalist id="idAssetReferenceNo">
                            </datalist>
                          </div>
                        </div>
                      </div>
                      <div class="form-row btnFilterRow">
                        <div class="col-md-2 offset-10">
                          <button type="button" class="btn btn-sm btn-secondary" id="btnFilter">Search</button>
                        </div>
                      </div>
                      <div id="divContainer">
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
    <!-- /#page-content-wrapper -->
    </div>`;

    $("#cover").fadeOut(4000);
    $("#menu-toggle").click((e) => {
      e.preventDefault();
      $("#wrapper").toggleClass("toggled");
    });

    this._getAccessToken();
    this._getTypeOfAssetList();
    this._getBuildingAndLocationList();
    this._getAllOffices();
    this.AddEventListeners();
    this._navigateToAddAssetForm();
    this._getAssetsAsync();
    this._loader();
    NavUtils.collapse();
    NavUtils.navTriggers();
    // NavUtils.cover();

    this._getFiltersFromLocalStorage();
  }

  private AddEventListeners(): any {
    document.getElementById('btnFilter').addEventListener('click', () => this._displayAssets());
    document.getElementById('btnFilter').addEventListener('click', () => this._loader());
    document.getElementById('idTypeOfAsset').addEventListener('click', () => this._getAssetsByFilters(this.accessTokenValue));
  }

  private _navigateToAddAssetForm() {
    $('#btnAdd').on('click', () => {
      Navigation.navigate(`${this.context.pageContext.web.absoluteUrl}/SitePages/${commonConfig.Page.AddAssets}`, true);
    });
  }

  private _loader() {
    let html: string = "";
    html += `<div id="cover"> <span class="fa-solid fa-rotate"></span> loading...</div>`;

    const listContainer: Element = this.domElement.querySelector('#loader');
    listContainer.innerHTML = html;

    NavUtils.cover();
  }

  //#region Filters
  private async _getListOfRefNo() {
    try {
      let html: string = '';

      if (selectedOfficeArr.length > 0) {
        filteredAssetRefNo.forEach((item: string) => {
          html += `
          <option value="${item}">${item}</option>`;
        });
      }
      else {
        this.assetByFilterList.forEach((asset: IDynamicField) => {
          html += `
            <option value="${asset.ReferenceNumber}">${asset.ReferenceNumber}</option>`;
        });
      }

      const listContainer: Element = this.domElement.querySelector('#idAssetReferenceNo');
      listContainer.innerHTML = html;
    }
    catch (error) {
      console.log(error);
      return error;
    }
  }

  private _getTypeOfAssetList() {
    try {
      this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${commonConfig.List.TypeOfAssetList}')/items`, SPHttpClient.configurations.v1)
        .then(response => {
          return response.json()
            .then(async (items: any): Promise<void> => {
              this.ListOfAssets = items.value;

              await this._displayTypeOfAssetList();
            });
        });
    }
    catch (error) {
      console.log(error);
      return error;
    }
  }

  private _displayTypeOfAssetList() {
    try {
      let html: string = '';

      if (selectedOfficeArr.length > 0) {
        filteredTypeOfAsset.forEach((item: string) => {
          html += `
          <option value="${item}">${item}</option>`;
        });
      }
      else {
        this.ListOfAssets.forEach((item: ITypeOfAssetList) => {
          html += `
          <option value="${item.Title}">${item.Title}</option>`;
        });
      }

      const listContainer: Element = this.domElement.querySelector('#idTypeOfAsset');
      listContainer.innerHTML = html;
    }
    catch (error) {
      console.log(error);
      return error;
    }
  }

  private _getAllOffices() {
    try {
      this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${commonConfig.List.OfficeList}')/items`, SPHttpClient.configurations.v1)
        .then(response => {
          return response.json()
            .then((items: any): void => {
              this.ListOfOffices = items.value;

              this.ListOfOfficeFiltered = this.ListOfOffices.filter((obj, pos, arr) => {
                return arr.map(mapObj =>
                  mapObj.Title).indexOf(obj.Title) == pos;
              });

              this._officeFilters();
            });
        });
    }
    catch (error) {
      console.log(error);
      return error;
    }
  }

  private _getBuildingAndLocationList() {
    try {
      let html: string = '';
      this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${commonConfig.List.BuildingList}')/items`, SPHttpClient.configurations.v1)
        .then(response => {
          return response.json()
            .then((items: any): void => {
              this.ListOfBuildings = items.value;

              this.ListOfBuildingsFiltered = this.ListOfBuildings.filter((obj, pos, arr) => {
                return arr.map(mapObj =>
                  mapObj.Location).indexOf(obj.Location) == pos;
              });

              // this._locationFilters();
              this._buildingFilters();
            });
        });
    }
    catch (error) {
      console.log(error);
      return error;
    }
  }

  private async _getAssetsByFilters(token: string) {
    try {
      var assetRefNoValue = (<HTMLInputElement>document.getElementById('myListAssetReferenceNo')).value;
      var typeOfAssetValue = (<HTMLInputElement>document.getElementById('myListTypeOfAsset')).value;
      var assetNameValue = "";
      var locationValue = "";
      var buildingValue = "";
      var officeValue = "";

      selectedLocationArr = [];
      selectedBuildingArr = [];
      selectedOfficeArr = [];

      $('#locationFilters input:checked').each(function () {
        selectedLocationArr.push($(this).attr('name'));
      });

      $('#buildingFilters input:checked').each(function () {
        selectedBuildingArr.push($(this).attr('name'));
      });
  
      $('#officeFilters input:checked').each(function () {
        selectedOfficeArr.push($(this).attr('name'));
      });

      console.log(assetRefNoValue);
      console.log(typeOfAssetValue);
      console.log(selectedLocationArr);
      console.log(selectedBuildingArr);
      console.log(selectedOfficeArr);

      if (selectedLocationArr.length > 0) {
        selectedLocationArr.forEach((location: string) => {
          locationValue += location + ";";
        });

        locationValue = locationValue.slice(0, -1);
      }

      if (selectedBuildingArr.length > 0) {
        selectedBuildingArr.forEach((building: string) => {
          buildingValue += building + ";";
        });

        buildingValue = buildingValue.slice(0, -1);
      }

      if (selectedOfficeArr.length > 0) {
        selectedOfficeArr.forEach((office: string) => {
          officeValue += office + ";";
        });

        officeValue = officeValue.slice(0, -1);
      }

      await $.ajax({
        type: 'GET',
        url: commonConfig.baseUrl + `/api/Asset/GetAssetsByFilters?refNo=${assetRefNoValue}&assetName=${assetNameValue}&typeOfAsset=${typeOfAssetValue}&location=${locationValue}&office=${officeValue}&buildingName=${buildingValue}`,
        headers: {
          Authorization: 'Bearer ' + token
        },
        success: (result) => {
          this.assetByFilterList = result;
          this._getListOfRefNo();
        },
        error: (result) => {
          console.log(result);
          return result;
        }
      });
    }
    catch (error) {
      console.log(error);
      return error;
    }
  }

  private async _displayAssets() {
    this._saveAssetFiltersInLocalStorage();

    await this._getAssetsByFilters(this.accessTokenValue);
    if (this.assetByFilterList.length > 0) {
      var table = $('#tbl_asset_list').DataTable();

      if (table.data().any()) {
        table.clear().draw();
        table.destroy();
      }
      this._renderTable(this.assetByFilterList);
      this._renderTableAsync();
    }
    else {
      this._displayNoDataAvailable();
    }
  }
  //#endregion

  private _locationFilters() {
    try {
      let html: string = "";

      this.ListOfBuildingsFiltered.forEach((item: IBuildings) => {
        html += `
        <div class="input-field">
          <div class="custom-control custom-checkbox">
            <input type="checkbox" class="custom-control-input location" id="${item.Location}" name="${item.Location}" value="${item.Location}" ${this.LocationsFilterFromLocalStorage.includes(item.Location) ? "checked" : ""}>
            <label for="${item.Location}" class="custom-control-label"> ${item.Location}</label><br>
          </div>
        </div>`;
      });

      const listContainer: Element = this.domElement.querySelector('#locationFilters');
      listContainer.innerHTML = html;

      $('.location').change(async () => {
        var elementId: string = $(event.currentTarget).attr("id");
        var element = <HTMLInputElement>document.getElementById(`${elementId}`);
        if (element.checked) {
          selectedLocationArr.push(elementId);
        }
        else {
          selectedLocationArr.forEach(async (item, index) => {
            if (item == elementId) {
              selectedLocationArr.splice(index, 1);
            }
          });
        }

        if (selectedLocationArr.length > 0) {
          await this._filterOfficesListOnLocationChange();
        }
      });
    }
    catch (error) {
      console.log(error);
      return error;
    }
  }

  private _buildingFilters() {
    try {
      let html: string = "";

      this.ListOfBuildings.forEach((item: IBuildings) => {
        html += `
        <div class="input-field">
          <div class="custom-control custom-checkbox">
            <input type="checkbox" class="custom-control-input building" id="${item.Title}" name="${item.Title}" value="${item.Title}" ${this.BuildingsFilterFromLocalStorage.includes(item.Title) ? "checked" : ""}>
            <label for="${item.Title}" class="custom-control-label"> ${item.Title}</label><br>
          </div>
        </div>`;
      });

      const listContainer: Element = this.domElement.querySelector('#buildingFilters');
      listContainer.innerHTML = html;

      $('.building').change(async () => {
        var elementId: string = $(event.currentTarget).attr("id");
        var element = <HTMLInputElement>document.getElementById(`${elementId}`);
        if (element.checked) {
          selectedBuildingArr.push(elementId);
        }
        else {
          selectedBuildingArr.forEach(async (item, index) => {
            if (item == elementId) {
              selectedBuildingArr.splice(index, 1);
            }
          });
        }

        // if (selectedBuildingArr.length > 0) {
        await this._filterOfficesListOnBuildingChange();
        // }
      });
    }
    catch (error) {
      console.log(error);
      return error;
    }
  }

  private _officeFilters() {
    try {
      let html: string = "";
      selectedOfficeArr = [];

      if (selectedLocationArr.length > 0 || selectedBuildingArr.length > 0) {
        filteredOfficeName.forEach((officeName: string) => {
          html += `
          <div class="input-field">
            <div class="custom-control custom-checkbox">
              <input type="checkbox" class="custom-control-input office" id="${officeName}" name="${officeName}" value="${officeName}" ${this.OfficesFilterFromLocalStorage.includes(officeName) ? "checked" : ""}>
              <label for="${officeName}" class="custom-control-label"> ${officeName}</label><br>
            </div>
          </div>`;
        });
      }
      else {
        this.ListOfOfficeFiltered.forEach((item: IOffices) => {
          html += `
          <div class="input-field">
            <div class="custom-control custom-checkbox">
              <input type="checkbox" class="custom-control-input office" id="${item.Title}" name="${item.Title}" value="${item.Title}" ${this.OfficesFilterFromLocalStorage.includes(item.Title) ? "checked" : ""}>
              <label for="${item.Title}" class="custom-control-label"> ${item.Title}</label><br>
            </div>
          </div>`;
        });
      }

      const listContainer: Element = this.domElement.querySelector('#officeFilters');
      listContainer.innerHTML = html;

      $('.office').change(async () => {
        var elementId: string = $(event.currentTarget).attr("id");
        var element = <HTMLInputElement>document.getElementById(`${elementId}`);
        if (element.checked) {
          selectedOfficeArr.push(elementId);
        }
        else {
          selectedOfficeArr.forEach(async (item, index) => {
            if (item == elementId) {
              selectedOfficeArr.splice(index, 1);
            }
          });
        }

        if (selectedOfficeArr.length > 0) {
          await this._filterAssetsOnOfficeChange();
        }
        else {
          await this._displayTypeOfAssetList();
        }
      });
    }
    catch (error) {
      console.log(error);
      return error;
    }
  }

  private _filterOfficesListOnLocationChange() {
    filteredOfficeName = [];
    filteredTypeOfAsset = [];
    filteredAssetRefNo = [];

    selectedLocationArr.forEach((location: string) => {
      this.ListOfBuildings.forEach((building: IBuildings) => {
        if (building.Location == location) {
          this.ListOfOffices.forEach((office: IOffices) => {
            if (building.ID == office.BuildingIDId) {
              filteredOfficeName.push(office.Title);
            }
          });
        }
      });
    });

    filteredOfficeName = filteredOfficeName.filter((element, index, self) => {
      return index === self.indexOf(element);
    });

    this._officeFilters();
  }

  private _filterOfficesListOnBuildingChange() {
    filteredOfficeName = [];
    filteredTypeOfAsset = [];
    filteredAssetRefNo = [];

    if(selectedBuildingArr.length > 0) {
      selectedBuildingArr.forEach((buildingName: string) => {
        this.ListOfBuildings.forEach((building: IBuildings) => {
          if (building.Title == buildingName) {
            this.ListOfOffices.forEach((office: IOffices) => {
              if (building.ID == office.BuildingIDId) {
                filteredOfficeName.push(office.Title);
              }
            });
          }
        });
      });
    }
    else {
      this.ListOfOffices.forEach((office: IOffices) => {
        filteredOfficeName.push(office.Title);
      });
    }

    //Check if buildings checked in local storage
    if (this.BuildingsFilterFromLocalStorage.length > 0) {
      this._filterOfficeWithBuildingsLocalStorage();
    }

    filteredOfficeName = filteredOfficeName.filter((element, index, self) => {
      return index === self.indexOf(element);
    });

    this._officeFilters();
  }

  private _filterOfficeWithBuildingsLocalStorage() {
    try {
      this.BuildingsFilterFromLocalStorage.forEach((buildingName: string) => {
        this.ListOfBuildings.forEach((building: IBuildings) => {
          if (building.Title == buildingName) {
            this.ListOfOffices.forEach((office: IOffices) => {
              if (building.ID == office.BuildingIDId) {
                filteredOfficeName.push(office.Title);
              }
            });
          }
        });
      });
    }
    catch (error) {
      console.log(error);
      return error;
    }
  }

  private _filterAssetsOnOfficeChange() {
    filteredTypeOfAsset = [];
    filteredAssetRefNo = [];

    //If building options have been selected
    if (selectedBuildingArr.length > 0) {
      selectedBuildingArr.forEach((building: string) => {
        this.assetList.forEach((asset: IDynamicField) => {
          if (asset.BuildingName == building) {
            selectedOfficeArr.forEach((office: string) => {
              if (asset.OfficeName == office) {
                filteredTypeOfAsset.push(asset.TypeOfAsset);
                filteredAssetRefNo.push(asset.ReferenceNumber);
              }
            });
          }
        });
      });
    }

    //If location options have been selected
    // if (selectedLocationArr.length > 0) {
    //   selectedLocationArr.forEach((location: string) => {
    //     this.assetList.forEach((asset: IDynamicField) => {
    //       if (asset.BuildingLocation == location) {
    //         selectedOfficeArr.forEach((office: string) => {
    //           if (asset.OfficeName == office) {
    //             filteredTypeOfAsset.push(asset.TypeOfAsset);
    //             filteredAssetRefNo.push(asset.ReferenceNumber);
    //           }
    //         });
    //       }
    //     });
    //   });
    // }
    else {
      this.assetList.forEach((asset: IDynamicField) => {
        selectedOfficeArr.forEach((office: string) => {
          if (asset.OfficeName == office) {
            filteredTypeOfAsset.push(asset.TypeOfAsset);
            filteredAssetRefNo.push(asset.ReferenceNumber);
          }
        });
      });
    }

    filteredTypeOfAsset = filteredTypeOfAsset.filter((element, index, self) => {
      return index === self.indexOf(element);
    });

    this._displayTypeOfAssetList();
    this._getListOfRefNo();
  }

  private _getAccessTokenForDisplay() {
    try {
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
    catch (error) {
      console.log(error);
      return error;
    }
  }

  private _getAccessToken(): void {
    try {
      var body = {
        grant_type: 'password',
        client_id: 'myClientId',
        client_secret: 'myClientSecret',
        username: "roukaiyan@frci.net",
        password: "Pa$$w0rd"
      };

      $.ajax({
        type: 'POST',
        url: commonConfig.baseUrl + '/token',
        dataType: 'json',
        data: body,
        contentType: 'application/x-www-form-urlencoded',
        success: (result) => {
          this.accessTokenValue = result["access_token"];
          this._getAllAssets(result["access_token"]);
          return this.accessTokenValue;
        },
        error: (result) => {
          return result;
        }
      });
    }
    catch (error) {
      console.log(error);
      return error;
    }
  }

  private _getAllAssets(token: string): void {
    try {
      $.ajax({
        type: 'GET',
        url: commonConfig.baseUrl + '/api/Asset/GetAssets',
        headers: {
          Authorization: 'Bearer ' + token
        },
        success: (result) => {
          this.assetList = result;
        },
        error: (result) => {
          return result;
        }
      });
    }
    catch (error) {
      console.log(error);
      return error;
    }
  }

  private async _getAssetsAsync() {
    try {
      let token = await this._getAccessTokenForDisplay();
      this._renderTableAsync();
    }
    catch (error) {
      console.log(error);
      return error;
    }
  }

  private _renderTable(listOfAssets: IDynamicField[]) {
    try {
      let html: string = `<table id="tbl_asset_list" class="table table-striped">
        <thead>
          <tr>
            <th class="text-left">Type of Assets</th>
            <th class="text-left">Asset Reference No</th>
            <th class="text-left">Asset Name</th>
            <th class="text-center">View</th>
            <th class="text-center">Delete</th>
          </tr>
        </thead>
        <tbody id="tb_asset_list">`;

      listOfAssets.forEach((item: IDynamicField) => {
        html += `
          <tr>
            <td class="text-left">${item.TypeOfAsset}</td>
            <td class="text-left">${item.ReferenceNumber}</td>
            <td class="text-left">${item.Name}</td>
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
    catch (error) {
      console.log(error);
      return error;
    }
  }

  private _displayNoDataAvailable() {
    try {
      let html: string = "";

      html += '<div id="noDataText">There is no data available.</div>';

      const listContainer: Element = this.domElement.querySelector('#divContainer');
      listContainer.innerHTML = html;
    }
    catch (error) {
      console.log(error);
      return error;
    }
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
          { orderable: false, targets: [3, 4] }
        ],
        order: [[0, "asc"]]
      });

      $('#TypeOfAssets').on('keyup', 'input', function () {
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

      $('#AssetName').on('keyup', 'input', function () {
        table
          .columns(2)
          .search(this.value)
          .draw();
      });

      //Click view btn
      $('#tbl_asset_list').on('click', '.view', function () {
        var data = table.row($(this).parents('tr')).data();
        var refNo = data[1];
        var url = new URL(`${commonConfig.url}/SitePages/${commonConfig.Page.AddAssets}`);
        url.searchParams.append('refNo', refNo);
        Navigation.navigate(url.toString(), true);
      });

      //Click delete btn
      $('#tbl_asset_list').on('click', '.delete', function () {
        if (confirm("Are you sure you want to delete this asset?")) {
          var data = table.row($(this).parents('tr')).data();
          $.ajax({
            type: 'DELETE',
            data: { 'action': 'delete' },
            url: commonConfig.baseUrl + '/api/Asset/delete/' + data[1],
            headers: {
              Authorization: 'Bearer ' + AddAssetsDashboardWebPart.accessToken
            },
            dataType: 'json',
            contentType: 'application/json',
            success: (result) => {
              var url = new URL(`${commonConfig.url}/SitePages/${commonConfig.Page.AssetList}`);
              Navigation.navigate(url.toString(), true);
              return result;
            },
            error: (result) => {
              return result;
            }
          });
        }
        else {

        }
      });
    }
    catch (error) {
      console.log(error);
      return error;
    }
  }

  private _saveAssetFiltersInLocalStorage() {
    var selectedLocationsFilters = [];
    var selectedBuildingsFilters = [];
    var selectedOfficesFilters = [];

    $('#locationFilters input:checked').each(function () {
      selectedLocationsFilters.push($(this).attr('name'));
    });

    $('#buildingFilters input:checked').each(function () {
      selectedBuildingsFilters.push($(this).attr('name'));
    });

    $('#officeFilters input:checked').each(function () {
      selectedOfficesFilters.push($(this).attr('name'));
    });

    var assetRefNoValue = (<HTMLInputElement>document.getElementById('myListAssetReferenceNo')).value;
    var typeOfAssetValue = (<HTMLInputElement>document.getElementById('myListTypeOfAsset')).value;
    var filters = '';

    filters += "Locations=";
    if (selectedLocationsFilters.length > 0) {
      filters += selectedLocationsFilters.join(',');
    }
    filters += "&";

    filters += "Buildings=";
    if (selectedBuildingsFilters.length > 0) {
      filters += selectedBuildingsFilters.join(',');
    }
    filters += "&";

    filters += "Offices=";
    if (selectedOfficesFilters.length > 0) {
      filters += selectedOfficesFilters.join(',');
    }
    filters += "&";

    filters += `RefNo=${assetRefNoValue}&`;

    filters += `TypeOfAsset=${typeOfAssetValue}`;

    localStorage.setItem('filter', filters);
  }

  private _getFiltersFromLocalStorage() {
    var allFilters = localStorage.getItem('filter');
    if (allFilters != null) {

      var splitedFilters = allFilters.split('&');
      var locationFilter = splitedFilters[0].split('=')[1].split(','); //array of locations
      var buildingFilter = splitedFilters[1].split('=')[1].split(','); //array of buildings
      var officeFilter = splitedFilters[2].split('=')[1].split(','); //array of offices
      var refNoFilter = splitedFilters[3].split('=')[1];
      var typeOfAssetFilter = splitedFilters[4].split('=')[1];

      this.LocationsFilterFromLocalStorage = locationFilter;
      this.BuildingsFilterFromLocalStorage = buildingFilter;
      this.OfficesFilterFromLocalStorage = officeFilter;
      this.AssetRefNoFilterFromLocalStorage = refNoFilter;
      this.TypeOfAssetFilterFromLocalStorage = typeOfAssetFilter;

      if (this.AssetRefNoFilterFromLocalStorage != '')
        $('#myListAssetReferenceNo').val(this.AssetRefNoFilterFromLocalStorage);

      if (this.TypeOfAssetFilterFromLocalStorage != '')
        $('#myListTypeOfAsset').val(this.TypeOfAssetFilterFromLocalStorage);
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
