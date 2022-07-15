//#region Imports
import { //Guid, 
  Version
} from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { Navigation } from 'spfx-navigation';
import * as strings from 'AddAssetsWebPartStrings';
import * as $ from "jquery";
import { StringIterator } from 'lodash';
import { Guid } from "guid-typescript";
import { sp } from '@pnp/sp/presets/all';
import * as moment from 'moment';
import * as mime from 'mime-types';

import { navUtils } from '../../utils/navUtils';
let NavUtils = new navUtils();

import { navbar } from '../../utils/navbar';
let Navbar = new navbar();

require('../../../node_modules/bootstrap/dist/css/bootstrap.min.css');
require('../../../node_modules/bootstrap/js/src/collapse.js');
require('../../../node_modules/@fortawesome/fontawesome-free/css/all.min.css');
require('../../../node_modules/@fortawesome/fontawesome-free/css/fontawesome.css');
require('../../styles/global.css');
require('../../styles/spcommon.css');
require('../../styles/test.css');

import * as commonConfig from "../../utils/commonConfig.json";
//#endregion

var fileInfos = [];
var populatedFileInfos = [];
var fileDownloadInfos = [];
var tempFileInfos = [];
var filestream;
var fixarray;
var fileByteArray = [];
var folderID;

//#region Interfaces
export interface IAddAssetsWebPartProps {
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
  FolderId: string;
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

export interface IOffices {
  Title: string;
  FloorNumber: number;
  BuildingIDId: number;
  ID: number;
  ShortForm: string;
}

export interface IBuildings {
  ID: number;
  Title: string;
  Location: string;
  ShortForm: string;
}
//#endregion

export default class AddAssetsWebPart extends BaseClientSideWebPart<IAddAssetsWebPartProps> {
  private dropdownListName: string = "";
  private arrFieldsRequired = [];
  private accessToken: string = "";
  private dynamicField: IDynamicField;
  private formDetails: IDynamicField;
  private floorNoFiltered: any = [];
  private OfficeNameFiltered: any = [];
  private ListOfOffices: IOffices[];
  private ListOfBuildings: IBuildings[];
  public fileGUID: Guid;
  public folderItemGUID: Guid;
  private mainFileByteArray = [];

  protected onInit(): Promise<void> {

    return super.onInit().then(_ => {
      sp.setup({
        spfxContext: this.context
      });
    });
  }

  public render(): void {
    this.domElement.innerHTML = `${Navbar.cover}
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
                    <h3 id="header_title">Asset Form</h3>
                  </div>
                </div>
              </nav>
              <div id="content2">
                <div class="w3-container" id="form">
                  <div id="content3">
                    <div class="form-row">
                      <div class="col-md-6">
                        <div>
                          <h7>Asset Reference No.</h7>
                        </div>
                        <div class="input-group">
                          <input type="text" id="idAssetRefNo" readonly autocomplete="off"/> 
                        </div>
                      </div>
                      <div class="col-md-6">
                        <div>
                          <h7>Asset Description</h7>
                        </div>
                        <div class="input-group">
                          <input type="text" id="idAssetName" autocomplete="off"/>
                        </div>
                      </div>
                    </div>
                    <div class="form-row">
                      <div class="col-md-6">
                        <div>
                          <h7>Building Name <span class="required" id="requiredBuildingName">*</span></h7>
                        </div>
                        <div class="input-group">
                          <select name="BuildingName" id="idBuildingName">
                          </select>
                        </div>
                      </div>
                      <div class="col-md-6">
                        <div>
                          <h7>Building Location <span class="required" id="requiredBuildingLocation">*</span></h7>
                        </div>
                        <div class="input-group">
                          <input type="text" id="idBuildingLocation" autocomplete="off"/> 
                        </div>
                      </div>
                    </div>
                    <div class="form-row">
                      <div class="col-md-6">
                        <div>
                          <h7>Office <span class="required" id="requiredOffice">*</span></h7>
                        </div>
                        <div class="input-group">
                          <select name="office" id="idOffice">
                          </select>
                        </div>
                      </div>
                      <div class="col-md-6">
                        <div>
                          <h7>Floor <span class="required" id="requiredFloor">*</span></h7>
                        </div>
                          <select name="Floor" id="idFloor">
                          </select>
                      </div>
                    </div>
                    <div class="form-row">
                      <div class="col-md-12 input-group">
                        <h7>Ownership</h7>
                        <input type="text" id="idOwnership" autocomplete="off"/>
                      </div>
                    </div>
                    <div class="form-row">
                      <div class="col-md-12">
                        <h7>Type Of Asset <span class="required" id="requiredTypeOfAsset">*</span></h7>
                        <div class="input-group">
                          <select id="typeOfAssetList">
                          </select>
                        </div>
                      </div>
                    </div>
                    <div id="dynamicFields" class="form-row">
                    </div>
                    <div class="form-row">
                      <div class="col-md-12">
                        <h7>Servicing / Test Required</h7>
                        <div class="input-group">
                          <input id="servicingRequired" type="radio" name="servicingReq" value="true" checked="true"/>
                          <label for="servicingRequired"><span>Yes</span></label>
                          <input id="servicingNotRequired" type="radio" name="servicingReq" value="false"/>
                          <label for="servicingNotRequired"><span>No</span></label>
                        </div>
                      </div>
                    </div>
                    <div class="form-row">
                      <div class="col-md-6">
                        <div>
                          <h7>Last Servicing / Test Date</h7>
                        </div>
                        <div class="input-group">
                          <input type="date" id="idLastServicingDate" autocomplete="off"/>
                        </div>
                      </div>
                      <div class="col-md-6">
                        <div>
                          <h7>Servicing / Test Period</h7>
                        </div>
                        <div class="form-row input-group">
                          <input type="number" id="idServicingPeriod" autocomplete="off" placeholder="Enter servicing/test period in months" min="0" class="form-control"/>
                          <div class="input-group-append">
                            <span class="input-group-text" id="basic-addon1">Month(s)</span> 
                          </div>
                        </div>
                      </div>
                    </div>
                    <div class="form-row">
                      <div class="popup-overlay col-md-12">
                        <div class="popup-content" id="popup">
                        </div>
                      </div>
                    </div>
                    <div class="form-row">
                      <div class="col-md-12 input-group">
                        <h7>Comments</h7>
                        <textarea rows="3" id="idComments" autocomplete="off"></textarea>
                      </div>
                    </div>
                    <div class="form-row">
                      <div class="col-md-12">
                        <h7>Attachments</h7>
                        <div class="custom-file">
                          <input type="file" id="customFile" name="files" multiple>
                        </div>
                      </div>
                    </div>
                    <div class="form-row">
                      <div class="col-md-9 table-responsive">
                        <table class="table" id="attachmentTable">
                          <thead>
                            <tr>
                              <th class="th-lg" scope="col">Attachment Name</th>
                              <th scope="col">Action</th>
                            </tr>
                          </thead>
                          <tbody id="tableAttachmentContainer">
                          </tbody>
                        </table>
                      </div>
                    </div>
                    
                    <div class="form-row">
                      <div class="col-md-4 offset-8">
                        <button id="btnSubmit" class="btn btn-secondary" type="button">Submit</button>
                        <button id="btnCancel" class="btn btn-secondary" type="button">Cancel</button>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
      <!-- /#page-content-wrapper -->
    </div>
  </div>`;

    this._renderTypeOfAssetListAsync();
    this._renderFieldRequiredListAsync();
    this._getOfficesList();
    this._getBuildingsList();
    this._checkAttachmentTable();
    this.AddEventListeners();
    this._getAccessToken();
    NavUtils.navTriggers();
    NavUtils.cover();
  }

  private AddEventListeners(): any {
    document.getElementById('btnSubmit').addEventListener('click', () => this._submit());
    document.getElementById('btnCancel').addEventListener('click', () => this._cancel());
    document.getElementById('customFile').addEventListener('change', () => this.blob());
    document.getElementById('servicingRequired').addEventListener('change', () => this._checkIfServicingRequiredChecked());
    document.getElementById('servicingNotRequired').addEventListener('change', () => this._checkIfServicingNotRequiredChecked());
    document.getElementById('typeOfAssetList').addEventListener('change', () => this._renderFieldRequiredList(this.arrFieldsRequired));
    document.getElementById('idBuildingName').addEventListener('change', () => this._populateLocation());
    document.getElementById('idBuildingName').addEventListener('change', () => this._populateOfficesList(this.ListOfOffices));
    document.getElementById('idOffice').addEventListener('change', () => this._getFloorNo());
    document.getElementById('idFloor').addEventListener('change', () => this._getLastSequenceAssetRefNo(this.accessToken));
  }

  //#region Type of Asset List
  private _getTypeOfAssetListData(): Promise<ITypeOfAssetLists> {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists/GetByTitle('${commonConfig.List.TypeOfAssetList}')/Items`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  private _renderTypeOfAssetList(items: ITypeOfAssetList[]): void {
    var arr = [];
    let html: string = `<option value="">Select Type Of Assets</option>`;
    items.forEach((item: ITypeOfAssetList) => {
      arr.push(item.Title);
      arr.sort();
    });

    for (let j = 0; j < arr.length; j++) {
      html += `<option value="${arr[j]}">${arr[j]}</option>`;
    }

    const listContainer: Element = this.domElement.querySelector('#typeOfAssetList');
    listContainer.innerHTML = html;
  }

  private _renderTypeOfAssetListAsync(): void {
    this._getTypeOfAssetListData().then((response) => {
      this._renderTypeOfAssetList(response.value);
    });
  }
  //#endregion

  //#region Dropdown List
  private _getDropdownListData(listName: string): Promise<IDropdownLists> {
    try {
      return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists/GetByTitle('${listName}')/Items`, SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) => {
          return response.json();
        });
    }
    catch (error) {
      return error;
    }
  }

  private _renderDropdownList(items: IDropdownList[]): void {
    let html: string;
    items.forEach((item: IDropdownList) => {
      html += `<option>${item.Title}</option>`;
    });

    const listContainer: Element = this.domElement.querySelector(`#id${this.dropdownListName}List`);
    listContainer.innerHTML = html;
  }

  private _renderDropdownListAsync(listName: string): void {
    try {
      this._getDropdownListData(listName).then((response) => {
        this._renderDropdownList(response.value);
      });
    }
    catch (error) {
      return error;
    }
  }
  //#endregion

  //#region Fields For Each Type of Assets
  private _getFieldRequiredListData(): Promise<IFieldsRequiredLists> {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists/GetByTitle('${commonConfig.List.FieldsRequiredList}')/Items?$expand=TypeOfAssets&$select=Title,TypeOfAssets/Title,FieldType,DropdownListName,Required`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  private _renderFieldRequiredList(items: IFieldsRequiredList[]): void {
    let html: string = "";
    var typeOfAssetsValue = (<HTMLInputElement>document.getElementById('typeOfAssetList')).value;

    if (items.length > 0) {
      items.forEach((item: IFieldsRequiredList) => {
        var titleItem = item.Title.replace(/ /g, "");
        if (typeOfAssetsValue == item.TypeOfAssets.Title) {
          //Item header
          html += `
          <div class="col-md-6">
            <div>
              <h7>${item.Title}</h7>
            </div>
            <div class="input-group">`;

          //If field is a textbox & check if it is required
          if (item.FieldType == "TextBox") {
            if (item.Required) {
              html += `
                <input type="text" id="id${titleItem}" required autocomplete="off"/>`;
            }
            else {
              html += `
                <input type="text" id="id${titleItem}" autocomplete="off"/>`;
            }
          }

          //If field is a dropdown & check if it is required
          else if (item.FieldType == "Dropdown") {
            if (item.Required) {
              html += `
                <input type="text" id="id${titleItem}" list="id${titleItem}List" name="my${titleItem}Browser" required autocomplete="off"/>
                <datalist id="id${titleItem}List">
                </datalist>`;
            }
            else {
              html += `
                <input type="text" id="id${titleItem}" list="id${titleItem}List" name="my${titleItem}Browser" autocomplete="off"/>
                <datalist id="id${titleItem}List">
                </datalist>`;
            }

            this.dropdownListName = item.DropdownListName;
            this._renderDropdownListAsync(this.dropdownListName);
          }

          //If field is a date & check if it is required
          else if (item.FieldType == "Date") {
            if (item.Required) {
              html += `
                <input type="date" id="id${titleItem}" autocomplete="off"/>`;
            }
            else {
              html += `
                <input type="date" id="id${titleItem}" autocomplete="off"/>`;
            }
          }
        }

        //Closing div tags
        html += `
            </div>
          </div>`;
      });
    }

    const listContainer: Element = this.domElement.querySelector('#dynamicFields');
    listContainer.innerHTML = html;
  }

  private _fromSPListToArr(response: IFieldsRequiredList[]): void {
    this.arrFieldsRequired = response;
  }

  private _renderFieldRequiredListAsync(): void {
    this._getFieldRequiredListData().then((response) => {
      this._fromSPListToArr(response.value);
    });
  }
  //#endregion

  private async _submit() {
    try {
      let html: string = "";

      var result1: string = await this._addServicingDetailsToList();
      var result2: boolean = await this._applicationDetails();

      if (result2) {
        this._saveAsset(this.accessToken);
      }
      else {
        html += `
        <h2>Error</h2>
        <p>An error has occurred.</p>
        <button class="closePopup">Close</button>`;
      }

      const listContainer: Element = this.domElement.querySelector('#popup');
      listContainer.innerHTML = html;

      $(".popup-overlay, .popup-content").addClass("active");
      $(".closePopup").on("click", () => {
        $(".popup-overlay, .popup-content").removeClass("active");
      });
    }
    catch (error) {
      console.log(error);
      return error;
    }
  }

  private _cancel(): void {
    var url = (`${commonConfig.url}/SitePages/${commonConfig.Page.AssetList}`);
    Navigation.navigate(url.toString(), true);
  }

  private _getAccessToken(): void {
    var body = {
      grant_type: 'password',
      client_id: 'myClientId',
      client_secret: 'myClientSecret',
      username: "admin2@lincolnrealty.mu",
      password: "Pa$$w0rd1"
    };

    $.ajax({
      type: 'POST',
      url: commonConfig.baseUrl + '/token',
      dataType: 'json',
      data: body,
      contentType: 'application/x-www-form-urlencoded',
      success: (result) => {
        this.accessToken = result["access_token"];
        this._getAssetById(result["access_token"]);
        if (this._checkURLParameter()) {
          this._getAssetFolderID(result["access_token"]);
        }
      },
      error: (result) => {
        return result;
      }
    });
  }

  //#region GETs and populate dropdowns
  private _getBuildingsList() {
    let html: string = '';
    return new Promise(async (resolve, reject) => {
      try {
        this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${commonConfig.List.BuildingList}')/items`, SPHttpClient.configurations.v1)
          .then(response => {
            return response.json()
              .then((items: any): void => {
                this.ListOfBuildings = items.value;

                html += '<option value="">Choose Building</option>';

                this.ListOfBuildings.forEach((item: IBuildings) => {
                  html += `
              <option value="${item.Title}">${item.Title}</option>`;
                });

                const listContainer: Element = this.domElement.querySelector('#idBuildingName');
                listContainer.innerHTML = html;
              });
          });
        resolve(this.ListOfBuildings);

      }
      catch (error) {
        console.log(error);
        reject(error);
      }
    });

  }

  private _populateLocation() {
    var idBuildingValue = (<HTMLInputElement>document.getElementById('idBuildingName')).value;

    $('#idBuildingLocation').val("");

    this.ListOfBuildings.forEach((item: IBuildings) => {
      if (idBuildingValue == item.Title) {
        $('#idBuildingLocation').val(item.Location);
      }
    });
  }

  private _getOfficesList() {
    this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${commonConfig.List.OfficeList}')/items`, SPHttpClient.configurations.v1)
      .then(response => {
        return response.json()
          .then((items: any): void => {
            this.ListOfOffices = items.value;
          });
      });
  }

  private _populateOfficesList(officesList: IOffices[]) {
    let html: string = '';
    var idBuildingValue = (<HTMLInputElement>document.getElementById('idBuildingName')).value;

    $('#idOffice').val("");
    $('#idFloor').val("");
    this.OfficeNameFiltered = [];

    html += '<option value="">Choose Office</option>';

    this.ListOfBuildings.forEach((itemBuilding: IBuildings) => {
      if (idBuildingValue == itemBuilding.Title) {
        officesList.forEach((itemOffice: IOffices) => {
          if (itemBuilding.ID == itemOffice.BuildingIDId) {
            if (jQuery.inArray(itemOffice.Title, this.OfficeNameFiltered) == -1) {
              this.OfficeNameFiltered.push(itemOffice.Title);
            }
          }
        });
      }
    });

    this.OfficeNameFiltered.forEach((office: string) => {
      html += `<option value="${office}">${office}</option>`;
    });

    const listContainer: Element = this.domElement.querySelector('#idOffice');
    listContainer.innerHTML = html;
  }

  private _getFloorNo(): void {
    let html: string = '';
    var idOfficeValue = (<HTMLInputElement>document.getElementById('idOffice')).value;

    $('#idFloor').val("");
    this.floorNoFiltered = [];

    html += '<option value="">Choose Floor</option>';

    this.ListOfOffices.forEach((item: IOffices) => {
      if (idOfficeValue == item.Title) {
        if (jQuery.inArray(item.FloorNumber, this.floorNoFiltered) == -1) {
          this.floorNoFiltered.push(item.FloorNumber);
        }
      }
    });

    this.floorNoFiltered.forEach((floor: number) => {
      html += `
        <option value="${floor}">${floor}</option>`;
    });

    const listContainer: Element = this.domElement.querySelector('#idFloor');
    listContainer.innerHTML = html;
  }

  private _getLastSequenceAssetRefNo(token: string) {
    var idBuildingValue = (<HTMLInputElement>document.getElementById('idBuildingName')).value;
    var idOfficeValue = (<HTMLInputElement>document.getElementById('idOffice')).value;
    var idFloorValue = (<HTMLInputElement>document.getElementById('idFloor')).value;

    $.ajax({
      type: 'GET',
      url: commonConfig.baseUrl + '/api/Asset/GetLastSequence?buildingName=' + idBuildingValue + '&floorNo=' + idFloorValue + '&officeName=' + idOfficeValue,
      headers: {
        Authorization: 'Bearer ' + token
      },
      success: (result) => {
        this._populateAssetRefNo(result);
      },
      error: (error) => {
        return error;
      }
    });
  }

  private _populateAssetRefNo(sequenceNum: number) {
    var buildingNameValue = (<HTMLInputElement>document.getElementById('idBuildingName')).value;
    var floorNoValue = (<HTMLInputElement>document.getElementById('idFloor')).value;
    var idOfficeValue = (<HTMLInputElement>document.getElementById('idOffice')).value;
    var substrOffice = (idOfficeValue.replace(/ /g, "")).substring(0, 3);
    var strFloorNum = (floorNoValue.replace(/ /g, "")).replace(",", "");

    var nextSequenceNumber: number = +sequenceNum;
    var strNextSequenceNumber: string = nextSequenceNumber.toString();
    while (strNextSequenceNumber.length < 3) {
      strNextSequenceNumber = "0" + strNextSequenceNumber;
    }
    this.ListOfBuildings.forEach((itemBuilding: IBuildings) => {
      if (buildingNameValue == itemBuilding.Title) {
        $('#idAssetRefNo').val(itemBuilding.ShortForm + "_" + strFloorNum + "_" + substrOffice + "_" + strNextSequenceNumber);
      }
    });
  }
  //#endregion

  private _saveAsset(token: string) {
    try {
      $.ajax({
        type: 'POST',
        url: commonConfig.baseUrl + '/api/Asset/SaveAsset',
        headers: {
          Authorization: 'Bearer ' + token
        },
        dataType: 'json',
        data: JSON.stringify(this.dynamicField),
        contentType: 'application/json',
        success: (result) => {
          this._popupMessage();

          console.log(result);
          return result;
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

  private _popupMessage() {
    let html: string = "";
    if (this._checkURLParameter()) {
      html += `
        <h2>Success</h2>
        <p>Asset successfully updated.</p>
        <button class="closePopup">Close</button>`;
    }
    else {
      html += `
        <h2>Success</h2>
        <p>Asset successfully added.</p>
        <button class="closePopup">Close</button>`;
    }

    const listContainer: Element = this.domElement.querySelector('#popup');
    listContainer.innerHTML = html;

    $(".popup-overlay, .popup-content").addClass("active");
    $(".closePopup").on("click", () => {
      $(".popup-overlay, .popup-content").removeClass("active");
      // If updating asset, reload page
      if (this._checkURLParameter()) {
        location.reload();
      }
      // If adding new asset, redirect to asset list page
      else {
        var url = new URL(`${commonConfig.url}/SitePages/${commonConfig.Page.AssetList}`);
        Navigation.navigate(url.toString(), true);
      }
    });
  }

  private _getAssetById(token: string): void {
    var queryParms = new URLSearchParams(document.location.search.substring(1));
    var myParm = queryParms.get("refNo");
    if (myParm) {
      var refNo = myParm.trim();
      $("#header_title").html(`Update Asset Form`);
      $.ajax({
        type: 'GET',
        url: commonConfig.baseUrl + `/api/Asset/GetAssetByRefNo?refNo=${refNo}`,
        headers: {
          Authorization: 'Bearer ' + token
        },
        success: (result) => {
          this.formDetails = result;
          this._populateFormById(result);
        },
        error: (result) => {
          return result;
        }
      });
    }
    else {
      $("#header_title").html(`Add New Asset Form`);
    }
  }

  private _getAssetFolderID(token: string): void {
    folderID = 0;
    var queryParms = new URLSearchParams(document.location.search.substring(1));
    var myParm = queryParms.get("refNo");
    var refNo = myParm.trim();

    $.ajax({
      type: 'GET',
      url: commonConfig.baseUrl + `/api/Asset/GetAssetByRefNo?refNo=${refNo}`,
      headers: {
        Authorization: 'Bearer ' + token
      },
      success: (result) => {
        console.log(result);
        folderID = result["FolderId"];
        return result;
      },
      error: (result) => {
        console.log(result);
        return result;
      }
    });
  }

  private _checkURLParameter() {
    var queryField = 'refNo';
    var url = window.location.href;

    if (url.indexOf('?' + queryField + '=') != -1) {
      return true;
    }
    else if (url.indexOf('&' + queryField + '=') != -1) {
      return true;
    }
    else {
      return false;
    }
  }

  private _populateFormById(formDetailsList: IDynamicField) {
    let htmlOffice = "";
    let htmlFloor = "";
    let htmlBuilding = "";

    try {
      if (this._checkURLParameter()) {
        $("#btnSubmit").html("Update");
        $('#idAssetName').val(formDetailsList.Name);
        $('#idAssetRefNo').val(formDetailsList.ReferenceNumber);
        $('#idOwnership').val(formDetailsList.Ownership);
        $('#typeOfAssetList').val(formDetailsList.TypeOfAsset);
        $('#idServicingPeriod').val(formDetailsList.ServicingPeriod);
        $('#idComments').val(formDetailsList.Comment);

        this.ListOfBuildings.forEach((item: IBuildings) => {
          if (formDetailsList.BuildingName == item.Title) {
            $('#idBuildingLocation').val(item.Location);
          }
        });

        this._renderFieldRequiredList(this.arrFieldsRequired);

        htmlOffice += `<option value="${formDetailsList.OfficeName}" selected>${formDetailsList.OfficeName}</option>`;
        htmlFloor += `<option value="${formDetailsList.FloorNo}" selected>${formDetailsList.FloorNo}</option>`;
        htmlBuilding += `<option value="${formDetailsList.BuildingName}" selected>${formDetailsList.BuildingName}</option>`;

        this.arrFieldsRequired.forEach((item: IFieldsRequiredList) => {
          if (item.TypeOfAssets.Title == formDetailsList.TypeOfAsset) {
            var itemTitle = item.Title.replace(/ /g, "");

            if (itemTitle.indexOf("Date") >= 0) {
              if (formDetailsList[`${itemTitle}`] != null) {
                $(`#id${itemTitle}`).val(formDetailsList[`${itemTitle}`].substring(0, 10));
              }
              else {
                $(`#id${itemTitle}`).val("");
              }
            }
            else {
              $(`#id${itemTitle}`).val(formDetailsList[`${itemTitle}`]);
            }
          }
        });

        if (formDetailsList.LastServicingDate == null) {
          $('#idLastServicingDate').val("");
        }
        else {
          $('#idLastServicingDate').val(formDetailsList.LastServicingDate.substring(0, 10));
        }

        if ((formDetailsList.ServicingRequired) == true) {
          $('#servicingRequired').prop("checked", true);
          this._checkIfServicingRequiredChecked();
        }
        else {
          $('#servicingNotRequired').prop("checked", true);
          this._checkIfServicingNotRequiredChecked();
        }

        if (formDetailsList.AssetAttachments.length == 0) {
          $('#attachmentTable').hide();
        }
        else {
          $('#attachmentTable').show();
          if (fileInfos.length == 0) {
            formDetailsList.AssetAttachments.forEach(async (file: IAttachmentDetails) => {
              await populatedFileInfos.push({
                "AttachmentGUID": file.AttachmentGUID,
                "AttachmentFileName": file.AttachmentFileName,
                "AttachmentFileContent": file.AttachmentFileContent
              });
            });
            this._populateDownloadInfos();
          }

          this._populateAttachmentTable();
        }

        const listContainer1: Element = this.domElement.querySelector('#idOffice');
        listContainer1.innerHTML = htmlOffice;

        const listContainer2: Element = this.domElement.querySelector('#idFloor');
        listContainer2.innerHTML = htmlFloor;

        const listContainer3: Element = this.domElement.querySelector('#idBuildingName');
        listContainer3.innerHTML = htmlBuilding;

        //Disable all fields on view
        $('#idAssetName').prop('disabled', true);
        $('#idAssetRefNo').prop('disabled', true);
        $('#idFloor').prop('disabled', true);
        $('#idBuildingName').prop('disabled', true);
        $('#idOwnership').prop('disabled', true);
        $('#typeOfAssetList').prop('disabled', true);
        $('#idBuildingLocation').prop('disabled', true);
        $('#idOffice').prop('disabled', true);
      }
    }
    catch (error) {
      console.log(error);
      return error;
    }
  }

  private async _applicationDetails() {
    try {
      var result: boolean = false;
      var servicingReq;
      var iFolderID;
      var attachmentDetails: IAttachmentDetails[] = [];
      var typeOfAssetsValue = (<HTMLInputElement>document.getElementById('typeOfAssetList')).value;
      var BuildingName = (<HTMLInputElement>document.getElementById('idBuildingName')).value;
      var OfficeName = (<HTMLInputElement>document.getElementById('idOffice')).value;
      var BuildingLocation = (<HTMLInputElement>document.getElementById('idBuildingLocation')).value;
      var FloorNo = (<HTMLInputElement>document.getElementById('idFloor')).value;

      if (OfficeName == "" || FloorNo == "" || BuildingName == "" || BuildingLocation == "" || typeOfAssetsValue == "") {
        if (OfficeName == "") {
          $('#requiredOffice').text("* Required Field!");
        }
        if (FloorNo == "") {
          $('#requiredFloor').text("* Required Field!");
        }
        if (BuildingName == "") {
          $('#requiredBuildingName').text("* Required Field!");
        }
        if (BuildingLocation == "") {
          $('#requiredBuildingLocation').text("* Required Field!");
        }
        if (typeOfAssetsValue == "") {
          $('#requiredTypeOfAsset').text("* Required Field!");
        }
        return result;
      }
      else {
        if ($('#servicingNotRequired').is(':checked')) {
          servicingReq = false;
        }
        else if ($('#servicingRequired').is(':checked')) {
          servicingReq = true;
        }

        await this._convertFileToBinary();

        if (fileInfos.length > 0 || this.mainFileByteArray.length > 0) {
          if (this._checkURLParameter() && fileInfos.length > 0) {
            fileInfos.forEach((file: any) => {
              if (!file.file) {
                attachmentDetails.push({
                  AttachmentGUID: file.AttachmentGUID,
                  AttachmentFileName: file.AttachmentFileName,
                  AttachmentFileContent: file.AttachmentFileContent
                });
              }
            });
          }
          if (this.mainFileByteArray.length > 0) {
            this.mainFileByteArray.forEach((file: any) => {
              attachmentDetails.push({
                AttachmentGUID: file.AttachmentGUID,
                AttachmentFileName: file.AttachmentFileName,
                AttachmentFileContent: file.AttachmentFileContent
              });
            });
          }
        }
        else {
          attachmentDetails = [];
        }

        if (folderID > 0) {
          iFolderID = folderID.toString();
        }
        else {
          iFolderID = "";
        }

        this.dynamicField = {
          Name: (<HTMLInputElement>document.getElementById('idAssetName')).value,
          ReferenceNumber: (<HTMLInputElement>document.getElementById('idAssetRefNo')).value,
          BuildingName: (<HTMLInputElement>document.getElementById('idBuildingName')).value,
          OfficeName: (<HTMLInputElement>document.getElementById('idOffice')).value,
          BuildingLocation: (<HTMLInputElement>document.getElementById('idBuildingLocation')).value,
          FloorNo: (<HTMLInputElement>document.getElementById('idFloor')).value,
          Ownership: (<HTMLInputElement>document.getElementById('idOwnership')).value,
          TypeOfAsset: (<HTMLInputElement>document.getElementById('typeOfAssetList')).value,
          LastServicingDate: (<HTMLInputElement>document.getElementById('idLastServicingDate')).value,
          ServicingPeriod: (<HTMLInputElement>document.getElementById('idServicingPeriod')).value,
          Comment: (<HTMLInputElement>document.getElementById('idComments')).value,
          AssetAttachments: attachmentDetails,
          ServicingRequired: servicingReq,
          FolderId: iFolderID
        };

        this.arrFieldsRequired.forEach((item) => {
          if (item.TypeOfAssets.Title == typeOfAssetsValue) {
            var itemTitle = item.Title.replace(/ /g, "");
            this.dynamicField[`${itemTitle}`] = (<HTMLInputElement>document.getElementById(`id${itemTitle}`)).value;
          }
        });

        result = true;
        return result;
      }
    }
    catch (error) {
      console.log(error);
      return error;
    }
  }

  //#region Servicing
  private _addServicingDetailsToList(): Promise<string> {
    return new Promise(async (resolve, reject) => {
      try {
        let listUri: string = '/sites/Lincoln/Lists/Asset%20Servicing';
        var refNo = $('#idAssetRefNo').val();
        var FolderID = null;
        var newTitle;
        this.folderItemGUID = Guid.create();

        var lastServicingDate = moment(new Date((<HTMLInputElement>document.getElementById('idLastServicingDate')).value), "DD-MM-YYYY").format('YYYY-MM-DD');
        var servicingPeriod = parseInt((<HTMLInputElement>document.getElementById('idServicingPeriod')).value);
        var nextServicingDate = moment(lastServicingDate, "YYYY-MM-DD").add(servicingPeriod, 'months').format('YYYY-MM-DD');
        var emailSendDate = moment(nextServicingDate, "YYYY-MM-DD").subtract(21, 'days').format('YYYY-MM-DD');

        //Check if this is a new application and if servicing/test is true
        if ($('#servicingRequired').is(':checked') && !this._checkURLParameter()) {
          console.log("Check if this is a new application and if servicing/test is true");
          sp.web.lists.getByTitle("Asset Servicing").items.add({ Title: `${refNo}`, ContentTypeId: "0x0120" })
            .then((resp) => {
              FolderID = resp.data.ID;
              folderID = resp.data.ID;
              console.log("folder created: " + folderID);

              resolve(folderID.toString());
            })
            .then(() => {
              sp.web.lists.getByTitle("Asset Servicing").items.add({
                Title: this.folderItemGUID.toString(),
                AssetRefNo: refNo,
                AssetName: $("#idAssetName").val(),
                LastServicingDate: lastServicingDate,
                ServicingPeriod: servicingPeriod,
                NextServicingDate: nextServicingDate,
                EmailSendDate: emailSendDate,
                Building: $("#idBuildingName").val(),
                Office: $("#idOffice").val(),
                Floor: $("#idFloor").val(),
                ReminderSent: false,
                ActiveFlag: true
              })
                .then((item: any) => {
                  sp.web
                    .getFileByServerRelativeUrl(`${listUri}/${item.data.ID}_.000`)
                    .moveTo(`${listUri}/${FolderID}_.000/${this.folderItemGUID}`);
                  console.log("Item added to folder.");
                });
            });
        }
        //Check if an existing asset is being updated and if servicing/test is true
        else if ($('#servicingRequired').is(':checked') && this._checkURLParameter()) {
          console.log("Check if an existing asset is being updated and if servicing/test is true");
          this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Asset Servicing')/items?$select=*&$filter=AssetRefNo eq '${refNo}'`, SPHttpClient.configurations.v1)
            .then(response => {
              return response.json()
                .then((items: any): void => {
                  //Check if folder exists
                  if (items["value"].length > 0) {
                    items["value"].forEach((itemF: any) => {
                      var strItemLSD = itemF.LastServicingDate.substring(0, 10);
                      var strLSD = lastServicingDate.toString().substring(0, 10);
                      newTitle = this.folderItemGUID;

                      //Check if last servicing date or servicing period has been updated
                      if (this._strCompare(strItemLSD, strLSD) == 0 || itemF.ServicingPeriod.toString() != servicingPeriod) {
                        //Check if reminder has already been sent
                        if (itemF.ReminderSent) {
                          //Create new line in folder
                          newTitle = this.folderItemGUID;
                        }
                        else {
                          //Update existing line in folder
                          newTitle = itemF.Title;
                        }

                        sp.web.lists.getByTitle("Asset Servicing").items.add({
                          Title: newTitle.toString(),
                          AssetRefNo: refNo,
                          AssetName: $("#idAssetName").val(),
                          LastServicingDate: lastServicingDate,
                          ServicingPeriod: servicingPeriod,
                          NextServicingDate: nextServicingDate,
                          EmailSendDate: emailSendDate,
                          Building: $("#idBuildingName").val(),
                          Office: $("#idOffice").val(),
                          Floor: $("#idFloor").val(),
                          ReminderSent: false,
                          ActiveFlag: true
                        })
                          .then((item: any) => {
                            sp.web
                              .getFileByServerRelativeUrl(`${listUri}/${item.data.ID}_.000`)
                              .moveTo(`${listUri}/${folderID}_.000/${newTitle}`);
                            console.log("Item added to folder UPDATED VERSION.");
                          });
                      }
                      resolve(folderID.toString());
                    });
                  }
                  //If folder doesn't exist
                  else {
                    sp.web.lists.getByTitle("Asset Servicing").items.add({ Title: `${refNo}`, ContentTypeId: "0x0120" })
                      .then((resp) => {
                        FolderID = resp.data.ID;
                        folderID = resp.data.ID;
                        console.log("folder created: " + folderID);

                        resolve(folderID.toString());
                      })
                      .then(() => {
                        sp.web.lists.getByTitle("Asset Servicing").items.add({
                          Title: this.folderItemGUID.toString(),
                          AssetRefNo: refNo,
                          AssetName: $("#idAssetName").val(),
                          LastServicingDate: lastServicingDate,
                          ServicingPeriod: servicingPeriod,
                          NextServicingDate: nextServicingDate,
                          EmailSendDate: emailSendDate,
                          Building: $("#idBuildingName").val(),
                          Office: $("#idOffice").val(),
                          Floor: $("#idFloor").val(),
                          ReminderSent: false,
                          ActiveFlag: true
                        })
                          .then((item: any) => {
                            sp.web
                              .getFileByServerRelativeUrl(`${listUri}/${item.data.ID}_.000`)
                              .moveTo(`${listUri}/${FolderID}_.000/${this.folderItemGUID}`);
                            console.log("Item added to folder.");
                          });
                      });
                  }
                });
            });
        }
        //Check if servicing is no and if asset is being updated
        else if ($('#servicingNotRequired').is(':checked') && this._checkURLParameter()) {
          console.log("Check if servicing is no and if asset is being updated");
          this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Asset Servicing')/items?$select=*&$filter=AssetRefNo eq '${refNo}'`, SPHttpClient.configurations.v1)
            .then(response => {
              return response.json()
                .then((items: any): void => {
                  //Check if folder exists
                  if (items["value"].length > 0) {
                    console.log("Folder exists");
                    items["value"].forEach((itemF: any) => {
                      var strItemLSD = itemF.LastServicingDate.substring(0, 10);
                      var strLSD = lastServicingDate.toString().substring(0, 10);
                      newTitle = this.folderItemGUID;

                      //Check if last servicing date or servicing period has been updated
                      if (this._strCompare(strItemLSD, strLSD) == 0 || itemF.ServicingPeriod.toString() != servicingPeriod) {
                        //Check if reminder has already been sent
                        if (!itemF.ReminderSent) {
                          //Update existing line in folder
                          newTitle = itemF.Title;

                        sp.web.lists.getByTitle("Asset Servicing").items.add({
                          Title: newTitle.toString(),
                          AssetRefNo: refNo,
                          AssetName: $("#idAssetName").val(),
                          LastServicingDate: itemF.LastServicingDate,
                          ServicingPeriod: itemF.ServicingPeriod,
                          NextServicingDate: itemF.NextServicingDate,
                          EmailSendDate: itemF.EmailSendDate,
                          Building: $("#idBuildingName").val(),
                          Office: $("#idOffice").val(),
                          Floor: $("#idFloor").val(),
                          ReminderSent: false,
                          ActiveFlag: false
                        })
                          .then((item: any) => {
                            sp.web
                              .getFileByServerRelativeUrl(`${listUri}/${item.data.ID}_.000`)
                              .moveTo(`${listUri}/${folderID}_.000/${newTitle}`);
                            console.log("Item added to folder UPDATED VERSION.");
                          });
                        }
                      }
                      resolve(folderID.toString());
                    });
                  }
                  else {
                    console.log("Folder does not exist");
                    resolve(folderID.toString());
                  }
                });
            });
        }
        else {
          folderID = "";
          resolve(folderID.toString());
          console.log("Nope");
        }
      }
      catch (error) {
        console.log(error);
        reject(error);
      }
    });
  }

  private _strCompare(date1: string, date2: string): number {
    if (date1 == date2) {
      return 1;
    }
    else {
      return 0;
    }
  }
  //#endregion

  //#region File functions
  private async _convertFileToBinary() {
    try {
      this.mainFileByteArray = [];
      for (var file of fileInfos) {
        if (file.file) {
          let result = await this.readFile(file.file);
          fileByteArray = [];
          filestream = result;
          fixarray = new Uint8Array(filestream);
          for (var element of fixarray) {
            fileByteArray.push(element);
          }
          this.mainFileByteArray.push({
            "AttachmentGUID": file.AttachmentGUID,
            "AttachmentFileName": file.AttachmentFileName,
            "AttachmentFileContent": fileByteArray
          });
        }
      }
    }
    catch (error) {
      console.log(error);
    }
  }

  private readFile(file) {
    return new Promise((resolve, reject) => {
      var fr = new FileReader();
      fr.onload = () => {
        resolve(fr.result);
      };
      fr.onerror = reject;
      fr.readAsArrayBuffer(file);
    });
  }

  private _populateDownloadInfos() {
    return new Promise(async (resolve, reject) => {
      try {
        populatedFileInfos.forEach(async (file: any) => {
          var fileExtension: string = file.AttachmentFileName.substr(file.AttachmentFileName.lastIndexOf(".") + 1);
          await fileDownloadInfos.push({
            "linkSource": `data:${mime.lookup(`${fileExtension}`)};base64,${file.AttachmentFileContent}`,
            "fileName": file.AttachmentFileName,
            "AttachmentGUID": file.AttachmentGUID
          });
          resolve(fileDownloadInfos);
        });
      }
      catch (error) {
        console.log(error);
        reject(error);
      }
    });
  }

  private _uploadToAttachmentTable() {
    this._checkAttachmentTable();
    var fileCount = (<HTMLInputElement>document.getElementById("customFile")).files.length;

    if (fileInfos.length > 0 && tempFileInfos.length == fileCount) {
      $('#attachmentTable').show();

      this._populateAttachmentTable();
    }
  }

  private _populateAttachmentTable() {
    let html: string = "";

    fileInfos.forEach((file: any) => {
      var fileNameNoSpace = file.AttachmentFileName.replace(/ /g, "");
      html += `<tr id="tr_${fileNameNoSpace}_${file.AttachmentGUID}"><td class="th-lg" scope="row">${file.AttachmentFileName}</td>
      <td>
        <ul class="list-inline m-0">
          <li class="list-inline-item delete">
            <button class="btn btn-secondary btn-sm rounded-circle" id="btn_${fileNameNoSpace}_${file.AttachmentGUID}" type="button" data-toggle="tooltip" data-placement="top" title="Delete"><i class="fa fa-trash"></i></button>
          </li>
        </ul>
      </td></tr>`;
    });

    populatedFileInfos.forEach((file: any) => {
      var fileNameNoSpace = file.AttachmentFileName.replace(/ /g, "");
      html += `<tr id="tr_${fileNameNoSpace}_${file.AttachmentGUID}"><td class="th-lg" scope="row">${file.AttachmentFileName}</td>
      <td>
        <ul class="list-inline m-0">
          <li class="list-inline-item view">
            <button class="btn btn-secondary btn-sm rounded-circle" id="btn_${fileNameNoSpace}_${file.AttachmentGUID}" type="button" data-toggle="tooltip" data-placement="top" title="View"><i class="fa fa-download"></i></button>
          </li>
          <li class="list-inline-item delete">
            <button class="btn btn-secondary btn-sm rounded-circle" id="btn_${fileNameNoSpace}_${file.AttachmentGUID}" type="button" data-toggle="tooltip" data-placement="top" title="Delete"><i class="fa fa-trash"></i></button>
          </li>
        </ul>
      </td></tr>`;
    });

    const listContainer: Element = this.domElement.querySelector('#tableAttachmentContainer');
    listContainer.innerHTML = html;

    $("#tableAttachmentContainer").on('click', '.view', function () {
      try {
        var trid = $(this).closest('tr').attr('id').substring(3);
        var tridFields = trid.split('_');
        var tridFileName = tridFields[0];
        var tridId = tridFields[1];

        if (fileDownloadInfos.length > 0) {
          fileDownloadInfos.forEach((file: any) => {
            var fileNameReplace = file.fileName.replace(/ /g, "");
            if (tridFileName == fileNameReplace && tridId == file.AttachmentGUID) {
              const linkSource = `${file.linkSource}`;
              const downloadLink = document.createElement("a");
              downloadLink.href = linkSource;
              downloadLink.download = file.fileName;
              downloadLink.click();

            }
          });
        }
      }
      catch (error) {
        console.log(error);
        return error;
      }
    });

    $("#tableAttachmentContainer").on('click', '.delete', function () {
      try {
        var trid = $(this).closest('tr').attr('id').substring(3);
        var tridFields = trid.split('_');
        var tridFileName = tridFields[0];
        var tridId = tridFields[1];

        if (fileInfos.length > 0) {
          $(this).closest('tr').remove();
          fileInfos.forEach((file: any) => {
            var fileNameReplace = file.AttachmentFileName.replace(/ /g, "");
            if (tridFileName == fileNameReplace && tridId == file.AttachmentGUID) {
              fileInfos = fileInfos.filter(item => item.AttachmentGUID !== file.AttachmentGUID);
              if (fileInfos.length == 0) {
                $('#attachmentTable').hide();
              }
            }
          });
        }
        else {
          $(this).closest('tr').remove();
          $('#attachmentTable').hide();
        }
      }
      catch (error) {
        console.log(error);
        return error;
      }
    });
  }

  private blob() {
    var input = (<HTMLInputElement>document.getElementById("customFile"));
    var fileCount = input.files.length;
    try {
      tempFileInfos = [];
      for (var i = 0; i < fileCount; i++) {
        var file = input.files[i];
        var reader = new FileReader();
        reader.onload = ((file1) => {
          return (e) => {
            this.fileGUID = Guid.create();

            fileInfos.push({
              "AttachmentGUID": this.fileGUID.toString(),
              "AttachmentFileName": file1.name,
              "AttachmentFileContent": e.target.result,
              "file": file1
            });

            tempFileInfos.push({
              "AttachmentGUID": this.fileGUID.toString(),
              "AttachmentFileName": file1.name,
              "AttachmentFileContent": e.target.result
            });

            this._uploadToAttachmentTable();
          };
        })(file);
        reader.readAsArrayBuffer(file);
      }
    }
    catch (error) {
      return error;
    }
  }
  //#endregion

  private _checkAttachmentTable(): void {
    var myFile = (<HTMLInputElement>document.getElementById('customFile')).files;

    if (myFile.length == 0) {
      $('#attachmentTable').hide();
    }
  }

  private _checkIfServicingNotRequiredChecked(): void {
    if ($('#servicingNotRequired').is(':checked')) {
      $('#idLastServicingDate').prop("disabled", true);
      $('#idServicingPeriod').prop("disabled", true);
      $('#idLastServicingDate').val("");
      $('#idServicingPeriod').val("");
    }
  }

  private _checkIfServicingRequiredChecked(): void {
    if ($('#servicingRequired').is(':checked')) {
      $('#idLastServicingDate').prop("disabled", false);
      $('#idServicingPeriod').prop("disabled", false);
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
