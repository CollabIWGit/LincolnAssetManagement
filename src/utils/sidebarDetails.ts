import * as $ from "jquery";
import { Navigation } from 'spfx-navigation';
import * as commonConfig from "./commonConfig.json";

export class sidebarDetails {
    public sidebarMenu(absoluteURL: string) {
        var navbar = `
        <nav id="sidebar">
          <!--<div class="custom-menu">
              <button type="button" id="sidebarCollapse" class="btn btn-primary">
                <i class="fa fa-bars"></i>
                <span class="sr-only">Toggle Menu</span>
            </div>-->
          <img id="imgLogo" src="${absoluteURL}/SiteAssets/Lincoln-Realty-Logo-orange.png"
            alternate="lincoln-logo">
          <ul class="list-unstyled components mb-5">
            <li class="active">
              <a id="homePage">
                <span class="fa fa-home mr-3"> </span>Home
              </a>
            </li>
            <li id="adminMgtComponent">
              <a id="adminMgt">
                <span class="fa fa-sliders-h mr-3"> </span>Admin Management
              </a>
            </li>
            <li>
              <a href="#caseManagementUl" id="CaseMgt" data-toggle="collapse" aria-expanded="false" class="dropdown-toggle">
                <span class="fas fa-file-contract mr-3"> </span>Case Management
              </a>
              <ul class="list-unstyled collapse" style="list-style-type:none;" id="caseManagementUl">
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
            </li>
            <li>
              <a href="#assetManagementUl" id="AssetMgt" data-toggle="collapse" aria-expanded="false" class="dropdown-toggle">
                <span class="fas fa-folder-open mr-3"> </span>Asset Management
              </a>
              <ul class="list-unstyled collapse" style="list-style-type:none;" id="assetManagementUl">
                <li>
                  <a id="assetDashboardPage">
                    <span class="fa fa-list"> </span>  List of Assets
                  </a>
                </li>
                <li>
                  <a id="addAssetPage">
                    <span class="fa fa-plus"> </span>  Add new Asset
                  </a>
                </li>
              </ul>
            </li>
          </ul>
        </nav>`;
        $("#nav-placeholder").html(navbar);
        this.sidebarNavigation(absoluteURL);
    }

    public sidebarNavigation(absoluteURL: string) {
        $("#homePage").on("click", () => {
            Navigation.navigate(`${absoluteURL}/SitePages/${commonConfig.Page.HomePage}`, true);
        });
        $("#caseList").on("click", () => {
            Navigation.navigate(`${absoluteURL}/SitePages/${commonConfig.Page.CaseDashboard}`, true);
        });
        $("#addCase").on("click", () => {
            Navigation.navigate(`${absoluteURL}/SitePages/${commonConfig.Page.AddCase}`, true);
        });
        $("#addAssetPage").on("click", () => {
            Navigation.navigate(`${absoluteURL}/SitePages/${commonConfig.Page.AddAssets}`, true);
        });
        $("#assetDashboardPage").on("click", () => {
            Navigation.navigate(`${absoluteURL}/SitePages/${commonConfig.Page.AssetDashboard}`, true);
        });
    }
}