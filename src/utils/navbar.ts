import * as commonConfig from "../utils/commonConfig.json";

export class navbar {
    public cover: string = `<div id="cover"> <i class="fa-solid fa-rotate"></i> loading...</div>`;

    public navbar: string = `
    <div id="sidebar-wrapper">
        <img id="imgLogo" src="${commonConfig.url}/SiteAssets/Lincoln-Realty-Logo-orange.png" alternate="lincoln-logo">
        <ul class="list-unstyled components mb-5">
            <li>
                <a id="home">
                    <span class="fa fa-home mr-3"> </span>Home
                </a>
            </li>
            <li>
                <a id="CaseMgt" href="#caseMngtSubmenu" data-toggle="collapse" aria-expanded="false" class="dropdown-toggle collapsed">
                    <span class="fas fa-file-contract mr-3"> </span>Case Management
                </a>
                <ul class="collapse list-unstyled" id="caseMngtSubmenu" style="list-style-type:none;">
                    <li>
                        <a id="caseList">
                            <span class="fa fa-list"> </span> List of Case
                        </a>
                    </li>
                    <li>
                        <a id="addCase">
                            <span class="fa fa-plus"> </span> Add new Case
                        </a>
                    </li>
                </ul>
            </li>
            <li>
                <a id="AssetMgt" href="#assetMngtSubmenu" data-toggle="collapse" aria-expanded="false" class="dropdown-toggle collapsed">
                    <span class="fas fa-folder-open mr-3"></span>Asset Management
                </a>
                <ul class="collapse list-unstyled" id="assetMngtSubmenu" style="list-style-type:none;">
                    <li>
                        <a id="assetList">
                            <span class="fa fa-list"> </span> List of Assets
                        </a>
                    </li>
                    <li>
                        <a id="addAsset">
                            <span class="fa fa-plus"> </span> Add new Asset
                        </a>
                    </li>
                </ul>
            </li>
            <li id="userMgtComponent">
                <a id="UserMgt" href="#UserMngtSubmenu" data-toggle="collapse" aria-expanded="false" class="dropdown-toggle collapsed">
                    <span class="fa fa-user mr-3"> </span>User Management
                </a>
                <ul class="collapse list-unstyled" id="UserMngtSubmenu" style="list-style-type:none;">
                    <li>
                        <a id="usersList">
                            <span class="fa fa-list"> </span> List of Users
                        </a>
                    </li>
                    <li>
                        <a id="addUser">
                            <span class="fa fa-plus"> </span> Add new User
                        </a>
                    </li>
                </ul>
            </li>
            <li>
                <a id="AdminManagement" href="#adminMngtSubmenu" data-toggle="collapse" aria-expanded="false" class="dropdown-toggle collapsed">
                    <span class="fa fa-user mr-3"> </span>Admin Management
                </a>
                <ul class="collapse list-unstyled" id="adminMngtSubmenu" style="list-style-type:none;">
                    <li>
                        <a id="officesList">
                            <span class="fa fa-list mr-1"> </span> Offices List
                        </a>
                    </li>
                    <li>
                        <a id="typeOfAsset">
                            <span class="fa fa-list mr-1"> </span>Type of Assets List
                        </a>
                    </li>
                    <li>
                        <a id="natureOfProblem">
                            <span class="fa fa-list mr-1"> </span>Nature Of Problem List
                        </a>
                    </li>
                </ul>
            </li>

            <li> 
            <a id="report">
                <span class="fa fa-chart-area mr-3"> </span>Reporting
            </a>
            
            </li>
        </ul>
    </div>`;
}