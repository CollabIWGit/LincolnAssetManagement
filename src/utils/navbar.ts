export class navbar {
    public cover: string = `<div id="cover"> <span class="glyphicon glyphicon-refresh w3-spin preloader-Icon"></span> loading...</div>`;

    public navbar: string = `
    <div id="sidebar-wrapper">
        <img id="imgLogo" src="https://frcidevtest.sharepoint.com/sites/Lincoln/SiteAssets/Lincoln-Realty-Logo-orange.png" alternate="lincoln-logo">
        <ul class="list-unstyled components mb-5">
            <li>
                <a id="home">
                    <span class="fa fa-home mr-3"> </span>Home
                </a>
            </li>
                <li id="adminMgtComponent">
                    <a id="adminMgt">
                        <span class="fa fa-sliders-h mr-3"> </span>Admin Management
                    </a>
                </li>
            <li id="userMgtComponent">
                <a id="UserMgt">
                    <span class="fa fa-user mr-3"> </span>User Management
                </a>
                <div class="collapse3 collapse">
                    <ul style="list-style-type:none;" id="caseManagementUl">
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
                </div>
            </li>
            <li>
                <a id="CaseMgt">
                    <span class="fas fa-file-contract mr-3"> </span>Case Management
                </a>
                <div class="collapse1 collapse">
                    <ul style="list-style-type:none;" id="caseManagementUl">
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
                                <span class="fa fa-list"> </span> List of Assets
                            </a>
                        </li>
                        <li>
                            <a id="addAsset">
                                <span class="fa fa-plus"> </span> Add new Asset
                            </a>
                        </li>
                    </ul>
                </div>
            </li>
        </ul>
    </div>`;
}