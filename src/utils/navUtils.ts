import * as $ from 'jquery';
import { Navigation } from 'spfx-navigation';
import * as commonConfig from "./commonConfig.json";

export class navUtils {
    public navTriggers() {
        $('#caseList').on('click', () => {
            Navigation.navigate(`https://lincolnrealtymu.sharepoint.com/sites/Lincoln/SitePages/${commonConfig.Page.CaseList}`, true);
        });

        $('#usersList').on('click', () => {
            Navigation.navigate(`https://lincolnrealtymu.sharepoint.com/sites/Lincoln/SitePages/${commonConfig.Page.UsersList}`, true);
        });

        $('#addUser').on('click', () => {
            Navigation.navigate(`https://lincolnrealtymu.sharepoint.com/sites/Lincoln/SitePages/${commonConfig.Page.AddUser}`, true);
        });

        $('#btnAdd').on('click', () => {
            Navigation.navigate(`https://lincolnrealtymu.sharepoint.com/sites/Lincoln/SitePages/${commonConfig.Page.AddCase}`, true);
        });

        $('#btnAddUser').on('click', () => {
            Navigation.navigate(`https://lincolnrealtymu.sharepoint.com/sites/Lincoln/SitePages/${commonConfig.Page.AddUser}`, true);
        });

        $('#addCase').on('click', () => {
            Navigation.navigate(`https://lincolnrealtymu.sharepoint.com/sites/Lincoln/SitePages/${commonConfig.Page.AddCase}`, true);
        });

        $('#home').on('click', () => {
            Navigation.navigate(`https://lincolnrealtymu.sharepoint.com/sites/Lincoln/SitePages/${commonConfig.Page.HomePage}`, true);
        });

        $('#addAsset').on('click', () => {
            Navigation.navigate(`https://lincolnrealtymu.sharepoint.com/sites/Lincoln/SitePages/${commonConfig.Page.AddAssets}`, true);
        });

        $('#assetList').on('click', () => {
            Navigation.navigate(`https://lincolnrealtymu.sharepoint.com/sites/Lincoln/SitePages/${commonConfig.Page.AssetList}`, true);
        });
    }

    public collapse() {
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

        $("#UserMgt").hover(
            () => {
                (<any>$(".collapse3")).show();
            },
            () => {
                (<any>$(".collapse3")).hide();
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

        $(".collapse3").hover(
            () => {
                (<any>$(".collapse3")).show();
            },
            () => {
                (<any>$(".collapse3")).hide();
            }
        );

        $("#btnLocation").click(() => {
            (<any>$(".collapsecardLocation")).slideToggle(500);
        });

        $("#btnOffice").click(() => {
            (<any>$(".collapsecardOffice")).slideToggle(500);
        });

        $("#sidebarCollapse").click(() => {
            (<any>$("#sidebar")).slideToggle(500);
        });
    }

    public cover() {
        $("#cover").fadeOut(4000);
        $("#menu-toggle").click((e) => {
            e.preventDefault();
            $("#wrapper").toggleClass("toggled");
        });
    }
}