import { override } from "@microsoft/decorators";
import { BaseApplicationCustomizer } from "@microsoft/sp-application-base";
import pnp from "sp-pnp-js";
import { graph } from "@pnp/graph/presets/all";
import "../../ExternalRef/Css/alertify.min.css";
import "../../ExternalRef/Css/style.css";
import "alertifyjs";

const alertify: any = require("../../ExternalRef/Js/alertify.min.js");
let restrictedUrlArr = [];

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */

export interface IListRestrictionApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class ListRestrictionApplicationCustomizer extends BaseApplicationCustomizer<IListRestrictionApplicationCustomizerProperties> {
  @override
  public onInit(): Promise<void> {
    return super.onInit().then(() => {
      pnp.setup({
        spfxContext: this.context,
      });

      graph.setup({
        spfxContext: this.context,
      });

      this.getListItems();
      document.querySelectorAll(".ms-HorizontalNavItem-link").forEach((btn) => {
        btn.addEventListener("click", () => {
          this.getListItems();
        });
      });
    });
  }

  async getListItems() {
    let IsCurrentUserAdmin = false;
    let homepageURL = this.context.pageContext.web.absoluteUrl;
    let currentUserName = this.context.pageContext.user.email;

    await pnp.sp.web.siteGroups
      .getByName("Pernix_Connect_Super_Admin")
      .users.get()
      .then((allItems: any[]) => {
        IsCurrentUserAdmin = allItems.some(
          (e: any) => e.Email.toLowerCase() === currentUserName.toLowerCase()
        );
      })
      .catch((err: any) => {
        console.log("Err >> ", err);
      });

    await pnp.sp.web.lists
      .getByTitle("RestrictedLists")
      .items.select("Title")
      .get()
      .then((allItems: any[]) => {
        if (allItems.length > 0) {
          for (let index = 0; index < allItems.length; index++) {
            let splitString = allItems[index].Title.split("?");
            restrictedUrlArr.push(splitString[0]);
          }
        }
      })
      .catch((err: any) => {
        console.log("Err >> ", err);
      });

    if (!IsCurrentUserAdmin) {
      document.querySelector(".ms-FocusZone.ms-CommandBar")["style"].display =
        "none";
      if (restrictedUrlArr.length > 0) {
        let locationURL = window.location.href.toLowerCase().split("?");
        let splittednewLoc = locationURL[0];
        let result = restrictedUrlArr.filter(function (urlvalue) {
          let toLower = urlvalue.toLowerCase();
          return splittednewLoc == toLower;
        });

        if (result.length > 0) {
          let message: string =
            "Sorry! You are not authorized to access this page";
          alertify
            .alert(message, function () {
              window.location.href = homepageURL;
            })
            .set({ closable: false })
            .setHeader("<em> Alert </em> ");
        }
      }
    }
  }
}
