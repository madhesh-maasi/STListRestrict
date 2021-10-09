import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';
import pnp from 'sp-pnp-js';
import * as strings from 'ListRestrictionCustomApplicationCustomizerStrings';
import "../../ExternalRef/Css/alertify.min.css";
import "../../ExternalRef/Css/style.css";
import "alertifyjs";

var alertify: any = require("../../ExternalRef/Js/alertify.min.js");
const LOG_SOURCE: string = 'ListRestrictionCustomApplicationCustomizer';
let restrictedUrlArr = [];

/** 
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IListRestrictionCustomApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
} 

/** A Custom Action which can be run during execution of a Client Side Application */
export default class ListRestrictionCustomApplicationCustomizer
  extends BaseApplicationCustomizer<IListRestrictionCustomApplicationCustomizerProperties> {

  @override
  public onInit(): Promise<void> {
 
    return super.onInit().then(() => {  
      pnp.setup({
         spfxContext: this.context
         });
         this.getListItems();
        //  return Promise.resolve();
        document.querySelectorAll(".ms-HorizontalNavItem-link").forEach((btn)=>{
          btn.addEventListener('click',()=>{
            this.getListItems();
          })
        })
    }); 
  }
  async getListItems(){
    
    let IsCurrentUserAdmin = false;
    var homepageURL=this.context.pageContext.web.absoluteUrl;
    var currentUserName=this.context.pageContext.user.email;
    console.log(currentUserName);
    await pnp.sp.web.lists.getByTitle('UserDJCList').items.select("*,User/Title,User/EMail").expand("User").get().then((allItems: any[])=>{
      console.log(allItems);
      let AdminDetail =  allItems.filter((li)=> li.SuperAdmin == true && li.User.EMail ==currentUserName );
      console.log(AdminDetail);
      AdminDetail.length > 0 ?IsCurrentUserAdmin = true:IsCurrentUserAdmin = false;
    });
    
    
    await pnp.sp.web.lists.getByTitle('RestrictedLists').items.select("Title").get().then((allItems: any[]) => {
      // console.log(allItems)
      if(allItems.length>0)
      {
        for (var index = 0; index < allItems.length; index++) {
          var splitString=allItems[index].Title.split('?');
          restrictedUrlArr.push(splitString[0]);
        }
      }
    });
    if(!IsCurrentUserAdmin){
      document.querySelector(".ms-FocusZone.ms-CommandBar")["style"].display="none";
      if(restrictedUrlArr.length>0)
{
  var locationURL=window.location.href.toLowerCase().split('?');
  var splittednewLoc=locationURL[0];
  let result = restrictedUrlArr.filter(function (urlvalue) {
    var toLower=urlvalue.toLowerCase();
    return splittednewLoc==toLower;
});
// console.log(result);
if(result.length>0)
{ 
  // let message: string = "Sorry! You are not authorized to access this page";
  // let answer = window.confirm(message);
  // if(answer){
  //   window.location.href = "https://stconsultingcomau.sharepoint.com/sites/iTimeSheet"
  // }
  let message: string = "Sorry! You are not authorized to access this page";
  alertify.alert(message, function() {
      window.location.href=homepageURL;
  }).set({ 'closable':false})
  .setHeader("<em> Alert </em> ");
}
}
    }
  }
}
