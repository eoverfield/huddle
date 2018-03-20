//based on https://github.com/SharePoint/sp-dev-fx-extensions/tree/master/samples/js-application-run-once
//by Mikael Svensen http://www.techmikael.com/
import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import { Site } from 'sp-pnp-js';
import { GraphHttpClient, HttpClientResponse, IGraphHttpClientOptions } from '@microsoft/sp-http';

import * as strings from 'CollabCommSiteSetupApplicationCustomizerStrings';

const LOG_SOURCE: string = 'CollabCommSiteSetupApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ICollabCommSiteSetupApplicationCustomizerProperties {
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class CollabCommSiteSetupApplicationCustomizer
  extends BaseApplicationCustomizer<ICollabCommSiteSetupApplicationCustomizerProperties> {

  @override
  public onInit(): Promise<void> {
      // Need to be admin in order to remove the customizer - if not skip doing the work
      // For Group sites, the owners will be site admins
      let isSiteAdmin = this.context.pageContext.legacyPageContext.isSiteAdmin;

      if (isSiteAdmin) {
          this.DoWork();
      }
      
      return Promise.resolve();
  }

  private async DoWork() {
    
    // use await if you want to block the dialog before continue
    //await Dialog.alert(data);
    
    // Group ID is not directly available yet, so we get it from the legacy context
    let groupId = this.context.pageContext.legacyPageContext.groupId;

    //go and get image binary data we want to submit.
    //TODO: open panel and ask to request image be provided, or use default one.
    //idea to obtain binary from existing image:
    //https://stackoverflow.com/questions/934012/get-image-data-in-javascript

    //now with binary image, post/patch to graph
    //https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/api/profilephoto_update
    this.context.graphHttpClient.post(`v1.0/groups`,GraphHttpClient.configurations.v1,{
      body: JSON.stringify({"description": "Self help community for library",
        "displayName": "Library Assist",
        "groupTypes": [
          "Unified"
        ],
        "mailEnabled": true,
        "mailNickname": "library",
        "securityEnabled": false
      })  
    }).then((response: HttpClientResponse) => {
      if (response.ok) {
        //if we got back a valid response from graph call, then go ahead and remove customizer
        this.removeCustomizer();
      } else {
        //else keep customizer in place and set up warning
        console.warn(response.status);
      }
    });
  }

  private async removeCustomizer() {
      // Remove custom action from current sute
      let site = new Site(this.context.pageContext.site.absoluteUrl);
      let customActions = await site.userCustomActions.get(); // if installed as web scope, change this line
      
      for (let i = 0; i < customActions.length; i++) {
          var instance = customActions[i];
          if (instance.ClientSideComponentId === this.componentId) {
              await site.userCustomActions.getById(instance.Id).delete();
              console.log("Extension removed");
              // reload the page once done if needed
              window.location.href = window.location.href;
              break;
          }
      }
  }
}
