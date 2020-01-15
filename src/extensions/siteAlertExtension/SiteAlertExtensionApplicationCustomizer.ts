import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer, 
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'SiteAlertExtensionApplicationCustomizerStrings';
import styles from './AppCustomizer.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from "@pnp/sp";
import { SPFetchClient } from "@pnp/nodejs";
import { SPHttpClient, SPHttpClientResponse, SPHttpClientConfiguration } from '@microsoft/sp-http';  
import { getIconClassName } from '@uifabric/styling';



const LOG_SOURCE: string = 'SiteAlertExtensionApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ISiteAlertExtensionApplicationCustomizerProperties {
  // This is an example; replace with your own property
  Top: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class SiteAlertExtensionApplicationCustomizer
  extends BaseApplicationCustomizer<ISiteAlertExtensionApplicationCustomizerProperties> {
    // These have been added
    private _topPlaceholder: PlaceholderContent | undefined;
    private _bottomPlaceholder: PlaceholderContent | undefined;
  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    // Wait for the placeholders to be created (or handle them being changed) and then
    // render.
    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);

    this._getSiteTitle();

    return Promise.resolve<void>();
  }
  private _getSiteTitle(): void {
    this.context.spHttpClient
    .get(`${this.context.pageContext.web.absoluteUrl}/_api/web?$select=Title`, SPHttpClient.configurations.v1)
    .then((res: SPHttpClientResponse): Promise<{ Title: string; }> => {
      return res.json();
    })
    .then((web: {Title: string}): void => {
      console.log(web.Title);
      this.properties.Top = web.Title;
    });
  }

  private _renderPlaceHolders(): void {
    console.log("Hi!");
    console.log("SiteAlertApplicationCustomizer._renderPlaceHolders()");
    console.log(
      "Available placeholders: ",
      this.context.placeholderProvider.placeholderNames
        .map(name => PlaceholderName[name])
        .join(", ")
    );

    // Handling the top placeholder
    if (!this._topPlaceholder) {
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top,
        { onDispose: this._onDispose }
      );

      // The extension should not assume that the expected placeholder is available.
      if (!this._topPlaceholder) {
        console.error("The expected placeholder (Top) was not found.");
        return;
      }

      if (this.properties) {

        this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/lists/GetByTitle('Alert')/items?$select=Title,Link,AlertType,Display&$orderby=Created%20desc&$top=1`,  
                SPHttpClient.configurations.v1)  
                .then((response: SPHttpClientResponse) => {  
                  response.json().then((responseJSON: any) => {  
                    let topString: string = responseJSON.value[0].Title;
                    // console.log ("Writing Object");
                    // console.log (responseJSON.value[0].Title);
                    let showFlag: string = responseJSON.value[0].Display;
                    let bannerColor: string = "";
                    let bannerIcon: string = "";
                    if (!topString) {
                      topString = "(Top property was not defined.)";
                    }

                    // Want to determine if banner needs shown ////////////////////////////
                    if(responseJSON.value[0].Display == "Hide") {
                      showFlag = "none";
                    } else  {
                      showFlag = "block";
                    }

                    //Making Adjustements for Each Alert Type /////////////////////
                    switch (responseJSON.value[0].AlertType) {
                      case "Other":
                        bannerColor = "#FFE380";
                        bannerIcon = "Info";
                      break;

                      case "System":
                        bannerColor = "#FFAC61";
                        bannerIcon = "DataManagementSettings";
                      break;

                      case "Safety/Security":  
                        bannerColor = "#ff8979";
                        bannerIcon = "ShieldAlert";
                      break;

                      default:
                        bannerColor = "#002d72";
                        bannerIcon = "";
                      break;
                    }


                    // Generating DOM for Banner /////////////////////////////////////
                    if (this._topPlaceholder.domElement) {
                      this._topPlaceholder.domElement.innerHTML = `
                      <a href="${responseJSON.value[0].Link}">
                      <div class="${styles.app}" style="display:${showFlag};">
                        <div class="${styles.top}" style="background-color:${bannerColor};" href="${responseJSON.value[0].Link}">
                          <i class="${getIconClassName(bannerIcon)}" aria-hidden="true"></i> ${escape(
                            topString
                          )}
                        </div>
                      </div></a>`;
                    }
                  });  
                }); 


      }
    }
  }
  private _onDispose(): void {
    console.log('[SiteAlertExtensionApplicationCustomizer._onDispose] Disposed custom top and bottom placeholders.');
  }
}
