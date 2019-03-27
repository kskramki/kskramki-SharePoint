import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './GetUnreadEmailsWebPart.module.scss';
import * as strings from 'GetUnreadEmailsWebPartStrings';
import {MSGraphClient} from  '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

export interface IGetUnreadEmailsWebPartProps {
  description: string;
}

export default class GetUnreadEmailsWebPart extends BaseClientSideWebPart<IGetUnreadEmailsWebPartProps> {
  
private Gcallback(r:any):void
{
  if(r==null)
 console.log("Operation Success") ;
 else
 console.log(r);

}
  public render(): void {

    this.context.msGraphClientFactory.getClient().then((client:MSGraphClient):void =>{
var content={"message": {"subject": "Hello Guys!Meet for lunch?", "body": {
      "contentType": "Text",
      "content": "The new cafeteria is open."
    },
    "toRecipients": [
      {
        "emailAddress": {
          "address": "ramcts@ramcts.onmicrosoft.com"
        }
      }
    ],
    
  },
  "saveToSentItems": "false"
};
//client.api("/me/mailFolders/AAMkAGMwYjhjOTdkLTk0ZWItNDRlNC1iN2RlLTRiYjE0Y2Y3ZDRhYgAuAAAAAADLNNuaWUwNTYuIOsdOe2FYAQB06bCmaxbCTJXvFWsW58fKAAAAAAEMAAA=/messages").get((error, rawResponse?: any) => {
  client.api("/users/ram@spcrackers.onmicrosoft.com/sendmail").post(content,this.Gcallback);
  
    client.api("/users/606e8b6c-36c2-4294-ab69-14944b5764a2/manager").get().then(response=>{
      console.log("Managere",response);
      console.log(response.mail);
    });
    client.api("/users/606e8b6c-36c2-4294-ab69-14944b5764a2/photo/$value").get().then(response=>{
      console.log("Managere",response);
     let  imgblob = response.blob();
      let imgsrc = window.URL.createObjectURL(imgblob);
      console.log(imgsrc);
    });
  });
    this.domElement.innerHTML = `
      <div class="${ styles.getUnreadEmails }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
              <p class="${ styles.description }">${escape(this.properties.description)}</p>
              <a href="https://aka.ms/spfx" class="${ styles.button }">
                <span class="${ styles.label }">Learn more</span>
                
              </a>
            </div>
          </div>
        </div>
      </div>`;
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
