import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SpDataTeamsWebPart.module.scss';
import * as strings from 'SpDataTeamsWebPartStrings';
import * as microsoftTeams from '@microsoft/teams-js';
import { SPComponentLoader } from '@microsoft/sp-loader';
import pnp, { sp, WebPart, Web, Item } from 'sp-pnp-js';
SPComponentLoader.loadCss("https://maxcdn.bootstrapcdn.com/bootstrap/3.4.0/css/bootstrap.min.css");
import {
  SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions
} from '@microsoft/sp-http';

export interface ISpDataTeamsWebPartProps {
  description: string;
}

export default class SpDataTeamsWebPart extends BaseClientSideWebPart<ISpDataTeamsWebPartProps> {


  private teamsContext: microsoftTeams.Context;

  protected onInit(): Promise<any> {
    let retVal: Promise<any> = Promise.resolve();
    if (this.context.microsoftTeams) {
      retVal = new Promise((resolve, reject) => {
        this.context.microsoftTeams.getContext(context => {
          this.teamsContext = context;
          resolve();
        });
      });
    }
    return retVal;
  }

  protected getData() {
    let web = new Web(this.context.pageContext.web.absoluteUrl);
    console.log(this.context.pageContext.web.absoluteUrl + "/_api/web/Lists/GetByTitle('TeamMembers')/items?");
    this.domElement.querySelector('#dataSelector').addEventListener('click', () => {
      pnp.sp.web.lists.getByTitle('TeamMembers').items.get().then((result) => {
        console.log(result);
        this.buildHTML(result);
      }).catch((error) => {
        console.log("Something went wrong" + error);
      });
      // let requestUrl = "https://.sharepoint.com/sites//_api/web/Lists/GetByTitle('TeamMembers')/items?";
      // this.context.spHttpClient.get(requestUrl, SPHttpClient.configurations.v1)
      //   .then((response: SPHttpClientResponse) => {
      //     if (response.ok) {
      //       response.json().then((responseJSON) => {
      //         if (responseJSON != null && responseJSON.value != null) {
      //           console.log(responseJSON.value);
      //           this.buildHTML(responseJSON.value);
      //         }
      //       });
      //     }
      //   });
    });
  }

  public buildHTML(res) {
    let bdy = "<table style=\"border:1px solid #ddd!important;\" class=\"table\"><thead style=\"background: #337ab7;color: white;\"><tr><td>Title</td><td>First Name</td><td>LastName</td></tr></thead><tbody>";
    res.map((item, i) => {
      bdy += `
      <tr>
        <td>"${item.Title}"</td> 
        <td>"${item.FirstName}"</td> 
        <td>"${item.LastName}"</td> 
     </tr>`;
    });
    bdy += "</tbody></table>";
    document.getElementById("tbl").innerHTML = bdy;
  }


  public render(): void {
    let title: string = '';
    let siteTabTitle: string = '';
    if (this.teamsContext) {
      title = "Welcome to MS Teams!";
      siteTabTitle = "Team: " + this.teamsContext.teamName;
    }
    else {
      title = "Welcome to SharePoint!";
      siteTabTitle = "SharePoint site: " + this.context.pageContext.web.title;
    }
    this.domElement.innerHTML = `
      <div class="${ styles.spDataTeams}">        
            <div class="panel panel-primary">
                <div class="panel-heading">Display List Items Using SPfX in MS Teams </div>
                <div class="panel-body">
                    <div class="row-fluid">
                    <button type="button" id="dataSelector" class="btn btn-primary">Get SharePoint List Data</button>
                    </div>
                    <br>
                    <br>
                    <div class="row-fluid" id="tbl"></div>
                </div>
              <br> 
              <br> 
            </div>
        </div>`;
    this.getData();
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
