import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './ListCreationWebPart.module.scss';
import * as strings from 'ListCreationWebPartStrings';
import { SPHttpClient,SPHttpClientResponse, ISPHttpClientOptions } from "@microsoft/sp-http";

export interface IListCreationWebPartProps {
  description: string;
}

export default class ListCreationWebPart extends BaseClientSideWebPart<IListCreationWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.listCreation }">
<h3>Creating a New List Dynamically </h3><br/><br/></br>

<p>Please fill out the below details to create a new list programatically </p><br/><br/>
New List Name: <br/><input type="text" id="txtnewListName"/><br/><br/>

New List Description: <br/><input type="text" id="txtnewListDescription"/><br/><br/>
<button id="btnCreateList">Create List</button><br/>
      </div>`;
      this.bindEvents();
  }

  private bindEvents(): void{
    this.domElement.querySelector('#btnCreateList').addEventListener('click', () => {
      this.createNewList();
    })
  }


  createNewList(): void{
    var newListName = this.domElement.querySelector('#txtnewListName')['value'];
    var newListDescription = this.domElement.querySelector('#txtnewListDescription')['value'];
    // var web = this.context.pageContext.web;
    // var ctx = this.context;
    // var url = web.absoluteUrl + "/_api/web/lists/add";
    const listUrl: string = this.context.pageContext.web.absoluteUrl + "/_api/web/lists/GetByTitle('"+newListName+"')";

    this.context.spHttpClient.get(listUrl, SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse) => {
      if(response.status===200){
        alert('A List already does exist with this name.');
        return;
      }
      if(response.status===404){
        const url: string = this.context.pageContext.web.absoluteUrl + "/_api/web/lists";
        const listDefinition: any = { "Title": newListName, "Description": newListDescription, "AllowContentTypes": true,  "BaseTemplate": 100, "ContentTypesEnabled": true };

        const spHttpClientOptions: ISPHttpClientOptions = {
          body: JSON.stringify(listDefinition)
        };
        this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
        .then((response: SPHttpClientResponse) => {
          if(response.status===201){
            alert('List Created Successfully');
          }else{
            alert('List Creation Failed'+response.status + "-" + response.statusText);
          }
        });
      }
      else{
        alert('List Creation Failed'+response.status + "-" + response.statusText);
      }
    });
  }

  // protected get dataVersion(): Version {
  //   return Version.parse('1.0');
  // }

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
