import { Version } from '@microsoft/sp-core-library';
import { SPComponentLoader } from '@microsoft/sp-loader';

import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SpfxTestAjaxWebPart.module.scss';
import * as strings from 'SpfxTestAjaxWebPartStrings';

// import node module external libraries
import * as $ from 'jquery';
import * as React from 'react';
require('bootstrap');

export interface ISpfxTestAjaxWebPartProps {
  description: string;
}

export default class SpfxTestAjaxWebPart extends BaseClientSideWebPart<ISpfxTestAjaxWebPartProps> {

  public render(): void {
    let teamName = "";
    let folderName = "";
    let subfolderName = "";
    let siteTitle = this.context.pageContext.web.title;
    var folderNamePrev = "";
    var subFolderNamePrev = "";
    var sfCount = 1;
    var fCount = 1;
    var folderID = "";
    var folderDoc = "";
    var subFolderGroupID = "";
    var subFolderID = "";
    var subFolderDoc = "";
    var folderString = "";
    var subFolderString = "";  
    var docFlag = false;
    let tabNum = 1;

    let bootstrapCssURL = "https://cdn.jsdelivr.net/npm/bootstrap@3.3.7/dist/css/bootstrap.min.css";
    let fontawesomeCssURL = "https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.11.2/css/regular.min.css";
    SPComponentLoader.loadCss(bootstrapCssURL);
    SPComponentLoader.loadCss(fontawesomeCssURL);

    let query = "<Query>"+
                    "<Where>"+                   
                        "<Or>"+                           
                            "<Contains><FieldRef Name='Knowledge_Team'/><Value Type='Text'>'Process Design'</Value></Contains>" +
                            "<Or>" +
                                "<Contains><FieldRef Name='Knowledge_SharedWith'/><Value Type='Text'>'Process Design'</Value></Contains>" + 
                                //"<Contains><FieldRef Name='Knowledge_SharedWith'/><Value Type='Text'>"+tabName+"</Value></Contains>"+
                            "</Or>" +
                        "</Or>"+
                    "</Where>"+
                    "<OrderBy>" +
                        "<FieldRef Name='Knowledge_Folder' Ascending='True' />"+
                        "<FieldRef Name='Knowledge_SubFolder' Ascending='True' />"+
                    "</OrderBy>"+
                "</Query>";
    let data = { 'query' :{'__metadata': { 'type': 'SP.CamlQuery' }, 'ViewXml': query}};

    $.ajax ({
      url:this.context.pageContext.web.absoluteUrl+"/_api/web/lists/GetByTitle('Policies')/items?$select=*,ID,TaxCatchAll/ID,TaxCatchAll/Term&$expand=TaxCatchAll", //?$filter=EntityPropertyName eq 'Medicals'",
      type:"GET",
      data: JSON.stringify(data), 
      headers:{"accept": "application/json;odata=verbose"},
      success: function(data) {
        console.log(data.d.results);
        let results=data.d.results;
        //_api/web/lists/getbytitle('TaxonomyHiddenList')/items?$filter=ID eq $select=ID,Title
        $.each(results,function(index,item){
          let teamID = item.Knowledge_Team.WssId;
          let folderID = item.Knowledge_Folder.WssId;
          if(item.Knowledge_Subfolder !== null){
            alert("data in subfolder");
            //let subfolderID = item.Knowledge_Subfolder.WssId;
          }
          let tax_len=item.TaxCatchAll.results.length;
/*
          for (var field in item) {
            // If it's the field we're interested in....
            if (item.hasOwnProperty(field) && field === ) {
                if (obj[field] !== null) {
                    // ... get the WssId from the field ...
                    var thisId = obj[field].WssId;
                    // ... and loop through the TaxCatchAll data to find the matching Term
                    for (var i = 0; i < obj.TaxCatchAll.length; i++) {
                        if (obj.TaxCatchAll[i].ID === thisId) {
                            // Augment the fieldName object with the Term value
                            obj[field].Term = obj.TaxCatchAll[i].Term;
                            return obj[field];
                        }
                    }
                }
            }
          }     
*/
          for (var i = 0; i < tax_len; i++) {
            //console.log("taxcatchall ID="+item.TaxCatchAll.results[i].ID);
            switch(item.TaxCatchAll.results[i].ID){
              case teamID:
                teamName = item.TaxCatchAll.results[i].Term;
                break;
              case folderID:
                folderName = item.TaxCatchAll.results[i].Term;
                break;
              //case subfolderID:
              //  subfolderName = item.TaxCatchAll.results[i].Term;
              //  break;
            }
          }
          console.log("TeamName="+teamName+" Foldername="+folderName);

          if (folderName !== folderNamePrev) {
            // ***** Setup Parent Folder 
            let folderTxt =  'knTab' + tabNum + '-Folder' + fCount;
            //console.log("powerUser="+PowerUser);
            folderDoc = folderTxt + "Doc";
            folderString = '<div class="card documentFolder">' +
                '<a class="card-link accordion-toggle" data-toggle="collapse" data-parent="#knLibrary' + tabNum + '" href="#' + folderTxt + '" style="text-decoration:none">' +
                '<div class="card-header">' +
                //'<div class="badge badge-info" id="'+folderID+'Count"></div>' +
                '<h5 class="folderTitle">' + folderName + '</h5>' +
                '</div>' +
                '</a>' +
                '<div id="' + folderTxt + '" class="collapse docList">' +
                '<div class="card-body" id="' + folderTxt + 'rootPanel">' +
                '<div id="' + folderDoc + '">' +
                '</div>' +
                '</div>' +
                '</div>' +
                '</div>';                           

            $('#folders').append(folderString);                                                           
            fCount++;
            folderNamePrev = folderName;
          }
        });       
      },
      error: function(Error){
        alert(JSON.stringify(Error));
      }
    }); 
        
    this.domElement.innerHTML = `
      <div class="${ styles.spfxTestAjax }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
              <p class="${ styles.description }">${escape(this.properties.description)}</p>
              <div class="newLoader"><div id="docLoader"></div></div>

              <!-- ***** Left Container - Navigation Tabs ***** -->
              <div id="knDocs">        
                  <div class="tabHeader w-100 p-1">   
                      <div class="scroller scroller-left mt-2"><i class="fa fa-chevron-left"></i></div>
                      <div class="scroller scroller-right mt-2"><i class="fa fa-chevron-right"></i></div>
                      <div class="customTabContainer"> 
                          <nav class="nav nav-pills list mt-2" role="tablist" id="TabNames"></nav>
                      </div>
                  </div>
        
                  <div class="tab-content documentSection" id="TabContent"> 
                    <div id="folders"></div>        
                  </div>  
                          
                  <div class="documentViewer">            
                      <iframe id="docViewer" src=""></iframe>
                  </div>                            
              </div>              
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
