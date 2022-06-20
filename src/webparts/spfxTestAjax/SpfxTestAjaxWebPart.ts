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
    let folderWssID: string[];
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
    let count = 0;
    let bootstrapCssURL = "https://cdn.jsdelivr.net/npm/bootstrap@3.3.7/dist/css/bootstrap.min.css";
    let fontawesomeCssURL = "https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.11.2/css/regular.min.css";
    SPComponentLoader.loadCss(bootstrapCssURL);
    SPComponentLoader.loadCss(fontawesomeCssURL);

    var list = "Policies"; 
    var webUrl = this.context.pageContext.web.absoluteUrl;
    var requestUri = "/_api/web/lists/GetByTitle('"+list+"')/items?$select=*,Folder,TaxCatchAll/ID,TaxCatchAll/Term,FieldValuesAsText/FileRef&$expand=TaxCatchAll,FieldValuesAsText,Folder";
    var _requestHeader = { "accept" : "application/json;odata=verbose" };
    var _contentType = "application/json;odata=verbose";
    var query = "<View><Query><OrderBy><FieldRef Name='Knowledge_Folder' Ascending='TRUE'/></OrderBy></Query></View>";
    var queryData = { "query" : {"__metadata": { "type": "SP.CamlQuery" }, "ViewXml":query}};
    //this.context.pageContext.web.absoluteUrl+"/_api/web/lists/GetByTitle('"+list+"')/items?$select=*,TaxCatchAll/ID,TaxCatchAll/Term,FieldValuesAsText/FileRef&$expand=TaxCatchAll,FieldValuesAsText",
    //_api/web/lists/getbytitle('TaxonomyHiddenList')/items?$filter=ID eq $select=ID,Title

    $.ajax ({
      url: webUrl+requestUri,
      type:"GET",
      headers : _requestHeader,
      data: queryData,
      success: function(data) {
        console.log(data.d.results);
        let results=data.d.results;

        //let sorted = [];
        //$(results).each(function(k, v) {
        //    for(var key in v) {
        //        sorted.push({key: key, value: v[key]})
        //    }
        //});
        //sorted.sort();
        //sorted = results.sort("Knowledge_Folder");
        //console.log(sorted);
        
        $.each(results, function (index, item) {

          let teamWssID = item.Knowledge_Team.WssId;
          this.folderWssID=item.Knowledge_Folder.WssId;
          
          if(item.Knowledge_SubFolder !== null){
            //alert("data in subfolder");
            //let subfolderWssID = item.Knowledge_SubFolder.WssId;
          }

          let tax_len=item.TaxCatchAll.results.length;
          console.log("folder id="+this.folderWssID);

/*          
          for (var i = 0; i < tax_len; i++) {
            //console.log("taxcatchall ID="+item.TaxCatchAll.results[i].ID);
            switch(item.TaxCatchAll.results[i].ID){
              case teamWssID:
                teamName = item.TaxCatchAll.results[i].Term;
                break;
              case folderWssID:
                folderName = item.TaxCatchAll.results[i].Term;
                break;
              //case subfolderWssID:
              //  subfolderName = item.TaxCatchAll.results[i].Term;
              //  break;
            }
          }
          console.log("TeamName="+teamName+" Foldername="+folderName);
          
          if (folderName !== folderNamePrev) {
            // ***** Setup Parent Folder 
            let folderTxt =  'dcTab' + tabNum + '-Folder' + fCount;
            //console.log("powerUser="+PowerUser);
            folderDoc = folderTxt + "Doc";
            folderString = '<div class="panel panel-default">' +
                              '<div class="panel-heading">' +
                                '<h4 class="panel-title">' +
                                  '<a data-toggle="collapse" data-parent="#accordion" href="#'+ folderTxt +'">'+
                                    '<span class="glyphicon glyphicon-menu-right text-success"></span> '+folderName+            
                                  '</a>'+
                                '</h4>'+
                              '</div>'+
                            '<div id="'+ folderTxt +'" class="panel-collapse collapse">'+
                              '<div class="panel-body">Document</div>'+
                            '</div>'+
                          '</div>';                           
            $('#folders').append(folderString);                                                           
            fCount++;
            folderNamePrev = folderName;
          }
*/
          //count++;
        });
        let folderLen = this.folderWssID.length;  
        console.log('folder length='+folderLen);     
      },
      error: function(Error){
        alert(JSON.stringify(Error));
      }
    }); 
    
    //folderWssID.sort();

    //for (var i = 0; i < folderWssID.length; i++) {
    //  $('#folders').append(folderWssID[i]); 
    //}

    this.domElement.innerHTML = `
      <div class="${ styles.spfxTestAjax }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
              <p class="${ styles.description }">${escape(this.properties.description)}</p>
              
              <div class="panel-group" id="accordion">
                <div class="panel panel-default">
                  <div class="panel-heading">
                    <h4 class="panel-title">
                      <a data-toggle="collapse" data-parent="#accordion" href="#collapseTwo">
                        <span class="glyphicon glyphicon-menu-right text-info"></span> Pay by PayPal            
                      </a>
                    </h4>
                  </div>
                  <div id="collapseTwo" class="panel-collapse collapse">
                    <div class="panel-body">Pay Pal</div>
                  </div>
                </div> 
                
                <div id='folders'/>                
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
