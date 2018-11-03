import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { Moment } from 'moment';
import * as strings from 'RegionalNewsWebPartStrings';
import 'jquery';
export interface IRegionalNewsWebPartProps {
  description: string;
}
let moment = require('moment');
export default class RegionalNewsWebPart extends BaseClientSideWebPart<IRegionalNewsWebPartProps> {

  public render(): void {
    let cssURL = "https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css";
    let fontUrl = "https://stackpath.bootstrapcdn.com/font-awesome/4.7.0/css/font-awesome.min.css";
    SPComponentLoader.loadCss(cssURL);
    SPComponentLoader.loadCss(fontUrl);
    "https://www.w3schools.com/w3css/4/w3.css"
    this.domElement.innerHTML = `
    
      <div class = "panel panel-default">
        <div class = "panel-heading col-md-12" style="background-color: #023576; color: #ccd6e4;height: 50px;">
            <div style="float: left; font-size: large;">Regional News </div> <div style="float: right;"> <button class="btn btn-warning btn-xs" type="button"style=" height: 31px;"><i class="fa fa-newspaper-o fa-lg" aria-hidden="true"></i></button></div>
        </div>
        
        <div class = "panel-body">
          <ul style="list-style-type:none;" id="RegionalNewsList">
          <!--  <li style="border-top: 0; border-left: 0; border-right: 0; margin-top: 10%;"><a href="#">&raquo;&nbsp; Cras justo odio</a><p style="font-size: x-small;">March 28, 2018</p></li>
            <hr>
            <li style="border-top: 0; border-left: 0; border-right: 0; "><a href="#">&raquo;&nbsp; Dapibus ac facilisis in</a><p style="font-size: x-small;">April 15, 2018</p></li>
            <hr>
            <li style="border-top: 0; border-left: 0; border-right: 0;"><a href="#">&raquo;&nbsp; Morbi leo risus</a><p style="font-size: x-small;">April 20, 2018</p></li>
            <hr>
            <li style="border-top: 0; border-left: 0; border-right: 0;"><a href="#">&raquo;&nbsp; Porta ac consectetur ac</a><p style="font-size: x-small;">April 25, 2018</p></li>
            <hr>
            <li style="border-top: 0; border-left: 0; border-right: 0;"><a href="#">&raquo;&nbsp; Vestibulum at eros</a><p style="font-size: x-small;">May 15, 2018</p></li> -->
          </ul>
        </div>
      </div>`;

      this.readyFunc();
  }

  readyFunc()
  {
    var Absourl = this.context.pageContext.web.absoluteUrl;   
    -
    $(document).ready(function(){
      DisplayRegionalNews();
      function DisplayRegionalNews(){
        var call = jQuery.ajax({
          url: Absourl + "/_api/Web/Lists/getByTitle('SpfxRegionalNews')/Items?$top=3&$orderby=Created desc",
          type: "GET",
          dataType: "json",
          async: false,  
          headers: {
              Accept: "application/json; odata=verbose",
              "Content-Type": "application/json;odata=verbose"
          }
      });
        call.done(function (data, textStatus, jqXHR) 
        {          
            $('#RegionalNewsList li').remove();
            var orderedList = $('#RegionalNewsList');
            var newsSize: any = data.d.results.length-1;
            $.each(data.d.results, function (idx, elem) 
            {
              var objDate = new Date(elem.Created),
              locale = "en-us",
              month = objDate.toLocaleString(locale, { month: "long" });
              var newsDescription = elem.NewsDescription.length > 68 ? elem.NewsDescription.substr(0,68)+"..." : elem.NewsDescription;
              if(idx == newsSize)
              {
                orderedList.append("<li style='margin-left:-30px;'><a href='https://acuvateuk.sharepoint.com/sites/TrainingDevSite/Lists/SpfxRegionalNews/AllItems.aspx' target='_blank' style='font-family: sans-serif; color: #345c93;' data-toggle='tooltip' title='"+elem.NewsDescription+"'>&raquo;&nbsp; "+newsDescription+"</a><p style='font-size: x-small;color: #949494;margin-top: 1%;'>"+ month+" "+objDate.getDate()+", "+objDate.getFullYear()+"</p></li>");   
              }
              else if(idx == "0")
              {
                orderedList.append("<li style='margin-left:-30px;margin-top: 40px;height: 80px;'><a href='https://acuvateuk.sharepoint.com/sites/TrainingDevSite/Lists/SpfxRegionalNews/AllItems.aspx' target='_blank' style='font-family: sans-serif;color: #345c93;' data-toggle='tooltip' title='"+elem.NewsDescription+"'>&raquo;&nbsp; "+newsDescription+"</a><p style='font-size: x-small;color: #949494;margin-top: 1%;'>"+ month+" "+objDate.getDate()+", "+objDate.getFullYear()+"</p><hr></li>");
              }
              else
              {
                orderedList.append("<li style='margin-left:-30px';height: 80px;><a href='https://acuvateuk.sharepoint.com/sites/TrainingDevSite/Lists/SpfxRegionalNews/AllItems.aspx' target='_blank' style='font-family: sans-serif;color: #345c93;' data-toggle='tooltip' title='"+elem.NewsDescription+"'>&raquo;&nbsp; "+newsDescription+"</a><p style='font-size: x-small;color: #949494;margin-top: 1%;'>"+ month+" "+objDate.getDate()+", "+objDate.getFullYear()+"</p><hr></li>");
              }            
            });
        });
        call.fail(function (jqXHR, textStatus, errorThrown) {
            var response = JSON.parse(jqXHR.responseText);
            var message = response ? response.error.message.value : textStatus;
            alert("Call failed. Error: " + message);
        });
      }
    });
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
