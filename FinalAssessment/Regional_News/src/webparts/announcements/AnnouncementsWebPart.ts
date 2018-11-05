import { Version, Environment, EnvironmentType } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape, times } from '@microsoft/sp-lodash-subset';

import styles from './AnnouncementsWebPart.module.scss';
import * as strings from 'AnnouncementsWebPartStrings';
import { SPComponentLoader } from '@microsoft/sp-loader';
import 'jquery';
import { SPHttpClientResponse, SPHttpClient } from '@microsoft/sp-http';
require('bootstrap');

export interface IAnnouncementsWebPartProps {
  description: string;
}

export default class AnnouncementsWebPart extends BaseClientSideWebPart<IAnnouncementsWebPartProps> {

  public render(): void {
    let Url = "https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css";
    //var AbsoluteURL = this.context.pageContext.web.absoluteUrl;
    SPComponentLoader.loadCss(Url);
    this.domElement.innerHTML = `
            
    <div id="Error"></div>
    <div id="AnnouncementsDisplay" class="carousel slide" data-ride="carousel">
    
      <!-- Wrapper for slides -->
      <div class="carousel-inner">

      </div>
    </div>`;

    $(document).ready(function () {
    });
    this.getCarousal();
  }
  getCarousal() {
    //alert("entered get carousal event");
    let html: string = '';
    if (Environment.type === EnvironmentType.Local) {
      this.domElement.querySelector('#Error').innerHTML = "Sorry this does not work in local workbench";
    } else {
      this.context.spHttpClient.get(
        this.context.pageContext.web.absoluteUrl + `/_api/Web/Lists/getByTitle('Announcements')/items?$select=Title,ImageURL,ID,Description&$top=5&$orderby=Created desc`, SPHttpClient.configurations.v1
      ).then((Response: SPHttpClientResponse) => {
        Response.json().then((listObjects: any) => {
          var ImageSliderCount = 0;
          listObjects.value.forEach(element => {
            if (ImageSliderCount === 0) {
              html += `<div class='item active'><img src="${element.ImageURL.Url}" alt="${element.Title}" style='width:100%;height:363px;'>
            <div style="background-color:black;opacity:0.5;position: absolute;top: 10%;left: 35%;transform: translate(-50%, -50%);
            font-size: 20px;color:#ffd633;">${element.Title}</div>
            <div style="background-color:black;opacity:0.5;position: absolute;top: 30%;left:35%;transform: translate(-50%, -50%);
            font-size:14px;color:white;display:block;">${element.Description}</div>
            </div>`;
              ImageSliderCount++;
            }
            else {
              html += `<div class="item"><img src="${element.ImageURL.Url}" alt="${element.Title}" style='width:100%;height:363px;'>
            <div style="background-color:black;opacity:0.5;position: absolute;top: 10%;left: 35%;transform: translate(-50%, -50%);
            font-size: 20px;color:#ffd633;">${element.Title}</div>
            <div style="background-color:black;opacity:0.5;position: absolute;top: 30%;left: 35%;transform: translate(-50%, -50%);
            font-size:14px;color:white;display:block;">${element.Description}</div>
            </div>`
                ;
              ImageSliderCount++;
            }
          });
          this.domElement.querySelector('.carousel-inner').innerHTML = html + `<div style="float:left;padding-right: 1%;">
          <a class="glyphicon glyphicon-chevron-left btn btn-warning" style="position: absolute; top: 85%; right: 90%;height:30px;" href="#AnnouncementsDisplay" data-slide="prev"></a>
          </div>
          
          <div>
          <a class="glyphicon glyphicon-chevron-right btn btn-warning" style="position: absolute; top: 85%; right: 80%;height:30px;" href="#AnnouncementsDisplay" data-slide="next"></a>
          </div>
          <div id="viewall">
            <a class="btn btn-warning" data-toggle="tooltip" data-placement="bottom" title="Click to see all Announcements" href="https://acuvateuk.sharepoint.com/sites/TrainingDevSite/Lists/Announcements/AllItems.aspx" target="_blank" style="position: absolute; top: 85%; left: 80%;height:30px;">View All</a>
          </div>`;
        });
      });
    }
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
