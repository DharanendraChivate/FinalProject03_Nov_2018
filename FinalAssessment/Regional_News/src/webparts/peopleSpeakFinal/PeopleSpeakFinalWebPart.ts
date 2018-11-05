import { Version,Environment, EnvironmentType } from '@microsoft/sp-core-library';
import {BaseClientSideWebPart,IPropertyPaneConfiguration,PropertyPaneTextField
       } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import{SPComponentLoader} from '@microsoft/sp-loader';//to load all the css cdn paths which were used
require('jquery');
import * as $ from 'jquery';//importing jquery
require("bootstrap");
import styles from './PeopleSpeakFinalWebPart.module.scss';
import * as strings from 'PeopleSpeakFinalWebPartStrings';
export interface IPeopleSpeakFinalWebPartProps {
description: string;
}
export default class PeopleSpeakFinalWebPart extends BaseClientSideWebPart<IPeopleSpeakFinalWebPartProps> {

    public render(): void {
    //all The Css Cdn's Which Were Used In My WebPart
    let Bootstrapurl="https://stackpath.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css";
    let w3style="https://www.w3schools.com/w3css/4/w3.css";
    let glyphicon='https://use.fontawesome.com/releases/v5.4.1/css/all.css';
    //Loading all the Css Cdn's Via SpComponenetLoader
    SPComponentLoader.loadCss(Bootstrapurl);
    SPComponentLoader.loadCss(w3style);
    SPComponentLoader.loadCss(glyphicon);
    this.domElement.innerHTML = `  

    <!-- Inserting Panel  -->
    <div class="panel panel-default">
    <div class="panel-heading col-md-12"  style="background-color: #023576; color: #ccd6e4;">
    <!-- <div class="panel-heading col-md-12" style='${styles.navheader}'>-->

    <div style="float: left; font-size:large;">People Speak</div> 
    <div style="float: right;"> 
    <button class="btn btn-warning btn-circle btn-xs" id="roundbutton">
    <i class='fas fa-microphone' style='font-size:20px;color:white;margin-top: 6px;'></i>
    </button>
    </div>   
    </div>

    <!-- Inserting Carousel in the Panel -->

    <div id="PeopleSpeakCarousel" class="carousel slide" data-ride="carousel">
    

    <!-- Wrapper for slides -->
    <div class="carousel-inner" id="DisplaySlides">

    </div>
    

    <!-- Left and right controls -->
    <a class="left carousel-control" href="#PeopleSpeakCarousel" data-slide="prev" style="margin-top:50%;height: max-content;">
    <button class="btn btn-sq btn-warning btn-xs" style="margin-right: 45%;">
    <i class='fas fa-angle-left fa-2x' style=''></i>
    </button>
    </a>
     <a class="right carousel-control" href="#PeopleSpeakCarousel" data-slide="next" style="margin-top:50%; height: max-content;">
     <button class="btn btn-sq btn-warning btn-xs" style="margin-left: 45%; btn-xs">
     <i class='fas fa-angle-right fa-2x' style=''></i>
     </button>
     </a> 
     </div>
     </div>  

    
     <button type="button" id="NavigateToList" class="btn btn-warning btn-sm center-block " style="color:#0000ff; position: absolute; top:92%; right:39%;">View All</button>
      `;

     this.DisplayData();

    }

   private DisplayData(){
    try{

    var absoluteUrl = this.context.pageContext.web.absoluteUrl;
    let html: string = '';
    if (Environment.type === EnvironmentType.Local) {
      this.domElement.querySelector('#test').innerHTML = "sorry this does not work in local workbench";
    }
    else{
      //ajaxcall Inorder TO Fetch The Data From The Sharepoint List And To Display
    var call = $.ajax({
    //REstapi To Fetch All The details From The Sharepoint ListWhich Are Required Inorder To Display.
    url: absoluteUrl + "/_api/web/lists/getByTitle('SpfxPeopleSpeak')/Items/?$select= Id,Title,Designation,Picture&$top=3&$orderby=Id desc",
    type: "GET",
    dataType: "json",
    headers: {
    Accept: "application/json;odata=verbose"
        }
      });

    call.done(function (data,textStatus,jqXHR) {
    var Slider = $("#DisplaySlides");
    var Active;
    $.each(data.d.results, function (index,value) {
     //The First Slide Of The Carousel Should Be in The ActiveClass
    if(index=='0')
    { Active="item active"
    }
    else{
    Active="item"
    };
    // binding data to wrapper for slides 
    if(value.Picture==null&&value.Designation!=null){
      Slider.append("<div class='"+Active+"'><img src='https://upmaa-pennmuseum.netdna-ssl.com/collections/images/image_not_available_300.jpg' style='width:100%;height:230px;opacity: 1;'><div class='carousel-caption' style='position: absolute;bottom: 8px;left: -40px; padding-bottom: 0px;'><h6><p style='background-color:#080808; color:white;opacity:0.6;'>"+value.Title+"<br/><i>"+value.Designation+"</i></p></h6></div>  </div>");
    }
    else if(value.Designation==null&&value.Picture!=null){
      Slider.append("<div class='"+Active+"'><img src='"+value.Picture.Description+"' style='width:100%;height:230px;opacity: 1;'><div class='carousel-caption' style='position: absolute;bottom: 8px;left: -40px; padding-bottom: 0px;'><h6><p style='background-color:#080808; color:white;opacity:0.6;'>"+value.Title+"<br/><i>Designation Not Found</i></p></h6></div>  </div>");
    }
    else if(value.Picture==null &&value.Designation==null){
      Slider.append("<div class='"+Active+"'><img src='https://upmaa-pennmuseum.netdna-ssl.com/collections/images/image_not_available_300.jpg' style='width:100%;height:230px;opacity: 1;'><div class='carousel-caption' style='position: absolute;bottom: 8px;left: -40px; padding-bottom: 0px;'><h6><p style='background-color:#080808; color:white;opacity:0.6;'>"+value.Title+"<br/><i>Designation Not Found</i></p></h6></div>  </div>");
    }
    else{
      Slider.append("<div class='"+Active+"'><img src='"+value.Picture.Description+"' style='width:100%;height:230px;opacity: 1;'><div class='carousel-caption' style='position: absolute;bottom: 8px;left: -40px; padding-bottom: 0px;'><h6><p style='background-color:#080808; color:white;opacity:0.6;'>"+value.Title+"<br/><i>"+value.Designation+"</i></p></h6></div>  </div>");
    }
    }); 
    });
    

    call.fail(function (jqXHR, textStatus, errorThrown) {
    var response = JSON.parse(jqXHR.responseText);
    var message = response ? response.error.message.value : textStatus;
    alert("Call failed. Error: " + message);
    });

    // function to navigate to the list page when clicking on the view all function 
    $(document).on("click","#NavigateToList",function(){
      //$(this).attr('target','_blank');
    window.open("https://acuvateuk.sharepoint.com/sites/TrainingDevSite/Lists/SpfxPeopleSpeak/AllItems.aspx",'_blank');
    })
    }
    }
     catch(error)
     {
      console.log(error);

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
