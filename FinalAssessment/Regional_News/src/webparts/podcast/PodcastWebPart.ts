import {
  Version,
  Log,
  ServiceScope,
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPComponentLoader } from '@microsoft/sp-loader';
import styles from './PodcastWebPart.module.scss';
import * as strings from 'PodcastWebPartStrings';
import * as $ from 'jquery';
import * as bs from 'bootstrap';
require('bootstrap');
import * as pnp from 'sp-pnp-js';
import { CurrentUser } from 'sp-pnp-js/lib/sharepoint/siteusers';
const LOG_SOURCE: string = 'PodcastWebPart';

export interface IPodcastWebPartProps {
  description: string;
}
var PodcastUser;
var PodcastId;
var CurrentUserId;
var IslikedBefore;
var LikesCount;
var AbsoluteUrl;

export default class PodcastWebPart extends BaseClientSideWebPart<IPodcastWebPartProps> {

  public render(): void {

    // this.context.statusRenderer.clearError(this.domElement);
    // this.context.statusRenderer.displayLoadingIndicator(this.domElement, strings.Loading);
    // Log.verbose('SpFxNuggets', 'Invoking render');

    // this._webInfoProvider.getWebInfo().then((webInfo: IWebInfo) => {
    //   if (this.properties.fail) {
    //     throw new Error('Mayday');
    //   }
    //   Log.info('SpFxNuggets', 'Service OK', this.context.serviceScope);
    //   this.context.statusRenderer.clearLoadingIndicator(this.domElement);
    //   this.context.domElement.innerHTML = `<h1>${webInfo.title}</h1>`;

    // }).catch((err) => {
    //   Log.error('SpFxNuggets', err);
    //   this.context.statusRenderer.clearLoadingIndicator(this.domElement);
    //   this.context.statusRenderer.renderError(this.domElement, err);
    // });

    AbsoluteUrl = this.context.pageContext.web.absoluteUrl;
    var contextuser = this.context.pageContext.user.email;

    let CssURL = "https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css";
    let FontUrl = "https://stackpath.bootstrapcdn.com/font-awesome/4.7.0/css/font-awesome.min.css";
    let glyph = "https://use.fontawesome.com/releases/v5.4.2/css/all.css";

    SPComponentLoader.loadCss(CssURL);
    SPComponentLoader.loadCss(FontUrl);
    SPComponentLoader.loadCss(glyph);

    this.domElement.innerHTML = `
    <div class="${ styles.podcast}">       <!-- Podcast -->
    <div class="${ styles.container}">   <!-- container -->
                <div class="${ styles.row}">        <!-- row -->
                <div><i class='fa fa-podcast' id="${styles.podcasticon}" style="padding-top: 3%; padding-left: 79.5%;"></i></div>
                <div><label id="${styles.podcastlabel}" style="margin-top: 11%;">Podcast</label></div>
                <br/>                 
                <div class="${ styles.column}">   <!-- column -->
                  <div style="text-align:center"  class="${styles.border}">    <!-- border-->
                      <div class="${styles.image}">
                      <img src="" class="img-responsive" id="image" alt="Cinque Terre" width="150" height="100"> 
                      </div>
                    <p class="${ styles.title}" id="Title" style="font-size: larger;font-weight: lighter;"></p>
                    <p class="${ styles.subTitle}" id="Role" style="font-size: small;font-weight: lighter;font-style: italic;"></p>
                    <p class="${ styles.description}" id="Description" style="font-size: smaller; font-weight: 100;font-style: italic;"></p>
                      
                      <i  class="${ styles.ThumbsUp} fa fa-thumbs-up fa-xs" style="color:#f2b914;" id="ThumbsUp"></i>
                      <i  class="${ styles.CommentsIcon} fa fa-comments fa-xs" id="CommentsIcon" style="color:#f2b914;"></i>
                    
                      <span>                    
                             <button  class="${ styles.hyperlinks} btn btn-link" Id="ReadMore">Read More</button> 
                             <a class="${ styles.hyperlinks}" target="_blank" href="https://acuvateuk.sharepoint.com/sites/TrainingDevSite/Lists/Podcast/AllItems.aspx?viewpath=%2Fsites%2FTrainingDevSite%2FLists%2FPodcast%2FAllItems.aspx">View All</a>
                          </span>
                    </div>                            <!-- border-->
                  </div>                              <!-- column -->
                </div>                                <!-- row -->
      
      
      
    
        <!-- The Modal -->
     <div class="modal fade" id="ReadMorePodCast">
          <div class="modal-dialog modal-lg">
              <div class="modal-content" id="Modalcontent">           
                      <!-- Modal Header -->
                      <div class="${styles["modal-header"]} modal-header" id="Modalheader">
                      <h4 class="modal-title"> </h4>
                      <button type="button" class="close" data-dismiss="modal">&times;</button>
                      </div>
                     <!-- Modal body -->
                     <div class="${styles["modal-body"]} modal-body" id="Modalbody">
                        <div class="container-fluid">
                            <div class="row">
                                <div class="col-md-3">
                                    <img src="" id="popupimage" class="img-responsive" alt="Cinque Terre" > 
                                    <p id ="popuprole" class="bg-success"></p>
                                </div>
                               <div class="col-md-3 bg-danger">
                                    <div class="row">
                                    <p class="bg-primary">Description.</p>
                                    </div>
                                    <div class="row">
                                    <div class="col" id= "${styles.scrollDescription}" >
                                    <p id = "popupdescription"></p>
                                    </div>
                                    </div>
                                </div>
                                <div class="col-md-6 ml-auto col bg-success">  
                                      <div class="row">
                                      <p class="bg-info">comments.</p>
                                      </div>  
                                        <div class="row">        
                                              <div class="col" id="${styles.scrollComments}">                       
                                                    <div class="${styles.popupcomments}" id = "popupcomments">
                                                      
                                                          <section class="${styles["comment-list"]} comment-list" id="commentlist">
                                                                <!-- dynamic comments -->
                                                          </section>
                                                      </div>   

                                                      <div class="widget-area no-padding blank">
                                                      <div class="status-upload">
                                                        <form>
                                                        <div class="row"> 
                                                        <div class="col-md-9">
                                                        <textarea id="CommentTextBox" style="width:100%;" placeholder="Post Your Comment" ></textarea>                                                  
                                                        </div> 
                                                        <div class="col">
                                                        <button type="button" id="SubmitComment" class="btn btn-success green" style="margin-top: 1.5%;"><i class="fa fa-share"></i> Share</button>
                                                        </div> 
                                                        </div> 
                                                        </form>
                                                    </div> <!-- Status Upload  -->
                                                   </div>   <!-- Widget Area -->
                                                </div>   
                                             
                                          </div>     
                                              
                                 </div>
                            </div>
                         </div>
                    </div>            
                     <!-- Modal footer -->
                     <div class=" ${styles["modal-footer"]} modal-footer">
                       <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
                     </div>             
                </div>
          </div>
      </div>
    <!-- Modal ends-->
        </div> <!-- container -->
      </div>   <!-- Podcast -->`;

    this.DisplayPodcast();
  }

  //---------------------------method to display podacst-----------------------------//
  DisplayPodcast() {
    if (Environment.type === EnvironmentType.Local) {
      this.domElement.querySelector('#Error').innerHTML = "Sorry this does not work in local workbench";
    }
    else {

      $(document).ready(function () {

        $(document).on('click', '#ReadMore', function () {
          SPFXPodcastPopupComment();

          $('#ReadMorePodCast').modal('show');
        });
        GetUserDetails();
        SPFXPodcast();
        SPFXPodcastLikesCount();
        SPFXPodcastCommentsCount();
        OnClickOfLike();
        SubmitComment();
      });
    }

    //-------------------------------------function to get the current user id-----------------------------//
    function GetUserDetails() {
      var url = AbsoluteUrl + "/_api/web/currentuser";
      $.ajax({
        url: url,
        headers: {
          Accept: "application/json;odata=verbose"
        },
        async: false,
        success: function (data) {
          CurrentUserId = data.d.Id;
        },
        error: function (data) {
          alert("An error occurred. Please try again.");
        }
      });
    }

    //-------------------------------------function to display the main part-----------------------------//
    function SPFXPodcast() {

      var call = $.ajax({
        url: AbsoluteUrl + `/_api/web/lists/GetByTitle('SPFXPodcast')/Items?$select = Title,Role,ImageURL,Description,LikesCount&$top = 1&$orderby=Created desc`,
        type: 'GET',
        async: false,
        dataType: "json",
        headers: {
          Accept: "application/json;odata=verbose"
        }
      });

      call.done(function (data, textStatus, jqXHR) {

        PodcastUser = data.d.results[0].Title;
        PodcastId = data.d.results[0].Id;
        if (data.d.results[0].URL != null) {
          $('#image').attr("src", data.d.results[0].URL.Url);
          $('#popupimage').attr("src", data.d.results[0].URL.Url);
        }
        else {
          $('#image').attr("src", "http://www.bsmc.net.au/wp-content/uploads/No-image-available.jpg");
          $('#popupimage').attr("src", "http://www.bsmc.net.au/wp-content/uploads/No-image-available.jpg");
        }
        $('#Title').text(data.d.results[0].Title);
        if (data.d.results[0].Role != null) {
          $('#Role').text(data.d.results[0].Role);
          $('#popuprole').text(data.d.results[0].Role);
        }
        else {
          $('#Role').text("Role is not available");
          $('#popuprole').text("Role is not available");
        }
        if (data.d.results[0].Description != null) {
          $('#Description').text((data.d.results[0].Description).substr(0, 50) + "...");
          $('#popupdescription').text(data.d.results[0].Description);
        }
        else {
          $('#Description').text("there is no Description for this person");
          $('#popupdescription').text(" there is no Description for this person available in the list ");
        }
        //assigning the data to the popup
        $('.modal-title').text(data.d.results[0].Title);
      });

      call.fail(function (jqXHR, textStatus, errorThrown) {
        var response = JSON.parse(jqXHR.responseText);
        var message = response ? response.error.message.value : textStatus;
        alert("Call failed. Error: " + message);
      });
    };

    //--------function to display the number of likes----------//
    function SPFXPodcastLikesCount() {

      var call = $.ajax({
        url: AbsoluteUrl + `/_api/web/lists/GetByTitle('SPFXPodcastLikes')/Items?$expand=Author,UserLookup&$select=Author/Id,Author/Title,UserLookup/Title&$filter=UserLookup eq ${PodcastId}`,
        type: 'GET',
        dataType: "json",
        async: false,
        headers: {
          Accept: "application/json;odata=verbose"
        }
      });
      call.done(function (data, textStatus, jqXHR) {
        LikesCount = data.d.results.filter(value => value.UserLookup.Title === PodcastUser).length;  //likescount
        $('#ThumbsUp').text(LikesCount);
        $('#popuprole').append("</br>");
        $('#popuprole').append(`<i class="fa fa-thumbs-up"> ${LikesCount} </i>`);
        IslikedBefore = data.d.results.filter(value => value.Author.Id === CurrentUserId).length;  //checking if the user liked the person before
      });
      call.fail(function (jqXHR, textStatus, errorThrown) {
        var response = JSON.parse(jqXHR.responseText);
        var message = response ? response.error.message.value : textStatus;
        alert("Call failed. Error: " + message);
      });
    };

    //---------------------------function to display the number of comments-----------------------------//
    function SPFXPodcastCommentsCount() {

      var call = $.ajax({
        url: AbsoluteUrl + `/_api/web/lists/GetByTitle('SPFXPodcastComments')/Items?$expand=Author,UserLookup&$select=Author/Id,Author/Title,UserLookup/Title`,
        type: 'GET',
        dataType: "json",
        async: false,
        headers: {
          Accept: "application/json;odata=verbose"
        }
      });
      call.done(function (data, textStatus, jqXHR) {
        var counter = data.d.results.filter(value => value.UserLookup.Title === PodcastUser).length; //commentscount
        $('#CommentsIcon').text(counter + "comment(s)");
      });
      call.fail(function (jqXHR, textStatus, errorThrown) {
        var response = JSON.parse(jqXHR.responseText);
        var message = response ? response.error.message.value : textStatus;
        alert("Call failed. Error: " + message);
      });
    };

    //------------------------------function to like the person----------------------------------//
    function OnClickOfLike() {
      $(document).on("click", "#ThumbsUp", function () {
        if (IslikedBefore <= 0) {              //checking if the user already liked the person
          pnp.sp.web.lists.getByTitle('SPFXPodcastLikes').items.add({ UserLookupId: PodcastId })   //adding like
            .then(() => {
              SPFXPodcastLikesCount();        //refreshing likescount after liking
              $('#ThumbsUp').append("Liked");
            });
        }
        else {
          alert("you already liked the picture");
        }
      });
    }
    //--------------------------------function to display comments on the modal-----------//
    function SPFXPodcastPopupComment() {
      var call = $.ajax({
        url: AbsoluteUrl + `/_api/web/lists/GetByTitle('SPFXPodcastComments')/Items?$expand=Author,UserLookup&$select=Created,Comment,Author/Id,Author/Title,UserLookup/Title&$filter=UserLookup eq ${PodcastId}`,
        type: 'GET',
        dataType: "json",
        headers: {
          Accept: "application/json;odata=verbose"
        }
      });
      //---------------dynamically adding comments--------------------//
      call.done(function (data, textStatus, jqXHR) {

        $('#commentlist').empty();
        $.each(data.d.results, function (index, value) {
          $('#commentlist').append(`       
                  <article class="row">                
                    <div class="col-md-10 col-sm-10">
                      <div class="panel panel-default arrow left">  
                        <div class="${styles["panel-body"]}">
                          <header class="text-left">
                            <div class="comment-user"><i class="fa fa-user"></i> ${value.Author.Title}</div>
                            <time class="comment-date" datetime="16-12-2014 01:05"><i class="fa fa-clock-o"></i> ${value.Created}</time>
                          </header>
                          <div class="comment-post">
                            <p>
                            ${value.Comment}
                             </p>
                           </div>                
                        </div>
                      </div>
                    </div>
                  </article>`);
        });
      })
      call.fail(function (jqXHR, textStatus, errorThrown) {
        var response = JSON.parse(jqXHR.responseText);
        var message = response ? response.error.message.value : textStatus;
        alert("Call failed. Error: " + message);
      });
    }
    //-------------------------------function to submit the comment in popup to the list----------------------------------//
    function SubmitComment() {

      $(document).on("click", "#SubmitComment", function () {
        var Comment = $("#CommentTextBox").val();    // getting text in text box
        $("#CommentTextBox").val('');                // emptying the text box
        if (!Comment) {                             //checking if the value is null
          alert("Please enter the comment");
        }
        else {
          pnp.sp.web.lists.getByTitle('SPFXPodcastComments').items.add({ Comment: Comment, UserLookupId: PodcastId })
            .then(() => {
              SPFXPodcastPopupComment();            //refreshing comments section after entering comment
              SPFXPodcastCommentsCount();
            });
        }
      })
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
    }
  }
}


