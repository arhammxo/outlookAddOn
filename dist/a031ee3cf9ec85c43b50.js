import"./taskpane.css";!function(){"use strict";var e,t;function i(e){$("#error-display").hide(),$("#not-configured").hide(),$("#gist-list-container").show(),getUserGists(e,(function(e,t){t||($("#gist-list").empty(),buildGistList($("#gist-list"),e,n))}))}function n(){$("#insert-button").removeAttr("disabled"),$(".ms-ListItem").removeClass("is-selected").removeAttr("checked"),$(this).children(".ms-ListItem").addClass("is-selected").attr("checked","checked")}function s(n){e=JSON.parse(n.message),setConfig(e,(function(n){t.close(),t=null,i(e.gitHubUserName)}))}function o(e){t=null}Office.initialize=function(n){jQuery(document).ready((function(){(e=getConfig())&&e.gitHubUserName?i(e.gitHubUserName):$("#not-configured").show(),$("#insert-button").on("click",(function(){var e=$(".ms-ListItem.is-selected").val();getGist(e,(function(e,t){Office.context.mailbox.item.body.setSelectedDataAsync(e,{coercionType:Office.CoercionType.Html},(function(e){e.status===Office.AsyncResultStatus.Failed&&function(e){$("#not-configured").hide(),$("#gist-list-container").hide(),$("#error-display").text(e),$("#error-display").show()}("Could not insert gist: "+e.error.message)}))}))})),$("#settings-icon").on("click",(function(){var i=new URI("dialog.html").absoluteTo(window.location).toString();e&&(i=i+"?gitHubUserName="+e.gitHubUserName+"&defaultGistId="+e.defaultGistId),Office.context.ui.displayDialogAsync(i,{width:20,height:40,displayInIframe:!0},(function(e){(t=e.value).addEventHandler(Office.EventType.DialogMessageReceived,s),t.addEventHandler(Office.EventType.DialogEventReceived,o)}))}))}))}}();