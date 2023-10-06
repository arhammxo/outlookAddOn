let config;
let btnEvent;
import './fl.html';

// The initialize function must be run each time a new page is loaded.
Office.initialize = function() {};

function showError(error) {
    Office.context.mailbox.item.notificationMessages.replaceAsync('github-error', {
        type: 'errorMessage',
        message: error
    }, function(result) {});
}

let settingsDialog;

function insertDefaultGist(event) {
    // const u = 'https://api.github.com/users/arhammxo';
    const u = 'https://arhammxo.github.io/jsonDem/meetingspace.json';
    // const u = 'http://52.66.254.76:8891/meetingspace';
    try {
        $.ajax({
            url: u,
            dataType: 'json'
        }).done(function(gist) {
            const startDate = new Date();
            startDate.setDate(startDate.getDate() + 2);
            Office.context.mailbox.item.start.setAsync(
                startDate, { asyncContext: { optionalVariable1: 1, optionalVariable2: 2 } },
                (asyncResult) => {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        console.log(asyncResult.error.message);
                        return;
                    }

                    console.log("Successfully set the start time.");
                    /*
                        Run additional operations appropriate to your scenario and
                        use the optionalVariable1 and optionalVariable2 values as needed.
                    */
                });
            const endDate = new Date();
            endDate.setDate(endDate.getDate() + 4);
            Office.context.mailbox.item.end.setAsync(
                endDate, { asyncContext: { optionalVariable1: 1, optionalVariable2: 2 } },
                (asyncResult) => {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        console.log(asyncResult.error.message);
                        return;
                    }

                    console.log("Successfully set the start time.");
                    /*
                        Run additional operations appropriate to your scenario and
                        use the optionalVariable1 and optionalVariable2 values as needed.
                    */
                });
            let roomName = 'Bamboo Room';
            let cB = '<!doctype html>';
            cB = cB + '<html>';
            cB = cB + '<head>';
            cB = cB + '<meta charset="UTF-8">';
            cB = cB + '<meta name="viewport" content="width=device-width, initial-scale=1.0">';
            cB = cB + '<link href="https://fonts.googleapis.com/css?family=Open Sans" rel="stylesheet">';
            cB = cB + '</head>';
            cB = cB + '<body style="font-family: Open Sans;">';
            cB = cB + '<div style="position: relative; margin-left:20px; display: flex; flex-direction: row; border-radius: 25px; background-image: linear-gradient(to right top, #ffffff, #f5f5f5, #ebebeb, #e2e2e2, #d8d8d8); padding: 10px; box-shadow: 5px 5px 7px lightslategray;">';
            cB = cB + '<div style="z-index: 1;">';
            cB = cB + '<img style="border-radius: 50px 40px; max-width: 600px;" src="https://localhost:3000/assets/br.png" alt="">';
            cB = cB + '</div>';
            // cB = cB + '<div style="position: absolute; bottom: 5px; left: 20px; z-index: 2;">';
            // cB = cB + '<p style="padding: 10px; padding-left: 20px; padding-right: 20px; background-color: whitesmoke; border-top-right-radius: 30px 30px;"><span style="font-weight: bold;">Elhi tru</span><span>Rooms</span></p>';
            // cB = cB + '</div>';
            cB = cB + '<div style="margin: 5px; margin-left: 30px;">';
            cB = cB + '<p style="margin-left: 5px; color: #36454F; margin-bottom: 0; font-size: xx-large; font-weight: bolder;">' + roomName + '</p>';
            cB = cB + '<p style="margin-left: 5px; color: #36454F;  margin-top: 0; "><span style="font-weight: bolder;">First Floor, Business Plaza</span> <span>(Last room on the left)</span></p>';
            // cB = cB + '<
            cB = cB + '<div style="margin-top: 15px">';
            cB = cB + '<div style="display: flex; flex-direction: row; position: relative;  background-color: #ffffff;"';

            let tl = "window.location.href='https://w3docs.com';"
            cB = cB + '<button style="color: #000000; background-color: #ffffff; border: none;  border-radius: 50px; display: flex; flex-direction:row; align-items: center; cursor: pointer;"></button>';
            cB = cB + '<a href="" style="text-decoration: none;"><button style="margin: 5px; color: #000000; background-color: #ffffff; padding: 10px 10px; border: none;  border-radius: 50px; box-shadow: rgba(0, 0, 0, 0.2) 0px 12px 28px 0px, rgba(0, 0, 0, 0.1) 0px 2px 4px 0px, rgba(255, 255, 255, 0.05) 0px 0px 0px 1px inset; lightslategray; display: flex; flex-direction:row; align-items: center; cursor: pointer;"> <img src="https://img.icons8.com/ios/50/zoom.png" alt="" style="width:25px; margin-left:0px; margin-right:9px; flex-direction: row-reverse;">Zoom</button></a>';
            cB = cB + '<a href="" style="text-decoration: none;"><button style="margin: 5px; color: #000000; background-color: #ffffff; padding: 10px 10px; border: none;  border-radius: 50px; box-shadow: rgba(0, 0, 0, 0.2) 0px 12px 28px 0px, rgba(0, 0, 0, 0.1) 0px 2px 4px 0px, rgba(255, 255, 255, 0.05) 0px 0px 0px 1px inset; lightslategray; display: flex; flex-direction:row; align-items: center; cursor: pointer;"><img src="https://img.icons8.com/ios/50/hdmi-cable.png" alt="" style="width:25px; margin-left:0px; margin-right:9px; flex-direction: row-reverse;">HDMI </button></a>';
            cB = cB + '<a href="" style="text-decoration: none;"><button style="margin: 5px; color: #000000; background-color: #ffffff; padding: 10px 10px; border: none;  border-radius: 50px; box-shadow: rgba(0, 0, 0, 0.2) 0px 12px 28px 0px, rgba(0, 0, 0, 0.1) 0px 2px 4px 0px, rgba(255, 255, 255, 0.05) 0px 0px 0px 1px inset; lightslategray; display: flex; flex-direction:row; align-items: center; cursor: pointer;"><img src="https://img.icons8.com/ios/50/usb-c.png" alt="" style="width:25px; margin-left:0px; margin-right:9px; flex-direction: row-reverse;">USB-C </button></a>';
            cB = cB + '<a href="" style="text-decoration: none;"><button style="margin: 5px; color: #000000; background-color: #ffffff; padding: 10px 10px; border: none;  border-radius: 50px; box-shadow: rgba(0, 0, 0, 0.2) 0px 12px 28px 0px, rgba(0, 0, 0, 0.1) 0px 2px 4px 0px, rgba(255, 255, 255, 0.05) 0px 0px 0px 1px inset; lightslategray; display: flex; flex-direction:row; align-items: center; cursor: pointer;"><img src="https://img.icons8.com/ios/50/air-play.png" alt="" style="width:25px; margin-left:0px; margin-right:9px; flex-direction: row-reverse;">Airplay </button></a>';

            cB = cB + '</div>';

            // cB = cB + '<a href="">    </a>';
            // cB = cB + '<a href="" style="text-align: center; color: #36454F; background-color: white; border: none; padding: 10px; text-align: center; text-decoration: none; display: inline-block; font-size: 12px;  border-radius: 30px; box-shadow: 2px 5px 7px lightslategray; margin: 10px;">  <img style="margin-right: 5px;" src="https://localhost:3000/assets/icons8-zoom-16.png" alt="">Zoom</a>';
            // cB = cB + '<a href="" style="text-align: center; color: #36454F; background-color: white; border: none; padding: 10px; text-align: center; text-decoration: none; display: inline-block; font-size: 12px;  border-radius: 30px; box-shadow: 2px 5px 7px lightslategray; margin: 10px;"><img style="margin-right: 5px;" src="https://localhost:3000/assets/icons8-hdmi-cable-16.png" alt="">HDMI</a>';
            // cB = cB + '<a href="" style="text-align: center; color: #36454F; background-color: white; border: none; padding: 10px; text-align: center; text-decoration: none; display: inline-block; font-size: 12px;  border-radius: 30px; box-shadow: 2px 5px 7px lightslategray; margin: 10px;"><img style="margin-right: 5px;" src="https://localhost:3000/assets/icons8-usb-c-16.png" alt="">USB-C</a>';
            // cB = cB + '<a href="" style="text-align: center; color: #36454F; background-color: white; border: none; padding: 10px; text-align: center; text-decoration: none; display: inline-block; font-size: 12px;  border-radius: 30px; box-shadow: 2px 5px 7px lightslategray; margin: 10px;"><img style="margin-right: 5px;" src="https://localhost:3000/assets/icons8-airplay-16.png" alt="">Airplay</a>';
            cB = cB + '</div>';
            cB = cB + '<div style="margin: 5px; display: flex; flex-direction: row;">';
            cB = cB + '<div>';
            cB = cB + '<img style="margin-top: 10px; margin-right: 20px; border-radius: 25px;" src="https://api.mapbox.com/styles/v1/mapbox/streets-v12/static/-122.4241,37.78,15.25,0,60/150x150?access_token=pk.eyJ1IjoiYXJoYW1teG8iLCJhIjoiY2xrd3JobnF6MTVlMTNzbWEzc253c2E0byJ9.k-2gE01ESFIhQc5jghw0kA" alt="">';
            cB = cB + '</div>';
            cB = cB + '<div>';
            cB = cB + '<br><span style="font-size:18px; font-weight: bolder;">Microsoft</span><br>';
            cB = cB + '<p style="margin-top: -20px; padding-right: 50px;">No. 5 (Epitome, Cyber City, 10th Floor, TowerB & C, DLF Building, DLF Phase 3, Gurugram, Harayana -122002</p>';
            cB = cB + '<a href="" style="text-decoration:none; color: green; margin-top: -20px ;">https://maps.app.goo.gl/PEhz6UsEKZSRgPNV8</a>';
            cB = cB + '</div>';
            cB = cB + '</div>';
            cB = cB + '<div style="float: right; position:relative; display:inline-block; padding-top: 30px;">';
            cB = cB + '<div style="display: flex; flex-direction: row; margin-bottom:0px">';
            cB = cB + '<div>';
            cB = cB + '<a href="" style="cursor: pointer; font-size: small; padding-left: 20px; padding-right: 20px; color: black; text-align: center; display: inline-block; text-decoration:none; margin: 5px;"><img src="https://localhost:3000/assets/icons8-checked-identification-documents-32.png" alt=""><br>Get Visitor Access</a>';
            cB = cB + '</div>';
            cB = cB + '<div>';
            cB = cB + '<a href="" style="cursor: pointer; font-size: small; padding-left: 20px; padding-right: 20px; color: black; text-align: center; display: inline-block; text-decoration:none; margin: 5px;"><img src="https://localhost:3000/assets/icons8-wifi-32.png" alt=""><br>Get Guest WiFi Access</a>';
            cB = cB + '</div>';
            cB = cB + '<div>';
            cB = cB + '<a href="" style="cursor: pointer; font-size: small; padding-left: 20px; padding-right: 20px; color: black; text-align: center; display: inline-block; text-decoration:none; margin: 5px;"><img src="https://localhost:3000/assets/icons8-dots-32.png" alt=""><br>More Information</a>';
            cB = cB + '</div>';
            cB = cB + '</div>';
            cB = cB + '</div>';
            cB = cB + '</div>';
            cB = cB + '</div>';
            cB = cB + '</body>';


            // cB = cB + '<body style="font-family: Open Sans;">';
            // cB = cB + '<div style="position: relative; margin-left:20px; display: flex; flex-direction: row; border-radius: 25px; background: whitesmoke; padding: 20px; box-shadow: 5px 5px 7px lightslategray;">';
            // cB = cB + '<div style="z-index: 1;">>';
            // cB = cB + '<img style="border-radius: 50px 40px; max-width: 500px;" src="https://localhost:3000/assets/br.png" alt="image">';
            // cB = cB + '</div>';
            // cB = cB + '<div style="position: absolute; position: absolute; bottom: 5px; left: 20px; z-index: 2;">';
            // cB = cB + '<p style="margin-left: 5px; margin-bottom: 0; font-size: xx-large; font-weight: bolder;">Bamboo Room</p>';
            // cB = cB + '<p style="margin-left: 5px; margin-top: 0;"><span style="font-weight: bolder;">First Floor, Business Plaza</span> <span>(Last room on the left)</span></p>';
            // cB = cB + '<div style="margin-top: -10px;>';
            // cB = cB + '<a href="" style="background-color: white; border: none; padding: 10px; text-align: center; text-decoration: none; display: inline-block; font-size: 12px;  border-radius: 10px; box-shadow: 2px 5px 7px lightslategray; margin: 5px;"> </a>';
            // cB = cB + '<a href="" style="background-color: white; border: none; padding: 10px; text-align: center; text-decoration: none; display: inline-block; font-size: 12px;  border-radius: 10px; box-shadow: 2px 5px 7px lightslategray; margin: 5px;">Zoom</a>';
            // cB = cB + '<a href="" style="background-color: white; border: none; padding: 10px; text-align: center; text-decoration: none; display: inline-block; font-size: 12px;  border-radius: 10px; box-shadow: 2px 5px 7px lightslategray; margin: 5px;">HDMI</a>';
            // cB = cB + '<a href="" style="background-color: white; border: none; padding: 10px; text-align: center; text-decoration: none; display: inline-block; font-size: 12px;  border-radius: 10px; box-shadow: 2px 5px 7px lightslategray; margin: 5px;">USB-C</a>';
            // cB = cB + '<a href="" style="background-color: white; border: none; padding: 10px; text-align: center; text-decoration: none; display: inline-block; font-size: 12px;  border-radius: 10px; box-shadow: 2px 5px 7px lightslategray; margin: 5px;">Airplay</a>';
            // cB = cB + '</div>';
            // cB = cB + '<div style="margin: 5px; padding:10px; display: flex; flex-direction: row;">';
            // cB = cB + '<div>';
            // cB = cB + '<img style="margin-top: 10px; margin-right: 20px; border-radius: 25px;" src="https://localhost:3000/assets/map.png" alt="">';
            // cB = cB + '</div>';
            // cB = cB + '<div>';
            // cB = cB + '<h2>Microsoft</h2>';
            // cB = cB + '<p style="margin-top: -10px; padding-right: 50px;">No. 5 (Epitome, Cyber City, 10th Floor, TowerB & C, DLF Building, DLF Phase 3, Gurugram, Harayana -122002</p>';
            // cB = cB + '<a href="" style="color: darkgreen; margin-top: -10px; font-size: small;">https://maps.app.goo.gl/PEhz6UsEKZSRgPNV8</a>';
            // cB = cB + '</div>';
            // cB = cB + '</div>';
            // cB = cB + '</div>';
            // cB = cB + '</div>';
            // cB = cB + '</body>'
            cB = cB + '</html>';

            // let cB = '<pre><code>';
            // for (i = 0; i < gist.length; i++) {
            //     cB = cB + gist[i].name;
            //     cB = cB + "<br></br>";
            // }
            // cB = cB + '</code></pre>';
            Office.context.mailbox.item.body.setSelectedDataAsync(cB, { coercionType: Office.CoercionType.Html }, function(result) {
                event.completed();
            });
            // Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/outlook/90-other-item-apis/get-set-location-appointment-organizer.yaml


            const location = "New Delhi";
            Office.context.mailbox.item.location.setAsync(location, (result) => {
                if (result.status !== Office.AsyncResultStatus.Succeeded) {
                    console.error(`Action failed with message ${result.error.message}`);
                    return;
                }
                console.log(`Successfully set location to ${location}`);
            });
            return codeBlock;

        }).fail(function(error) {
            const sks = JSON.stringify(error);
            Office.context.mailbox.item.body.setSelectedDataAsync("diff err " + sks, { coercionType: Office.CoercionType.Html }, function(result) {
                event.completed();
            });
            return converter.makeHtml(error);
        });
    } catch (e) {
        const sk = JSON.stringify(e);
        Office.context.mailbox.item.body.setSelectedDataAsync("erroe" + sk, { coercionType: Office.CoercionType.Html }, function(result) {
            event.completed();
        });
    }




    // try {
    //     const z = extr();
    //     Office.context.mailbox.item.body.setSelectedDataAsync(z, { coercionType: Office.CoercionType.Html }, function(result) {
    //         event.completed();
    //     });
    // } catch (e) {
    //     const sk = JSON.stringify(e);
    //     Office.context.mailbox.item.body.setSelectedDataAsync("erroe" + sk, { coercionType: Office.CoercionType.Html }, function(result) {
    //         event.completed();
    //     });
    // }


    // $.ajax({
    //     url: endpoint,
    //     data: JSON.stringify({'1':'2'}),
    //     // headers: {'X-Requested-With': 'XMLHttpRequest'},
    //     contentType: 'text/plain',
    //     type: 'POST',
    //     dataType: 'json',
    //     error: function(xhr, status, error) {
    //         // error
    //       }
    // }).done(function(data) {
    //     // done
    //   });

    // config = getConfig();
    // dat = JSON.stringify('/default/meeting.json');

    // // Check if the add-in has been configured.
    // if (config && config.defaultGistId) {
    //     // Get the default gist content and insert.
    //     try {
    //         getGist(config.defaultGistId, function(gist, error) {
    //             if (gist) {
    //                 buildBodyContent(gist, function(content, error) {
    //                     if (content) {
    //                         Office.context.mailbox.item.body.setSelectedDataAsync(dat, { coercionType: Office.CoercionType.Html }, function(result) {
    //                             event.completed();
    //                         });
    //                     } else {
    //                         showError(error);
    //                         event.completed();
    //                     }
    //                 });
    //             } else {
    //                 showError(error);
    //                 event.completed();
    //             }
    //         });
    //     } catch (err) {
    //         showError(err);
    //         event.completed();
    //     }

    // } else {
    //     // Save the event object so we can finish up later.
    //     btnEvent = event;
    //     // Not configured yet, display settings dialog with
    //     // warn=1 to display warning.
    //     const url = new URI('dialog.html?warn=1').absoluteTo(window.location).toString();
    //     const dialogOptions = { width: 20, height: 40, displayInIframe: true };

    //     Office.context.ui.displayDialogAsync(url, dialogOptions, function(result) {
    //         settingsDialog = result.value;
    //         settingsDialog.addEventHandler(Office.EventType.DialogMessageReceived, receiveMessage);
    //         settingsDialog.addEventHandler(Office.EventType.DialogEventReceived, dialogClosed);
    //     });
    // }
}

// Register the function.
Office.actions.associate("insertDefaultGist", insertDefaultGist);

function receiveMessage(message) {
    config = JSON.parse(message.message);
    setConfig(config, function(result) {
        settingsDialog.close();
        settingsDialog = null;
        btnEvent.completed();
        btnEvent = null;
    });
}

function dialogClosed(message) {
    settingsDialog = null;
    btnEvent.completed();
    btnEvent = null;
}