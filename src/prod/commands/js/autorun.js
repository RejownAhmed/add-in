const SIGNATURE_API_URL = "api/v2/app/reach/get-signature";
const API_BASE_URL = getBaseUrl('production');
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// Contains code for event-based activation on Outlook on web, on Windows, and on Mac (new UI preview).

/**
 * Checks if signature exists.
 * If not, displays the information bar for user.
 * If exists, insert signature into new item (appointment or message).
 * @param {*} eventObj Office event object
 * @returns
 */
function checkSignature(eventObj) {
   

  Office.context.mailbox.item.from.getAsync(function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          const msgFrom = asyncResult.value;
          console.log("Message from: " + msgFrom.emailAddress);
          insert_auto_signature(msgFrom.emailAddress, eventObj);
      } else {
          console.log(asyncResult.error);
      }
  });

 //// if (eventObj.source.id == "OnNewMessageCompose") {
 //     let user_info_str = Office.context.roamingSettings.get("user_info");
 //     let user_info = JSON.parse(user_info_str);
 //   let  user_email = user_info.email;

 // insert_auto_signature(user_email, eventObj);

}

  function insert_auto_signature(user_info, eventObj)
  {
      if (Office.context.mailbox.item.getComposeTypeAsync) {
          // Find out if the compose type is "newEmail", "reply", or "forward" so that we can apply the correct template.
          Office.context.mailbox.item.getComposeTypeAsync(
              {
                  asyncContext: {
                      user_info: user_info,
                      eventObj: eventObj,
                  },
              },
              function (asyncResult) {
                  if (asyncResult.status === "succeeded") {
                      get_signature_info(
                          asyncResult.asyncContext.user_info,
                          asyncResult.asyncContext.eventObj
                          
                      );
                  }
              }
          );
      }


  
}


function addTemplateSignature(signatureDetails, eventObj) {

  //Image is not embedded, or is referenced from template HTML
  Office.context.mailbox.item.body.setSignatureAsync(
      signatureDetails,
      {
          coercionType: "html",
          asyncContext: eventObj,
      },
      function (asyncResult) {
          asyncResult.asyncContext.completed();
      }
  );

}

/**
* Creates information bar to display when new message or appointment is created
*/
function display_insight_infobar() {
  Office.context.mailbox.item.notificationMessages.addAsync("fd90eb33431b46f58a68720c36154b4a", {
      type: "insightMessage",
      message: "Please set your signature with the Office Add-ins sample.",
      icon: "Icon.16x16",
      actions: [
          {
              actionType: "showTaskPane",
              actionText: "Set signatures",
              commandId: get_command_id(),
              contextData: "{''}",
          },
      ],
  });
}

/**
* Gets template name (A,B,C) mapped based on the compose type
* @param {*} compose_type The compose type (reply, forward, newMail)
* @returns Name of the template to use for the compose type
*/
function get_template_name(compose_type) {
  if (compose_type === "reply") return Office.context.roamingSettings.get("reply");
  if (compose_type === "forward") return Office.context.roamingSettings.get("forward");
  return Office.context.roamingSettings.get("newMail");
}

/**
* Gets HTML signature in requested template format for given user
* @param {\} template_name Which template format to use (A,B,C)
* @param {*} user_info Information details about the user
* @returns HTML signature in requested template format
*/
function get_signature_info(user_info, eventObj) {
  let apiUrl = "https://reachapi.reach.app/api/v2/app/reach/get-signature/outlook?email=" + user_info;

      fetch(apiUrl)
          .then(function (response) {
              if (!response.ok) {
                  console.log("Network response was not ok");
              }
              return response.text();
          })
          .then(function (signature) {
              // Assuming the API returns the signature as text
              // You may need to adjust this part based on the actual response format
            
              console.log(signature);
              if (signature !== "") {
                  addTemplateSignature(signature,eventObj)

                  // Your logic when the signature is not empty
              }
          })
          .catch(function (error) {
              console.log('Fetch error:', error);
              // Handle errors as needed
          });


  

}



/**
* Gets correct command id to match to item type (appointment or message)
* @returns The command id
*/
function get_command_id() {
  if (Office.context.mailbox.item.itemType == "appointment") {
      return "MRCS_TpBtn1";
  }
  return "MRCS_TpBtn0";
}


Office.actions.associate("checkSignature", checkSignature);
