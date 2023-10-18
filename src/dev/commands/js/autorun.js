/* global Office, axios, console */
import "../../../apis";

const SIGNATURE_API_URL = "api/v2/app/reach/get-signature";
const API_BASE_URL = getBaseUrl('development');

// let isFirstTime = 0,
//   signature;

// function init() {
//   signature = Office.context.roamingSettings.get("signature");
//   isFirstTime = Office.context.roamingSettings.get("isFirstTime");

//   console.log(isFirstTime);
// }

function getSignature(token) {
  // `${_host}/domain/user/view_signature/raw?token=${e}&email=${_email}&client=outlook&platform=${_platform}&client_version=${_client_version}&addin_js_version=1.3.0`
  const email = Office.context.mailbox.userProfile.emailAddress;
  const url = `${API_BASE_URL}/${SIGNATURE_API_URL}/outlook/`;
  return axios.get(url, {
    params: {
      email,
      token,
    },
  });
}

function checkSignature(e) {
  // init();
  // checkIfFirstTime();
  // console.log(e);
  Office.context.mailbox.item.saveAsync(function (t) {
    // console.log(t);
    if (t.status === Office.AsyncResultStatus.Succeeded) {
      Office.context.mailbox.getCallbackTokenAsync({ isRest: !0 }, function (t) {
        getSignature(t.value)
          .then((res) => {
            Office.context.mailbox.item.body.setSignatureAsync(res.data, {
              coercionType: Office.MailboxEnums.BodyType.Html,
            });
            // TODO: Later
            // saveUserSignature(res.data);
          })
          .catch((error) => {
            if (error.response?.data?.message) {
              console.error(error.response?.data?.message);
            } else {
              console.log(error);
            }
          });
      });
    } else {
      console.error("Item save failed: " + t.error.message);
    }
  });
  e.completed();
}

// TODO: Later implement below methods

// function checkIfFirstTime() {
//   saveRoamingSettings("isFirstTime", void 0 === isFirstTime ? !signature : isFirstTime);
// }

// function saveUserSignature(signature) {
//   Office.context.roamingSettings.set("signature", signature);
//   Office.context.roamingSettings.saveAsync(function (e) {
//     console.log("saveToRoamingSettings - " + JSON.stringify(e));
//   });
// }

// function saveRoamingSettings(key, payload) {
//   Office.context.roamingSettings.set(key, payload);
//   Office.context.roamingSettings.saveAsync(function (e) {
//     console.log("saveToRoamingSettings - " + JSON.stringify(e));
//   });
// }

Office.actions.associate("checkSignature", checkSignature);
