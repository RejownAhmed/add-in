/* global Office, axios, console */
const SIGNATURE_API_URL = "api/v2/app/reach/get-signature";
const API_BASE_URL = "https://reach-api.test";

function getSignature(token) {
  // `${_host}/domain/user/view_signature/raw?token=${e}&email=${_email}&client=outlook&platform=${_platform}&client_version=${_client_version}&addin_js_version=1.3.0`
  const email = Office.context.mailbox.userProfile.emailAddress;
  const url = `${API_BASE_URL}/${SIGNATURE_API_URL}/outlook/`;
  return axios.get(url, {
    params: {
      email,
    },
  });
}
function checkSignature(e) {
  console.log(e);
  Office.context.mailbox.item.saveAsync(function (t) {
    // console.log(t);
    Office.context.mailbox.getCallbackTokenAsync({ isRest: !0 }, function (t) {
      getSignature(t.value)
        .then((res) => {
          Office.context.mailbox.item.body.setSignatureAsync(res.data, {
            coercionType: Office.MailboxEnums.BodyType.Html,
          });
        })
        .catch((error) => {
          console.log(error);
          if (error.response?.data?.message) {
            console.error(error.response?.data?.message);
          }
        });
    });
  });
}

Office.actions.associate("checkSignature", checkSignature);
