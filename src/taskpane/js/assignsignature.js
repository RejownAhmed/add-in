/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    // showLoader(), initConstants(), void 0 === _isFirstTime || _isFirstTime ? showWelcomeSection() :
    renderMainPage();
  }
});

async function renderMainPage() {
  Office.context.mailbox.item.saveAsync(function (e) {
    console.log(e, "assignsignature page");
  });
}
