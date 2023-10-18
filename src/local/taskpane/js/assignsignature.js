/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, Office */

Office.onReady(() => {
  renderMainPage();
});

async function renderMainPage() {
  console.log("assignsignature page");
  Office.context.mailbox.item.saveAsync(function (e) {
    console.log(e, "assignsignature page");
  });
}
