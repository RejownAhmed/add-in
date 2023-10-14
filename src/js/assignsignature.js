function renderMainPage() {
  Office.context.mailbox.item.saveAsync(function (e) {
    console.log(e, "assignsignature page");
  });
}
