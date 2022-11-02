function onOpen() {
  let ui = SpreadsheetApp.getUi()

  ui.createMenu("Process")
    .addItem("Format crawl export", "crawlSheetFormat")
    .addToUi()
}
