/**
 * Functionalities to add:
 * 
 * - Possibility to have relative path instead of absolute path
 * - Relative path adaptable to sub-domains and different TLDs
 * 
**/

function crawlSheetFormat() {
  // Storing the active sheet reference
  let crawlSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()

  // Splitting .csv file to columns
  crawlSheet.getRange("A:A").splitTextToColumns()

  // Storing useful ranges in variables
  let crawlRange = crawlSheet.getDataRange()
  let numberOfColumns = crawlRange.getNumColumns()
  let numberOfRows = crawlRange.getNumRows()
  let crawlFirstRow = crawlSheet.getRange("1:1")
  let crawlFirstColumn = crawlSheet.getRange("A:A")

  // Hiding useless columns
  let columnsToHide = [{ "value": "Title 1 Length"},
                       { "value": "Meta Description 1 Length"},
                       { "value": "Meta Keywords 1"},
                       { "value": "Meta Keywords 1 Length"},
                       { "value": "H1-1 Length"},
                       { "value": "H1-2"},
                       { "value": "H1-2 Length"},
                       { "value": "H2-1"},
                       { "value": "H2-1 Length"},
                       { "value": "H2-2"},
                       { "value": "H2-2 Length"},
                       { "value": "X-Robots-Tag 1"},
                       { "value": "Meta Refresh 1"},
                       { "value": 'HTTP rel="next" 1'},
                       { "value": 'HTTP rel="prev" 1'},
                       { "value": "amphtml Link Element"},
                       { "value": "Unique JS Inlinks"},
                       { "value": "% of Total"},
                       { "value": "Unique JS Outlinks"},
                       { "value": "Unique External JS Outlinks"},
                       { "value": "Closest Similarity Match"},
                       { "value": "No. Near Duplicates"},
                       { "value": "Spelling Errors"},
                       { "value": "Grammar Errors"},
                       { "value": "Hash"},
                       { "value": "Response Time"},
                       { "value": "Crawl Timestamp"}]

  const columnToHideHandler = (sheet, cellContent, rowIndex, columnIndex) => {
    sheet.hideColumn(sheet.getRange(rowIndex,columnIndex))
  }

  callbackFromRowValueLookup(crawlSheet,columnsToHide,1,columnToHideHandler)


  // Adding an hyperlink to the first column
  let adresses = crawlFirstColumn.getValues()
  let hyperlinks = []
  adresses.forEach((url,index) => hyperlinks[index] = [`=HYPERLINK("${url}")`])
  hyperlinks[0] = ["Address"]
  crawlFirstColumn.setValues(hyperlinks)


  // Freeze the first row/column
  crawlSheet.setFrozenColumns(1)
  crawlSheet.setFrozenRows(1)


  // Apply general style
  crawlRange.setHorizontalAlignment("left")
  crawlRange.setVerticalAlignment("middle")
  crawlRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  crawlSheet.setRowHeightsForced(1,numberOfRows,35)


  // Apply first row style
  crawlFirstRow.setHorizontalAlignment("center")
  crawlFirstRow.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  crawlSheet.setRowHeightsForced(1,1,45)
  

  // Modify column widths
  let columnsToModifyWidth = [{"value": "Address", "width": 600},
                              {"value": "Content Type", "width": 150},
                              {"value": "Indexability Status", "width": 120},
                              {"value": "Title 1", "width": 200},
                              {"value": "Meta Description 1", "width": 200},
                              {"value": "H1-1", "width": 200},
                              {"value": "Meta Robots 1", "width": 160},
                              {"value": "Canonical Link Element 1", "width": 400},
                              {"value": "Redirect URL", "width": 400},
                              {"value": "URL Encoded Address", "width": 400}]

  const columnToModifyWidthHandler = (sheet, cellContent, rowIndex, columnIndex) => {
    sheet.setColumnWidth(columnIndex, cellContent.width)
  }

  callbackFromRowValueLookup(crawlSheet,columnsToModifyWidth,1,columnToModifyWidthHandler)


  // Apply filter to sheet
  crawlRange.createFilter()


  // Add conditional formating to columns
  let crawlConditionalRules = crawlSheet.getConditionalFormatRules()

  // --- Status code
  let statusCodeColumnIndex = getColumnIndexFromValue(crawlSheet, "Status Code", 1)
  let statusCodeRange = crawlSheet.getRange(2, statusCodeColumnIndex, numberOfRows - 1, 1)

  crawlConditionalRules.push(createNewConditionalFormatRule(statusCodeRange, "numberEqualTo", 404, "red"))
  crawlConditionalRules.push(createNewConditionalFormatRule(statusCodeRange, "numberEqualTo", 0, "orange"))
  crawlConditionalRules.push(createNewConditionalFormatRule(statusCodeRange, "numberNotEqualTo", 200, "yellow"))

  // --- Indexability
  let indexabilityColumnIndex = getColumnIndexFromValue(crawlSheet, "Indexability", 1)
  let indexabilityRange = crawlSheet.getRange(2, indexabilityColumnIndex, numberOfRows - 1, 1)

  let indexabilityStatusColumnIndex = getColumnIndexFromValue(crawlSheet, "Indexability Status", 1)
  let indexabilityStatusRange = crawlSheet.getRange(2, indexabilityStatusColumnIndex, numberOfRows - 1, 1)

  crawlConditionalRules.push(createNewConditionalFormatRule(indexabilityRange, "formulaSatisfied", `=IF(E2<>"Indexable",1,0)`, "yellow"))
  crawlConditionalRules.push(createNewConditionalFormatRule(indexabilityStatusRange, "cellNotEmpty", "", "yellow"))

  // --- Title
  let titleColumnIndex = getColumnIndexFromValue(crawlSheet, "Title 1", 1)
  let titleRange = crawlSheet.getRange(2, titleColumnIndex, numberOfRows - 1, 1)

  let titlePixelWidthColumnIndex = getColumnIndexFromValue(crawlSheet, "Title 1 Pixel Width", 1)
  let titlePixelWidthRange = crawlSheet.getRange(2, titlePixelWidthColumnIndex, numberOfRows - 1, 1)

  crawlConditionalRules.push(createNewConditionalFormatRule(titleRange, "cellEmpty", "", "orange"))
  crawlConditionalRules.push(createNewConditionalFormatRule(titlePixelWidthRange, "numberGreaterThan", 575, "orange"))
  crawlConditionalRules.push(createNewConditionalFormatRule(titlePixelWidthRange, "numberLessThan", 280, "yellow"))

  // --- Meta Description
  let metaDescriptionColumnIndex = getColumnIndexFromValue(crawlSheet, "Meta Description 1", 1)
  let metaDescriptionRange = crawlSheet.getRange(2, metaDescriptionColumnIndex, numberOfRows - 1, 1)

  let metaDescriptionPixelWidthColumnIndex = getColumnIndexFromValue(crawlSheet, "Meta Description 1 Pixel Width", 1)
  let metaDescriptionPixelWidthRange = crawlSheet.getRange(2, metaDescriptionPixelWidthColumnIndex, numberOfRows - 1, 1)

  crawlConditionalRules.push(createNewConditionalFormatRule(metaDescriptionRange, "cellEmpty", "", "orange"))
  crawlConditionalRules.push(createNewConditionalFormatRule(metaDescriptionPixelWidthRange, "numberGreaterThan", 920, "orange"))
  crawlConditionalRules.push(createNewConditionalFormatRule(metaDescriptionPixelWidthRange, "numberLessThan", 340, "yellow"))

  // --- H1
  let h1ColumnIndex = getColumnIndexFromValue(crawlSheet, "H1-1", 1)
  let h1Range = crawlSheet.getRange(2, h1ColumnIndex, numberOfRows - 1, 1)

  crawlConditionalRules.push(createNewConditionalFormatRule(h1Range, "cellEmpty", "", "orange"))

  // --- Word count
  let wordCountColumnIndex = getColumnIndexFromValue(crawlSheet, "Word Count", 1)
  let wordCountRange = crawlSheet.getRange(2, wordCountColumnIndex, numberOfRows - 1, 1)

  crawlConditionalRules.push(createNewConditionalFormatRule(wordCountRange, "numberLessThan", 500, "orange"))

  // --- Crawl depth
  let crawlDepthColumnIndex = getColumnIndexFromValue(crawlSheet, "Crawl Depth", 1)
  let crawlDepthRange = crawlSheet.getRange(2, crawlDepthColumnIndex, numberOfRows - 1, 1)

  crawlConditionalRules.push(createNewConditionalFormatRule(crawlDepthRange, "numberGreaterThanOrEqualTo", 3, "orange"))

  // --- Inlinks
  let inlinksColumnIndex = getColumnIndexFromValue(crawlSheet, "Inlinks", 1)
  let inlinksRange = crawlSheet.getRange(2, inlinksColumnIndex, numberOfRows - 1, 1)

  let uniqueInlinksColumnIndex = getColumnIndexFromValue(crawlSheet, "Unique Inlinks", 1)
  let uniqueInlinksRange = crawlSheet.getRange(2, uniqueInlinksColumnIndex, numberOfRows - 1, 1)

  crawlConditionalRules.push(createNewConditionalFormatRule(inlinksRange, "numberLessThan", 10, "orange"))
  crawlConditionalRules.push(createNewConditionalFormatRule(uniqueInlinksRange, "numberLessThan", 10, "orange"))
  
  // --- Applying the rules
  crawlSheet.setConditionalFormatRules(crawlConditionalRules)

  
  // Set today's date as tab name
  let tabName = new Date()
  crawlSheet.setName(tabName.toLocaleDateString())


  // Focus on Content-type cell to prepare filtering HTML content
  crawlSheet.setCurrentCell(crawlSheet.getRange("B1"))
}



/**
 * Function:
 * - Return the column index by looking-up a value on a selected row
 * 
 * Parameters:
 * - Sheet
 * - Value to look-up
 * - Index of the row to search
 * 
 * Return:
 * - Column index of the first matched cell
 */
function getColumnIndexFromValue(sheet, value, rowIndex) {
  let numberOfColumns = sheet.getDataRange().getNumColumns()
  let rowRange = sheet.getRange(rowIndex,1,1,numberOfColumns)
  let rowCellValues = rowRange.getValues()[0]
  let rowCellValue = ""

  for (let i = 0; i < rowCellValues.length; i++) {
    rowCellValue = rowCellValues[i]
    if (rowCellValue == value) {
      return i+1
    }
  }
}



/**
 * Function:
 * - Look-up for an array of values on the contents of one row
 *   Performs a callback containing useful data when a result is found
 * 
 * Parameters:
 * - Sheet
 * - Array of objects to look-up
 *     - "value" key is required (when an equal value is found, callback is called)
 *     - additional keys can be added to use in the callback
 * - Index of the row to search
 * - Callback performed when a match exists between row value and object value
 * 
 * Return:
 * - Nothing
 * 
 * Callback parameters:
 * - Sheet
 * - matched Object
 * - Row index of the matched cell
 * - Column index of the matched cell
 */
function callbackFromRowValueLookup(sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(), cells = [], rowIndex = 1, callback) {
  let numberOfColumns = sheet.getDataRange().getNumColumns()
  let rowRange = sheet.getRange(rowIndex,1,1,numberOfColumns)
  let rowCellValues = rowRange.getValues()[0]
  let cellContent = {}
  let rowCellValue = ""
  
  for (let i = 0; i < cells.length; i++) {
    cellContent = cells[i]
    for (let j = 0; j < rowCellValues.length; j ++) {
      rowCellValue = rowCellValues[j]
      if (cellContent.value == rowCellValue) {
        // sheet reference passed to the callback
        // cell content object passed to the callback
        // cell row index passed to the callback
        // j+1 <=> cell column index passed to the callback
        callback(sheet, cellContent, rowIndex, j+1)
      }
    }
  }
}



/**
 * Function:
 * - Use parameters to push a new conditional formatting rule to the array of rules
 *   Do not forget to apply the rules once everything is set
 * 
 * - Parameters:
 * - Array of rules
 * - Range where the rule will be applied (one range only)
 * - Type of rule to use (see acceptable values below)
 * - Value to use in the rule
 * - Color to use:
 *   - green
 *   - yellow
 *   - orange
 *   - red
 * 
 * - Return:
 * - Array of rules with the new rule pushed in
 */
function createNewConditionalFormatRule(ruleRange, ruleType, ruleValue, ruleColor) {
  let rule = SpreadsheetApp.newConditionalFormatRule()

  switch (ruleType) {
    case 'numberEqualTo':
      rule.whenNumberEqualTo(ruleValue)
      break
    case 'numberNotEqualTo':
      rule.whenNumberNotEqualTo(ruleValue)
      break
    case 'numberGreaterThan':
      rule.whenNumberGreaterThan(ruleValue)
      break
    case 'numberGreaterThanOrEqualTo':
      rule.whenNumberGreaterThanOrEqualTo(ruleValue)
      break
    case 'numberLessThan':
      rule.whenNumberLessThan(ruleValue)
      break
    case 'numberLessThanOrEqualTo':
      rule.whenNumberLessThanOrEqualTo(ruleValue)
      break
    case 'formulaSatisfied':
      rule.whenFormulaSatisfied(ruleValue)
      break
    case 'cellEmpty':
      rule.whenCellEmpty()
      break
    case 'cellNotEmpty':
      rule.whenCellNotEmpty()
      break
    default:
      alert("Conditional rule: 'ruleType' error")
      return false
  }

  let greenColor = "#d9ead3"
  let yellowColor = "#fff2cc"
  let orangeColor = "#fce5cd"
  let redColor = "#f4cccc"

  switch (ruleColor) {
    case 'green':
      rule.setBackground(greenColor)
      break
    case 'yellow':
      rule.setBackground(yellowColor)
      break
    case 'orange':
      rule.setBackground(orangeColor)
      break
    case 'red':
      rule.setBackground(redColor)
      break
    default:
      alert("Conditional rule: 'ruleColor' error")
      return false
  }

  rule.setRanges([ruleRange])

  rule.build()

  return rule
}
