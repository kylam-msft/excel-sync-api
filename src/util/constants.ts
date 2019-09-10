import { stripSpaces } from "./strip-spaces";

export const SAMPLES = {
  BASIC_SAMPLE: stripSpaces(`
    const workbook = ExcelScript.workbook;
    const sheet = workbook.getWorksheet("Sheet1");
    const range = sheet.getRange("A1");
    range.setValues([['Hello']]);
    range.getFormat().getFill().setColor('yellow');

    const address = range.getAddress();
    console.log("Address: " + address);
  `),
  INVENTORY_MAKER: stripSpaces(`
  const workbook = ExcelScript.workbook;
  const worksheetPrefix = "reorder";
  /**
   * Load the vendor list into an object for lookup
   */
  console.log("Reading vendor data");
  const vendorList = workbook
    .getWorksheet("Vendor List")
    .getTable("vendor")
    .getDataBodyRange()
    .getValues();
  const vendorObj = {};
  for (let i = 0; i < vendorList.length; i++) {
    let key = vendorList[i][0];
    let val = vendorList[i][2];
    vendorObj[key] = val;
  }
  /**
   * Get inventory table rows
   */
  const inventoryTable = workbook.getWorksheet("inventory").getTable("inventory");
  const rows = inventoryTable.getRows();
  /**
   * Table header constant row
   */
  /**
   * Table header constant row
   */
  const tableHeaders = [
    "Inventory ID",
    "Name",
    "Unit Price",
    "Quantity in Stock",
    "Quantity in Reorder",
    "Reorder price",
    "Vendor email",
  ];
  /**
   * Prepare output array to be used for re-order table.
   */
  let outArray = [tableHeaders];
  /**
   * Read through each rows
   */
  console.log("Preparing re-order data");
  for (let i = 0; i < rows.length; i++) {
    let inventoryValues = rows[i].getValues();

    // Ignore discontinued rows.
    if (inventoryValues[0][9] === "Yes") continue;
    // If inventory is less than re-order level, add to the output array. Select certain columns.

    if (inventoryValues[0][4] < inventoryValues[0][6]) {
      outArray.push([
        inventoryValues[0][0],
        inventoryValues[0][1],
        inventoryValues[0][3],
        inventoryValues[0][4],
        inventoryValues[0][8],
        // Make room for formula
        "",
        vendorObj[inventoryValues[0][0]],
      ]);
    }
  }
  console.log(outArray.length - 1 + " items to be added to reorder table for this period");
  /**
   * Create re-order worksheet with name of reorder{date} .
   */
  var today = new Date();
  var res = today
    .toISOString()
    .slice(0, 10)
    .replace(/-/g, "");
  const wsname = worksheetPrefix + res;
  workbook.getWorksheet(wsname).delete();
  console.log("Creating new reorder worksheet: " + wsname);
  const reorderWorksheet = workbook.addWorksheet(wsname);
  /**
   * Get target range reference based on the size of the output array and update its values and auto fit columns.
   */
  console.log("Adding re-order range data and adding a table on top");
  let targetRange = reorderWorksheet.getRange("A1").getAbsoluteResizedRange(outArray.length, tableHeaders.length);
  targetRange.setValues(outArray);
  targetRange.getFormat().autofitColumns();
  /**
   * Create re-order table on top of the range data
   */

  let reorderTable = reorderWorksheet.addTable(targetRange.getAddress(), true);
  /**
   * Update the formula to calculate reorder price.
   */
  const reorderValueColRange = reorderTable
    .getColumn("Reorder price")
    .getDataBodyRange()
    .getCell(0, 0);
  reorderValueColRange.setValues([["=[@[Unit Price]]*[@[Quantity in Reorder]]"]]);
  /**
   * Run through the reorder table and color rows with current quantity of 0 with red and ones with < 10 with yellow.
   */
  for (let i = 0; i < outArray.length - 1; i++) {
    let row = reorderTable.getRowAt(i);
    let rowValues = row.getValues();
    if (rowValues[0][3] == "0") {
      row.getRange().getFormat().getFill().setColor("red");
      continue;
    }
    if (rowValues[0][3] < "10") {
      row.getRange().getFormat().getFill().setColor("yellow");
      continue;
    }
  }
  /**
   * Add legend for yellow and red colors exactly 1 column after the table to the right.
   * First calculate the cell address where the legend data should be added.
   */
  let legendRow = reorderTable.getHeaderRowRange().getResizedRange(1, 2);
  const legendColumnCount = legendRow.getColumnCount();
  const legendYellow = legendRow.getCell(0, legendColumnCount - 1);
  const legendRed = legendRow.getCell(1, legendColumnCount - 1);
  legendRed.setValues([["Out of stock"]]);
  legendRed.getFormat().getFill().setColor("Red");
  legendYellow.setValues([["Low inventory"]]);
  legendYellow.getFormat().getFill().setColor("Yellow");
  legendRed.getFormat().autofitColumns();
  /**
   * Add index sheet each time and refresh the list of reorder sheet links
   */
  workbook.getWorksheet("index").delete();
  let worksheets = workbook.getWorksheets();

  // Select only sheets beginning with reorder
  let indexes = [];
  for (const worksheet of worksheets) {
    const name = worksheet.getName();
    if (name.startsWith("reorder")) {
      indexes.push(worksheet.getName());
    }
  }
  // Transpose indexes array to be used in range update.
  let indexesForRangeUpdate = [];
  for (let s of indexes) {
    indexesForRangeUpdate.push([s]);
  }
  // Add index worksheet and position it as the first worksheet
  console.log("Adding index sheet");
  const indexSheet = workbook.addWorksheet("index");
  indexSheet.setPosition(0);
  // Update the index sheet range with list of all reorder worksheets
  let indexTargetRange = indexSheet.getRange("A1").getAbsoluteResizedRange(indexes.length, 1);
  indexTargetRange.setValues(indexesForRangeUpdate);
  indexTargetRange.getFormat().autofitColumns();

  console.log("Done!");
  `),
};
