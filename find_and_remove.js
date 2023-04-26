// Find certain string and remove the entire row

// Input to adjust
const CHOSEN_RANGE = "A3:A880";
const STRING_TO_CHECK = "Total"

// Code to run

function isStringInCellText(inputString: string) {
  return inputString.includes(STRING_TO_CHECK);
}

function main(workbook: ExcelScript.Workbook) {
  let selectedSheet = workbook.getActiveWorksheet();

  let range = selectedSheet.getRange(CHOSEN_RANGE);

  for (let i = 0; i < range.getCellCount(); i++) {
    let currentCell = range.getCell(i, 0);
    let result = currentCell.getText();
    if (isStringInCellText(result)) {
      console.log(result);
      let rowToDelete = currentCell.getRowIndex() + 1;
      let cellsToDelete = 'A' + rowToDelete + ':T' + rowToDelete;
      console.log(cellsToDelete)
      let rangeToDelete = selectedSheet.getRange(cellsToDelete);
      rangeToDelete.delete(ExcelScript.DeleteShiftDirection.up);
    }
  };

  console.log('Done !')
}