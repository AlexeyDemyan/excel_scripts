// Change character in a cell

// amend input parameters here:
const requieredRange = "B1:B27";
const characterToRemove = "3"
// ----------------------------


function replaceString(inputString: string) {

    let newString = "";

    let replaced = false;

    for (let i = 0; i < inputString.length; i++) {
      if (inputString[i] === characterToRemove && replaced === false) {
            newString += "";
            replaced = true;
        }
        else {
            newString += inputString[i]
        }
    }
    return newString;
}

function main(workbook: ExcelScript.Workbook) {
    let selectedSheet = workbook.getActiveWorksheet();

  let range = selectedSheet.getRange(requieredRange);

  let start = Date.now();

    for (let i = 0; i < range.getCellCount(); i++) {
        let result = range.getCell(i, 0).getText();
        
        let currentCell = range.getCell(i, 0);
        currentCell.setValue(replaceString(result));
    };

  let end = Date.now();

  console.log(`Executed in: ${(end - start) * 0.001} seconds`);
}

// Check If Certain Number Comes Immediately After Alphabet Characters

function checkValue(inputString: string) {

    let checked = false;

    for (let i = 0; i < inputString.length; i++) {
        if (checked === false) {
            if (
                inputString[i] === "0" ||
                inputString[i] === "1" ||
                inputString[i] === "2" ||
                inputString[i] === "3" ||
                inputString[i] === "4" ||
                inputString[i] === "5" ||
                inputString[i] === "6" ||
                inputString[i] === "7" ||
                inputString[i] === "8" ||
                inputString[i] === "9" 
            ) {
                checked = true;
                if (inputString[i] !== "3") {
                    console.log(inputString);
                }
            }
        }
    }
}

function main(workbook: ExcelScript.Workbook) {
    let selectedSheet = workbook.getActiveWorksheet();

    let range = selectedSheet.getRange("B5754:B9369");

    for (let i = 0; i < range.getCellCount(); i++) {
        let result = range.getCell(i, 0).getText();
        checkValue(result);
    };

    console.log('Done !')
}

// Look for certain string

// Input to adjust
const CHOSEN_RANGE = "A3:A880";
const STRING_TO_CHECK = "Total"

// Code to run

function checkValue(inputString: string) {

    if (inputString.includes(STRING_TO_CHECK)) {
        console.log(inputString)
    }
}

function main(workbook: ExcelScript.Workbook) {
    let selectedSheet = workbook.getActiveWorksheet();

    let range = selectedSheet.getRange(CHOSEN_RANGE);

    for (let i = 0; i < range.getCellCount(); i++) {
        let result = range.getCell(i, 0).getText();
        checkValue(result);
    };

    console.log('Done !')
}

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
