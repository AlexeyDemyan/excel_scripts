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