// adding namespace just to silence TS in VSCode

namespace ExcelScript {
    export interface Workbook {
        getActiveWorksheet():any;
        addWorksheet():any;
        getWorksheets(): any
    }
}

// actual code from here

function main(workbook: ExcelScript.Workbook) {

    // Input to adjust

    let sourceSheet = workbook.getActiveWorksheet();
    const NEW_SHEET_NAME = 'result';

    const RANGE = {
        columnStart: 'A',
        columnFinish: 'D',
        rowStart: 1,
        rowFinish: 100000,
    };

    // Code to run 

    // Removing already existing Result worksheet so we don't have to do it manually each time

    let worksheets = workbook.getWorksheets();
    worksheets.forEach((sheet) => {
        if (sheet.getName() === NEW_SHEET_NAME) {
            sheet.delete();
        }
    });

    let result = workbook.addWorksheet();
    result.setName(NEW_SHEET_NAME);
    let resultTotalRange = result.getRange('A1:A1');
    let currentTotalRange = sourceSheet.getRange(`${RANGE.columnStart}${RANGE.rowStart}:${RANGE.columnFinish}${RANGE.rowStart}`)

    // Arrangineg Headers

    resultTotalRange.getCell(0, 1).setValue("Item Code");
    resultTotalRange.getCell(0, 2).setValue("Difference");
    resultTotalRange.getCell(0, 3).setValue("Stock Entry Number");

    let resultLineCounter = 1; 

    for (let i = 1; i < RANGE.rowFinish; i++) {
      if (i % 100 === 0) {
        console.log(`Iteration ${i}`)
      }
      if (currentTotalRange.getCell(i, 0).getText() === "") {
        break
      }
      let currentValuation = Number(Number(currentTotalRange.getCell(i, 2).getText()).toFixed(4));
      let valuationToCompare = Number(Number(currentTotalRange.getCell(i + 1, 2).getText()).toFixed(4));
      if (currentTotalRange.getCell(i, 1).getText() === currentTotalRange.getCell(i + 1, 1).getText()
        && (valuationToCompare - currentValuation) !== 0) {
        resultTotalRange.getCell(resultLineCounter, 1).setValue(currentTotalRange.getCell(i, 1).getText());
        resultTotalRange.getCell(resultLineCounter, 2).setValue(Math.abs(valuationToCompare - currentValuation).toFixed(4));
        resultTotalRange.getCell(resultLineCounter, 3).setValue(currentTotalRange.getCell(i + 1, 3).getText());

        resultLineCounter++;
      }
    }

    console.log('Done !')
}
