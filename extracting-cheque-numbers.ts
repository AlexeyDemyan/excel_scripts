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
    const RANGE = {
        columnStart: 'E',
        columnFinish: 'E',
        rowStart: 2,
        rowFinish: 1500,
    };

    // Code to run    
    // Currently assuming cheque number always has 5 digits only

    let totalRange = sourceSheet.getRange(`${RANGE.columnStart}${RANGE.rowStart}:${RANGE.columnStart}${RANGE.rowFinish}`);

    // extracting cheque number

    for (let i = 0; i < RANGE.rowFinish; i++) {
        let currentCell = totalRange.getCell(i, 0);
        let currentCellText: string = currentCell.getText();
        let result: string[] = [];
        let currentCellTextArrayed = Array.from(currentCellText);
        for (let j = 0; j < currentCellTextArrayed.length; j++) {
          let resultingValue = currentCellTextArrayed.slice(j, j + 5).join("");
          if (Number(resultingValue) && Number(resultingValue).toString().length === 5) {
            console.log(resultingValue);
            let targetCell = totalRange.getCell(i, 1);
            targetCell.setValue(resultingValue);
          }
        }
    }

    console.log('All Done!')
}
