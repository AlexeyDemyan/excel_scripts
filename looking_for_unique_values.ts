// adding namespace just to silence TS in VSCode

namespace ExcelScript {
  export interface Workbook {
      getActiveWorksheet():any;
      addWorksheet():any;
  }
}

// actual code from here

function main(workbook: ExcelScript.Workbook) {

  // Input to adjust

  let sourceSheet = workbook.getActiveWorksheet();
  const RANGE = {
    columnStart: 'B',
    columnFinish: 'G',
    rowStart: 1,
    rowFinish: 1050,
  };

  // Code to run    

  let totalRange = sourceSheet.getRange(`${RANGE.columnStart}${RANGE.rowStart}:${RANGE.columnStart}${RANGE.rowFinish}`);

  let arrayToCheck: String[] = [];

  console.log(totalRange.getCell(1, 5).getText());

  for (let i = 1; i < RANGE.rowFinish; i++) {
    let currentCell = totalRange.getCell(i, 0);
    let currentCellText: string = currentCell.getText();
    if (currentCellText === "") { break }
    arrayToCheck.push(currentCellText)
  }

  console.log(arrayToCheck.length);

  arrayToCheck.forEach(value => {
    if(arrayToCheck.indexOf(value) === arrayToCheck.lastIndexOf(value)) {
      totalRange.getCell(arrayToCheck.indexOf(value) + 1, 5).setValue('Unique');
    }
  })

  console.log('All Done!');
}