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
    columnStart: 'A',
    columnFinish: 'I',
    rowStart: 1,
    rowFinish: 1100,
  };

  // Code to run    

  let totalRange = sourceSheet.getRange(`${RANGE.columnStart}${RANGE.rowStart}:${RANGE.columnStart}${RANGE.rowFinish}`);

  for (let i = 1; i < 125; i++) {
    console.log(`checking ${i}`)
    let arrayToCheck: String[] = [];

    arrayToCheck.push(totalRange.getCell(i, 0).getText());
    arrayToCheck.push(totalRange.getCell(i, 1).getText());
    arrayToCheck.push(totalRange.getCell(i, 2).getText());
    arrayToCheck.push(totalRange.getCell(i, 3).getText());

    for (let j = 1; j < 125; j++) {
      let arrayToCheckAgainst: String[] = [];

      arrayToCheckAgainst.push(totalRange.getCell(j, 5).getText());
      arrayToCheckAgainst.push(totalRange.getCell(j, 6).getText());
      arrayToCheckAgainst.push(totalRange.getCell(j, 7).getText());
      arrayToCheckAgainst.push(totalRange.getCell(j, 8).getText());

      if (arrayToCheck[0] === arrayToCheckAgainst[0] &&
        arrayToCheck[1] === arrayToCheckAgainst[1] &&
        arrayToCheck[2] === arrayToCheckAgainst[2] &&
        arrayToCheck[3] === arrayToCheckAgainst[3]) {

        totalRange.getCell(i, 0).setValue('');
        totalRange.getCell(i, 1).setValue('');
        totalRange.getCell(i, 2).setValue('');
        totalRange.getCell(i, 3).setValue('');

        totalRange.getCell(j, 5).setValue('');
        totalRange.getCell(j, 6).setValue('');
        totalRange.getCell(j, 7).setValue('');
        totalRange.getCell(j, 8).setValue('');
        break
      }
    }
  }

  console.log('All Done!');
}