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
      columnFinish: 'J',
      rowStart: 1,
      rowFinish: 10,
    };

    let EMPTY_CELLS_TO_VERIFY_THAT_RANGE_FINISHED = 10;

    // Code to run    
    // important to run list consistency check before moving rows
   
    let sourceTotalRange = sourceSheet.getRange(`${RANGE.columnStart}${RANGE.rowStart}:${RANGE.columnStart}${RANGE.rowFinish}`)
  
    let start = Date.now();

    let counter = 0;
    let i = 0;
    let currentCell = sourceTotalRange.getCell(i, 0);

    while (counter < EMPTY_CELLS_TO_VERIFY_THAT_RANGE_FINISHED) {
        currentCell = sourceTotalRange.getCell(i, 0);
        let currentCellText = currentCell.getText();
        i++;

        if (currentCellText !== '') {
            counter = 0;
        }

        if (currentCellText === '') {
            counter++;
        }
    }
    
    RANGE.rowFinish = i;
    console.log(RANGE.rowFinish);

    counter = 0;
    i = 0;

    let end = Date.now();
  
    console.log(`Executed in: ${(end - start) * 0.001} seconds`);
  }