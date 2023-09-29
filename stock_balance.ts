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
    const RESULT_SHEET_NAME = 'Result';
    const RANGE = {
      columnStart: 'A',
      columnFinish: 'AB',
      rowStart: 1,
      rowFinish: 10,
    };

    const COLUMNS_DICTIONARY = {
        'Item': 'Item',
        'Balance Qty': 'Balance Qty'
    };

    let balanceQtyColumn = 0;

    // Removing already existing Result worksheet so we don't have to do it manually each time

    let worksheets = workbook.getWorksheets();
    worksheets.forEach((sheet) => {
        if (sheet.getName() === RESULT_SHEET_NAME) {
            sheet.delete();
        }
    });
  
    const ORDERED_LIST: Array<string> = ['Item', 'mock', 'mock'];

    const EMPTY_CELLS_TO_VERIFY_THAT_RANGE_FINISHED = 10;
  
    // Code to run    
    // important to run list consistency check before moving rows
  
    let result = workbook.addWorksheet();
    result.setName(RESULT_SHEET_NAME);
  
    let sourceTotalRange = sourceSheet.getRange(`${RANGE.columnStart}${RANGE.rowStart}:${RANGE.columnStart}${RANGE.rowFinish}`)
  
    let start = Date.now();

    let counter = 0;
    let i = 0;
    let currentCell = sourceTotalRange.getCell(i, 0);
    console.log(`checking: ${currentCell.getText()}`);

    // Establishing Total Rows Range

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
    sourceTotalRange = sourceSheet.getRange(`${RANGE.columnStart}${RANGE.rowStart}:${RANGE.columnStart}${RANGE.rowFinish}`)

    console.log(`Range is verified, rows established: ${i}`)

    counter = 0;
    i = 0;

    // Establishing Total Columns Range

    while (counter < EMPTY_CELLS_TO_VERIFY_THAT_RANGE_FINISHED) {
        currentCell = sourceTotalRange.getCell(0, i);
        let currentCellText = currentCell.getText();
        console.log(currentCellText);
        i++;

        if (currentCellText === COLUMNS_DICTIONARY['Balance Qty']) {
            balanceQtyColumn = i;
        }

        if (currentCellText !== '') {
            counter = 0;
        }

        if (currentCellText === '') {
            counter++;
        }
    }

    counter = 0;
    i = 0;

    ORDERED_LIST.forEach(element => {
        // let sumOfItemBalance = 0;

      for (let j = 0; j < sourceTotalRange.getCellCount(); j++) {
        let currentCell = sourceTotalRange.getCell(j, 0);
        let currentCellText = currentCell.getText();

        if (currentCellText === element) {
          let rangeToMove = `${RANGE.columnStart}${currentCell.getRowIndex() + 1}:${RANGE.columnFinish}${currentCell.getRowIndex() + 1}`;
          let sourceRange = sourceSheet.getRange(rangeToMove);
          let resultRange = result.getRange(`${RANGE.columnStart}${RANGE.rowStart + counter}:${RANGE.columnFinish}${RANGE.rowStart + counter}`);
          resultRange.setValues(sourceRange.getValues());

          let balanceQtyCell = resultRange.getCell(i, balanceQtyColumn);
        
          console.log(balanceQtyCell.getText());

          counter++;
        }
        
      }

    console.log('finished one item');
    counter++;
      
    })

    let end = Date.now();
  
    console.log(`Executed in: ${(end - start) * 0.001} seconds`);
  }