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
    const NEW_SHEET_NAME = '';
  
    const RANGE = {
      columnStart: 'A',
      columnFinish: 'U',
      rowStart: 1,
      rowFinish: 100,
    };
  
    const HEADERS = ['',''];
  
    const STOCK_ENTRY_TYPE = '';
    const COMPANY = '';
    const DEFAULT_UOM = '';
    const DEFAULT_SOURCE_WAREHOUSE = ''
  
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
  
    for (let i = 0; i < HEADERS.length; i++) {
      resultTotalRange.getCell(0, i).setValue(HEADERS[i]);
    };
  
    resultTotalRange.getCell(1, 0).setValue(STOCK_ENTRY_TYPE);
    resultTotalRange.getCell(1, 1).setValue(COMPANY);
  
    // Moving data
  
    function moveColumn(column: string, destination: string) {
      let rangeToMove = sourceSheet.getRange(`${column}${2}:${column}${RANGE.rowFinish}`);
      let destinationRange = result.getRange(`${destination}${2}:${destination}${RANGE.rowFinish}`);
      destinationRange.setValues(rangeToMove.getValues());
    }
    
    // moving item codes and quantities
  
    moveColumn('B', 'C');
    moveColumn('J', 'D');
  
    
    
    for (let i = 1; i < RANGE.rowFinish; i++) {
      let quantity = resultTotalRange.getCell(i, 3).getText();
      if (quantity === '') {
        break
      }
      // converting quantities to dozens
      resultTotalRange.getCell(i, 3).setValue(Number(quantity)/12);
  
      // copying quantities to column on the right
      resultTotalRange.getCell(i, 4).setValue(quantity);
  
      // entering DOZENS as UOMs in both columns
      resultTotalRange.getCell(i, 5).setValue(DEFAULT_UOM);
      resultTotalRange.getCell(i, 6).setValue(DEFAULT_UOM);
  
      // entering conversion factor
      resultTotalRange.getCell(i, 7).setValue(1);
  
      // entering default source warehouse
      resultTotalRange.getCell(i, 8).setValue(DEFAULT_SOURCE_WAREHOUSE);
    }
  
    console.log('Done !')
  }