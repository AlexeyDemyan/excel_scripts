function main(workbook: ExcelScript.Workbook) {

    // Input to adjust
  
    let sourceSheet = workbook.getActiveWorksheet();
    const NEW_SHEET_NAME = 'Result';
    const RANGE = {
      columnStart: 'A',
      columnFinish: 'T',
      rowStart: 3,
      rowFinish: 65,
    };
  
    const ORDERED_LIST: Array<string> = ['item name 1', 'item name 2', 'item name 3'];
  
    // Code to run    
    // important to run list consistency check before moving rows
  
    let result = workbook.addWorksheet();
    result.setName(NEW_SHEET_NAME);
  
    let sourceTotalRange = sourceSheet.getRange(`${RANGE.columnStart}${RANGE.rowStart}:${RANGE.columnStart}${RANGE.rowFinish}`)
  
    let counter = 0;
    ORDERED_LIST.forEach(element => {
      for (let i = 0; i < sourceTotalRange.getCellCount(); i++) {
        let currentCell = sourceTotalRange.getCell(i, 0);
        let currentCellText = currentCell.getText();
        if (currentCellText === element) {
          let rangeToMove = `${RANGE.columnStart}${currentCell.getRowIndex() + 1}:${RANGE.columnFinish}${currentCell.getRowIndex() + 1}`;
          let sourceRange = sourceSheet.getRange(rangeToMove);
          let resultRange = result.getRange(`${RANGE.columnStart}${RANGE.rowStart + counter}:${RANGE.columnFinish}${RANGE.rowStart + counter}`);
          resultRange.setValues(sourceRange.getValues());
          counter++;
        }
        
      }
    })
  
    console.log('Done !')
  }