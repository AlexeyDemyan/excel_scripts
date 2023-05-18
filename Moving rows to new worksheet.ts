function main(workbook: ExcelScript.Workbook) {

    // Input to adjust

    let sourceSheet = workbook.getActiveWorksheet();
    const NEW_SHEET_NAME = 'Result';
    const RANGE = {
      columnStart: 'A',
      columnFinish: 'T',
      rowStart: 3,
      rowFinish: 10,
    };

    // Code to run    

    // important to run list consistency check before moving rows

    let rangeToMove = sourceSheet.getRange(`${RANGE.columnStart}${RANGE.rowStart}:${RANGE.columnFinish}${RANGE.rowStart}`);
     let result = workbook.addWorksheet();
     result.setName(NEW_SHEET_NAME);

    const resultRange = result.getRange(`${RANGE.columnStart}${RANGE.rowStart}:${RANGE.columnFinish}${RANGE.rowStart}`);

     resultRange.setValues(rangeToMove.getValues());

    let sourceTotalRange = sourceSheet.getRange(`${RANGE.columnStart}${RANGE.rowStart}:${RANGE.columnStart}${RANGE.rowFinish}`)

    for (let i = 0; i < sourceTotalRange.getCellCount(); i++) {  
      let rangeToMove = `${RANGE.columnStart}${RANGE.rowStart + i}:${RANGE.columnFinish}${RANGE.rowStart + 1}`;
      let sourceRange = sourceSheet.getRange(rangeToMove);
      let resultRange = result.getRange(rangeToMove);

      resultRange.setValues(sourceRange.getValues());
    }

    console.log('Done !')
}