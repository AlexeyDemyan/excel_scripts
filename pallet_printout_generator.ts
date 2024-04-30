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
        columnStart: 'B',
        columnFinish: 'C',
        rowStart: 1,
        rowFinish: 2500,
    };

    const SETTINGS = {
        product: '=Settings!C2',
        batch: '=Settings!C3',
        release: '=Settings!C4'
    }

    // Code to run    

    let totalRange = sourceSheet.getRange(`${RANGE.columnStart}${RANGE.rowStart}:${RANGE.columnStart}${RANGE.rowFinish}`);

    let lineNumber = 0;
    let palletNumber = 1;

    while(lineNumber < RANGE.rowFinish) {
        for (let j = 1; j <= 15; j++) {

            if ([2,7,12].includes(j)) {
                totalRange.getCell(lineNumber,0).setValue('PRODUCT:')
                totalRange.getCell(lineNumber, 1).setValue(SETTINGS.product)
                };
            if([3,8,13].includes(j)) {
                totalRange.getCell(lineNumber,0).setValue('PALLET NO:')
                totalRange.getCell(lineNumber, 1).setValue(palletNumber)
                };
            if([4,9,14].includes(j)) {
                totalRange.getCell(lineNumber,0).setValue('BATCH:')
                totalRange.getCell(lineNumber, 1).setValue(SETTINGS.batch)
                };
            if([5,10,15].includes(j)) {
                totalRange.getCell(lineNumber,0).setValue('RELEASE:')
                totalRange.getCell(lineNumber, 1).setValue(SETTINGS.release)
                };

            if (j === 6) {totalRange.getRow(lineNumber).getFormat().setRowHeight(112.5)}
            if (j === 11) {totalRange.getRow(lineNumber).getFormat().setRowHeight(60)}

            lineNumber++;
        }

        palletNumber++;
    }

    console.log('All Done!')
}