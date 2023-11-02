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
        columnFinish: 'B',
        rowStart: 3,
        rowFinish: 373,
    };

    // Code to run    
    // important to run list consistency check before moving rows

    let totalRange = sourceSheet.getRange(`${RANGE.columnStart}${RANGE.rowStart}:${RANGE.columnStart}${RANGE.rowFinish}`);

    // truncating ID numbers

    for (let i = 0; i < RANGE.rowFinish; i++) {
        let currentCell = totalRange.getCell(i, 0);
        let currentCellText: string = currentCell.getText();
        let result: string[] = [];
        Array.from(currentCellText).forEach((character) => {
            if (character.toLowerCase() === 'm' || character.toLowerCase() === 'g') {
                character = ''
            }
            result.push(character);
        })

        // moving truncated ID numbers to empty cells on the right:

        let targetCell = totalRange.getCell(i, 1);
        targetCell.setValue(result.join(''));
    }

    console.log('All Done!')
}
