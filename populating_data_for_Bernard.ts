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
        rowStart: 1,
        rowFinish: 2400,
    };

    const currentTotalRange = sourceSheet.getRange(`${RANGE.columnStart}${RANGE.rowStart}:${RANGE.columnFinish}${RANGE.rowStart}`);

    let currentItem = "";
    let currentName = "";

    for (let i = 1; i < RANGE.rowFinish; i++) {
        let itemCell = currentTotalRange.getCell(i,0);
        let nameCell = currentTotalRange.getCell(i, 2);
        let controlCell = currentTotalRange.getCell(i, 4);

        // Old code:
        // if (itemCell.getText() !== "") {
        //     currentItem = itemCell.getText();
        // }

        // Refactor 1:
        //currentItem = itemCell.getText() ? itemCell.getText() : currentItem;

        // Refactor 2:
        currentItem = itemCell.getText() || currentItem;
        currentName = nameCell.getText() || currentName;

        if (controlCell.getText() !== "") {
            itemCell.setValue(currentItem);
            nameCell.setValue(currentName);
        }

        if ((i % 1000) === 0) {
            console.log(i);
        }
    }

    console.log('All Done!');
}
