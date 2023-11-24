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
    const NEW_SHEET_NAME = 'result';

    const RANGE = {
        columnStart: 'G',
        columnFinish: 'AV',
        rowStart: 1,
        rowFinish: 80,
    };

    const HEADERS = ['', ''];

    const UOMS = {
        METRIC_TON: 'METRIC TON',
        GRAMMES: 'GRAMMES',
    }

    const COMPANY = '';

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

    let resultLineCounter = 1;

    for (let i = 8; i < RANGE.rowFinish; i++) {
        if (currentTotalRange.getCell(i, 0).getText() !== "") {

            let theoreticalHL = Number(currentTotalRange.getCell(i, 4).getText());

            let wineCode = currentTotalRange.getCell(i, 2).getText();
            resultTotalRange.getCell(resultLineCounter, 0).setValue(wineCode);
            resultTotalRange.getCell(resultLineCounter, 1).setValue(COMPANY);
            resultTotalRange.getCell(resultLineCounter, 2).setValue(1);

            let grapeCode = currentTotalRange.getCell(i, 0).getText();
            resultTotalRange.getCell(resultLineCounter, 3).setValue(grapeCode);
            let grapeQty = Number(currentTotalRange.getCell(i, 1).getText()) / theoreticalHL * 0.10;
            resultTotalRange.getCell(resultLineCounter, 4).setValue(grapeQty);
            resultTotalRange.getCell(resultLineCounter, 5).setValue(UOMS.METRIC_TON);

            resultLineCounter++;

            for (let j = 5; j < 100; j++) {
                let itemQty = Number(currentTotalRange.getCell(i, j).getText()) / theoreticalHL;
                if (itemQty > 0) {
                    resultTotalRange.getCell(resultLineCounter, 4).setValue(itemQty);
                    let itemCode = currentTotalRange.getCell(1, j).getText();
                    resultTotalRange.getCell(resultLineCounter, 3).setValue(itemCode);
                    resultTotalRange.getCell(resultLineCounter, 5).setValue(UOMS.GRAMMES);
                    resultLineCounter++;
                }
            }            
        }
    }

    console.log('Done !')
}
