function main(workbook: ExcelScript.Workbook) {

    // Input to adjust

    let sourceSheet = workbook.getActiveWorksheet();
    const NEW_SHEET_NAME = 'result';

    const RANGE = {
        columnStart: 'A',
        columnFinish: 'D',
        rowStart: 1,
        rowFinish: 14300,
    };

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

    // Preparing array of journal entries

    let resultingList: String[] = [];

    for (let i = 1; i < RANGE.rowFinish; i++) {        
        if (currentTotalRange.getCell(i, 3).getText() === "") {
            break
        }

        let currentMultipleLinesJournalNumber = currentTotalRange.getCell(i,3).getText();
        resultingList.push(currentMultipleLinesJournalNumber);
    }

    console.log(resultingList);

    let resultLineCounter = 0;

    for (let i = 1; i < RANGE.rowFinish; i++) {

        if (i % 1000 === 0) {
            console.log(`Iteration ${i}`)
        }

        let currentJourlanNumberToCheck = currentTotalRange.getCell(i, 0).getText();
        if (resultingList.includes(currentJourlanNumberToCheck)) {
            resultTotalRange.getCell(resultLineCounter, 0).setValue(currentTotalRange.getCell(i, 0).getText());
            resultTotalRange.getCell(resultLineCounter, 1).setValue(currentTotalRange.getCell(i, 1).getText());
            resultTotalRange.getCell(resultLineCounter, 2).setValue(currentTotalRange.getCell(i, 2).getText());
            resultLineCounter++;
        }
    }

    console.log('Done !')
}