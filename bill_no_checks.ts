function main(workbook: ExcelScript.Workbook) {

    // Input to adjust

    let sourceSheet = workbook.getActiveWorksheet();

    const RESULT = 'result';

    const RANGE = {
        columnStart: 'A',
        columnFinish: 'C',
        rowStart: 1,
        rowFinish: 4000,
    };

    // Code to run 

    // Removing already existing Result worksheet so we don't have to do it manually each time

    let worksheets = workbook.getWorksheets();
    worksheets.forEach((sheet) => {
        if (sheet.getName() === RESULT) {
            sheet.delete();
        }
    });

    // Creating new sheets

    let resultSheet = workbook.addWorksheet();
    resultSheet.setName(RESULT);
    let resultSheetTotalRange = resultSheet.getRange('A1:A1');

    let currentTotalRange = sourceSheet.getRange(`${RANGE.columnStart}${RANGE.rowStart}:${RANGE.columnFinish}${RANGE.rowStart}`);

    // Main Code

    let billNoToCheckAgainst = '69';
    let possibleDuplicateInvoiceNo = '';
    let resultLineCount = 1;
    let partyList: String[] = [];

    for (let i = 0; i < RANGE.rowFinish; i++) {
        console.log(`checking ${i}`);
        if (currentTotalRange.getCell(i, 0).getText() === "") {
            console.log(`Script completed at ${i} iterations`);
            break;
        }

        let currentBillNo = currentTotalRange.getCell(i, 0).getText();
        let currentInvoiceNo = currentTotalRange.getCell(i, 1).getText();
        let currentParty = currentTotalRange.getCell(i, 2).getText();

        if (currentBillNo !== billNoToCheckAgainst) {
            billNoToCheckAgainst = currentBillNo;
            partyList = [];
            partyList.push(currentParty);
        } else {
            if (partyList.includes(currentParty)) {
                if (resultSheetTotalRange.getCell(resultLineCount - 1, 0).getValue() !== currentBillNo) {
                    resultLineCount++;
                }

                if (resultSheetTotalRange.getCell(resultLineCount - 1, 1).getValue() !== possibleDuplicateInvoiceNo) {   
                resultSheetTotalRange.getCell(resultLineCount, 0).setValue(currentBillNo);
                resultSheetTotalRange.getCell(resultLineCount, 1).setValue(possibleDuplicateInvoiceNo);
                resultSheetTotalRange.getCell(resultLineCount, 2).setValue(currentParty);
                resultLineCount++;
                }
                resultSheetTotalRange.getCell(resultLineCount, 0).setValue(currentBillNo);
                resultSheetTotalRange.getCell(resultLineCount, 1).setValue(currentInvoiceNo);
                resultSheetTotalRange.getCell(resultLineCount, 2).setValue(currentParty);
                resultLineCount++;
            } else {
                partyList.push(currentParty);
            }
        }
        possibleDuplicateInvoiceNo = currentInvoiceNo;

    }

    console.log('Done !')
}
