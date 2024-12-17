function main(workbook: ExcelScript.Workbook) {

    // Input to adjust

    let sourceSheet = workbook.getActiveWorksheet();
    const NEW_SHEET_NAME = 'Journal Entry';

    const RANGE = {
        columnStart: 'A',
        columnFinish: 'H',
        rowStart: 1,
        rowFinish: 100,
    };

    const HEADERS = ["Entry Type", "Posting Date", "Account (Accounting Entries)", "Debit (Accounting Entries)", "Credit (Accounting Entries)", "User Remark (Accounting Entries)"];

    const ENTRY_TYPE = 'Journal Entry';
    const POSTING_DATE = new Date();
    const DEBIT_ACCOUNT = 'Account 1';
    const CREDIT_ACCOUNT = 'Account 2';

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

    resultTotalRange.getCell(1, 0).setValue(ENTRY_TYPE);
    resultTotalRange.getCell(1, 1).setValue(`${POSTING_DATE.getDate()}/${POSTING_DATE.getMonth() + 1}/${POSTING_DATE.getFullYear()}`);

    // Main Part

    let rowCount = 1;

    for (let i = 1; i < RANGE.rowFinish; i++) {
        let currentCell = currentTotalRange.getCell(i, 0);
        let currentCellText = currentCell.getText();
        let currentCreditCellAmount = currentTotalRange.getCell(i, 6).getValue();
        let description = currentTotalRange.getCell(i, 4).getValue();
        if (currentCellText === "") {
            break
        }

        let targetAccountCell = resultTotalRange.getCell(rowCount, 2);
        targetAccountCell.setValue(DEBIT_ACCOUNT)

        let targetDebitCell = resultTotalRange.getCell(rowCount, 3);
        targetDebitCell.setValue(currentCreditCellAmount);

        let targetRemarkLine = resultTotalRange.getCell(rowCount, 5);
        targetRemarkLine.setValue(description);

        rowCount++;

        targetAccountCell = resultTotalRange.getCell(rowCount, 2);
        targetAccountCell.setValue(CREDIT_ACCOUNT);

        let targetCreditCell = resultTotalRange.getCell(rowCount, 4);
        targetCreditCell.setValue(currentCreditCellAmount);

        targetRemarkLine = resultTotalRange.getCell(rowCount, 5);
        targetRemarkLine.setValue(description);

        rowCount++;
    }

    console.log('Done !')
}