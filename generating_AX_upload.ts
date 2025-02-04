function main(workbook: ExcelScript.Workbook) {

    // Input to adjust

    let sourceSheet = workbook.getActiveWorksheet();
    const NEW_SHEET_NAME = 'UPLOAD';

    const RANGE = {
        columnStart: 'A',
        columnFinish: 'F',
        rowStart: 1,
        rowFinish: 1000,
    };

    const HEADERS = ['JournalName', 'JournalNum', 'LineNum', 'CurrencyCode', 'TransDate', 'AccountType', 'Voucher', 'Txt', 'LedgerDimension', 'AmountCurDebit', 'AmountCurCredit'];

    const CURRENCY_CODE = 'EUR';
    const JOURNAL_NAME = 'GL_Journal';
    const ACCOUNT_TYPE = 'Ledger';
    const LEDGER_DIMENSION_DEBIT = '999101';
    const LEDGER_DIMENSION_CREDIT = '999006';
    const SEQUENCE_NAME = 'MRMUPLD';
    const SEQUENCE_DIGIT_COUNT = 3;

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

    // Moving data

    function moveColumn(column: string, destination: string) {
        let rangeToMove = sourceSheet.getRange(`${column}${2}:${column}${RANGE.rowFinish}`);
        let destinationRange = result.getRange(`${destination}${2}:${destination}${RANGE.rowFinish}`);
        destinationRange.setValues(rangeToMove.getValues());
    }

    // moving item codes and quantities

    moveColumn('D', 'J');
    moveColumn('E', 'K');
    moveColumn('F', 'H');

    const currentTime = new Date();
    const currentYear = (currentTime.getFullYear().toString());
    const currentMonth = ((currentTime.getMonth() + 1).toString().padStart(2, '0'));
    const currentDate = (currentTime.getDate().toString().padStart(2, '0'));

    const currentHours = (currentTime.getHours().toString().padStart(2, '0'));
    const currentMinutes = (currentTime.getMinutes().toString().padStart(2, '0'));
    const currentSeconds = (currentTime.getSeconds().toString().padStart(2, '0'));

    const currentTimeStamp = currentYear.concat(currentMonth, currentDate, currentHours, currentMinutes, currentSeconds);

    let sequenceNumber = 1;

    for (let i = 1; i < RANGE.rowFinish; i++) {
        let txtCellValue = resultTotalRange.getCell(i, 7).getText();
        if (txtCellValue === '') {
            break
        }

        resultTotalRange.getCell(i, 0).setValue(JOURNAL_NAME);
        resultTotalRange.getCell(i, 2).setValue(i);
        resultTotalRange.getCell(i, 3).setValue(CURRENCY_CODE);
        resultTotalRange.getCell(i, 5).setValue(ACCOUNT_TYPE);
        if ((i % 2) !== 0) {
            resultTotalRange.getCell(i, 8).setValue(LEDGER_DIMENSION_DEBIT);
            resultTotalRange.getCell(i, 6).setValue(SEQUENCE_NAME + currentTimeStamp + sequenceNumber.toString().padStart(SEQUENCE_DIGIT_COUNT, '0'));
        } else {
            resultTotalRange.getCell(i, 8).setValue(LEDGER_DIMENSION_CREDIT);
            resultTotalRange.getCell(i, 6).setValue(SEQUENCE_NAME + currentTimeStamp + sequenceNumber.toString().padStart(SEQUENCE_DIGIT_COUNT, '0'));
            sequenceNumber++;
        }
    }

    console.log('Done !')
}