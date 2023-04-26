// Look for certain string

// Input to adjust
const CHOSEN_RANGE = "A3:A880";
const STRING_TO_CHECK = "Total"

// Code to run

function checkValue(inputString: string) {

    if (inputString.includes(STRING_TO_CHECK)) {
        console.log(inputString)
    }
}

function main(workbook: ExcelScript.Workbook) {
    let selectedSheet = workbook.getActiveWorksheet();

    let range = selectedSheet.getRange(CHOSEN_RANGE);

    for (let i = 0; i < range.getCellCount(); i++) {
        let result = range.getCell(i, 0).getText();
        checkValue(result);
    };

    console.log('Done !')
}
