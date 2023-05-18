function main(workbook: ExcelScript.Workbook) {

    // Input to adjust

    const CHOSEN_RANGE = 'A3:A60';

    // Code to run

    let selectedSheet = workbook.getActiveWorksheet();
    let range = selectedSheet.getRange(CHOSEN_RANGE);

    let storedItems: Array<string> = [];

    for (let i = 0; i < range.getCellCount(); i++) {
        let currentCell = range.getCell(i, 0);
        let currentCellText = currentCell.getText();
        if (storedItems[storedItems.length - 1] !== currentCellText) {
            if (storedItems.includes(currentCellText)) {
                console.log(`duplication detected - ${currentCellText}`)
            }
            storedItems.push(currentCellText);
        }
    }
    console.log(storedItems)
    console.log('Done !')
}