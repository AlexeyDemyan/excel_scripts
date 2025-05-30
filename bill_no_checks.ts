function main(workbook: ExcelScript.Workbook) {

    // Input to adjust

    let sourceSheet = workbook.getActiveWorksheet();
    const MATERIAL_RECEIPT = 'Material Receipt';
    const MATERIAL_ISSUE = 'Material Issue';

    const RANGE = {
        columnStart: 'A',
        columnFinish: 'E',
        rowStart: 1,
        rowFinish: 10000,
    };

    const HEADERS = ["Stock Entry Type", "Company", "Item Code (Items)", "Qty (Items)", "Qty as per Stock UOM (Items)", "UOM (Items)", "Stock UOM (Items)", "Conversion Factor (Items)", "Source Warehouse (Items)"];

    const COMPANY = 'Marsovin Winery Ltd (B&V)';
    const DEFAULT_WAREHOUSE = 'RM (PET) - M'

    // Code to run 

    // Removing already existing Result worksheet so we don't have to do it manually each time

    let worksheets = workbook.getWorksheets();
    worksheets.forEach((sheet) => {
        if (sheet.getName() === MATERIAL_RECEIPT || sheet.getName() === MATERIAL_ISSUE) {
            sheet.delete();
        }
    });

    // Creating new sheets

    let materialReceiptSheet = workbook.addWorksheet();
    materialReceiptSheet.setName(MATERIAL_RECEIPT);
    let materialReceiptSheetTotalRange = materialReceiptSheet.getRange('A1:A1');

    let materialIssueSheet = workbook.addWorksheet();
    materialIssueSheet.setName(MATERIAL_ISSUE);
    let materialIssueSheetTotalRange = materialIssueSheet.getRange('A1:A1');

    let currentTotalRange = sourceSheet.getRange(`${RANGE.columnStart}${RANGE.rowStart}:${RANGE.columnFinish}${RANGE.rowStart}`);

    // Arranging Headers

    for (let i = 0; i < HEADERS.length; i++) {
        materialReceiptSheetTotalRange.getCell(0, i).setValue(HEADERS[i]);
        materialIssueSheetTotalRange.getCell(0, i).setValue(HEADERS[i]);
        if (i === 8) {
            materialReceiptSheetTotalRange.getCell(0, i).setValue("Target Warehouse (Items)");
        }
    };

    materialReceiptSheetTotalRange.getCell(1, 0).setValue(MATERIAL_RECEIPT);
    materialReceiptSheetTotalRange.getCell(1, 1).setValue(COMPANY);

    materialIssueSheetTotalRange.getCell(1, 0).setValue(MATERIAL_ISSUE);
    materialIssueSheetTotalRange.getCell(1, 1).setValue(COMPANY);

    // Main Code

    let materialReceiptLineCounter = 1;
    let materialIssueLineCounter = 1;

    for (let i = 0; i < RANGE.rowFinish; i++) {
        if (currentTotalRange.getCell(i, 0).getText() === "") {
            console.log("Finished!")
            break;
        }
        let currentItem = currentTotalRange.getCell(i, 0).getText();
        let currentAmount = currentTotalRange.getCell(i, 4).getText();

        console.log(`Checking item ${currentItem}`)
        let itemUOM = currentTotalRange.getCell(i, 2).getText();
        let itemAmount = currentTotalRange.getCell(i, 3).getText();

        let absoluteDifference = Math.abs(Number(currentAmount) - Number(itemAmount));

        if (Number(currentAmount) > Number(itemAmount)) {
            materialReceiptSheetTotalRange.getCell(materialReceiptLineCounter, 2).setValue(currentItem);
            materialReceiptSheetTotalRange.getCell(materialReceiptLineCounter, 3).setValue(absoluteDifference);
            materialReceiptSheetTotalRange.getCell(materialReceiptLineCounter, 4).setValue(absoluteDifference);
            materialReceiptSheetTotalRange.getCell(materialReceiptLineCounter, 5).setValue(itemUOM);
            materialReceiptSheetTotalRange.getCell(materialReceiptLineCounter, 6).setValue(itemUOM);
            materialReceiptSheetTotalRange.getCell(materialReceiptLineCounter, 7).setValue(1);
            materialReceiptSheetTotalRange.getCell(materialReceiptLineCounter, 8).setValue(DEFAULT_WAREHOUSE);
            materialReceiptLineCounter++
        }

        if (Number(currentAmount) < Number(itemAmount)) {
            materialIssueSheetTotalRange.getCell(materialIssueLineCounter, 2).setValue(currentItem);
            materialIssueSheetTotalRange.getCell(materialIssueLineCounter, 3).setValue(absoluteDifference);
            materialIssueSheetTotalRange.getCell(materialIssueLineCounter, 4).setValue(absoluteDifference);
            materialIssueSheetTotalRange.getCell(materialIssueLineCounter, 5).setValue(itemUOM);
            materialIssueSheetTotalRange.getCell(materialIssueLineCounter, 6).setValue(itemUOM);
            materialIssueSheetTotalRange.getCell(materialIssueLineCounter, 7).setValue(1);
            materialIssueSheetTotalRange.getCell(materialIssueLineCounter, 8).setValue(DEFAULT_WAREHOUSE);
            materialIssueLineCounter++
        }
    }

    console.log('Done !')
}
