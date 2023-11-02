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
        columnStart: 'C',
        columnFinish: 'C',
        rowStart: 3,
        rowFinish: 373,
    };

    // Code to run    
    // important to run list consistency check before moving rows

    let totalRange = sourceSheet.getRange(`${RANGE.columnStart}${RANGE.rowStart}:${RANGE.columnStart}${RANGE.rowFinish}`);

    for (let i = 0; i < RANGE.rowFinish; i++) {
        let idNumber: string = totalRange.getCell(i, 0).getText();
        let addressId: string;
        //console.log(idNumber);
        
        for (let j = 0; j < 1400; j++) {
          if (j === 1399) {
            break;
          }
          let cellToCheck: string = totalRange.getCell(j, 6).getText();
          if (idNumber === cellToCheck) {
            //console.log(j);
            //console.log(totalRange.getCell(j, 5).getText());
            addressId = totalRange.getCell(j, 10).getText();
            for (let k = 0; k < 1400; k++) {
              let idCellToCheck: string = totalRange.getCell(k, 12).getText();
              if (addressId === idCellToCheck) {
                //console.log(totalRange.getCell(k, 12).getText());
                let addressToInput = `${totalRange.getCell(k, 13).getText()}, ${totalRange.getCell(k, 14).getText()}`;
                let targetCell = totalRange.getCell(i, 3);
                targetCell.setValue(addressToInput);
                console.log(`value set on iteration ${i}`);
              }
            }
          }
        }
    }

    console.log('All Done!')
}


