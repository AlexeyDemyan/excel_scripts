// Check If Certain Number Comes Immediately After Alphabet Characters

function checkValue(inputString: string) {

  let checked = false;

  for (let i = 0; i < inputString.length; i++) {
      if (checked === false) {
          if (
              inputString[i] === "0" ||
              inputString[i] === "1" ||
              inputString[i] === "2" ||
              inputString[i] === "3" ||
              inputString[i] === "4" ||
              inputString[i] === "5" ||
              inputString[i] === "6" ||
              inputString[i] === "7" ||
              inputString[i] === "8" ||
              inputString[i] === "9" 
          ) {
              checked = true;
              if (inputString[i] !== "3") {
                  console.log(inputString);
              }
          }
      }
  }
}

function main(workbook: ExcelScript.Workbook) {
  let selectedSheet = workbook.getActiveWorksheet();

  let range = selectedSheet.getRange("B5754:B9369");

  for (let i = 0; i < range.getCellCount(); i++) {
      let result = range.getCell(i, 0).getText();
      checkValue(result);
  };

  console.log('Done !')
}