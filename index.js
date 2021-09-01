const Excel = require("exceljs");
const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");

// const toBeCopy = "About_Me_Dtls1";

// const copiedVersion = "Copied_Version";

const toBeCopy = "EX-DATA_CHECKS2";
const copiedVersion = "EX-DATA_ISSUES";
// Using a function to set default app path
function getDir() {
  if (process.pkg) {
    return path.resolve(process.execPath + "/..");
  } else {
    return path.join(require.main ? require.main.path : process.cwd());
  }
}

const currentDir = getDir();
/**
 *return void 
 just console the the notification message file has been changed 
 */

(async () => {
  let files = fs.readdirSync(`${currentDir}/My_Data`);
  if (!files) {
    console.log(err);
    throw new Error("No file found");
  }
  files = files?.filter((value) => !value.startsWith(".~"));
  let filePath = `${currentDir}/My_Data/${files[0]}`;
  // let SavedFilePath = `${currentDir}/My_Data/${files[0]}`;

  // if (filePath.indexOf(".xlsm") !== -1) {
  //   foundXLSM = true;
  // }
  /**
   * create the work book instance for target
   */
  let targetWorkbook = new Excel.Workbook();

  /**
   * provide the target file with hardcoded name
   */
  targetWorkbook = await targetWorkbook.xlsx.readFile(filePath);

  /**
   * provide the worksheet on which new data copy
   */
  const targetWorksheet = targetWorkbook.getWorksheet(copiedVersion);

  // you can add new sheet as well.
  //   here Sheet1 is hardcoded as you allow but but we attact from like that
  //   const targetSheet = targetWorkbook._worksheets[1].name;

  /**
   * create for read the file
   */
  let sourceWorkbook = new Excel.Workbook();

  /**
   * give the name to file which you want read
   */
  sourceWorkbook = await sourceWorkbook.xlsx.readFile(filePath);

  /**
   * gave the sheet name
   */
  const sourceWorksheet = sourceWorkbook.getWorksheet(toBeCopy);
  /**
   * load into memory and create file pointer
   */
  sourceWorksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    /**
     * add row condition in description you wanna 500
     */
    if (rowNumber > 500) return;
    const targetRow = targetWorksheet.getRow(rowNumber);
    row.eachCell({ includeEmpty: false }, (cell, cellNumber) => {
      targetRow.getCell(cellNumber).value = "asdfasd";
      // targetRow.getCell(cellNumber).value = cell.value?.result
      //   ? cell.value?.result
      //   : cell.value;
      // targetRow.getCell(cellNumber).style = cell.style;
    });
    row.commit();
  });

  await targetWorkbook.xlsx.writeFile(filePath);

  console.log(
    `file sheet has been copied from ${toBeCopy} sheet to ${copiedVersion} on path ${filePath}`
  );
  // });
})();

// if (foundXLSM) {
//   original_file.Sheets[copiedVersion] =
//     XLSX.readFile(filePath).Sheets[copiedVersion];
//   fs.unlinkSync(filePath);
//   XLSX.writeFile(original_file, SavedFilePath, { bookVBA: true });
// }

/**
 * write into file
 */
// let original_file = "";
// if (foundXLSM) {
//   original_file = XLSX.readFile(filePath, { bookVBA: true });
//   filePath = filePath.replace(".xlsm", ".xlsx");
// }
