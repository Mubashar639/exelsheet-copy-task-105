const Excel = require("exceljs");
const fs = require("fs");
const path = require("path");

const toBeCopy = "About_Me_Dtls1";

const copiedVersion = "Copied_Version";

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
  const files = fs.readdirSync(`${currentDir}/My_Data`);
  if (!files) {
    console.log(err);
    throw new Error("No file found");
  }
  console.log(files);

  const filePath = `${currentDir}/My_Data/${files[0]}`;

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
      targetRow.getCell(cellNumber).value = cell.value?.result
        ? cell.value?.result
        : cell.value;
      targetRow.getCell(cellNumber).style = cell.style;
    });
    row.commit();
  });
  /**
   * write into file
   */
  await targetWorkbook.xlsx.writeFile(filePath);

  console.log(
    `file sheet has been copied from ${toBeCopy} sheet to ${copiedVersion} on path ${filePath}`
  );
  // });
})();
