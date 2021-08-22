const Excel = require("exceljs");

async function copyExcel() {
  /**
   * create the work book instance for target
   */
  let targetWorkbook = new Excel.Workbook();

  /**
   * provide the target file with hardcoded name
   */
  targetWorkbook = await targetWorkbook.xlsx.readFile("my_data/target.xlsx");

  /**
   * provide the worksheet on which new data copy
   */
  const targetWorksheet = targetWorkbook.getWorksheet(`Sheet1`);

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
  sourceWorkbook = await sourceWorkbook.xlsx.readFile("my_data/source.xlsx");

  /**
   * gave the sheet name
   */
  const sourceWorksheet = sourceWorkbook.getWorksheet("Sheet1");

  /**
   * load into memory and create file pointer
   */
  sourceWorksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    /**
     * add row condition in description you wanna 500
     */
    if (rowNumber > 500) return;
    var targetRow = targetWorksheet.getRow(rowNumber);
    row.eachCell({ includeEmpty: false }, (cell, cellNumber) => {
      targetRow.getCell(cellNumber).value = cell.value;
    });
    row.commit();
  });
  /**
   * write into file
   */
  await targetWorkbook.xlsx.writeFile("my_data/target.xlsx");
  console.log("file is copied in target.xlsx");
}

copyExcel();
