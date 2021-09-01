const Excel = require("exceljs");
const fs = require("fs");
const path = require("path");

//const toBeCopy = "About_Me_Dtls1";
//const copiedVersion = "Copied_Version";

const toBeCopy = "EX-DATA";
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

(async () => {
  let files = fs.readdirSync(`${currentDir}/My_Data`);
  if (!files) {
    console.log(err);
    throw new Error("No file found");
  }
  files = files?.filter((value) => !value.startsWith(".~"));
  let filePath = `${currentDir}/My_Data/${files[0]}`;

  let workBook = new Excel.Workbook();

  workBook = await workBook.xlsx.readFile(filePath);

  let targetWorksheet = workBook.getWorksheet(copiedVersion);
  if (targetWorksheet) {
    await targetWorksheet.destroy();
    targetWorksheet = await workBook.addWorksheet(copiedVersion);
  }

  const sourceWorksheet = workBook.getWorksheet(toBeCopy);

  sourceWorksheet.eachRow({ includeEmpty: false }, async (row, rowNumber) => {
    if (rowNumber > 500) return;
    const targetRow = targetWorksheet.getRow(rowNumber);
    const sourceRow = sourceWorksheet.getRow(rowNumber);
    row.eachCell({ includeEmpty: false }, async (cell, cellNumber) => {
      targetRow.getCell(cellNumber).value =
        typeof sourceRow.getCell(cellNumber).value === "object"
          ? sourceRow.getCell(cellNumber).value?.result
            ? sourceRow.getCell(cellNumber).value?.result
            : ""
          : sourceRow.getCell(cellNumber).value;
      targetRow.getCell(cellNumber).style = sourceRow.getCell(cellNumber).style;
      // targetRow.getCell(cellNumber).value.formula = null;
      await targetRow.commit();
      // await targetRow.commit();
      // await sourceWorksheet.commit;
      // await sourceRow.commit();
    });
    await row.commit();
  });

  await workBook.xlsx.writeFile(filePath);

  console.log(
    `file sheet has been copied from ${toBeCopy} sheet to ${copiedVersion} on path ${filePath}`
  );
})();

// console.log(
//   typeof sourceRow.getCell(cellNumber).value === "object"
//     ? sourceRow.getCell(cellNumber).value?.result
//       ? sourceRow.getCell(cellNumber).value.result
//       : ""
//     : sourceRow.getCell(cellNumber).value
// );
