const ExcelJS = require("exceljs");
const path = require("path");
const fs = require("fs");

const filename = path.resolve(__dirname, "../data", "data.xlsx");
const readFileSync = async () => {
  const options = {
    sharedStrings: "emit",
    hyperlinks: "emit",
    worksheets: "emit",
  };
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filename);

  // 访问工作表
  workbook.eachSheet(function (worksheet, sheetId) {
    worksheet.eachRow(function (row, rowNumber) {
      const answers = row.getCell(8);
      const pattern = /[a-dA-D]/g;
      const answerMap = {
        A: "D",
        B: "E",
        C: "F",
        D: "G",
      };

      if (pattern.test(answers)) {
        answers
          .toString()
          .split("")
          .forEach((answer) => {
            const col = answerMap[answer];
            const index = `${col}${rowNumber}`;

            if (!Char) {
              return;
            }

            console.log(index);

            worksheet.getCell(index).style = {
              fill: {
                type: "pattern",
                pattern: "solid",
                fgColor: { argb: "FFFFFF00" },
                bgColor: { argb: "FF00FF00" },
              },
            };
          });
      }
    });
  });

  await workbook.xlsx.writeFile(filename);
};

const readFileAsync = () => {};

readFileSync();
