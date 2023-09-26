

const XlsxPopulate = require("xlsx-populate");

function sheet(sheetData) {
    XlsxPopulate.fromBlankAsync()
    .then(async function (workbook) {
      const sheet = workbook?.sheet(0);
      
      // Set headers
      const keys = Object.keys(sheetData[0]);
      let i = 65;
      for (let key of keys) {
        let letter = String.fromCharCode(i);
        i++;
        sheet.cell(`${letter}1`).value(key);
      }

      // Add data from JSON
      sheetData.forEach((entry, index) => {
        const row = index + 2;

        let j = 65;
        for (let k of keys) {
          let col = String.fromCharCode(j);
          sheet.cell(`${col}${row}`).value(entry[k]);
          j++;
        }
      });

      return workbook.toFileAsync('example.xlsx');
      
    //   // Set a password to protect the sheet
    //   const password = "mayo@123";
    //   // password protection removed only for testing purpose
    //   const finalExcel = await workbook.outputAsync({ password });
    })
    .catch(err => {
      context.log(`errror in XlsxPopulate: ${err}`);
      rej(errorinside);
    });
}
sheetData = [
  {
    name: "John",
    age: 24,
    contact: 1234567890,
  }, {
    name: "Jane",
    age: 22,
    contact: 1234567890,
  }, {
    name: "Janet",
    age: 26,
    contact: 1234567890,
  }, {
    name: "Jill",
    age: 28,
    contact: 1234567890,
  }, {
    name: "Jack",
    age: 30,
    contact: 1234567890,
  }
]
sheet(sheetData)
