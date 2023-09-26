

const XlsxPopulate = require("xlsx-populate");
function sheet(start, end, columns) {
  XlsxPopulate.fromBlankAsync()
    .then(async function (workbook) {
      const sheet = workbook?.sheet(0);
      let row = 1;
      for (let i = start; i < end; i += columns) {
        let j = 65;
        for (let k = i; k < i + columns; k++) {
          let col = String.fromCharCode(j++);
          if (k > end) {
            break;
          }
          sheet.cell(`${col}${row}`).value(k);
        }
        row++;
      }
      return workbook.toFileAsync('example.xlsx');
    })
    .catch(err => {
      console.log(`errror in XlsxPopulate: ${err}`);
    });
}
sheet(10, 1000, 25)