const Excel = require('exceljs');

const fs = require('fs');

const wb = new Excel.Workbook();
const path = require('path');
const filePath = path.resolve(__dirname,'qr-codes.xlsx');

wb.xlsx.readFile(filePath).then(() => {

    // appendToFile(wb.getWorksheet("1 - 15000-rows BUD"), 2);
    // appendToFile(wb.getWorksheet("2 - 5000-rows BUD"), 2);
    // appendToFile(wb.getWorksheet("3 - 10000-rows BUD"), 2);
    // appendToFile(wb.getWorksheet("4 - 51-rows Stella"), 3);
    // appendToFile(wb.getWorksheet("5 - 10000-rows Brahma DM"), 1);
    // appendToFile(wb.getWorksheet("6 - 25000-rows BUD"), 2);
    appendToFile(wb.getWorksheet("7 - 25000-rows Brahma DM"), 1);
});

const appendToFile = (sheet, campaignType) => {

    const codes = [];

    for (i = 2; i <= sheet.rowCount; i++) {
        const code = sheet.getRow(i).getCell(2).value;
        codes.push(code);
    }

    const chunked = chunk(codes, 100);

    chunked.forEach(x => {
        fs.appendFileSync('7 - 25000-rows Brahma DM.sql', `INSERT INTO product (code, event_campaign_type) VALUES \n`);


        for (i = 0; i < x.length; i++) {
            const sign = i + 1 == x.length ? ';' : ',';
            fs.appendFileSync('7 - 25000-rows Brahma DM.sql', `\t('${x[i]}', ${campaignType})${sign}\n`);
        }

    });
}

const chunk = (input, size) => {
    return input.reduce((arr, item, idx) => {
      return idx % size === 0
        ? [...arr, [item]]
        : [...arr.slice(0, -1), [...arr.slice(-1)[0], item]];
    }, []);
  };