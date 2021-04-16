const Excel = require('exceljs');
const generator = require('randomstring');

const NEW_CODES_QUANTITY = 25000;

const createCsvWriter = require('csv-writer').createObjectCsvWriter;
const csvWriter = createCsvWriter({
  path: '7 - 25000-rows Brahma DM.csv',
  fieldDelimiter: ';',
  header: [
    {id: "_", title: "_"},
    {id: 'code', title: 'code'},
    {id: 'url', title: 'url'}
  ]
});

const wb = new Excel.Workbook();
const path = require('path');
const filePath = path.resolve(__dirname,'qr-codes.xlsx');

wb.xlsx.readFile(filePath).then(() => {
    const sh1 = wb.getWorksheet("1 - 15000-rows BUD");
    const sh2 = wb.getWorksheet("2 - 5000-rows BUD");
    const sh3 = wb.getWorksheet("3 - 10000-rows BUD");
    const sh4 = wb.getWorksheet("4 - 51-rows Stella");
    const sh5 = wb.getWorksheet("5 - 10000-rows Brahma DM");
    const sh6 = wb.getWorksheet("6 - 25000-rows BUD");
    const sh7 = wb.getWorksheet("7 - 25000-rows Brahma DM");

    const allCodes = [];

    const newCodes = [];

    for (i = 1; i < sh1.rowCount; i++) {
        allCodes.push(sh1.getRow(i).getCell(2).value);
    }

    for (i = 1; i < sh2.rowCount; i++) {
        allCodes.push(sh2.getRow(i).getCell(2).value);
    }

    for (i = 1; i < sh3.rowCount; i++) {
        allCodes.push(sh3.getRow(i).getCell(2).value);
    }

    for (i = 1; i < sh4.rowCount; i++) {
        allCodes.push(sh4.getRow(i).getCell(2).value);
    }

    for (i = 1; i < sh5.rowCount; i++) {
        allCodes.push(sh5.getRow(i).getCell(2).value);
    }

    for (i = 1; i < sh6.rowCount; i++) {
        allCodes.push(sh6.getRow(i).getCell(2).value);
    }
    
    console.log('allCodes.length', allCodes.length);

    while (newCodes.length !== NEW_CODES_QUANTITY) {
        const code = generator.generate(7);

        if (!allCodes[code.toLocaleLowerCase()]) {
            newCodes.push(code.toLocaleLowerCase());
        }

        console.log(newCodes.length);
    }
    console.log('allCodes', allCodes.length);

    csvWriter.writeRecords(newCodes.map(code => ({_: "https://orig.app/", code: code,url: `https://orig.app/${code}` }))).then(()=> console.log('The CSV file was written successfully'));
});