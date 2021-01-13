const Excel = require('exceljs');
const generator = require('randomstring');

const createCsvWriter = require('csv-writer').createObjectCsvWriter;
const csvWriter = createCsvWriter({
  path: 'out.csv',
  header: [
    {id: 'url', title: 'url'},
    {id: 'code', title: 'code'}
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

    console.log('allCodes', allCodes.length);


    while (newCodes.length !== 10000) {
        const code = generator.generate(7);

        if (!allCodes[code.toLocaleLowerCase()]) {
            console.log(code);
            newCodes.push(code.toLocaleLowerCase());
            console.log('New codes: ', newCodes.length);
        }
    }

    csvWriter.writeRecords(newCodes.map(code => ({url: `https://orig.app/${code}`, code: code}))).then(()=> console.log('The CSV file was written successfully'));


    // sh1.getRow(1).getCell(2).value = 32;
    // wb.xlsx.writeFile("sample2.xlsx");
    // console.log("Row-3 | Cell-2 - "+sh1.getRow(3).getCell(2).value);
    
    // console.log(sh1.rowCount);
    // //Get all the rows data [1st and 2nd column]
    // for (i = 1; i <= sh1.rowCount; i++) {
    //     console.log(sh1.getRow(i).getCell(1).value);
    //     console.log(sh1.getRow(i).getCell(2).value);
    // }
});