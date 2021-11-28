const Excel = require('exceljs');

// EDIT THESE:
const fileName = 'BE-survey-analysis';
const worksheetName = 'anketa33921-2021-11-28-3';
//

const filePath = `C:\\Users\\David\\Documents\\Ekonomska Fakulteta\\3rd-YEAR\\BUS. ENVIRONMENT\\${fileName}.xlsx`;

async function excelOperations() {
    let workbook = new Excel.Workbook();
    workbook = await workbook.xlsx.readFile(filePath);
    let worksheet = workbook.getWorksheet(worksheetName);

    //Create results object
    let results = {};

    let peopleWhoKnewOne = 0;

    let column = 13;
    for(let row = 3; row <= worksheet.rowCount; row++) {
        //get string
        let str = worksheet.getCell(row, column).value;

        if(typeof str === 'number') {
            continue;
        }

        peopleWhoKnewOne++;

        //convert string to array of strings + sanitize
        let arr = str.split(',');
        arr = arr.map(name => {
            return name.trim().toLowerCase();
        });

        //iterate over array
        arr.forEach(name => {
            if(results[name]) {
                results[name]++;
            }
            else if(!results[name]) {
                results[name] = 1
            }
        });
    }
    console.log(results);
    console.log(peopleWhoKnewOne);
}

excelOperations();