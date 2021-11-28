const Excel = require('exceljs');

// EDIT THESE:
const fileName = 'BE-survey-analysis';
const worksheetName = 'survey-working-final';
//

const filePath = `C:\\Users\\David\\Documents\\Ekonomska Fakulteta\\3rd-YEAR\\BUS. ENVIRONMENT\\${fileName}.xlsx`;

async function excelOperations() {
    let workbook = new Excel.Workbook();
    workbook = await workbook.xlsx.readFile(filePath);
    
    let worksheet = workbook.getWorksheet(worksheetName);
    
    // Replace missing answers with 'null', so they are not included in correlation analysis
    for(let column = 1; column <= worksheet.columnCount; column++) {
        for(let row = 2; row <= worksheet.rowCount; row++) {
            let val = worksheet.getCell(row, column).value;
            if(val < 0) {
                worksheet.getCell(row, column).value = null;
            }
        }
    }
    workbook.xlsx.writeFile(filePath);
    console.log('Done');
}

excelOperations();