'use strict'
const config = require('./Config');


try {
    RetrieveRecords(config.DatabaseConfig, config.SQLStatement)
        .then(recordset => WriteToExcel(recordset[0], config.FilePath))
        .catch(err => console.log(err))
        .finally(() => console.log('finished.'));
} catch (err) {
    console.log(err);
}

function WriteToExcel(objectArray, filePath, sheetName) {
    if (!objectArray.length) return;

    const Excel = require('exceljs')
    let workbook = new Excel.Workbook()
    let worksheet = workbook.addWorksheet(sheetName ? sheetName : 'Sheet1');


    //Create Headers
    worksheet.columns = Object.keys(objectArray[0]).map(value => {
        return {
            header: value,
            key: value
        };
    });
    objectArray.forEach(element => worksheet.addRow(element));

    workbook.xlsx.writeFile(filePath || "file1.xlsx");
    console.log('File Saved!!!');

}
async function RetrieveRecords(databaseConfig, query) {
    // connect to your database
    const sql = require("mssql/msnodesqlv8");
    let sqlResult = '';
    await sql.connect(databaseConfig);
    sqlResult = await sql.query(query);
    //console.log(sqlResult.recordsets);
    return sqlResult.recordsets;

}