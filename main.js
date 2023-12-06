const ExcelJS = require('exceljs');
const jp = require('jsonpath');
const process = require('process');

const DEFAULT_NEW_COLUMN_NAME = 'JsonPathResult'

async function parseExcelToJson(excelFileName, titleField, jsonPath, newColumnName) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(excelFileName);

    const worksheet = workbook.worksheets[0];
    const titleRowIndex = 1; // 假设第一行是标题行
    let titleFieldIndex = null;

    // 获取标题字段的列索引
    worksheet.getRow(titleRowIndex).eachCell((cell, colNumber) => {
        if (cell.value === titleField) {
            titleFieldIndex = colNumber;
        }
    });

    if (titleFieldIndex === null) {
        console.error(`Title field '${titleField}' not found.`);
        return;
    }

    // 添加新列用于存放JSONPath结果
    const jsonPathColumnIndex = worksheet.columnCount + 1;
    worksheet.getRow(titleRowIndex).getCell(jsonPathColumnIndex).value = newColumnName;

    // 遍历所有行，应用JSONPath表达式
    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        if (rowNumber !== titleRowIndex) { // 跳过标题行
            try {
                const jsonData = JSON.parse(row.getCell(titleFieldIndex).value);
                const jsonPathResult = jp.query(jsonData, jsonPath);
                row.getCell(jsonPathColumnIndex).value = JSON.stringify(jsonPathResult);
            } catch (e) {
                row.getCell(jsonPathColumnIndex).value = 'Error parsing JSON';
            }
        }
    });

    // 保存更新后的Excel文件
    await workbook.xlsx.writeFile(`updated_${excelFileName}`);
}

// 获取命令行参数
const args = process.argv.slice(2);
var newColumnName = DEFAULT_NEW_COLUMN_NAME
if (args.length !== 4 ) {
    if( args.length !== 3) {
        console.log('Usage: node excelToJson.js <excel file> <title field> <json path>');
        process.exit(1);
    }
}
else {
    newColumnName = args[3];
}

parseExcelToJson(args[0], args[1], args[2], newColumnName);
