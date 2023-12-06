const ExcelJS = require('exceljs');
const jp = require('jsonpath');
const process = require('process');
// const chalk = require('chalk');
const prettyjson = require('prettyjson');

function prettyPrintJson(json) {
    const options = {
        noColor: false,
    };

    // 使用prettyjson格式化JSON
    let formattedJson = prettyjson.render(json, options);

    // 使用chalk添加额外的颜色
    // formattedJson = chalk.blue(formattedJson);

    console.log(formattedJson);
}

async function parseExcelToJson(excelFileName, titleField, jsonPath, sampleRow) {
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
    worksheet.getRow(titleRowIndex).getCell(jsonPathColumnIndex).value = 'JsonPathResult';

    // 遍历所有行，应用JSONPath表达式
    var output = false;
    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        if (rowNumber == sampleRow || (sampleRow == 0 && rowNumber != titleRowIndex)) { // 跳过标题行
            try {
                const jsonData = JSON.parse(row.getCell(titleFieldIndex).value);
                const jsonPathResult = jp.query(jsonData, jsonPath);
                const val = JSON.stringify(jsonPathResult, null, 2);
                console.log('' + rowNumber + ': ')
                prettyPrintJson(val)
                console.log('')
            } catch (e) {
                row.getCell(jsonPathColumnIndex).value = 'Error parsing JSON';
            }
        }
    });
}

// 获取命令行参数
const args = process.argv.slice(2);
var targetRow = 0;
if (args.length !== 4) {
    if( args.length == 3) {
        targetRow = 0;
    }
    else {
        console.log('Usage: node excelToJson.js <excel file> <title field> <json path>');
        process.exit(1);
    }
} else {
    targetRow = parseInt(args[3])
}

parseExcelToJson(args[0], args[1], args[2], targetRow);
