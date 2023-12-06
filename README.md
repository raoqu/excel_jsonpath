# excel_jsonpath
Excel jsonpath operate

# Dependencies
npm install exceljs jsonpath chalk prettyjson
#### jsonpath test
node test.js <xlsx> <columnTitle> <jsonPaht> [row|0]
node test.js input.xlsx request_data "$.processConfig.auditConfig" 2
#### excel update
node main.js <xlsx> <columnTitle> <jsonPath> [newColumnName]
node main.js output.xlsx auditConfig "$[0].auditConfig.ENROLL_SUBMIT.extProperties.fourthUsers" fourthUsers
