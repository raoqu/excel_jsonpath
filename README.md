# excel_jsonpath
Excel jsonpath operate

# Dependencies
```bash
npm install exceljs jsonpath chalk prettyjson
```

# jsonpath test
```bash
node test.js <xlsx> <columnTitle> <jsonPaht> [row|0]
```

```bash
node test.js input.xlsx request_data "$.config" 2
```

# excel update
```bash
node main.js <xlsx> <columnTitle> <jsonPath> [newColumnName]
```

```bash
node main.js output.xlsx auditConfig "$[0].config.users" users
```
