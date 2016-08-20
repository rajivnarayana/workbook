var workbook = require('workbook');
var wb = new workbook();
var sheet = wb.sheet();
wb.set(sheet, "A1", 2);
wb.set(sheet, "A2", 2);
wb.set(sheet, "A3", "=SUM(A1:A2)");
console.log(`Set A1: ${2}, A2: ${2}, Value of A3 calculated as '=SUM(A1:A2)' is ${wb.get(sheet, 'A3').valueOf()}`);
process.exit(0);