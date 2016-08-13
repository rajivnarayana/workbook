var assert = require('assert');
var workbook = require('./lib/workbook');
var wb = new workbook();
var ws = wb.sheet();
wb.nameRange("Values", "Sheet1!A1:A3");
wb.nameRange("Total", "Sheet1!A4");
wb.nameRange("Name", "Sheet1!B1");
wb.on("set", function(sheetName, row, col, oldValue, newValue) {
    console.log('row: '+row+', oldValue: '+oldValue+', newValue: '+newValue);
})
wb.on("updated", function(sheetName, row, col, oldValue, newValue) {
    console.log('row: '+row+', oldValue: '+oldValue+', newValue: '+newValue);
})

wb.set(ws, "A1:A4", [1,2,3, "=SUM(Values)"]);
wb.set(ws, "B1", "World");
assert( wb.run("Total = 6") );
assert( wb.run("2=AVERAGE(Values)") );
wb.set(ws, "A1", 5);
assert(wb.get(ws, "A4") == 10);
assert( wb.run('"Hello, World" = CONCATENATE("Hello, ", Name)') );
assert( wb.run('"Hello, World" = "Hello, " & Name' ) );