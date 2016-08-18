var workbook = require('./lib/workbook');
var assert = require('assert');

var wb = new workbook();
var sheet = wb.sheet();
var set_count = 0;
var update_count = 0;

wb.on("set", function(sheetName, row, col, oldValue, newValue) {
    set_count++;
});

wb.on("updated", function(sheetName, row, col, oldValue, newValue) {
    update_count++;
});

wb.set(sheet, "A1", 2);
wb.set(sheet, "A2", 2);
wb.set(sheet, "A3", "=SUM(A1:A2)");

assert( +wb.get(sheet, "A3") === 4, "A3 should be 4");

wb.set(0, "A2", 20);

assert( +wb.get(sheet, "A3") === 22, "A3 should be 22");

assert( set_count === 4, "Set count should be 4" );
assert( update_count === 1, "Update count should be 1" );
