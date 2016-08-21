var workbook = require('../lib/workbook');
var assert = function(condition, errorMessage) {
    expect(condition).toBeTruthy();
}

describe("Simple Workbook ", () => {
    var wb = new workbook();
    var sheet = wb.sheet();
    wb.set(sheet, "A1", "=B1");
    wb.set(sheet, "A2", "=B2");
    wb.set(sheet, "B1", 2);
    wb.set(sheet, "B2", 2);
    it("Should update cell with formula when dependents update", () => {
        //Addition
        wb.set(sheet, "A3", "=SUM(A1:A2)");
        assert( +wb.get(sheet, "A3") === 4, "A3 should be 4");

        //Subtraction
        wb.set(sheet, "A4", "=A1-A2");
        assert( +wb.get(sheet, "A4") === 0, "A3 should be 4");

        //multiplication
        wb.set(sheet, "A5", "=A1*A2");
        assert( +wb.get(sheet, "A5") === 4, "A3 should be 4");

        //division
        wb.set(sheet, "A6", "=A1/A2");
        assert( +wb.get(sheet, "A6") === 1, "A3 should be 4");
    });
})

