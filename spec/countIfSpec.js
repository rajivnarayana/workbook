var workbook = require('../lib/workbook');
var assert = function(condition, errorMessage) {
    expect(condition).toBeTruthy();
}

describe("Simple Workbook ", () => {
    it ("read workbook", (done) => {
        var wb = new workbook();
        var sheet = wb.sheet();
        wb.set(sheet, "A1", 2);
        wb.set(sheet, "A2", 2);
        assert(+wb.run(sheet,`=COUNTIF(A1:A2,"!=0")`)==2);
    });
})

