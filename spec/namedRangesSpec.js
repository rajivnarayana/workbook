var workbook = require('../lib/workbook');
describe("Named ranges :: ", () => {
    var wb = new workbook();
    var ws = wb.sheet();
    wb.nameRange("Values", "Sheet1!A1:A3");
    wb.nameRange("Total", "Sheet1!A4");
    wb.nameRange("Name", "Sheet1!B1");
    it("Test", () => {
        wb.set(ws, "A1:A4", [1,2,3, "=SUM(Values)"]);
        wb.set(ws, "B1", "World");
        expect(wb.run("Total = 6")).toBeTruthy();
        expect( wb.run("2=AVERAGE(Values)") ).toBeTruthy();
        wb.set(ws, "A1", 5);
        expect(wb.get(ws, "A4") == 10);
        expect( wb.run('"Hello, World" = CONCATENATE("Hello, ", Name)') ).toBeTruthy();
        expect( wb.run('"Hello, World" = "Hello, " & Name' ) ).toBeTruthy();
    })
})
