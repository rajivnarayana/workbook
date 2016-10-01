const workbook = require("../lib/workbook");
describe("dependency graph", () => {
    it("should not update lot of times", () => {
        let wb = new workbook();
        let sheet = wb.sheet();
        wb.set(sheet, "A1", "1" );
        wb.set(sheet, "B1", "=A1+1" );
        wb.set(sheet, "C1", "=B1+1" );
        wb.set(sheet, "D1", "=C1+1" );
        wb.set(sheet, "E1", "=D1+1" );
        let counter = 0;
        wb.on("updated", () => counter++);
        wb.set(sheet,"A1","2")
        expect(counter).toBe(4);
    })
})