var workbook = require('../lib/workbook');
describe("Named ranges :: ", () => {
    var wb = new workbook();
    var ws = wb.sheet();
    var ws2 = wb.sheet({name : "Projections"});
    it("Test", () => {
        wb.set(ws, "A1", 0.3125);
        wb.set(ws, "B1", 0.0465);
        wb.set(ws, "C1", "=IF(ROUND(Projections!$K$1*(A1*(1+B1)),0)>=Projections!$K$1, Projections!$K$1, ROUND(Projections!$K$1*(A1*(1+B1)),0))");
        wb.set(ws2, "K1", 96);
        expect(+wb.get(ws,"C1")).toBe(31);
    })
})
