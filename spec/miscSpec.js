const workbook = require("../");
describe("Miscellaneous", () => {
    it("COUNTA tests ", ()=>{
        const wb = new workbook();
        const sheet = wb.sheet();
        var set = (ref, val) => wb.set(sheet, ref, val);
        new Array(3).fill(0).forEach((val, index) => set(`A${index+1}`, index+1));
        expect(+wb.run(sheet, "=COUNTA(A1,A3)")).toBe(2);
        expect(+wb.run(sheet, "=COUNTA(A1:A3)")).toBe(3);
        expect(+wb.run(sheet, "=COUNTA(A1:A10)")).toBe(3);
        expect(+wb.run(sheet, "=COUNTA(B1:B3)")).toBe(0);
    })
})