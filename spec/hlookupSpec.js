var workbook = require('../lib/workbook');

describe("V and H LOOKUP tests ", () => {
    var wb = new workbook();
    it ("should hLookup ", ()=> {
        let set = (ref, val) => wb.set(sheet, ref, val)
        var sheet = wb.sheet();
        const rows = 5;
        let counter = 1;
        ["Sunday","Monday","Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"].forEach((day, index)=>{
            set(`${String.fromCharCode(65+index)}1`, day);
            new Array(rows).fill(0).forEach((val, colIndex) => set(`${String.fromCharCode(65+index)}${colIndex+2}`, rows * index + colIndex+ 1))
        });
        expect(+wb.run(sheet, `=HLOOKUP("Tuesday",A1:G6,3,0)`)).toBe(12)
    })

    it ("should vLookup ", ()=> {
        let set = (ref, val) => wb.set(sheet, ref, val)
        var sheet = wb.sheet();
        const rows = 5;
        let counter = 1;
        ["Sunday","Monday","Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"].forEach((day, index)=>{
            set(`A${index+1}`, day);
            new Array(rows).fill(0).forEach((val, colIndex) => set(`${String.fromCharCode(65+colIndex+1)}${index+1}`, rows * index + colIndex+ 1))
        });
        expect(+wb.run(sheet, `=VLOOKUP("Tuesday",A1:E7,3,0)`)).toBe(12)
    })

    it ("should not crash when the range is invalid", () => {
        let set = (ref, val) => wb.set(sheet, ref, val)
        var sheet = wb.sheet();
        expect(+wb.run(sheet, `=VLOOKUP("Tuesday",A1:E7,3,0)`)).toBeNaN();
    })
});
