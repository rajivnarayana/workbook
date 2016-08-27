let workbook = require("../lib/workbook");
describe("Dependencies Workbook ", () => {
    var wb = new workbook();
    var sheet = wb.sheet();
    let set = (ref, val) => wb.set(sheet, ref, val)
    let get = ref => +wb.get(sheet,ref)
    it ("should update", () => {
        const columns = 3;
        wb.on('set', (sheetName, row, col, oldVal, val) => {
            // console.log(`${row} ${oldVal}`);
        })
        wb.on('updated', (sheetName, row, col, newVal, oldVal) => {
            expect(newVal-oldVal).toBe(1);
        })

        for(var row=1; row<=3; row++) {
            for (var col=1; col<=columns; col++) {
                set(String.fromCharCode(64+col)+row, 1);
            }
            if (row == 1) {
                set(`${String.fromCharCode(64+columns+1)}1`,`=SUM(A1:${String.fromCharCode(64+columns)}1)`);
            } else {
                set(`${String.fromCharCode(64+columns+1)}${row}`,`=SUM(A${row}:${String.fromCharCode(64+columns)}${row},${String.fromCharCode(64+columns+1)}${row-1})`);
            }
        }
        expect(get(`${String.fromCharCode(64+columns+1)}1`)).toBe(3);
        expect(get(`${String.fromCharCode(64+columns+1)}2`)).toBe(6);
        expect(get(`${String.fromCharCode(64+columns+1)}3`)).toBe(9);
        set('A1', 2);
        expect(get(`${String.fromCharCode(64+columns+1)}1`)).toBe(4);
        expect(get(`${String.fromCharCode(64+columns+1)}2`)).toBe(7);
        expect(get(`${String.fromCharCode(64+columns+1)}3`)).toBe(10);
    })
});
