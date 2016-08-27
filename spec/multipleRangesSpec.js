var workbook = require('../lib/workbook');
describe("Multiple ranges :: ", () => {
    var wb = new workbook();
    wb.debug();
    let sideLength = 3;
    var squareArray = function(sideLength = 3) {
        return Array(sideLength * sideLength).fill(0).map((v,index) => index+1).reduce((prev, element, index) => {
            if (index %sideLength == 0)prev.push([element]);else prev[prev.length-1].push(element); return prev;}, []) 
    }
    let sheet = wb.sheet({ name: "PriceTable" , store : 'row', data : squareArray(sideLength)});
    it ('test', ()=> {

        wb.set(sheet, 'D1','=SUM(A1:C1)');
        wb.set(sheet, 'D2','=SUM(A2:C2)');
        wb.set(sheet, 'D3','=SUM(A3:C3)');
        wb.set(sheet, 'A4','=SUM(A1:A3)');
        wb.set(sheet, 'B4','=SUM(B1:B3)');
        wb.set(sheet, 'C4','=SUM(C1:C3)');
        wb.set(sheet, 'D4','=SUM(D1:D3)');

        expect(+wb.get(sheet, 'A1')).toBe(1);
        expect(+wb.get(sheet, 'B1')).toBe(2);
        expect(+wb.get(sheet, 'C1')).toBe(3);
        expect(+wb.get(sheet, 'D1')).toBe(6);
        expect(+wb.get(sheet, 'A4')).toBe(12);

        expect(+wb.get(sheet, 'A2')).toBe(4);
        expect(+wb.get(sheet, 'B2')).toBe(5);
        expect(+wb.get(sheet, 'C2')).toBe(6);
        expect(+wb.get(sheet, 'D2')).toBe(15);
        expect(+wb.get(sheet, 'B4')).toBe(15);

        expect(+wb.get(sheet, 'A3')).toBe(7);
        expect(+wb.get(sheet, 'B3')).toBe(8);
        expect(+wb.get(sheet, 'C3')).toBe(9);
        expect(+wb.get(sheet, 'D3')).toBe(24);
        expect(+wb.get(sheet, 'C4')).toBe(18);

        expect(+wb.get(sheet, 'D4')).toBe(45);

        wb.set(sheet, 'A1', 11);
        expect(+wb.get(sheet, 'D4')).toBe(55);
    })
})