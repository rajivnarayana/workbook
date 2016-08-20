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
    let value = ref => wb.get(sheet, ref).valueOf();
    it ('test', ()=> {

        wb.set(sheet, 'D1','=SUM(A1:C1)');
        wb.set(sheet, 'D2','=SUM(A2:C2)');
        wb.set(sheet, 'D3','=SUM(A3:C3)');
        wb.set(sheet, 'A4','=SUM(A1:A3)');
        wb.set(sheet, 'B4','=SUM(B1:B3)');
        wb.set(sheet, 'C4','=SUM(C1:C3)');
        wb.set(sheet, 'D4','=SUM(D1:D3)');

        expect(value('A1')).toBe(1);
        expect(value('B1')).toBe(2);
        expect(value('C1')).toBe(3);
        expect(value('D1')).toBe(6);
        expect(value('A4')).toBe(12);

        expect(value('A2')).toBe(4);
        expect(value('B2')).toBe(5);
        expect(value('C2')).toBe(6);
        expect(value('D2')).toBe(15);
        expect(value('B4')).toBe(15);

        expect(value('A3')).toBe(7);
        expect(value('B3')).toBe(8);
        expect(value('C3')).toBe(9);
        expect(value('D3')).toBe(24);
        expect(value('C4')).toBe(18);

        expect(value('D4')).toBe(45);

        wb.set(sheet, 'A1', 11);
        expect(value('D4')).toBe(55);
    })
})