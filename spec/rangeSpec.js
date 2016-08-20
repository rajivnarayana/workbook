var workbook = require('../lib/workbook');
describe("Range Tests :: ", () => {
    var wb = new workbook("CustomPricingModel");
    var pricing_table = wb.sheet({ name: "PriceTable" , store : 'row', data : [
        ["Programs:", 100,                                      200,                      300,                      400],
        ["Rule1",     "=CustomerPrice!A1", "=CustomerPrice!A1*0.95",  "=CustomerPrice!A1*0.8", "=CustomerPrice!A1*0.78"],
        ["Rule2",     "=CustomerPrice!A1", "=CustomerPrice!A1*0.90", "=CustomerPrice!A1*0.85", "=CustomerPrice!A1*0.82"]
    ]});
    var customer_price = wb.sheet({ name: "CustomerPrice" });

    // it("Formula on another sheet", () => {
        // wb.set(customer_price,"A1", "=SUM(PriceTable!B1:E1)");
        // expect( wb.get(customer_price, "A1").valueOf()).toBe(1000);
    // })

    xit("test", () => {
        // wb.set(customer_price,"A1", "=SUM(PriceTable!B1:E1)");
        // expect( wb.get(customer_price, "A1").valueOf()).toBe(1000);
        wb.set(customer_price,"A1", "42"); 
        // customer program
        wb.set(customer_price, "A2", "100"); 

        // lookup discounted price by matching the position of 100 in the
        // decision table and moving 0 column to the right of PriceTable!A2.
        wb.set(customer_price, "A3", '=OFFSET(PriceTable!A2, 0, MATCH(A2, PriceTable!B1:E1, 0))');
        let x = wb.get(customer_price,"A3" ).valueOf();
        expect(x.valueOf()).toBe(42);

    })
})
