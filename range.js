var assert = require('assert');
require('./lib/workbook');
var workbook = global.workbook;
var wb = new workbook("CustomPricingModel");
var pricing_table = wb.sheet({ name: "PriceTable" , store : 'row', data : [
        ["Programs:", 100,                                      200,                      300,                      400],
        ["Rule1",     "=CustomerPrice!A1", "=CustomerPrice!A1*0.95",  "=CustomerPrice!A1*0.8", "=CustomerPrice!A1*0.78"],
        ["Rule2",     "=CustomerPrice!A1", "=CustomerPrice!A1*0.90", "=CustomerPrice!A1*0.85", "=CustomerPrice!A1*0.82"]
    ]});
var customer_price = wb.sheet({ name: "CustomerPrice" });

wb.set(customer_price,"A1", "=SUM(PriceTable!B1:E1)");
console.log(wb.cell(customer_price, "A1").valueOf());
// assert( wb.get(customer_price, "A1") === 1000, "Price should be 1000");