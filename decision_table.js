var assert = require('assert');
require('./lib/workbook');
var workbook = global.workbook;

function Sheet(sheet, workbook) {
    this.sheet = sheet;
    this.workbook = workbook;
}

Sheet.prototype.run = function() {

}

Sheet.prototype.set = function() {

}

Sheet.prototype.get = function() {

}

   var wb = new workbook("CustomPricingModel");
   var pricing_table = wb.sheet({ name: "PriceTable" , store : 'row', data : [
	     ["Programs:", 100,                                      200,                      300,                      400],
	     ["Rule1",     "=CustomerPrice!A1", "=CustomerPrice!A1*0.95",  "=CustomerPrice!A1*0.8", "=CustomerPrice!A1*0.78"],
	     ["Rule2",     "=CustomerPrice!A1", "=CustomerPrice!A1*0.90", "=CustomerPrice!A1*0.85", "=CustomerPrice!A1*0.82"]
	 ]});
   var customer_price = wb.sheet({ name: "CustomerPrice" });

//    wb.set(pricing_table,
// 	 [
// 	     ["Programs:", 100,                                      200,                      300,                      400],
// 	     ["Rule1",     "=CustomerPrice!A1", "=CustomerPrice!A1*0.95",  "=CustomerPrice!A1*0.8", "=CustomerPrice!A1*0.78"],
// 	     ["Rule2",     "=CustomerPrice!A1", "=CustomerPrice!A1*0.90", "=CustomerPrice!A1*0.85", "=CustomerPrice!A1*0.82"]
// 	 ]
//    );


   // order total before discount
   wb.set(customer_price,"A1", "42"); 
   // customer program
   wb.set(customer_price,"A2", "100"); 

   // lookup discounted price by matching the position of 100 in the
   // decision table and moving 0 column to the right of PriceTable!A2.
   wb.set(customer_price,"A3", '=OFFSET(PriceTable!A2, 0, MATCH(A2, PriceTable!B1:E1, 0))');
   wb.on("updated", function(sheetName, row, col, oldValue, newValue) {
        console.log('row: '+row+", col: "+col+", value: "+newValue);
    });
   assert( +wb.get(customer_price,"A3").valueOf() === 42, "Price should be 42");

//    wb.set(customer_price,"A2", "200");

   // Now the discounted price is found 1 column to the right.
//    assert( +customer_price.get("A3") === (42*0.95), "Price should be " + (42*0.95));

   // Now the discounted price is found 2 columns to the right.
//    wb.set(customer_price,"A2", "300");

   assert( +customer_price.get("A3") === (42*0.8), "Price should be " + (42*0.8));

   wb.set(customer_price,"A2", "400");

   assert( +customer_price.get("A3") === (42*0.78), "Price should be " + (42*0.78));

   wb.set(customer_price,"A2", "foo");

   assert( customer_price.run("A3 = NA()"), "Error should be #N/A");
