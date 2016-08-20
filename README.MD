# Introduction

Modularized code from [FormulaPages](https://github.com/FormulaPages/org)

# Usage
Look at [sample](./sample/index.js) for how it used
```javascript
var workbook = require('workbook');
var wb = new workbook();
var sheet = wb.sheet();
wb.set(sheet, "A1", 2);
wb.set(sheet, "A2", 2);
wb.set(sheet, "A3", "=SUM(A1:A2)");
console.log('Calculated A3 is '+wb.get(sheet, 'A3').valueOf());
process.exit(0);
```
