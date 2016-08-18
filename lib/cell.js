var CELLINDEX = require('./cellindex');
var INDEX2ADDR = require('./index2addr');
var workbookUtils = require('./workbook_utils');
var fn = require("./utils");

  function cell(wb, sheetIndex) {
      var cellIndex, addr, rowIndex, colIndex;

      if (arguments.length === 3) {
          cellIndex = arguments[2];
          var addr = INDEX2ADDR(cellIndex);
          rowIndex = addr.rowIndex;
          colIndex = addr.colIndex;

      } else if (arguments.length === 4) {
          rowIndex = arguments[2];
          colIndex = arguments[3];
          cellIndex = CELLINDEX( rowIndex, colIndex );
      }

      if (wb.cells[sheetIndex][cellIndex]) {
          return wb.cells[sheetIndex][cellIndex];
      } else {

          this.workbook = wb;
          this.sheetIndex = sheetIndex;
          this.cellIndex = cellIndex;
          this.rowIndex = rowIndex;
          this.colIndex = colIndex;

          wb.cells[sheetIndex][cellIndex] = this;

          return this;
      }
  };

  cell.prototype.valueOf = cell.prototype.val = cell.prototype.value = function() {
      return this.workbook.getValue(this.sheetIndex,
                                    this.rowIndex,
                                    this.colIndex);
  };

  cell.prototype.addr = function() {
      return fn.ADDRESS(this.rowIndex+1, this.colIndex+1, 0);
  }

  cell.prototype.toString = function() {
      return workbookUtils.sheetName(this.sheetName, this.addr());
  };

  module.exports = cell;
  