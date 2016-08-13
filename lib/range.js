var fn = require("./utils");
var CELLINDEX = require("./cellindex");
  function range(wb, sheetIndex, topLeft, bottomRight) {

      // link back to workbook
      this.workbook = wb;

      // the sheet on which the range is set.
      this.sheetIndex = sheetIndex;

      if (topLeft.cellIndex > bottomRight.cellIndex) {
          throw Error("topLeft must be smaller then bottomRight");
      }

      if (typeof topLeft !== 'function' && !fn.ISCELL(topLeft)) {
          throw Error("topLeft must be function or reference");
      }

      if (typeof bottomRight !== 'function' && !fn.ISREF(bottomRight)) {
          throw Error("bottomRight must be function or reference");
      }

      if (fn.ISFUNCTION(topLeft) || fn.ISFUNCTION(bottomRight)) {
          this.dynamic = true;
          this.topLeft = topLeft;
          this.bottomRight = bottomRight;
          return this;
      } else {
          this.dynamic = false;

          // combine cellIndexes to get a unique key for ranges.
          this.rangeIndex = topLeft.cellIndex + bottomRight.cellIndex;
      }


      // if already created then return it from the index.
      if (wb.rangeLookup[sheetIndex] &&
          wb.rangeLookup[sheetIndex][this.rangeIndex]) {
          var rangeId = wb.rangeLookup[sheetIndex][this.rangeIndex];
          return wb.ranges[rangeId];
      }

      this.rangeId = wb.ranges.length;
      this.sheetIndex = topLeft.sheetIndex;
      this.topLeft = topLeft;
      this.bottomRight = bottomRight;

      if (!wb.rangeLookup[sheetIndex]) {
           wb.rangeLookup[sheetIndex] = [];
      }

      wb.rangeLookup[sheetIndex][this.rangeIndex] = this.rangeId;
      wb.ranges[this.rangeId] = this;

      return this;
  }   

  range.prototype.cells = function(mode) {
      var a,b, sheetIndex, wb;                             

      a = workbook.resolveCell(this.topLeft);
      b = workbook.resolveCell(this.bottomRight);    

      if (a.sheetIndex !== b.sheetIndex) {
          throw Error("topLeft and bottomRight must be on the same sheet.");
      }          

      if (a === null || b === null) {
          return [];
      }

      wb = a.workbook;
      sheetIndex = a.sheetIndex;

      var cells = [];
      var colA = a.colIndex;
      var rowA = a.rowIndex;

      var colB = b.colIndex;
      var rowB = b.rowIndex;

      for (var rowNum = rowA; rowNum <= rowB; rowNum++) {
          for (var colNum = colA; colNum <= colB; colNum++) {

              if (mode === 2) {
                  cells.push(fn.ADDRESS(rowNum +1, colNum+1, 0));
              } else if (mode === 3) {
                  cells.push(wb.sheetName( wb.sheetNames[sheetIndex],
                                                 fn.ADDRESS(rowNum +1, colNum+1, 0) ));
              } else if (mode === 4) {
                  cells.push( new cell( wb, sheetIndex, CELLINDEX( rowNum, colNum ) ) );
              } else {
                  cells.push( CELLINDEX( rowNum, colNum ) );
              }
          }
      }

      return cells;
  }

  // return true if the address is inside the range
  range.prototype.hit = function(addr) {
      var a,b,target;

      target = workbook.cellInfo(addr);
      a = workbook.resolveCell(this.topLeft);
      b = workbook.resolveCell(this.bottomRight);        

      if (a.sheetIndex !== b.sheetIndex) {
          throw Error("topLeft and bottomRight must be on the same sheet.");
      }

      return (target.colIndex >= a.colIndex &&
              target.colIndex <= b.colIndex &&
              target.rowIndex >= a.rowIndex &&
              target.rowIndex <= b.rowIndex);
  }

  range.prototype.set = function(values) {

      if (!fn.ISARRAY(values)) {
          throw TypeError("Only accepts array of values.");
      }
      var topLeft = workbook.resolveCell(this.topLeft);

      var cells = this.cells();
      for (var i =0; i < cells.length; i++) {
          this.workbook.set(topLeft.sheetIndex, cells[i], values[i]);
      }
  }

  range.prototype.values = range.prototype.valueOf = function() {
    var a,b, index;

    a = workbook.resolveCell(this.topLeft);
    b = workbook.resolveCell(this.bottomRight);
    
    if (a.sheetIndex !== b.sheetIndex) {
        throw Error("topLeft and bottomRight must be on the same sheet.");
    }

    if (a === null || b === null) {
        return [];
    }

    this.workbook = a.workbook;
    index = this.sheetIndex = a.sheetIndex;

    var data = [];
    var colA = a.colIndex;
    var rowA = a.rowIndex;

    var colB = b.colIndex;
    var rowB = b.rowIndex;
    
    for (var rowNum = rowA; rowNum <= rowB; rowNum++) {
        // When only 1 column then flatten the output
        if (colA === colB) {
            data[rowNum-rowA] = this.workbook.getValue(index, rowNum, colA);
        } else {
            for (var colNum = colA; colNum <= colB; colNum++) {
                // When only 1 row then flatten the output
                if (rowA === rowB) {                        
                    data[colNum-colA] = this.workbook.getValue(index, rowNum, colNum); 
                } else {
                    if (typeof data[rowNum-rowA] === 'undefined') {
                        data[rowNum-rowA] = [];
                    }
                    data[rowNum-rowA][colNum-colA] = this.workbook.getValue(index, rowNum, colNum);
                }
            }
        }
    }

    return data;

}

range.prototype.toString = function() {
    return workbook.sheetName(this.workbook.sheetNames[this.sheetIndex], this.topLeft.addr) + ":" + this.bottomRight.addr;
}

module.exports = range;