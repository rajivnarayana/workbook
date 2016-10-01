var workbook = require('../lib/workbook');

let pendingUpdates = [];

var fn = {
    ISCELL : function(ref) {
        return ((value && typeof value == 'object') && ref.constructor.name === "cell");
    }
}
workbook.prototype.update = function (sheetIndex, cell) {
    var dirtyCells = this.findDependents(sheetIndex, cell);

    if (dirtyCells) {
        for (var i = 0; i < dirtyCells.length; i++) {
            this.updateDirtyCellInNextTick(dirtyCells[i]);
        }
    }
}

workbook.prototype.updateDirtyCellInNextTick = function(cellRef) {
    if (pendingUpdates.some(cell => {
            return cell.sheetIndex == cellRef.sheetIndex && cell.row == cellRef.rowIndex && cell.col == cellRef.colIndex
        })) {
            return;
        }
        let ref = {sheetIndex : cellRef.sheetIndex, row : cellRef.rowIndex, col : cellRef.colIndex}
        pendingUpdates.push(ref);
        setTimeout(function() {
            pendingUpdates.splice(pendingUpdates.indexOf(ref), 1);
            this.updateDirtyCell(cellRef);
            if (pendingUpdates.length == 0) {
                console.log('done');
            }
        }.bind(this),0)
}
workbook.prototype.updateDirtyCell = function(dirtyCell) {
                var cellIndex = dirtyCell.cellIndex,
                rowIndex = dirtyCell.rowIndex,
                colIndex = dirtyCell.colIndex,
                sheetIndex = dirtyCell.sheetIndex;
                console.log(`Updating dependent ${rowIndex}, ${colIndex}`);
            let oldValue = this.getValue(sheetIndex, rowIndex, colIndex);
            let newValue = this.recalculate(sheetIndex, cellIndex);

            if (oldValue !== newValue) {
                if (fn.ISCELL(newValue) || oldValue !== newValue.valueOf()) {
                    this.set(sheetIndex,
                        rowIndex,
                        colIndex,
                        newValue.valueOf());

                    // Trigger notification on the callbacks
                    this.triggerEvent("updated",
                        [sheetIndex,
                            rowIndex,
                            colIndex,
                            newValue,
                            oldValue]);

                }
            }

}
describe("Set delayed spec", () => {
    
    var wb = new workbook();
/*
    let originalUpdate = wb.update;
    let pendingUpdates = [];
    wb.update = function(sheetIndex, cellRef) {
        console.log(`${cellRef.rowIndex}, ${cellRef.colIndex}`)
        if (pendingUpdates.some(cell => {
            return cell.sheetIndex == sheetIndex && cell.row == cellRef.rowIndex && cell.col == cellRef.colInex
        })) {
            return;
        }
        let ref = {sheetIndex : sheetIndex, row : cellRef.rowIndex, col : cellRef.colIndex}
        pendingUpdates.push(ref);
        setTimeout(function() {
            pendingUpdates.splice(pendingUpdates.indexOf(ref), 1);
            console.log(`pending update ${cellRef.rowIndex}, ${cellRef.colIndex}`)
            originalUpdate.call(wb, sheetIndex, cellRef);
        }, 0)
    }
*/
    var ws = wb.sheet();

    it("should update on next tick", (done) => {
        wb.on('updated', (sheetIndex, row, col, newValue, oldValue) => {
            console.log(`${row} ${col} ${oldValue} ${newValue}`);
        })
        wb.set(ws,'B1', '=SUM(A1:A10)');
        new Array(10).fill(0).forEach((val, index) => {
            wb.set(ws, 'A'+(index + 1), 1);
        })
        setTimeout(done, 1000);
    })
})