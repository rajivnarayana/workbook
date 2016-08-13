var index2Row = require('./index2row');
var index2Col = require('./index2col');

module.exports = function INDEX2ADDR(index) {
    return {rowIndex : index2Row(index),
        colIndex : index2Col(index)
    }
}