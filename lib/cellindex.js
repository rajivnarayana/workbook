var _formulaConstants = require("./constants");
module.exports =  function CELLINDEX(row, col) {
return Math.floor(row * _formulaConstants.MAX_COLS) + col;
}
