var _formulaConstants = require('./constants');
module.exports = function INDEX2ROW(index) {
    return Math.floor(index / _formulaConstants.MAX_COLS);
}
