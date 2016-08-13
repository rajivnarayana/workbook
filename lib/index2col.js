var constants = require('./constants');
var INDEX2ROW = require('./index2row');

module.exports = function INDEX2COL(index) {
  return index - (INDEX2ROW(index) * constants.MAX_COLS);
}