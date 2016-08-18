var error = require('./error.js');
var workbook = require('./workbook_utils');

exports.flatten = function (array) {
    var a = array;
    if (a.constructor.name === "range") {
        var a = a.values();
    }

    var r = [];

    function _flatten(a) {
        for (var i = 0; i < a.length; i++) {
            if (typeof a[i] === "undefined" || a[i] === null) {
                continue; // empty cell or variable...just keep rolling
            } else if (a[i].constructor.name === "range") {
                _flatten(a[i].values());
            } else if (a[i].constructor == Array) {
                _flatten(a[i]);
            } else {
                r.push(a[i]);
            }
        }
        return r;
    }

    return _flatten(a);
}
// isAnyError

exports.isAnyError = function () {
    var n = arguments.length;
    while (n--) {
        if (arguments[n] instanceof Error) {
            return true;
        }
    }
    return false;
};
// parseNumber

exports.parseNumber = function (string) {
    if (string === undefined || string === '') {
        return error.value;
    }
    if (!isNaN(string)) {
        return parseFloat(string);
    }
    return error.value;
};
// parseDate

exports.parseDate = function (date) {
    if (typeof date === 'string') {
        date = new Date(date);
        if (!isNaN(date)) {
            return date;
        }
    } else if (date === date) {
        if (date instanceof Date) {
            return new Date(date);
        }
        var d = parseInt(date, 10);
        if (d < 0) {
            return error.num;
        }
        if (d <= 60) {
            return new Date(d1900.getTime() + (d - 1) * MilliSecondsInDay);
        }
        return new Date(d1900.getTime() + (d - 2) * MilliSecondsInDay);
    }

    return error.value;
};
// parseBool

exports.parseBool = function (bool) {
    if (typeof bool === 'boolean') {
        return bool;
    }

    if (bool instanceof Error) {
        return bool;
    }

    if (typeof bool === 'number') {
        return bool !== 0;
    }

    if (typeof bool === 'string') {
        var up = bool.toUpperCase();
        if (up === 'TRUE') {
            return true;
        }

        if (up === 'FALSE') {
            return false;
        }
    }

    if (bool instanceof Date && !isNaN(bool)) {
        return true;
    }

    return error.value;
};
// serial
/*
Dates as sequential serial numbers so they can be used in calculations. By default, January 1, 1900 is serial number 1, and January 1, 2008 is serial number 39448 because it is 39,448 days after January 1, 1900.
*/
exports.serial = function (date) {
    var addOn = (date > -2203891200000) ? 2 : 1;
    return Math.ceil((date - d1900) / MilliSecondsInDay) + addOn;
}
// serialTime
/*
Dates as sequential serial numbers so they can be used in calculations. By default, January 1, 1900 is serial number 1, and January 1, 2008 is serial number 39448 because it is 39,448 days after January 1, 1900.
*/
// DONT THINK THIS WORKS!
exports.serialTime = function (date) {
    var addOn = (date > -2203891200000) ? 2 : 1;
    return ((date - d1900) / MilliSecondsInDay) + addOn;
}
// validDate

exports.validDate = function (d) {
    return d && d.getTime && !isNaN(d.getTime());
}
// GUID

// credit to http://stackoverflow.com/questions/105034/create-guid-uuid-in-javascript
// rfc4122 version 4 compliant solution
exports.GUID = function () {
    var guid = 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
        var r = Math.random() * 16 | 0, v = c == 'x' ? r : (r & 0x3 | 0x8);
        return v.toString(16);
    });
    return guid;
};
// Date

// DAY

exports.DAY = function (date) {
    if (typeof date === "string") {
        return new Date(date).getDate();
    } else {
        return exports.parseDate(date).getDate();
    }
}
// DATE

exports.DATE = function (year, month, day) {
    return exports.serial(new Date(year, month - 1, day));
}
// DATEDIF

exports.DATEDIF = function (start_date, end_date, unit) {
    var s, e; // start and end serial

    if (typeof start_date === "string") {
        start_date = exports.parseDate(start_date);
        s = exports.serial(start_date);
    } else {
        s = start_date;
        start_date = exports.parseDate(s);
    }

    if (typeof end_date === "string") {
        end_date = exports.parseDate(end_date);
        e = exports.serial(end_date);
    } else {
        e = end_date;
        end_date = exports.parseDate(e);
    }

    if (end_date === start_date) {
        return 0;
    }

    if (end_date < start_date) {
        return error.num;
    }

    switch (unit) {
        case "Y":
            return end_date.getFullYear() - start_date.getFullYear();
        case "M":
            return (((end_date.getFullYear() - start_date.getFullYear()) * 12)
                - (start_date.getMonth())
                + end_date.getMonth());
        case "D":
            return e - s;
        case "MD":
            return "Not supported"
        case "YM":
            return "Not supported"
        case "YD":
            return "Not supported"
    }

    return error.na;
}
// DATEVALUE

exports.DATEVALUE = function () {
    if (!arguments.length) {
        return new Date();
    }

    if (arguments.length === 1) {
        return exports.serial(new Date(arguments[0]));
    }

    var args = arguments;
    args[1] = args[1] - 1; // Monthes are between 0 and 11.
    return exports.serial(new (Date.bind.apply(Date, [Date].concat([].splice.call(args, 0))))());
}
// DAYS360

exports.DAYS360 = function (start_date, end_date, method) {
    method = typeof method === 'undefined' ? false : exports.parseBool(method);
    start_date = exports.parseDate(exports.DATEVALUE(start_date));
    end_date = exports.parseDate(exports.DATEVALUE(end_date));

    if (start_date instanceof Error) {
        return start_date;
    }
    if (end_date instanceof Error) {
        return end_date;
    }
    if (method instanceof Error) {
        return method;
    }
    var sm = start_date.getMonth();
    var em = end_date.getMonth();
    var sd, ed;
    if (method) {
        sd = start_date.getDate() === 31 ? 30 : start_date.getDate();
        ed = end_date.getDate() === 31 ? 30 : end_date.getDate();
    } else {
        var smd = new Date(start_date.getFullYear(), sm + 1, 0).getDate();
        var emd = new Date(end_date.getFullYear(), em + 1, 0).getDate();
        sd = start_date.getDate() === smd ? 30 : start_date.getDate();
        if (end_date.getDate() === emd) {
            if (sd < 30) {
                em++;
                ed = 1;
            } else {
                ed = 30;
            }
        } else {
            ed = end_date.getDate();
        }
    }
    return 360 * (end_date.getFullYear() - start_date.getFullYear()) +
        30 * (em - sm) + (ed - sd);
};
// EDATE

exports.EDATE = function (start_date, months) {
    start_date = exports.parseDate(start_date);
    if (start_date instanceof Error) {
        return start_date;
    }
    if (isNaN(months)) {
        return error.value;
    }
    months = parseInt(months, 10);
    start_date.setMonth(start_date.getMonth() + months);
    return exports.serial(start_date);
};
// EOMONTH

exports.EOMONTH = function (start_date, months) {
    start_date = exports.parseDate(start_date);
    if (exports.ISERROR(start_date)) {
        return error.na;
    }

    if (months !== months) { // NaN(months)
        return error.value;
    }

    months = parseInt(months, 10);
    return exports.serial(new Date(start_date.getFullYear(), start_date.getMonth() + months + 1, 0));

};
// HOUR

// Returns the hour from 0 to 24.

exports.HOUR = function (serial_num) {
    // remove numbers before decimal place and convert fraction to 24 hour scale.
    return exports.FLOOR((serial_num - exports.FLOOR(serial_num)) * 24);
}
// ISLEAPYEAR

exports.ISLEAPYEAR = function (date) {
    date = exports.parseDate(date);
    var year = date.getFullYear();
    return (((year % 4 === 0) && (year % 100 !== 0)) ||
        (year % 400 === 0));
}
// ISOWEEKNUM

exports.ISOWEEKNUM = function (date) {
    date = exports.parseDate(date);
    if (date instanceof Error) {
        return date;
    }

    date.setHours(0, 0, 0);
    date.setDate(date.getDate() + 4 - (date.getDay() || 7));
    var yearStart = new Date(date.getFullYear(), 0, 1);
    return Math.ceil((((date - yearStart) / MilliSecondsInDay) + 1) / 7);
};
// MINUTE

// Extract minute part from serial_num.

exports.MINUTE = function (serial_num) {

    var startval = (serial_num - Math.floor(serial_num)) * SecondsInDay; // get date/time parts
    var hrs = Math.floor(startval / SecondsInHour);
    var startval = startval - hrs * SecondsInHour;

    return Math.floor(startval / 60);
}
// MONTH

exports.MONTH = function (timestamp) {
    var date;
    if (typeof timestamp === "string") {
        date = exports.parseDate(exports.DATEVALUE(timestamp));
    } else {
        date = exports.parseDate(timestamp);
    }

    if (date && !exports.ISERROR(date)) {
        return date.getMonth() + 1;
    } else {
        return date;
    }

}
// NOW

exports.NOW = function () {
    return exports.serialTime(new Date());
}
// SECOND

// Extract minute part from serial_num.

exports.SECOND = function (serial_num) {

    var startval = (serial_num - Math.floor(serial_num)) * SecondsInDay; // get date/time parts
    var hrs = Math.floor(startval / SecondsInHour);
    var startval = startval - hrs * SecondsInHour;
    var mins = Math.floor(startval / 60);
    return Math.floor(startval - mins * 60);
}
// TIME

exports.TIME = function (hour, minute, second) {
    hour = exports.parseNumber(hour);
    minute = exports.parseNumber(minute);
    second = exports.parseNumber(second);
    if (exports.isAnyError(hour, minute, second)) {
        return error.value;
    }
    if (hour < 0 || minute < 0 || second < 0) {
        return error.num;
    }
    return (((SecondsInHour * hour) + (SecondsInMinute * minute) + second) / SecondsInDay);
};
// TIMEVALUE

exports.TIMEVALUE = function (time_text) {
    // The JavaScript new Date() does not accept only time.
    // To workaround the issue we put 1/1/1900 at the front.

    var last2Characters = time_text.substr(-2).toUpperCase();
    var date;

    if (time_text.length === 7 && (last2Characters === AM || last2Characters === PM)) {
        time_text = "1/1/1900 " + time_text.substr(0, 5) + " " + last2Characters;
    } else if (time_text.length < 9) {
        time_text = "1/1/1900 " + time_text;
    }

    date = new Date(time_text);

    if (date instanceof Error) {
        return date;
    }

    return (SecondsInHour * date.getHours() +
        SecondsInMinute * date.getMinutes() +
        date.getSeconds()) / SecondsInDay;
};
// TODAY

exports.TODAY = function () {
    return exports.FLOOR(exports.NOW())
}
// YEAR

exports.YEAR = function (serial_num) {
    return Math.floor(serial_num / 365) + 1900;
}
// YEARFRAC

exports.YEARFRAC = function (start_date, end_date, basis) {
    start_date = exports.parseDate(start_date);
    if (start_date instanceof Error) {
        return start_date;
    }
    end_date = exports.parseDate(end_date);
    if (end_date instanceof Error) {
        return end_date;
    }

    basis = basis || 0;
    var sd = start_date.getDate();
    var sm = start_date.getMonth() + 1;
    var sy = start_date.getFullYear();
    var ed = end_date.getDate();
    var em = end_date.getMonth() + 1;
    var ey = end_date.getFullYear();

    switch (basis) {
        case 0:
            // US (NASD) 30/360
            if (sd === 31 && ed === 31) {
                sd = 30;
                ed = 30;
            } else if (sd === 31) {
                sd = 30;
            } else if (sd === 30 && ed === 31) {
                ed = 30;
            }
            return ((ed + em * 30 + ey * 360) - (sd + sm * 30 + sy * 360)) / 360;
        case 1:
            // Actual/actual
            var feb29Between = function (date1, date2) {
                var year1 = date1.getFullYear();
                var mar1year1 = new Date(year1, 2, 1);
                if (isLeapYear(year1) && date1 < mar1year1 && date2 >= mar1year1) {
                    return true;
                }
                var year2 = date2.getFullYear();
                var mar1year2 = new Date(year2, 2, 1);
                return (isLeapYear(year2) && date2 >= mar1year2 && date1 < mar1year2);
            };
            var ylength = 365;
            if (sy === ey || ((sy + 1) === ey) && ((sm > em) || ((sm === em) && (sd >= ed)))) {
                if ((sy === ey && isLeapYear(sy)) ||
                    feb29Between(start_date, end_date) ||
                    (em === 1 && ed === 29)) {
                    ylength = 366;
                }
                return daysBetween(start_date, end_date) / ylength;
            }
            var years = (ey - sy) + 1;
            var days = (new Date(ey + 1, 0, 1) - new Date(sy, 0, 1)) / 1000 / 60 / 60 / 24;
            var average = days / years;
            return daysBetween(start_date, end_date) / average;
        case 2:
            // Actual/360
            return daysBetween(start_date, end_date) / 360;
        case 3:
            // Actual/365
            return daysBetween(start_date, end_date) / 365;
        case 4:
            // European 30/360
            return ((ed + em * 30 + ey * 360) - (sd + sm * 30 + sy * 360)) / 360;
    }
};
// Engineering

// BIN2DEC

exports.BIN2DEC = function (value) {
    var valueAsString;

    if (typeof value === "string") {
        valueAsString = value;
    } else if (typeof value !== "undefined") {
        valueAsString = value.toString();
    } else {
        return error.NA;
    }

    if (valueAsString.length > 10) return error.NUM;

    if (valueAsString.length === 10 && valueAsString[0] === '1') {
        return parseInt(valueAsString.substring(1), 2) - 512;
    }

    // Convert binary number to decimal
    return parseInt(valueAsString, 2);

};
// BIN2HEX

// DEC2BIN

// Logical
/*
Name	Published	Alpha	Circle CI
IF	X		
NOT	X		
AND	X		
OR	X		
IFERROR	X		
IFNA	X		
XOR	X		
SWITCH	X		
CHOOSE	X	X	
IF

NOT

AND

OR

IFERROR
*/
exports.IFERROR = function (value, valueIfError) {
    if (exports.ISERROR(value)) {
        return valueIfError;
    }
    return value;
};
// IFNA

exports.IFNA = function (value, value_if_na) {
    return value === error.na ? value_if_na : value;
};
// XOR

exports.XOR = function () {
    var args = exports.flatten(arguments);
    var result = 0;
    for (var i = 0; i < args.length; i++) {
        if (args[i]) {
            result++;
        }
    }
    return (result & 1) ? true : false;
};
// SWITCH

exports.SWITCH = function () {
    var result;
    if (arguments.length > 0) {
        var targetValue = arguments[0];
        var argc = arguments.length - 1;
        var switchCount = Math.floor(argc / 2);
        var switchSatisfied = false;
        var defaultClause = argc % 2 === 0 ? null : arguments[arguments.length - 1];

        if (switchCount) {
            for (var index = 0; index < switchCount; index++) {
                if (targetValue === arguments[index * 2 + 1]) {
                    result = arguments[index * 2 + 2];
                    switchSatisfied = true;
                    break;
                }
            }
        }

        if (!switchSatisfied && defaultClause) {
            result = defaultClause;
        }
    }

    return result;
};
// CHOOSE

// MORE LISPY Fucking shit in here...haha
exports.CHOOSE = function (index) {
    if (arguments.length < 2) {
        return error.value;
    }

    var values = [];
    for (var i = 1; i < arguments.length; i++) {
        var item = arguments[i];
        if (exports.ISRANGE(item)) {
            values = values.concat(item.values());
        } else {
            values.push(item);
        }
    }

    var retVal = values[index - 1];

    if (exports.ISCELL(retVal)) {
        return (retVal).valueOf();
    }

    return retVal;

}
// Lookup & Reference

// ADDRESS

// Absolute Relative Modes:

// None
// Both
// Row
// Column
exports.ADDRESS = function (row, col, absolute_relative_mode, use_a1_notation, sheet) {
    switch (absolute_relative_mode) {
        case 0:
            return workbook.toColumn(col - 1) + (row).toString()
        case 2:
            return workbook.toColumn(col - 1) + "$" + (row).toString()
        case 3:
            return "$" + workbook.toColumn(col - 1) + (row).toString()
        default:
            return "$" + workbook.toColumn(col - 1) + "$" + (row).toString()
    }
}
// COLUMN

exports.COLUMN = function (ref) {

    if (exports.ISTEXT(ref)) {
        return workbook.extractCellInfo(ref).col;
    }

    if (exports.ISCELL(ref)) {
        return ref.col;
    }

    return error.na;
}
// COLUMNS

// Returns the number of columns in the array or range.

exports.COLUMNS = function (ref) {
    var cols = 0;

    if (exports.ISARRAY(ref)) {

        for (var i = 0; i < ref.length; i++) {
            if (cols === 0) {
                if (!exports.ISARRAY(ref[i])) {
                    return ref.length;
                }
                cols = ref[i].length;
            } else if (cols === ref[i].length) {
                continue;
            } else { // not all columns name size; raise error
                return error.value;
            }
        }

        return cols;
    }

    if (exports.ISRANGE(ref)) {
        return ref.bottomRight.colIndex - ref.topLeft.colIndex;
    }

    return error.na;

}
// LOOKUP

// The LOOKUP function supports two forms:

// Vector Form
// LOOKUP(lookup_value, lookup_vector, results_vector)

// Array Form
// LOOKUP(lookup_value, array)

exports.LOOKUP = function () {
    var lookup_value, lookup_array, lookup_vector, results_vector;
    if (arguments.length === 2) { // array form
        var wide = false;

        lookup_value = arguments[0].valueOf();
        lookup_array = arguments[1];

        if (exports.ISRANGE(lookup_array)) {
            lookup_array = lookup_array.valueOf();
        }

        for (var i = 0; i < lookup_array.length; i++) {
            if (typeof lookup_array[i] !== 'undefined' && lookup_value === lookup_array[i].valueOf()) {
                return lookup_array[i];
            }
        }

    } else if (arguments.length === 3) { // vector form`
        lookup_value = arguments[0].valueOf();
        lookup_vector = arguments[1];
        results_vector = arguments[2];

        if (exports.ISRANGE(lookup_vector)) {
            lookup_vector = lookup_vector.valueOf();
        }

        if (exports.ISRANGE(results_vector)) {
            results_vector = results_vector.valueOf();
        }

        for (var i = 0; i < lookup_vector.length; i++) {
            if (typeof lookup_vector[i] !== 'undefined' && lookup_value === lookup_vector[i].valueOf()) {
                return results_vector[i];
            }
        }

    }

    return error.na;

}
// HLOOKUP

exports.HLOOKUP = function (needle, table, index, exactmatch) {
    if (exports.ISRANGE(table)) {
        table = table.valueOf();
    }

    if (exports.ISCELL(needle)) {
        needle = needle.valueOf();
    }

    if (typeof needle === "undefined" || exports.ISBLANK(needle)) {
        return null;
    }

    index = index || 0;

    var row = table[0];
    for (var i = 0; i < row.length; i++) {

        if ((exactmatch && row[i] === needle) ||
            row[i].toLowerCase().indexOf(needle.toLowerCase()) !== -1) {
            return (index < (table.length + 1) ? table[index - 1][i] : table[0][i]);
        }
    }

    return error.na;
}
// INDIRECT

exports.INDIRECT = function (cell_string) {
    return this.cell(cell_string);
}
// VLOOKUP

// MATCH
exports.MATCH = function (lookup_reference, array_reference, matchType) {
    var lookupArray, lookupValue;
    var isRef = false;


    // Gotta have only 2 arguments folks!
    if (arguments.length === 2) {
        matchType = 1;
    }

    // Find the lookup value inside a worksheet cell, if needed.
    if (exports.ISREF(lookup_reference)) {
        lookupValue = lookup_reference.valueOf();
    } else {
        lookupValue = lookup_reference;
    }


    // Find the array inside a worksheet range, if needed.
    if (exports.ISREF(array_reference)) {
        isRef = true;
        lookupArray = array_reference.values();
    } else if (exports.ISARRAY(array_reference)) {
        lookupArray = array_reference;
    } else {
        return error.na;
    }

    // Gotta have both lookup value and array
    if (!lookupValue && !lookupArray) {
        return error.na;
    }

    // Bail on weird match types!
    if (matchType !== -1 && matchType !== 0 && matchType !== 1) {
        return error.na;
    }



    var index;
    var indexValue;


    for (var idx = 0; idx < lookupArray.length; idx++) {
        if (matchType === 1) {
            if (lookupArray[idx] === lookupValue) {
                return idx + 1;
            } else if (lookupArray[idx] < lookupValue) {
                if (!indexValue) {
                    index = idx + 1;
                    indexValue = lookupArray[idx];
                } else if (lookupArray[idx] > indexValue) {
                    index = idx + 1;
                    indexValue = lookupArray[idx];
                }
            }
        } else if (matchType === 0) {
            if (typeof lookupValue === 'string') {
                // '?' is mapped to the regex '.'
                // '*' is mapped to the regex '.*'
                // '~' is mapped to the regex '\?'
                if (idx === 0) {
                    lookupValue = "^" + lookupValue.replace(/\?/g, '.').replace(/\*/g, '.*').replace(/~/g, '\\?') + "$";
                }
                if (typeof lookupArray[idx] !== "undefined") {
                    if (String(lookupArray[idx]).toLowerCase().match(String(lookupValue).toLowerCase())) {
                        return idx + 1;
                    }
                }
            } else {
                if (typeof lookupArray[idx] !== "undefined" && lookupArray[idx] !== null && lookupArray[idx].valueOf() === lookupValue) {
                    return idx + 1;
                }
            }
        } else if (matchType === -1) {
            if (lookupArray[idx] === lookupValue) {
                return idx + 1;
            } else if (lookupArray[idx] > lookupValue) {
                if (!indexValue) {
                    index = idx + 1;
                    indexValue = lookupArray[idx];
                } else if (lookupArray[idx] < indexValue) {
                    index = idx + 1;
                    indexValue = lookupArray[idx];
                }
            }
        }
    }

    return index ? index : error.na;
};

// OFFSET

exports.OFFSET = function (ref, rows, cols, height, width) {
    var topLeft, bottomRight,
        rowsVal = 0, colsVal = 0;

    // handle case when cell object is placed in.
    // reference is the string value (e.g. A1)
    var isCell = exports.ISCELL(ref);
    var reference = isCell ? ref.addr() : ref;

    try {


        if (reference.indexOf(':') > 0) {
            topLeft = workbook.cellInfo(reference.split(':')[0]);
        } else {
            topLeft = workbook.cellInfo(reference);
        }

        // clone object to avoid messing with memorized cells
        topLeft = JSON.parse(JSON.stringify(topLeft));

        if (exports.ISBLANK(rows) || exports.ISBLANK(cols)) {
            return error.na;
        }

        rowsVal = exports.ISCELL(rows) ? rows.valueOf() : rows;
        colsVal = exports.ISCELL(cols) ? cols.valueOf() : cols;

        if (exports.ISERROR(rowsVal) || exports.ISERROR(colsVal)) {
            return error.na;
        }

        topLeft.rowIndex += rowsVal;
        topLeft.colIndex += colsVal;

        bottomRight = JSON.parse(JSON.stringify(topLeft));

        if (typeof height !== "undefined" && typeof height === "number") {
            bottomRight.rowIndex += width;
        }

        if (typeof height !== "undefined" && typeof width === "number") {
            bottomRight.colIndex += height;
        }

        var _s = function (point) { return workbook.toColumn(point.colIndex) + (point.rowIndex + 1).toString(); }
        topLeft = _s(topLeft);
        bottomRight = _s(bottomRight);
        if (topLeft === bottomRight) {
            return ref.workbook.cell(ref.sheetIndex, topLeft);
        } else {
            return ref.workbook.range(ref.sheetIndex, topLeft, bottomRight);
        }

    } catch (e) {
        return workbook.errors.value;
    }
};
// INDEX

// Find a reference with a row and column offset from an array or reference.

exports.INDEX = function (reference, row_num, column_num) {
    var cell, addr, col, row;

    if (exports.ISREF(reference)) {

        if (exports.ISRANGE(reference)) {
            cell = workbook.cellInfo(reference.topLeft);
        } else {
            if (!exports.ISCELL(reference)) { return workbook.errors.na; }
            cell = workbook.cellInfo(reference.addr());
        }

        column_num = column_num || 1;
        col = cell.colIndex + column_num - 1;
        row = cell.rowIndex + row_num;

        addr = workbook.toColumn(col) + (row);
        return reference.sheet.cell(addr);

    } else if (exports.ISARRAY(reference)) {

        column_num = column_num || 1;
        return reference[row_num - 1][column_num - 1];
    }
}
// ROW

exports.ROW = function (ref) {

    if (exports.ISTEXT(ref)) {
        return workbook.extractCellInfo(ref).row;
    }

    if (exports.ISCELL(ref)) {
        return ref.row;
    }

    return error.na;

}
// ROWS

exports.ROWS = function (ref) {
    var cols = 0;

    if (exports.ISARRAY(ref)) {

        return ref.length;
    }

    if (exports.ISRANGE(ref)) {
        return ref.bottomRight.rowIndex - ref.topLeft.rowIndex;
    }

    return error.na;

}
// Information

// CELL

exports.CELL = function (info_type, reference) {

    if (!exports.ISCELL(reference)) {
        return error.NA;
    }

    switch (info_type) {

        case "address":
            return reference.addr();
        case "col":
            return reference.colIndex + 1;
        case "row":
            return reference.row;
        case "color":
            return error.missing;
        case "contents":
            return error.missing;
        case "format":
            return error.missing;
        case "parentheses":
            return error.missing;
        case "prefix":
            return error.missing;
        case "protect":
            return error.missing;
        case "type":
            return error.missing;
        case "width":
            return error.missing;
    }

};
// DEPENDENTS

exports.DEPENDENTS = function (cell) {
    return cell.sheet.findDependents(cell).map(function (n) {
        return workbook.sheetName(cell.sheet, n.addr());
    });
};
// INFO

exports.INFO = function (text_type) {
};
// ISARRAY

exports.ISARRAY = function (arr) {
    return Object.prototype.toString.call(arr) === '[object Array]';
};
// ISBLANK

exports.ISBLANK = function (value) {
    if (exports.ISCELL(value)) {
        value = value.valueOf();
    }

    return typeof value === 'undefined' || value === null;
};
// ISBINARY

exports.ISBINARY = function (number) {
    return (/^[01]{1,10}$/).test(number);
};
// ISCELL

exports.ISCELL = function (ref) {
    return (exports.ISOBJECT(ref) && ref.constructor.name === "cell");
}
// ISEMAIL

exports.ISEMAIL = function (email) {
    var re = /^([\w-]+(?:\.[\w-]+)*)@((?:[\w-]+\.)*\w[\w-]{0,66})\.([a-z]{2,6}(?:\.[a-z]{2})?)$/i;
    return re.test(email);
};
// ISERR

exports.ISERR = function (value) {
    if (typeof value === 'undefined' || value === null) { return false; }
    value = value.valueOf();
    return ([error.value, error.ref, error.div0, error.num, error.name, error.nil]).indexOf(value) >= 0 ||
        (typeof value === 'number' && (value !== value || !isFinite(value))); // ensure numbers are not NaN or Infinity
};
// ISERROR

exports.ISERROR = function (value) {
    if (exports.ISCELL(value)) { value = value.valueOf(); }
    return exports.ISERR(value) || value === error.na;
};
// ISEVEN

exports.ISEVEN = function (value) {
    if (exports.ISCELL(value)) { value = value.valueOf(); }
    return (Math.floor(Math.abs(value)) & 1) ? false : true;
};
// ISFORMULA

exports.ISFORMULA = function (ref) {
    return exports.ISCELL(ref) && ref.workbook.cells[ref.sheetIndex][ref.cellIndex].hasOwnProperty('fid');
};
// ISFUNCTION

exports.ISFUNCTION = function (fun) {
    return fun && Object.prototype.toString.call(fun) == '[object Function]';
};
// ISLOGICAL

exports.ISLOGICAL = function (value) {
    return value === true || value === false;
};
// ISNA

exports.ISNA = function (value) {
    if (exports.ISCELL(value)) { value = value.valueOf(); }
    return value === error.na;
};
// ISNONTEXT

exports.ISNONTEXT = function (value) {
    return typeof (value) !== 'string';
};
// ISNUMBER

exports.ISNUMBER = function (value) {
    return typeof (value) === 'number' && !isNaN(value) && isFinite(value);
};
// ISODD

exports.ISODD = function (value) {
    if (exports.ISCELL(value)) { value = value.valueOf(); }
    return (Math.floor(Math.abs(value)) & 1) ? true : false;
};
// ISOBJECT

exports.ISOBJECT = function (value) {
    return (value && typeof value == 'object') || false;
};
// ISRANGE

exports.ISRANGE = function (ref) {
    return (exports.ISOBJECT(ref) && ref.constructor.name === "range");
};
// ISREF

exports.ISREF = function (ref) {
    return (exports.ISOBJECT(ref) && (ref.constructor.name === "cell" || ref.constructor.name === "range"));
};
// ISTEXT

exports.ISTEXT = function (value) {
    return typeof (value) === 'string';
};
// N

exports.N = function (value) {
    if (exports.ISNUMBER(value)) {
        return value;
    }
    if (value instanceof Date) {
        return value.getTime();
    }
    if (value === true) {
        return 1;
    }
    if (value === false) {
        return 0;
    }
    if (exports.ISERROR(value)) {
        return value;
    }
    return 0;
};
// NA

exports.NA = function () {
    return error.na;
};
// PRECEDENTS

exports.PRECEDENTS = function (cell) {

    if (cell.fid) {
        return cell.workbook.functions[cell.fid].precedents;
    }

    return [];

};
// SHEET

exports.SHEET = function (name) {
    return this.workbook.get(name);
};
// SHEETS

exports.SHEETS = function () {
    return Object.keys(this.workbook.sheets).length;
};
// TYPE

exports.TYPE = function (value) {
    if (exports.ISLOGICAL(value)) {
        return 4;
    }
    else if (exports.ISTEXT(value)) {
        return 2;
    }
    else if (exports.ISNUMBER(value)) {
        return 1;
    }
    else if (exports.ISERROR(value)) {
        return 16;
    }
    else if (exports.ISARRAY(value)) {
        return 64;
    }
};
// Misc Functions

// NUMBERS

exports.NUMBERS = function () {
    var possibleNumbers = exports.flatten(arguments);

    var arr = [], n;

    for (var i = 0; i < possibleNumbers.length; i++) {
        n = possibleNumbers[i];
        if (exports.ISNUMBER(n)) {
            arr.push(n);
        }
    }

    return arr;

};
// UNIQUE

exports.UNIQUE = function () {
    var result = [];
    var range = exports.flatten(arguments)
    for (var i = 0; i < range.length; ++i) {
        var hasElement = false;
        var element = range[i];

        // Check if we've already seen this element.
        for (var j = 0; j < result.length; ++j) {
            hasElement = result[j] === element;
            if (hasElement) { break; }
        }

        // If we did not find it, add it to the result.
        if (!hasElement) {
            result.push(element);
        }
    }
    return result;
};
// Financial

// ACCRINT

exports.ACCRINT = function (issue, first, settlement, rate, par, frequency, basis) {
    // Return error if either date is invalid
    var issueDate = exports.parseDate(exports.DATEVALUE(issue.valueOf()));
    var firstDate = exports.parseDate(exports.DATEVALUE(first.valueOf()));
    var settlementDate = exports.parseDate(exports.DATEVALUE(settlement.valueOf()));

    // Set default values
    par = par.valueOf() || 0;
    basis = basis.valueOf() || 0;
    rate = rate.valueOf();

    var validDate = exports.validDate;

    if (!validDate(issueDate) || !validDate(firstDate) || !validDate(settlementDate)) {
        return error.value;
    }


    // Return error if either rate or par are lower than or equal to zero
    if (rate <= 0 || par <= 0) {
        return error.num;
    }

    // Return error if frequency is neither 1, 2, or 4
    if ([1, 2, 4].indexOf(frequency.valueOf()) === -1) {
        return error.num;
    }

    // Return error if basis is neither 0, 1, 2, 3, or 4
    if ([0, 1, 2, 3, 4].indexOf(basis) === -1) {
        return error.num;
    }

    // Return error if settlement is before or equal to issue
    if (settlementDate <= issueDate) {
        return error.num;
    }


    // Compute accrued interest
    return par * rate * exports.YEARFRAC(issueDate, settlementDate, basis);
};
// CUMIPMT

exports.CUMIPMT = function (rate, periods, value, start, end, type) {
    // Credits: algorithm inspired by Apache OpenOffice
    // Credits: Hannes Stiebitzhofer for the translations of function and variable names

    rate = exports.parseNumber(rate);
    periods = exports.parseNumber(periods);
    value = exports.parseNumber(value);
    if (exports.isAnyError(rate, periods, value)) {
        return error.value;
    }

    // Return error if either rate, periods, or value are lower than or equal to zero
    if (rate <= 0 || periods <= 0 || value <= 0) {
        return error.num;
    }

    // Return error if start < 1, end < 1, or start > end
    if (start < 1 || end < 1 || start > end) {
        return error.num;
    }

    // Return error if type is neither 0 nor 1
    if (type !== 0 && type !== 1) {
        return error.num;
    }

    // Compute cumulative interest
    var payment = exports.PMT(rate, periods, value, 0, type);
    var interest = 0;

    if (start === 1) {
        if (type === 0) {
            interest = -value;
            start++;
        }
    }

    for (var i = start; i <= end; i++) {
        if (type === 1) {
            interest += exports.FV(rate, i - 2, payment, value, 1) - payment;
        } else {
            interest += exports.FV(rate, i - 1, payment, value, 0);
        }
    }
    interest *= rate;

    // Return cumulative interest
    return interest;
};
// PMT

exports.PMT = function (rate, periods, present, future, type) {

    future = exports.N(future) || 0;
    type = exports.N(type) || 0;

    var payment;
    if (rate === 0) {
        payment = (present + future) / periods;
    } else {
        var term = Math.pow(1 + rate, periods);
        if (type === 1) {
            payment = (future * rate / (term - 1) + present * rate / (1 - 1 / term)) / (1 + rate);
        } else {
            payment = future * rate / (term - 1) + present * rate / (1 - 1 / term);
        }
    }
    return -payment;
};
// FV

exports.FV = function (rate, periods, payment, value, type) {

    if (typeof rate === 'undefined') throw Error("rate is undefined");
    if (typeof periods === 'undefined') throw Error("rate is undefined");
    if (typeof payment === 'undefined') throw Error("rate is undefined");

    value = value || 0;
    type = type || 0;


    var fv;
    if (rate === 0) {
        fv = value + payment * periods;
    } else {
        var term = Math.pow(1 + rate, periods);
        if (type === 1) {
            fv = value * term + payment * (1 + rate) * (term - 1) / rate;
        } else {
            fv = value * term + payment * (term - 1) / rate;
        }
    }
    return -fv;
};
// PV

exports.PV = function (rate, periods, payment, future, type) {
    if (typeof rate === 'undefined') throw Error("rate is undefined");
    if (typeof periods === 'undefined') throw Error("rate is undefined");
    if (typeof payment === 'undefined') throw Error("rate is undefined");

    future = future || 0;
    type = type || 0;

    if (rate === 0) {
        return -payment * periods - future;
    } else {
        return (((1 - Math.pow(1 + rate, periods)) / rate) * payment * (1 + rate * type) - future) / Math.pow(1 + rate, periods);
    }
};
// NPV

exports.NPV = function (rate) {
    rate = rate * 1;
    var factor = 1,
        sum = 0;

    for (var i = 1; i < arguments.length; i++) {
        var factor = factor * (1 + rate);
        sum += arguments[i] / factor;
    }

    return sum;
}
// IPMT

exports.IPMT = function (rate, per, nper, pv, fv, type) {
    var pmt = exports.PMT(rate, nper, pv, fv, type),
        fv = exports.FV(rate, per - 1, pmt, pv, type),
        result = fv * rate;

    // account for payments at beginning of period versus end.
    if (type) {
        result /= (1 + rate);
    }

    return result;
}
// NPER

exports.NPER = function (rate, pmt, pv, fv, type) {
    var log,
        result;
    rate = parseFloat(rate || 0);
    pmt = parseFloat(pmt || 0);
    pv = parseFloat(pv || 0);
    fv = (fv || 0);
    type = (type || 0);

    log = function (prim) {
        if (isNaN(prim)) {
            return Math.log(0);
        }
        var num = Math.log(prim);
        return num;
    }

    if (rate == 0.0) {
        result = (-(pv + fv) / pmt);
    } else if (type > 0.0) {
        result = (log(-(rate * fv - pmt * (1.0 + rate)) / (rate * pv + pmt * (1.0 + rate))) / (log(1.0 + rate)));
    } else {
        result = (log(-(rate * fv - pmt) / (rate * pv + pmt)) / (log(1.0 + rate)));
    }

    if (isNaN(result)) {
        result = 0;
    }

    return result;
}
// Math

// ABS

exports.ABS = function (v) { return Math.abs(v); }
// ACOS

exports.ACOS = function (number) {

    if (!exports.ISNUMBER(number)) {
        return error.value;
    }

    return Math.acos(number);
};
// ACOSH

exports.ACOSH = function (number) {

    if (!exports.ISNUMBER(number)) {
        return error.value;
    }

    return Math.log(number + Math.sqrt(number * number - 1));
};
// ACOT

exports.ACOT = function (number) {

    if (!exports.ISNUMBER(number)) {
        return error.value;
    }

    return Math.atan(1 / number);
};

// ACOTH

exports.ACOTH = function (number) {

    if (!exports.ISNUMBER(number)) {
        return error.value;
    }

    return 0.5 * Math.log((number + 1) / (number - 1));
};
// ADD

exports.ADD = function (a, b) {

    // force conversion to number
    a = +a;
    b = +b;

    // check for NaN
    if (a !== a || b !== b) {
        return error.value;
    }

    return a + b;
}
// ASIN

exports.ASIN = function (number) {

    if (!exports.ISNUMBER(number)) {
        return error.value;
    }

    return Math.asin(number);
};
// ASINH

exports.ASINH = function (number) {

    if (!exports.ISNUMBER(number)) {
        return error.value;
    }

    return Math.log(number + Math.sqrt(number * number + 1));
};
// ATAN

exports.ATAN = function (number) {

    if (!exports.ISNUMBER(number)) {
        return error.value;
    }

    return Math.atan(number);
};
// ATAN2

exports.ATAN2 = function (number_x, number_y) {

    if (!exports.ISNUMBER(number_x)) {
        return error.value;
    }

    if (!exports.ISNUMBER(number_y)) {
        return error.value;
    }


    return Math.atan2(number_x, number_y);
};
// ATANH

exports.ATANH = function (number) {

    if (!exports.ISNUMBER(number)) {
        return error.value;
    }

    return Math.log((1 + number) / (1 - number)) / 2;
};
// BASE

exports.BASE = function (number, radix, min_length) {
    min_length = min_length || 0;

    number = exports.parseNumber(number);
    radix = exports.parseNumber(radix);
    min_length = exports.parseNumber(min_length);
    if (exports.isAnyError(number, radix, min_length)) {
        return error.value;
    }
    min_length = (min_length === undefined) ? 0 : min_length;
    var result = number.toString(radix);
    return new Array(Math.max(min_length + 1 - result.length, 0)).join('0') + result;

}
// CEILING

exports.CEILING = function (number, significance, mode) {
    significance = (significance === undefined) ? 1 : Math.abs(significance);
    mode = mode || 0;

    number = exports.parseNumber(number);
    significance = exports.parseNumber(significance);
    mode = exports.parseNumber(mode);
    if (exports.isAnyError(number, significance, mode)) {
        return error.value;
    }
    if (significance === 0) {
        return 0;
    }
    var precision = -Math.floor(Math.log(significance) / Math.log(10));
    if (number >= 0) {
        return exports.ROUND(Math.ceil(number / significance) * significance, precision);
    } else {
        if (mode === 0) {
            return -exports.ROUND(Math.floor(Math.abs(number) / significance) * significance, precision);
        } else {
            return -exports.ROUND(Math.ceil(Math.abs(number) / significance) * significance, precision);
        }
    }
};
// CONBIN

exports.COMBIN = function (number, number_chosen) {
    number = exports.parseNumber(number);
    number_chosen = exports.parseNumber(number_chosen);
    if (exports.isAnyError(number, number_chosen)) {
        return error.value;
    }
    return exports.FACT(number) / (exports.FACT(number_chosen) * exports.FACT(number - number_chosen));
};
// COS

exports.COS = function (number) {

    if (!exports.ISNUMBER(number)) {
        return error.value;
    }

    return Math.cos(number);
};
// COSH

exports.COSH = function (number) {

    if (!exports.ISNUMBER(number)) {
        return error.value;
    }

    return (Math.exp(number) + Math.exp(-number)) / 2;
};
// COT

exports.COT = function (number) {

    if (!exports.ISNUMBER(number)) {
        return error.value;
    }

    return 1 / Math.tan(number);
};
// COTH

exports.COTH = function (number) {

    if (!exports.ISNUMBER(number)) {
        return error.value;
    }

    var e2 = Math.exp(2 * number);
    return (e2 + 1) / (e2 - 1);

};
// CSC

exports.CSC = function (number) {

    if (!exports.ISNUMBER(number)) {
        return error.value;
    }

    return 1 / Math.sin(number);
};
// CSCH

exports.CSCH = function (number) {

    if (!exports.ISNUMBER(number)) {
        return error.value;
    }

    return 2 / (Math.exp(number) - Math.exp(-number));
};
// DECIMAL

exports.DECIMAL = function (number, radix) {
    if (arguments.length < 1) {
        return error.value;
    }

    return parseInt(number, radix);
};
// DEGREES

exports.DEGREES = function (number) {
    if (!exports.ISNUMBER(number)) {
        return error.value;
    }

    return number * 180 / Math.PI;
};
// DIVIDE

exports.DIVIDE = function (a, b) {

    // force conversion to number
    a = +a;
    b = +b;

    // check for NaN
    if (a !== a || b !== b) {
        return error.value;
    }

    return a / b;
}
// EVEN

exports.EVEN = function (number) {
    number = exports.parseNumber(number);
    if (number instanceof Error) {
        return number;
    }
    return exports.CEILING(number, -2, -1);
};
// EQ

exports.EQ = function (a, b) {

    if (exports.ISCELL(a)) {
        a = a.valueOf();
    }

    if (exports.ISCELL(b)) {
        b = b.valueOf();
    }

    // Unlike the host language the string comparisions are 
    if (typeof a === "string" && typeof b === "string") {
        return a.toLowerCase() === b.toLowerCase()
    } else {
        return a === b;
    }

}
// EXP

exports.EXP = function (num) {
    return Math.pow(Math.E, num);
}
// FACT

var MEMOIZED_FACT = [];
exports.FACT = function (number) {
    number = exports.parseNumber(number);
    if (number instanceof Error) {
        return number;
    }
    var n = Math.floor(number);
    if (n === 0 || n === 1) {
        return 1;
    } else if (MEMOIZED_FACT[n] > 0) {
        return MEMOIZED_FACT[n];
    } else {
        MEMOIZED_FACT[n] = exports.FACT(n - 1) * n;
        return MEMOIZED_FACT[n];
    }
};
// FACTDOUBLE

exports.FACTDOUBLE = function (number) {
    number = exports.parseNumber(number);
    if (number instanceof Error) {
        return number;
    }
    var n = Math.floor(number);
    if (n <= 0) {
        return 1;
    } else {
        return n * exports.FACTDOUBLE(n - 2);
    }
};
// FLOOR

exports.FLOOR = function (value, significance) {
    significance = significance || 1;

    if (
        (value < 0 && significance > 0)
        || (value > 0 && significance < 0)
    ) {
        var result = new Number(0);
        return result;
    }
    if (value >= 0) {
        return Math.floor(value / significance) * significance;
    } else {
        return Math.ceil(value / significance) * significance;
    }
}

// GCD

// adapted http://rosettacode.org/wiki/Greatest_common_divisor#JavaScript
exports.GCD = function () {
    var range = exports.parseNumberArray(exports.flatten(arguments));
    if (range instanceof Error) {
        return range;
    }
    var n = range.length;
    var r0 = range[0];
    var x = r0 < 0 ? -r0 : r0;
    for (var i = 1; i < n; i++) {
        var ri = range[i];
        var y = ri < 0 ? -ri : ri;
        while (x && y) {
            if (x > y) {
                x %= y;
            } else {
                y %= x;
            }
        }
        x += y;
    }
    return x;
};
// GT

exports.GT = function (a, b) {

    if (exports.ISCELL(a)) {
        a = a.valueOf();
    }

    if (exports.ISCELL(b)) {
        b = b.valueOf();
    }

    return a > b;

}
// GTE

exports.GTE = function (a, b) {

    if (exports.ISCELL(a)) {
        a = a.valueOf();
    }

    if (exports.ISCELL(b)) {
        b = b.valueOf();
    }

    return a >= b;

}
// LCM

exports.LCM = function () {
    // Credits: Jonas Raoni Soares Silva
    var o = exports.parseNumberArray(exports.flatten(arguments));
    if (o instanceof Error) {
        return o;
    }
    for (var i, j, n, d, r = 1;
	       (n = o.pop()) !== undefined;) {
        while (n > 1) {
            if (n % 2) {
                for (i = 3, j = Math.floor(Math.sqrt(n)); i <= j && n % i; i += 2) {
                    //empty
                }
                d = (i <= j) ? i : n;
            } else {
                d = 2;
            }
            for (n /= d, r *= d, i = o.length; i;
                (o[--i] % d) === 0 && (o[i] /= d) === 1 && o.splice(i, 1)) {
                //empty
            }
        }
    }
    return r;
};
// LN

exports.LN = function (number) {

    if (!exports.ISNUMBER(number)) {
        return error.value;
    }

    return Math.log(number);
};
// LOG

exports.LOG = function (number, base) {

    base = base || 10;

    if (!exports.ISNUMBER(number)) {
        return error.value;
    }

    if (!exports.ISNUMBER(base)) {
        return error.value;
    }

    return Math.log(number) / Math.log(base);

};
// LOG10

exports.LOG10 = function (number) {
    if (!exports.ISNUMBER(number)) {
        return error.value;
    }

    return Math.log(number) / Math.log(10);
}
// LT

exports.LT = function (a, b) {

    if (exports.ISCELL(a)) {
        a = a.valueOf();
    }

    if (exports.ISCELL(b)) {
        b = b.valueOf();
    }

    return a < b;

}
// LTE

exports.LTE = function (a, b) {

    if (exports.ISCELL(a)) {
        a = a.valueOf();
    }

    if (exports.ISCELL(b)) {
        b = b.valueOf();
    }

    return a <= b;

}
// INT

exports.INT = function (num) {
    return Math.floor(num);
}
// MINUS

exports.MINUS = function (a, b) {

    // force conversion to number
    a = +a;
    b = +b;

    // check for NaN
    if (a !== a || b !== b) {
        return error.value;
    }

    return a - b;
}
// MOD

exports.MOD = function (a, b) {
    return a % b;
}
// MULTIPLY

exports.MULTIPLY = function (a, b) {

    // force conversion to number
    a = +a;
    b = +b;

    // check for NaN
    if (a !== a || b !== b) {
        return error.value;
    }

    return a * b;
}
// NE

exports.NE = function (a, b) {

    if (exports.ISCELL(a)) {
        a = a.valueOf();
    }

    if (exports.ISCELL(b)) {
        b = b.valueOf();
    }

    return a !== b;

}
// PI

exports.PI = function () { return Math.PI };
// POWER

exports.POW = exports.POWER = function (val, nth) {
    return Math.pow(val, nth);
}
// PRODUCT

exports.PRODUCT = function () {
    var range = exports.flatten(arguments);
    var result = range[0];

    for (var i = 1; i < range.length; i++) {
        result *= range[i];
    }

    return result;
}
// QUOTIENT

exports.QUOTIENT = function (a, b) {
    var q = Math.floor(a / b);
    if (q !== q) {
        return error.value;
    }
    return q;
}
// RADIANS

exports.RADIANS = function (number) {
    number = exports.parseNumber(number);
    if (number instanceof Error) {
        return number;
    }
    return number * Math.PI / 180;
};
// RAND

exports.RAND = function () {
    return Math.random();
};
// RANDBETWEEN

exports.RANDBETWEEN = function (bottom, top) {
    bottom = exports.parseNumber(bottom);
    top = exports.parseNumber(top);
    if (exports.isAnyError(bottom, top)) {
        return error.value;
    }
    // Creative Commons Attribution 3.0 License
    // Copyright (c) 2012 eqcode
    return bottom + Math.ceil((top - bottom + 1) * Math.random()) - 1;
};
// ROUND

exports.ROUNDUP = function (number, precision) {
    var factors = [1, 10, 100, 1000, 10000, 100000, 1000000, 10000000, 100000000, 1000000000];
    var factor = factors[precision];
    if (number > 0) {
        return Math.ceil(number * factor) / factor;
    } else {
        return Math.floor(number * factor) / factor;
    }
}

// ROUNDUP

exports.ROUNDUP = function (number, precision) {
    var factors = [1, 10, 100, 1000, 10000, 100000, 1000000, 10000000, 100000000, 1000000000];
    var factor = factors[precision];
    if (number > 0) {
        return Math.ceil(number * factor) / factor;
    } else {
        return Math.floor(number * factor) / factor;
    }
}
// SUM

exports.SUM = function () {
    var numbers = exports.flatten(arguments);
    var result = 0;
    for (var i = 0; i < numbers.length; i++) {
        if (numbers[i] instanceof Array) {
            for (var j = 0; j < numbers[i].length; j++) {
                result += (exports.ISNUMBER(numbers[i][j])) ? numbers[i][j] : exports.ISCELL(numbers[i][j]) ? numbers[i][j].valueOf() : 0;
            }
        } else {
            result += (exports.ISNUMBER(numbers[i])) ? numbers[i] : exports.ISCELL(numbers[i]) ? numbers[i].valueOf() : 0;
        }
    }

    return result;
};

// SUMIF

exports.SUMIF = function (range, criteria) {
    range = exports.flatten(range);
    var result = 0;
    for (var i = 0; i < range.length; i++) {
        result += (eval(range[i] + criteria)) ? range[i] : 0;
    }
    return result;
}
// Stats

// AVERAGE

exports.AVERAGE = function () {
    var set = exports.flatten.apply(this, arguments);
    return exports.SUM(set) / set.length;
}
// COUNT

exports.COUNT = function () {
    var count = 0,
        v = arguments,
        i = v.length - 1;

    if (i < 0) {
        return count;
    }

    do {
        if (v[i] !== null) {
            count++;
        }
    } while (i--);

    return count;
}
// COUNTIF

exports.COUNTIF = function (range, criteria) {
    range = exports.flatten(range);
    if (!/[<>=!]/.test(criteria)) {
        criteria = '=="' + criteria + '"';
    }
    var matches = 0;
    for (var i = 0; i < range.length; i++) {
        if (typeof range[i] !== 'string') {

            if (exports.ISERROR(range[i])) { continue; }
            if (eval(range[i] + criteria)) {
                matches++;
            }
        } else {
            if (eval('"' + range[i] + '"' + criteria)) {
                matches++;
            }
        }
    }
    return matches;
};
// MIN

// MAX

// SLOPE

exports.SLOPE = function (data_y, data_x) {
    /*
       data_y = utils.parseNumberArray(utils.flatten(data_y));
       data_x = utils.parseNumberArray(utils.flatten(data_x));
 
       if (utils.anyIsError(data_y, data_x)) {
           return error.value;
       }
       var xmean = jStat.mean(data_x);
       var ymean = jStat.mean(data_y);
       
       var n = data_x.length;
 
       var num = 0;
       var den = 0;
 
       for (var i = 0; i < n; i++) {
           num += (data_x[i] - xmean) * (data_y[i] - ymean);
           den += Math.pow(data_x[i] - xmean, 2);
       }
 
       return num / den;
    */
};
// Text

// CHAR

exports.CHAR = function (number) {
    number = exports.parseNumber(number);
    if (number instanceof Error) {
        return number;
    }
    return String.fromCharCode(number);
};
// CLEAN

// Removes the first 32 non-printing characters from the ASCII character set.

exports.CLEAN = function (text) {
    var out = "";

    for (var i = 0; i < text.length; i++) {
        if (text.charCodeAt(i) > 31) {
            out += text[i];
        }
    }

    return out;
}
// CODE

exports.CODE = function (text) {
    text = text || '';
    return text.charCodeAt(0);
};
// CONCATENATE

exports.CONCAT = exports.CONCATENATE = function () {
    var args = exports.flatten(arguments);

    for (var i = 0; i < args.length; i++) {
        if (args[i] === true) { args[i] = 'TRUE'; }
        else if (args[i] === false) { args[i] = 'FALSE'; }
        else { args[i] = args[i] ? args[i].valueOf() : '' }
    }

    return args.join('');
};
// DOLLAR

exports.DOLLAR = function (num) {
    return exports.TEXT(num, '$#,##0.00_);($#,##0.00)');
}
// EXACT

exports.EXACT = function (a, b) {

    if (exports.ISCELL(a)) {
        a = a.valueOf();
    }

    if (exports.ISCELL(b)) {
        b = b.valueOf();
    }

    if (typeof a !== "string" || typeof b !== "string") {
        return error.na;
    }

    return a === b;
}
// FIND

exports.FIND = function (find_text, within_text, position) {
    if (!within_text) { return null; }
    position = (typeof position === 'undefined') ? 1 : position;
    position = within_text.indexOf(find_text, position - 1) + 1;
    return position === 0 ? workbook.errors.value : position;
}
// FIXED

exports.FIXED = function (num, decimals, commas) {
    if (commas) { return num.toFixed(decimals).replace(/\B(?=(\d{3})+(?!\d))/g, ","); }
    return num.toFixed(decimals);
}
// LEFT

exports.LEFT = function (text, number) {
    number = (number === undefined) ? 1 : number;
    number = exports.parseNumber(number);
    if (number instanceof Error || typeof text !== 'string') {
        return error.value;
    }
    return text ? text.substring(0, number) : null;
};
// LEN

LEN = function (text) {
    if (arguments.length === 0) {
        return error.error;
    }

    if (typeof text === 'string') {
        return text ? text.length : 0;
    }

    if (text.length) {
        return text.length;
    }

    return error.value;
};
// LOWER

exports.LOWER = function (text) {
    if (typeof text === 'undefined' || text === null) return "";
    return text.toLowerCase();
}
// JOIN

exports.JOIN = function (array, separator) {
    return array.join(separator);
};
// MID

exports.MID = function (text, start, number) {
    start = exports.parseNumber(start);
    number = exports.parseNumber(number);
    if (exports.isAnyError(start, number) || typeof text !== 'string') {
        return number;
    }
    return text.substring(start - 1, number + 1);
};
// PROPER

exports.PROPER = function (text) {
    if (text === undefined || text.length === 0) {
        return error.value;
    }
    if (text === true) {
        text = 'TRUE';
    }
    if (text === false) {
        text = 'FALSE';
    }
    if (isNaN(text) && typeof text === 'number') {
        return error.value;
    }
    if (typeof text === 'number') {
        text = '' + text;
    }

    return text.replace(/\w\S*/g, function (txt) {
        return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
    });
};
// REPLACE

exports.REPLACE = function (text, position, length, new_text) {
    position = exports.parseNumber(position);
    length = exports.parseNumber(length);
    if (exports.isAnyError(position, length) ||
        typeof text !== 'string' ||
        typeof new_text !== 'string') {
        return error.value;
    }
    return text.substr(0, position - 1) + new_text + text.substr(position - 1 + length);
};
// REPT

exports.REPT = function (t, n) {
    var r = "";
    for (var i = 0; i < n; i++) {
        r += t;
    }
    return r;
}
// RIGHT

exports.RIGHT = function (text, number) {
    number = (number === undefined) ? 1 : number;
    number = exports.parseNumber(number);
    if (number instanceof Error) {
        return number;
    }
    return text ? text.substring(text.length - number) : null;
};
// SEARCH

exports.SEARCH = function (find_text, within_text, position) {
    if (!within_text) { return null; }
    position = (typeof position === 'undefined') ? 1 : position;

    // The SEARCH function translated the find_text into a regex.
    var find_exp = find_text
        .replace(/([^~])\?/g, '$1.')   // convert ? into .
        .replace(/([^~])\*/g, '$1.*')  // convert * into .*
        .replace(/([~])\?/g, '\\?')    // convert ~? into \?
        .replace(/([~])\*/g, '\\*');   // convert ~* into \*

    position = new RegExp(find_exp, "i").exec(within_text);

    if (position) { return position.index + 1 }
    return workbook.errors.value;
}
// SPLIT

exports.SPLIT = function (text, separator) {
    return text.split(separator);
};
// SUBSTITUTE

exports.SUBSTITUTE = function (text, old_text, new_text, occurrence) {
    if (!text || !old_text || !new_text) {
        return text;
    } else if (occurrence === undefined) {
        return text.replace(new RegExp(old_text, 'g'), new_text);
    } else {
        var index = 0;
        var i = 0;
        while (text.indexOf(old_text, index) > 0) {
            index = text.indexOf(old_text, index + 1);
            i++;
            if (i === occurrence) {
                return text.substring(0, index) + new_text + text.substring(index + old_text.length);
            }
        }
    }
};
// T

exports.T = function (value) {
    return (typeof value === "string") ? value : '';
};
// TEXT

exports.TEXT = function (value, format) {
    return workbook.FormatNumber.formatNumberWithFormat(value, format);
};
// TRIM

exports.TRIM = function (text) {
    if (typeof text !== 'string') {
        return error.value;
    }
    return text.replace(/ +/g, ' ').trim();
};
// UPPER

exports.UPPER = function (text) {
    if (typeof text === 'undefined' || text === null) return "";
    return text.toUpperCase();
}
// VALUE

// This is hopelessly inadequate.

exports.VALUE = function (t) {
    return Number.parseFloat(t);
}
// Formula Parser Code

