var fn = require('./utils');
var error = require('./error');
var parser = require('./parser');
var compiledNumber = 0;

module.exports = workbook = {};

// these functions always need to use .apply(ws, ...)
workbook.CONTEXTUAL_FUNCTIONS = ["SHEETS", "SHEET", "INDIRECT"];
workbook.ALWAYS_CONTEXTUAL = false; // always use .apply to set this to worksheet

// Define function library container
workbook.fn = fn;
workbook.errors = error;

// Matches $A$1, $A1, A$1 or A1. 
var cellRegex = /^(?:[$])?([a-zA-Z]+)(?:[$])?([0-9]+)$/;
var cellInfoMem = {};
workbook.cellInfo = function (addr) {

    if (cellInfoMem.hasOwnProperty(addr)) {
        return cellInfoMem[addr];
    }

    var matches = addr.match(cellRegex);

    if (matches === null) {
        throw Error("Ref does not match cell reference: " + addr);
    } else {
        var result = {
            addr: addr,
            row: +matches[2],
            col: matches[1],
            colIndex: workbook.toColumnIndex(matches[1]),
            rowIndex: +matches[2] - 1
        };
        cellInfoMem[addr] = result;
        return result;
    }
}

var precedents, suppress = false;
workbook.compile = function (exp, mode) {
    var ast = exp,
        jsCode,
        functionCode,
        f,
        context = this;

    mode = mode || 1; // default to compile JSFunction

    // convert to AST when string provided
    if (typeof ast === 'string') {
        ast = workbook.parse(this, ast);
    }

    precedents = []; // reset shared precedents
    jsCode = workbook.compiler(ast);

    switch (mode) {
        case 1:
            var id = compiledNumber++;
            f = Function("context", "// formula: " + exp + "\nreturn " + jsCode + "\n//# sourceURL=formula_function_" +
                id + ".js");
            f.id = id;
            f.js = jsCode;
            f.exp = exp;
            f.ast = ast;
            f.precedents = precedents;

            return f;
        case 2:
            return jsCode;
        case 3:
            functionCode = "// formula: " + exp + "\nfunction(context) {\n  return " + jsCode + ";\n}";
            return functionCode;
        case 4:
            return precedents;
    }

}

// define a compiler function to handle recurse the AST.
workbook.compiler = function (node) {

    var lhs, rhs, _name, dynamic = false;

    var compiler = workbook.compiler;

    // The node is expected to be either an operator, function or a value.
    switch (node.type) {
        case 'operator':
            switch (node.subtype) {
                case 'prefix-plus':
                    return '+' + compiler(node.operands[0]);
                case 'prefix-minus':
                    return '-' + compiler(node.operands[0]);
                case 'infix-add':
                    return ("workbook.fn.ADD(" + compiler(node.operands[0]) + ',' +
                        compiler(node.operands[1]) + ")");
                case 'infix-subtract':
                    return ("workbook.fn.MINUS(" + compiler(node.operands[0]) + ',' +
                        compiler(node.operands[1]) + ")");
                case 'infix-multiply':
                    return ("workbook.fn.MULTIPLY(" + compiler(node.operands[0]) + ',' +
                        compiler(node.operands[1]) + ")");
                case 'infix-divide':
                    return ("workbook.fn.DIVIDE(" + compiler(node.operands[0]) + ',' +
                        compiler(node.operands[1]) + ")");
                case 'infix-power':
                    return ('workbook.fn.POWER(' + compiler(node.operands[0]) + ','
                        + compiler(node.operands[1]) + ')')
                case 'infix-concat':
                    lhs = compiler(node.operands[0]);
                    rhs = compiler(node.operands[1]);

                    return "workbook.fn.CONCAT(" + workbook.wrapString(lhs) + ', ' + workbook.wrapString(rhs) + ")";
                case 'infix-eq':
                    return ("workbook.fn.EQ(" + compiler(node.operands[0]) + ',' +
                        compiler(node.operands[1]) + ")");
                case 'infix-ne':
                    return ("workbook.fn.NE(" + compiler(node.operands[0]) + ',' +
                        compiler(node.operands[1]) + ")");
                case 'infix-gt':
                    return ("workbook.fn.GT(" + compiler(node.operands[0]) + ',' +
                        compiler(node.operands[1]) + ")");
                case 'infix-gte':
                    return ("workbook.fn.GTE(" + compiler(node.operands[0]) + ',' +
                        compiler(node.operands[1]) + ")");
                case 'infix-lt':
                    return ("workbook.fn.LT(" + compiler(node.operands[0]) + ',' +
                        compiler(node.operands[1]) + ")");
                case 'infix-lte':
                    return ("workbook.fn.LTE(" + compiler(node.operands[0]) + ',' +
                        compiler(node.operands[1]) + ")");
            }
            throw TypeException("Unknown operator: " + node.subtype);
        case 'group':
            return ('(' + compiler(node.exp) + ')');
        case 'function':
            switch (node.name) {
                case 'IF':
                    if (node.args.length > 3) { throw Error("IF sent too many arguments."); }
                    if (node.args.length !== 3) { throw Error("IF expects 3 arguments"); }
                    return ('((' + compiler(node.args[0]) +
                        ')?' + compiler(node.args[1]) +
                        ':' + compiler(node.args[2]) + ')');

                case 'NOT':
                    if (node.args.length !== 1) { throw Error("NOT only accepts one argument"); }
                    return 'workbook.fn.NOT(' + compiler(node.args[0]) + ')';
                case 'AND':
                    return ('workbook.fn.AND(' +
                        node.args.map(function (n) {
                            return compiler(n);
                        }).join(', ') + ')');
                case 'OR':
                    return ('workbook.fn.OR(' +
                        node.args.map(function (n) {
                            return compiler(n);
                        }).join(', ') + ')');

                default:

                    _name = function (name) {
                        return workbook.fn.hasOwnProperty(name) ? "workbook.fn." + name : name;
                    }


                    if (workbook.ALWAYS_CONTEXTUAL || workbook.CONTEXTUAL_FUNCTIONS.indexOf(node.name) >= 0) {
                        return (_name(node.name) + '.apply( context, ' +
                            (node.args.length > 0 ? "[" + node.args.map(function (n) {
                                return compiler(n);
                            }).join(',') + '] )' : '[] )'));
                    } else {
                        return (_name(node.name) + '( ' + node.args.map(function (n) {
                            return compiler(n);
                        }).join(',') + ' )');
                    }


            }
        case 'cell':
            if (typeof precedents !== "undefined" && !suppress) { precedents.push(node); }

            if (node.subtype === "remote") {
                return 'context.ref(\"' + node.worksheet + '\", \"' + node.addr + '\")';
            } else {
                return 'context.ref(\"' + node.addr + '\")';
            }
        case 'range':

            if (typeof precedents !== "undefined") { precedents.push(node); suppress = true; }
            node.bottomRight.subtype = node.topLeft.subtype;
            node.bottomRight.worksheet = node.topLeft.worksheet;
            lhs = compiler(node.topLeft);
            rhs = compiler(node.bottomRight);
            suppress = false;

            // anonymous functions are the perfect solution for dynamic ranges but was not immediately obvious to me
            if (node.topLeft.type === "function") {
                lhs = "function() { return (" + lhs + "); }"
            }

            if (node.bottomRight.type === "function") {
                rhs = "function() { return (" + rhs + "); }"
            }

            return ('context.range( ' + lhs + ', ' + rhs + ' )');

        case 'value':
            switch (node.subtype) {
                case 'array':
                    return ('[' +
                        node.items.map(function (n) {
                            return compiler(n);
                        }).join(',') + ']');
                case 'string':
                    return "'" + node.value.replace(/'/g, "''") + "'";
                case 'variable':

                    if (precedents && !suppress) { precedents.push(node); }

                    if (node.subtype === "remote-named") {
                        return 'context.ref(\"' + node.worksheet + '\", \"' + node.value + '\")';
                    } else {
                        return 'context.ref(\"' + node.value + '\")';
                    }


                default:
                    return node.value;
            }
    }
}


workbook.extractCellInfo = function (col, row) {
    var cellInfo, cellName;
    if (typeof col === "number" && typeof row === 'number') {
        return workbook.cellInfo(workbook.toColumn(col) + (row + 1).toString());
    } else if (typeof col === "string" && typeof row === 'number') {
        return workbook.cellInfo(col + row.toString());
    } else if (typeof col === "string" && typeof row === 'undefined') {
        return workbook.cellInfo(col);
    } else {
        throw Error("Expects either row, col or cell name (e.g. A1)");
    }
}

workbook.FormatNumber = {};

workbook.FormatNumber.format_definitions = {}; // Parsed formats are stored here globally


// Other constants

workbook.FormatNumber.commands =
    {
        copy: 1, color: 2, integer_placeholder: 3, fraction_placeholder: 4, decimal: 5,
        currency: 6, general: 7, separator: 8, date: 9, comparison: 10, section: 11, style: 12
    };



workbook.FormatNumber.formatNumberWithFormat = function (rawvalue, format_string, currency_char) {

    var scfn = workbook.FormatNumber;

    var op, operandstr, fromend, cval, operandstrlc;
    var startval, estartval;
    var hrs, mins, secs, ehrs, emins, esecs, ampmstr, ymd;
    var minOK, mpos;
    var result = "";
    var thisformat;
    var section, gotcomparison, compop, compval, cpos, oppos;
    var sectioninfo;
    var i, decimalscale, scaledvalue, strvalue, strparts, integervalue, fractionvalue;
    var integerdigits2, integerpos, fractionpos, textcolor, textstyle, separatorchar, decimalchar;
    var value; // working copy to change sign, etc.

    rawvalue = rawvalue - 0; // make sure a number
    value = rawvalue;
    if (!isFinite(value)) return "NaN";

    var negativevalue = value < 0 ? 1 : 0; // determine sign, etc.
    if (negativevalue) value = -value;
    var zerovalue = value == 0 ? 1 : 0;

    currency_char = currency_char || DefaultCurrency;

    scfn.parse_format_string(scfn.format_definitions, format_string); // make sure format is parsed
    thisformat = scfn.format_definitions[format_string]; // Get format structure

    if (!thisformat) throw "Format not parsed error!";

    section = thisformat.sectioninfo.length - 1; // get number of sections - 1

    if (thisformat.hascomparison) { // has comparisons - determine which section
        section = 0; // set to which section we will use
        gotcomparison = 0; // this section has no comparison
        for (cpos = 0; ; cpos++) { // scan for comparisons
            op = thisformat.operators[cpos];
            operandstr = thisformat.operands[cpos]; // get next operator and operand
            if (!op) { // at end with no match
                if (gotcomparison) { // if comparison but no match
                    format_string = "General"; // use default of General
                    scfn.parse_format_string(scfn.format_definitions, format_string);
                    thisformat = scfn.format_definitions[format_string];
                    section = 0;
                }
                break; // if no comparision, matches on this section
            }
            if (op == scfn.commands.section) { // end of section
                if (!gotcomparison) { // no comparison, so it's a match
                    break;
                }
                gotcomparison = 0;
                section++; // check out next one
                continue;
            }
            if (op == scfn.commands.comparison) { // found a comparison - do we meet it?
                i = operandstr.indexOf(":");
                compop = operandstr.substring(0, i);
                compval = operandstr.substring(i + 1) - 0;
                if ((compop == "<" && rawvalue < compval) ||
                    (compop == "<=" && rawvalue <= compval) ||
                    (compop == "=" && rawvalue == compval) ||
                    (compop == "<>" && rawvalue != compval) ||
                    (compop == ">=" && rawvalue >= compval) ||
                    (compop == ">" && rawvalue > compval)) { // a match
                    break;
                }
                gotcomparison = 1;
            }
        }
    }
    else if (section > 0) { // more than one section (separated by ";")
        if (section == 1) { // two sections
            if (negativevalue) {
                negativevalue = 0; // sign will provided by section, not automatically
                section = 1; // use second section for negative values
            }
            else {
                section = 0; // use first for all others
            }
        }
        else if (section == 2) { // three sections
            if (negativevalue) {
                negativevalue = 0; // sign will provided by section, not automatically
                section = 1; // use second section for negative values
            }
            else if (zerovalue) {
                section = 2; // use third section for zero values
            }
            else {
                section = 0; // use first for positive
            }
        }
    }

    sectioninfo = thisformat.sectioninfo[section]; // look at values for our section

    if (sectioninfo.commas > 0) { // scale by thousands
        for (i = 0; i < sectioninfo.commas; i++) {
            value /= 1000;
        }
    }
    if (sectioninfo.percent > 0) { // do percent scaling
        for (i = 0; i < sectioninfo.percent; i++) {
            value *= 100;
        }
    }

    decimalscale = 1; // cut down to required number of decimal digits
    for (i = 0; i < sectioninfo.fractiondigits; i++) {
        decimalscale *= 10;
    }
    scaledvalue = Math.floor(value * decimalscale + 0.5);
    scaledvalue = scaledvalue / decimalscale;

    if (typeof scaledvalue != "number") return "NaN";
    if (!isFinite(scaledvalue)) return "NaN";

    strvalue = scaledvalue + ""; // convert to string (Number.toFixed doesn't do all we need)

    //   strvalue = value.toFixed(sectioninfo.fractiondigits); // cut down to required number of decimal digits
    // and convert to string

    if (scaledvalue == 0 && (sectioninfo.fractiondigits || sectioninfo.integerdigits)) {
        negativevalue = 0; // no "-0" unless using multiple sections or General
    }

    if (strvalue.indexOf("e") >= 0) { // converted to scientific notation
        return rawvalue + ""; // Just return plain converted raw value
    }

    strparts = strvalue.match(/^\+{0,1}(\d*)(?:\.(\d*)){0,1}$/); // get integer and fraction parts
    if (!strparts) return "NaN"; // if not a number
    integervalue = strparts[1];
    if (!integervalue || integervalue == "0") integervalue = "";
    fractionvalue = strparts[2];
    if (!fractionvalue) fractionvalue = "";

    if (sectioninfo.hasdate) { // there are date placeholders
        if (rawvalue < 0) { // bad date
            return "??-???-??&nbsp;??:??:??";
        }
        startval = (rawvalue - Math.floor(rawvalue)) * SecondsInDay; // get date/time parts
        estartval = rawvalue * SecondsInDay; // do elapsed time version, too
        hrs = Math.floor(startval / SecondsInHour);
        ehrs = Math.floor(estartval / SecondsInHour);
        startval = startval - hrs * SecondsInHour;
        mins = Math.floor(startval / 60);
        emins = Math.floor(estartval / 60);
        secs = startval - mins * 60;
        decimalscale = 1; // round appropriately depending if there is ss.0
        for (i = 0; i < sectioninfo.fractiondigits; i++) {
            decimalscale *= 10;
        }
        secs = Math.floor(secs * decimalscale + 0.5);
        secs = secs / decimalscale;
        esecs = Math.floor(estartval * decimalscale + 0.5);
        esecs = esecs / decimalscale;
        if (secs >= 60) { // handle round up into next second, minute, etc.
            secs = 0;
            mins++; emins++;
            if (mins >= 60) {
                mins = 0;
                hrs++; ehrs++;
                if (hrs >= 24) {
                    hrs = 0;
                    rawvalue++;
                }
            }
        }
        fractionvalue = (secs - Math.floor(secs)) + ""; // for "hh:mm:ss.000"
        fractionvalue = fractionvalue.substring(2); // skip "0."

        ymd = workbook.FormatNumber.convert_date_julian_to_gregorian(Math.floor(rawvalue + JulianOffset));

        minOK = 0; // says "m" can be minutes if true
        mspos = sectioninfo.sectionstart; // m scan position in ops
        for (; ; mspos++) { // scan for "m" and "mm" to see if any minutes fields, and am/pm
            op = thisformat.operators[mspos];
            operandstr = thisformat.operands[mspos]; // get next operator and operand
            if (!op) break; // don't go past end
            if (op == scfn.commands.section) break;
            if (op == scfn.commands.date) {
                if ((operandstr.toLowerCase() == "am/pm" || operandstr.toLowerCase() == "a/p") && !ampmstr) {
                    if (hrs >= 12) {
                        hrs -= 12;
                        ampmstr = operandstr.toLowerCase() == "a/p" ? PM1 : PM; // "P" : "PM";
                    }
                    else {
                        ampmstr = operandstr.toLowerCase() == "a/p" ? AM1 : AM; // "A" : "AM";
                    }
                    if (operandstr.indexOf(ampmstr) < 0)
                        ampmstr = ampmstr.toLowerCase(); // have case match case in format
                }
                if (minOK && (operandstr == "m" || operandstr == "mm")) {
                    thisformat.operands[mspos] += "in"; // turn into "min" or "mmin"
                }
                if (operandstr.charAt(0) == "h") {
                    minOK = 1; // m following h or hh or [h] is minutes not months
                }
                else {
                    minOK = 0;
                }
            }
            else if (op != scfn.commands.copy) { // copying chars can be between h and m
                minOK = 0;
            }
        }
        minOK = 0;
        for (--mspos; ; mspos--) { // scan other way for s after m
            op = thisformat.operators[mspos];
            operandstr = thisformat.operands[mspos]; // get next operator and operand
            if (!op) break; // don't go past end
            if (op == scfn.commands.section) break;
            if (op == scfn.commands.date) {
                if (minOK && (operandstr == "m" || operandstr == "mm")) {
                    thisformat.operands[mspos] += "in"; // turn into "min" or "mmin"
                }
                if (operandstr == "ss") {
                    minOK = 1; // m before ss is minutes not months
                }
                else {
                    minOK = 0;
                }
            }
            else if (op != scfn.commands.copy) { // copying chars can be between ss and m
                minOK = 0;
            }
        }
    }

    integerdigits2 = 0; // init counters, etc.
    integerpos = 0;
    fractionpos = 0;
    textcolor = "";
    textstyle = "";
    separatorchar = SeparatorChar;
    if (separatorchar.indexOf(" ") >= 0) separatorchar = separatorchar.replace(/ /g, "&nbsp;");
    decimalchar = DecimalChar;
    if (decimalchar.indexOf(" ") >= 0) decimalchar = decimalchar.replace(/ /g, "&nbsp;");

    oppos = sectioninfo.sectionstart;

    while (op = thisformat.operators[oppos]) { // execute format
        operandstr = thisformat.operands[oppos++]; // get next operator and operand

        if (op == scfn.commands.copy) { // put char in result
            result += operandstr;
        }

        else if (op == scfn.commands.color) { // set color
            textcolor = operandstr;
        }

        else if (op == scfn.commands.style) { // set style
            textstyle = operandstr;
        }

        else if (op == scfn.commands.integer_placeholder) { // insert number part
            if (negativevalue) {
                result += "-";
                negativevalue = 0;
            }
            integerdigits2++;
            if (integerdigits2 == 1) { // first one
                if (integervalue.length > sectioninfo.integerdigits) { // see if integer wider than field
                    for (; integerpos < (integervalue.length - sectioninfo.integerdigits); integerpos++) {
                        result += integervalue.charAt(integerpos);
                        if (sectioninfo.thousandssep) { // see if this is a separator position
                            fromend = integervalue.length - integerpos - 1;
                            if (fromend > 2 && fromend % 3 == 0) {
                                result += separatorchar;
                            }
                        }
                    }
                }
            }
            if (integervalue.length < sectioninfo.integerdigits
                && integerdigits2 <= sectioninfo.integerdigits - integervalue.length) { // field is wider than value
                if (operandstr == "0" || operandstr == "?") { // fill with appropriate characters
                    result += operandstr == "0" ? "0" : "&nbsp;";
                    if (sectioninfo.thousandssep) { // see if this is a separator position
                        fromend = sectioninfo.integerdigits - integerdigits2;
                        if (fromend > 2 && fromend % 3 == 0) {
                            result += separatorchar;
                        }
                    }
                }
            }
            else { // normal integer digit - add it
                result += integervalue.charAt(integerpos);
                if (sectioninfo.thousandssep) { // see if this is a separator position
                    fromend = integervalue.length - integerpos - 1;
                    if (fromend > 2 && fromend % 3 == 0) {
                        result += separatorchar;
                    }
                }
                integerpos++;
            }
        }
        else if (op == scfn.commands.fraction_placeholder) { // add fraction part of number
            if (fractionpos >= fractionvalue.length) {
                if (operandstr == "0" || operandstr == "?") {
                    result += operandstr == "0" ? "0" : "&nbsp;";
                }
            }
            else {
                result += fractionvalue.charAt(fractionpos);
            }
            fractionpos++;
        }

        else if (op == scfn.commands.decimal) { // decimal point
            if (negativevalue) {
                result += "-";
                negativevalue = 0;
            }
            result += decimalchar;
        }

        else if (op == scfn.commands.currency) { // currency symbol
            if (negativevalue) {
                result += "-";
                negativevalue = 0;
            }
            result += operandstr;
        }

        else if (op == scfn.commands.general) { // insert "General" conversion

            // *** Cut down number of significant digits to avoid floating point artifacts:

            if (value != 0) { // only if non-zero
                var factor = Math.floor(Math.LOG10E * Math.log(value)); // get integer magnitude as a power of 10
                factor = Math.pow(10, 13 - factor); // turn into scaling factor
                value = Math.floor(factor * value + 0.5) / factor; // scale positive value, round, undo scaling
                if (!isFinite(value)) return "NaN";
            }
            if (negativevalue) {
                result += "-";
            }
            strvalue = value + ""; // convert original value to string
            if (strvalue.indexOf("e") >= 0) { // converted to scientific notation
                result += strvalue;
                continue;
            }
            strparts = strvalue.match(/^\+{0,1}(\d*)(?:\.(\d*)){0,1}$/); // get integer and fraction parts
            integervalue = strparts[1];
            if (!integervalue || integervalue == "0") integervalue = "";
            fractionvalue = strparts[2];
            if (!fractionvalue) fractionvalue = "";
            integerpos = 0;
            fractionpos = 0;
            if (integervalue.length) {
                for (; integerpos < integervalue.length; integerpos++) {
                    result += integervalue.charAt(integerpos);
                    if (sectioninfo.thousandssep) { // see if this is a separator position
                        fromend = integervalue.length - integerpos - 1;
                        if (fromend > 2 && fromend % 3 == 0) {
                            result += separatorchar;
                        }
                    }
                }
            }
            else {
                result += "0";
            }
            if (fractionvalue.length) {
                result += decimalchar;
                for (; fractionpos < fractionvalue.length; fractionpos++) {
                    result += fractionvalue.charAt(fractionpos);
                }
            }
        }
        else if (op == scfn.commands.date) { // date placeholder
            operandstrlc = operandstr.toLowerCase();
            if (operandstrlc == "y" || operandstrlc == "yy") {
                result += (ymd.year + "").substring(2);
            }
            else if (operandstrlc == "yyyy") {
                result += ymd.year + "";
            }
            else if (operandstrlc == "d") {
                result += ymd.day + "";
            }
            else if (operandstrlc == "dd") {
                cval = 1000 + ymd.day;
                result += (cval + "").substr(2);
            }
            else if (operandstrlc == "ddd") {
                cval = Math.floor(rawvalue + 6) % 7;
                result += DayNames3[cval];
            }
            else if (operandstrlc == "dddd") {
                cval = Math.floor(rawvalue + 6) % 7;
                result += DayNames[cval];
            }
            else if (operandstrlc == "m") {
                result += ymd.month + "";
            }
            else if (operandstrlc == "mm") {
                cval = 1000 + ymd.month;
                result += (cval + "").substr(2);
            }
            else if (operandstrlc == "mmm") {
                result += MonthNames3[ymd.month - 1];
            }
            else if (operandstrlc == "mmmm") {
                result += MonthNames[ymd.month - 1];
            }
            else if (operandstrlc == "mmmmm") {
                result += MonthNames[ymd.month - 1].charAt(0);
            }
            else if (operandstrlc == "h") {
                result += hrs + "";
            }
            else if (operandstrlc == "h]") {
                result += ehrs + "";
            }
            else if (operandstrlc == "mmin") {
                cval = (1000 + mins) + "";
                result += cval.substr(2);
            }
            else if (operandstrlc == "mm]") {
                if (emins < 100) {
                    cval = (1000 + emins) + "";
                    result += cval.substr(2);
                }
                else {
                    result += emins + "";
                }
            }
            else if (operandstrlc == "min") {
                result += mins + "";
            }
            else if (operandstrlc == "m]") {
                result += emins + "";
            }
            else if (operandstrlc == "hh") {
                cval = (1000 + hrs) + "";
                result += cval.substr(2);
            }
            else if (operandstrlc == "s") {
                cval = Math.floor(secs);
                result += cval + "";
            }
            else if (operandstrlc == "ss") {
                cval = (1000 + Math.floor(secs)) + "";
                result += cval.substr(2);
            }
            else if (operandstrlc == "am/pm" || operandstrlc == "a/p") {
                result += ampmstr;
            }
            else if (operandstrlc == "ss]") {
                if (esecs < 100) {
                    cval = (1000 + Math.floor(esecs)) + "";
                    result += cval.substr(2);
                }
                else {
                    cval = Math.floor(esecs);
                    result += cval + "";
                }
            }
        }
        else if (op == scfn.commands.section) { // end of section
            break;
        }

        else if (op == scfn.commands.comparison) { // ignore
            continue;
        }

        else {
            result += "!! Parse error !!";
        }
    }

    if (textcolor) {
        result = '<span style="color:' + textcolor + ';">' + result + '</span>';
    }
    if (textstyle) {
        result = '<span style="' + textstyle + ';">' + result + '</span>';
    }

    return result;

}

workbook.FormatNumber.parse_format_string = function (format_defs, format_string) {

    var scfn = workbook.FormatNumber;

    var thisformat, section, sectionfinfo;
    var integerpart = 1; // start out in integer part
    var lastwasinteger; // last char was an integer placeholder
    var lastwasslash; // last char was a backslash - escaping following character
    var lastwasasterisk; // repeat next char
    var lastwasunderscore; // last char was _ which picks up following char for width
    var inquote, quotestr; // processing a quoted string
    var inbracket, bracketstr, bracketdata; // processing a bracketed string
    var ingeneral, gpos; // checks for characters "General"
    var ampmstr, part; // checks for characters "A/P" and "AM/PM"
    var indate; // keeps track of date/time placeholders
    var chpos; // character position being looked at
    var ch; // character being looked at

    if (format_defs[format_string]) return; // already defined - nothing to do

    thisformat = { operators: [], operands: [], sectioninfo: [{}] }; // create info structure for this format
    format_defs[format_string] = thisformat; // add to other format definitions

    section = 0; // start with section 0
    sectioninfo = thisformat.sectioninfo[section]; // get reference to info for current section
    sectioninfo.sectionstart = 0; // position in operands that starts this section
    sectioninfo.integerdigits = 0; // number of integer-part placeholders
    sectioninfo.fractiondigits = 0; // fraction placeholders
    sectioninfo.commas = 0; // commas encountered, to handle scaling
    sectioninfo.percent = 0; // times to scale by 100

    for (chpos = 0; chpos < format_string.length; chpos++) { // parse
        ch = format_string.charAt(chpos); // get next char to examine
        if (inquote) {
            if (ch == '"') {
                inquote = 0;
                thisformat.operators.push(scfn.commands.copy);
                thisformat.operands.push(quotestr);
                continue;
            }
            quotestr += ch;
            continue;
        }
        if (inbracket) {
            if (ch == ']') {
                inbracket = 0;
                bracketdata = workbook.FormatNumber.parse_format_bracket(bracketstr);
                if (bracketdata.operator == scfn.commands.separator) {
                    sectioninfo.thousandssep = 1; // explicit [,]
                    continue;
                }
                if (bracketdata.operator == scfn.commands.date) {
                    sectioninfo.hasdate = 1;
                }
                if (bracketdata.operator == scfn.commands.comparison) {
                    thisformat.hascomparison = 1;
                }
                thisformat.operators.push(bracketdata.operator);
                thisformat.operands.push(bracketdata.operand);
                continue;
            }
            bracketstr += ch;
            continue;
        }
        if (lastwasslash) {
            thisformat.operators.push(scfn.commands.copy);
            thisformat.operands.push(ch);
            lastwasslash = false;
            continue;
        }
        if (lastwasasterisk) {
            thisformat.operators.push(scfn.commands.copy);
            thisformat.operands.push(ch + ch + ch + ch + ch); // do 5 of them since no real tabs
            lastwasasterisk = false;
            continue;
        }
        if (lastwasunderscore) {
            thisformat.operators.push(scfn.commands.copy);
            thisformat.operands.push("&nbsp;");
            lastwasunderscore = false;
            continue;
        }
        if (ingeneral) {
            if ("general".charAt(ingeneral) == ch.toLowerCase()) {
                ingeneral++;
                if (ingeneral == 7) {
                    thisformat.operators.push(scfn.commands.general);
                    thisformat.operands.push(ch);
                    ingeneral = 0;
                }
                continue;
            }
            ingeneral = 0;
        }
        if (indate) { // last char was part of a date placeholder
            if (indate.charAt(0) == ch) { // another of the same char
                indate += ch; // accumulate it
                continue;
            }
            thisformat.operators.push(scfn.commands.date); // something else, save date info
            thisformat.operands.push(indate);
            sectioninfo.hasdate = 1;
            indate = "";
        }
        if (ampmstr) {
            ampmstr += ch;
            part = ampmstr.toLowerCase();
            if (part != "am/pm".substring(0, part.length) && part != "a/p".substring(0, part.length)) {
                ampstr = "";
            }
            else if (part == "am/pm" || part == "a/p") {
                thisformat.operators.push(scfn.commands.date);
                thisformat.operands.push(ampmstr);
                ampmstr = "";
            }
            continue;
        }
        if (ch == "#" || ch == "0" || ch == "?") { // placeholder
            if (integerpart) {
                sectioninfo.integerdigits++;
                if (sectioninfo.commas) { // comma inside of integer placeholders
                    sectioninfo.thousandssep = 1; // any number is thousands separator
                    sectioninfo.commas = 0; // reset count of "thousand" factors
                }
                lastwasinteger = 1;
                thisformat.operators.push(scfn.commands.integer_placeholder);
                thisformat.operands.push(ch);
            }
            else {
                sectioninfo.fractiondigits++;
                thisformat.operators.push(scfn.commands.fraction_placeholder);
                thisformat.operands.push(ch);
            }
        }
        else if (ch == ".") { // decimal point
            lastwasinteger = 0;
            thisformat.operators.push(scfn.commands.decimal);
            thisformat.operands.push(ch);
            integerpart = 0;
        }
        else if (ch == '$') { // currency char
            lastwasinteger = 0;
            thisformat.operators.push(scfn.commands.currency);
            thisformat.operands.push(ch);
        }
        else if (ch == ",") {
            if (lastwasinteger) {
                sectioninfo.commas++;
            }
            else {
                thisformat.operators.push(scfn.commands.copy);
                thisformat.operands.push(ch);
            }
        }
        else if (ch == "%") {
            lastwasinteger = 0;
            sectioninfo.percent++;
            thisformat.operators.push(scfn.commands.copy);
            thisformat.operands.push(ch);
        }
        else if (ch == '"') {
            lastwasinteger = 0;
            inquote = 1;
            quotestr = "";
        }
        else if (ch == '[') {
            lastwasinteger = 0;
            inbracket = 1;
            bracketstr = "";
        }
        else if (ch == '\\') {
            lastwasslash = 1;
            lastwasinteger = 0;
        }
        else if (ch == '*') {
            lastwasasterisk = 1;
            lastwasinteger = 0;
        }
        else if (ch == '_') {
            lastwasunderscore = 1;
            lastwasinteger = 0;
        }
        else if (ch == ";") {
            section++; // start next section
            thisformat.sectioninfo[section] = {}; // create a new section
            sectioninfo = thisformat.sectioninfo[section]; // get reference to info for current section
            sectioninfo.sectionstart = 1 + thisformat.operators.length; // remember where it starts
            sectioninfo.integerdigits = 0; // number of integer-part placeholders
            sectioninfo.fractiondigits = 0; // fraction placeholders
            sectioninfo.commas = 0; // commas encountered, to handle scaling
            sectioninfo.percent = 0; // times to scale by 100
            integerpart = 1; // reset for new section
            lastwasinteger = 0;
            thisformat.operators.push(scfn.commands.section);
            thisformat.operands.push(ch);
        }
        else if (ch.toLowerCase() == "g") {
            ingeneral = 1;
            lastwasinteger = 0;
        }
        else if (ch.toLowerCase() == "a") {
            ampmstr = ch;
            lastwasinteger = 0;
        }
        else if ("dmyhHs".indexOf(ch) >= 0) {
            indate = ch;
        }
        else {
            lastwasinteger = 0;
            thisformat.operators.push(scfn.commands.copy);
            thisformat.operands.push(ch);
        }
    }

    if (indate) { // last char was part of unsaved date placeholder
        thisformat.operators.push(scfn.commands.date);
        thisformat.operands.push(indate);
        sectioninfo.hasdate = 1;
    }

    return;

}


workbook.FormatNumber.parse_format_bracket = function (bracketstr) {

    var scfn = workbook.FormatNumber;

    var bracketdata = {};
    var parts;

    if (bracketstr.charAt(0) == '$') { // currency
        bracketdata.operator = scfn.commands.currency;
        parts = bracketstr.match(/^\$(.+?)(\-.+?){0,1}$/);
        if (parts) {
            bracketdata.operand = parts[1] || DefaultCurrency || '$';
        }
        else {
            bracketdata.operand = bracketstr.substring(1) || DefaultCurrency || '$';
        }
    }
    else if (bracketstr == '?$') {
        bracketdata.operator = scfn.commands.currency;
        bracketdata.operand = '[?$]';
    }
    else if (AllowedColors[bracketstr.toUpperCase()]) {
        bracketdata.operator = scfn.commands.color;
        bracketdata.operand = AllowedColors[bracketstr.toUpperCase()];
    }
    else if (parts = bracketstr.match(/^style=([^"]*)$/)) { // [style=...]
        bracketdata.operator = scfn.commands.style;
        bracketdata.operand = parts[1];
    }
    else if (bracketstr == ",") {
        bracketdata.operator = scfn.commands.separator;
        bracketdata.operand = bracketstr;
    }
    else if (AllowedDates[bracketstr.toUpperCase()]) {
        bracketdata.operator = scfn.commands.date;
        bracketdata.operand = AllowedDates[bracketstr.toUpperCase()];
    }
    else if (parts = bracketstr.match(/^[<>=]/)) { // comparison operator
        parts = bracketstr.match(/^([<>=]+)(.+)$/); // split operator and value
        bracketdata.operator = scfn.commands.comparison;
        bracketdata.operand = parts[1] + ":" + parts[2];
    }
    else { // unknown bracket
        bracketdata.operator = scfn.commands.copy;
        bracketdata.operand = "[" + bracketstr + "]";
    }

    return bracketdata;

}

workbook.FormatNumber.convert_date_gregorian_to_julian = function (year, month, day) {

    var juliandate;

    juliandate = day - 32075 + workbook.intFunc(1461 * (year + 4800 + workbook.intFunc((month - 14) / 12)) / 4);
    juliandate += workbook.intFunc(367 * (month - 2 - workbook.intFunc((month - 14) / 12) * 12) / 12);
    juliandate = juliandate - workbook.intFunc(3 * workbook.intFunc((year + 4900 + workbook.intFunc((month - 14) / 12)) / 100) / 4);

    return juliandate;

}


workbook.FormatNumber.convert_date_julian_to_gregorian = function (juliandate) {

    var L, N, I, J, K;

    L = juliandate + 68569;
    N = Math.floor(4 * L / 146097);
    L = L - Math.floor((146097 * N + 3) / 4);
    I = Math.floor(4000 * (L + 1) / 1461001);
    L = L - Math.floor(1461 * I / 4) + 31;
    J = Math.floor(80 * L / 2447);
    K = L - Math.floor(2447 * J / 80);
    L = Math.floor(J / 11);
    J = J + 2 - 12 * L;
    I = 100 * (N - 49) + I + L;

    return { year: I, month: J, day: K };

}

workbook.makeFunction = function (name, args, exp) {
    if (typeof name === 'undefined') throw Error("Must provide name to make formula");
    if (typeof exp === 'undefined') { // handle overload for (name, exp)
        exp = args;
        args = "";
    }
    fn[name] = Function("return " + workbook.compile(exp, 2) + ";");
}

workbook.parse = function (context, exp) {
    if (typeof exp === 'undefined' && typeof context === 'string') {
        exp = context;
        context = {}
    } else if (typeof exp === 'undefined') {
        throw TypeError("No formula!");
    }

    if (typeof context !== 'object' && typeof context !== 'function') {
        throw TypeError("Context must be an object or a function!");
    }

    parser.yy.context = context;

    return parser.parse(exp);
}

workbook.resolveCell = function (ref) {
    if (fn.ISCELL(ref)) {
        return ref;
    } else if (typeof ref === 'function') {
        return ref();
    } else {
        throw Error("Expects cell or function");
    }
}

workbook.sheetName = function (name, cell) {
    return ((name.indexOf(' ') > 0 ?
        ("'" + name + "'!" + cell) :
        (name + "!" + cell)));

}

// A function to split Sheet!Cell or Sheet!Ref into 2 parts
workbook.splitReference = function (ref) {
    var parts = [], worksheet;

    if (typeof ref !== "string") {
        throw "expects string";
    }

    parts = ref.split('!');

    if (parts.length !== 2) {
        throw "reference should have two parts";
    }

    worksheet = parts[0];

    // strip leading and trailing ' if present
    if (worksheet.indexOf("'") === 0) {
        parts[0] = worksheet.substring(1, worksheet.length - 1);
    }

    return parts;

}

// ```

// toColumn
// --------

// Converts a column index (26) into the column (e.g. A or AB).

// ``` javascript
var toColumnMem = {};
workbook.toColumn = function (column_index) {

    if (toColumnMem.hasOwnProperty(column_index)) {
        return toColumnMem[column_index];
    }

    // The column is determined by applying a modified Hexavigesimal algorithm.
    // Normally BA follows Z but spreadsheets count wrong and nobody cares. 

    // Instead they do it in a way that makes sense to most people but
    // is mathmatically incorrect. So AA follows Z which in the base 10
    // system is like saying 01 follows 9. 

    // In the least significant digit
    // A..Z is 0..25

    // For the second to nth significant digit
    // A..Z is 1..26

    var converted = ""
        , secondPass = false
        , remainder;

    value = Math.abs(column_index);

    do {
        remainder = value % 26;

        if (secondPass) {
            remainder--;
        }

        converted = String.fromCharCode((remainder + 'A'.charCodeAt(0))) + converted;
        value = Math.floor((value - remainder) / 26);

        secondPass = true;
    } while (value > 0);

    toColumnMem[column_index] = converted;
    return converted;
}

// ```

// toColumnIndex
// -------------

// Convert column name back to the column index.

// ``` javascript
// Many thanks to http://en.wikipedia.org/wiki/Hexavigesimal for explaining the number theory
// and the basic structure of the algorithm (modified base26).
workbook.toColumnIndex = function (column) {

    // convert the column name into the column index

    // see toColumn for rant on why this is sensible even though it is illogical.
    var s = 0, secondPass;

    if (column != null && column.length > 0) {

        s = column.charCodeAt(0) - 'A'.charCodeAt(0);

        for (var i = 1; i < column.length; i++) {
            s += 1 // compensate for the weirdos that invented spreadsheet column naming
            s *= 26;
            s += column.charCodeAt(i) - 'A'.charCodeAt(0);
            secondPass = true;
        }

    } else {
        throw Error("Must provide a column");
    }

    return s;

}

// ```

// wrapString
// ----------

// ``` javascript
workbook.wrapString = function (s) {

    if (s[0] == "'" && s[s.length - 1] === "'") {
        return s;
    }

    return 'String(' + s + '.valueOf())';

}


workbook.intFunc = function (n) {
    if (n < 0) {
        return -Math.floor(-n);
    }
    else {
        return Math.floor(n);
    }
}

workbook.start = new Date();
workbook.beat = 0;

workbook.heartbeat = function () {
    workbook.beat++;
}

setInterval(workbook.heartbeat, 50);
workbook.hit = function (range, addr) {
    return range.hit(addr);
}
