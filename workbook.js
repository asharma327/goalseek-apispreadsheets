const fs = require('fs');
// --- Simple date helper: parse a string into a JS Date.
// --- If we fail, we just return null. A real solution might want to handle more formats robustly.
// function tryParseDate(dateStr) {
//     if (!dateStr) return null;
//     let d = new Date(dateStr);
//     if (isNaN(d.getTime())) return null;
//     return d;
// }

function tryParseDate(dateStr) {
    if (!dateStr) return null;

    // First, try the built-in Date parser.
    let d = new Date(dateStr);
    if (!isNaN(d.getTime())) {
        return d;
    }

    // Check for MM-DD-YYYY format.
    // e.g., "03-18-2025" should become "2025-03-18".
    let match = dateStr.match(/^(\d{2})-(\d{2})-(\d{4})$/);
    if (match) {
        // Rearrange to YYYY-MM-DD.
        let formatted = `${match[3]}-${match[1]}-${match[2]}`;
        let d2 = new Date(formatted);
        if (!isNaN(d2.getTime())) {
            return d2;
        }
    }

    return null;
}


// --- Excel-like helper functions ---
function excelIf(condition, trueVal, falseVal) {
    return condition ? trueVal : falseVal;
}

function excelAnd(...args) {
    return args.every((a) => !!a);
}

function excelOr(...args) {
    return args.some((a) => !!a);
}

// EDATE-like helper: shift a date by N months.
// Very naive logic: if day is out-of-range, we clamp to 28.
function excelEdate(startDate, months) {
    if (!(startDate instanceof Date) || isNaN(startDate.valueOf())) {
        return null;
    }
    let y = startDate.getFullYear();
    let m = startDate.getMonth() + 1;
    let d = startDate.getDate();

    let newMonth = m + months;
    let newYear = y;
    while (newMonth > 12) {
        newMonth -= 12;
        newYear += 1;
    }
    while (newMonth < 1) {
        newMonth += 12;
        newYear -= 1;
    }
    let safeDay = Math.min(d, 28);
    return new Date(newYear, newMonth - 1, safeDay);
}


// EOMONTH: shift a date by N months, then go to last day of that month
function excelEomonth(startDate, months) {
    let tmp = excelEdate(startDate, months);
    if (!tmp) return null;
    // "last day of that month":
    let newY = tmp.getFullYear();
    let newM = tmp.getMonth(); // 0-based
    // move to 1st day of next month, subtract 1 day
    let firstNext = new Date(newY, newM + 1, 1);
    let final = new Date(firstNext.getTime() - 86400000); // subtract 1 day in ms
    return final;
}

// Round down
function excelRounddown(number, digits) {
    let factor = Math.pow(10, digits);
    // floor can misbehave with negative numbers if you prefer "trunc" behavior, but we’ll keep it as floor
    return Math.floor(number * factor) / factor;
}

function excelMin(...args) {
    return Math.min(...args);
}

function excelMax(...args) {
    return Math.max(...args);
}

function excelYear(dateVal) {
    if (!(dateVal instanceof Date) || isNaN(dateVal.valueOf())) {
        return 0;
    }
    return dateVal.getFullYear();
}

// Excel-like rounding. Node's `Math.round` does "bankers rounding" for .5 in older versions?
// We'll do a naive approach that might differ slightly from Excel in edge .5 cases.
function excelRound(value, digits) {
    let f = Math.pow(10, digits);
    return Math.round(value * f) / f;
}

/**
 * The main big class to replicate the Python version in JavaScript.
 *
 * usage:
 *    let wb = new PdZppHJFxW4lidr7({ "15yrlump": { Xinput_datumvandaag: '2025-03-18', ... } });
 *    let results = wb.calculateOutputCells({ "15yrlump": ["F31", "H50", ...] }, preActions, postActions);
 *
 */
class PdZppHJFxW4lidr7 {
    constructor(inputCells) {
        // inputCells is expected to be like: { "15yrlump": { "Xinput_datumvandaag": "2025-03-18", ... } }
        this.inputCells = inputCells || {};
        this._staticValues = new Map(); //  store by key =>  (sheet, cell) => value

        // We define the “named_cells” structure similarly to Python
        this.namedCells = {
            "15yrlump": {
                rngTotal: "M64",
                Xinput_aanbiederkeuze: "F20",
                Xinput_datumvandaag: "F5",
                Xinput_geboortedatumaanvrager1: "F8",
                Xinput_geboortedatumaanvrager2: "F10",
                Xinput_hypotheekjanee: "F17",
                Xinput_hypotheeksaldo: "F18",
                Xinput_marktwaarde: "F15",
                Xinput_partnerjanee: "F9",
                Xinput_wozwaarde: "F16",
            },
        };

        // Fill in the "hard-coded" _staticValues that you had in your Python’s __init__:
        // I'm omitting many of them for brevity, but you can add them all similarly:
        this._fillStaticValues();

        // Apply user input to override some cells:
        this._updateStaticWithInputCells();
    }

    _fillStaticValues() {
        this._staticValues.set("15yrlump,E2", "Date update");
        this._staticValues.set("15yrlump,F2", new Date(2024, 10, 7)); // 2024-11-07 (JS month is 0-indexed)
        this._staticValues.set("15yrlump,E3", "Inputs");
        this._staticValues.set("15yrlump,F5", new Date(2024, 8, 5));
        this._staticValues.set("15yrlump,F8", new Date(1952, 1, 1));
        this._staticValues.set("15yrlump,E4", "Date/ time inputs");
        this._staticValues.set("15yrlump,E5", "Date today");
        this._staticValues.set("15yrlump,E6", "Date in 2 months");
        this._staticValues.set("15yrlump,E7", "Data applicants");
        this._staticValues.set("15yrlump,E8", "Birtdate applicant 1");
        this._staticValues.set("15yrlump,E9", "Do you have a partner? (ja/nee)");
        this._staticValues.set("15yrlump,E10", "Birtdate applicant 2");
        this._staticValues.set("15yrlump,E11", "Age applicant 1");
        this._staticValues.set("15yrlump,E12", "Age applicant 2");
        this._staticValues.set("15yrlump,E13", "Age for calculation");
        this._staticValues.set("15yrlump,E14", "Home/ mortgage data");
        this._staticValues.set("15yrlump,E15", "Market value");
        this._staticValues.set("15yrlump,F15", 550000);
        this._staticValues.set("15yrlump,E16", "WOZ-value (in general 85% of market value)");
        this._staticValues.set("15yrlump,F16", 220000);
        this._staticValues.set("15yrlump,E17", "Do you have a mortgage (ja/nee)");
        this._staticValues.set("15yrlump,F17", "Ja");
        this._staticValues.set("15yrlump,E18", "Input mortgage");
        this._staticValues.set("15yrlump,F18", 95000);
        this._staticValues.set("15yrlump,E19", "Mortgage for calculation");
        this._staticValues.set("15yrlump,E25", "Calculation - 15 year fixed");
        this._staticValues.set("15yrlump,E26", "Step 1");
        this._staticValues.set("15yrlump,E27", "Goalseek LTV/ EUR %");
        this._staticValues.set("15yrlump,E28", "LTV (goalseek)");
        this._staticValues.set("15yrlump,E29", "Does LTV > Exceed CAP?");
        this._staticValues.set("15yrlump,E30", "Step 2");
        this._staticValues.set("15yrlump,E31", "Minimum annual payments (year 2-14)");
        this._staticValues.set("15yrlump,E33", "Step 3 - result from Step 1 and 2");
        this._staticValues.set("15yrlump,E34", "First year (maximum) pay out%");
        this._staticValues.set("15yrlump,E35", "First year pay OUT");
        this._staticValues.set("15yrlump,E37", "Conclusion: first year pay-out is maximum of F36 or maximum first year payout based on goal seek.");
        this._staticValues.set("15yrlump,E39", "Walkthrough");
        this._staticValues.set("15yrlump,E40", "- Determine LTV (max)");
        this._staticValues.set("15yrlump,E41", "- Determine annual payments ( year 2-14)");
        this._staticValues.set("15yrlump,E42", "- goal seek initial pay-out");
        this._staticValues.set("15yrlump,C48", "15yrfixed");
        this._staticValues.set("15yrlump,D49", "Year");
        this._staticValues.set("15yrlump,E49", "Age (minimum)");
        this._staticValues.set("15yrlump,F49", "WOZ-value");
        this._staticValues.set("15yrlump,G49", "Current mortgage");
        this._staticValues.set("15yrlump,H49", "To be received");
        this._staticValues.set("15yrlump,I49", "Rate");
        this._staticValues.set("15yrlump,J49", "Annual interest");
        this._staticValues.set("15yrlump,K49", "Cum interest");
        this._staticValues.set("15yrlump,L49", "Totale mortgage");
        this._staticValues.set("15yrlump,M49", "Total");
        this._staticValues.set("15yrlump,N49", "LTV ( for review only)");
        this._staticValues.set("15yrlump,C50", 1);
        this._staticValues.set("15yrlump,H50", -254.39000000000524);
        this._staticValues.set("15yrlump,C51", 2);
        this._staticValues.set("15yrlump,H51", 2100);
        this._staticValues.set("15yrlump,C52", 3);
        this._staticValues.set("15yrlump,H52", 2100);
        this._staticValues.set("15yrlump,C53", 4);
        this._staticValues.set("15yrlump,H53", 2100);
        this._staticValues.set("15yrlump,C54", 5);
        this._staticValues.set("15yrlump,H54", 2100);
        this._staticValues.set("15yrlump,C55", 6);
        this._staticValues.set("15yrlump,H55", 2100);
        this._staticValues.set("15yrlump,C56", 7);
        this._staticValues.set("15yrlump,H56", 2100);
        this._staticValues.set("15yrlump,C57", 8);
        this._staticValues.set("15yrlump,H57", 2100);
        this._staticValues.set("15yrlump,C58", 9);
        this._staticValues.set("15yrlump,H58", 2100);
        this._staticValues.set("15yrlump,C59", 10);
        this._staticValues.set("15yrlump,H59", 2100);
        this._staticValues.set("15yrlump,C60", 11);
        this._staticValues.set("15yrlump,H60", 2100);
        this._staticValues.set("15yrlump,C61", 12);
        this._staticValues.set("15yrlump,H61", 2100);
        this._staticValues.set("15yrlump,C62", 13);
        this._staticValues.set("15yrlump,H62", 2100);
        this._staticValues.set("15yrlump,C63", 14);
        this._staticValues.set("15yrlump,H63", 2100);
        this._staticValues.set("15yrlump,C64", 15);
        this._staticValues.set("15yrlump,H64", 2100);
        this._staticValues.set("15yrlump,C65", 16);
        this._staticValues.set("15yrlump,H65", 0);
        this._staticValues.set("15yrlump,C66", 17);
        this._staticValues.set("15yrlump,H66", 0);
        this._staticValues.set("15yrlump,C67", 18);
        this._staticValues.set("15yrlump,H67", 0);
        this._staticValues.set("15yrlump,C68", 19);
        this._staticValues.set("15yrlump,H68", 0);
        this._staticValues.set("15yrlump,C69", 20);
        this._staticValues.set("15yrlump,H69", 0);
        this._staticValues.set("15yrlump,C70", 21);
        this._staticValues.set("15yrlump,H70", 0);
        this._staticValues.set("15yrlump,C71", 22);
        this._staticValues.set("15yrlump,H71", 0);
        this._staticValues.set("15yrlump,C72", 23);
        this._staticValues.set("15yrlump,H72", 0);
        this._staticValues.set("15yrlump,C73", 24);
        this._staticValues.set("15yrlump,H73", 0);
        this._staticValues.set("15yrlump,C74", 25);
        this._staticValues.set("15yrlump,H74", 0);
        this._staticValues.set("15yrlump,C75", 26);
        this._staticValues.set("15yrlump,H75", 0);
        this._staticValues.set("15yrlump,C76", 27);
        this._staticValues.set("15yrlump,H76", 0);
        this._staticValues.set("15yrlump,C77", 28);
        this._staticValues.set("15yrlump,H77", 0);
        this._staticValues.set("15yrlump,C78", 29);
        this._staticValues.set("15yrlump,H78", 0);
        this._staticValues.set("15yrlump,C79", 30);
        this._staticValues.set("15yrlump,H79", 0);
        this._staticValues.set("15yrlump,E83", "Conclusion");
        this._staticValues.set("15yrlump,E84", "Total payout");
        this._staticValues.set("15yrlump,E85", "Interest");
        this._staticValues.set("15yrlump,I85", "Is first year payout exceed minimum>");
        this._staticValues.set("15yrlump,E86", "Check on Yes");
        this._staticValues.set("15yrlump,I86", "Is total payout > 25.000");
        this._staticValues.set("15yrlump,I87", "Is first year pay>");
        this._staticValues.set("param15yrlump,D4", "Parameters");
        this._staticValues.set("param15yrlump,D5", "Interest");
        this._staticValues.set("param15yrlump,E5", 0.0659);
        this._staticValues.set("param15yrlump,D6", "Digits");
        this._staticValues.set("param15yrlump,E6", 2);
        this._staticValues.set("param15yrlump,D7", "Additional first year pay out");
        this._staticValues.set("param15yrlump,E7", 5000);
        this._staticValues.set("param15yrlump,D8", "CAP on maximum payout (cumulative)");
        this._staticValues.set("param15yrlump,E8", 550000);
        this._staticValues.set("param15yrlump,D9", "Minimal mortgage");
        this._staticValues.set("param15yrlump,E9", 25000);
        this._staticValues.set("param15yrlump,D11", "LTV op WOZ-value");
        this._staticValues.set("param15yrlump,E11", "MAX LTV ON (START) AGE");
        this._staticValues.set("param15yrlump,F11", "Initial Pay out");
        this._staticValues.set("param15yrlump,D13", 55);
        this._staticValues.set("param15yrlump,E13", 0.0);
        this._staticValues.set("param15yrlump,F13", 0.0);
        this._staticValues.set("param15yrlump,D14", 56);
        this._staticValues.set("param15yrlump,E14", 0.0);
        this._staticValues.set("param15yrlump,F14", 0.0);
        this._staticValues.set("param15yrlump,D15", 57);
        this._staticValues.set("param15yrlump,E15", 0.0);
        this._staticValues.set("param15yrlump,F15", 0.0);
        this._staticValues.set("param15yrlump,D22", 64);
        this._staticValues.set("param15yrlump,E22", 0.55);
        this._staticValues.set("param15yrlump,F22", 0.1);
        this._staticValues.set("param15yrlump,D23", 65);
        this._staticValues.set("param15yrlump,E23", 0.55);
        this._staticValues.set("param15yrlump,F23", 0.1);
        this._staticValues.set("param15yrlump,D24", 66);
        this._staticValues.set("param15yrlump,E24", 0.57);
        this._staticValues.set("param15yrlump,F24", 0.106);
        this._staticValues.set('param15yrlump,D25', 67)
        this._staticValues.set('param15yrlump,E25', 0.59)
        this._staticValues.set('param15yrlump,F25', 0.112)
        this._staticValues.set('param15yrlump,D26', 68)
        this._staticValues.set('param15yrlump,E26', 0.61)
        this._staticValues.set('param15yrlump,F26', 0.118)
        this._staticValues.set('param15yrlump,D27', 69)
        this._staticValues.set('param15yrlump,E27', 0.63)
        this._staticValues.set('param15yrlump,F27', 0.124)
        this._staticValues.set('param15yrlump,D28', 70)
        this._staticValues.set('param15yrlump,E28', 0.65)
        this._staticValues.set('param15yrlump,F28', 0.128)
        this._staticValues.set('param15yrlump,D29', 71)
        this._staticValues.set('param15yrlump,E29', 0.66)
        this._staticValues.set('param15yrlump,F29', 0.132)
        this._staticValues.set('param15yrlump,D30', 72)
        this._staticValues.set('param15yrlump,E30', 0.665)
        this._staticValues.set('param15yrlump,F30', 0.136)
        this._staticValues.set('param15yrlump,D31', 73)
        this._staticValues.set('param15yrlump,E31', 0.67)
        this._staticValues.set('param15yrlump,F31', 0.14)
        this._staticValues.set('param15yrlump,D32', 74)
        this._staticValues.set('param15yrlump,E32', 0.675)
        this._staticValues.set('param15yrlump,F32', 0.144)
        this._staticValues.set('param15yrlump,D33', 75)
        this._staticValues.set('param15yrlump,E33', 0.68)
        this._staticValues.set('param15yrlump,F33', 0.148)
        this._staticValues.set('param15yrlump,D34', 76)
        this._staticValues.set('param15yrlump,E34', 0.685)
        this._staticValues.set('param15yrlump,F34', 0.15)
        this._staticValues.set('param15yrlump,D35', 77)
        this._staticValues.set('param15yrlump,E35', 0.69)
        this._staticValues.set('param15yrlump,F35', 0.152)
        this._staticValues.set('param15yrlump,D36', 78)
        this._staticValues.set('param15yrlump,E36', 0.695)
        this._staticValues.set('param15yrlump,F36', 0.154)
        this._staticValues.set('param15yrlump,D37', 79)
        this._staticValues.set('param15yrlump,E37', 0.7)
        this._staticValues.set('param15yrlump,F37', 0.158)
        this._staticValues.set('param15yrlump,D38', 80)
        this._staticValues.set('param15yrlump,E38', 0.705)
        this._staticValues.set('param15yrlump,F38', 0.16)
        this._staticValues.set('param15yrlump,D39', 81)
        this._staticValues.set('param15yrlump,E39', 0.71)
        this._staticValues.set('param15yrlump,F39', 0.162)
        this._staticValues.set('param15yrlump,D40', 82)
        this._staticValues.set('param15yrlump,E40', 0.715)
        this._staticValues.set('param15yrlump,F40', 0.164)
        this._staticValues.set('param15yrlump,D41', 83)
        this._staticValues.set('param15yrlump,E41', 0.72)
        this._staticValues.set('param15yrlump,F41', 0.166)
        this._staticValues.set('param15yrlump,D42', 84)
        this._staticValues.set('param15yrlump,E42', 0.725)
        this._staticValues.set('param15yrlump,F42', 0.168)
        this._staticValues.set('param15yrlump,D43', 85)
        this._staticValues.set('param15yrlump,E43', 0.725)
        this._staticValues.set('param15yrlump,D44', 86)
        this._staticValues.set('param15yrlump,E44', 0.725)
        this._staticValues.set('param15yrlump,D45', 87)
        this._staticValues.set('param15yrlump,E45', 0.725)
        this._staticValues.set('param15yrlump,D46', 88)
        this._staticValues.set('param15yrlump,E46', 0.725)
        this._staticValues.set('param15yrlump,D47', 89)
        this._staticValues.set('param15yrlump,E47', 0.725)
        this._staticValues.set('param15yrlump,D48', 90)
        this._staticValues.set('param15yrlump,E48', 0.725)
        this._staticValues.set('param15yrlump,D49', 91)
        this._staticValues.set('param15yrlump,E49', 0.725)
        this._staticValues.set('param15yrlump,D50', 92)
        this._staticValues.set('param15yrlump,E50', 0.725)
        this._staticValues.set('param15yrlump,D51', 93)
        this._staticValues.set('param15yrlump,E51', 0.725)
        this._staticValues.set('param15yrlump,D52', 94)
        this._staticValues.set('param15yrlump,E52', 0.725)
        this._staticValues.set('param15yrlump,D53', 95)
        this._staticValues.set('param15yrlump,E53', 0.725)
        this._staticValues.set('param15yrlump,D54', 96)
        this._staticValues.set('param15yrlump,E54', 0.725)
        this._staticValues.set('param15yrlump,D55', 97)
        this._staticValues.set('param15yrlump,E55', 0.725)
        this._staticValues.set('param15yrlump,D56', 98)
        this._staticValues.set('param15yrlump,E56', 0.725)
        this._staticValues.set('param15yrlump,D57', 99)
        this._staticValues.set('param15yrlump,E57', 0.725)
        this._staticValues.set('param15yrlump,D58', 100)
        this._staticValues.set('param15yrlump,E58', 0.725)
        this._staticValues.set('param15yrlump,D59', 101)
        this._staticValues.set('param15yrlump,E59', 0.725)
        this._staticValues.set('param15yrlump,D60', 102)
        this._staticValues.set('param15yrlump,E60', 0.725)
        this._staticValues.set('param15yrlump,D61', 103)
        this._staticValues.set('param15yrlump,E61', 0.725)
        this._staticValues.set('param15yrlump,D62', 104)
        this._staticValues.set('param15yrlump,E62', 0.725)
        this._staticValues.set('param15yrlump,D63', 105)
        this._staticValues.set('param15yrlump,E63', 0.725)
        this._staticValues.set('param15yrlump,D64', 106)
        this._staticValues.set('param15yrlump,E64', 0.725)
        this._staticValues.set('param15yrlump,D65', 107)
        this._staticValues.set('param15yrlump,E65', 0.725)
        this._staticValues.set('param15yrlump,D66', 108)
        this._staticValues.set('param15yrlump,E66', 0.725)
        this._staticValues.set('param15yrlump,D67', 109)
        this._staticValues.set('param15yrlump,E67', 0.725)
        this._staticValues.set('param15yrlump,D68', 110)
        this._staticValues.set('param15yrlump,E68', 0.725)
        // etc. up to the age 110 row.
        // For brevity, you can fill them exactly as in your Python code or omit the ones that are 0.725 etc.

        // Done with the missing lines
    }


    // Mimic the Python method that merges user input into the _staticValues
    _updateStaticWithInputCells() {
        for (let sheetName in this.inputCells) {
            if (!this.namedCells[sheetName]) continue;
            let sheetInputs = this.inputCells[sheetName];
            for (let cellRef in sheetInputs) {
                let actualCell = this.namedCells[sheetName][cellRef];
                if (!actualCell) continue;

                let rawVal = sheetInputs[cellRef];

                // If rawVal is a string, trim it.
                if (typeof rawVal === "string") {
                    rawVal = rawVal.trim();

                    // Check for YYYY-MM-DD format
                    if (/^\d{4}-\d{2}-\d{2}$/.test(rawVal)) {
                        let d = new Date(rawVal);
                        if (!isNaN(d.getTime())) {
                            this.setValue(sheetName, actualCell, d);
                            continue; // Date successfully parsed, skip further processing
                        }
                    }

                    // Check for MM-DD-YYYY format
                    if (/^\d{2}-\d{2}-\d{4}$/.test(rawVal)) {
                        let parts = rawVal.split("-");
                        let month = parseInt(parts[0], 10);
                        let day = parseInt(parts[1], 10);
                        let year = parseInt(parts[2], 10);
                        let d = new Date(year, month - 1, day);
                        if (!isNaN(d.getTime())) {
                            this.setValue(sheetName, actualCell, d);
                            continue; // Date successfully parsed, skip further processing
                        }
                    }
                }

                // Remove commas if present.
                if (typeof rawVal === "string") {
                    rawVal = rawVal.replace(/,/g, "");
                }

                // Now attempt to parse as a float.
                let valNum = parseFloat(rawVal);
                if (!isNaN(valNum)) {
                    this.setValue(sheetName, actualCell, valNum);
                } else {
                    this.setValue(sheetName, actualCell, String(rawVal));
                }
            }
        }
    }


    // The equivalent of `set_value(sheet, cell, new_val)`
    setValue(sheetName, cellName, newValue) {
        // Use string keys consistently
        this._staticValues.set(`${sheetName},${cellName}`, newValue);
    }

    _vlookupExact(lookupValue, sheetName, firstCol, lastCol, startRow, endRow, returnColIndex) {
        // Build an array of column letters from firstCol to lastCol.
        let cols = [];
        let startCode = firstCol.charCodeAt(0);
        let endCode = lastCol.charCodeAt(0);
        for (let c = startCode; c <= endCode; c++) {
            cols.push(String.fromCharCode(c));
        }

        // Determine the lookup column (first column) and the return column.
        let lookupColumn = cols[0];
        let retCol = cols[returnColIndex - 1];
        if (!retCol) return 0; // Return 0 if the return column is invalid.

        // Define a small tolerance to allow for floating point imprecision.
        const tolerance = 1e-10;

        // Iterate through the specified rows.
        for (let row = startRow; row <= endRow; row++) {
            let cellAddress = `${lookupColumn}${row}`;
            let lookupCellVal = this.getValue(sheetName, cellAddress);

            // Skip this row if the lookup cell is null/undefined.
            if (lookupCellVal == null) continue;

            // Try to convert the lookup cell to a number.
            let parsedVal = parseFloat(lookupCellVal);
            if (isNaN(parsedVal)) continue;

            // Compare using tolerance.
            if (Math.abs(parsedVal - lookupValue) < tolerance) {
                // If a match is found, get the return cell's value.
                let returnVal = this.getValue(sheetName, `${retCol}${row}`);
                // Attempt to parse the return value to a number; if not possible, return it as-is.
                let parsedReturn = parseFloat(returnVal);
                return isNaN(parsedReturn) ? returnVal : parsedReturn;
            }
        }
        // If no match is found, return 0.
        return 0;
    }


    _excelEdate(startDate, months) {
        if (!(startDate instanceof Date) || isNaN(startDate.valueOf())) {
            return null;
        }
        // naive approach:
        let year = startDate.getFullYear();
        let month = startDate.getMonth() + 1; // +1 because JS months are 0-based
        let day = startDate.getDate();

        let newMonth = month + months;
        let newYear = year;
        while (newMonth > 12) {
            newMonth -= 12;
            newYear += 1;
        }
        while (newMonth < 1) {
            newMonth += 12;
            newYear -= 1;
        }
        // clamp day to 28 if it is out of range
        let safeDay = Math.min(day, 28);
        return new Date(newYear, newMonth - 1, safeDay);
    }

    /**
     * A Node version of the Python "_get_input_from_named" method.
     * We check if "inputCells[sheet][named_range]" is set; if so, return it.
     * Otherwise, fallback to "this._staticValues.get(fallback_key)".
     */
    _getInputFromNamed(sheet, namedRange, fallbackKey) {
        // fallbackKey is something like ("15yrlump","F5"). We'll handle it as [sheetName, cellName].
        let fallbackSheet = fallbackKey[0];
        let fallbackCell  = fallbackKey[1];

        // let sheetInputs = this.inputCells[sheet] || {};
        // let val = sheetInputs[namedRange];
        // if (val !== undefined && val !== null) {
        //     return val;
        // }

        // fallback => read from staticValues
        return this.getValue(fallbackSheet, fallbackCell);
    }

    /**
     * rngTotal:
     * Named range for '15yrlump'!M64
     */
    rngTotal() {
        // In Python, you had a commented-out line referencing _get_input_from_named,
        // but ultimately just returned get_value("15yrlump","M64").
        return this.getValue("15yrlump", "M64");
    }

    /**
     * Xinput_aanbiederkeuze:
     * Named range for '15yrlump'!F20
     */
    Xinput_aanbiederkeuze() {
        return this._getInputFromNamed('15yrlump', 'Xinput_aanbiederkeuze', ['15yrlump','F20']);
    }

    /**
     * Xinput_datumvandaag:
     * Named range for '15yrlump'!F5
     */
    Xinput_datumvandaag() {
        return this._getInputFromNamed('15yrlump','Xinput_datumvandaag', ['15yrlump','F5']);
    }

    /**
     * Xinput_geboortedatumaanvrager1:
     * Named range for '15yrlump'!F8
     */
    Xinput_geboortedatumaanvrager1() {
        return this._getInputFromNamed('15yrlump','Xinput_geboortedatumaanvrager1', ['15yrlump','F8']);
    }

    /**
     * Xinput_geboortedatumaanvrager2:
     * Named range for '15yrlump'!F10
     * We attempt to parse it as a Date if possible.
     */
    Xinput_geboortedatumaanvrager2() {
        let val = this._getInputFromNamed('15yrlump','Xinput_geboortedatumaanvrager2', ['15yrlump','F10']);
        if (!val) return val;

        // Attempt to parse as date
        try {
            // e.g. if you have a parse function or do new Date(...)
            let d = new Date(val);
            if (!isNaN(d.getTime())) {
                return d;
            }
        } catch(e) {
            // if parse fails, just fall through
        }
        return val;
    }

    /**
     * Xinput_hypotheekjanee:
     * Named range for '15yrlump'!F17
     */
    Xinput_hypotheekjanee() {
        return this._getInputFromNamed('15yrlump','Xinput_hypotheekjanee', ['15yrlump','F17']);
    }

    /**
     * Xinput_hypotheeksaldo:
     * Named range for '15yrlump'!F18
     * We try to parse it as float. If that fails, just return the raw value.
     */
    Xinput_hypotheeksaldo() {
        let val = this._getInputFromNamed('15yrlump','Xinput_hypotheeksaldo', ['15yrlump','F18']);
        if (val === null || val === undefined) return val;

        let f = parseFloat(val);
        if (!isNaN(f)) {
            return f;
        }
        return val; // if parseFloat fails, fallback
    }

    /**
     * Xinput_marktwaarde:
     * Named range for '15yrlump'!F15
     */
    Xinput_marktwaarde() {
        return this._getInputFromNamed('15yrlump','Xinput_marktwaarde', ['15yrlump','F15']);
    }

    /**
     * Xinput_partnerjanee:
     * Named range for '15yrlump'!F9
     */
    Xinput_partnerjanee() {
        return this._getInputFromNamed('15yrlump','Xinput_partnerjanee', ['15yrlump','F9']);
    }

    /**
     * Xinput_wozwaarde:
     * Named range for '15yrlump'!F16
     * We parse the result as a float.
     */
    Xinput_wozwaarde() {
        let val = this._getInputFromNamed('15yrlump','Xinput_wozwaarde', ['15yrlump','F16']);
        if (val === null || val === undefined) return 0;
        let f = parseFloat(val);
        return isNaN(f) ? 0 : f;
    }


    getDValueForRow(rowNum) {
        if (rowNum === 50){
            let currDt = this._staticValues.get("15yrlump,F5")
            return currDt.getFullYear()
        } else {
            let prevD = parseFloat(this.getValue("15yrlump", `D${rowNum - 1}`) || 0);
            return prevD + 1;
        }
    }

    /**
     *  read E(row-1), add 1
     */
    getEValueForRow(rowNum) {
        if (rowNum === 50){
            return this.getValue("15yrlump", "F13")
        } else {
            let prevE = parseFloat(this.getValue("15yrlump", `E${rowNum - 1}`) || 0);
            return prevE + 1;
        }

    }

    /**
     *  always = $F$16
     *  so we can just do return this.getValue("15yrlump","F16")
     */
    getFValueForRow(rowNum) {
        // ignoring rowNum because it’s the same for all rows
        return this.getValue("15yrlump", "F16");
    }

    /**
     *  always = $F$19
     */
    getGValueForRow(rowNum) {
        return this.getValue("15yrlump", "F19");
    }

    /**
     *  always = param15yrlump!E5
     */
    getIValueForRow(rowNum) {
        return this.getValue("param15yrlump", "E5");
    }

    /**
     *  J(row) = ROUND((H(row) + L(row-1)) * (1 + I(row)/(1 - I(row)))^(1/12)^12 - (H(row) + L(row-1)), E6)
     */
    getJValueForRow(rowNum) {
        if (rowNum === 50){
            let h50 = parseFloat(this.getValue("15yrlump", `H${rowNum}`) || 0);
            let i50 = parseFloat(this.getValue("15yrlump", `I${rowNum}`) || 0);
            let digits = parseInt(this.getValue("param15yrlump", "E6") || 0, 10);
            let base;
            if (Math.abs(1 - i50) > 1e-12) {
                // avoid division by zero
                base = 1 + i50 / (1 - i50);
            }  else {
                base = 1e9
            }
            let monthly = Math.pow(base, 1 / 12);
            let finalFactor = Math.pow(monthly, 12);
            let expr = (h50 * finalFactor) - h50;
            let val = this._excelRound(expr, digits);
            return val
        } else {
            let hVal = parseFloat(this.getValue("15yrlump", `H${rowNum}`) || 0);
            let lPrev = parseFloat(this.getValue("15yrlump", `L${rowNum - 1}`) || 0);
            let iVal = parseFloat(this.getValue("15yrlump", `I${rowNum}`) || 0);
            let digits = parseInt(this.getValue("param15yrlump", "E6") || 0, 10);
            let base;
            if (Math.abs(1 - iVal) < 1e-12) {
                // avoid division by zero
                base = 1e9
            }  else {
                base = 1 + iVal / (1 - iVal);
            }
            let monthly = Math.pow(base, 1 / 12);
            let finalFactor = Math.pow(monthly, 12);
            let expr = (hVal + lPrev) * finalFactor - (hVal + lPrev);
            let val = this._excelRound(expr, digits);
            return val
        }

    }

    /**
     *  K(row) = J(row) + K(row-1)
     */
    getKValueForRow(rowNum) {
        let jVal = parseFloat(this.getValue("15yrlump", `J${rowNum}`) || 0);
        // For row 50, ignore previous row (header) value and assume it’s zero.
        if (rowNum === 50) {
            return jVal;
        }
        let kPrev = parseFloat(this.getValue("15yrlump", `K${rowNum - 1}`) || 0);
        return jVal + kPrev;
    }


    /**
     *  L(row) = H(row) + J(row) + L(row-1)
     */
    getLValueForRow(rowNum) {
        if (rowNum === 50) {
            let hVal = parseFloat(this.getValue("15yrlump", `H${rowNum}`) || 0);
            let kVal = parseFloat(this.getValue("15yrlump", `K${rowNum}`) || 0);

            return hVal + kVal;
        } else {
            let hVal = parseFloat(this.getValue("15yrlump", `H${rowNum}`) || 0);
            let jVal = parseFloat(this.getValue("15yrlump", `J${rowNum}`) || 0);
            let lPrev = parseFloat(this.getValue("15yrlump", `L${rowNum - 1}`) || 0);

            return hVal + jVal + lPrev;
        }
    }

    /**
     *  M(row) = L(row) + G(row)
     */
    getMValueForRow(rowNum) {
        let lVal = parseFloat(this.getValue("15yrlump", `L${rowNum}`) || 0);
        let gVal = parseFloat(this.getValue("15yrlump", `G${rowNum}`) || 0);
        return lVal + gVal;
    }

    /**
     *  N(row) = M(row) / F(row)
     */
    getNValueForRow(rowNum) {
        let mVal = parseFloat(this.getValue("15yrlump", `M${rowNum}`) || 0);
        let fVal = parseFloat(this.getValue("15yrlump", `F${rowNum}`) || 1);
        if (Math.abs(fVal) < 1e-12) return 0;
        return mVal / fVal;
    }


    // The equivalent of `get_value(sheetName, cellName)`
    // which in Python sometimes calls a formula method.
    // For brevity, we’ll just do direct lookups plus certain known formula references below.
    getValue(sheetName, cellName) {
        let key = `${sheetName},${cellName}`;
        let formulaMap = this._formulaMethodMap();
        let formulaMethod = formulaMap[key];
        if (formulaMethod && typeof this[formulaMethod] === "function") {
            return this[formulaMethod]();
        }
        return this._staticValues.get(key);
    }

    // ***EXAMPLE***: replicate a few formula cells from your Python code:
    sheet_15yrlump_F6() {
        // =EDATE(F5, 2)
        let f5Val = this.getValue("15yrlump", "F5");
        return excelEdate(f5Val, 2);
    }

    sheet_15yrlump_F11() {
        // =ROUNDDOWN((F6 - F8)/365.25, 0)
        let d6 = this.getValue("15yrlump", "F6");
        let d8 = this.getValue("15yrlump", "F8");
        if (!d6 || !d8 || !(d6 instanceof Date) || !(d8 instanceof Date)) return 0;
        let diffMs = d6.getTime() - d8.getTime();
        let diffDays = diffMs / 86400000;
        let years = diffDays / 365.25;
        return Math.floor(years);
    }

    sheet_15yrlump_F12() {
        // check the partner cell => F9
        let partnerVal = this.getValue("15yrlump", "F9"); // "Nee" or "Ja"
        if (typeof partnerVal === "string" && partnerVal.toLowerCase() === "nee") {
            return "";  // empty string
        }
        // otherwise compute
        let d6 = this.getValue("15yrlump", "F6");
        let d10 = this.getValue("15yrlump", "F10");
        if (!(d6 instanceof Date) || !(d10 instanceof Date)) {
            return "";
        }
        let diffDays = (d6 - d10) / (1000*3600*24);
        let years = diffDays / 365.25;
        return Math.floor(years);
    }

// =IF(F9="Nee", F11, IF(F9="Ja", MIN(F11:F12)))
    sheet_15yrlump_F13() {
        let f9 = this.getValue("15yrlump", "F9");  // "Nee" or "Ja"
        let valF11 = this.getValue("15yrlump", "F11");
        let valF12 = this.getValue("15yrlump", "F12");

        if (typeof f9 === "string" && f9.toLowerCase() === "nee") {
            return valF11;
        }
        if (typeof f9 === "string" && f9.toLowerCase() === "ja") {
            // If F12 is "", treat that as ignoring it or 9999? We’ll do a simple check:
            if (valF12 === "") {
                return valF11;
            }
            return Math.min(parseFloat(valF11||0), parseFloat(valF12||0));
        }
        return null;
    }

// =IF(F17="Nee", 0, Xinput_hypotheeksaldo)
    sheet_15yrlump_F19() {
        let f17 = this.getValue("15yrlump", "F17"); // "Nee" or "Ja"
        if (typeof f17 === "string" && f17.toLowerCase() === "nee") {
            return 0;
        }
        // else => return F18
        let hypoSaldo = parseFloat(this.getValue("15yrlump", "F18") || 0);
        return hypoSaldo;
    }

// =VLOOKUP(F13, param15yrlump!D13:F68, 2, FALSE)
    sheet_15yrlump_F27() {
        let valF13 = parseFloat(this.getValue("15yrlump", "F13") || 0);
        // we do an exact VLOOKUP on param15yrlump D13..F68 => returning col2 => 'E'
        let vlookupMatch = this._vlookupExact(valF13, "param15yrlump", "D", "F", 13, 68, 2);
        return vlookupMatch
    }

// =F27 * Xinput_wozwaarde
    sheet_15yrlump_F28() {
        let valF27 = parseFloat(this.getValue("15yrlump", "F27") || 0);
        let woz = parseFloat(this.getValue("15yrlump", "F16") || 0);
        // or if you store the woz in "F16" or do Xinput_wozwaarde()
        return valF27 * woz;
    }

// =IF(F28>param15yrlump!E8, param15yrlump!E8, F28)
    sheet_15yrlump_F29() {
        let valF28 = parseFloat(this.getValue("15yrlump", "F28") || 0);
        let e8 = parseFloat(this.getValue("param15yrlump", "E8") || 0);
        return (valF28 > e8) ? e8 : valF28;
    }

// =MAX( 1000 + (0.005 * Xinput_wozwaarde), (1% * (Xinput_wozwaarde - F19)) )
    sheet_15yrlump_F31() {
        let woz = parseFloat(this.getValue("15yrlump", "F16") || 0);
        let f19val = parseFloat(this.sheet_15yrlump_F19() || 0);
        let val1 = 1000 + (0.005 * woz);
        let val2 = 0.01*(woz - f19val);
        return Math.max(val1, val2);
    }

// =VLOOKUP(F13, param15yrlump!D23:F51, 3, FALSE)
    sheet_15yrlump_F34() {
        let valF13 = parseFloat(this.getValue("15yrlump", "F13") || 0);
        return this._vlookupExact(valF13, "param15yrlump", "D", "F", 23, 51, 3);
    }

// =F34 * Xinput_wozwaarde
    sheet_15yrlump_F35() {
        let valF34 = parseFloat(this.sheet_15yrlump_F34() || 0);
        let wozVal = parseFloat(this.getValue("15yrlump", "F16") || 0);
        return valF34 * wozVal;
    }

// = + fixed 5.000
// (In your Python code, E36 = 5000. It's basically a constant cell.)
    sheet_15yrlump_E36() {
        return 5000;
    }

// = F35 + param15yrlump!$E$7
    sheet_15yrlump_F36() {
        let f35 = parseFloat(this.sheet_15yrlump_F35() || 0);
        let e7 = parseFloat(this.getValue("param15yrlump", "E7") || 0);
        return f35 + e7;
    }

// =IF(F86="Yes", SUM(H50:H64), 0)  (Python code had it as F84, but we can replicate)
    sheet_15yrlump_F84() {
        let f86Val = this.getValue("15yrlump", "F86");
        if (f86Val === "Yes") {
            let total = 0;
            for (let r=50; r<=64; r++){
                let v = parseFloat(this.getValue("15yrlump", `H${r}`) || 0);
                total += v;
            }
            return total;
        }
        return 0;
    }

// = param15yrlump!E5 * 100
    sheet_15yrlump_F85() {
        let e5 = parseFloat(this.getValue("param15yrlump", "E5") || 0);
        return e5 * 100;
    }

// =IF(H50>F31, "Yes","No")
    sheet_15yrlump_J85() {
        let h50 = parseFloat(this.getValue("15yrlump", "H50") || 0);
        let f31 = parseFloat(this.sheet_15yrlump_F31() || 0);
        return (h50 > f31) ? "Yes" : "No";
    }

// =IF(AND(J85="Yes", J86="Yes"), "Yes", "No")
    sheet_15yrlump_F86() {
        let j85 = this.sheet_15yrlump_J85();
        let j86 = this.sheet_15yrlump_J86();
        if (j85 === "Yes" && j86 === "Yes") {
            return "Yes";
        }
        return "No";
    }

    sheet_param15yrlump_F43() {
        return this.getValue("param15yrlump", "F42");
    }
    sheet_param15yrlump_F44() {
        return this.getValue("param15yrlump", "F43");
    }
    sheet_param15yrlump_F45() {
        return this.getValue("param15yrlump", "F44");
    }
    sheet_param15yrlump_F46() {
        return this.getValue("param15yrlump", "F45");
    }
    sheet_param15yrlump_F47() {
        return this.getValue("param15yrlump", "F46");
    }
    sheet_param15yrlump_F48() {
        return this.getValue("param15yrlump", "F47");
    }
    sheet_param15yrlump_F49() {
        return this.getValue("param15yrlump", "F48");
    }
    sheet_param15yrlump_F50() {
        return this.getValue("param15yrlump", "F49");
    }
    sheet_param15yrlump_F51() {
        return this.getValue("param15yrlump", "F50");
    }
    sheet_param15yrlump_F52() {
        return this.getValue("param15yrlump", "F51");
    }
    sheet_param15yrlump_F53() {
        return this.getValue("param15yrlump", "F52");
    }
    sheet_param15yrlump_F54() {
        return this.getValue("param15yrlump", "F53");
    }
    sheet_param15yrlump_F55() {
        return this.getValue("param15yrlump", "F54");
    }
    sheet_param15yrlump_F56() {
        return this.getValue("param15yrlump", "F55");
    }
    sheet_param15yrlump_F57() {
        return this.getValue("param15yrlump", "F56");
    }
    sheet_param15yrlump_F58() {
        return this.getValue("param15yrlump", "F57");
    }
    sheet_param15yrlump_F59() {
        return this.getValue("param15yrlump", "F58");
    }
    sheet_param15yrlump_F60() {
        return this.getValue("param15yrlump", "F59");
    }
    sheet_param15yrlump_F61() {
        return this.getValue("param15yrlump", "F60");
    }
    sheet_param15yrlump_F62() {
        return this.getValue("param15yrlump", "F61");
    }
    sheet_param15yrlump_F63() {
        return this.getValue("param15yrlump", "F62");
    }
    sheet_param15yrlump_F64() {
        return this.getValue("param15yrlump", "F63");
    }
    sheet_param15yrlump_F65() {
        return this.getValue("param15yrlump", "F64");
    }
    sheet_param15yrlump_F66() {
        return this.getValue("param15yrlump", "F65");
    }
    sheet_param15yrlump_F67() {
        return this.getValue("param15yrlump", "F66");
    }
    sheet_param15yrlump_F68() {
        return this.getValue("param15yrlump", "F67");
    }

// =IF(SUM(H50:H64)>25000, "Yes","No")
    sheet_15yrlump_J86() {
        let sumH = 0;
        for (let r=50; r<=64; r++){
            sumH += parseFloat(this.getValue("15yrlump", `H${r}`) || 0);
        }
        if (sumH > 25000) {
            return "Yes";
        }
        return "No";
    }

    // ... and so on, replicating the rest.

    /********************************************************************
     *  D50..D79
     ********************************************************************/
    sheet_15yrlump_D50() {
        let val = this.getDValueForRow(50);
        return val
    }
    sheet_15yrlump_D51() {
        return this.getDValueForRow(51);
    }
    sheet_15yrlump_D52() {
        return this.getDValueForRow(52);
    }
    sheet_15yrlump_D53() {
        return this.getDValueForRow(53);
    }
    sheet_15yrlump_D54() {
        return this.getDValueForRow(54);
    }
    sheet_15yrlump_D55() {
        return this.getDValueForRow(55);
    }
    sheet_15yrlump_D56() {
        return this.getDValueForRow(56);
    }
    sheet_15yrlump_D57() {
        return this.getDValueForRow(57);
    }
    sheet_15yrlump_D58() {
        return this.getDValueForRow(58);
    }
    sheet_15yrlump_D59() {
        return this.getDValueForRow(59);
    }
    sheet_15yrlump_D60() {
        return this.getDValueForRow(60);
    }
    sheet_15yrlump_D61() {
        return this.getDValueForRow(61);
    }
    sheet_15yrlump_D62() {
        return this.getDValueForRow(62);
    }
    sheet_15yrlump_D63() {
        return this.getDValueForRow(63);
    }
    sheet_15yrlump_D64() {
        return this.getDValueForRow(64);
    }
    sheet_15yrlump_D65() {
        return this.getDValueForRow(65);
    }
    sheet_15yrlump_D66() {
        return this.getDValueForRow(66);
    }
    sheet_15yrlump_D67() {
        return this.getDValueForRow(67);
    }
    sheet_15yrlump_D68() {
        return this.getDValueForRow(68);
    }
    sheet_15yrlump_D69() {
        return this.getDValueForRow(69);
    }
    sheet_15yrlump_D70() {
        return this.getDValueForRow(70);
    }
    sheet_15yrlump_D71() {
        return this.getDValueForRow(71);
    }
    sheet_15yrlump_D72() {
        return this.getDValueForRow(72);
    }
    sheet_15yrlump_D73() {
        return this.getDValueForRow(73);
    }
    sheet_15yrlump_D74() {
        return this.getDValueForRow(74);
    }
    sheet_15yrlump_D75() {
        return this.getDValueForRow(75);
    }
    sheet_15yrlump_D76() {
        return this.getDValueForRow(76);
    }
    sheet_15yrlump_D77() {
        return this.getDValueForRow(77);
    }
    sheet_15yrlump_D78() {
        return this.getDValueForRow(78);
    }
    sheet_15yrlump_D79() {
        return this.getDValueForRow(79);
    }

    /********************************************************************
     *  E50..E79
     ********************************************************************/
    sheet_15yrlump_E50() {
        let val = this.getEValueForRow(50);
        return val
    }
    sheet_15yrlump_E51() {
        return this.getEValueForRow(51);
    }
    sheet_15yrlump_E52() {
        return this.getEValueForRow(52);
    }
    sheet_15yrlump_E53() {
        return this.getEValueForRow(53);
    }
    sheet_15yrlump_E54() {
        return this.getEValueForRow(54);
    }
    sheet_15yrlump_E55() {
        return this.getEValueForRow(55);
    }
    sheet_15yrlump_E56() {
        return this.getEValueForRow(56);
    }
    sheet_15yrlump_E57() {
        return this.getEValueForRow(57);
    }
    sheet_15yrlump_E58() {
        return this.getEValueForRow(58);
    }
    sheet_15yrlump_E59() {
        return this.getEValueForRow(59);
    }
    sheet_15yrlump_E60() {
        return this.getEValueForRow(60);
    }
    sheet_15yrlump_E61() {
        return this.getEValueForRow(61);
    }
    sheet_15yrlump_E62() {
        return this.getEValueForRow(62);
    }
    sheet_15yrlump_E63() {
        return this.getEValueForRow(63);
    }
    sheet_15yrlump_E64() {
        return this.getEValueForRow(64);
    }
    sheet_15yrlump_E65() {
        return this.getEValueForRow(65);
    }
    sheet_15yrlump_E66() {
        return this.getEValueForRow(66);
    }
    sheet_15yrlump_E67() {
        return this.getEValueForRow(67);
    }
    sheet_15yrlump_E68() {
        return this.getEValueForRow(68);
    }
    sheet_15yrlump_E69() {
        return this.getEValueForRow(69);
    }
    sheet_15yrlump_E70() {
        return this.getEValueForRow(70);
    }
    sheet_15yrlump_E71() {
        return this.getEValueForRow(71);
    }
    sheet_15yrlump_E72() {
        return this.getEValueForRow(72);
    }
    sheet_15yrlump_E73() {
        return this.getEValueForRow(73);
    }
    sheet_15yrlump_E74() {
        return this.getEValueForRow(74);
    }
    sheet_15yrlump_E75() {
        return this.getEValueForRow(75);
    }
    sheet_15yrlump_E76() {
        return this.getEValueForRow(76);
    }
    sheet_15yrlump_E77() {
        return this.getEValueForRow(77);
    }
    sheet_15yrlump_E78() {
        return this.getEValueForRow(78);
    }
    sheet_15yrlump_E79() {
        return this.getEValueForRow(79);
    }

    /********************************************************************
     *  F50..F79
     ********************************************************************/
    sheet_15yrlump_F50() {
        let val = this.getFValueForRow(50);
        return val
    }
    sheet_15yrlump_F51() {
        return this.getFValueForRow(51);
    }
    sheet_15yrlump_F52() {
        return this.getFValueForRow(52);
    }
    sheet_15yrlump_F53() {
        return this.getFValueForRow(53);
    }
    sheet_15yrlump_F54() {
        return this.getFValueForRow(54);
    }
    sheet_15yrlump_F55() {
        return this.getFValueForRow(55);
    }
    sheet_15yrlump_F56() {
        return this.getFValueForRow(56);
    }
    sheet_15yrlump_F57() {
        return this.getFValueForRow(57);
    }
    sheet_15yrlump_F58() {
        return this.getFValueForRow(58);
    }
    sheet_15yrlump_F59() {
        return this.getFValueForRow(59);
    }
    sheet_15yrlump_F60() {
        return this.getFValueForRow(60);
    }
    sheet_15yrlump_F61() {
        return this.getFValueForRow(61);
    }
    sheet_15yrlump_F62() {
        return this.getFValueForRow(62);
    }
    sheet_15yrlump_F63() {
        return this.getFValueForRow(63);
    }
    sheet_15yrlump_F64() {
        return this.getFValueForRow(64);
    }
    sheet_15yrlump_F65() {
        return this.getFValueForRow(65);
    }
    sheet_15yrlump_F66() {
        return this.getFValueForRow(66);
    }
    sheet_15yrlump_F67() {
        return this.getFValueForRow(67);
    }
    sheet_15yrlump_F68() {
        return this.getFValueForRow(68);
    }
    sheet_15yrlump_F69() {
        return this.getFValueForRow(69);
    }
    sheet_15yrlump_F70() {
        return this.getFValueForRow(70);
    }
    sheet_15yrlump_F71() {
        return this.getFValueForRow(71);
    }
    sheet_15yrlump_F72() {
        return this.getFValueForRow(72);
    }
    sheet_15yrlump_F73() {
        return this.getFValueForRow(73);
    }
    sheet_15yrlump_F74() {
        return this.getFValueForRow(74);
    }
    sheet_15yrlump_F75() {
        return this.getFValueForRow(75);
    }
    sheet_15yrlump_F76() {
        return this.getFValueForRow(76);
    }
    sheet_15yrlump_F77() {
        return this.getFValueForRow(77);
    }
    sheet_15yrlump_F78() {
        return this.getFValueForRow(78);
    }
    sheet_15yrlump_F79() {
        return this.getFValueForRow(79);
    }

    /********************************************************************
     *  G50..G79
     ********************************************************************/
    sheet_15yrlump_G50() {
        return this.getGValueForRow(50);
    }
    sheet_15yrlump_G51() {
        return this.getGValueForRow(51);
    }
    sheet_15yrlump_G52() {
        return this.getGValueForRow(52);
    }
    sheet_15yrlump_G53() {
        return this.getGValueForRow(53);
    }
    sheet_15yrlump_G54() {
        return this.getGValueForRow(54);
    }
    sheet_15yrlump_G55() {
        return this.getGValueForRow(55);
    }
    sheet_15yrlump_G56() {
        return this.getGValueForRow(56);
    }
    sheet_15yrlump_G57() {
        return this.getGValueForRow(57);
    }
    sheet_15yrlump_G58() {
        return this.getGValueForRow(58);
    }
    sheet_15yrlump_G59() {
        return this.getGValueForRow(59);
    }
    sheet_15yrlump_G60() {
        return this.getGValueForRow(60);
    }
    sheet_15yrlump_G61() {
        return this.getGValueForRow(61);
    }
    sheet_15yrlump_G62() {
        return this.getGValueForRow(62);
    }
    sheet_15yrlump_G63() {
        return this.getGValueForRow(63);
    }
    sheet_15yrlump_G64() {
        let val = this.getGValueForRow(64);
        return this.getGValueForRow(64);
    }
    sheet_15yrlump_G65() {
        return this.getGValueForRow(65);
    }
    sheet_15yrlump_G66() {
        return this.getGValueForRow(66);
    }
    sheet_15yrlump_G67() {
        return this.getGValueForRow(67);
    }
    sheet_15yrlump_G68() {
        return this.getGValueForRow(68);
    }
    sheet_15yrlump_G69() {
        return this.getGValueForRow(69);
    }
    sheet_15yrlump_G70() {
        return this.getGValueForRow(70);
    }
    sheet_15yrlump_G71() {
        return this.getGValueForRow(71);
    }
    sheet_15yrlump_G72() {
        return this.getGValueForRow(72);
    }
    sheet_15yrlump_G73() {
        return this.getGValueForRow(73);
    }
    sheet_15yrlump_G74() {
        return this.getGValueForRow(74);
    }
    sheet_15yrlump_G75() {
        return this.getGValueForRow(75);
    }
    sheet_15yrlump_G76() {
        return this.getGValueForRow(76);
    }
    sheet_15yrlump_G77() {
        return this.getGValueForRow(77);
    }
    sheet_15yrlump_G78() {
        return this.getGValueForRow(78);
    }
    sheet_15yrlump_G79() {
        return this.getGValueForRow(79);
    }

    /********************************************************************
     *  I50..I79
     ********************************************************************/
    sheet_15yrlump_I50() {
        let val = this.getIValueForRow(50);
        return val
    }
    sheet_15yrlump_I51() {
        return this.getIValueForRow(51);
    }
    sheet_15yrlump_I52() {
        return this.getIValueForRow(52);
    }
    sheet_15yrlump_I53() {
        return this.getIValueForRow(53);
    }
    sheet_15yrlump_I54() {
        return this.getIValueForRow(54);
    }
    sheet_15yrlump_I55() {
        return this.getIValueForRow(55);
    }
    sheet_15yrlump_I56() {
        return this.getIValueForRow(56);
    }
    sheet_15yrlump_I57() {
        return this.getIValueForRow(57);
    }
    sheet_15yrlump_I58() {
        return this.getIValueForRow(58);
    }
    sheet_15yrlump_I59() {
        return this.getIValueForRow(59);
    }
    sheet_15yrlump_I60() {
        return this.getIValueForRow(60);
    }
    sheet_15yrlump_I61() {
        return this.getIValueForRow(61);
    }
    sheet_15yrlump_I62() {
        return this.getIValueForRow(62);
    }
    sheet_15yrlump_I63() {
        return this.getIValueForRow(63);
    }
    sheet_15yrlump_I64() {
        return this.getIValueForRow(64);
    }
    sheet_15yrlump_I65() {
        return this.getIValueForRow(65);
    }
    sheet_15yrlump_I66() {
        return this.getIValueForRow(66);
    }
    sheet_15yrlump_I67() {
        return this.getIValueForRow(67);
    }
    sheet_15yrlump_I68() {
        return this.getIValueForRow(68);
    }
    sheet_15yrlump_I69() {
        return this.getIValueForRow(69);
    }
    sheet_15yrlump_I70() {
        return this.getIValueForRow(70);
    }
    sheet_15yrlump_I71() {
        return this.getIValueForRow(71);
    }
    sheet_15yrlump_I72() {
        return this.getIValueForRow(72);
    }
    sheet_15yrlump_I73() {
        return this.getIValueForRow(73);
    }
    sheet_15yrlump_I74() {
        return this.getIValueForRow(74);
    }
    sheet_15yrlump_I75() {
        return this.getIValueForRow(75);
    }
    sheet_15yrlump_I76() {
        return this.getIValueForRow(76);
    }
    sheet_15yrlump_I77() {
        return this.getIValueForRow(77);
    }
    sheet_15yrlump_I78() {
        return this.getIValueForRow(78);
    }
    sheet_15yrlump_I79() {
        return this.getIValueForRow(79);
    }

    /********************************************************************
     *  J50..J79
     ********************************************************************/
    sheet_15yrlump_J50() {
        let val = this.getJValueForRow(50);
        return val
    }
    sheet_15yrlump_J51() {
        return this.getJValueForRow(51);
    }
    sheet_15yrlump_J52() {
        return this.getJValueForRow(52);
    }
    sheet_15yrlump_J53() {
        return this.getJValueForRow(53);
    }
    sheet_15yrlump_J54() {
        return this.getJValueForRow(54);
    }
    sheet_15yrlump_J55() {
        return this.getJValueForRow(55);
    }
    sheet_15yrlump_J56() {
        return this.getJValueForRow(56);
    }
    sheet_15yrlump_J57() {
        return this.getJValueForRow(57);
    }
    sheet_15yrlump_J58() {
        return this.getJValueForRow(58);
    }
    sheet_15yrlump_J59() {
        return this.getJValueForRow(59);
    }
    sheet_15yrlump_J60() {
        return this.getJValueForRow(60);
    }
    sheet_15yrlump_J61() {
        return this.getJValueForRow(61);
    }
    sheet_15yrlump_J62() {
        return this.getJValueForRow(62);
    }
    sheet_15yrlump_J63() {
        return this.getJValueForRow(63);
    }
    sheet_15yrlump_J64() {
        return this.getJValueForRow(64);
    }
    sheet_15yrlump_J65() {
        return this.getJValueForRow(65);
    }
    sheet_15yrlump_J66() {
        return this.getJValueForRow(66);
    }
    sheet_15yrlump_J67() {
        return this.getJValueForRow(67);
    }
    sheet_15yrlump_J68() {
        return this.getJValueForRow(68);
    }
    sheet_15yrlump_J69() {
        return this.getJValueForRow(69);
    }
    sheet_15yrlump_J70() {
        return this.getJValueForRow(70);
    }
    sheet_15yrlump_J71() {
        return this.getJValueForRow(71);
    }
    sheet_15yrlump_J72() {
        return this.getJValueForRow(72);
    }
    sheet_15yrlump_J73() {
        return this.getJValueForRow(73);
    }
    sheet_15yrlump_J74() {
        return this.getJValueForRow(74);
    }
    sheet_15yrlump_J75() {
        return this.getJValueForRow(75);
    }
    sheet_15yrlump_J76() {
        return this.getJValueForRow(76);
    }
    sheet_15yrlump_J77() {
        return this.getJValueForRow(77);
    }
    sheet_15yrlump_J78() {
        return this.getJValueForRow(78);
    }
    sheet_15yrlump_J79() {
        return this.getJValueForRow(79);
    }

    /********************************************************************
     *  K50..K79
     ********************************************************************/
    sheet_15yrlump_K50() {
        let val = this.getKValueForRow(50);
        return val
    }
    sheet_15yrlump_K51() {
        return this.getKValueForRow(51);
    }
    sheet_15yrlump_K52() {
        return this.getKValueForRow(52);
    }
    sheet_15yrlump_K53() {
        return this.getKValueForRow(53);
    }
    sheet_15yrlump_K54() {
        return this.getKValueForRow(54);
    }
    sheet_15yrlump_K55() {
        return this.getKValueForRow(55);
    }
    sheet_15yrlump_K56() {
        return this.getKValueForRow(56);
    }
    sheet_15yrlump_K57() {
        return this.getKValueForRow(57);
    }
    sheet_15yrlump_K58() {
        return this.getKValueForRow(58);
    }
    sheet_15yrlump_K59() {
        return this.getKValueForRow(59);
    }
    sheet_15yrlump_K60() {
        return this.getKValueForRow(60);
    }
    sheet_15yrlump_K61() {
        return this.getKValueForRow(61);
    }
    sheet_15yrlump_K62() {
        return this.getKValueForRow(62);
    }
    sheet_15yrlump_K63() {
        return this.getKValueForRow(63);
    }
    sheet_15yrlump_K64() {
        return this.getKValueForRow(64);
    }
    sheet_15yrlump_K65() {
        return this.getKValueForRow(65);
    }
    sheet_15yrlump_K66() {
        return this.getKValueForRow(66);
    }
    sheet_15yrlump_K67() {
        return this.getKValueForRow(67);
    }
    sheet_15yrlump_K68() {
        return this.getKValueForRow(68);
    }
    sheet_15yrlump_K69() {
        return this.getKValueForRow(69);
    }
    sheet_15yrlump_K70() {
        return this.getKValueForRow(70);
    }
    sheet_15yrlump_K71() {
        return this.getKValueForRow(71);
    }
    sheet_15yrlump_K72() {
        return this.getKValueForRow(72);
    }
    sheet_15yrlump_K73() {
        return this.getKValueForRow(73);
    }
    sheet_15yrlump_K74() {
        return this.getKValueForRow(74);
    }
    sheet_15yrlump_K75() {
        return this.getKValueForRow(75);
    }
    sheet_15yrlump_K76() {
        return this.getKValueForRow(76);
    }
    sheet_15yrlump_K77() {
        return this.getKValueForRow(77);
    }
    sheet_15yrlump_K78() {
        return this.getKValueForRow(78);
    }
    sheet_15yrlump_K79() {
        return this.getKValueForRow(79);
    }

    /********************************************************************
     *  L50..L79
     ********************************************************************/
    sheet_15yrlump_L50() {
        let val = this.getLValueForRow(50);
        return val
    }
    sheet_15yrlump_L51() {
        return this.getLValueForRow(51);
    }
    sheet_15yrlump_L52() {
        return this.getLValueForRow(52);
    }
    sheet_15yrlump_L53() {
        return this.getLValueForRow(53);
    }
    sheet_15yrlump_L54() {
        return this.getLValueForRow(54);
    }
    sheet_15yrlump_L55() {
        return this.getLValueForRow(55);
    }
    sheet_15yrlump_L56() {
        return this.getLValueForRow(56);
    }
    sheet_15yrlump_L57() {
        return this.getLValueForRow(57);
    }
    sheet_15yrlump_L58() {
        return this.getLValueForRow(58);
    }
    sheet_15yrlump_L59() {
        return this.getLValueForRow(59);
    }
    sheet_15yrlump_L60() {
        return this.getLValueForRow(60);
    }
    sheet_15yrlump_L61() {
        return this.getLValueForRow(61);
    }
    sheet_15yrlump_L62() {
        return this.getLValueForRow(62);
    }
    sheet_15yrlump_L63() {
        return this.getLValueForRow(63);
    }
    sheet_15yrlump_L64() {
        let val = this.getLValueForRow(64);
        return val
    }
    sheet_15yrlump_L65() {
        return this.getLValueForRow(65);
    }
    sheet_15yrlump_L66() {
        return this.getLValueForRow(66);
    }
    sheet_15yrlump_L67() {
        return this.getLValueForRow(67);
    }
    sheet_15yrlump_L68() {
        return this.getLValueForRow(68);
    }
    sheet_15yrlump_L69() {
        return this.getLValueForRow(69);
    }
    sheet_15yrlump_L70() {
        return this.getLValueForRow(70);
    }
    sheet_15yrlump_L71() {
        return this.getLValueForRow(71);
    }
    sheet_15yrlump_L72() {
        return this.getLValueForRow(72);
    }
    sheet_15yrlump_L73() {
        return this.getLValueForRow(73);
    }
    sheet_15yrlump_L74() {
        return this.getLValueForRow(74);
    }
    sheet_15yrlump_L75() {
        return this.getLValueForRow(75);
    }
    sheet_15yrlump_L76() {
        return this.getLValueForRow(76);
    }
    sheet_15yrlump_L77() {
        return this.getLValueForRow(77);
    }
    sheet_15yrlump_L78() {
        return this.getLValueForRow(78);
    }
    sheet_15yrlump_L79() {
        return this.getLValueForRow(79);
    }

    /********************************************************************
     *  M50..M79
     ********************************************************************/
    sheet_15yrlump_M50() {
        let val = this.getMValueForRow(50);
        return val
    }
    sheet_15yrlump_M51() {
        return this.getMValueForRow(51);
    }
    sheet_15yrlump_M52() {
        return this.getMValueForRow(52);
    }
    sheet_15yrlump_M53() {
        return this.getMValueForRow(53);
    }
    sheet_15yrlump_M54() {
        return this.getMValueForRow(54);
    }
    sheet_15yrlump_M55() {
        return this.getMValueForRow(55);
    }
    sheet_15yrlump_M56() {
        return this.getMValueForRow(56);
    }
    sheet_15yrlump_M57() {
        return this.getMValueForRow(57);
    }
    sheet_15yrlump_M58() {
        return this.getMValueForRow(58);
    }
    sheet_15yrlump_M59() {
        return this.getMValueForRow(59);
    }
    sheet_15yrlump_M60() {
        return this.getMValueForRow(60);
    }
    sheet_15yrlump_M61() {
        return this.getMValueForRow(61);
    }
    sheet_15yrlump_M62() {
        return this.getMValueForRow(62);
    }
    sheet_15yrlump_M63() {
        return this.getMValueForRow(63);
    }
    sheet_15yrlump_M64() {
        let val = this.getMValueForRow(64);
        return val
    }
    sheet_15yrlump_M65() {
        return this.getMValueForRow(65);
    }
    sheet_15yrlump_M66() {
        return this.getMValueForRow(66);
    }
    sheet_15yrlump_M67() {
        return this.getMValueForRow(67);
    }
    sheet_15yrlump_M68() {
        return this.getMValueForRow(68);
    }
    sheet_15yrlump_M69() {
        return this.getMValueForRow(69);
    }
    sheet_15yrlump_M70() {
        return this.getMValueForRow(70);
    }
    sheet_15yrlump_M71() {
        return this.getMValueForRow(71);
    }
    sheet_15yrlump_M72() {
        return this.getMValueForRow(72);
    }
    sheet_15yrlump_M73() {
        return this.getMValueForRow(73);
    }
    sheet_15yrlump_M74() {
        return this.getMValueForRow(74);
    }
    sheet_15yrlump_M75() {
        return this.getMValueForRow(75);
    }
    sheet_15yrlump_M76() {
        return this.getMValueForRow(76);
    }
    sheet_15yrlump_M77() {
        return this.getMValueForRow(77);
    }
    sheet_15yrlump_M78() {
        return this.getMValueForRow(78);
    }
    sheet_15yrlump_M79() {
        return this.getMValueForRow(79);
    }

    /********************************************************************
     *  N50..N79
     *
     *  Some code uses "N(row) = M(row)/F(row)"
     *  or we can call getNValueForRow(rowNum).
     ********************************************************************/
    sheet_15yrlump_N50() {
        let val = this.getNValueForRow(50);
        return val
    }
    sheet_15yrlump_N51() {
        return this.getNValueForRow(51);
    }
    sheet_15yrlump_N52() {
        return this.getNValueForRow(52);
    }
    sheet_15yrlump_N53() {
        return this.getNValueForRow(53);
    }
    sheet_15yrlump_N54() {
        return this.getNValueForRow(54);
    }
    sheet_15yrlump_N55() {
        return this.getNValueForRow(55);
    }
    sheet_15yrlump_N56() {
        return this.getNValueForRow(56);
    }
    sheet_15yrlump_N57() {
        return this.getNValueForRow(57);
    }
    sheet_15yrlump_N58() {
        return this.getNValueForRow(58);
    }
    sheet_15yrlump_N59() {
        return this.getNValueForRow(59);
    }
    sheet_15yrlump_N60() {
        return this.getNValueForRow(60);
    }
    sheet_15yrlump_N61() {
        return this.getNValueForRow(61);
    }
    sheet_15yrlump_N62() {
        return this.getNValueForRow(62);
    }
    sheet_15yrlump_N63() {
        return this.getNValueForRow(63);
    }
    sheet_15yrlump_N64() {
        return this.getNValueForRow(64);
    }
    sheet_15yrlump_N65() {
        return this.getNValueForRow(65);
    }
    sheet_15yrlump_N66() {
        return this.getNValueForRow(66);
    }
    sheet_15yrlump_N67() {
        return this.getNValueForRow(67);
    }
    sheet_15yrlump_N68() {
        return this.getNValueForRow(68);
    }
    sheet_15yrlump_N69() {
        return this.getNValueForRow(69);
    }
    sheet_15yrlump_N70() {
        return this.getNValueForRow(70);
    }
    sheet_15yrlump_N71() {
        return this.getNValueForRow(71);
    }
    sheet_15yrlump_N72() {
        return this.getNValueForRow(72);
    }
    sheet_15yrlump_N73() {
        return this.getNValueForRow(73);
    }
    sheet_15yrlump_N74() {
        return this.getNValueForRow(74);
    }
    sheet_15yrlump_N75() {
        return this.getNValueForRow(75);
    }
    sheet_15yrlump_N76() {
        return this.getNValueForRow(76);
    }
    sheet_15yrlump_N77() {
        return this.getNValueForRow(77);
    }
    sheet_15yrlump_N78() {
        return this.getNValueForRow(78);
    }
    sheet_15yrlump_N79() {
        return this.getNValueForRow(79);
    }




    _excelRound(value, digits) {
        // let factor = Math.pow(10, digits);
        // return Math.round(value * factor) / factor;

        const factor = Math.pow(10, digits);
        const shiftedValue = value * factor;
        const roundedValue = Math.floor(shiftedValue + 0.5);

        // Check if it was exactly .5 and apply Bankers' Rounding
        if (Math.abs(shiftedValue - roundedValue) === 0.5) {
            // Round to the nearest even number
            return (Math.floor(shiftedValue) % 2 === 0)
                ? Math.floor(shiftedValue) / factor
                : Math.ceil(shiftedValue) / factor;
        }

        return roundedValue / factor;
    }

    _calculate_formula_method_map(){
        return {
            '15yrlump,F6': 'sheet_15yrlump_F6',
            '15yrlump,F11': 'sheet_15yrlump_F11',
            '15yrlump,F12': 'sheet_15yrlump_F12',
            '15yrlump,F13': 'sheet_15yrlump_F13',
            '15yrlump,F19': 'sheet_15yrlump_F19',
            '15yrlump,F27': 'sheet_15yrlump_F27',
            '15yrlump,F28': 'sheet_15yrlump_F28',
            '15yrlump,F29': 'sheet_15yrlump_F29',
            '15yrlump,F31': 'sheet_15yrlump_F31',
            '15yrlump,F34': 'sheet_15yrlump_F34',
            '15yrlump,F35': 'sheet_15yrlump_F35',
            '15yrlump,E36': 'sheet_15yrlump_E36',
            '15yrlump,F36': 'sheet_15yrlump_F36',
            '15yrlump,F84': 'sheet_15yrlump_F84',
            '15yrlump,F85': 'sheet_15yrlump_F85',
            '15yrlump,J85': 'sheet_15yrlump_J85',
            '15yrlump,F86': 'sheet_15yrlump_F86',
            '15yrlump,J86': 'sheet_15yrlump_J86'
        }
    }

    // A minimal “calculate” function. In Python, you recalc formulas.
    // In Node, we might just call the key formula methods and store them:
    /**
     * Partial recalc approach:
     *   - only recalc formula cells in this._dirty_formula_cells
     *   - do them in an arbitrary order once
     *   - clear the dirty set
     * If your formulas feed each other, you might need a topological order or repeated passes.
     */
    calculate() {
        // const formulaMap = this._formulaMethodMap();
        const formulaMap = this._calculate_formula_method_map()
        for (let key of Object.keys(formulaMap)) {
            let methodName = formulaMap[key];
            if (methodName && typeof this[methodName] === "function") {
                let val = this[methodName]();
                let [sheet, cell] = key.split(",");
                this.setValue(sheet.trim(), cell.trim(), val);
            }
        }
        // (b) run the row-based approach (D/E/F/G..., etc.)
        this.calculate_rows_15yrlump();
    }

    /**
     * Handle all repeated formulas for rows 50..79 in one pass.
     * We read F16, F19, param15yrlump!E5, E6, etc. as needed, then we loop row=50..79
     * and compute Drow, Erow, Frow, ... Nrow in one go.
     * We store the results in _staticValues.
     */
    calculate_rows_15yrlump() {
        // read some inputs
        let f16_val = parseFloat(this.getValue("15yrlump","F16") || 0);
        let f19_val = parseFloat(this.getValue("15yrlump","F19") || 0);
        let e5_val  = parseFloat(this.getValue("param15yrlump","E5") || 0);
        let e6_val  = parseFloat(this.getValue("param15yrlump","E6") || 0);

        for (let row = 50; row <= 79; row++) {
            let d_prev = 0;
            if (row === 50) {

                d_prev = this.Xinput_datumvandaag();
            } else {
                d_prev = parseFloat(this._staticValues.get("15yrlump," + `D${row-1}`) || 0);
            }
            let d_val = parseInt(d_prev,10)+1;
            this.setValue("15yrlump", `D${row}`, d_val);

            // E(row) = E(row-1)+1, row=50 => E(49) => or F13
            let e_val;
            if (row===50) {
                let e_prev = parseFloat(this.getValue("15yrlump","F13") || 0);
                e_val = e_prev+1;
            } else {
                let e_prev = parseFloat(this._staticValues.get("15yrlump," + `E${row-1}`)||0);
                e_val = e_prev+1;
            }
            this.setValue("15yrlump", `E${row}`, e_val);

            // F(row) => $F$16
            this.setValue("15yrlump", `F${row}`, f16_val);

            // G(row) => $F$19
            this.setValue("15yrlump", `G${row}`, f19_val);

            // I(row) => param15yrlump!E5 => e5_val
            this.setValue("15yrlump", `I${row}`, e5_val);

            // J(row) => Round( (H(row)+L(row-1))*(1+I/(1-I))^(1/12)^12 - (H(row)+L(row-1)) , e6_val)
            let h_val = parseFloat(this._staticValues.get("15yrlump," + `H${row}`)||0);
            let l_prev=0;
            if (row===50) {
                // let hv = parseFloat(this.getValue("15yrlump","H50")||0);
                // let kv = parseFloat(this.getValue("15yrlump","K50")||0);
                // l_prev= hv+kv;
                l_prev = 0;
            } else {
                l_prev= parseFloat(this._staticValues.get("15yrlump," + `L${row-1}`)||0);
            }
            // i_val => e5_val
            let i_val = e5_val;
            let base= (Math.abs(1-i_val)<1e-12)?1e9:(1+ i_val/(1-i_val));
            let monthly= Math.pow(base,1/12);
            let final_= Math.pow(monthly,12);
            let expr= (h_val + l_prev)*final_ - (h_val + l_prev);

            const excelRound = (x, digits)=>{
                let factor = Math.pow(10,digits);
                return Math.round(x*factor)/factor;
            };
            let j_val = excelRound(expr,e6_val);
            this.setValue("15yrlump", `J${row}`, j_val);

            // K(row) = J(row)+K(row-1)
            let k_prev = (row===50)
                ? parseFloat(this.getValue("15yrlump","J50")||0)
                : parseFloat(this._staticValues.get("15yrlump," +`K${row-1}`)||0);
            let k_val= j_val + k_prev;
            this.setValue("15yrlump", `K${row}`, k_val);

            // L(row) = H(row)+J(row)+L(row-1)
            let l_val = h_val + j_val + l_prev;
            this.setValue("15yrlump", `L${row}`, l_val);

            // M(row) = L(row)+ G(row)
            let m_val = l_val + f19_val;
            this.setValue("15yrlump", `M${row}`, m_val);

            // N(row) = M(row)/ F(row)
            let n_val = (Math.abs(f16_val)<1e-12) ? 0 : (m_val/f16_val);
            this.setValue("15yrlump", `N${row}`, n_val);
        }
    }

    // lump15yrls_linear() {
    //     // Initial recalc to sync up values.
    //     this.calculate();
    //     const prv_dbl_Max_Mortgage = 550000;
    //
    //     // Partner logic
    //     let g2 = this.Xinput_geboortedatumaanvrager2();
    //     if (g2 && g2 instanceof Date && g2.getFullYear() > 1924) {
    //         this.setValue("15yrlump", "F9", "Ja");
    //     } else {
    //         this.setValue("15yrlump", "F9", "Nee");
    //     }
    //
    //     // Mortgage logic
    //     let saldo = this.Xinput_hypotheeksaldo() || 0;
    //     if (saldo > 0) {
    //         this.setValue("15yrlump", "F17", "Ja");
    //     } else {
    //         this.setValue("15yrlump", "F17", "Nee");
    //     }
    //
    //     // Fill F50..F79 with Xinput_wozwaarde
    //     let wozVal = this.Xinput_wozwaarde();
    //     for (let row = 50; row < 80; row++) {
    //         this.setValue("15yrlump", "F" + row, wozVal);
    //     }
    //
    //     // Optional: checks for total loan and payout – omitted here
    //
    //     let dblGoalSeek = Math.round(this.getValue("15yrlump", "F28") || 0);
    //     let dblMinAnnPay = this.getValue("15yrlump", "F31") || 0;
    //     for (let row = 51; row < 65; row++) {
    //         this.setValue("15yrlump", "H" + row, dblMinAnnPay);
    //     }
    //     let init_h50 = this.getValue("15yrlump", "F36") || 0;
    //     this.setValue("15yrlump", "H50", init_h50);
    //     this.calculate();
    //
    //     // Define _diff to log H50 and rngTotal differences.
    //     const _diff = (x) => {
    //         this.setValue("15yrlump", "H50", x);
    //         this.calculate();
    //         let rt = this.rngTotal() || 0;
    //         console.log(`_diff(${x}): H50=${this._staticValues.get("15yrlump,H50")} rngTotal=${rt} goal=${dblGoalSeek} diff=${rt - dblGoalSeek}`);
    //         return rt - dblGoalSeek;
    //     };
    //
    //     // Measure at x0 and x1
    //     let x0 = init_h50;
    //     let f0 = _diff(x0);
    //     let x1 = init_h50 + 100;
    //     let f1 = _diff(x1);
    //     let slope = (Math.abs(x1 - x0) > 1e-15) ? ((f1 - f0) / (x1 - x0)) : null;
    //     let final_h50;
    //     if (slope === null || Math.abs(slope) < 1e-15) {
    //         final_h50 = init_h50;
    //     } else {
    //         final_h50 = x0 - (f0 / slope);
    //     }
    //     console.log("Final H50 computed:", final_h50);
    //
    //     this.setValue("15yrlump", "H50", final_h50);
    //     this.calculate();
    //     console.log("After final recalc, rngTotal:", this.rngTotal());
    //
    //     let dblTotalMortgage = Math.round(this.getValue("15yrlump", "L64") || 0);
    //     if (dblTotalMortgage > prv_dbl_Max_Mortgage) {
    //         // (Optional second pass if needed)
    //     }
    // }

    lump15yrls_linear() {
        // Initial recalc to sync up values.
        this.calculate();
        const prv_dbl_Max_Mortgage = 550000;

        // Partner logic: if the second applicant’s birthdate exists and its year is > 1924, then set partner to "Ja"
        let g2 = this.Xinput_geboortedatumaanvrager2();
        if (g2 && g2 instanceof Date && g2.getFullYear() > 1924) {
            this.setValue("15yrlump", "Xinput_partnerjanee", "Ja");
        } else {
            this.setValue("15yrlump", "Xinput_partnerjanee", "Nee");
        }

        // Mortgage logic: if the hypotheeksaldo is > 0, mark it as "Ja"
        let saldo = this.Xinput_hypotheeksaldo() || 0;
        if (saldo > 0) {
            this.setValue("15yrlump", "Xinput_hypotheeksaldo", "Ja");
        } else {
            this.setValue("15yrlump", "Xinput_hypotheeksaldo", "Nee");
        }

        // Fill F50..F79 with the WOZ value.
        let wozVal = this.Xinput_wozwaarde();
        for (let row = 50; row < 80; row++) {
            this.setValue("15yrlump", "F" + row, wozVal);
        }

        // Calculate total loan and compare with cap (optional, no action taken here)
        let rngTotal_val = this.rngTotal();
        let dblTotalLoan = saldo + (rngTotal_val || 0);
        let limit_e8 = this.getValue("param15yrlump", "E8") || 99999999;
        if (dblTotalLoan > limit_e8) {
            // (Optional handling)
        }

        // Sum H50 to H64 and compare with a minimum payout (optional)
        let sum_h50_h64 = 0;
        for (let r = 50; r < 65; r++) {
            sum_h50_h64 += this.getValue("15yrlump", "H" + r) || 0;
        }
        let min_e9 = this.getValue("param15yrlump", "E9") || 0;
        if (sum_h50_h64 < min_e9) {
            // (Optional handling)
        }

        // Get the goal seek target and the minimum annual payment.
        let dblGoalSeek = Math.round(this.getValue("15yrlump", "F28") || 0);
        let dblMinAnnPay = this.getValue("15yrlump", "F31") || 0;

        // Set H51..H64 to the minimum annual payment.
        for (let row = 51; row < 65; row++) {
            this.setValue("15yrlump", "H" + row, dblMinAnnPay);
        }

        // Get the initial H50 value from F36, set H50, and recalc.
        let init_h50 = this.getValue("15yrlump", "F36") || 0;
        this.setValue("15yrlump", "H50", init_h50);
        this.calculate();

        // Define a helper function _diff that updates H50, recalculates, and returns the difference (rngTotal - goal).
        const _diff = (x) => {
            this.setValue("15yrlump", "H50", x);
            this.calculate();
            // console.log(`_diff(${x}): H50 =${this._staticValues.get("15yrlump,H50")} rngTotal =${this.rngTotal()} goal =${dblGoalSeek} diff =${this.rngTotal() - dblGoalSeek}`);
            return (this.rngTotal() || 0) - dblGoalSeek;
        };

        // Measure difference at the initial H50 (x0) and at x0 + 100.
        let x0 = init_h50;
        let f0 = _diff(x0);
        let x1 = init_h50 + 100;
        let f1 = _diff(x1);

        // Compute the slope and solve for the H50 that zeroes the difference.
        let slope = (Math.abs(x1 - x0) > 1e-15) ? ((f1 - f0) / (x1 - x0)) : null;
        let final_h50;
        if (slope === null || Math.abs(slope) < 1e-15) {
            final_h50 = init_h50;
        } else {
            final_h50 = x0 - (f0 / slope);
        }

        // console.log("Final H50 computed:", final_h50);

        // Set H50 to the computed value and recalc.
        this.setValue("15yrlump", "H50", final_h50);
        this.calculate();

        // Optionally check total mortgage versus the maximum allowed (not refined here).
        let dblTotalMortgage = Math.round(this.getValue("15yrlump", "L64") || 0);
        if (dblTotalMortgage > prv_dbl_Max_Mortgage) {
            // (Optional second pass)
        }
    }



    // Example macro (like lump15yrls) that might rely on repeated calls.
    // For brevity we’ll do a scaled-down version:
    lump15yrls() {
        // This is the same approach as your snippet:
        // Must do lumpsum logic, then call this.calculate().
        // If you do it AFTER the final read, you'll see no changes.
        // So be sure to call it in preFormulasActions.
        // console.log("running macros")
        let prv_dbl_Max_Mortgage= 550000;

        // partner logic
        let valG2 = this.getValue("15yrlump","F10");
        if (valG2 instanceof Date && valG2.getFullYear()>1924) {
            this.setValue("15yrlump","F9","Ja");
        } else {
            this.setValue("15yrlump","F9","Nee");
        }

        // mortgage logic
        let saldoVal = parseFloat(this.getValue("15yrlump","F18")||0);
        if (saldoVal>0) {
            this.setValue("15yrlump","F17","Ja");
        } else {
            this.setValue("15yrlump","F17","Nee");
        }

        // fill F50..F79 => woz
        let wozVal= parseFloat(this.getValue("15yrlump","F16")||0);
        for (let r=50;r<=79;r++){
            this.setValue("15yrlump",`F${r}`, wozVal);
        }

        // initial calc
        this.calculate();

        let dblGoalSeek= Math.round(parseFloat(this.getValue("15yrlump","F28")||0));
        let dblMinAnnPay= parseFloat(this.getValue("15yrlump","F31")||0);

        // fill H51..H64 => dblMinAnnPay
        for (let r=51;r<=64;r++){
            this.setValue("15yrlump",`H${r}`, dblMinAnnPay);
        }

        // set H50 => F36
        let h50Init= parseFloat(this.getValue("15yrlump","F36")||0);
        this.setValue("15yrlump","H50", h50Init);
        this.calculate();

        const getRngTotal= ()=>{
            return parseFloat(this.getValue("15yrlump","M64")||0);
        };

        let iterationCount=0;
        const maxIters=200;
        while(true){
            let currentVal= Math.round(getRngTotal());
            if (currentVal=== dblGoalSeek) break;
            if (iterationCount> maxIters) break;

            if (iterationCount===0) {
                // sum J50..J64
                let sumJ=0;
                for (let rr=50;rr<=64;rr++){
                    sumJ+= parseFloat(this.getValue("15yrlump",`J${rr}`)||0);
                }
                let dblTemp= dblGoalSeek - (dblMinAnnPay*14) - sumJ;
                this.setValue("15yrlump","H50", dblTemp);
                this.calculate();
            } else {
                let diff= currentVal- dblGoalSeek;
                let oldH50= parseFloat(this.getValue("15yrlump","H50")||0);
                let newVal= oldH50;

                if (currentVal> dblGoalSeek) {
                    if (diff>20000) newVal= oldH50-10000;
                    else if (diff>10000) newVal= oldH50-5000;
                    else if (diff>1000) newVal= oldH50-500;
                    else if (diff>100) newVal= oldH50-50;
                    else if (diff>10) newVal= oldH50-5;
                    else if (diff>5) newVal= oldH50-1;
                    else if (diff>2) newVal= oldH50-1;
                    else if (diff>1) newVal= oldH50-0.5;
                    else if (diff>0.5) newVal= oldH50-0.1;
                    else newVal= oldH50-0.01;
                } else {
                    // currentVal< dblGoalSeek => increment lumpsum
                    let absDiff= Math.abs(diff);
                    if (absDiff>20000) newVal= oldH50+10000;
                    else if (absDiff>10000) newVal= oldH50+5000;
                    else if (absDiff>1000) newVal= oldH50+500;
                    else if (absDiff>100) newVal= oldH50+50;
                    else if (absDiff>10) newVal= oldH50+5;
                    else if (absDiff>5) newVal= oldH50+1;
                    else if (absDiff>2) newVal= oldH50+1;
                    else if (absDiff>1) newVal= oldH50+0.5;
                    else if (absDiff>0.5) newVal= oldH50+0.1;
                    else newVal= oldH50+0.01;
                }

                this.setValue("15yrlump","H50", newVal);
                this.calculate();
            }
            iterationCount++;
        }

        let dblTotalMortgage= Math.round(parseFloat(this.getValue("15yrlump","L64")||0));
        if (dblTotalMortgage> prv_dbl_Max_Mortgage) {
            let iteration2=0;
            while (true){
                if (iteration2>500) break;
                dblTotalMortgage= Math.round(parseFloat(this.getValue("15yrlump","L64")||0));
                let diff2= dblTotalMortgage- prv_dbl_Max_Mortgage;
                if (Math.abs(diff2)<1) break;

                let oldH50_2= parseFloat(this.getValue("15yrlump","H50")||0);
                let newVal2= oldH50_2;

                if (diff2>0){
                    if (diff2>500000) newVal2= oldH50_2-50000;
                    else if (diff2>100000) newVal2= oldH50_2-99000;
                    // etc. your big cascade
                    else newVal2= oldH50_2-0.01;
                } else {
                    newVal2= oldH50_2+0.01;
                }
                this.setValue("15yrlump","H50", newVal2);
                this.calculate();
                iteration2++;
            }
        }
    }

    // If you want macros by name:
    macroMap() {
        return {
            // lump15yrls: () => this.lump15yrls(),
            lump15yrls: () => this.lump15yrls_linear(),
        };
    }

    // The logic from “calculate_output_cells” in Python
    calculateOutputCells(outputCells, preFormulasActions, postFormulasActions) {
        // possibly run “pre” macros
        // if (Array.isArray(postFormulasActions)) {
        //     this.runMacrosAsNeeded(postFormulasActions);
        // }

        // do main “calculate”
        // let results = this._computeOutputResults(outputCells);
        // console.log(results)
        let results;
        // possibly run “post” macros, then recalc + final results
        if (Array.isArray(postFormulasActions)) {
            this.runMacrosAsNeeded(postFormulasActions);
            results = this._computeOutputResults(outputCells);
        }
        return results;
    }

    runMacrosAsNeeded(actions) {
        let macros = this.macroMap();
        for (let action of actions) {
            if (action.type === "macro") {
                let name = action.parameters && action.parameters.name;
                if (name && macros[name]) {
                    macros[name]();
                }
            }
        }
    }

    _computeOutputResults(outputCells) {
        // outputCells is an object like: { "15yrlump": ["F31", "H50"] }
        let finalRes = {};
        for (let sheet in outputCells) {
            let arr = outputCells[sheet];
            let partialRes = {};
            arr.forEach((cellAddr) => {
                partialRes[cellAddr] = this.getValue(sheet, cellAddr);
            });
            finalRes[sheet] = partialRes;
        }
        return finalRes;
    }

    _formulaMethodMap(){
        return {
            "15yrlump,F6": "sheet_15yrlump_F6",
            "15yrlump,F11": "sheet_15yrlump_F11",
            "15yrlump,F12": "sheet_15yrlump_F12",
            "15yrlump,F13": "sheet_15yrlump_F13",
            "15yrlump,F19": "sheet_15yrlump_F19",
            "15yrlump,F27": "sheet_15yrlump_F27",
            "15yrlump,F28": "sheet_15yrlump_F28",
            "15yrlump,F29": "sheet_15yrlump_F29",
            "15yrlump,F31": "sheet_15yrlump_F31",
            "15yrlump,F34": "sheet_15yrlump_F34",
            "15yrlump,F35": "sheet_15yrlump_F35",
            "15yrlump,E36": "sheet_15yrlump_E36",
            "15yrlump,F36": "sheet_15yrlump_F36",

            "15yrlump,D50": "sheet_15yrlump_D50",
            "15yrlump,E50": "sheet_15yrlump_E50",
            "15yrlump,F50": "sheet_15yrlump_F50",
            "15yrlump,G50": "sheet_15yrlump_G50",
            "15yrlump,I50": "sheet_15yrlump_I50",
            "15yrlump,J50": "sheet_15yrlump_J50",
            "15yrlump,K50": "sheet_15yrlump_K50",
            "15yrlump,L50": "sheet_15yrlump_L50",
            "15yrlump,M50": "sheet_15yrlump_M50",
            "15yrlump,N50": "sheet_15yrlump_N50",

            "15yrlump,D51": "sheet_15yrlump_D51",
            "15yrlump,E51": "sheet_15yrlump_E51",
            "15yrlump,F51": "sheet_15yrlump_F51",
            "15yrlump,G51": "sheet_15yrlump_G51",
            "15yrlump,I51": "sheet_15yrlump_I51",
            "15yrlump,J51": "sheet_15yrlump_J51",
            "15yrlump,K51": "sheet_15yrlump_K51",
            "15yrlump,L51": "sheet_15yrlump_L51",
            "15yrlump,M51": "sheet_15yrlump_M51",
            "15yrlump,N51": "sheet_15yrlump_N51",

            "15yrlump,D52": "sheet_15yrlump_D52",
            "15yrlump,E52": "sheet_15yrlump_E52",
            "15yrlump,F52": "sheet_15yrlump_F52",
            "15yrlump,G52": "sheet_15yrlump_G52",
            "15yrlump,I52": "sheet_15yrlump_I52",
            "15yrlump,J52": "sheet_15yrlump_J52",
            "15yrlump,K52": "sheet_15yrlump_K52",
            "15yrlump,L52": "sheet_15yrlump_L52",
            "15yrlump,M52": "sheet_15yrlump_M52",
            "15yrlump,N52": "sheet_15yrlump_N52",

            "15yrlump,D53": "sheet_15yrlump_D53",
            "15yrlump,E53": "sheet_15yrlump_E53",
            "15yrlump,F53": "sheet_15yrlump_F53",
            "15yrlump,G53": "sheet_15yrlump_G53",
            "15yrlump,I53": "sheet_15yrlump_I53",
            "15yrlump,J53": "sheet_15yrlump_J53",
            "15yrlump,K53": "sheet_15yrlump_K53",
            "15yrlump,L53": "sheet_15yrlump_L53",
            "15yrlump,M53": "sheet_15yrlump_M53",
            "15yrlump,N53": "sheet_15yrlump_N53",

            "15yrlump,D54": "sheet_15yrlump_D54",
            "15yrlump,E54": "sheet_15yrlump_E54",
            "15yrlump,F54": "sheet_15yrlump_F54",
            "15yrlump,G54": "sheet_15yrlump_G54",
            "15yrlump,I54": "sheet_15yrlump_I54",
            "15yrlump,J54": "sheet_15yrlump_J54",
            "15yrlump,K54": "sheet_15yrlump_K54",
            "15yrlump,L54": "sheet_15yrlump_L54",
            "15yrlump,M54": "sheet_15yrlump_M54",
            "15yrlump,N54": "sheet_15yrlump_N54",

            // Row 55
            "15yrlump,D55": "sheet_15yrlump_D55",
            "15yrlump,E55": "sheet_15yrlump_E55",
            "15yrlump,F55": "sheet_15yrlump_F55",
            "15yrlump,G55": "sheet_15yrlump_G55",
            "15yrlump,I55": "sheet_15yrlump_I55",
            "15yrlump,J55": "sheet_15yrlump_J55",
            "15yrlump,K55": "sheet_15yrlump_K55",
            "15yrlump,L55": "sheet_15yrlump_L55",
            "15yrlump,M55": "sheet_15yrlump_M55",
            "15yrlump,N55": "sheet_15yrlump_N55",

            // Row 56
            "15yrlump,D56": "sheet_15yrlump_D56",
            "15yrlump,E56": "sheet_15yrlump_E56",
            "15yrlump,F56": "sheet_15yrlump_F56",
            "15yrlump,G56": "sheet_15yrlump_G56",
            "15yrlump,I56": "sheet_15yrlump_I56",
            "15yrlump,J56": "sheet_15yrlump_J56",
            "15yrlump,K56": "sheet_15yrlump_K56",
            "15yrlump,L56": "sheet_15yrlump_L56",
            "15yrlump,M56": "sheet_15yrlump_M56",
            "15yrlump,N56": "sheet_15yrlump_N56",

            // Row 57
            "15yrlump,D57": "sheet_15yrlump_D57",
            "15yrlump,E57": "sheet_15yrlump_E57",
            "15yrlump,F57": "sheet_15yrlump_F57",
            "15yrlump,G57": "sheet_15yrlump_G57",
            "15yrlump,I57": "sheet_15yrlump_I57",
            "15yrlump,J57": "sheet_15yrlump_J57",
            "15yrlump,K57": "sheet_15yrlump_K57",
            "15yrlump,L57": "sheet_15yrlump_L57",
            "15yrlump,M57": "sheet_15yrlump_M57",
            "15yrlump,N57": "sheet_15yrlump_N57",

            // Row 58
            "15yrlump,D58": "sheet_15yrlump_D58",
            "15yrlump,E58": "sheet_15yrlump_E58",
            "15yrlump,F58": "sheet_15yrlump_F58",
            "15yrlump,G58": "sheet_15yrlump_G58",
            "15yrlump,I58": "sheet_15yrlump_I58",
            "15yrlump,J58": "sheet_15yrlump_J58",
            "15yrlump,K58": "sheet_15yrlump_K58",
            "15yrlump,L58": "sheet_15yrlump_L58",
            "15yrlump,M58": "sheet_15yrlump_M58",
            "15yrlump,N58": "sheet_15yrlump_N58",

            // Row 59
            "15yrlump,D59": "sheet_15yrlump_D59",
            "15yrlump,E59": "sheet_15yrlump_E59",
            "15yrlump,F59": "sheet_15yrlump_F59",
            "15yrlump,G59": "sheet_15yrlump_G59",
            "15yrlump,I59": "sheet_15yrlump_I59",
            "15yrlump,J59": "sheet_15yrlump_J59",
            "15yrlump,K59": "sheet_15yrlump_K59",
            "15yrlump,L59": "sheet_15yrlump_L59",
            "15yrlump,M59": "sheet_15yrlump_M59",
            "15yrlump,N59": "sheet_15yrlump_N59",

            // Row 60
            "15yrlump,D60": "sheet_15yrlump_D60",
            "15yrlump,E60": "sheet_15yrlump_E60",
            "15yrlump,F60": "sheet_15yrlump_F60",
            "15yrlump,G60": "sheet_15yrlump_G60",
            "15yrlump,I60": "sheet_15yrlump_I60",
            "15yrlump,J60": "sheet_15yrlump_J60",
            "15yrlump,K60": "sheet_15yrlump_K60",
            "15yrlump,L60": "sheet_15yrlump_L60",
            "15yrlump,M60": "sheet_15yrlump_M60",
            "15yrlump,N60": "sheet_15yrlump_N60",

            // Row 61
            "15yrlump,D61": "sheet_15yrlump_D61",
            "15yrlump,E61": "sheet_15yrlump_E61",
            "15yrlump,F61": "sheet_15yrlump_F61",
            "15yrlump,G61": "sheet_15yrlump_G61",
            "15yrlump,I61": "sheet_15yrlump_I61",
            "15yrlump,J61": "sheet_15yrlump_J61",
            "15yrlump,K61": "sheet_15yrlump_K61",
            "15yrlump,L61": "sheet_15yrlump_L61",
            "15yrlump,M61": "sheet_15yrlump_M61",
            "15yrlump,N61": "sheet_15yrlump_N61",

            // Row 62
            "15yrlump,D62": "sheet_15yrlump_D62",
            "15yrlump,E62": "sheet_15yrlump_E62",
            "15yrlump,F62": "sheet_15yrlump_F62",
            "15yrlump,G62": "sheet_15yrlump_G62",
            "15yrlump,I62": "sheet_15yrlump_I62",
            "15yrlump,J62": "sheet_15yrlump_J62",
            "15yrlump,K62": "sheet_15yrlump_K62",
            "15yrlump,L62": "sheet_15yrlump_L62",
            "15yrlump,M62": "sheet_15yrlump_M62",
            "15yrlump,N62": "sheet_15yrlump_N62",

            // Row 63
            "15yrlump,D63": "sheet_15yrlump_D63",
            "15yrlump,E63": "sheet_15yrlump_E63",
            "15yrlump,F63": "sheet_15yrlump_F63",
            "15yrlump,G63": "sheet_15yrlump_G63",
            "15yrlump,I63": "sheet_15yrlump_I63",
            "15yrlump,J63": "sheet_15yrlump_J63",
            "15yrlump,K63": "sheet_15yrlump_K63",
            "15yrlump,L63": "sheet_15yrlump_L63",
            "15yrlump,M63": "sheet_15yrlump_M63",
            "15yrlump,N63": "sheet_15yrlump_N63",

            // Row 64
            "15yrlump,D64": "sheet_15yrlump_D64",
            "15yrlump,E64": "sheet_15yrlump_E64",
            "15yrlump,F64": "sheet_15yrlump_F64",
            "15yrlump,G64": "sheet_15yrlump_G64",
            "15yrlump,I64": "sheet_15yrlump_I64",
            "15yrlump,J64": "sheet_15yrlump_J64",
            "15yrlump,K64": "sheet_15yrlump_K64",
            "15yrlump,L64": "sheet_15yrlump_L64",
            "15yrlump,M64": "sheet_15yrlump_M64",
            "15yrlump,N64": "sheet_15yrlump_N64",

            // Row 65
            "15yrlump,D65": "sheet_15yrlump_D65",
            "15yrlump,E65": "sheet_15yrlump_E65",
            "15yrlump,F65": "sheet_15yrlump_F65",
            "15yrlump,G65": "sheet_15yrlump_G65",
            "15yrlump,I65": "sheet_15yrlump_I65",
            "15yrlump,J65": "sheet_15yrlump_J65",
            "15yrlump,K65": "sheet_15yrlump_K65",
            "15yrlump,L65": "sheet_15yrlump_L65",
            "15yrlump,M65": "sheet_15yrlump_M65",
            "15yrlump,N65": "sheet_15yrlump_N65",

            // Row 66
            "15yrlump,D66": "sheet_15yrlump_D66",
            "15yrlump,E66": "sheet_15yrlump_E66",
            "15yrlump,F66": "sheet_15yrlump_F66",
            "15yrlump,G66": "sheet_15yrlump_G66",
            "15yrlump,I66": "sheet_15yrlump_I66",
            "15yrlump,J66": "sheet_15yrlump_J66",
            "15yrlump,K66": "sheet_15yrlump_K66",
            "15yrlump,L66": "sheet_15yrlump_L66",
            "15yrlump,M66": "sheet_15yrlump_M66",
            "15yrlump,N66": "sheet_15yrlump_N66",

            // Row 67
            "15yrlump,D67": "sheet_15yrlump_D67",
            "15yrlump,E67": "sheet_15yrlump_E67",
            "15yrlump,F67": "sheet_15yrlump_F67",
            "15yrlump,G67": "sheet_15yrlump_G67",
            "15yrlump,I67": "sheet_15yrlump_I67",
            "15yrlump,J67": "sheet_15yrlump_J67",
            "15yrlump,K67": "sheet_15yrlump_K67",
            "15yrlump,L67": "sheet_15yrlump_L67",
            "15yrlump,M67": "sheet_15yrlump_M67",
            "15yrlump,N67": "sheet_15yrlump_N67",

            // Row 68
            "15yrlump,D68": "sheet_15yrlump_D68",
            "15yrlump,E68": "sheet_15yrlump_E68",
            "15yrlump,F68": "sheet_15yrlump_F68",
            "15yrlump,G68": "sheet_15yrlump_G68",
            "15yrlump,I68": "sheet_15yrlump_I68",
            "15yrlump,J68": "sheet_15yrlump_J68",
            "15yrlump,K68": "sheet_15yrlump_K68",
            "15yrlump,L68": "sheet_15yrlump_L68",
            "15yrlump,M68": "sheet_15yrlump_M68",
            "15yrlump,N68": "sheet_15yrlump_N68",

            // Row 69
            "15yrlump,D69": "sheet_15yrlump_D69",
            "15yrlump,E69": "sheet_15yrlump_E69",
            "15yrlump,F69": "sheet_15yrlump_F69",
            "15yrlump,G69": "sheet_15yrlump_G69",
            "15yrlump,I69": "sheet_15yrlump_I69",
            "15yrlump,J69": "sheet_15yrlump_J69",
            "15yrlump,K69": "sheet_15yrlump_K69",
            "15yrlump,L69": "sheet_15yrlump_L69",
            "15yrlump,M69": "sheet_15yrlump_M69",
            "15yrlump,N69": "sheet_15yrlump_N69",

            // Row 70
            "15yrlump,D70": "sheet_15yrlump_D70",
            "15yrlump,E70": "sheet_15yrlump_E70",
            "15yrlump,F70": "sheet_15yrlump_F70",
            "15yrlump,G70": "sheet_15yrlump_G70",
            "15yrlump,I70": "sheet_15yrlump_I70",
            "15yrlump,J70": "sheet_15yrlump_J70",
            "15yrlump,K70": "sheet_15yrlump_K70",
            "15yrlump,L70": "sheet_15yrlump_L70",
            "15yrlump,M70": "sheet_15yrlump_M70",
            "15yrlump,N70": "sheet_15yrlump_N70",

            // Row 71
            "15yrlump,D71": "sheet_15yrlump_D71",
            "15yrlump,E71": "sheet_15yrlump_E71",
            "15yrlump,F71": "sheet_15yrlump_F71",
            "15yrlump,G71": "sheet_15yrlump_G71",
            "15yrlump,I71": "sheet_15yrlump_I71",
            "15yrlump,J71": "sheet_15yrlump_J71",
            "15yrlump,K71": "sheet_15yrlump_K71",
            "15yrlump,L71": "sheet_15yrlump_L71",
            "15yrlump,M71": "sheet_15yrlump_M71",
            "15yrlump,N71": "sheet_15yrlump_N71",

            // Row 72
            "15yrlump,D72": "sheet_15yrlump_D72",
            "15yrlump,E72": "sheet_15yrlump_E72",
            "15yrlump,F72": "sheet_15yrlump_F72",
            "15yrlump,G72": "sheet_15yrlump_G72",
            "15yrlump,I72": "sheet_15yrlump_I72",
            "15yrlump,J72": "sheet_15yrlump_J72",
            "15yrlump,K72": "sheet_15yrlump_K72",
            "15yrlump,L72": "sheet_15yrlump_L72",
            "15yrlump,M72": "sheet_15yrlump_M72",
            "15yrlump,N72": "sheet_15yrlump_N72",

            // Row 73
            "15yrlump,D73": "sheet_15yrlump_D73",
            "15yrlump,E73": "sheet_15yrlump_E73",
            "15yrlump,F73": "sheet_15yrlump_F73",
            "15yrlump,G73": "sheet_15yrlump_G73",
            "15yrlump,I73": "sheet_15yrlump_I73",
            "15yrlump,J73": "sheet_15yrlump_J73",
            "15yrlump,K73": "sheet_15yrlump_K73",
            "15yrlump,L73": "sheet_15yrlump_L73",
            "15yrlump,M73": "sheet_15yrlump_M73",
            "15yrlump,N73": "sheet_15yrlump_N73",

            // Row 74
            "15yrlump,D74": "sheet_15yrlump_D74",
            "15yrlump,E74": "sheet_15yrlump_E74",
            "15yrlump,F74": "sheet_15yrlump_F74",
            "15yrlump,G74": "sheet_15yrlump_G74",
            "15yrlump,I74": "sheet_15yrlump_I74",
            "15yrlump,J74": "sheet_15yrlump_J74",
            "15yrlump,K74": "sheet_15yrlump_K74",
            "15yrlump,L74": "sheet_15yrlump_L74",
            "15yrlump,M74": "sheet_15yrlump_M74",
            "15yrlump,N74": "sheet_15yrlump_N74",

            // Row 75
            "15yrlump,D75": "sheet_15yrlump_D75",
            "15yrlump,E75": "sheet_15yrlump_E75",
            "15yrlump,F75": "sheet_15yrlump_F75",
            "15yrlump,G75": "sheet_15yrlump_G75",
            "15yrlump,I75": "sheet_15yrlump_I75",
            "15yrlump,J75": "sheet_15yrlump_J75",
            "15yrlump,K75": "sheet_15yrlump_K75",
            "15yrlump,L75": "sheet_15yrlump_L75",
            "15yrlump,M75": "sheet_15yrlump_M75",
            "15yrlump,N75": "sheet_15yrlump_N75",

            // Row 76
            "15yrlump,D76": "sheet_15yrlump_D76",
            "15yrlump,E76": "sheet_15yrlump_E76",
            "15yrlump,F76": "sheet_15yrlump_F76",
            "15yrlump,G76": "sheet_15yrlump_G76",
            "15yrlump,I76": "sheet_15yrlump_I76",
            "15yrlump,J76": "sheet_15yrlump_J76",
            "15yrlump,K76": "sheet_15yrlump_K76",
            "15yrlump,L76": "sheet_15yrlump_L76",
            "15yrlump,M76": "sheet_15yrlump_M76",
            "15yrlump,N76": "sheet_15yrlump_N76",

            // Row 77
            "15yrlump,D77": "sheet_15yrlump_D77",
            "15yrlump,E77": "sheet_15yrlump_E77",
            "15yrlump,F77": "sheet_15yrlump_F77",
            "15yrlump,G77": "sheet_15yrlump_G77",
            "15yrlump,I77": "sheet_15yrlump_I77",
            "15yrlump,J77": "sheet_15yrlump_J77",
            "15yrlump,K77": "sheet_15yrlump_K77",
            "15yrlump,L77": "sheet_15yrlump_L77",
            "15yrlump,M77": "sheet_15yrlump_M77",
            "15yrlump,N77": "sheet_15yrlump_N77",

            // Row 78
            "15yrlump,D78": "sheet_15yrlump_D78",
            "15yrlump,E78": "sheet_15yrlump_E78",
            "15yrlump,F78": "sheet_15yrlump_F78",
            "15yrlump,G78": "sheet_15yrlump_G78",
            "15yrlump,I78": "sheet_15yrlump_I78",
            "15yrlump,J78": "sheet_15yrlump_J78",
            "15yrlump,K78": "sheet_15yrlump_K78",
            "15yrlump,L78": "sheet_15yrlump_L78",
            "15yrlump,M78": "sheet_15yrlump_M78",
            "15yrlump,N78": "sheet_15yrlump_N78",

            // Row 79
            "15yrlump,D79": "sheet_15yrlump_D79",
            "15yrlump,E79": "sheet_15yrlump_E79",
            "15yrlump,F79": "sheet_15yrlump_F79",
            "15yrlump,G79": "sheet_15yrlump_G79",
            "15yrlump,I79": "sheet_15yrlump_I79",
            "15yrlump,J79": "sheet_15yrlump_J79",
            "15yrlump,K79": "sheet_15yrlump_K79",
            "15yrlump,L79": "sheet_15yrlump_L79",
            "15yrlump,M79": "sheet_15yrlump_M79",
            "15yrlump,N79": "sheet_15yrlump_N79",

            // Additional cells: F84, F85, F86, J85, J86
            "15yrlump,F84": "sheet_15yrlump_F84",
            "15yrlump,F85": "sheet_15yrlump_F85",
            "15yrlump,F86": "sheet_15yrlump_F86",
            "15yrlump,J85": "sheet_15yrlump_J85",
            "15yrlump,J86": "sheet_15yrlump_J86",

            // param15yrlump!F43..F68
            "param15yrlump,F43": "sheet_param15yrlump_F43",
            "param15yrlump,F44": "sheet_param15yrlump_F44",
            "param15yrlump,F45": "sheet_param15yrlump_F45",
            "param15yrlump,F46": "sheet_param15yrlump_F46",
            "param15yrlump,F47": "sheet_param15yrlump_F47",
            "param15yrlump,F48": "sheet_param15yrlump_F48",
            "param15yrlump,F49": "sheet_param15yrlump_F49",
            "param15yrlump,F50": "sheet_param15yrlump_F50",
            "param15yrlump,F51": "sheet_param15yrlump_F51",
            "param15yrlump,F52": "sheet_param15yrlump_F52",
            "param15yrlump,F53": "sheet_param15yrlump_F53",
            "param15yrlump,F54": "sheet_param15yrlump_F54",
            "param15yrlump,F55": "sheet_param15yrlump_F55",
            "param15yrlump,F56": "sheet_param15yrlump_F56",
            "param15yrlump,F57": "sheet_param15yrlump_F57",
            "param15yrlump,F58": "sheet_param15yrlump_F58",
            "param15yrlump,F59": "sheet_param15yrlump_F59",
            "param15yrlump,F60": "sheet_param15yrlump_F60",
            "param15yrlump,F61": "sheet_param15yrlump_F61",
            "param15yrlump,F62": "sheet_param15yrlump_F62",
            "param15yrlump,F63": "sheet_param15yrlump_F63",
            "param15yrlump,F64": "sheet_param15yrlump_F64",
            "param15yrlump,F65": "sheet_param15yrlump_F65",
            "param15yrlump,F66": "sheet_param15yrlump_F66",
            "param15yrlump,F67": "sheet_param15yrlump_F67",
            "param15yrlump,F68": "sheet_param15yrlump_F68"
        };
    }


    exportKeysToFile(filename) {
        // Get all keys from the map
        const keys = Array.from(this._staticValues.keys());

        // Join keys with newline character
        const keyString = keys.join('\n');

        // Write to the specified file
        fs.writeFile(filename, keyString, (err) => {
            if (err) {
                console.error('Error writing to file:', err);
            } else {
                console.log('Keys written to', filename);
            }
        });
    }

}

module.exports = { PdZppHJFxW4lidr7 };