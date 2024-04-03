/**
 * Converts a package to a configuration workbook.
 * This function reads data from the active spreadsheet, processes it based on certain conditions,
 * and creates or updates sheets in the spreadsheet accordingly.
 */
function convertPackageToConfigWorkbook() {
    var ss = SpreadsheetApp.getActive(); // Gets the active spreadsheet
    var range = ss.getDataRange(); // Gets the data range of the spreadsheet
    var values = range.getValues(); // Gets the values within the data range
    var resetTypeOpenFlag = true; // Flag to reset type open status
    var memberList = []; // List to store members
    var map1 = new Map(); // Map to store component names and their respective members

    // Iterate through each value in the spreadsheet
    values.forEach(result => {
        if (result == "<types>" && resetTypeOpenFlag) {
            resetTypeOpenFlag = false;
        } else if (result == "</types>" && !resetTypeOpenFlag) {
            resetTypeOpenFlag = true;
        } else if (result[0].includes('members') && !resetTypeOpenFlag) {
            let memberName = result[0].substring(9, (result[0].length - 10));
            memberList.push(memberName);
        } else if (result[0].includes('name') && !resetTypeOpenFlag) {
            let componentName = result[0].substring(6, (result[0].length - 7));
            map1.set(componentName, memberList);
            memberList = [];
        }
    });

    var startRow = 1;
    var startCol = 1;

    // Iterate through the entries of map1
    for (const [key, value] of map1.entries()) {
        var sheetValue = [];
        try {
            var rs = SpreadsheetApp.getActive().getSheetByName(key);
            sheetValue.push(['Component Name', 'Purpose']);
            value.forEach((val, index) => {
                sheetValue.push([val, '']);
            });
            if (!rs) {
                rs = ss.insertSheet(key);
            }

            rs.setColumnWidth(1, 400);
            rs.setColumnWidth(2, 600);
            var range = rs.getRange(1, 1, 1, 2);
            range.setBackground("blue");
            range.setFontColor("white");
            range.setFontWeight("bold");
            range.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

            var valueRange = rs.getRange(startRow, startCol, sheetValue.length, sheetValue[0].length);
            valueRange.setValues(sheetValue);
            valueRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

        } catch (error) {
            Logger.log('Key is : ' + error);
        }

    }
}
