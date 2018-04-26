var GoogleSpreadsheet = require("google-spreadsheet");

function loadSpreadsheet(options, cb) {

    var my_sheet = new GoogleSpreadsheet(options.key);

    if (options.private_key_id) {
        options.type = 'service_account';
        my_sheet.useServiceAccountAuth({
            private_key_id: options.private_key_id,
            private_key: options.private_key,
            client_email: options.client_email,
            client_id: options.client_id
        }, function () { processSpreadsheet(my_sheet, options, cb) });
    } else {
        processSpreadsheet(my_sheet, options, cb)
    }

}


function processSpreadsheet(my_sheet, options, cb) {

    my_sheet.getInfo(function (err, info) {
        if (err) {
            return cb(err)
        }

        var sheet = info.worksheets[options.sheet - 1];
        var firstRow = options.firstRow || 1;
        var colCount = options.colCount || sheet.colCount
        var rowCount = sheet.rowCount

        my_sheet.getCells(options.sheet, {
            'min-row': firstRow,
            'max-row': rowCount,
            'min-col': 1,
            'max-col': colCount,
            'return-empty': true
        }, function (err, row_data) {
            if (err) {
                return cb(err)
            }
            var converted = {};
            var langs = [];
            var commentsColumnIndex;

            for (var i = 1; i < colCount; i++) {
                if (options.ignoreCommentsColumn == true && row_data[i].value == 'comments') {
                    // do nothing
                    commentsColumnIndex = i;
                } else {
                    if (options.warnOnMissingValues && row_data[i].value.length == 0) {
                        console.log('Column is missing key at index ' + i)
                    }
                    if (options.errorOnMissingValues && row_data[i].value.length == 0) {
                        throw new Error('Column is missing key at index ' + i);
                    }
                    langs[i] = row_data[i].value;
                }
            }

            for (var i = firstRow + 1; i <= rowCount; i++) {
                for (var j = 1; j < colCount; j++) {
                    if (options.ignoreCommentsColumn && j == commentsColumnIndex) {
                        // do nothing
                    } else {
                        if (options.warnOnMissingValues && row_data[(i - 1) * colCount + j].value.length == 0) {
                            console.log('Cell is missing value at col ' + i + ', row ' + j)
                        }
                        if (options.errorOnMissingValues && row_data[(i - 1) * colCount + j].value.length == 0) {
                            throw new Error('Cell is missing value at col ' + i + ', row ' + j);
                        }
                        var lang = langs[j];
                        converted[lang] = converted[lang] || {};
                        converted[lang][row_data[(i - 1) * colCount].value] = row_data[(i - 1) * colCount + j].value;
                    }
                }
            }
            return cb(null, converted);
        });
    });

}

function gulpI18n(options) {

    if (!options) {
        throw new Error('Missing options!')
    }
    
    return new Promise((resolve, reject) => {
        loadSpreadsheet(options, function (err, data) {
            if (err) {
                throw err;
            } else {
                resolve(data)
            }
            
        });
    })
    
}

module.exports = gulpI18n;
