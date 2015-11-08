'use strict';

//dependencies
var fs = require('fs');
var _ = require('lodash');
var xlsx = new require('xlsx-tools');

var excel = {};

function before() {
    //if flat sheet disable multi sheets
    if (this.options.flat) {
        this.options.multi = false;
    }
    //if multi sheets disable flat sheet
    if (this.options.multi) {
        this, options.flat = false
    }
}


function prepareHeaders() {
    //this refer to js-export instance
    var headers = [];

    //collect object headers
    _.forEach(this.data, function(data) {
        //get keys
        headers = _.union(headers, _.keys(data));
    });

    //compact headers
    headers = _.compact(headers);

    return headers;
}

function prepareRowsData(headers) {
    //this refer to js-export instance
    var rowsData = [];

    _.forEach(this.data, function(data) {
        var _data = {};

        _.forEach(headers, function(head) {
            _data[head] = data[head] || this.options.missing;
        }.bind(this));

        rowsData = _.union(rowsData, [_data]);

    }.bind(this));

    //compact data
    rowsData = _.compact(rowsData);

    return rowsData;
}


function prepareFlatSheet() {
    //this refer to js-export instance

    var sheet = {};

    //set sheet name
    sheet.name = this.options.sheet;

    sheet.columns = [];
    sheet.rows = [];

    //prepare sheet columns
    function prepareColumns(data, key, ref) {

        _.forEach(data, function(value, _key) {
            var _ref = undefined;

            if (_.isString(_key)) {
                //normalize ref and key	
                _ref = ref && _.isString(ref) ? [ref, _key].join('.') : _key;
                _key = key && _.isString(ref) ? [key, _key].join('_') : _key;
            }

            //check for column existance
            var exists = _.find(sheet.columns, {
                name: _key
            });

            //add column if not exists
            if (!exists) {

                //flatten plain object
                if (_.isPlainObject(value)) {
                    prepareColumns(value, _key, _ref);
                }

                //add simple columns
                else {
                    var column = {
                        name: _key,
                        ref: _ref
                    };
                    sheet.columns.push(column);
                }
            }

        });
    }

    prepareColumns(this.data, null, null);

    //prepare sheet rows
    _.forEach(this.data, function(data) {
        var _data = {};

        _.forEach(sheet.columns, function(column) {
            _data[column.name] = _.get(data, column.ref, this.options.missing);
        }.bind(this));

        sheet.rows = _.union(sheet.rows, [_data]);

    }.bind(this));

    return sheet;
}


function prepareSheets() {
    //this refer to js-export instance

    before.call(this);

    var sheets = [];

    //prepare single flat sheet
    if (this.options.flat) {
        var sheet = prepareFlatSheet.call(this);
        sheets.push(sheet);
    }

    //prepare multi sheets
    if (this.options.multi) {

    }


    return sheets;
}


function prepareWorkbook() {
    //this refer to js-export instance

    var sheets = prepareSheets.call(this);

    //prepare template workbook
    var workbook = xlsx.workbook();

    //create worksheets
    _.forEach(sheets, function(sheet) {

        //create worksheet
        var worksheet = xlsx.worksheet(sheet.name);

        //add workbook sheets
        workbook.addSheet(worksheet);

        //build worksheet header
        worksheet.setHeader(_.map(sheet.columns, 'name'));

        //build worksheet data
        _.forEach(sheet.rows, function(row) {
            worksheet.addRow(row);
        });

    });

    return workbook;
}



excel.write = function(filename, options, done) {
    //normalize arguments
    if (filename && _.isFunction(filename)) {
        done = filename;
        filename = Date.now() + '.xls';
    }

    if (filename && !_.isFunction(filename) && options && _.isFunction(options)) {
        done = options;
        options = {};
    }

    //prepare workbook
    var workbook = prepareWorkbook.call(this);

    //prepare file data
    var data = new Buffer(workbook.write(), 'binary');

    //write data to a file
    fs.writeFile(filename, data, options, done);
};


excel.download = function(response, options) {
    // body...
};


module.exports = excel;