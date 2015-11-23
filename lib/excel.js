'use strict';

//dependencies
var fs = require('fs');
var _ = require('lodash');
var xlsx = new require('xlsx-tools');

var excel = {};

function isPrimitive(value) {
    return _.isBoolean(value) || _.isDate(value) ||
        _.isNumber(value) || _.isString(value) ||
        value.toString() !== '[object Object]';
}

function before() {
    //if multi sheets disable flat sheet
    if (this.options.multi) {
        this.options.flat = false
    }

    //if flat sheet disable multi sheets
    if (this.options.flat) {
        this.options.multi = false;
    }

}


function prepareFlatSheet() {
    //this refer to js-export instance

    /* jshint validthis:true */
    var self = this;

    var sheet = {};

    //set sheet name
    sheet.name = this.options.sheet;

    sheet.columns = [];
    sheet.rows = [];

    //prepare sheet columns
    function prepareColumns(data, key, ref) {

        _.forEach(data, function(value, _key) {
            var _ref = _key;

            if (_.isString(_key)) {
                //normalize ref and key
                if (ref && _.isString(ref)) {
                    _ref = [ref, _key].join('.');
                }

                if (key && _.isString(key) && self.options.joinFieldName) {
                    _key = [key, _key].join(self.options.fieldSeparator);
                }
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

function prepareMultiSheets() {
    //this refer to js-export instance

    /* jshint validthis:true */
    var self = this;

    var sheets = [];

    function prepareSheets(sheetName, data) {

        function prepareSheet(sheetName, data, ref) {

            _.forEach(data, function(value, key) {

                //handle primitives
                if (isPrimitive(value)) {

                    //find existing sheet
                    var sheet = _.find(sheets, {
                        name: sheetName
                    });

                    //add sheet
                    if (!sheet) {
                        sheet = {
                            name: sheetName,
                            columns: []
                        };
                        sheets.push(sheet);
                    }

                    //check for column existance
                    var column = _.find(sheet.columns, {
                        name: key
                    });

                    //add column
                    if (!column) {
                        column = {
                            name: key,
                            ref: _.compact([ref, key]).join('.')
                        }
                        sheet.columns.push(column);
                    }
                } else {
                    prepareSheet(key, value, _.compact([ref, key]).join('.'));
                }

            });

        }

        _.forEach(data, function(_data) {
            prepareSheet(sheetName, _data, null);
        });

    }

    prepareSheets(this.options.sheet, this.data);

    //prepare sheets rows
    _.forEach(this.data, function(data) {

        _.forEach(sheets, function(sheet) {
            var _data = {};

            _.forEach(sheet.columns, function(column) {
                _data[column.name] = _.get(data, column.ref, this.options.missing);
            }.bind(this));

            sheet.rows = _.union(sheet.rows, [_data]);

        }.bind(this));

    }.bind(this));

    return sheets;
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
        var _sheets = prepareMultiSheets.call(this);
        sheets = _.union(sheets, _sheets);
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
    //merge options with default
    options = _.merge({
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        filename: Date.now() + '.xls'
    }, options);

    //prepare workbook
    var workbook = prepareWorkbook.call(this);

    //prepare data
    var data = new Buffer(workbook.write(), 'binary');

    //set response type
    response.type(options.type);

    //set file name
    response.attachment(options.filename);

    //respond with excel data
    response.send(data);
};


module.exports = excel;