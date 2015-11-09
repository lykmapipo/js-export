'use strict';

//dependencies
var fs = require('fs');
var _ = require('lodash');
var xlsx = new require('xlsx-tools');

var excel = {};

function isPrimitive(value) {
    return _.isBoolean(value) || _.isDate(value) ||
        _.isNumber(value) || _.isString(value);
}

function havePrimitives(object) {
    return _.reduce(object, function(has, value) {
        return has || isPrimitive(value);
    }, false);
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

    //prepare sheets
    function prepareSheets(data, key) {

        _.forEach(data, function(value, _key) {
            //prepare sheet name
            var name = _key && _.isString(_key) ? _key : key;

            //create or reuse sheet
            var isValidSheet = havePrimitives(value);
            if (isValidSheet) {

                var sheet = _.find(sheets, {
                    name: name
                }) || {
                    name: name
                };

                sheet.columns = [];
                sheet.rows = [];

                //prepare sheet columns
                _.forEach(value, function(subvalue, subkey) {
                    if (isPrimitive(subvalue)) {
                        var column = {
                            name: subkey,
                            ref: subkey
                        };
                        sheet.columns.push(column);
                    }
                });

                //add sheet
                sheets.push(sheet);

            }

            //continue prepare sheets for sub values
            var subvalues = {};
            _.forEach(value, function(subvalue, subkey) {
                if (_.isPlainObject(subvalue)) {
                    if (self.options.joinSheetName) {
                        subkey = [name, subkey].join(self.options.fieldSeparator);
                    }
                    subvalues[subkey] = subvalue;
                }
            });
            prepareSheets(subvalues, name);

        });

    }

    prepareSheets(this.data, this.options.sheet);

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
    // body...
};


module.exports = excel;