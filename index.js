'use strict';

//dependencies
var path = require('path');
var _ = require('lodash');

//load default engines
var excel = require(path.join(__dirname, 'lib', 'excel'));

//defaults options
var defaults = {
    sheet: 'Sheet',
    flat: true,
    joinFieldName: true,
    fieldSeparator: '_',
    multi: false,
    joinSheetName: false,
    missing: 'NA'
};


function uniqueFields(data) {
    //collect unique fields
    var fields = [];

    //collect data fields
    _.forEach(data, function(value, key) {
        //obtain fields from plain object data type
        if (_.isPlainObject(value)) {
            fields = _.union(fields, _.keys(value));
        }

        //obtain fields from simple data types
        if (!_.isPlainObject(value) && _.isString(key)) {
            fields = _.union(fields, [key]);
        }
    });

    //compact fields
    fields = _.compact(fields);

    return fields;
}


function normalizeFields(data, fields, missing) {
    //normalize data fields
    _.forEach(fields, function(field) {
        if (!_.has(data, field)) {
            //set missing value
            data[field] = missing;
        }
    });

    return data;

}

function normalize() {
    //this refer to js-export instance

    /* jshint validthis:true */
    var self = this;

    var data = [];

    var fields = uniqueFields(this.data);

    _.forEach(this.data, function(value) {
        data.push(normalizeFields(value, fields, self.options.missing));
    });

    //set data to normalized data
    this.data = data;
}

/**
 * @constructor
 * @description export js data into different acceptable format  
 * @param {Object} options export options
 * @return {Object} export instance
 * @public
 */
function Export(data, options) {

    //normalize options
    options = options || {};

    //merge defaults
    options = _.merge({}, defaults, options);

    //bind options
    this.options = options;

    //bind data
    this.data = data;
    if (this.data && !_.isArray(this.data)) {
        this.data = [data];
    }


    //normalize data structure
    normalize.call(this);

    //set default engines
    this.use('excel', excel);

}


/**
 * @function
 * @description plugin new export engine
 * @param  {String} name     name of the export engine
 * @param  {Object} exporter valid export engine definition
 * @public
 */
Export.prototype.use = function(name, engine) {
    //extend export with additional exporter engine
    var methods = ['write', 'download'];

    _.forEach(methods, function(method) {
        //prepare method name
        var _method = method + _.capitalize(name);

        //bind export engine functions into export
        this[_method] = function() {

            //invoke export engine method
            return engine[method].apply(this, arguments);

        };

    }.bind(this));
};


//export xlsform template
module.exports = exports = Export;