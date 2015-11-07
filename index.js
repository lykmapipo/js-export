'use strict';

//dependencies
var _ = require('lodash');


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

    //bind options
    this.options = options;

    //bind data
    this.data = data;

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
            return engine[method].apply(this, arguments);
        };

    }.bind(this));
};


//export xlsform template
module.exports = exports = Export;