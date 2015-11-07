'use strict';

//dependencies
var path = require('path');
var expect = require('chai').expect;

//jsexport
var JSExport = require(path.join(__dirname, '..'));

describe('excel engine', function() {

    it('should be able enable by default', function() {
        var jsexport = new JSExport();

        expect(jsexport.writeExcel).to.exist;
        expect(jsexport.downloadExcel).to.exist;
        expect(jsexport.writeExcel).to.be.a('function');
        expect(jsexport.downloadExcel).to.be.a('function');

    });

});