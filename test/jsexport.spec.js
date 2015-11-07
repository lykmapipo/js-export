'use strict';

//dependencies
var path = require('path');
var expect = require('chai').expect;

//jsexport
var JSExport = require(path.join(__dirname, '..'));

describe('js export', function() {

    it('should export functional constructor', function() {
        expect(JSExport).to.be.a('function');
    });

    it('should be able to plugin new export engine', function() {
        var simpleEngine = {
            write: function( /*path, done*/ ) {

            },
            download: function( /*response, options*/ ) {

            }
        };

        var jsexport = new JSExport();
        jsexport.use('simple', simpleEngine);

        expect(jsexport.writeSimple).to.exist;
        expect(jsexport.downloadSimple).to.exist;
        expect(jsexport.writeSimple).to.be.a('function');
        expect(jsexport.downloadSimple).to.be.a('function');
    });

});