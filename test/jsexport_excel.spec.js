'use strict';

//dependencies
var fs = require('fs');
var path = require('path');
var del = require('del');
var mkdir = require('mkdir-p');
var chai = require('chai');
chai.use(require('chai-fs'));
var expect = chai.expect;
var generatedPath = path.join(__dirname, 'fixtures', 'generated');


//jsexport
var JSExport = require(path.join(__dirname, '..'));

describe('excel engine', function() {

    beforeEach(function() {
        mkdir.sync(generatedPath);
    });

    afterEach(function() {
        //cleanup file
        del.sync([generatedPath + '**']);
    });

    it('should be enable by default', function() {
        var jsexport = new JSExport();

        expect(jsexport.writeExcel).to.exist;
        expect(jsexport.downloadExcel).to.exist;
        expect(jsexport.writeExcel).to.be.a('function');
        expect(jsexport.downloadExcel).to.be.a('function');

    });

    it('should be able to flat data, join field name and export into excel file', function(done) {
        var filename =
            path.join(generatedPath, Date.now() + '.xls');

        var simple = path.join(__dirname, 'fixtures', 'simple.xls');

        var data =
            JSON.parse(fs.readFileSync(path.join(__dirname, 'fixtures', 'simple.json'), 'utf-8'));

        var jsexport = new JSExport(data);

        jsexport.writeExcel(filename, function(error) {

            expect(error).to.not.exist;
            expect(filename).to.be.a.file().and.not.empty;

            expect(fs.readFileSync(filename).toString('base64').substr(0, 12))
                .to.be.equal(fs.readFileSync(simple).toString('base64').substr(0, 12));

            done();
        });

    });

    it('should be able to flat data, unjoin field name and export into excel file', function(done) {
        var filename =
            path.join(generatedPath, Date.now() + '.xls');

        var farmer = path.join(__dirname, 'fixtures', 'farmer.xls');

        var data =
            JSON.parse(fs.readFileSync(path.join(__dirname, 'fixtures', 'farmer.json'), 'utf-8'));

        var jsexport = new JSExport(data, {
            joinFieldName: false,
            sheet: 'Farmer'
        });

        jsexport.writeExcel(filename, function(error) {

            expect(error).to.not.exist;
            expect(filename).to.be.a.file().and.not.empty;

            expect(fs.readFileSync(filename).toString('base64').substr(0, 12))
                .to.be.equal(fs.readFileSync(farmer).toString('base64').substr(0, 12));

            done();
        });

    });

});