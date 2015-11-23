'use strict';

//dependencies
var fs = require('fs');
var path = require('path');
var express = require('express');
var request = require('supertest');
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
        var filename = path.join(generatedPath, Date.now() + '.xls');

        var simple = path.join(__dirname, 'fixtures', 'simple_flat.xls');

        var data = require(path.join(__dirname, 'fixtures', 'simple.json'));

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
        var filename = path.join(generatedPath, Date.now() + '.xls');

        var farmer = path.join(__dirname, 'fixtures', 'farmer_flat.xls');

        var data = require(path.join(__dirname, 'fixtures', 'farmer.json'));

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


    it('should be able to export data into multi sheet excel file', function(done) {
        var filename = path.join(generatedPath, Date.now() + '.xls');

        var data = require(path.join(__dirname, 'fixtures', 'simple.json'));

        var multi = path.join(__dirname, 'fixtures', 'simple_multi.xls');

        var jsexport = new JSExport(data, {
            sheet: 'basic',
            multi: true
        });

        jsexport.writeExcel(filename, function(error) {

            expect(error).to.not.exist;
            expect(filename).to.be.a.file().and.not.empty;

            expect(fs.readFileSync(filename).toString('base64').substr(0, 12))
                .to.be.equal(fs.readFileSync(multi).toString('base64').substr(0, 12));

            done();
        });

    });

    it('should be able to export and download data in specified http path', function(done) {
        var filename = Date.now() + '.xls';
        var data = require(path.join(__dirname, 'fixtures', 'simple.json'));
        var simple = path.join(__dirname, 'fixtures', 'simple_flat.xls');

        var app = express();
        app.get('/xlsexport', function(request, response) {
            var jsexport = new JSExport(data);
            jsexport.downloadExcel(response, {
                filename: filename
            });
        });

        function binaryParser(response, callback) {
            response.setEncoding('binary');
            response.data = '';
            response.on('data', function(chunk) {
                response.data += chunk;
            });
            response.on('end', function() {
                callback(null, new Buffer(response.data, 'binary'));
            });
        }

        request(app)
            .get('/xlsexport')
            .expect(200)
            .parse(binaryParser)
            .end(function(error, response) {

                expect(error).to.not.exist;

                expect(response.headers['content-type'])
                    .to.equal('application/vnd.ms-excel');

                expect(response.headers['content-disposition']).to.exist;

                expect(response.headers['content-disposition'])
                    .to.be.equal('attachment; filename="' + filename + '"');

                expect(response.body.toString('base64').substr(0, 12))
                    .to.be.equal(fs.readFileSync(simple).toString('base64').substr(0, 12));

                done(error, response);

            });
    });

});