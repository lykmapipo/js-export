'use strict';

//dependencies
var path = require('path');
var expect = require('chai').expect;
var faker = require('faker');

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

    it('should be able to export data into excel file', function(done) {
        var data = [];
        for (var i = 0; i < 10; i++) {
            var user = faker.helpers.userCard();
            if ((i % 2) > 0) {
                delete user.email;
            }
            data.push(user);
        }

        var jsexport = new JSExport(data);

        jsexport.writeExcel(done);
    });

});