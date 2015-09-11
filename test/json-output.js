var chai = require('chai');
var fs = require('fs');
var path = require('path');
var should = chai.should();
var kexcel = require('..');


describe('JSON Output', function () {
    var workbook;
    it('JSON Output', function (done) {
        kexcel.open(path.join(__dirname, 'input-files', 'beer_per_capita.xlsx'), function (err, wb) {
            var sheet = wb.getSheet(0);
            var json = sheet.toJSON();

            json[0].name.should.equal('Czech Republic');
            json[0].beer_per_capita.should.equal('148.6');

            json[49].name.should.equal('India');
            json[49].beer_per_capita.should.equal('2');
            done();
        });
    });

});
