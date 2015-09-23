var chai = require('chai');
var fs = require('fs');
var path = require('path');
var should = chai.should();
var kexcel = require('..');
var expect = chai.expect;
describe('Return values instead of functions', function () {
    var workbook;
    it('Open a input file...', function (done) {
        kexcel.open(path.join(__dirname, 'input-files', '42.xlsx'), function (err, wb) {
            workbook = wb;
            done();
        });
    });

    it('Get the value', function () {
        workbook.getSheet(0).getCellValue(1,1).should.equal('42');
    });

    it('Trailing white space test', function () {
        workbook.getSheet(0).getCellValue(2,1).should.equal('test with trailing white space ');
    });

    it('Null test', function () {
        expect( workbook.getSheet(0).getCellValue(3,1)).to.be.null;
    });

});
