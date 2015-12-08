var chai = require('chai');
var expect = chai.expect;
chai.should();

var fs = require('fs');
var path = require('path');

var kexcel = require('..');

var devnull = require('dev-null');

describe('Modify an existing .xlsx sheet', function () {
    var workbook;

    before(function (done) {
        kexcel.open(path.join(__dirname, 'input-files', 'example.xlsx')).then(function (wb) {
            workbook = wb;
            done();
        });
    });

    it('should have Hello World in it.', function () {
        workbook.getSheet(0).getCellValue(1, 1).should.equal('Hello');
        workbook.getSheet(0).getCellValue(1, 2).should.equal('World');
    });

    it('should return a computed / calculated value', function () {
        workbook.getSheet(0).getCellValue('D1').should.equal('42');
    });

    it('should return the cell function', function () {
        var bla = workbook.getSheet(0).getCellFunction('D1');
        workbook.getSheet(0).getCellFunction('D1').should.equal('=18+24');
    });

    it('should return undefined', function () {
        expect(workbook.getSheet(0).getCellFunction('D24')).to.be.undefined;
    });

    it('should be able to handle trailing white spaces', function () {
        workbook.getSheet(0).getCellValue('D2').should.equal('Trailing white spaces    ');
    });

    it('should be able to handle libre office text formats', function () {
        workbook.getSheet(0).getCellValue('D3').should.equal('Formatted Text');
    });

    after(function (done) {
        var stream = workbook.pipe(devnull());
        stream.on('finish', done);
    });

});

describe('Open an existing .xlsx stream', function () {
    var workbook;

    before(function (done) {
        kexcel.open(fs.createReadStream(path.join(__dirname, 'input-files', 'example.xlsx'))).then(function (wb) {
            workbook = wb;
            done();
        });
    });

    it('should have Hello World in it.', function () {
        workbook.getSheet(0).getCellValue(1, 1).should.equal('Hello');
        workbook.getSheet(0).getCellValue(1, 2).should.equal('World');
    });

    after(function (done) {
        var stream = workbook.pipe(devnull());
        stream.on('finish', done);
    });

});

describe('Copy cell style', function () {
    var workbook;

    before(function (done) {
        kexcel.open(path.join(__dirname, 'input-files', 'example.xlsx')).then(function (wb) {
            workbook = wb;
            done();
        });
    });

    it('should be able to copy cell style from another cell', function () {
        workbook.getSheet(0).setCellValue('A2', 'Test', 'A1');
    });

    after(function (done) {
        var stream = workbook.pipe(fs.createWriteStream(path.join(__dirname, 'output-files', 'advanced.xlsx')));
        stream.on('finish', done);
    });

});

describe('Open a non-existing file', function () {
    it('should throw an error', function (done) {
        kexcel.open('fileNotFound.xlsx').catch(function (err) {
            done();
        });
    });

});

describe('Open a corrupted file', function () {
    it('should throw an error', function (done) {
        kexcel.open('package.json').catch(function (err) {
            done();
        });
    });
});