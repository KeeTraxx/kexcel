var chai = require('chai');
var fs = require('fs');
var should = chai.should();
var expect = chai.expect;

var kexcel = require('..');
var path = require('path');

describe('Basic kexcel sheet test', function () {
    var workbook;
    it('Create a new sheet', function (done) {
        kexcel.new(function (err, wb) {
            if (err) throw err;
            workbook = wb;
            done();
        });
    });

    it('Put "Hello World" in cell A1', function () {
        workbook.getSheet(0).setCellValue(1, 1, 'Hello World');
    });

    it('Put data into a row', function () {
        workbook.getSheet(0).setRowValues(2, ['Hello', 'World', 'and', 19, 'plus', 23, 'equals', '=D2+F2', '< This is a formula']);
    });

    it('Get data from a cell', function () {
        workbook.getSheet(0).getCellValue(1, 1).should.contain('Hello World');
        workbook.getSheet(0).getCellValue(2, 7).should.contain('equals');
        var row2 = workbook.getSheet(0).getRowValues(2);
        row2.should.be.a.array;
        row2.length.should.equal(9);
        row2.should.contain('Hello');
        row2.should.contain('< This is a formula');
        row2.should.contain('=D2+F2');
        row2.should.contain(19);
    });

    it('Duplicate a sheet', function () {
        var sheet = workbook.duplicateSheet(0, 'Duplicated Sheet');

        sheet.getCellValue(1, 1).should.equal('Hello World');
        sheet.setCellValue(2, 2, 'Another World');

        workbook.getSheet(0).getCellValue(2, 2).should.equal('World');
        sheet.getCellValue(2, 2).should.equal('Another World');

    });

    it('Get a cell by Excel reference', function () {
        workbook.getSheet(0).getCellValue('A1').should.equal('Hello World');
    });

    it('Set a cell value by Excel reference', function () {
        workbook.getSheet(1).setCellValueByRef('A1', 'This is Sparta!!');
        workbook.getSheet(1).getCellValue('A1').should.equal('This is Sparta!!');
    });

    it('Set using undefined values does nothing', function () {
        // TODO: Actually this should empty the cell?
        workbook.getSheet(0).setCellValue(1, 1, null);
    });

    it('Get an undefined cell', function () {
        expect(workbook.getSheet(0).getCellValue(200, 200)).to.be.undefined;
        expect(workbook.getSheet(0).getCellValue(1, 200)).to.be.undefined;
    });

    it('Get a undefined row', function () {
        expect(workbook.getSheet(0).getRowValues(200)).to.be.undefined;
    });

    it('Get the last row number', function () {
        workbook.getSheet(0).getLastRowNumber().should.equal(2);
    });

    it('Append a row', function () {
        workbook.getSheet(0).appendRow(['one', 'two']);
        workbook.getSheet(0).getLastRowNumber().should.equal(3);
        workbook.getSheet(0).getCellValue(3, 2).should.equal('two');
    });

    it('Add a row somewhere', function () {
        workbook.getSheet(0).setRowValues(42, ['row', 'forty', 'two']);
        workbook.getSheet(0).getLastRowNumber().should.equal(42);
        workbook.getSheet(0).getCellValue(42, 2).should.equal('forty');
    });

    it('Append a row', function () {
        workbook.getSheet(0).appendRow(['Forty', 'Three']);
        workbook.getSheet(0).getLastRowNumber().should.equal(43);
        workbook.getSheet(0).getCellValue(3, 2).should.equal('two');
    });

    it('Save sheet to output.xlsx', function (done) {
        var file = path.join(__dirname, 'output-files', 'output.xlsx');
        var ws = fs.createWriteStream(file);
        ws.on('close', function (err) {
            should.not.exist(err);
            fs.existsSync(file).should.be.true;
            done();
        });
        workbook.pipe(ws);
    });


});
