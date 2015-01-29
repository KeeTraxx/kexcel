var chai = require('chai');
var fs = require('fs');
var should = chai.should();
var kexcel = require('..');

describe('Basic kexcel sheet test', function () {
    var workbook;
    it('Create a new sheet', function (done) {
        kexcel.new(function (err, wb) {
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
        workbook.getSheet(0).getCellValue(1,1).should.contain('Hello World');
        workbook.getSheet(0).getCellValue(2,7).should.contain('equals');
        var row2 = workbook.getSheet(0).getRowValues(2);
        row2.should.be.a.array;
        row2.length.should.equal(9);
        row2.should.contain('Hello');
        row2.should.contain('< This is a formula');
        row2.should.contain('=D2+F2');
        row2.should.contain(19);
    });

    it('Save sheet to output.xlsx', function (done) {
        var ws = fs.createWriteStream('output.xlsx');
        ws.on('close', function (err) {
            should.not.exist(err);
            fs.existsSync('output.xlsx').should.be.true;
            done();
        });
        workbook.pipe(ws);
    });

});
