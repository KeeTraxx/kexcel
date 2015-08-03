var chai = require('chai');
var fs = require('fs');
var path = require('path');
var should = chai.should();
var kexcel = require('..');


describe('Issue #3 github', function () {
    var workbook;
    it('Opening the sheet with non-standard cell formattings', function (done) {
        kexcel.open(path.join(__dirname, 'input-files', 'issue_3.xlsx'), function (err, wb) {
            workbook = wb;
            done();
        });
    });

    it('Overwriting a cell', function () {
        var sheet1 = workbook.getSheet(0);
        sheet1.setCellValue(3, 1, 'Overwrite');
    });

    it('Save sheet to output.xlsx', function (done) {
        var ws = fs.createWriteStream(path.join(__dirname, 'output-files', 'issue_3.xlsx'));
        ws.on('close', function (err) {
            should.not.exist(err);
            fs.existsSync(path.join(__dirname, 'output-files', 'issue_3.xlsx')).should.be.true;
            done();
        });
        workbook.pipe(ws);
    });

});

describe('Callback issue', function () {
    var workbook;
    it('Create a new workbook and some data', function (done) {
        kexcel.new(function (err, wb) {
            workbook = wb;

            var sheet = workbook.getSheet(0);

            sheet.setCellValue(1, 1, 'test');

            done();
        });
    });

    it('Save sheet to output2.xlsx with callback', function (done) {
        var ws = fs.createWriteStream(path.join(__dirname, 'output-files', 'issue_3.xlsx'));

        workbook.pipe(ws, function (err, wb) {
            should.not.exist(err);
            should.exist(wb);
            var sheet = wb.getSheet(0);
            sheet.getCellValue(1, 1).should.equal('test');
            done();
        });
    });
});

describe('Complicated formulae', function () {
    var workbook;
    it('Create a new workbook and some data', function (done) {

        var randomNumbers = [];

        for (var row = 1; row < 11; row++) {
            randomNumbers.push(~~(Math.random() * 50))
        }

        kexcel.new(function (err, wb) {
            workbook = wb;
            var sheet = wb.getSheet(0);
            for (var row = 1; row < 11; row++) {
                sheet.setCellValue(row, 1, randomNumbers[row - 1]);
            }

            var row = 11;

            sheet.setCellValue(row, 2, '=SUM(A1:A10)');
            sheet.setCellValue(row, 1, 'Sum');
            row++;
            sheet.setCellValue(row, 2, '=AVERAGE(A1;A2)');
            sheet.setCellValue(row, 1, 'Average');
            row++;
            sheet.setCellValue(row, 2, '=CONCATENATE(A1;A2;A3;A4;A5;A6;A7;A8;A9)');
            sheet.setCellValue(row, 1, 'Concat');

            done();
        });

    });


    it('Save sheet to output2.xlsx with callback', function (done) {
        var ws = fs.createWriteStream(path.join(__dirname, 'output-files', 'complicated_formulas.xlsx'));
        workbook.pipe(ws, done);
    });

});