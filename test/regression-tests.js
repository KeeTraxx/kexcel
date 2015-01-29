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
        sheet1.setCellValue(3,1, 'Overwrite');
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