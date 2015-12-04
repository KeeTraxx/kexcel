var chai = require('chai');
var expect = chai.expect;
chai.should();

var fs = require('fs');
var path = require('path');

var kexcel = require('..');

var devnull = require('dev-null');

describe('KExcel new instantiation', function () {
	var workbook;

	before(function (done) {
		kexcel.new().then(function (wb) {
			workbook = wb;
			done();
		});
	});



	it('should instanciate a new workbook', function () {
		expect(workbook).to.exist;
		expect(workbook).to.be.an.instanceof(kexcel);
	});

	it('should return an empty workbook (with 1 sheet)', function () {
		workbook.getSheet(0).should.exist;
		expect(workbook.getSheet(1)).to.not.exist;
	});

	it('should have a sheet named "Sheet1"', function () {
		workbook.getSheet(0).should.exist;
		workbook.getSheet(0).getName().should.equal('Sheet1');
	});

	it('should be able to get sheet by name', function () {
		var sheet = workbook.getSheet('Sheet1');
		sheet.getName().should.equal('Sheet1');

		expect(workbook.getSheet('FooBar')).to.be.undefined;
	});

	after(function (done) {
		var stream = workbook.pipe(devnull());
		stream.on('finish', done);
	});

});

describe('KExcel sheet modification', function () {
	var workbook;

	before(function (done) {
		kexcel.new().then(function (wb) {
			workbook = wb;
			done();
		});
	});

	it('should be able to change the name of the first sheet', function () {
		workbook.getSheet(0).setName('MySheet');
		workbook.getSheet(0).getName().should.equal('MySheet');
	});

	it('should be able to set a value in a cell', function () {
		var sheet = workbook.getSheet(0);
		sheet.setCellValue(1, 1, 'Hello world!');
		sheet.getCellValue(1, 1).should.equal('Hello world!');
		sheet.getCellValue(1, 1).should.have.length(12);
	});

	it('should be able to set a number in a cell', function () {
		var sheet = workbook.getSheet(0);
		sheet.setCellValue(3, 1, 42);
		sheet.getCellValue(3, 1).should.equal(42);
	});

	it('should be able to set a formula in a cell', function () {
		var sheet = workbook.getSheet(0);
		sheet.setCellValue(2, 1, '=A1');
		sheet.getCellValue(2, 1).should.equal('=A1');
	});

	it('should be able to remove the values of a cell', function () {
		var sheet = workbook.getSheet(0);
		sheet.setCellValue(4, 1, 'To be removed');
		sheet.getCellValue(4, 1).should.equal('To be removed');
		sheet.setCellValue(4, 1, undefined);
		expect(sheet.getCellValue(4, 1)).to.be.undefined;
	});

	it('should be able to get a cell value by reference (e.g. "A1")', function () {
		var sheet = workbook.getSheet(0);
		sheet.getCellValue('A1').should.equal('Hello world!');
		expect(sheet.getCellValue('Z55')).to.be.undefined;
	});

	it('should be able to create a new sheet without a name', function () {
		var sheet = workbook.createSheet();
		sheet.getName().should.equal('Sheet2');
	});

	it('should be able to create a new sheet with a name', function () {
		var sheet = workbook.createSheet('FooSheet');
		sheet.getName().should.equal('FooSheet');
	});

	it('should be able to copy contents from another sheet', function () {
		var sheet = workbook.getSheet('Sheet2');
		sheet.copyFrom(workbook.getSheet('MySheet'));
		sheet.getCellValue('A1').should.equal('Hello world!');
	});

	it('should deep copy contents', function () {
		var mysheet = workbook.getSheet('MySheet');
		var sheet2 = workbook.getSheet('Sheet2');
		sheet2.setCellValue(1, 2, 'Thanks for all the fish');
		sheet2.getCellValue(1, 2).should.not.equal(mysheet.getCellValue(1, 2));
	});

	it('should be able to set row values', function () {
		var sheet = workbook.getSheet('FooSheet');
		sheet.setRow(1, [1, 2, '=A1+B1']);
	});

	it('should be able to get row values', function () {
		var sheet = workbook.getSheet('FooSheet');
		sheet.getRow(1).should.be.eql([1, 2, '=A1+B1']);
	});

	it('should be able to append row values', function () {
		var sheet = workbook.getSheet('FooSheet');
		sheet.appendRow(['This', 'is', 'an', 'appended', 'row']);
		sheet.getRow(2).should.be.eql(['This', 'is', 'an', 'appended', 'row']);
	});

	after(function (done) {
		var stream = workbook.pipe(devnull());
		stream.on('finish', done);
	});

});

describe('KExcel output possibilites', function () {
	var workbook;

	before(function (done) {
		kexcel.new().then(function (wb) {
			workbook = wb;
			done();
		});
	});

	it('should be able to output JSON', function () {
		var sheet = workbook.getSheet(0);
		sheet.appendRow(['firstname', 'lastname']);
		sheet.appendRow(['Spike', 'Spiegel']);
		sheet.appendRow(['Gandalf', 'The Grey']);

		sheet.toJSON().should.eql([
			{ firstname: 'Spike', lastname: 'Spiegel' },
			{ firstname: 'Gandalf', lastname: 'The Grey' }
		])
	});
	it('should be able to save to an .xlsx file', function (done) {
		var filename = path.join(__dirname, 'output-files', 'basic.xlsx');
		var stream = workbook.pipe(fs.createWriteStream(filename));
		stream.on('finish', function () {
			fs.exists(filename, function (exists) {
				if (exists) {
					done();
				} else {
					done(new Error('Outputfile does not exist: ' + path));
				}
			});
		});
	});
});