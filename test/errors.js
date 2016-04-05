var chai = require('chai');
var assert = chai.assert;
var expect = chai.expect;
chai.should();

var fs = require('fs');
var path = require('path');

var kexcel = require('..');

var devnull = require('dev-null');
var util = require('util');
describe('KExcel open a non-existing file', function () {
    it('should should throw an error', function () {
	kexcel.open('non-existing-file.xlsx').then(function(wb){
	    console.log('should never get here');
	    assert(false);
	}).catch(function(err){
	    //console.log(util.inspect(err));
	    expect(err).to.be.defined;
	    expect(err).to.equal('non-existing-file.xlsx not found.');
	});
    });
});