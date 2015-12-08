var chai = require('chai');
var expect = chai.expect;
chai.should();

var fs = require('fs');
var path = require('path');

var SharedStrings = require('../js/SharedStrings');
var Sheet = require('../js/Sheet');
var XMLFile = require('../js/XMLFile');

var mockWorkbook = {
    getXML: function () {
    },
    tempDir: path.join(__dirname, 'input-files')
};

describe('Reload an already loaded SharedStrings file', function () {
    var ss;
    it('Should not reload from file again', function (done) {
        ss = new SharedStrings(path.join(__dirname, 'input-files', 'sharedStrings.xml'), mockWorkbook);
        var firstXml;
        ss.load().then(function (xml) {
            firstXml = xml;
            return ss.load();
        }).then(function (xml) {
            expect(firstXml).to.equal(xml);
            done();
        });
    });

    it('Should return undefined if looking for a non-existing string', function () {
        expect(ss.getString(42)).to.be.undefined;
    });
});

describe('Reload an already loaded Sheet file', function () {
    var sheet;
    it('Should not reload from file again', function (done) {
        sheet = new Sheet(mockWorkbook, {}, {
            $: {
                Target: 'sheet1.xml',
                Id: 'rId1'
            }
        });
        var firstXml;
        sheet.load().then(function (xml) {
            firstXml = xml;
            return sheet.load();
        }).then(function (xml) {
            expect(firstXml).to.equal(xml);
            done();
        });
    });

});

describe('Reload an already loaded XML file', function () {
    var xmlfile;
    it('Should not reload from file again', function (done) {
        xmlfile = new XMLFile(path.join(__dirname, 'input-files', 'sharedStrings.xml'));
        var firstXml;
        xmlfile.load().then(function (xml) {
            firstXml = xml;
            return xmlfile.load();
        }).then(function (xml) {
            expect(firstXml).to.equal(xml);
            done();
        });
    });

});