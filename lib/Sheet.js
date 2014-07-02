// Core node modules
var path = require('path');

// NPM Packages
var elementtree = require('elementtree');
var subElement = elementtree.SubElement;
var _ = require('underscore');

var XMLFile = require('./XMLFile');

module.exports = Sheet;

function Sheet(workbook, attribs) {
    var sheet = this;
    _.forEach(attribs, function (value, key) {
        sheet[key] = value;
    });
    this.sharedstrings = workbook.sharedStringsXml;
    this.xml = new XMLFile(path.join(workbook.temppath, this.path), this.sheetXML);
    var sheetView = this.xml.find('./sheetViews/sheetView');
    sheetView.set('tabSelected', null);
}

Sheet.prototype.setCellValue = function (rownum, col, cellvalue, cellstyle) {
    var sheetData = this.xml.find('./sheetData');
    var row = _.find(sheetData.findall('./row'), function (d) {
        return d.get('r') == '' + rownum;
    });


    var ref = intToExcelColumn(col) + rownum;

    if (!row) {
        row = subElement(sheetData, 'row');
        row.set('r', rownum);
    }

    var c = row.find('./c[@r="' + ref + '"]');
    if (!c) {
        c = subElement(row, 'c');
        c.set('r', ref);
    }

    if (cellstyle) {
        var s = sheetData.find('.//c[@r="' + cellstyle + '"]') ? sheetData.find('.//c[@r="' + cellstyle + '"]').get('s') : undefined;
        if (s) c.set('s', s);
    }

    if (!_.isNumber(cellvalue)) {
        c.set('t', 's');
    }

    var v = c.find('./v');

    if (!v) {
        v = subElement(c, 'v');
    }

    if (!_.isNumber(cellvalue)) {
        v.text = this.sharedstrings.get(cellvalue).toString();
    } else {
        v.text = String(cellvalue);
    }

};

Sheet.prototype.save = function () {
    this.xml.save();
};

function intToExcelColumn(col) {
    var result = '';

    var mod;

    while (col > 0) {
        mod = (col - 1) % 26;
        result = String.fromCharCode(65 + mod) + result;
        col = Math.floor((col - mod) / 26);
    }

    return result;

}

Sheet.prototype.getCellValue = function (row, col) {
    var ref = intToExcelColumn(col) + row;
    var c = this.xml.find('.//c[@r="' + ref + '"]');
    if(!c) return undefined;
    switch (c.get('t')) {
        case 's':
            var index = parseInt(c.find('./v').text);
            return this.sharedstrings.get(index);
            break;
        default:
            return c.find('./v') ? c.find('./v').text : null;
            break;
    }
};