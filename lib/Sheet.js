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

    if ( cellvalue == undefined || cellvalue == null ) {
        return;
    }

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

    if (_.isNumber(cellvalue)) {
        var v = c.find('./v');
        if (!v) {
            v = subElement(c, 'v');
        }
        v.text = String(cellvalue);
    } else if( cellvalue[0] == '=' ) {
        // assume formula
        var f = c.find('./f');
        if (!f) {
            f = subElement(c, 'f');
        }
        f.text = String(cellvalue);
    } else {
        var v = c.find('./v');
        if (!v) {
            v = subElement(c, 'v');
        }
        v.text = this.sharedstrings.get(cellvalue).toString();
    }

};


Sheet.prototype.replaceRow = function (rownum, values) {
    var self = this;
    if (values.length == 0) {
        return;
    }

    var sheetData = this.xml.find('./sheetData');
    var row = _.find(sheetData.findall('./row'), function (d) {
        return d.get('r') == '' + rownum;
    });


    if (!row) {
        row = subElement(sheetData, 'row');
        row.set('r', rownum);
    }

    var col = 0;

    _.each(values, function(cellvalue){
        col++;
        var ref = intToExcelColumn(col) + rownum;
        var c = row.find('./c[@r="' + ref + '"]');
        var v;
        if (!c) {
            c = subElement(row, 'c');
            c.set('r', ref);
        }

        if (!_.isNumber(cellvalue)) {
            c.set('t', 's');
        }

        if (_.isNumber(cellvalue)) {
            v = c.find('./v');
            if (!v) {
                v = subElement(c, 'v');
            }
            v.text = String(cellvalue);
        } else if( cellvalue[0] == '=' ) {
            // assume formula
            var f = c.find('./f');
            if (!f) {
                f = subElement(c, 'f');
            }
            f.text = String(cellvalue);
        } else {
            v = c.find('./v');
            if (!v) {
                v = subElement(c, 'v');
            }
            v.text = self.sharedstrings.get(cellvalue).toString();
        }
    });
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

function excelColumnToInt(ref) {
    var number = 0;
    var pow = 1;
    for (var i = ref.length - 1; i >= 0; i--) {
        var c = ref.charCodeAt(i) - 64;
        if (c > 0 && c < 27) {
            number += (c) * pow;
            pow *= 26;
        }
    }

    return number;
}

Sheet.prototype.getCellValue = function (row, col) {
    var ref = intToExcelColumn(col) + row;
    var c = this.xml.find('.//c[@r="' + ref + '"]');
    if (!c) return undefined;
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

Sheet.prototype.getRowValues = function (rowId) {
    var row = _.find( this.xml.findall('./sheetData/row'), function(d) {
        return parseInt(d.get('r')) == rowId;
    } );
    if (!row) return undefined;

    var cells = row.findall('./c');
    var result = [];

    var sharedstrings = this.sharedstrings;

    _.each(cells, function (c) {

        var index = excelColumnToInt(c.get('r'));

        switch (c.get('t')) {
            case 's':
                var i = parseInt(c.find('./v').text);
                result[index] = sharedstrings.get(i);
                break;
            default:
                result[index] = c.find('./v') ? c.find('./v').text : undefined;
                break;
        }
    });
    return result;
};