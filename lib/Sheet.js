/**
 * Created by ktran on 21.03.14.
 */
var fs = require('fs');
var path = require('path');
var elementtree = require('elementtree');
var subElement = elementtree.SubElement;
var _ = require('underscore');
var XMLFile = require('./XMLFile');

module.exports = Sheet;

function Sheet(workbook, attribs) {
    for( var i in attribs ) {
            this[i] = attribs[i];
    }
    this.sharedstrings = workbook.sharedStringsXml;
    console.log(this.sharedstrings);
    this.xml = new XMLFile(path.join(workbook.temppath, this.path ), this.sheetXML);
    var sheetView = this.xml.find('./sheetViews/sheetView');
    sheetView.set('tabSelected',null);
}

Sheet.prototype.replaceRow = function(rownum, obj) {
    var sheetData = this.xml.find('./sheetData');
    var row = sheetData.find('./row[@r="'+rownum+'"]');

    if ( !row ) {
        row = subElement(sheetData,'row');
        row.set('r', rownum);
    }

    var sheet = this;

    _.each(obj, function(d, ref){
        var c = row.find('./c[@r="'+ref+'"]');
        if ( !c ) {
            c = subElement(row,'c');
            c.set('r', ref);
        }
        c.set('t', 's');
        var v =  c.find('./v');

        if ( !v ) {
            v = subElement(c, 'v');
        }

        v.text = sheet.sharedstrings.get(d);
    });
};

Sheet.prototype.setCellValue = function(rownum,col,cellvalue,cellstyle) {
    var sheetData = this.xml.find('./sheetData');
    var row = sheetData.find('./row[@r="'+rownum+'"]');

    var ref = intToExcelColumn(col)+rownum;

    if ( !row ) {
        row = subElement(sheetData,'row');
        row.set('r', rownum);
    }

    var c = row.find('./c[@r="'+ref+'"]');
    if ( !c ) {
        c = subElement(row,'c');
        c.set('r', ref);
    }
    c.set('t', 's');

    if (cellstyle) {
        var s = sheetData.find('.//c[@r="'+cellstyle+'"]') ? sheetData.find('.//c[@r="'+cellstyle+'"]').get('s') : undefined;
        if (s) c.set('s', s);
    }

    var v =  c.find('./v');

    if ( !v ) {
        v = subElement(c, 'v');
    }
/*
    var t = subElement( is,'t' );
    t.text = cellvalue;
    */
    v.text = this.sharedstrings.get(cellvalue).toString();

}

Sheet.prototype.save = function() {
    this.xml.save();
};

function intToExcelColumn(col) {
    var result = '';

    var mod;

    while(col > 0) {
        mod = (col - 1) % 26;
        result = String.fromCharCode(65+mod) + result;
        col = Math.floor((col-mod)/26);
    }

    return result;

}