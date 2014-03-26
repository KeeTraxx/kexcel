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

    _.each(obj, function(d, ref){
        var c = row.find('./c[@r="'+ref+'"]');
        if ( !c ) {
            c = subElement(row,'c');
            c.set('r', ref);
        }
        c.set('t', 'inlineStr');
        var is =  c.find('./is');

        if ( !is ) {
            is = subElement(c, 'is');
        }

        var t = subElement( is,'t' );
        t.text = d;
    });
};

Sheet.prototype.setCellValue = function(rownum,col,cellvalue,cellstyle) {
    var sheetData = this.xml.find('./sheetData');
    var row = sheetData.find('./row[@r="'+rownum+'"]');

    var ref = intToExcelColumn(col)+rownum;

    if ( !row ) {
        row = subElement(sheetData,'row');
    }

    var c = row.find('./c[@r="'+ref+'"]');
    if ( !c ) {
        c = subElement(row,'c');
        c.set('r', ref);
    }
    c.set('t', 'inlineStr');

    if (cellstyle) {
        var s = sheetData.find('.//c[@r="'+cellstyle+'"]') ? sheetData.find('.//c[@r="'+cellstyle+'"]').get('s') : undefined;
        if (s) c.set('s', s);
    }

    var is =  c.find('./is');

    if ( !is ) {
        is = subElement(c, 'is');
    }

    var t = subElement( is,'t' );
    t.text = cellvalue;

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