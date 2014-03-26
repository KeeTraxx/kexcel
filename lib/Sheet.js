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
    }

    _.each(obj, function(d, ref){
        var c = row.find('./c[@r="'+ref+'"]');
        var is;
        if ( c ) {
            is =  c.find('./is');
        } else {
            c.set('r', ref);
            is = subElement(row, 'is');
        }
        c.set('t', 'inlineStr');

        if ( !is ) {
            is = subElement(c, 'is');
        }

        var t = subElement( is,'t' );
        t.text = d;
    });
};

Sheet.prototype.save = function() {
    this.xml.save();
};