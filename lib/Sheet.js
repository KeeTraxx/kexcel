/**
 * Created by ktran on 21.03.14.
 */
var fs = require('fs');
var path = require('path');
var elementtree = require('elementtree');
var subElement = elementtree.SubElement;
var _ = require('underscore');

module.exports = Sheet;

function Sheet(workbook, id, name) {
    this.sheetId = id;
    this.name = name;

    this.filename = path.join(workbook.temppath, 'xl/worksheets/sheet'+id+'.xml' );
    var contents = fs.readFileSync(this.filename);
    this.etree = elementtree.parse(contents.toString());

}

Sheet.prototype.getName = function() {
    return this.name;
}

Sheet.prototype.getId = function() {
    return this.id;
}

Sheet.prototype.getTree = function() {
    return this.etree;
}

Sheet.prototype.replaceRow = function(rownum, obj) {
    var sheetData = this.etree.find('./sheetData');
    var row = sheetData.find('./row[@r="'+rownum+'"]');

    if ( !row ) {
        row = subElement(sheetData,'row');
    }

    _.each(obj, function(d, ref){
        try{
            var c = row.find('./c[@r="'+ref+'"]');
            var isel;
            if ( c ) {
                isel =  c.find('./is');
            } else {
                c.set('r', ref);
                isel = subElement(row, 'is');
            }
            c.set('t', 'is');

            if ( !isel ) {
                isel = subElement(c, 'is');
            }
            console.log(isel);

            var t = subElement( isel,'t' );
            console.log(t);
            t.text = d;
        } catch(e){
            console.log(e);
        }
    });
}

Sheet.prototype.save = function() {
    console.log('b');
    var contents = this.etree.write();
    console.log(contents);
    try{
        fs.writeFileSync(this.filename, contents, {flag: 'w' });
    } catch(e) {
        console.log(e);
    }
    console.log('d');
}