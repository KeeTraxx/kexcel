/**
 * Created by ktran on 21.03.14.
 */
var elementtree = require('elementtree');
var subElement = elementtree.SubElement;
var _ = require('underscore');
var XMLFileOld = require('./XMLFileOld');

module.exports = SharedStrings;

function SharedStrings(filename, sheetXML) {
    this.xml = new XMLFileOld(filename, sheetXML);
    this.filename = filename;
    this.strings = [];
    this.parseStrings();
}

SharedStrings.prototype.parseStrings = function() {
    var obj = this;
    _.each( this.xml.findall('./si'), function(d) {
        obj.strings.push(d.findtext('t'));
    });
};

SharedStrings.prototype.save = function() {
    this.xml = new XMLFileOld(this.filename, '<sst count="'+this.strings.length+'" uniqueCount="'+this.strings.length+'" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></sst>');
    var root = this.xml.find('./');
    _.each(this.strings, function(d){
        var si = subElement(root, 'si');
        var t = subElement(si, 't');
        t.text = d;
    });

    this.xml.save();

};

SharedStrings.prototype.get = function(stringOrId) {
    if( typeof stringOrId === 'number') {
        return this.strings[stringOrId];
    }

    var index = _.indexOf(this.strings, stringOrId);

    if ( index != -1) {
        return index;
    } else {
        this.strings.push(stringOrId);
        return this.strings.length-1;
    }

};