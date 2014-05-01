/**
 * Created by ktran on 21.03.14.
 */
var fs = require('fs');
var path = require('path');
var elementtree = require('elementtree');
var subElement = elementtree.SubElement;
var _ = require('underscore');
var XMLFile = require('./XMLFile');

module.exports = SharedStrings;

function SharedStrings(filename, sheetXML) {
    this.xml = new XMLFile(filename, sheetXML);
}
