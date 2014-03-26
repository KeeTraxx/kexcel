/**
 * Created by ktran on 26.03.14.
 */

module.exports = XMLFile;

var fs = require('fs');
var elementtree = require('elementtree');
var ET = elementtree.ElementTree;

ET.prototype.save = function(){
    var contents = this.write();
    fs.writeFileSync(this.path, contents, {flag: 'w' });
};

function XMLFile(p, sheetXML) {
    this.path = p;
    var contents;
    if ( sheetXML ) {
        contents = sheetXML;
    } else {
        contents = fs.readFileSync(this.path);
    }
    var xml = elementtree.parse(contents.toString());
    xml.path = this.path;
    return xml;

}
