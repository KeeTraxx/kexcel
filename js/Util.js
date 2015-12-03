var fs = require("fs");
var xml2js = require("xml2js");
var Promise = require("bluebird");
var parseString = Promise.promisify(xml2js.parseString);
var readFile = Promise.promisify(fs.readFile);
var writeFile = Promise.promisify(fs.writeFile);
var builder = new xml2js.Builder();
function parseXML(input) {
    return parseString(input);
}
exports.parseXML = parseXML;
function loadXML(path) {
    return readFile(path).then(function (buffer) {
        return parseString(buffer.toString());
    });
}
exports.loadXML = loadXML;
function saveXML(xmlobj, path) {
    var contents = builder.buildObject(xmlobj);
    return writeFile(path, contents).thenReturn(path);
}
exports.saveXML = saveXML;
//# sourceMappingURL=Util.js.map