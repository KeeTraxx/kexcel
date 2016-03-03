"use strict";
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var Promise = require("bluebird");
var Util = require("./Util");
var Saveable = require('./Saveable');
var XMLFile = (function (_super) {
    __extends(XMLFile, _super);
    function XMLFile(path) {
        _super.call(this, path);
        this.path = path;
    }
    XMLFile.prototype.load = function () {
        var _this = this;
        return this.xml ? Promise.resolve(this.xml) : Util.loadXML(this.path).then(function (xml) {
            return _this.xml = xml;
        });
    };
    return XMLFile;
}(Saveable));
module.exports = XMLFile;
//# sourceMappingURL=XMLFile.js.map