var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var Promise = require("bluebird");
var Util = require("./Util");
var Saveable = require('./Saveable');
var SharedStrings = (function (_super) {
    __extends(SharedStrings, _super);
    function SharedStrings(path, workbook) {
        _super.call(this, path);
        this.path = path;
        this.workbook = workbook;
        this.cache = {};
    }
    SharedStrings.prototype.load = function () {
        var _this = this;
        return this.xml ?
            Promise.resolve(this.xml) :
            Util.loadXML(this.path).then(function (xml) {
                _this.xml = xml;
            }).catch(function () {
                return Promise.all([
                    _this.addRelationship(),
                    _this.addContentType()
                ]).then(function () {
                    return Util.parseXML('<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"><si/></sst>');
                }).then(function (xml) {
                    _this.xml = xml;
                });
            }).finally(function () {
                _this.xml.sst.si.forEach(function (si, index) {
                    if (si.t) {
                        _this.cache[si.t[0]] = index;
                    }
                    else {
                    }
                });
                return _this.xml;
            });
    };
    SharedStrings.prototype.getIndex = function (s) {
        return this.cache[s] || this.storeString(s);
    };
    SharedStrings.prototype.getString = function (n) {
        var sxml = this.xml.sst.si[n];
        if (!sxml)
            return undefined;
        return sxml.hasOwnProperty('t') ? sxml.t[0] : _.compact(sxml.r.map(function (d) {
            return _.isString(d.t[0]) ? d.t[0] : null;
        })).join(' ');
    };
    SharedStrings.prototype.storeString = function (s) {
        var index = this.xml.sst.si.push({ t: [s] }) - 1;
        this.cache[s] = index;
        return index;
    };
    SharedStrings.prototype.addRelationship = function () {
        var relationships = this.workbook.getXML('xl/_rels/workbook.xml.rels');
        relationships.Relationships.Relationship.push({
            '$': {
                Id: 'rId1ss',
                Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
                Target: 'sharedStrings.xml'
            }
        });
    };
    SharedStrings.prototype.addContentType = function () {
        var contentTypes = this.workbook.getXML('[Content_Types].xml');
        contentTypes.Types.Override.push({
            '$': {
                PartName: '/xl/sharedStrings.xml',
                ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'
            }
        });
    };
    return SharedStrings;
})(Saveable);
module.exports = SharedStrings;
//# sourceMappingURL=SharedStrings.js.map