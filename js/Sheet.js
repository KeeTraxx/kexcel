var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var Promise = require("bluebird");
var _ = require("lodash");
var Util = require("./Util");
var Saveable = require('./Saveable');
var path = require('path');
var Sheet = (function (_super) {
    __extends(Sheet, _super);
    function Sheet(workbook, workbookXml, relationshipXml) {
        _super.call(this, null);
        this.workbook = workbook;
        this.workbookXml = workbookXml;
        this.relationshipXml = relationshipXml;
        if (this.workbookXml) {
            this.filename = this.relationshipXml.$.Target;
            this.path = path.join(this.workbook.tempDir, 'xl', this.filename);
            this.id = this.relationshipXml.$.Id;
        }
    }
    Sheet.prototype.load = function () {
        var _this = this;
        return this.xml ?
            Promise.resolve(this.xml) :
            Util.loadXML(this.path)
                .then(function (xml) {
                return _this.xml = xml;
            });
    };
    Sheet.prototype.getName = function () {
        return this.workbookXml.$.name;
    };
    Sheet.prototype.setName = function (name) {
        this.workbookXml.$.name = name;
    };
    Sheet.prototype.create = function () {
        this.addRelationship();
        this.addContentType();
        this.addToWorkbook();
        this.xml = _.cloneDeep(this.workbook.emptySheet.xml);
    };
    Sheet.prototype.copyFrom = function (sheet) {
        this.xml = _.cloneDeep(sheet.xml);
        delete this.xml.worksheet.sheetViews;
    };
    Sheet.prototype.setCellValue = function (rownum_or_ref, colnum, cellvalue, copyCellStyle) {
        var cell = this.getCell(rownum_or_ref, colnum);
        var value = typeof colnum == 'string' ? colnum : cellvalue;
        var from = typeof colnum == 'string' ? cellvalue : copyCellStyle;
        if (cellvalue === undefined || cellvalue === null) {
            var matches = cell.$.r.match(Sheet.refRegex);
            var rownum = parseInt(matches[2]);
            var row = this.getRowXML(rownum);
            row.c.splice(row.c.indexOf(cell), 1);
        }
        else {
            this.setValue(cell, value);
            if (from !== undefined) {
                cell.$.s = this.getCell(from).$.s;
            }
        }
    };
    Sheet.prototype.setValue = function (cell, cellvalue) {
        if (typeof cellvalue == 'number') {
            cell.v = [cellvalue];
            delete cell.f;
        }
        else if (cellvalue[0] == '=') {
            cell.f = [cellvalue.substr(1).replace(/;/g, ',')];
        }
        else {
            cell.v = [this.workbook.sharedStrings.getIndex(cellvalue)];
            cell.$.t = 's';
            delete cell.$.s;
        }
        return cell;
    };
    Sheet.prototype.getCellValue = function (r, colnum) {
        var cell = this.getCell(r, colnum);
        if (cell.$.t == 's') {
            return this.workbook.sharedStrings.getString(cell.v[0]);
        }
        else if (cell.f && cell.v) {
            var value = cell.v[0].hasOwnProperty('_') ? cell.v[0]._ : cell.v[0];
            return (value != '') ? value : undefined;
        }
        else if (cell.f) {
            return '=' + cell.f[0];
        }
        else {
            return cell.hasOwnProperty('v') ? cell.v[0] : undefined;
        }
    };
    Sheet.prototype.getCellFunction = function (r, colnum) {
        var cell = this.getCell(r, colnum);
        if (cell === undefined || cell === null || !cell.f)
            return undefined;
        var func = cell.f[0].hasOwnProperty('_') ? cell.f[0]._ : cell.f[0];
        return '=' + func;
    };
    Sheet.prototype.getCell = function (rownum_or_ref, colnum) {
        var rownum;
        var cellId;
        if (typeof rownum_or_ref == 'string') {
            var matches = rownum_or_ref.match(Sheet.refRegex);
            rownum = parseInt(matches[2]);
            cellId = rownum_or_ref;
        }
        else if (typeof rownum_or_ref == 'number') {
            rownum = rownum_or_ref;
            cellId = Sheet.intToExcelColumn(colnum) + rownum;
        }
        else {
            return rownum_or_ref;
        }
        var row = this.getRowXML(rownum);
        var cell = _.find(row.c, function (cell) {
            return cell.$.r == cellId;
        });
        if (cell === undefined) {
            cell = { $: { r: cellId } };
            row.c = row.c || [];
            row.c.push(cell);
            row.c.sort(function (a, b) {
                return Sheet.excelColumnToInt(a.$.r) - Sheet.excelColumnToInt(b.$.r);
            });
        }
        return cell;
    };
    Sheet.prototype.getRow = function (r) {
        var _this = this;
        var row = r;
        if (typeof r == 'number') {
            row = this.getRowXML(r);
        }
        if (!row.c)
            return undefined;
        var result = [];
        row.c.forEach(function (cell) {
            result[Sheet.excelColumnToInt(cell.$.r) - 1] = _this.getCellValue(cell);
        });
        return result;
    };
    Sheet.prototype.setRow = function (r, values) {
        var _this = this;
        var row = r;
        if (typeof r == 'number') {
            row = this.getRowXML(r);
        }
        var rownum = row.$.r;
        row.c = _.compact(values.map(function (value, index) {
            if (!value)
                return undefined;
            var cellId = Sheet.intToExcelColumn(index + 1) + rownum;
            return _this.setValue({ $: { r: cellId } }, value);
        }));
    };
    Sheet.prototype.appendRow = function (values) {
        var row = this.getRowXML(this.getLastRowNumber() + 1);
        this.setRow(row, values);
        return row.$.r;
    };
    Sheet.prototype.getLastRowNumber = function () {
        if (this.xml.worksheet.sheetData[0].row) {
            this.xml.worksheet.sheetData[0].row = this.xml.worksheet.sheetData[0].row.filter(function (row) {
                return row.c && row.c.length > 0;
            });
            return _.last(this.xml.worksheet.sheetData[0].row).$.r || 0;
        }
        else {
            return 0;
        }
    };
    Sheet.prototype.getRowXML = function (rownum) {
        if (!this.xml.worksheet.sheetData[0]) {
            this.xml.worksheet.sheetData[0] = { row: [] };
        }
        var rows = this.xml.worksheet.sheetData[0].row;
        var row = _.find(rows, function (r) {
            return r.$.r == rownum;
        });
        if (!row) {
            row = { $: { r: rownum } };
            rows.push(row);
            rows.sort(function (row1, row2) {
                return row1.$.r - row2.$.r;
            });
        }
        return row;
    };
    Sheet.prototype.toJSON = function () {
        var _this = this;
        var keys = this.getRow(1);
        var rows = this.xml.worksheet.sheetData[0].row.slice(1);
        return rows.map(function (row) {
            return _.zipObject(keys, _this.getRow(row));
        });
    };
    Sheet.prototype.addRelationship = function () {
        var relationships = this.workbook.getXML('xl/_rels/workbook.xml.rels');
        this.id = 'rId' + (relationships.Relationships.Relationship.length + 1);
        this.filename = 'worksheets/kexcel_' + this.id + '.xml';
        this.relationshipXml = {
            '$': {
                Id: this.id,
                Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
                Target: this.filename
            }
        };
        relationships.Relationships.Relationship.push(this.relationshipXml);
    };
    Sheet.prototype.addContentType = function () {
        var contentTypes = this.workbook.getXML('[Content_Types].xml');
        this.path = path.join(this.workbook.tempDir, 'xl', 'worksheets', 'kexcel_' + this.id + '.xml');
        contentTypes.Types.Override.push({
            '$': {
                PartName: '/xl/' + this.filename,
                ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
            }
        });
    };
    Sheet.prototype.addToWorkbook = function () {
        var wbxml = this.workbook.getXML('xl/workbook.xml');
        var sheets = wbxml.workbook.sheets[0].sheet;
        this.workbookXml = { '$': { name: 'Sheet' + (sheets.length + 1), sheetId: sheets.length + 1, 'r:id': this.id } };
        sheets.push(this.workbookXml);
    };
    Sheet.intToExcelColumn = function (col) {
        var result = '';
        var mod;
        while (col > 0) {
            mod = (col - 1) % 26;
            result = String.fromCharCode(65 + mod) + result;
            col = Math.floor((col - mod) / 26);
        }
        return result;
    };
    Sheet.excelColumnToInt = function (ref) {
        var number = 0;
        var pow = 1;
        for (var i = ref.length - 1; i >= 0; i--) {
            var c = ref.charCodeAt(i) - 64;
            if (c > 0 && c < 27) {
                number += (c) * pow;
                pow *= 26;
            }
        }
        return number;
    };
    Sheet.refRegex = /^([A-Z]+)(\d+)$/i;
    return Sheet;
})(Saveable);
module.exports = Sheet;
//# sourceMappingURL=Sheet.js.map