import * as fs from "fs";
import * as xml2js from "xml2js";
import * as Promise from "bluebird";
import * as _ from "lodash";

import * as Util from "./Util";
import * as K from ".."

import XMLFile = require('./XMLFile');
import Workbook = require('./Workbook');
import Saveable = require('./Saveable');
import path = require('path');

var parseString: any = Promise.promisify(xml2js.parseString);
var readFile = Promise.promisify(fs.readFile);
var fileExists = Promise.promisify(fs.exists);
var writeFile: any = Promise.promisify(fs.writeFile);
var builder = new xml2js.Builder();

class Sheet extends Saveable {

    protected id: string;
    public xml: any;
    public filename: string;

    constructor(protected workbook: Workbook, protected workbookXml?: any, protected relationshipXml?: any) {
        super(null);
        if (this.workbookXml) {
            this.filename = this.relationshipXml.$.Target;
            this.path = path.join(this.workbook.tempDir, 'xl', this.filename);
            this.id = this.relationshipXml.$.Id;
        }
    }

    public load(): Promise<any> {
        return this.xml ?
            Promise.resolve<any>(this.xml) :
            Util.loadXML(this.path)
                .then(xml => {
                    return this.xml = xml;
                });
    }

    public setName(name: string): void {
        this.workbookXml.$.name = name;
    }

    public getName(): string {
        return this.workbookXml.$.name;
    }

    public create(): void {
        this.addRelationship();
        this.addContentType();
        this.addToWorkbook();
        this.xml = this.workbook.emptySheet.xml;
    }

    public copyFrom(sheet: Sheet) {
        this.xml = _.cloneDeep(sheet.xml);
        // delete selections if any
        delete this.xml.worksheet.sheetViews;
    }

    public setCellValue(rownum: number, colnum: number, cellvalue: any, copyCellStyle?: K.Cell) {
        var cell = this.getCell(rownum, colnum);
        if (cellvalue === undefined || cellvalue === null) {
            // delete cell
            var row: K.Row = this.getRow(rownum);
            row.c.splice(row.c.indexOf(cell), 1);
        } else {
            this.setValue(cell, cellvalue);
            if (copyCellStyle !== undefined) {
                cell.$.s = copyCellStyle.$.s;
            }
        }
    }

    public getCellValueByRef(ref: string) {
        var matches = ref.match(/^([A-Z]+)(\d+)$/i);
        return this.getValue(this.getCell(parseInt(matches[2]), Sheet.excelColumnToInt(matches[1])));
    }

    private getCell(rownum: number, colnum: number): K.Cell {
        var row: K.Row = this.getRow(rownum);
        var cellId: string = Sheet.intToExcelColumn(colnum) + rownum;

        var cell = _.find(row.c, cell => {
            return cell.$.r == cellId;
        });

        if (cell === undefined) {
            cell = { $: { r: cellId } };
            row.c = row.c || [];
            row.c.push(cell);

            row.c.sort((a, b) => {
                return Sheet.excelColumnToInt(a.$.r) - Sheet.excelColumnToInt(b.$.r);
            });
        }

        return cell;
    }

    public setRowValues(rownum: number, values: Array<string | number>): void {
        var row: K.Row = this.getRow(rownum);
        this.setRow(row, values);
    }

    private setRow(row: K.Row, values: Array<string | number>): void {
        var rownum = row.$.r;
        row.c = _.compact(values.map((value, index) => {
            if (!value) return undefined;
            var cellId = Sheet.intToExcelColumn(index + 1) + rownum;
            return this.setValue({ $: { r: cellId } }, value);
        }));
    }

    public appendRowValues(values: Array<string | number>): void {
        var row: K.Row = this.getRow(this.getRowCount() + 1);
        this.setRow(row, values);
    }

    public getRowValues(rownum: number): Array<string | number> {
        var row = this.getRow(rownum);
        return this.getRowData(row);
    }

    private getRowData(row: K.Row): Array<string | number> {

        if (!row.c) return undefined;

        var result: Array<string | number> = [];

        row.c.forEach((cell, i) => {
            result[Sheet.excelColumnToInt(cell.$.r) - 1] = this.getValue(cell);
        });
        return result;
    }

    public getRowCount(): number {
        if (this.xml.worksheet.sheetData[0].row) {
            return _.last<K.Row>(this.xml.worksheet.sheetData[0].row).$.r || 0;
        } else {
            return 0;
        }

    }

    public getCellValue(rownum: number, colnum: number): string {
        return this.getValue(this.getCell(rownum, colnum));
    }

    public getValue(cell: K.Cell): string {
        if (cell === undefined || cell === null) return undefined;

        if (cell.$.t == 's') {
            // Sharedstring
            return this.workbook.sharedStrings.getString(cell.v[0]);
        } else if (cell.f && cell.v) {
            var value = cell.v[0].hasOwnProperty('_') ? cell.v[0]._ : cell.v[0];
            return (value != '') ? value : undefined;
        } else if (cell.f) {
            return '=' + cell.f[0];
        } else {
            return cell.hasOwnProperty('v') ? cell.v[0] : undefined;
        }
    }

    private getRow(rownum: number): K.Row {
        if (!this.xml.worksheet.sheetData[0]) {
            this.xml.worksheet.sheetData[0] = { row: [] };
        }
        var rows: Array<K.Row> = this.xml.worksheet.sheetData[0].row;
        var row: K.Row = _.find<K.Row>(rows, r => { return r.$.r == rownum });

        if (!row) {
            row = { $: { r: rownum } };
            rows.push(row);
            rows.sort((row1, row2) => {
                return row1.$.r - row2.$.r;
            });
        }

        return row;

    }

    private setValue(cell: K.Cell, cellvalue: any): K.Cell {
        if (cellvalue === null || cellvalue === undefined) {
            return;
        }

        if (typeof cellvalue == 'number') {
            // number
            cell.v = [cellvalue];
            delete cell.f;
        } else if (cellvalue[0] == '=') {
            // function
            cell.f = [cellvalue.substr(1).replace(/;/g, ',')];
        } else {
            // assume string
            cell.v = [this.workbook.sharedStrings.getIndex(cellvalue)];
            cell.$.t = 's';

            // reset cell type
            delete cell.$.s;
        }
        return cell;
    }

    public toJSON(): any {
        var keys: Array<string | number> = this.getRowValues(1);
        var rows: Array<K.Row> = this.xml.worksheet.sheetData[0].row.slice(1);
        return rows.map(row => {
            return _.zipObject(keys, this.getRowData(row));
        });
    }

    protected addRelationship(): void {
        var relationships = this.workbook.getXML('xl/_rels/workbook.xml.rels')
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

    }

    protected addContentType(): void {
        var contentTypes = this.workbook.getXML('[Content_Types].xml');
        this.path = path.join(this.workbook.tempDir, 'xl', 'worksheets', 'kexcel_' + this.id + '.xml');
        contentTypes.Types.Override.push({
            '$': {
                PartName: '/xl/' + this.filename,
                ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
            }
        });
    }

    protected addToWorkbook(): void {
        var wbxml = this.workbook.getXML('xl/workbook.xml');
        var sheets = wbxml.workbook.sheets[0].sheet;
        this.workbookXml = { '$': { name: 'Sheet' + (sheets.length + 1), sheetId: sheets.length + 1, 'r:id': this.id } };
        sheets.push(this.workbookXml);
    }

    public static intToExcelColumn(col: number): string {
        var result = '';

        var mod;

        while (col > 0) {
            mod = (col - 1) % 26;
            result = String.fromCharCode(65 + mod) + result;
            col = Math.floor((col - mod) / 26);
        }

        return result;
    }

    public static excelColumnToInt(ref: string): number {
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
    }

}

export = Sheet;