import * as Promise from "bluebird";
import * as _ from "lodash";

import * as Util from "./Util";
import * as K from ".."

import path = require('path');
import {Saveable} from "./Saveable";
import {Workbook} from "./Workbook";

export class Sheet extends Saveable {

    protected id:string;
    public xml:any;
    public filename:string;
    private static refRegex:RegExp = /^([A-Z]+)(\d+)$/i;

    constructor(protected workbook:Workbook, protected workbookXml?:any, protected relationshipXml?:any) {
        super(null);
        if (this.workbookXml) {
            this.filename = this.relationshipXml.$.Target;
            this.path = path.join(this.workbook.tempDir, 'xl', this.filename);
            this.id = this.relationshipXml.$.Id;
        }
    }

    public load():Promise<any> {
        return this.xml ?
            Promise.resolve<any>(this.xml) :
            Util.loadXML(this.path)
                .then(xml => {
                    return this.xml = xml;
                });
    }

    public getName():string {
        return this.workbookXml.$.name;
    }

    public setName(name:string):void {
        this.workbookXml.$.name = name;
    }

    public create():void {
        this.addRelationship();
        this.addContentType();
        this.addToWorkbook();
        this.xml = _.cloneDeep(this.workbook.emptySheet.xml);
    }

    public copyFrom(sheet:Sheet) {
        this.xml = _.cloneDeep(sheet.xml);
        // delete selections if any
        delete this.xml.worksheet.sheetViews;
    }

    public setCellValue(rownum:number, colnum:number, cellvalue:string | number, copyCellStyle?:string);
    public setCellValue(ref:string, cellvalue:string | number, copyCellStyle?:string);
    public setCellValue(rownum_or_ref:any, colnum:number, cellvalue:string | number, copyCellStyle?:string):void {
        var cell = this.getCell(rownum_or_ref, colnum);
        var value:string|number = typeof colnum == 'string' ? colnum : cellvalue;
        var from:any = typeof colnum == 'string' ? cellvalue : copyCellStyle;
        if (cellvalue === undefined || cellvalue === null) {
            var matches = cell.$.r.match(Sheet.refRegex);
            var rownum = parseInt( matches[2] );
            // delete cell
            var row:K.Row = this.getRowXML(rownum);
            row.c.splice(row.c.indexOf(cell), 1);
        } else {
            this.setValue(cell, value);
            if (from !== undefined) {
                cell.$.s = this.getCell(from).$.s;
            }
        }
    }


    private setValue(cell:K.Cell, cellvalue:any):K.Cell {

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

    public getCellValue(rownum:number, colnum:number):string | number;
    public getCellValue(ref:string):string | number;
    public getCellValue(cell:K.Cell):string | number;
    public getCellValue(r:any, colnum?:number):string | number {
        var cell:K.Cell = this.getCell(r,colnum);

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

    public getCellFunction(rownum:number, colnum:number):string | number;
    public getCellFunction(ref:string):string | number;
    public getCellFunction(cell:K.Cell):string | number;
    public getCellFunction(r:any, colnum?:number):string {
        var cell:K.Cell = this.getCell(r,colnum);
        if (cell === undefined || cell === null || !cell.f) return undefined;

        var func = cell.f[0].hasOwnProperty('_') ? cell.f[0]._ : cell.f[0];

        return  '=' + func;
    }

    private getCell(rownum:number, colnum:number);
    private getCell(ref:string);
    private getCell(cell:K.Cell);
    private getCell(rownum_or_ref:any, colnum?:number):K.Cell {
        var rownum:number;
        var cellId:string;
        if (typeof rownum_or_ref == 'string') {
            var matches = rownum_or_ref.match(Sheet.refRegex);
            rownum = parseInt( matches[2] );
            //colnum = Sheet.excelColumnToInt(matches[1]);
            cellId = rownum_or_ref;
        } else if (typeof rownum_or_ref == 'number') {
            rownum = rownum_or_ref;
            //colnum = colnum;
            cellId = Sheet.intToExcelColumn(colnum) + rownum;
        } else {
            return rownum_or_ref;
        }
        var row:K.Row = this.getRowXML(rownum);
        var cell = _.find(row.c, cell => {
            return cell.$.r == cellId;
        });

        if (cell === undefined) {
            cell = {$: {r: cellId}};
            row.c = row.c || [];
            row.c.push(cell);

            row.c.sort((a, b) => {
                return Sheet.excelColumnToInt(a.$.r) - Sheet.excelColumnToInt(b.$.r);
            });
        }

        return cell;
    }

    public getRow(rownum:number):Array<string | number>;
    public getRow(row:K.Row):Array<string | number>
    public getRow(r:any):Array<string | number> {
        var row:K.Row = r;
        if (typeof r == 'number') {
            row = this.getRowXML(r);
        }

        if (!row.c) return undefined;

        var result:Array<string | number> = [];

        row.c.forEach((cell) => {
            result[Sheet.excelColumnToInt(cell.$.r) - 1] = this.getCellValue(cell);
        });
        return result;

    }

    public setRow(rownum:number, values:Array<string | number>):void;
    public setRow(row:K.Row, values:Array<string | number>):void;
    public setRow(r:any, values:Array<string | number>):void {
        var row:K.Row = r;
        if (typeof r == 'number') {
            row = this.getRowXML(r);
        }

        var rownum = row.$.r;

        row.c = _.compact(values.map((value, index) => {
            if (!value) return undefined;
            var cellId = Sheet.intToExcelColumn(index + 1) + rownum;
            return this.setValue({$: {r: cellId}}, value);
        }));

    }

    public appendRow(values:Array<string | number>):number {
        var row:K.Row = this.getRowXML(this.getLastRowNumber() + 1);
        this.setRow(row, values);
        return row.$.r;
    }


    public getLastRowNumber():number {
        if (this.xml.worksheet.sheetData && this.xml.worksheet.sheetData[0].row) {
            // Remove empty rows
            this.xml.worksheet.sheetData[0].row = this.xml.worksheet.sheetData[0].row.filter(function(row){
                return row.c && row.c.length > 0;
            });
            return +_.last<K.Row>(this.xml.worksheet.sheetData[0].row).$.r || 0;
        } else {
            return 0;
        }

    }

    private getRowXML(rownum:number):K.Row {
        if (!this.xml.worksheet.sheetData) {
            this.xml.worksheet.sheetData = [];
        }

        if (!this.xml.worksheet.sheetData[0]) {
            this.xml.worksheet.sheetData[0] = {row: []};
        }

        var rows:Array<K.Row> = this.xml.worksheet.sheetData[0].row;
        var row:K.Row = _.find<K.Row>(rows, r => {
            return r.$.r == rownum
        });

        if (!row) {
            row = {$: {r: rownum}};
            rows.push(row);
            rows.sort((row1, row2) => {
                return row1.$.r - row2.$.r;
            });
        }

        return row;

    }

    public toJSON():{} {
        var keys:Array<string | number> = this.getRow(1);
        var rows:Array<K.Row> = this.xml.worksheet.sheetData[0].row.slice(1);
        return rows.map(row => {
            return _.zipObject(keys, this.getRow(row));
        });
    }

    protected addRelationship():void {
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

    }

    protected addContentType():void {
        var contentTypes = this.workbook.getXML('[Content_Types].xml');
        this.path = path.join(this.workbook.tempDir, 'xl', 'worksheets', 'kexcel_' + this.id + '.xml');
        contentTypes.Types.Override.push({
            '$': {
                PartName: '/xl/' + this.filename,
                ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
            }
        });
    }

    protected addToWorkbook():void {
        var wbxml = this.workbook.getXML('xl/workbook.xml');
        var sheets = wbxml.workbook.sheets[0].sheet;
        this.workbookXml = {'$': {name: 'Sheet' + (sheets.length + 1), sheetId: sheets.length + 1, 'r:id': this.id}};
        sheets.push(this.workbookXml);
    }

    public static intToExcelColumn(col:number):string {
        var result = '';

        var mod;

        while (col > 0) {
            mod = (col - 1) % 26;
            result = String.fromCharCode(65 + mod) + result;
            col = Math.floor((col - mod) / 26);
        }

        return result;
    }

    public static excelColumnToInt(ref:string):number {
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
