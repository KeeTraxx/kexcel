import * as fs from "fs";
import * as stream from "stream";
import * as Promise from "bluebird";
import * as path from "path";
import * as _ from "lodash";
import * as rimraf from "rimraf";

var unzip = require('unzip');
var mkTempDir:any = Promise.promisify(require('temp').mkdir);
var fstream = require('fstream');
var archiver = require('archiver');

import XMLFile = require("./XMLFile");
import SharedStrings = require("./SharedStrings");
import Sheet = require("./Sheet");
import Saveable = require("./Saveable");
import Util = require("./Util");

class Workbook {
    private static autoload:Array<string> = [
        'xl/workbook.xml',
        'xl/_rels/workbook.xml.rels',
        '[Content_Types].xml'
    ];
    public tempDir:string;
    private files:{ [path: string]: Saveable; } = {};
    private sheets:Array<Sheet> = [];
    public emptySheet:XMLFile;
    public sharedStrings:SharedStrings;

    protected source:stream.Readable;

    constructor(input:stream.Readable);
    constructor(input:string);
    constructor(input:any) {
        this.source = typeof input == 'string' ? fs.createReadStream(input) : input;
    }

    public static new():Promise<Workbook> {
        return Workbook.open(path.join(__dirname, '..', 'templates', 'empty.xlsx'));
    }

    public static open(input:any):Promise<Workbook> {
        var workbook = new Workbook(input);
        return workbook.init();
    }

    protected init():Promise<Workbook> {
        return this.extract().then(() => {
            var p:Array<Promise<void>> = Workbook.autoload.map(filepath => {
                var xmlfile = new XMLFile(path.join(this.tempDir, filepath));
                this.files[filepath] = xmlfile;
                return xmlfile.load();
            });
            return Promise.all(p);
        }).then(() => {
            this.emptySheet = new XMLFile(path.join(__dirname, '..', 'templates', 'emptysheet.xml'));
            return this.emptySheet.load();
        }).then(() => {
            return Promise.all([this.initSharedStrings(), this.initSheets()])
        }).thenReturn(this);

    }

    protected initSharedStrings():Promise<void> {
        this.sharedStrings = new SharedStrings(path.join(this.tempDir, 'xl', 'sharedStrings.xml'), this);
        return this.sharedStrings.load();
    }

    protected initSheets():Promise<void[]> {
        var wbxml = this.getXML('xl/workbook.xml');
        var relxml = this.getXML('xl/_rels/workbook.xml.rels');
        var p:Array<Promise<void>> = _.map<any, Promise<void>>(wbxml.workbook.sheets[0].sheet, sheetXml => {
            var r:any = _.find<any>(relxml.Relationships.Relationship, rel => {
                return rel.$.Id == sheetXml.$['r:id'];
            });
            var sheet = new Sheet(this, sheetXml, r);
            this.files[sheet.filename] = sheet;
            this.sheets.push(sheet);
            return sheet.load();
        });
        return Promise.all(p);
    }

    private extract():Promise<void> {
        return mkTempDir('xlsx').then(tempDir => {
            this.tempDir = tempDir;
            return new Promise((resolve, reject) => {
                var outstream = this.source.pipe(unzip.Parse()).pipe(fstream.Writer(tempDir));

                outstream.on('close', () => {
                    resolve(this)
                });
                outstream.on('error', () => reject);

            })
        });
    }

    public getXML(filePath:string):any {
        return this.files[filePath].xml;
    }

    public createSheet(name?:string):Sheet {
        var sheet = new Sheet(this);
        sheet.create();
        this.sheets.push(sheet);
        this.files[sheet.filename] = sheet;
        if (name != undefined) {
            sheet.setName(name);
        }
        return sheet;
    }

    public getSheet(index:number):Sheet;
    public getSheet(name:string):Sheet;
    public getSheet(input:any):Sheet {
        if (typeof input == 'number') return this.sheets[input];

        return _.find<Sheet>(this.sheets, sheet => {
            return sheet.getName() == input;
        });
    }

    public pipe<T extends stream.Writable>(destination:T, options?:{ end?: boolean }):T {
        var archive = archiver('zip');
        Promise.all(_.map(this.files, function (file:XMLFile) {
            return file.save();
        })).then(() => {
            return this.sharedStrings.save();
        }).then(() => {
            archive.on('finish', () => {
                // Async version somehow doesn't work?
                /*rimraf(this.tempDir, function(error){
                 console.log('errr');
                 console.log(error);
                 });*/
                rimraf.sync(this.tempDir);
            });
            archive.pipe(destination, options);
            archive.bulk([
                {expand: true, cwd: this.tempDir, src: ['**', '_rels/.rels'], data: {date: new Date()}}
            ]);

            archive.finalize();
        });
        return archive;
    }

}

export = Workbook;