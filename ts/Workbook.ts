import {WriteStream} from "fs";
const fs = require('fs');
import * as Promise from "bluebird";
import * as path from "path";
import * as _ from "lodash";
import * as rimraf from "rimraf";

const unzip = require('unzip');
var mkTempDir:any = Promise.promisify(require('temp').mkdir);
var fstream = require('fstream');
import * as archiver from "archiver";

import {ReadStream} from "fs";
import {Saveable} from "./Saveable";
import {Sheet} from "./Sheet";
import {XMLFile} from "./XMLFile";
import {SharedStrings} from "./SharedStrings";

/**
 * Main class for KExcel. Use .new() and .open(file | stream) to open an .xlsx file.
 */
export class Workbook {

    /**
     * These files are automatically loaded into files[]
     * @type {string[]}
     */
    private static autoload:Array<string> = [
        'xl/workbook.xml',
        'xl/_rels/workbook.xml.rels',
        '[Content_Types].xml'
    ];

    /**
     * Temporary directory (created using the 'temp' library)
     */
    public tempDir:string;

    /**
     * Dictionary which holds pointers to files.
     * @type {{}}
     */
    private files:{ [path: string]: Saveable; } = {};

    /**
     * The array of sheets in this workbook.
     * @type {Array}
     */
    public sheets:Sheet[] = [];

    /**
     * Template for an empty sheet
     */
    public emptySheet:XMLFile;

    /**
     * Manage SharedStrings in the workbook.
     */
    public sharedStrings:SharedStrings;

    /**
     * Holds the path of the source file (if opened from a file)
     */
    private filename:string;

    /**
     * The source stream (from another stream or file)
     */
    protected source:ReadStream;

    constructor(input:any) {
        if (typeof input == 'string') {
            this.filename = input;
            this.source = fs.createReadStream(input);
        } else {
            this.source = input;
        }
    }

    /**
     * Creates an empty workbook
     * @returns {Promise<Workbook>}
     */
    public static new():Promise<Workbook> {
        return Workbook.open(path.join(__dirname, '..', 'templates', 'empty.xlsx'));
    }

    /**
     * Opens an existing .xlsx file
     * @param input
     * @returns {Promise<Workbook>}
     */
    public static open(input:ReadStream | string):Promise<Workbook> {
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
        if ( this.filename && !fs.existsSync(this.filename)) return Promise.reject(this.filename + ' not found.');
        return mkTempDir('xlsx').then(tempDir => {
            this.tempDir = tempDir;
            return new Promise((resolve, reject) => {
                var parser = unzip.Parse();
                var writer = fstream.Writer(tempDir);
                var outstream = this.source.pipe(parser).pipe(writer);

                outstream.on('close', () => {
                    resolve(this);
                });

                parser.on('error', error => {
                    reject(error);
                });

            });
        });
    }

    /**
     * Returns the xml object for the requested file
     * @param filePath
     * @returns {any}
     */
    public getXML(filePath:string):any {
        return this.files[filePath].xml;
    }

    /**
     * Creates a new sheet in the workbook
     * @param name Optionally set the name of the sheet
     * @returns {Sheet}
     */
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

    /**
     * Get sheet at the specified index
     * @param index
     */
    public getSheet(index:number):Sheet;
    public getSheet(name:string):Sheet;
    public getSheet(input:any):Sheet {
        if (typeof input == 'number') return this.sheets[input];

        return _.find<Sheet>(this.sheets, sheet => {
            return sheet.getName() == input;
        });
    }

    public pipe<T extends WriteStream>(destination:T, options?:{ end?: boolean }):T {
        var archive:any = archiver('zip');
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

module.exports = {
    Workbook,
    open: Workbook.open,
    'new': Workbook.new
};