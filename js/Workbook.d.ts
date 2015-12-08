import * as stream from "stream";
import * as Promise from "bluebird";
import XMLFile = require("./XMLFile");
import SharedStrings = require("./SharedStrings");
import Sheet = require("./Sheet");
declare class Workbook {
    private static autoload;
    tempDir: string;
    private files;
    private sheets;
    emptySheet: XMLFile;
    sharedStrings: SharedStrings;
    private filename;
    protected source: stream.Readable;
    constructor(input: stream.Readable);
    constructor(input: string);
    static new(): Promise<Workbook>;
    static open(input: any): Promise<Workbook>;
    protected init(): Promise<Workbook>;
    protected initSharedStrings(): Promise<void>;
    protected initSheets(): Promise<void[]>;
    private extract();
    getXML(filePath: string): any;
    createSheet(name?: string): Sheet;
    getSheet(index: number): Sheet;
    getSheet(name: string): Sheet;
    pipe<T extends stream.Writable>(destination: T, options?: {
        end?: boolean;
    }): T;
}
export = Workbook;
