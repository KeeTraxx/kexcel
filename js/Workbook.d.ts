import * as stream from "stream";
import * as Promise from "bluebird";
import XMLFile = require("./XMLFile");
import SharedStrings = require("./SharedStrings");
import Sheet = require("./Sheet");
/**
 * Main class for KExcel. Use .new() and .open(file | stream) to open an .xlsx file.
 */
declare class Workbook {
    /**
     * These files are automatically loaded into files[]
     * @type {string[]}
     */
    private static autoload;
    /**
     * Temporary directory (created using the 'temp' library)
     */
    tempDir: string;
    /**
     * Dictionary which holds pointers to files.
     * @type {{}}
     */
    private files;
    /**
     * The array of sheets in this workbook.
     * @type {Array}
     */
    private sheets;
    /**
     * Template for an empty sheet
     */
    emptySheet: XMLFile;
    /**
     * Manage SharedStrings in the workbook.
     */
    sharedStrings: SharedStrings;
    /**
     * Holds the path of the source file (if opened from a file)
     */
    private filename;
    /**
     * The source stream (from another stream or file)
     */
    protected source: stream.Readable;
    constructor(input: any);
    /**
     * Creates an empty workbook
     * @returns {Promise<Workbook>}
     */
    static new(): Promise<Workbook>;
    /**
     * Opens an existing .xlsx file
     * @param input
     * @returns {Promise<Workbook>}
     */
    static open(input: stream.Readable | string): Promise<Workbook>;
    protected init(): Promise<Workbook>;
    protected initSharedStrings(): Promise<void>;
    protected initSheets(): Promise<void[]>;
    private extract();
    /**
     * Returns the xml object for the requested file
     * @param filePath
     * @returns {any}
     */
    getXML(filePath: string): any;
    /**
     * Creates a new sheet in the workbook
     * @param name Optionally set the name of the sheet
     * @returns {Sheet}
     */
    createSheet(name?: string): Sheet;
    /**
     * Get sheet at the specified index
     * @param index
     */
    getSheet(index: number): Sheet;
    getSheet(name: string): Sheet;
    pipe<T extends stream.Writable>(destination: T, options?: {
        end?: boolean;
    }): T;
}
export = Workbook;
