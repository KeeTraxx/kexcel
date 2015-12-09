var fs = require("fs");
var Promise = require("bluebird");
var path = require("path");
var _ = require("lodash");
var rimraf = require("rimraf");
var unzip = require('unzip');
var mkTempDir = Promise.promisify(require('temp').mkdir);
var fstream = require('fstream');
var archiver = require('archiver');
var XMLFile = require("./XMLFile");
var SharedStrings = require("./SharedStrings");
var Sheet = require("./Sheet");
/**
 * Main class for KExcel. Use .new() and .open(file | stream) to open an .xlsx file.
 */
var Workbook = (function () {
    function Workbook(input) {
        /**
         * Dictionary which holds pointers to files.
         * @type {{}}
         */
        this.files = {};
        /**
         * The array of sheets in this workbook.
         * @type {Array}
         */
        this.sheets = [];
        if (typeof input == 'string') {
            this.filename = input;
            this.source = fs.createReadStream(input);
        }
        else {
            this.source = input;
        }
    }
    /**
     * Creates an empty workbook
     * @returns {Promise<Workbook>}
     */
    Workbook.new = function () {
        return Workbook.open(path.join(__dirname, '..', 'templates', 'empty.xlsx'));
    };
    /**
     * Opens an existing .xlsx file
     * @param input
     * @returns {Promise<Workbook>}
     */
    Workbook.open = function (input) {
        var workbook = new Workbook(input);
        return workbook.init();
    };
    Workbook.prototype.init = function () {
        var _this = this;
        return this.extract().then(function () {
            var p = Workbook.autoload.map(function (filepath) {
                var xmlfile = new XMLFile(path.join(_this.tempDir, filepath));
                _this.files[filepath] = xmlfile;
                return xmlfile.load();
            });
            return Promise.all(p);
        }).then(function () {
            _this.emptySheet = new XMLFile(path.join(__dirname, '..', 'templates', 'emptysheet.xml'));
            return _this.emptySheet.load();
        }).then(function () {
            return Promise.all([_this.initSharedStrings(), _this.initSheets()]);
        }).thenReturn(this);
    };
    Workbook.prototype.initSharedStrings = function () {
        this.sharedStrings = new SharedStrings(path.join(this.tempDir, 'xl', 'sharedStrings.xml'), this);
        return this.sharedStrings.load();
    };
    Workbook.prototype.initSheets = function () {
        var _this = this;
        var wbxml = this.getXML('xl/workbook.xml');
        var relxml = this.getXML('xl/_rels/workbook.xml.rels');
        var p = _.map(wbxml.workbook.sheets[0].sheet, function (sheetXml) {
            var r = _.find(relxml.Relationships.Relationship, function (rel) {
                return rel.$.Id == sheetXml.$['r:id'];
            });
            var sheet = new Sheet(_this, sheetXml, r);
            _this.files[sheet.filename] = sheet;
            _this.sheets.push(sheet);
            return sheet.load();
        });
        return Promise.all(p);
    };
    Workbook.prototype.extract = function () {
        var _this = this;
        if (this.filename && !fs.existsSync(this.filename))
            return Promise.reject(this.filename + ' not found.');
        return mkTempDir('xlsx').then(function (tempDir) {
            _this.tempDir = tempDir;
            return new Promise(function (resolve, reject) {
                var parser = unzip.Parse();
                var writer = fstream.Writer(tempDir);
                var outstream = _this.source.pipe(parser).pipe(writer);
                outstream.on('close', function () {
                    resolve(_this);
                });
                parser.on('error', function (error) {
                    reject(error);
                });
            });
        });
    };
    /**
     * Returns the xml object for the requested file
     * @param filePath
     * @returns {any}
     */
    Workbook.prototype.getXML = function (filePath) {
        return this.files[filePath].xml;
    };
    /**
     * Creates a new sheet in the workbook
     * @param name Optionally set the name of the sheet
     * @returns {Sheet}
     */
    Workbook.prototype.createSheet = function (name) {
        var sheet = new Sheet(this);
        sheet.create();
        this.sheets.push(sheet);
        this.files[sheet.filename] = sheet;
        if (name != undefined) {
            sheet.setName(name);
        }
        return sheet;
    };
    Workbook.prototype.getSheet = function (input) {
        if (typeof input == 'number')
            return this.sheets[input];
        return _.find(this.sheets, function (sheet) {
            return sheet.getName() == input;
        });
    };
    Workbook.prototype.pipe = function (destination, options) {
        var _this = this;
        var archive = archiver('zip');
        Promise.all(_.map(this.files, function (file) {
            return file.save();
        })).then(function () {
            return _this.sharedStrings.save();
        }).then(function () {
            archive.on('finish', function () {
                // Async version somehow doesn't work?
                /*rimraf(this.tempDir, function(error){
                 console.log('errr');
                 console.log(error);
                 });*/
                rimraf.sync(_this.tempDir);
            });
            archive.pipe(destination, options);
            archive.bulk([
                { expand: true, cwd: _this.tempDir, src: ['**', '_rels/.rels'], data: { date: new Date() } }
            ]);
            archive.finalize();
        });
        return archive;
    };
    /**
     * These files are automatically loaded into files[]
     * @type {string[]}
     */
    Workbook.autoload = [
        'xl/workbook.xml',
        'xl/_rels/workbook.xml.rels',
        '[Content_Types].xml'
    ];
    return Workbook;
})();
module.exports = Workbook;
//# sourceMappingURL=Workbook.js.map