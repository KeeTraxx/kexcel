/**
 * Created by ktran on 21.03.14.
 */
var Q = require('q');
var fs = require('fs');
var fstream = require('fstream');
var unzip = require('unzip');
var archiver = require('archiver');
var path = require('path');
var elementtree = require('elementtree');
var subElement = elementtree.SubElement;
var _ = require('underscore');
var XMLFile = require('./XMLFile');
var rimraf = require('rimraf');

var Sheet = require('./Sheet');
var SharedStrings = require('./SharedStrings');

var temp = require('temp');
temp.track();

exports.open = function (file, callback) {
    var kexcel = new Workbook(file);
    kexcel.then(function (workbook) {
        callback(null, workbook);
    });

};

function Workbook(file) {
    var deferred = Q.defer();

    var workbook = this;
    temp.mkdir('xlsx', function (err, dirPath) {

        if (err) {
            deferred.reject(err);
        }

        workbook.temppath = dirPath;

        var xlsxFile = fs.createReadStream(file);
        var outputDir = fstream.Writer(dirPath);
        xlsxFile.pipe(unzip.Parse()).pipe(outputDir);
        outputDir.on('close', function () {
            parseSheets();
            deferred.resolve(workbook);
        });

    });

    var parseSheets = function () {
        //console.log(workbook.temppath, '/xl/workbook.xml');
        workbook.workbookXml = new XMLFile(path.join(workbook.temppath, '/xl/workbook.xml'));
        workbook.relationshipXml = new XMLFile(path.join(workbook.temppath, '/xl/_rels/workbook.xml.rels'));
        workbook.contentTypesXml = new XMLFile(path.join(workbook.temppath, '/[Content_Types].xml'));

        var sharedstringsfile = workbook.relationshipXml.find('.//Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"]');
        if (sharedstringsfile) {
            workbook.sharedStringsXml = new SharedStrings(path.join(workbook.temppath, 'xl', sharedstringsfile.attrib.Target));
        } else {
            var r = subElement(workbook.relationshipXml.find('./'), 'Relationship');
            var filename = 'sharedStrings.xml';
            // Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"
            r.attrib.Target = filename;
            r.attrib.Type = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings';
            r.attrib.Id = 'rId' + (workbook.relationshipXml.findall('.//Relationship').length + 1);
            workbook.sharedStringsXml = new SharedStrings(path.join(workbook.temppath, 'xl', filename), '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"></sst>');

            var root = workbook.contentTypesXml.find('./');
            var or = subElement(root, 'Override');
            or.set('PartName', '/xl/'+filename);
            or.set('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml');
        }

        _.each(workbook.relationshipXml.findall('.//Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"]'), function (sheetdata) {
            var workbooksheet = workbook.workbookXml.find('./sheets/sheet[@r:id="' + sheetdata.attrib.Id + '"]');
            var attribs = {
                name: workbooksheet.get('name'),
                path: 'xl/' + sheetdata.attrib.Target,
                id: sheetdata.attrib.Id,
                sheetId: workbooksheet.attrib.sheetId
            };
            workbook.sheets.push(new Sheet(workbook, attribs));
        });

        //console.log(workbook.sharedStringsXml);

    };

    this.duplicateSheet = function (sheet, newname) {
        var id = 'rdupId' + ( workbook.relationshipXml.findall('./Relationship').length + 1 );
        var newsheet = new Sheet(workbook, {
            id: id,
            name: newname,
            path: 'xl/worksheets/' + id + '.xml',
            sheetXML: sheet.xml.write()
        });
        var sheetId = _.max(workbook.sheets, function (s) {
            return s.sheetId
        });
        this.sheets.push(newsheet);

        var sheets = workbook.workbookXml.find('./sheets');
        var sheet = subElement(sheets, 'sheet');

        sheet.set('name', newname);
        sheet.set('sheetId', ++(sheetId.sheetId));
        sheet.set('r:id', id);

        var relationship = subElement(workbook.relationshipXml.find('./'), 'Relationship');
        relationship.set('Id', id);
        relationship.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet');
        relationship.set('Target', 'worksheets/' + id + '.xml');

        return newsheet;
    };

    this.deleteSheetAt = function(sheetIndex) {
        var sheet = this.sheets[sheetIndex];
        var sheetElement = this.workbookXml.find("./sheets/sheet[@r:id='"+sheet.id+"']");
        this.workbookXml.find('./sheets').remove(sheetElement);

        var relationshipElement = this.relationshipXml.find("./Relationship[@Id='"+sheet.id+"']");
        this.relationshipXml.find('./').remove(relationshipElement);

        this.sheets.splice(sheetIndex,1);
    };

    this.pipe = function (output, callback) {
        var archive = archiver('zip');
        _.each(this.sheets, function (sheet) {
            // TODO only save if dirty aka something has changed.
            sheet.save();
        });

        this.relationshipXml.save();
        this.workbookXml.save();
        this.sharedStringsXml.save();
        this.contentTypesXml.save();


        output.on('close', function () {
            callback && callback(null, workbook);
        });

        archive.on('error', function (err) {
            throw err;
        });

        archive.pipe(output);
        archive.bulk([
            {expand: true, cwd: workbook.temppath, src: ['**', '_rels/.rels'], data: { date: new Date() } }
        ]);
        archive.finalize();
    };

    this.close = function (callback) {
        rimraf(workbook.temppath, callback || function () {
        });
    };

    this.sheets = [];

    return deferred.promise;
}