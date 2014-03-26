/**
 * Created by ktran on 21.03.14.
 */
var Q = require('q');
var fs = require('fs');
var fstream = require('fstream');
var unzip = require('unzip');
var archiver = require('archiver');
var archive = archiver('zip');
var path = require('path');
var elementtree = require('elementtree');
var subElement = elementtree.SubElement;
var _ = require('underscore');
var XMLFile = require('./XMLFile');

var Sheet = require('./Sheet');

var temp = require('temp');
temp.track();

var uuid = require('node-uuid');

exports.open = function(file, callback) {
    var kexcel = new Workbook(file);
    kexcel.then(function(){
        callback(null, kexcel );
    });

};

function Workbook(file) {
    var file = file;
    var deferred = Q.defer();

    var workbook = this;
    temp.mkdir('xlsx', function(err, dirPath){

        if ( err ) {
            deferred.reject(err);
        }

        workbook.temppath = dirPath;

        var xlsxfile = fs.createReadStream(file);
        var outputDir = fstream.Writer(dirPath);
        xlsxfile.pipe(unzip.Parse()).pipe(outputDir);
        outputDir.on('close', function(){
            parseSheets();
            deferred.resolve(workbook);
        });

    });

    var parseSheets = function() {
        //console.log(workbook.temppath, '/xl/workbook.xml');
        workbook.workbookXml = new XMLFile(path.join( workbook.temppath, '/xl/workbook.xml' ));
        workbook.relationshipXml = new XMLFile( path.join( workbook.temppath, '/xl/_rels/workbook.xml.rels' ));

        _.each( workbook.relationshipXml.findall('.//Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"]'), function(sheetdata) {
            var workbooksheet = workbook.workbookXml.find('./sheets/sheet[@r:id="'+sheetdata.attrib.Id+'"]');
            var attribs = {
                name: workbooksheet.get('name'),
                path : 'xl/'+sheetdata.attrib.Target,
                id: sheetdata.attrib.Id,
                sheetId: workbooksheet.attrib.sheetId
            }
            workbook.sheets.push( new Sheet( workbook, attribs ) );
        } );
    }

    this.duplicateSheet = function(sheet, newname) {
        //var id = uuid.v4();
        var id = 'rdupId'+( workbook.relationshipXml.findall('./Relationship').length+1 );
        var newsheet = new Sheet( workbook, {
            id: id,
            name: newname,
            path: 'xl/worksheets/'+id+'.xml',
            sheetXML: sheet.xml.write()
        } );
        var sheetId = _.max(workbook.sheets, function(s){return s.sheetId});
        this.sheets.push(newsheet);

        var sheets = workbook.workbookXml.find('./sheets');
        var sheet = subElement(sheets,'sheet');

        sheet.set('name', newname);
        sheet.set('sheetId', ++(sheetId.sheetId));
        sheet.set('r:id', id);

        var relationship = subElement( workbook.relationshipXml.find('./'), 'Relationship' );
        relationship.set('Id', id);
        relationship.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet');
        relationship.set('Target', 'worksheets/'+id+'.xml');

        return newsheet;
    }

    this.pipe = function(output, callback) {
        _.each(this.sheets, function(sheet){
            // TODO only save if dirty aka something has changed.
            sheet.save();
        });

        this.relationshipXml.save();
        this.workbookXml.save();


        output.on('close', function(){
            callback(null, workbook);
        });

        archive.on('error', function(err){
            throw err;
        });

        archive.pipe(output);
        archive.bulk([{expand: true, cwd: workbook.temppath, src: ['**','_rels/.rels']}]);
        archive.finalize();
    }

    this.sheets = [];

    return deferred.promise;
}