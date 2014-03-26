/**
 * Created by ktran on 21.03.14.
 */
var Q = require('q');
var fs = require('fs');
var fstream = require('fstream');
var unzip = require('unzip');
var path = require('path');
var elementtree = require('elementtree');
var _ = require('underscore');

var Sheet = require('./Sheet');

var temp = require('temp');
//temp.track();

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

        var inputPath = path.join(dirPath,'test.txt');
        fs.writeFile(inputPath, 'test', console.log );

        var xlsxfile = fs.createReadStream(file);
        var outputDir = fstream.Writer(dirPath);
        xlsxfile.pipe(unzip.Parse()).pipe(outputDir);
        outputDir.on('close', function(){
            parseSheets();
            deferred.resolve(workbook);
        });

    });

    var parseSheets = function() {
        var filename = path.join(workbook.temppath, 'xl/workbook.xml' );
        var contents = fs.readFileSync(filename);
        var etree = elementtree.parse(contents.toString());

        _.each( etree.findall('sheets/sheet'), function(sheetdata) {
            console.log(sheetdata);
            console.log(sheetdata.attrib);
            workbook.sheets.push( new Sheet( workbook, sheetdata.attrib.sheetId, sheetdata.attrib.name ) );
        } );
    }

    this.sheets = [];

    return deferred.promise;
};