/**
 * Created by ktran on 21.03.14.
 */
var Q = require('q');
var fs = require('fs');
var fstream = require('fstream');
var unzip = require('unzip');
var path = require('path');
var util = require('util');
var elementtree = require('elementtree');
var _ = require('underscore');

var temp = require('temp');
//temp.track();

exports.open = function(file, callback) {
    var kexcel = new KExcel(file);
    kexcel.then(function(){
        callback(null, kexcel );
    });

};


function KExcel(file) {
    var file = file;

    var deferred = Q.defer();
    var kexcel = this;
    temp.mkdir('xlsx', function(err, dirPath){

        if ( err ) {
            deferred.reject(err);
        }

        kexcel.temppath = dirPath;

        var inputPath = path.join(dirPath,'test.txt');
        fs.writeFile(inputPath, 'test', console.log );

        var xlsxfile = fs.createReadStream(file);
        var outputDir = fstream.Writer(dirPath);
        xlsxfile.pipe(unzip.Parse()).pipe(outputDir);
        outputDir.on('close', function(){
            //console.warn('end');
            deferred.resolve(kexcel);
        });

    });

    this.getSheets = function(callback) {
        var filename = path.join(kexcel.temppath, 'xl/workbook.xml' );
        var contents = fs.readFileSync(filename);
        var etree = elementtree.parse(contents.toString());
        return _.map( etree.findall('sheets/sheet'), function(d){return {name: d.attrib.name, sheetId: d.attrib.sheetId} } );
    };

    this.getSheet = function(id) {
        var filename = path.join(kexcel.temppath, 'xl/worksheets/sheet'+id+'.xml' );
        var contents = fs.readFileSync(filename);
        var etree = elementtree.parse(contents.toString());
        return etree;
    };

    return deferred.promise;
};