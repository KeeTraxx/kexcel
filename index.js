/**
 * Created by ktran on 21.03.14.
 */
var Q = require('q');
var fs = require('fs');
var fstream = require('fstream');
var unzip = require('unzip');
var path = require('path');
var xml2js = require('xml2js');
var util = require('util');

var temp = require('temp');
//temp.track();

exports.open = function(file) {
    return new KExcel(file);
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

    this.getSheets = function() {
        var d = Q.defer();
        var filename = path.join(kexcel.temppath, 'xl/workbook.xml' );
        var contents = fs.readFileSync(filename);

        xml2js.parseString( contents , function(err, result){
            var builder = new xml2js.Builder();
            var xml = builder.buildObject(result);
            //console.log();
            d.resolve(util.inspect( result.workbook.sheets, false, null ));
        });

        return d.promise;
    };

    return deferred.promise;
};