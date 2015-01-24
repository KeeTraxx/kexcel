var fs = require('fs');
var path = require('path');

var temp = require('temp');

var XMLFile = require('./XMLFile');

var xml2js = require('xml2js');
var parser = new xml2js.Parser();
var builder = new xml2js.Builder();
var fstream = require('fstream');
var unzip = require('unzip');
var archiver = require('archiver');
var _ = require('underscore');

var async = require('async');

function kexcel(file, callback) {
    var self = this;

    var tempdir = temp.mkdirSync('xlsx');
    var outputDir = fstream.Writer(tempdir);

    fs.createReadStream(file).pipe(unzip.Parse()).pipe(outputDir).on('close', function () {
        var readXmlFile = async.compose(parser.parseString, fs.readFile);
        self.files = {};
        var files = [
            'xl/workbook.xml',
            'xl/_rels/workbook.xml.rels',
            '[Content_Types].xml',
            'xl/sharedStrings.xml'
        ];

        if (!fs.existsSync(path.join(tempdir, 'xl/sharedStrings.xml')))
            fs.writeFileSync(path.join(tempdir, 'xl/sharedStrings.xml'), '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"></sst>');

        async.each(files, function (file, done) {
                readXmlFile(path.join(tempdir, file), function (err, xml) {
                    self.files[file] = xml;
                    done();
                });
            }, function (err) {
                // TODO: Check for Sharedstrings
                callback(err, self);
            }
        );
    });

    this.pipe = function (output, callback) {
        var archive = archiver('zip');
        async.each(Object.keys( self.files ), function(filename, next){
            fs.writeFile( path.join(tempdir, filename ), builder.buildObject(self.files[filename]), next );
        }, function() {
            output.on('close', function () {
                callback && callback(null, workbook);
            });

            archive.on('error', function (err) {
                throw err;
            });

            archive.pipe(output);
            archive.bulk([
                {expand: true, cwd: tempdir, src: ['**', '_rels/.rels'], data: { date: new Date() } }
            ]);
            archive.finalize();
        });
    };

}

exports.open = function (file, callback) {
    return new kexcel(file, callback);
};

exports.new = function (callback) {
    return new kexcel(path.join(__dirname, '..', 'templates', 'empty.xlsx'), callback);
};