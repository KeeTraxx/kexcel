var fs = require('fs');
var path = require('path');

var temp = require('temp');

var XMLFile = require('./XMLFile');
var Sheet = require('./Sheet');

var fstream = require('fstream');
var unzip = require('unzip');
var archiver = require('archiver');
var _ = require('lodash');
var async = require('async');

function kexcel(file, callback) {
    var self = this;
    var files = [];
    var sheets = [];

    var tempdir = temp.mkdirSync('xlsx');
    var outputDir = fstream.Writer(tempdir);

    function getFile(path) {
        return _.find(files, function (file) {
            return file.path == path;
        });
    }

    fs.createReadStream(file).pipe(unzip.Parse()).pipe(outputDir).on('close', function () {
        var fileList = [
            'xl/workbook.xml',
            'xl/_rels/workbook.xml.rels',
            '[Content_Types].xml',
            'xl/sharedStrings.xml'
        ];

        if (!fs.existsSync(path.join(tempdir, 'xl/sharedStrings.xml')))
            fs.writeFileSync(path.join(tempdir, 'xl/sharedStrings.xml'), '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"><si/></sst>');

        async.each(fileList, function (file, done) {
                files.push(XMLFile.readXmlFile(tempdir, file, done));
                /*readXmlFile(path.join(tempdir, file), function (err, xml) {
                 self.files[file] = xml;
                 done();
                 });*/
            }, function (err) {
                var relationships = getFile('xl/_rels/workbook.xml.rels').xml.Relationships.Relationship;
                var workbook = getFile('xl/workbook.xml').xml.workbook;
                var ss = _.find(relationships, function (relationship) {
                    return relationship.$.Target == 'sharedStrings.xml';
                });

                if (!ss) {
                    var overrides = getFile('[Content_Types].xml').xml.Types.Override;
                    overrides.push({
                        '$': {
                            PartName: '/xl/sharedStrings.xml',
                            ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'
                        }
                    });
                    relationships.push({
                            '$': {
                                Id: 'rId1ss',
                                Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
                                Target: 'sharedStrings.xml'
                            }
                        }
                    );
                }

                var sheetFilenames = _.map(workbook.sheets[0].sheet, function(sheet){
                    var rId = sheet.$['r:id'];
                    return _.find(relationships, function(relationship){
                        return relationship.$.Id == rId;
                    }).$.Target;
                });

                async.eachSeries(sheetFilenames, function (file, done) {
                    Sheet.readXmlFile(tempdir, path.join('xl', file), getFile('xl/sharedStrings.xml').xml.sst.si, function (err, sheet) {
                        sheets.push(sheet);
                        done();
                    });
                    /*
                     readXmlFile(path.join(tempdir, 'xl', file), function (err, xml) {
                     self.sheets['xl/' + file] = xml;
                     done();
                     });*/
                }, function (err) {
                    console.log(sheets);
                    callback(err, self);
                });

            }
        );
    });

    this.pipe = function (output, callback) {
        var archive = archiver('zip');
        async.each(files.concat(sheets), function (file, next) {
            file.save(next);
        }, function () {
            output.on('close', function () {
                callback && callback(null, workbook);
            });

            archive.on('error', function (err) {
                throw err;
            });

            archive.pipe(output);
            archive.bulk([
                {expand: true, cwd: tempdir, src: ['**', '_rels/.rels'], data: {date: new Date()}}
            ]);
            archive.finalize();
        });
    };

    this.getSheet = function (index) {
        return sheets[index];
    };

    this.duplicateSheet = function (sheetIndex, newTitle) {
        var relationships = getFile('xl/_rels/workbook.xml.rels').xml.Relationships.Relationship;
        var id = relationships.length + 1;
        var contents = _.cloneDeep(sheets[sheetIndex].xml);
        delete contents.worksheet.sheetViews;
        var sheet = Sheet.newXmlFile(tempdir, path.join('xl', 'worksheets', 'dup' + id + '.xml'), contents, getFile('xl/sharedStrings.xml').xml.sst.si);
        sheets.push(sheet);

        var sheetsXml = getFile('xl/workbook.xml').xml.workbook.sheets[0].sheet;
        sheetsXml.push(
            { '$': { name: newTitle || ('Sheet'+id), sheetId: id, 'r:id': 'rId' + id } }
        );
        //console.log(relationships);
        relationships.push({ '$':
        { Id: 'rId'+id,
            Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
            Target: 'worksheets/dup' + id + '.xml' } });

        var overrides = getFile('[Content_Types].xml').xml.Types.Override;

        overrides.push({ '$':
        { PartName: ['/xl', 'worksheets', 'dup' + id + '.xml'].join('/'),
            ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml' } });

        return sheet;
    }
}

exports.open = function (file, callback) {
    return new kexcel(file, callback);
};

exports.new = function (callback) {
    return new kexcel(path.join(__dirname, '..', 'templates', 'empty.xlsx'), callback);
};