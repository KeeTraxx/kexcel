var fs = require('fs');
var path = require('path');
var xml2js = require('xml2js');
var parser = new xml2js.Parser();

var builder = new xml2js.Builder();
var async = require('async');

function XMLFile(basedir, file, callback) {
    var readXmlFile = async.compose(parser.parseString, fs.readFile);
    var self = this;
    self.path = file;
    readXmlFile(path.join(basedir, file), function(err, obj){
        self.xml = obj;
        callback(err, self);
    });

    this.save = function(callback) {
        fs.writeFile(path.join(basedir, file), builder.buildObject(self.xml), callback);
    }
}

exports.readXmlFile = function(basedir, file, callback) {
    return new XMLFile(basedir, file, callback);
};

exports.XMLFile = XMLFile;