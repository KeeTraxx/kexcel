var fs = require('fs');
var xml2js = require('xml2js');
var parser = new xml2js.Parser();

var builder = new xml2js.Builder();
var async = require('async');

function XMLFile(file, callback) {
    var readXmlFile = async.compose(parser.parseString, fs.readFile);
    var self = this;
    self.path = file;
    readXmlFile(file, function(err, obj){
        self.xml = obj;
        callback(err, self);
    });
}

exports.readXmlFile = function(file, callback) {
    new XMLFile(file, callback);
};