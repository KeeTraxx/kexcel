var fs = require('fs');
var path = require('path');
var util = require('util');
var xml2js = require('xml2js');
var parser = new xml2js.Parser();
var builder = new xml2js.Builder();
var async = require('async');

var _ = require('lodash');


function Sheet(basedir, file, contents, sharedStrings) {
    var self = this;
    self.path = file;
    this.xml = contents;

    this.save = function (callback) {
        console.log(util.inspect(self.xml, false, null));
        fs.writeFile(path.join(basedir, file), builder.buildObject(self.xml), callback);
    };

    function getCellbyRef(ref) {
        var matches = ref.match(/([A-Z]+)([0-9]+)/);

        var rownum = matches[2];
        var colnum = excelColumnToInt(matches[1]);

        var rows = self.xml.worksheet.sheetData[0].row;

        var row = _.find(rows, function (r) {
            return r.$.r == rownum;
        });

        if ( row && row.hasOwnProperty('c') ) {
            return _.find( row.c, function(cell){
                return cell.$.r == ref;
            });
        } else {
            return undefined;
        }


    }

    this.setCellValue = function (rownum, colnum, cellvalue, copyCellStyleFrom) {
        // Do nothing if cellvalue isn't set
        if (cellvalue == null || cellvalue == undefined) {
            return;
        }


        this.xml.worksheet.sheetData[0] = this.xml.worksheet.sheetData[0] || {row : []};

        var rows = this.xml.worksheet.sheetData[0].row;

        var row = _.find(rows, function (r) {
            return r.$.r == rownum;
        });

        if (!row) {
            row = {
                '$': {r: rownum}
            };
            rows.push(row);

            // TODO: Only sort on save?
            rows.sort(function(row1,row2){
                return parseInt(row1.$.r) - parseInt(row2.$.r);
            });
        }

        var cellId = intToExcelColumn(colnum) + rownum;

        var cell = _.find(row.c, function (c) {
            return c.$.r == cellId;
        });

        if (!cell) {
            cell = {'$': {r: cellId}};
            row.c = row.c || [];
            row.c.push(cell);
        }

        setValue(cell, cellvalue);

        if ( copyCellStyleFrom ) {
            var refCell = getCellbyRef(copyCellStyleFrom);
            if( refCell ){
                cell.$.s = refCell.$.s
            }
        }

        // TODO: only sort on save?
        row.c.sort(function(cell1, cell2){
            return excelColumnToInt(cell1.$.r) - excelColumnToInt(cell2.$.r);
            //return cell1.$.r.localeCompare(cell2.$.r);
        });

    };

    function setValue(cell, cellvalue) {
        if (typeof cellvalue == 'number') {
            // number
            cell.v = [cellvalue];
            delete cell.f;
        } else if (cellvalue[0] == '=') {
            // function
            cell.f = [cellvalue];
        } else {
            // assume string
            var index = _.findIndex(sharedStrings, function (string) {
                return string.t == cellvalue;
            });

            if (index == -1) {
                sharedStrings = sharedStrings || [];
                sharedStrings.push({t: [cellvalue]});
                index = sharedStrings.length - 1;
            }

            cell.v = [index];
            cell.$.t = 's';
        }
    }

    this.getCellValue = function (rownum, colnum) {
        var rows = this.xml.worksheet.sheetData[0].row;
        var row = _.find(rows, function (r) {
            return r.$.r == rownum;
        });
        var cellId = intToExcelColumn(colnum) + rownum;

        if ( row && row.hasOwnProperty('c') ) {
            var cell = _.find(row.c, function (c) {
                return c.$.r == cellId;
            });

            return getValue(cell);
        } else {
            return undefined;
        }

    };

    function getValue(cell) {
        if (cell) {
            if (cell.$.t == 's') {
                return sharedStrings[parseInt(cell.v)].t[0];
            } else if (cell.f) {
                // formula
                return '=' + cell.f;
            } else {
                // other values
                return cell.v;
            }
        } else {
            return undefined;
        }
    }

    this.setRowValues = function (rownum, values) {
        var rows = this.xml.worksheet.sheetData[0].row;

        if ( !rows ) {
            rows = this.xml.worksheet.sheetData[0].row = [];
        }

        var row = _.find(rows, function (r) {
            return r.$.r == rownum;
        });

        if (!row) {
            row = {
                '$': {r: rownum}
            };
            rows.push(row);
            rows.sort(function(row1,row2){
                return parseInt(row1.$.r) - parseInt(row2.$.r);
            });
        }

        row.c = [];

        _.each(values, function(value, index){
            var cellId = intToExcelColumn(index+1) + rownum;
            var cell = {'$': {r: cellId}};
            row.c.push(cell);

            setValue(cell, value);
        });

        console.log(util.inspect( row, false, null ));

    };

    this.getRowValues = function (rownum) {
        var rows = this.xml.worksheet.sheetData[0].row;
        var row = _.find(rows, function (r) {
            return r.$.r == rownum;
        });

        if ( row && row.hasOwnProperty('c')) {
            return _.map(row.c, function(cell){
                return getValue(cell);
            });
        } else {
            return undefined;
        }

    };

}

function intToExcelColumn(col) {
    var result = '';

    var mod;

    while (col > 0) {
        mod = (col - 1) % 26;
        result = String.fromCharCode(65 + mod) + result;
        col = Math.floor((col - mod) / 26);
    }

    return result;
}

function excelColumnToInt(ref) {
    var number = 0;
    var pow = 1;
    for (var i = ref.length - 1; i >= 0; i--) {
        var c = ref.charCodeAt(i) - 64;
        if (c > 0 && c < 27) {
            number += (c) * pow;
            pow *= 26;
        }
    }

    return number;
}

exports.readXmlFile = function (basedir, file, sharedStrings, callback) {
    var readXmlFile = async.compose(parser.parseString, fs.readFile);
    readXmlFile(path.join(basedir, file), function (err, obj) {
        callback(null, new Sheet(basedir, file, obj, sharedStrings));
    });
};

exports.newXmlFile = function(basedir, file, contents, sharedStrings, callback) {
    var sheet = new Sheet(basedir, file, contents, sharedStrings);
    if ( callback ) {
        callback(sheet);
    }
    return sheet;
};