# kexcel

Manipulate .xlsx files - easy.

[![Build Status](https://travis-ci.org/KeeTraxx/kexcel.svg?branch=master)](https://travis-ci.org/KeeTraxx/kexcel) 
[![Coverage Status](https://coveralls.io/repos/KeeTraxx/kexcel/badge.svg?branch=master&service=github)](https://coveralls.io/github/KeeTraxx/kexcel?branch=master)

## Notes and information

## Installation
    npm install kexcel --save

### What KExcel supports
 * Reading .xlsx files from a file or a stream
 * Pure javascript, works on all operating systems.
 * Read sheet data and get values from:
   * Cells
   * Rows
   * Whole sheet
   
 * Supports creating empty files and sheets
   * Set strings / numbers or formulas to cells
   * Copy contents from other sheets
   * Copy cell styles from other cells
   
### What KExcel does not (yet?)
 * Calculations - you may set formulas, but they are only computed the next time you open the file in an spreadsheet editor (e.g. Excel or LibreOffice Calc)
 * Styles -  Setting custom styles
 * Support for other spreadsheet cell types (e.g. date, hours, etc.) 

## TL;DR Simple Usage Examples

### Create a new .xlsx file from scratch

````javascript
var fs = require('fs');
var kexcel = require('kexcel');
kexcel.new().then(function (wb) {
  var sheet = wb.getSheet(0);
  sheet.setCellValue(1, 1, 'Hello world!');
  return wb.pipe(fs.createWriteStream('output.xlsx'));
});
````

### Modify an existing .xlsx file and send it through a http response (express)
````javascript
var path = require('path');
var express = require('express');
var app = express();
var kexcel = require('kexcel');

app.get('/', function (req, res) {
    kexcel.open(path.join(__dirname, 'example.xlsx')).then(function(workbook) {
        var sheet = workbook.getSheet(0);
        sheet.setCellValue(1,1,'Hello World!');
        sheet.setRow(2, ['Hello', 'even', 'more', 'Worlds']);
        sheet.setRow(3, [1, '+', 2, 'equals','=A3+C3']);

        res.setHeader('Content-disposition', 'attachment; filename=example.xlsx');
        res.setHeader('Content-type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        workbook.pipe(res);
    });
});

var server = app.listen(3000, function () {
    var host = server.address().address;
    var port = server.address().port;

    console.log('KExcel app listening at http://%s:%s', host, port);
});
````