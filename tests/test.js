var kexcel = require('../');
var fs = require('fs');
/*kexcel.open(__dirname + '/Mappe1.xlsx',function(err, workbook){
    //workbook.pipe(fs.createWriteStream('super.xlsx'));
});*/

kexcel.new(function(err, workbook){

    // Get first sheet
    var sheet1 = workbook.getSheet(0);

    // Duplicate a sheet
    var duplicatedSheet = workbook.duplicateSheet(0,'My duplicated sheet');

    // Add some data to the first sheet
    // Caution!! Row and column are 1-based
    sheet1.setCellValue(1,1,'foo in first row and first column');
    sheet1.setCellValue(5,1,'bar in fifth row and first column');
    sheet1.setCellValue(5,8,'Somewhere...');

    // Insert cell value, also copy style from another cell.
    //sheet1.setCellValue(6,1,'This cell copies the style from cell A1', 'A1');

    // Put random numbers in the duplicated sheet
    for(var r=1;r<100;r++) {
        for(var c=1;c<100;c++) {
            duplicatedSheet.setCellValue( r, c,  ~~(Math.random() * 300) );
        }
    }

    // Save the file
    var output = fs.createWriteStream(__dirname + '/tester.xlsx');
    workbook.pipe(output);
});

var express = require('express');

var app = express();

app.get('/', function (req, res) {
    kexcel.new(function(err, workbook) {
        var sheet = workbook.getSheet(0);
        sheet.setCellValue(1,1,'Hello World!');
        sheet.setRowValues(2, ['Hello', 'even', 'more', 'Worlds']);
        sheet.setRowValues(3, [1, '+', 2, 'equals','=A3+C3']);

        res.setHeader('Content-disposition', 'attachment; filename=myfile.xlsx');
        res.setHeader('Content-type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        workbook.pipe(res);

    });
});

var server = app.listen(3000, function () {

    var host = server.address().address
    var port = server.address().port

    console.log('Example app listening at http://%s:%s', host, port)

});