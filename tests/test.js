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