var kexcel = require('../');
var fs = require('fs');
/*kexcel.open(__dirname + '/Mappe1.xlsx',function(err, workbook){
    //workbook.pipe(fs.createWriteStream('super.xlsx'));
});*/

kexcel.open('tests/Mappe1.xlsx',function(err, workbook){
    workbook.getSheet(0).setCellValue(1,10,42);
    workbook.getSheet(0).setCellValue(1,11,'TEST');
    workbook.getSheet(0).setCellValue(1,12,'Test');
    workbook.getSheet(0).setCellValue(1,6,'Inserted');

    workbook.getSheet(0).setCellValue(3,3,'Copied style from A1', 'A1');

    console.log(workbook.getSheet(0).getCellValue(1,1));
    console.log(workbook.getSheet(0).getRowValues(1));


    workbook.getSheet(0).setRowValues(10,['A', 'b', 4242]);

    workbook.duplicateSheet(0);

    workbook.pipe(fs.createWriteStream('super.xlsx'));
});