/**
 * Created by ktran on 24.03.14.
 */
var kexcel = require('./');

var fs = require('fs');
var Q = require('q');

kexcel.open( 'export.xlsx', function(err, workbook) {
    try{
        var newsheet = workbook.duplicateSheet(workbook.sheets[0],'testduplicate');
        newsheet.replaceRow(12,{A12:'bla2'});
        workbook.sheets[0].replaceRow(12, {A12: 'bla'});
        workbook.sheets[0].replaceRow(13, {B13: 'blamore'});
        workbook.sheets[0].setCellValue(14,3,'tester', 'C12');
        var output = fs.createWriteStream(__dirname + '/tester.xlsx');
        workbook.pipe(output,function(){
            console.log('done!');
            workbook.close();
        });
    } catch(e){
        console.log(e.stack);
        process.exit(1);
    }
});