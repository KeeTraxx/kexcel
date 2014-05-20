/**
 * Created by ktran on 24.03.14.
 */
var kexcel = require('./');

var fs = require('fs');
var Q = require('q');

Q.nfcall( kexcel.open, 'templates/empty.xlsx')
    .then(function(kexcel){
        try{
            var newsheet = kexcel.duplicateSheet(kexcel.sheets[0],'testduplicate');

            for(var r=1;r<100;r++) {
                for(var c=1;c<100;c++) {
                    newsheet.setCellValue( r, c,  ~~(Math.random() * 300) );
                }
            }

            kexcel.sheets[0].setCellValue(13,3,'tester');
            kexcel.sheets[0].setCellValue(13,5,'tester2');
            kexcel.sheets[0].setCellValue(14,3,'tester');

            kexcel.deleteSheetAt(0);

            var output = fs.createWriteStream(__dirname + '/tester.xlsx');
            kexcel.pipe(output,function(){
                console.log('done!');
                kexcel.close();
            });
        }catch(e){
            console.log(e.stack);
        }
    });