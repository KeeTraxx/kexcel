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
            console.log(kexcel.sharedStringsXml.xml.write());
            newsheet.replaceRow(12,{A12:'bla2'});
            kexcel.sheets[0].replaceRow(12, {A12: 'bla'});
            kexcel.sheets[0].replaceRow(13, {B13: 'blamore'});
            kexcel.sheets[0].setCellValue(14,3,'tester', 'C12');
            var output = fs.createWriteStream(__dirname + '/tester.xlsx');
            kexcel.pipe(output,function(){
                console.log('done!');
                kexcel.close();
            });
        }catch(e){
            console.log(e.stack);
        }
    });