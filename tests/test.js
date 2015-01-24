var kexcel = require('../');
var fs = require('fs');
kexcel.open(__dirname + '/Mappe1.xlsx',function(err, workbook){
    workbook.pipe(fs.createWriteStream('super.xlsx'));
});

kexcel.new(function(err, workbook){

});