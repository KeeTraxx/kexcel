const kexcel = require('./dist/kexcel.bundle');
const fs = require('fs');

kexcel.new().then(function(wb){
   console.log(wb);
   wb.getSheet(0).setCellValue(1,1,'test');
   wb.pipe(fs.createWriteStream('test.xlsx'));
});