/**
 * Created by ktran on 24.03.14.
 */
var kexcel = require('./');

var fs = require('fs');

kexcel.open('export.xlsx').then(function(kexcelfile){
    console.log('wtf');
    kexcelfile.getSheets().then(console.log);
});
