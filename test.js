/**
 * Created by ktran on 24.03.14.
 */
var kexcel = require('./');

var fs = require('fs');
var Q = require('q');

Q.nfcall( kexcel.open, 'export.xlsx')
    .then(function(kexcel){
        console.log( kexcel.sheets );
        console.log( kexcel.sheets[0].getName() );
        console.log( kexcel.sheets[0].getTree() );
        kexcel.sheets[0].replaceRow(12, {A12: 'bla'});
        console.log('a');
        kexcel.sheets[0].save();
        console.log('f');
    });