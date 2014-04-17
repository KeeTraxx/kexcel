kexcel
======

kexcel.js lets you open, edit and stream .xlsx files.
The main purpose for kexcel is to generate .xlsx reports from .xlsx templates.

1. Open a pre-formatted/layouted excel file.
2. Fill in some data.
3. Save the resulting Workbook to a file / stream it to the browser.
4. ???
5. Profit!

Install
=======
    npm install kexcel

Use
===
    var kexcel = require('kexcel');

    kexcel.open('sheet.xlsx', function(err, workbook) {
      if(err) throw err;
        // do some stuff with workbook.
    });

Test
====
Run `npm test`

MIT License.

**Thanks to [all other contributors](https://github.com/keetraxx/kexcel/graphs/contributors).**