var kexcel = require('../');
var fs = require('fs');
var path = require('path');
var filepath = path.join(__dirname, '../examples/input', 'empty.xlsx');

kexcel.open(filepath, function (err, workbook) {
    try {
        // Get first sheet
        var sheet1 = workbook.sheets[0];

        // Duplicate a sheet
        var duplicatedSheet = workbook.duplicateSheet(sheet1, 'My duplicated sheet');

        // Add some data to the first sheet
        // Caution!! Row and column are 1-based
        sheet1.setCellValue(1, 1, 'foo in first row and first column');
        sheet1.setCellValue(5, 1, 'bar in fifth row and first column');
        sheet1.setCellValue(5, 8, 'Somewhere...');
        sheet1.setCellValue(5, 26, 'Z Column');
        sheet1.setCellValue(5, 27, 'AA Column');
        sheet1.setCellValue(6, 1, 'This cell copies the style from cell A1', 'A1');
        sheet1.setCellValue(7,1,'=HYPERLINK("http://www.google.ch","Google")');
        sheet1.setCellValue(8,1,'=1+1');

        sheet1.replaceRow(9,['a', 'b', 'c']);

        console.log(sheet1.getRowValues(5));

        console.log('Should print: Somewhere...', sheet1.getCellValue(5, 8));

        // Put random numbers in the duplicated sheet
        for (var r = 1; r < 100; r++) {
            for (var c = 1; c < 100; c++) {
                duplicatedSheet.setCellValue(r, c, ~~(Math.random() * 300));
            }
        }

        console.log('Random number: ', duplicatedSheet.getCellValue(5, 8));

        // Save the file
        var output = fs.createWriteStream(__dirname + '/output.xlsx');
        workbook.pipe(output, function () {
            console.log('done!');
            workbook.close();
        });

    } catch (e) {
        console.log(e.stack);
    }
});
