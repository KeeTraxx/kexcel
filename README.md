# kexcel

Excel 2007+ file manipulator.

## Features

 * Read Excel 2007+ .xlsx files
   * Support for reading String and Number values from cells
 * Create / Modify / Write Excel 2007+ .xlsx files
   * Write Strings / Numbers / Formulas to cells
   * Copy/Duplicate existing sheets
   * Copy styles from other cells
   * No support for setting custom styles yet.

## Usage

### Create a new excel file

```javascript
var kexcel = require('kexcel');
var fs = require('fs');
kexcel.new( function(err, workbook) {

    // Get first sheet
    var sheet1 = workbook.getSheet(0);

    // Duplicate a sheet
    var duplicatedSheet = workbook.duplicateSheet(0,'My duplicated sheet');

    // Add some data to the first sheet
    // Caution!! Row and column are 1-based
    sheet1.setCellValue(1,1,'foo in first row and first column');
    sheet1.setCellValue(5,1,'bar in fifth row and first column');
    sheet1.setCellValue(5,8,'Somewhere...');

    // Insert cell value, also copy style from another cell.
    sheet1.setCellValue(6,1,'This cell copies the style from cell A1', 'A1');

    // Put random numbers in the duplicated sheet
    for(var r=1;r<100;r++) {
        for(var c=1;c<100;c++) {
            duplicatedSheet.setCellValue( r, c,  ~~(Math.random() * 300) );
        }
    }

    // Save the file
    var output = fs.createWriteStream(__dirname + '/tester.xlsx');
    workbook.pipe(output);
});
```

### Open an existing excel file

```javascript
var kexcel = require('kexcel');
var fs = require('fs');
kexcel.open( 'myspreadsheet.xlsx', function(err, workbook) {

    // Get first sheet
    var sheet1 = workbook.getSheet(0);

    // Duplicate a sheet
    var duplicatedSheet = workbook.duplicateSheet(0,'My duplicated sheet');

    // Save the file
    var output = fs.createWriteStream(__dirname + '/tester.xlsx');
    workbook.pipe(output);
});
```