Make Excel
=========

An extension of exceljs to create an Excel(xlsx) file from a simple json.

## Installation

  `npm install make-excel`

## Usage

* ###  alphaToNum
    converts Alphabets to Numbers
    Eg: 'BD' to 55
    - @param {string} alpha
    - @return {number}
* ###  numToAlpha
    Converts Number to Alphabets
    Eg: 55 to 'BD'
    - @param {number} num
    - @return {string}
* ### createWorkbook
    Create and return workbook
    - @return {Workbook}
    
    Eg:     
    ```
    const workbook = ExcelHelper.createWorkbook();
    const worksheet = ExcelHelper.addWorksheet(workbook, 'Test1');
    ```
* ### addWorksheet
    Adds a worksheet to the workbook with a given name 
    - @param {Workbook} workbook
    - @param {string} name
    - @return {Worksheet}
   
    Eg:
    ```
    const worksheet = ExcelHelper.addWorksheet(workbook, 'Test1');
    ```
* ### getNextColumn
    Get the next column to the current one
    Optional: Use the second argument to skip to a column n steps ahead of the
    given column
    - @param {string} currentColumn
    - @param {number} steps
    - @return {string}
    
    Eg: 
    ```
    const nextColumn = getNextColumn('C');
    //nextColumn = D
    const nextColumn = getNextColumn('F', 3);
    //nextColumn = I
    ```
* ### addCellBorder
    Change or add a border to a cell
    TODO: Add a better addCellborder function.
    - @param {Cell} cell
    - @param {object} border
    - @return {string}
   
    Eg: 
    ```
    const cell = worksheet.getCell('C3')
    addCellBorder(cell, { top: { style:thin } }) 
    // would add a top border to a cell or overwrite it.
    ```
* ### createOuterBorder
    Create an outer border to a given range of cells. 
    The start and the end objects should be of the following format
       { column:'B', row:'5' }
    The border width can be any of the following 'medium', 'thick', 'thin'
    - @param {object} start
    - @param {object} end
    - @param {Worksheet} worksheet
    - @param {string} borderWidth

    Eg: 
    ```
    const worksheet = ExcelHelper.addWorksheet(workbook, 'Test1');
    createOuterBorder({ column:'B', row:'5' }, { column:'F', row:'19' }, worksheet,'medium')
    //will create a outer border along the edge of these range of cells
    ```
* ### createSheetFromArray
    This function(createSheetFromArray) writes rows of data into a spreadsheet of an excel      file.
    Each row in rows contains an array of cellData holding the value and
    properties of each cell.
    The function also takes start row, start column, worksheet you want to fill
    arguments.
    The Rows Object should have the following structure
    ```
    [
        [cellData1, cellData2,...], //single row
        [cellData1, cellData2,..., skipRows: number] //optional skipRows property
    ]
    ```
    The skipRows attribute is used to skip a number of rows and continue
    writing from there.
    The object Cell Data that you can pass through should be of the form
    ```
    cellData = {
       value,
       border: {
                top: { style:'thin/thick/medium' },
                left: { style:'thin/thick/medium' },
                bottom: { style:'thin/thick/medium' },
                right: { style:'thin/thick/medium' }
               }
       font: { name: 'Arial', color: {argb: 'FF00FF00'}, size: 14, italic: true},
       alignment: { vertical: 'top/bottom,', horizontal: 'left/center/right'},
       fill: { type: 'pattern', pattern:'darkVertical', fgColor:{argb:'FFFF0000'}
      }
      //optional
      columnWidth: number,
      skipToColumn: column,
      columnOffset: column,
      mergeNumber: number
    }
    ```
    - The columnWidth attribute is used to give width to the entire column of the cell.
    - The skipToColumn attribute is used to skip to a pirtucular column and continue writing from there.
    - The columnOffset attribute is used to shift the current column and continue writing from there.
    - The mergerNumber attribute will merge the current cell till the number specified.
## Tests

  `npm test`

## Contributing

In lieu of a formal style guide, take care to maintain the existing coding style. Add unit tests for any new or changed functionality. Lint and test your code.