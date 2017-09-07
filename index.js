/**
 * Make-Excel module.
 * @module make-excel
 */

/**
 * A Workbook instance.
 * @typedef {object} Workbook
 * 
 * A Worksheet instance.
 * @typedef {object} Worksheet
 * 
 * A Cell instance.
 * @typedef {object} Cell
 */

'use strict';
const Excel = require('exceljs');
const _ = require('lodash');

/**
 * converts Alphabets to Numbers
 * 
 * Eg: 'BD' to 55
 * @param {string} alpha
 * @return {number}
 */
exports.alphaToNum = (alpha) => {

  if (_.isString(alpha)) {
    let num = 0;
    const uppercaseAlpha = _.upperCase(alpha);
    const len = uppercaseAlpha.length;
    for (let i = 0; i < len; ++i) {
      num = num * 26 + uppercaseAlpha.charCodeAt(i) - 0x40;
    }
    return num - 1;
  }
  throw Error('Argument provided is not a string');
};

/**
 * Converts Number to Alphabets
 * 
 * Eg: 55 to 'BD'
 * @param {number} num
 * @return {string}
 */
exports.numToAlpha = (num) => {

  if (isNumber(num)) {
    let alpha = '';
    for (; num >= 0; num = parseInt(num / 26, 10) - 1) {
      alpha = String.fromCharCode(num % 26 + 0x41) + alpha;
    }
    return alpha;
  }
  throw Error('Argument provided is not a number');
};

exports.createWorkbook = () => new Excel.Workbook();

const isNumber = (num) => _.isNumber(num) && _.isFinite(num);

/**
 * Adds a worksheet to the workbook with a given name 
 * @param {Workbook} workbook
 * @param {string} name
 * @return {Worksheet}
 */
exports.addWorksheet = (workbook, name) => {
  if (_.isString(name)) {
    const worksheet = workbook.addWorksheet(name);
    worksheet.lastRowWritten = 0;
    return worksheet;
  }
  throw Error('Name provided is not a string');
};

/**
 * Get the next column to the current one
 * 
 * Eg: getNextColumn('C') would result in 'D'
 * 
 * Optional: Use the second argument to skip to a column n steps ahead of the
 * given column
 * @param {string} currentColumn
 * @param {number} steps
 * @return {string}
 */
exports.getNextColumn = (currentColumn, steps = 1) => {

  if (_.isString(currentColumn) && isNumber(steps)) {
    let columnNumber = exports.alphaToNum(currentColumn);
    columnNumber += steps;
    return exports.numToAlpha(columnNumber);
  }
  throw Error('Incorrect arguments provided');
};

/**
 * Change or add a border to a cell
 * Eg: addCellBorder(cell, { top: { style:thin } }) would add a top border or overwrite it.
 * TODO: Add a better addCellborder function.
 * @param {Cell} cell
 * @param {object} border
 * @return {string}
 */
exports.addCellBorder = (cell, border) => {

  if (!cell.border) {
    cell.border = border;
  }
  else {
    cell.border = Object.assign(cell.border, border);
  }
  return cell;
};

/**
 * Create an outer border to a given range of cells. 
 * The start and the end objects should be of the following format
 *    { column:'B', row:'5' }
 * The border width can be any of the following 'medium', 'thick', 'thin'
 * Eg: createOuterBorder({ column:'B', row:'5' }, { column:'F', row:'19' }, worksheet,'medium')
 * will create a outer border along the edge of these range of cells
 * TODO: Add a better addCellborder function.
 * @param {object} start
 * @param {object} end
 * @param {Worksheet} worksheet
 * @param {string} borderWidth
 */
exports.createOuterBorder = (start, end, worksheet, borderWidth = 'medium') => {


  if (!_.has(start, 'column') || !_.has(start, 'row') || !_.has(end, 'column') || !_.has(end, 'row')) {
    throw Error('Invalid start, end arguments');
  }
  if (!_.hasIn(worksheet, 'getCell')) {
    throw Error('A valid worksheet object is not provided');
  }
  if (!_.isString(borderWidth)) {
    throw Error('A valid border width is not provided');
  }
  const startColNumber = exports.alphaToNum(start.column);
  const endColNumber = exports.alphaToNum(end.column);
  const colRange = endColNumber - startColNumber + 1;
  const borderStyle = {
    style: borderWidth
  };
  for (let i = start.row; i <= end.row; ++i) {
    const leftBorderCell = worksheet.getCell(start.column + i);
    const rightBorderCell = worksheet.getCell(end.column + i);
    const leftBorder = { left: borderStyle };
    const rightBorder = { right: borderStyle };
    exports.addCellBorder(leftBorderCell, leftBorder);
    exports.addCellBorder(rightBorderCell, rightBorder);
  }

  for (let i = 0; i < colRange; ++i) {

    const currentColumn = exports.numToAlpha(startColNumber + i);
    const topBorderCell = worksheet.getCell(currentColumn + start.row);
    const bottomBorderCell = worksheet.getCell(currentColumn + end.row);
    const topBorder = { top: borderStyle };
    const bottomBorder = { bottom: borderStyle };
    exports.addCellBorder(topBorderCell, topBorder);
    exports.addCellBorder(bottomBorderCell, bottomBorder);
  }
};

/**
* This function(createSheetFromArray) writes rows of data into a spreadsheet of an excel file.
* Each row in rows contains an array of cellData holding the value and
* properties of each cell.
* The function also takes start row, start column, worksheet you want to fill
* arguments.
* The Rows Object should have the following structure
* [
*  [cellData1, cellData2,...] //single row
*  [cellData1, cellData2,..., skipRows: number] //optional skipRows property
*  .
*  .
*  .
*  .
* ]
* The skipRows attribute is used to skip a number of rows and continue
* writing from there.
* The object Cell Data that you can pass through should be of the form
* cellData = {
*    value,
*    border: {
*             top: { style:'thin/thick/medium' },
*             left: { style:'thin/thick/medium' },
*             bottom: { style:'thin/thick/medium' },
*             right: { style:'thin/thick/medium' }
*            }
*    font: { name: 'Arial', color: {argb: 'FF00FF00'}, size: 14, italic: true},
*    alignment: { vertical: 'top/bottom,', horizontal: 'left/center/right'},
*    fill: { type: 'pattern', pattern:'darkVertical', fgColor:{argb:'FFFF0000'}
*   }
*   //optional
*   columnWidth: number,
*   skipToColumn: column,
*   columnOffset: column,
*   mergeNumber: number
*  }
* The columnWidth attribute is used to give width to the entire column of the
* cell.
* The skipToColumn attribute is used to skip to a pirtucular column and continue
* writing from there.
* The columnOffset attribute is used to shift the current column and continue
* writing from there.
* The mergerNumber attribute will merge the current cell till the number
* specified.
*
*/
exports.createSheetFromArray = (worksheet, rows, rowOffset, startColumn, hasOuterBorder = false) => {

  let maxColumn = startColumn;
  let currentColumn = startColumn;
  let currentRow = worksheet.lastRowWritten + 1 + rowOffset;
  const startRow = currentRow;
  _.forEach(rows, (row) => {

    currentColumn = startColumn;
    if (!isNaN(row.skipRows)) {
      currentRow += row.skipRows;
    }
    _.forEach(row, (cellData) => {

      if (cellData.skipToColumn) {
        currentColumn = cellData.skipToColumn;
      }
      if (!isNaN(cellData.columnOffset)) {

        currentColumn = exports.getNextColumn(currentColumn, cellData.columnOffset);
      }

      const cell = worksheet.getCell(currentColumn + currentRow);

      if (!isNaN(cellData.mergeNumber)) {
        const rightMostColumn = exports.getNextColumn(currentColumn, cellData.mergeNumber);
        worksheet.mergeCells(`${currentColumn + currentRow}:${rightMostColumn + currentRow}`);
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        currentColumn = rightMostColumn;
      }
      if (!isNaN(cellData.columnWidth)) {
        worksheet.getColumn(currentColumn).width = cellData.columnWidth;
      }
      _.assignIn(cell, cellData);

      if (currentColumn > maxColumn) {
        maxColumn = currentColumn;
      }
      currentColumn = exports.getNextColumn(currentColumn);
    });
    currentRow++;
  });
  const start = {
    row: startRow,
    column: startColumn
  };
  const end = {
    row: --currentRow,
    column: maxColumn
  };
  if (hasOuterBorder) {
    exports.createOuterBorder(start, end, worksheet);
  }
  worksheet.lastRowWritten = currentRow;
};

exports.createCellData = (value = null) => {

  return {
    value
  };
};

exports.addCellStyle = (cell, font = {}, border = {}, alignment = { vertical: 'middle', horizontal: 'left' }) => {

  cell.font = font;
  cell.border = Object.assign({}, border);
  cell.alignment = alignment;
};

exports.addSolidFill = (cell, color) => {

  cell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: color }
  };
};
/**
* This function(editWidth) can be used to edit the width of multiple columns.
* the columns with their sizes should be sent as key pair value
*
*/
exports.editWidth = (columnsWithSize, worksheet) => {

  _.forEach(columnsWithSize, (size, column) => {

    worksheet.getColumn(column).width = size;
  });
};

exports.createXLSXFile = (workbook, fileName) => {

  return workbook.xlsx.writeFile(fileName);
};

exports.readXlSXFile = (fileName) => {

  const workbook = exports.createWorkbook();
  return workbook.xlsx.readFile(fileName)
    .then(() => {

      return workbook;
    });
};
