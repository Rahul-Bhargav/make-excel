'use strict';

const expect = require('chai').expect;
const ExcelHelper = require('../index');

describe('Create a worksheet', function() {
  it('should Create a worksheet with the given name', function() {
    const workbook = ExcelHelper.createWorkbook();
    const worksheet = ExcelHelper.addWorksheet(workbook, 'Test1');
    expect(worksheet.name).to.equal('Test1');
    expect(worksheet.lastRowWritten).to.equal(0);
  });
  it('should give error if name is not a string', function() {
    const workbook = ExcelHelper.createWorkbook();
    try {
      ExcelHelper.addWorksheet(workbook, 342);
    }
    catch (err) {
      expect(err.message).to.equal('Name provided is not a string');
    }
  });
});

describe('Alpha to num', function() {
  it('should Convert an alphabet to integer', function() {
    const num = ExcelHelper.alphaToNum('E');
    expect(num).to.equal(4);
  });
  it('should Convert an 2 letter word to integer', function() {
    const num = ExcelHelper.alphaToNum('Bd');
    expect(num).to.equal(55);
  });
  it('should return an error if the input type is not a string', function() {
    try {
      ExcelHelper.alphaToNum(353);
    }
    catch (err) {
      expect(err.message).to.equal('Argument provided is not a string');
    }
  });
});

describe('Alpha to num', function() {
  it('should Convert an integer to alphabet', function() {
    const alpha = ExcelHelper.numToAlpha(4);
    expect(alpha).to.equal('E');
  });
  it('should Convert an integer to alphabet', function() {
    const alpha = ExcelHelper.numToAlpha(55);
    expect(alpha).to.equal('BD');
  });
  it('should return an error if the input type is not a string', function() {
    try {
      ExcelHelper.numToAlpha([]);
    }
    catch (err) {
      expect(err.message).to.equal('Argument provided is not a number');
    }
  });
});

describe('Get next column', function() {
  it('should get the next column', function() {
    const nextCol = ExcelHelper.getNextColumn('C');
    expect(nextCol).to.equal('D');
  });
  it('should get the next column with appropriate skipstep', function() {
    const nextCol = ExcelHelper.getNextColumn('F', 3);
    expect(nextCol).to.equal('I');
  });
  it('should return an error if the input type is not a string', function() {
    try {
      ExcelHelper.getNextColumn([]);
    }
    catch (err) {
      expect(err.message).to.equal('Incorrect arguments provided');
    }
  });
  it('should return an error if the skipsteps is not a number', function() {
    try {
      ExcelHelper.getNextColumn('C', 'asd');
    }
    catch (err) {
      expect(err.message).to.equal('Incorrect arguments provided');
    }
  });
});

describe('Add cell border', function() {
  it('should add border to a cell', function() {
    const cell = { border: { top: { style: 'thin' } } };
    const formattedCell = ExcelHelper.addCellBorder(cell,
      {
        top: { style: 'medium' },
        left: { style: 'thin' }
      }
    );
    expect(formattedCell.border.top.style).to.equal('medium');
    expect(formattedCell.border.left.style).to.equal('thin');
  });
});

describe('Create Outer border ', function() {
  const workbook = ExcelHelper.createWorkbook();
  const worksheet = ExcelHelper.addWorksheet(workbook, 'testBorder');
  it('should create an outer border given a range of cells', function() {

    const start = {
      column: 'C',
      row: '5'
    };
    const end = {
      column: 'G',
      row: '9'
    };
    ExcelHelper.createOuterBorder(start, end, worksheet, 'medium');
    expect(worksheet.getCell('C5').border.top.style).to.equal('medium');
    expect(worksheet.getCell('C5').border.left.style).to.equal('medium');
    expect(worksheet.getCell('G9').border.bottom.style).to.equal('medium');
    expect(worksheet.getCell('G9').border.right.style).to.equal('medium');
    expect(worksheet.getCell('E5').border.top.style).to.equal('medium');
  });
  it('should throw an error if the range of cells is invalid', function() {
    try {
      const start = {
        column: 'C',
      };
      const end = {
        row: '9'
      };
      ExcelHelper.createOuterBorder(start, end, worksheet, 'medium');
    }
    catch (err) {
      expect(err.message).to.equal('Invalid start, end arguments');
    }
  });
  it('should return an error if the border width is not a string', function() {
    try {
      const start = {
        column: 'C',
        row: '5'
      };
      const end = {
        column: 'G',
        row: '9'
      };
      ExcelHelper.createOuterBorder(start, end, worksheet, 56);
    }
    catch (err) {
      expect(err.message).to.equal('A valid border width is not provided');
    }
  });
});