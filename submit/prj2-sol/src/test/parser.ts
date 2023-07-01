import parse from '../lib/expr-parser.js';
import { CellRef } from '../lib/expr-parser.js';

import { panic } from 'cs544-js-utils';

import { assert, expect } from 'chai';

describe('parse', function() {

 
  let A1_CELL: CellRef;
  let A5_CELL: CellRef;

  before(() => {
    const cellA1Result = CellRef.parse('a1');
    if (!cellA1Result.isOk) panic('cannot parse cell-ref "a1"');
    A1_CELL = cellA1Result.val;
    const cellA5Result = CellRef.parse('a5');
    if (!cellA5Result.isOk) panic('cannot parse cell-ref "e9"');
    A5_CELL = cellA5Result.val;
  });

  
  it ('must parse simple arith expr correctly', function () {
    const astResult = parse('1 + 2*3', A1_CELL);
    assert(astResult.isOk === true);
    const ast = astResult.val;
    expect(ast.toText(A1_CELL)).to.equal('1+2*3');
  });

  it ('must parse simple prec expr correctly', function () {
    const astResult = parse('(  (1 + 2)*3 )', A1_CELL);
    assert(astResult.isOk === true);
    const ast = astResult.val;
    expect(ast.toText(A1_CELL)).to.equal('(1+2)*3');
  });

  it ('must parse assoc expr correctly', function () {
    const astResult = parse('(1 + 2 - 3 )', A1_CELL);
    assert(astResult.isOk === true);
    const ast = astResult.val;
    expect(ast.toText(A1_CELL)).to.equal('1+2-3');
  });

  it ('must parse paren-assoc expr correctly', function () {
    const astResult = parse('(1 + (2 - 3) )', A1_CELL);
    assert(astResult.isOk === true);
    const ast = astResult.val;
    expect(ast.toText(A1_CELL)).to.equal('1+(2-3)');
  });

  it ('must parse unary - expr correctly', function () {
    const astResult = parse('(--1 + (2))', A1_CELL);
    assert(astResult.isOk === true);
    const ast = astResult.val;
    expect(ast.toText(A1_CELL)).to.equal('--1+2');
  });

  it ('must parse function expr correctly', function () {
    const astResult = parse('(1 + max((2 + 3)*4, 5, 6))', A1_CELL);
    assert(astResult.isOk === true);
    const ast = astResult.val;
    assert.equal(ast.toText(A1_CELL), '1+max((2+3)*4, 5, 6)');
  });

  it ('must parse nested function expr correctly', function () {
    const astResult = parse('(1 + max((2 + 3)*4, min(5, 6)))', A1_CELL);
    assert(astResult.isOk === true);
    const ast = astResult.val;
    expect(ast.toText(A1_CELL)).to.equal('1+max((2+3)*4, min(5, 6))');
  });

  it ('must parse cell ref correctly', function () {
    const astResult = parse('c$1', A1_CELL);
    assert(astResult.isOk === true);
    const ast = astResult.val;
    expect(ast.toText(A1_CELL)).to.equal('c$1');
  });

  it ('must translate cell ref correctly', function () {
    const astResult = parse('c2', A1_CELL);
    assert(astResult.isOk === true);
    const ast = astResult.val;
    expect(ast.toText(A5_CELL)).to.equal('c6');
  });

  it ('must translate cell ref without translating abs', function () {
    const astResult = parse('f$2', A1_CELL);
    assert(astResult.isOk === true);
    const ast = astResult.val;
    expect(ast.toText(A5_CELL)).to.equal('f$2');
  });

  it ('must parse and translate complex formula', function () {
    const astResult = parse('((1 + F$2)*$b3)', A1_CELL);
    assert(astResult.isOk === true);
    const ast = astResult.val;
    expect(ast.toText(A5_CELL)).to.equal('(1+f$2)*$b7');
  });


});	
