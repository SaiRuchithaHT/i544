import { scan, CellRef } from '../lib/expr-parser.js';

import { ObjMap } from '../lib/utils.js';

import { panic } from 'cs544-js-utils';

import { assert, expect } from 'chai';

describe('scan', function() {

  let CELL_A1: CellRef;
  let CELL_E9: CellRef;

  before(() => {
    const cellA1Result = CellRef.parse('a1');
    if (!cellA1Result.isOk) panic('cannot parse cell-ref "a1"');
    CELL_A1 = cellA1Result.val;
    const cellE9Result = CellRef.parse('e9');
    if (!cellE9Result.isOk) panic('cannot parse cell-ref "e9"');
    CELL_E9 = cellE9Result.val;
  });

  it ('an integer should scan correctly', () => {
    const tokensResult = scan(' 123 ', CELL_A1);
    assert(tokensResult.isOk === true);
    const tokens = tokensResult.val;
    expect(tokens.length).to.equal(2);
    expect(tokens[0].kind).to.equal('num');
    expect(tokens[0].lexeme).to.equal('123');
    expect((tokens[0]).value).to.equal(123);
  });

  it ('a number with decimal point should scan correctly', () => {
    const tokensResult = scan(' 1.23 ', CELL_A1);
    assert(tokensResult.isOk === true);
    const tokens = tokensResult.val;
    expect(tokens.length).to.equal(2);
    expect(tokens[0].kind).to.equal('num');
    expect(tokens[0].lexeme).to.equal('1.23');
    expect(tokens[0].value).to.equal(1.23);
  });

  it ('a number with exponent should scan correctly', () => {
    const tokensResult = scan(' 1.23e2 ', CELL_A1);
    assert(tokensResult.isOk === true);
    const tokens = tokensResult.val;
    expect(tokens.length).to.equal(2);
    expect(tokens[0].kind).to.equal('num');
    expect(tokens[0].lexeme).to.equal('1.23e2');
    expect(tokens[0].value).to.equal(123);
  });

  it ('a number with negative exponent should scan correctly', () => {
    const tokensResult = scan(' 123e-2 ', CELL_A1);
    assert(tokensResult.isOk === true);
    const tokens = tokensResult.val;
    expect(tokens.length).to.equal(2);
    expect(tokens[0].kind).to.equal('num');
    expect(tokens[0].lexeme).to.equal('123e-2');
    expect(tokens[0].value).to.equal(1.23);
  });

  it ('a rel/rel ref should scan correctly', () => {
    const tokensResult = scan(' b4 ', CELL_A1);
    assert(tokensResult.isOk === true);
    const tokens = tokensResult.val;
    expect(tokens.length).to.equal(2);
    expect(tokens[0].kind).to.equal('ref');
    expect(tokens[0].lexeme).to.equal('b4');
    expect(tokens[0].value).to.deep.equal({
      col: { isAbs: false, index: 1 },
      row: { isAbs: false, index: 3 },
    });
  });

  it ('a rel/rel ref relative to a base should scan correctly', () => {
    const tokensResult = scan(' C5 ', CELL_E9);
    assert(tokensResult.isOk === true);
    const tokens = tokensResult.val;
    expect(tokens.length).to.equal(2);
    expect(tokens[0].kind).to.equal('ref');
    expect(tokens[0].lexeme).to.equal('C5');
    expect(tokens[0].value).to.deep.equal({
      col: { isAbs: false, index: -2 },
      row: { isAbs: false, index: -4 },
    });
  });

  it ('a rel/abs ref relative to a base should scan correctly', () => {
    const tokensResult = scan(' c$5 ', CELL_E9);
    assert(tokensResult.isOk === true);
    const tokens = tokensResult.val;
    expect(tokens.length).to.equal(2);
    expect(tokens[0].kind).to.equal('ref');
    expect(tokens[0].lexeme).to.equal('c$5');
    expect(tokens[0].value).to.deep.equal({
      col: { isAbs: false, index: -2 },
      row: { isAbs: true, index: 4 },
    });
  });

  it ('an abs/abs ref relative to a base should scan correctly', () => {
    const tokensResult = scan(' $c$5 ',  CELL_E9);
    assert(tokensResult.isOk === true);
    const tokens = tokensResult.val;
    expect(tokens.length).to.equal(2);
    expect(tokens[0].kind).to.equal('ref');
    expect(tokens[0].lexeme).to.equal('$c$5');
    expect(tokens[0].value).to.deep.equal({
      col: { isAbs: true, index: 2 },
      row: { isAbs: true, index: 4 },
    });
  });

  it ('multiple tokens should scan correctly', () => {
    const tokensResult = scan('123e-2 + ( $A2 * e11 )', CELL_E9);
    assert(tokensResult.isOk === true);
    const tokens = tokensResult.val;
    expect(tokens.length).to.equal(8);
    expect(tokens.map(t => t.kind)).to.deep.equal(
      [ 'num', '+', '(', 'ref', '*', 'ref', ')', 'EOF' ]);
    expect(tokens[0].value).to.equal(1.23);
    expect(tokens[3].value).to.deep.equal( {
      col: { isAbs: true, index: 0 },
      row: { isAbs: false, index: -7 },
    });
    expect(tokens[5].value).to.deep.equal({
      col: { isAbs: false, index: 0 },
      row: { isAbs: false, index: 2 },
    });
  });

  it ('a bad ref should result in a syntax error', () => {
    const tokensResult = scan('123e-2 + $$a2', CELL_A1);
    assert(tokensResult.isOk === false);
    expect(tokensResult.errors).to.have.length(1);
    expect(tokensResult.errors[0].options.code).to.equal('SYNTAX');
  });
    

});	
