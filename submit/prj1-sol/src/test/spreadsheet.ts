import { default as makeSpreadsheet, Spreadsheet } from '../lib/spreadsheet.js';

import { assert, expect } from 'chai';

describe('spreadsheet', function() {

  let spreadsheet: Spreadsheet;

  beforeEach(async () => {
    const result = await makeSpreadsheet('test');
    assert(result.isOk === true);
    spreadsheet = result.val;
  });

  it ('must evaluate a single number formula', async () => {
    const evalResult = await spreadsheet.eval('a1', '22');
    assert(evalResult.isOk === true);
    expect(evalResult.val).to.deep.equal({ a1: 22 });
  });

  it ('must evaluate a purely numeric formula', async () => {
    const evalResult = await spreadsheet.eval('a1', '(1 + 2)*-3 + 4');
    assert(evalResult.isOk === true);
    expect(evalResult.val).to.deep.equal({ a1: -5 });
  });

  it ('must evaluate a formula with a single reference', async () => {
    await spreadsheet.eval('a1', '22');
    const evalResult = await spreadsheet.eval('a2', 'a1');
    assert(evalResult.isOk === true);
    expect(evalResult.val).to.deep.equal({ a2: 22 });
  });

  it ('must evaluate a reference formula', async () => {
    await spreadsheet.eval('a1', '22');
    const evalResult = await spreadsheet.eval('a2', 'a1 * a1 + a1');
    assert(evalResult.isOk === true);
    expect(evalResult.val).to.deep.equal({ a2: 22*22 + 22 });
  });

  it ('must evaluate an undefined cell as 0', async () => {
    await spreadsheet.eval('a1', '22');
    const evalResult = await spreadsheet.eval('a2', 'a1 * b1');
    assert(evalResult.isOk === true);
    expect(evalResult.val).to.deep.equal({ a2: 0 });
  });

  it ('must cascade an update', async () => {
    await spreadsheet.eval('a1', '22');
    await spreadsheet.eval('a2', 'a1 * b1');
    const evalResult = await spreadsheet.eval('b1', '3');
    assert(evalResult.isOk === true);
    expect(evalResult.val).to.deep.equal({ b1: 3, a2: 66,  });
  });

  it ('must evaluate a multi-level formula', async () => {
    await spreadsheet.eval('a1', '22');
    await spreadsheet.eval('a2', 'a1 * b1');
    await spreadsheet.eval('b1', '3');
    const evalResult = await spreadsheet.eval('a3', 'a1 + a2');
    assert(evalResult.isOk === true);
    expect(evalResult.val).to.deep.equal({ a3: 88,  });
  });

  it ('must cascade an update through multiple levels', async () => {
    await spreadsheet.eval('a1', '22');
    await spreadsheet.eval('a2', 'a1 * b1');
    await spreadsheet.eval('b1', '3');
    await spreadsheet.eval('a3', 'a1 + a2');
    const evalResult = await spreadsheet.eval('a1', '3');
    assert(evalResult.isOk === true);
    expect(evalResult.val).to.deep.equal({ a1: 3, a2: 9, a3: 12,   });
  });

  it ('must detect a syntax error', async () => {
    const evalResult = await spreadsheet.eval('a1', 'a1 ++ 1');
    assert(evalResult.isOk === false, 'expected syntax error');
    expect(evalResult.errors).to.have.length(1);
    expect(evalResult.errors[0].options.code).to.equal('SYNTAX');
  });
  
  it ('must detect a direct circular reference', async () => {
    const evalResult = await spreadsheet.eval('a1', 'a1 + 1');
    assert(evalResult.isOk === false, 'expected direct circular reference');
    expect(evalResult.errors).to.have.length(1);
    expect(evalResult.errors[0].options.code).to.equal('CIRCULAR_REF');
  });
  
  it ('must detect an indirect circular reference', async () => {
    await spreadsheet.eval('a1', '22');
    await spreadsheet.eval('a2', 'a1 * b1');
    await spreadsheet.eval('b1', '3');
    await spreadsheet.eval('a3', 'a1 + a2');
    const evalResult = await spreadsheet.eval('a1', 'a3 + 1');
    assert(evalResult.isOk === false, 'expected indirect circular reference');
    expect(evalResult.errors).to.have.length(1);
    expect(evalResult.errors[0].options.code).to.equal('CIRCULAR_REF');
  });

  it ('must recover from an error', async function() {
    await spreadsheet.eval('a1', '22');
    await spreadsheet.eval('a2', 'a1 * b1');
    await spreadsheet.eval('b1', '3');
    await spreadsheet.eval('a3', 'a1 + a2');
    const evalResult1 = await spreadsheet.eval('a1', 'a3 + 1'); 
    assert(evalResult1.isOk === false);
    expect(evalResult1.errors).to.have.length(1);
    expect(evalResult1.errors[0].options.code).to.equal('CIRCULAR_REF');
    const evalResult2 = await spreadsheet.eval('a4', 'a1 + a3');
    assert(evalResult2.isOk === true);
    expect(evalResult2.val).to.deep.equal({ a4: 110,  });
  });

});	
