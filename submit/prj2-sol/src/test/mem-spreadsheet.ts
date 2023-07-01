import { makeSpreadsheet, Spreadsheet } from '../lib/spreadsheet.js';

import { assert, expect } from 'chai';

describe('in-memory spreadsheet', function() {

  let spreadsheet: Spreadsheet;

  beforeEach(() => {
    spreadsheet = makeSpreadsheet('test');
  });

  it ('must evaluate a single number formula', () => {
    const evalResult = spreadsheet.eval('a1', '22');
    assert(evalResult.isOk === true);
    expect(evalResult.val).to.deep.equal({ a1: 22 });
  });

  it ('must evaluate a purely numeric formula', () => {
    const evalResult = spreadsheet.eval('a1', '(1 + 2)*-3 + 4');
    assert(evalResult.isOk === true);
    expect(evalResult.val).to.deep.equal({ a1: -5 });
  });

  it ('must query an empty cell as 0 with empty formula', () => {
    const results = spreadsheet.query('a1');
    expect(results).to.deep.equal({ value: 0, expr: '' });
  });

  it ('must evaluate a formula with a single reference', () => {
    spreadsheet.eval('a1', '22');
    const evalResult = spreadsheet.eval('a2', 'a1');
    assert(evalResult.isOk === true);
    expect(evalResult.val).to.deep.equal({ a2: 22 });
  });

  it ('must evaluate a reference formula', () => {
    spreadsheet.eval('a1', '22');
    const evalResult = spreadsheet.eval('a2', 'a1 * a1 + a1');
    assert(evalResult.isOk === true);
    expect(evalResult.val).to.deep.equal({ a2: 22*22 + 22 });
  });

  it ('must evaluate an undefined cell as 0', () => {
    spreadsheet.eval('a1', '22');
    const evalResult = spreadsheet.eval('a2', 'a1 * b1');
    assert(evalResult.isOk === true);
    expect(evalResult.val).to.deep.equal({ a2: 0 });
  });

  it ('must cascade an update', () => {
    spreadsheet.eval('a1', '22');
    spreadsheet.eval('a2', 'a1 * b1');
    const evalResult = spreadsheet.eval('b1', '3');
    assert(evalResult.isOk === true);
    expect(evalResult.val).to.deep.equal({ b1: 3, a2: 66,  });
  });

  it ('must evaluate a multi-level formula', () => {
    spreadsheet.eval('a1', '22');
    spreadsheet.eval('a2', 'a1 * b1');
    spreadsheet.eval('b1', '3');
    const evalResult = spreadsheet.eval('a3', 'a1 + a2');
    assert(evalResult.isOk === true);
    expect(evalResult.val).to.deep.equal({ a3: 88,  });
  });

  it ('must cascade an update through multiple levels', () => {
    spreadsheet.eval('a1', '22');
    spreadsheet.eval('a2', 'a1 * b1');
    spreadsheet.eval('b1', '3');
    spreadsheet.eval('a3', 'a1 + a2');
    const evalResult = spreadsheet.eval('a1', '3');
    assert(evalResult.isOk === true);
    expect(evalResult.val).to.deep.equal({ a1: 3, a2: 9, a3: 12,   });
  });

  it ('must detect a syntax error', () => {
    const evalResult = spreadsheet.eval('a1', 'a1 ++ 1');
    assert(evalResult.isOk === false, 'expected syntax error');
    expect(evalResult.errors).to.have.length(1);
    expect(evalResult.errors[0].options.code).to.equal('SYNTAX');
  });
  
  it ('must detect a direct circular reference', () => {
    const evalResult = spreadsheet.eval('a1', 'a1 + 1');
    assert(evalResult.isOk === false, 'expected direct circular reference');
    expect(evalResult.errors).to.have.length(1);
    expect(evalResult.errors[0].options.code).to.equal('CIRCULAR_REF');
  });
  
  it ('must detect an indirect circular reference', () => {
    spreadsheet.eval('a1', '22');
    spreadsheet.eval('a2', 'a1 * b1');
    spreadsheet.eval('b1', '3');
    spreadsheet.eval('a3', 'a1 + a2');
    const evalResult = spreadsheet.eval('a1', 'a3 + 1');
    assert(evalResult.isOk === false, 'expected indirect circular reference');
    expect(evalResult.errors).to.have.length(1);
    expect(evalResult.errors[0].options.code).to.equal('CIRCULAR_REF');
  });

  it ('must recover from an error', function() {
    spreadsheet.eval('a1', '22');
    spreadsheet.eval('a2', 'a1 * b1');
    spreadsheet.eval('b1', '3');
    spreadsheet.eval('a3', 'a1 + a2');
    const evalResult1 = spreadsheet.eval('a1', 'a3 + 1'); 
    assert(evalResult1.isOk === false);
    expect(evalResult1.errors).to.have.length(1);
    expect(evalResult1.errors[0].options.code).to.equal('CIRCULAR_REF');
    const evalResult2 = spreadsheet.eval('a4', 'a1 + a3');
    assert(evalResult2.isOk === true);
    expect(evalResult2.val).to.deep.equal({ a4: 110,  });
  });

  it ('must copy formula with relative references', () => {
    const ss = spreadsheet;
    //c2: 2, d2: 4, e2: 6, f2: (c2 + d2)*e2
    const data = addData(ss, 4, 'c2', () => 0, () => 1,
			 i => (i < 3) ? String((i+1)*2) : '(c2 + d2)*e2');
    expect(ss.query('f2').value).to.equal(36);
    //e4: 3, f4: 6, g4: 9, h4: 12, i4: h4*2
    addData(ss, 5, 'e4', () => 0, () => 1,
	    i => (i < 4) ? String((i + 1)*3) : 'h4*2');
    expect(ss.query('i4').value).to.equal(24);
    const copyResult = ss.copy('h4', 'f2'); //h4 = (e4 + f4)*g4
    assert(copyResult.isOk === true);
    expect(copyResult.val).to.deep.equal({ h4: 81, i4: 162 });
  });

  it('must copy formula with relative/absolute references', () => {
    const ss = spreadsheet;
    //c2: 2, d2: 4, e2: 6, f2: (c2 + d2)*$e$2
    const formula = '(c2 + d2)*$e$2';
    const data = addData(ss, 4, 'c2', () => 0, () => 1,
			 i => (i < 3) ? String((i+1)*2) : formula);
    expect(ss.query('f2').value).to.equal(36);
    //e4: 3, f4: 6, g4: 9, h4: 12, i4: h4*2
    addData(ss, 5, 'e4', () => 0, () => 1,
	    i => (i < 4) ? String((i + 1)*3) : 'h4*2');
    expect(ss.query('i4').value).to.equal(24);
    const copyResult = ss.copy('h4', 'f2'); //h4 = (e4 + f4)*$e$2
    assert(copyResult.isOk === true);
    expect(copyResult.val).to.deep.equal({ h4: 54, i4: 108 });
  });

  it ('must detect circular references when copying', () => {
    const ss = spreadsheet;
    ss.eval('d2', '42');
    ss.eval('c1', '$d$2 + 1'); //43
    ss.eval('c2', 'c1*2');     //86
    ss.eval('c3', 'c2 + 1');   //87
    expect(ss.query('c3').value).to.equal(87);
    ss.copy('d1', 'c1');
    const copyResult = ss.copy('d2', 'c3');  //circular ref
    assert(copyResult.isOk === false);
    expect(copyResult.errors.length).to.be.above(0);
    expect(copyResult.errors[0].options.code).to.equal('CIRCULAR_REF');
  });

  it ('must cascade copy of an empty cell', () => {
    const ss = spreadsheet;
    //c3: 42, c4: 44, c5: 46, c6: 48, c7: 50
    const data =
      addData(ss, 5, 'c3', () => 1, () => 0, 
	      (i, x: CellInfo[]) => (i === 0) ? '42' : `${x[i-1].relRel}+2`);
    expect(ss.query('c7').value).to.equal(50);
    const copyResult = ss.copy('c3', 'x4');
    assert(copyResult.isOk === true);
    expect(copyResult.val)
      .to.deep.equal({ 'c4': 2, 'c5': 4, 'c6': 6, 'c7': 8 });
  });

  it ('must clear spreadsheet', () => {
    const ss = spreadsheet;
    const data = addData(ss, 10);
    ss.clear();
    const cellValuePairs = Object.keys(data).map(c => [c, (ss.query(c)).expr]);
    const results = Object.fromEntries(cellValuePairs);
    const expected = Object.fromEntries(Object.keys(data).map(k => [k, '']));
    expect(results).to.deep.equal(expected);
  });

  it ('must remove cells', () => {
    const ss = spreadsheet;
    const data = addData(ss, 10);
    for (const k of Object.keys(data)) { ss.remove(k); }
    const cellValuePairs = Object.keys(data).map(c => [c, (ss.query(c)).expr]);
    const results = Object.fromEntries(cellValuePairs);
    const expected = Object.fromEntries(Object.keys(data).map(k => [k, '']));
    expect(results).to.deep.equal(expected);
  });

  it ('must delete empty cells', () => {
    const ss = spreadsheet;
    const data = addData(ss, 10, 'b1');
    ss.remove('a1'); ss.remove('a10');
    const cellValuePairs = Object.keys(data).map(c => [c, (ss.query(c)).expr]);
    const results = Object.fromEntries(cellValuePairs);
    const expected = data;
    expect(results).to.deep.equal(expected);
  });

  it ('must delete cells with cascade', async function() {
    const ss = spreadsheet;
    //c3: 42, c4: 44, c5: 46, c6: 48, c7: 50
    const data =
      addData(ss, 5, 'c3', () => 1, () => 0, 
	      (i, x: CellInfo[]) => (i === 0) ? '42' : `${x[i-1].relRel}+2`);
    expect(ss.query('c7').value).to.equal(50);
    const rmResult = ss.remove('c3');
    assert(rmResult.isOk === true);
    expect(rmResult.val).to.deep.equal({ 'c4': 2, 'c5': 4, 'c6': 6, 'c7': 8 });
  });

  it ('must dump empty spreadsheet', () => {
    const ss = spreadsheet;
    const results = ss.dump();
    expect(results).to.deep.equal([]);
  });

  it ('must dump spreadsheet in lexical order', () => {
    const ss = spreadsheet;
    const data = addData(ss, 10);
    const results = ss.dump();
    expect(results).to.deep.equal(Object.entries(data));
  });

  it ('must dump spreadsheet in topological/lexical order', () => {
    spreadsheet.eval('a1', '22');
    spreadsheet.eval('a3', 'a1 * b1');
    spreadsheet.eval('b1', '3');
    spreadsheet.eval('a2', 'a1 + a3');
    const results = spreadsheet.dump();
    expect(results).to.deep.equal([
      [ 'a1', '22' ], [ 'b1', '3' ],
      [ 'a3', 'a1*b1' ],
      [ 'a2', 'a1+a3' ],
    ]);
  });
  
  it ('must cascade cell definitions across copy', () => {
    //corresponds to LOG
    const ss = spreadsheet;
    const evalResult1 = spreadsheet.eval('a1', 'b1 * 2 + 3');
    assert(evalResult1.isOk === true);
    expect(evalResult1.val).to.deep.equal({ a1: 3 });
    const evalResult2 = spreadsheet.eval('b1', '5');
    assert(evalResult2.isOk === true);
    expect(evalResult2.val).to.deep.equal({ a1: 13, b1: 5, });
    const copyResult = spreadsheet.copy('d3', 'a1'); //d3 = e3 * 2 + 3
    assert(copyResult.isOk === true);
    expect(copyResult.val).to.deep.equal({ d3: 3, });
    const evalResult3 = spreadsheet.eval('e3', 'e2 * 3');
    assert(evalResult3.isOk === true);
    expect(evalResult3.val).to.deep.equal({ e3: 0, d3: 3, });
    const evalResult4 = spreadsheet.eval('e2', '4');
    assert(evalResult4.isOk === true);
    expect(evalResult4.val).to.deep.equal({ e2: 4, e3: 12, d3: 27, });
    const dump = ss.dump();
    expect(dump).to.deep.equal([
      ['b1', '5'], ['e2', '4'], ['a1', 'b1*2+3'], ['e3', 'e2*3'],
      ['d3', 'e3*2+3']
    ]);
    const rmResult = ss.remove('e2');
    assert(rmResult.isOk === true);
    expect(rmResult.val).to.deep.equal({e3: 0, d3: 3});
  });

});	

type CellInfo = {
  relRel: string,
  relAbs: string,
  absRel: string,
  absAbs: string,
  formula: string,
};
  

/** Add nData formulas to spreadsheet ss starting at cell startCellId.
 *  Successive cells are computed by incrmenting row / col id's by
 *  rowIncFn() / colIncFn(). Formula for a cell is the result of
 *  calling formulaFn().  All functions are called passing in the
 *  index in [0, nData) and a map of previously determined cells.
 *
 *  Returns map from cell-ids to formulas.
 */
function addData(ss: Spreadsheet, nData: number,
		 startCellId= "a1",
		 rowIncFn = (i?:number, infos: CellInfo[]=[]) => 1,
		 colIncFn = (i?:number, infos: CellInfo[]=[]) => 1,
		 formulaFn = (i: number, infos?: CellInfo[]) => String(i + 2))
{
  console.assert(nData < 26);
  const data: { [cellId: string]: string } = {};
  console.assert(/^[a-zA-Z]\d+$/.test(startCellId));
  const a = 'a'.codePointAt(0)!;
  let colIndex = startCellId[0].toLowerCase().codePointAt(0)! - a;
  let rowIndex = Number(startCellId.slice(1)) - 1;
  const cellInfos: CellInfo[] = [];
  for (let i = 0; i < nData; i++) {
    const r = String(1 + rowIndex);
    const c = String.fromCodePoint(a + colIndex);
    const cellId = `${c}${r}`;
    const formula = formulaFn(i, cellInfos);
    rowIndex += rowIncFn(i, cellInfos);
    colIndex += colIncFn(i, cellInfos);
    cellInfos.push({
      relRel: cellId,
      relAbs: `${c}$${r}`,
      absRel: `$${c}${r}`,
      absAbs: `$${c}$${r}`,
      formula,
    });
    data[cellId] = formula;
  }
  for (const [k, v] of Object.entries(data)) { ss.eval(k, v); }
  return data;
}

