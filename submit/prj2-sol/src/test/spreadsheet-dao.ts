//will run the project DAO using an in-memory mongodb server
import MemSpreadsheetDao from './spreadsheet-mem-dao.js';

import { makeSpreadsheetDao, SpreadsheetDao } from '../lib/spreadsheet-dao.js';


import { assert, expect } from 'chai';

const TEST_SPREADSHEET_NAME = 'test';

describe('spreadsheet DAO', () => {
 
  let dao: SpreadsheetDao;

  beforeEach(async () => {
    dao = await MemSpreadsheetDao.setup(TEST_SPREADSHEET_NAME);
  });

  afterEach(async () => {
    await MemSpreadsheetDao.tearDown(dao);
  });

  it('must return spreadsheet name', () => {
    expect(dao.getSpreadsheetName()).to.equal(TEST_SPREADSHEET_NAME);
  });

  it('must query an empty cell', async () => {
    const cellId = 'x2';
    const queryResult = await dao.query(cellId);
    assert(queryResult.isOk === true);
    expect(queryResult.val).to.equal('');
  });

  it('must set cell expression', async () => {
    const [cellId, expr] = ['a1', 'a2 * 3'];
    const setResult = await dao.setCellExpr(cellId, expr);
    assert(setResult.isOk === true);
    const queryResult = await dao.query(cellId);
    assert(queryResult.isOk === true);
    expect(queryResult.val).to.equal(expr);
  });

  
  it('must set multiple cell expressions', async () => {
    const cellExprs = [
      ['a1', 'a2 * 3'],
      ['a2', 'a3 + 4'],
      ['a3', 'd5 / 6'],
      ['d5', '42' ],
    ];
    for (const [cellId, expr] of cellExprs) {
      const setResult = await dao.setCellExpr(cellId, expr);
      assert(setResult.isOk === true);
    }
    for (const [cellId, expr] of cellExprs) {
      const queryResult = await dao.query(cellId);
      assert(queryResult.isOk === true);
      expect(queryResult.val).to.equal(expr);
    }
  });

  it('must clear spreadsheet', async () => {
    //TODO
  });
  
  it('must remove cells from spreadsheet', async () => {
    const cellExprs = [
      ['a1', 'a2 * 3'],
      ['a2', 'a3 + 4'],
      ['a3', 'd5 / 6'],
      ['d5', '42' ],
    ];
    for (const [cellId, expr] of cellExprs) {
      const setResult = await dao.setCellExpr(cellId, expr);
      assert(setResult.isOk === true);
    }
    const rmCellIds = [ 'a1', 'd5', 'a8' ];
    for (const cellId of rmCellIds) {
      const rmResult = await dao.remove(cellId);
      assert(rmResult.isOk === true);
    }
    for (const [cellId, expr] of cellExprs) {
      const queryResult = await dao.query(cellId);
      assert(queryResult.isOk === true);
      const expected = rmCellIds.includes(cellId) ? '' : expr;
      expect(queryResult.val).to.equal(expected);
    }
  });

  it('must get data from empty spreadsheet', async () => {
    const dataResult = await dao.getData();
    assert(dataResult.isOk === true);
    expect(dataResult.val).to.deep.equal([]);
  });

  it('must get data from spreadsheet', async () => {
    const cellExprs = [
      ['a1', 'a2 * 3'],
      ['a2', 'a3 + 4'],
      ['a3', 'd5 / 6'],
      ['d5', '42' ],
    ];
    for (const [cellId, expr] of cellExprs) {
      const setResult = await dao.setCellExpr(cellId, expr);
      assert(setResult.isOk === true);
    }
    const dataResult = await dao.getData();
    assert(dataResult.isOk === true);
    const cellSort =
      (a: [string, string], b: [string, string]) => a[0].localeCompare(b[0]);
    expect(dataResult.val.sort(cellSort)).to.deep.equal(cellExprs);
  });

});


describe('verify actual spreadsheet dao persistence', () => {

  const dbUrl = 'mongodb://localhost:27017/spreadsheets';

  async function makeDao() {
    const makeResult = await makeSpreadsheetDao(dbUrl, TEST_SPREADSHEET_NAME);
    assert(makeResult.isOk === true);
    return makeResult.val;
  }

  beforeEach(async () => {
    const dao = await makeDao();
    const clearResult = await dao.clear();
    assert(clearResult.isOk === true);
    const closeResult = await dao.close();
    assert(closeResult.isOk === true);
  });

  it('must set cell expression', async () => {
    const dao1 = await makeDao();
    const [cellId, expr] = ['a1', 'a2 * 3'];
    const setResult = await dao1.setCellExpr(cellId, expr);
    assert(setResult.isOk === true);
    const closeResult1 = await dao1.close();
    assert(closeResult1.isOk === true);

    const dao2 = await makeDao();
    const queryResult = await dao2.query(cellId);
    assert(queryResult.isOk === true);
    expect(queryResult.val).to.equal(expr);
    const closeResult2 = await dao2.close();
    assert(closeResult2.isOk === true);
  });

  
  it('must set multiple cell expressions', async () => {
    const cellExprs = [
      ['a1', 'a2 * 3'],
      ['a2', 'a3 + 4'],
      ['a3', 'd5 / 6'],
      ['d5', '42' ],
    ];
    const dao1 = await makeDao();
    for (const [cellId, expr] of cellExprs) {
      const setResult = await dao1.setCellExpr(cellId, expr);
      assert(setResult.isOk === true);
    }
    const closeResult1 = await dao1.close();
    assert(closeResult1.isOk === true);

    const dao2 = await makeDao();
    for (const [cellId, expr] of cellExprs) {
      const queryResult = await dao2.query(cellId);
      assert(queryResult.isOk === true);
      expect(queryResult.val).to.equal(expr);
    }
    const closeResult2 = await dao2.close();
    assert(closeResult2.isOk === true);
  });
  
});

describe('must catch bad spreadsheet creation', () => {

  it('must catch badly formed db url', async () => {
    const daoResult = await makeSpreadsheetDao('localhost:27017/', 'test');
    assert(daoResult.isOk === false);
    expect(daoResult.errors.length).to.be.above(0);
    expect(daoResult.errors[0].options.code).to.equal('DB');
  });
  
});
