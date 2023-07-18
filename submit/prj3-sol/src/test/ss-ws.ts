import STATUS from 'http-status';

import { App, makeApp } from '../lib/ss-ws.js';

import supertest from 'supertest';

//will run the project DAO using an in-memory mongodb server
import MemSpreadsheetDao from './spreadsheet-mem-dao.js';

import { SpreadsheetDao, makeSpreadsheetServices } from 'cs544-prj2-sol';

import { assert, expect } from 'chai';

const BASE = '/spreadsheets';

const SS_NAME1 = 'test1';

describe('web services', () => {
  
  //mocha will run beforeEach() before each test to set up these variables
  let ws: ReturnType<typeof supertest>;
  let dao: SpreadsheetDao;
  
  beforeEach(async function () {
    dao = await MemSpreadsheetDao.setup();
    const ssServices = makeSpreadsheetServices(dao);
    const app: App = makeApp(ssServices, BASE);
    ws = supertest(app);
  });
	 
  //mocha runs this after each test; we use this to clean up the DAO.
  afterEach(async function () {
    await MemSpreadsheetDao.tearDown(dao);
  });

  
  describe('cells GET and PATCH web services', () => {
    
    it('must store cell', async () => {
      await setCells(SS_NAME1, [
	[ 'a1', { expr: '42', result: { a1: 42 } }, ],
      ]);
    });
    
    it('must query stored cell', async () => {
      const cellId = 'a1';
      await setCells(SS_NAME1, [
	[ cellId, { expr: '42', result: { a1: 42 } }, ],
      ]);
      await queryCells(SS_NAME1, [
	[ 'a1', { expr: '42', value: 42 } ],
      ]);
    });
    
    it('must propagate updates', async () => {
      await setCells(SS_NAME1, [
	[ 'a1', { expr: 'b1 + b2', result: { a1: 0 } }, ],
	[ 'b1', { expr: '42', result: { a1: 42, b1: 42 } }, ],
	[ 'b2', { expr: '2', result: { a1: 44, b2: 2 } }, ],
      ]);
    });

    it('must propagate complex updates', async () => {
      await setCells(SS_NAME1, [
	[ 'a1', { expr: '42', result: { a1: 42 } }, ],
	[ 'b1', { expr: 'a1', result: { b1: 42 } }, ],
	[ 'b2', { expr: 'a1 + b1', result: { b2: 84 } }, ],
	[ 'a1', { expr: '4', result: { a1: 4, b1: 4, b2: 8 } }, ],
      ]);
    });

    it('must copy relative formula', async () => {
      await setCells(SS_NAME1, [
	[ 'a1', { expr: '42', result: { a1: 42 } }, ],
	[ 'b1', { expr: 'a1', result: { b1: 42 } }, ],
	[ 'b2', { expr: 'a1 + b1', result: { b2: 84 } }, ],
      ]);
      const params = { srcCellId: 'b2' };
      const q = new URLSearchParams(params).toString();
      const url = `${BASE}/${SS_NAME1}/c2?${q}`; //b1 + c1
      const res = await
      ws.patch(url)
	.set('Content-Type', 'application/json')
	.send();
      expect(res.status === STATUS.OK);
      expect(res.body?.isOk).to.equal(true);
      const links = res.body?.links;
      expect(links?.self?.method).to.equal('PATCH');
      expect(links?.self?.href).to.equal(url);
      expect(res.body.result).to.deep.equal({ c2: 42, });
      await setCells(SS_NAME1, [
	[ 'c1', { expr: '4', result: { c1: 4, c2: 46 } }, ],
	[ 'a1', { expr: '2', result: { a1: 2, b1: 2, b2: 4, c2: 6 } }, ],
      ]);
    });


    it('must error PATCH request without query parameters', async () => {
      const url = `${BASE}/${SS_NAME1}/a1`;
      const res1 = await
      ws.patch(url)
	.set('Content-Type', 'application/json')
	.send();
      expect(res1.status === STATUS.BAD_REQUEST);
      expect(res1.body?.isOk === false);
      expect(res1.body?.errors).to.have.length(1);
      expect(res1.body?.errors[0].options?.code).to.equal('BAD_REQ');
    });
    
    it('must error PATCH request with both query parameters', async () => {
      const params = { expr: '42', srcCellId: 'b1' };
      const q = new URLSearchParams(params).toString();
      const url = `${BASE}/${SS_NAME1}/a1?${q}`;
      const res1 = await
      ws.patch(url)
	.set('Content-Type', 'application/json')
	.send();
      expect(res1.status === STATUS.BAD_REQUEST);
      expect(res1.body?.isOk === false);
      expect(res1.body?.errors).to.have.length(1);
      expect(res1.body?.errors[0].options?.code).to.equal('BAD_REQ');
    });
    
    it('must error bad cell id', async () => {
      const params = { expr: '42' };
      const q = new URLSearchParams(params).toString();
      const url = `${BASE}/${SS_NAME1}/aa1?${q}`;
      const res1 = await
      ws.patch(url)
	.set('Content-Type', 'application/json')
	.send();
      expect(res1.status === STATUS.BAD_REQUEST);
      expect(res1.body?.isOk === false);
      expect(res1.body?.errors).to.have.length(1);
      expect(res1.body?.errors[0].options?.code).to.equal('SYNTAX');
    });
    
    
    it('must error bad cell formula', async () => {
      const params = { expr: '42 +' };
      const q = new URLSearchParams(params).toString();
      const url = `${BASE}/${SS_NAME1}/a1?${q}`;
      const res1 = await
      ws.patch(url)
	.set('Content-Type', 'application/json')
	.send();
      expect(res1.status === STATUS.BAD_REQUEST);
      expect(res1.body?.isOk === false);
      expect(res1.body?.errors).to.have.length(1);
      expect(res1.body?.errors[0].options?.code).to.equal('SYNTAX');
    });
    
    it('must error direct circular ref', async () => {
      const params = { expr: 'a1' };
      const q = new URLSearchParams(params).toString();
      const url = `${BASE}/${SS_NAME1}/a1?${q}`;
      const res1 = await
      ws.patch(url)
	.set('Content-Type', 'application/json')
	.send();
      expect(res1.status === STATUS.BAD_REQUEST);
      expect(res1.body?.isOk === false);
      expect(res1.body?.errors).to.have.length(1);
      expect(res1.body?.errors[0].options?.code).to.equal('CIRCULAR_REF');
    });

    it('must error indirect circular ref', async () => {
      await setCells(SS_NAME1, [
	[ 'a1', { expr: 'b1 + b2', result: { a1: 0 } }, ],
      ]);
      const params = { expr: 'a1* c1' };
      const q = new URLSearchParams(params).toString();
      const url = `${BASE}/${SS_NAME1}/b1?${q}`;
      const res1 = await
      ws.patch(url)
	.set('Content-Type', 'application/json')
	.send();
      expect(res1.status === STATUS.BAD_REQUEST);
      expect(res1.body?.isOk === false);
      expect(res1.body?.errors).to.have.length(1);
      expect(res1.body?.errors[0].options?.code).to.equal('CIRCULAR_REF');
    });

  });

  describe('cells DELETE web service', () => {

    it('must remove cell', async () => {
      await setCells(SS_NAME1, [
	[ 'a1', { expr: '42', result: { a1: 42 } }, ],
	[ 'b1', { expr: 'a1', result: { b1: 42 } }, ],
	[ 'b2', { expr: 'b1', result: { b2: 42 } }, ],
      ]);
      const url = `${BASE}/${SS_NAME1}/a1`;
      const res = await ws.delete(url);
      expect(res.status === STATUS.OK);
      expect(res.body?.isOk).to.equal(true);
      const links = res.body?.links;
      expect(links?.self?.method).to.equal('DELETE');
      expect(links?.self?.href).to.equal(url);
      expect(res.body.result).to.deep.equal({b1: 0, b2: 0});
    });

    it('removing a cell should not affect dependent cell formula', async () => {
      await setCells(SS_NAME1, [
	[ 'a1', { expr: '42', result: { a1: 42 } }, ],
	[ 'b1', { expr: 'a1', result: { b1: 42 } }, ],
	[ 'b2', { expr: 'b1', result: { b2: 42 } }, ],
      ]);
      const url = `${BASE}/${SS_NAME1}/a1`;
      const res = await ws.delete(url);
      expect(res.status === STATUS.OK);
      expect(res.body?.isOk).to.equal(true);
      const links = res.body?.links;
      expect(links?.self?.method).to.equal('DELETE');
      expect(links?.self?.href).to.equal(url);
      expect(res.body.result).to.deep.equal({b1: 0, b2: 0});
      await queryCells(SS_NAME1, [
	[ 'a1', { expr: '', value: 0, } ],
	[ 'b1', { expr: 'a1', value: 0, } ],
	[ 'b2', { expr: 'b1', value: 0, } ],
      ]);
    });

    
  });

  describe('spreadsheet DELETE web service', () => {

    it('all cells disappear after spreadsheet DELETE', async () => {
      await setCells(SS_NAME1, [
	[ 'a1', { expr: '42', result: { a1: 42 } }, ],
	[ 'b1', { expr: 'a1', result: { b1: 42 } }, ],
	[ 'b2', { expr: 'b1 + a1', result: { b2: 84 } }, ],
      ]);
      const url = `${BASE}/${SS_NAME1}`;
      const res = await ws.delete(url);
      expect(res.status === STATUS.OK);
      expect(res.body?.isOk).to.equal(true);
      const links = res.body?.links;
      expect(links?.self?.method).to.equal('DELETE');
      expect(links?.self?.href).to.equal(url);
      expect(res.body?.result).to.be.undefined;
      await queryCells(SS_NAME1, [
	[ 'a1', { expr: '', value: 0, } ],
	[ 'b1', { expr: '', value: 0, } ],
	[ 'b2', { expr: '', value: 0, } ],
	[ 'a2', { expr: '', value: 0, } ],
      ]);
    });
    
    it('deleting an unknown spreadsheet is okay', async () => {
      const url = `${BASE}/${SS_NAME1}`;
      const res = await ws.delete(url);
      expect(res.status === STATUS.OK);
      expect(res.body?.isOk).to.equal(true);
      const links = res.body?.links;
      expect(links?.self?.method).to.equal('DELETE');
      expect(links?.self?.href).to.equal(url);
      expect(res.body?.result).to.be.undefined;
      await queryCells(SS_NAME1, [
	[ 'a1', { expr: '', value: 0, } ],
	[ 'b1', { expr: '', value: 0, } ],
	[ 'b2', { expr: '', value: 0, } ],
	[ 'a2', { expr: '', value: 0, } ],
      ]);
    });

  });
  
  describe('spreadsheet PUT web service', () => {

    it('loading a spreadsheet must read back data', async () => {
      const data: SSDump = [
	[ 'a1',  '42' ],
	[ 'b1', 'a1' ],
        [ 'b2', 'a1 + b1' ],
	[ 'a2',  '2' ],
        [ 'a3', '(a1 + a2) * a2' ],
      ];
      await loadSS(SS_NAME1, data);
      await queryCells(SS_NAME1, [
	[ 'a1',  { expr: '42', value: 42 } ],
	[ 'b1', { expr: 'a1', value: 42 } ],
        [ 'b2', { expr: 'a1+b1', value: 84 } ],
	[ 'a2', { expr: '2', value: 2 } ],
        [ 'a3', {expr: '(a1+a2)*a2', value: 88 } ],
      ]);
    });

    it('loading a spreadsheet must clear earlier data', async () => {
      await setCells(SS_NAME1, [
	[ 'c1', { expr: '42', result: { c1: 42 } } ],
	[ 'c2', { expr: 'c1', result: { c2: 42 } } ],
	[ 'c3', { expr: 'c1 + c2', result: { c3: 84 } } ],
      ]);
      const data: SSDump = [
	[ 'a1',  '42' ],
	[ 'b1', 'a1' ],
        [ 'b2', 'a1 + b1' ],
	[ 'a2',  '2' ],
        [ 'a3', '(a1 + a2) * a2' ],
      ];
      await loadSS(SS_NAME1, data);
      await queryCells(SS_NAME1, [
	[ 'a1', { expr: '42', value: 42 } ],
	[ 'b1', { expr: 'a1', value: 42 } ],
        [ 'b2', { expr: 'a1+b1', value: 84 } ],
	[ 'a2', { expr: '2', value: 2 } ],
        [ 'a3', { expr: '(a1+a2)*a2', value: 88 } ],
	[ 'c1', {expr: '', value: 0 } ], 
	[ 'c2', {expr: '', value: 0 } ], 
	[ 'c3', {expr: '', value: 0 } ], 
      ]);
    });

  });

  
  describe('spreadsheet GET web service', () => {

    it('reading a spreadsheet must read all current data', async () => {
      const evalTestData: CellEvalTest[] = [
	[ 'c1', { expr: '42', result: { c1: 42 } } ],
	[ 'c2', { expr: 'c1', result: { c2: 42 } } ],
	[ 'c3', { expr: 'c1 + c2', result: { c3: 84 } } ],
      ];
      await setCells(SS_NAME1, evalTestData);
      await readSS(SS_NAME1, evalTestToSSDump(evalTestData));
    });
  });
  
  
  /********************** Testing Utility Functions ********************/

  type CellEvalTest = [
    string,  //cellId
    {
      expr: string,
      result: { [cellId: string]: number },
    }
  ];

  async function setCells(ssName: string, updates: CellEvalTest[]) {
    for (const [cellId, {expr, result}] of updates) {
      const params = { expr };
      const q = new URLSearchParams(params).toString();
      const url = `${BASE}/${ssName}/${cellId}?${q}`;
      const res = await
        ws.patch(url)
	  .set('Content-Type', 'application/json')
	  .send();
      expect(res.status === STATUS.OK);
      expect(res.body?.isOk).to.equal(true);
      const links = res.body?.links;
      expect(links?.self?.method).to.equal('PATCH');
      expect(links?.self?.href).to.equal(url);
      expect(res.body.result).to.deep.equal(result);
    }
  }

  type CellQueryTest = [
    string, //cellId,
    { value: number, expr: string }
  ];

  async function queryCells(ssName: string, tests: CellQueryTest[]) {
    for (const [cellId, result] of tests) {
      const url = `${BASE}/${ssName}/${cellId}`;
      const res = await ws.get(url);
      expect(res.status === STATUS.OK);
      expect(res.body?.isOk).to.equal(true);
      const links = res.body?.links;
      expect(links?.self?.method).to.equal('GET');
      expect(links?.self?.href).to.equal(url);
      expect(res.body.result).to.deep.equal(result);
    }
  }

  type SSDump = [ string /* cellId */, string /* expr */ ][];
  
  async function loadSS(ssName: string, dump: SSDump) {
    const url = `${BASE}/${ssName}`;
      const res = await
        ws.put(url)
	  .set('Content-Type', 'application/json')
	  .send(dump);
    expect(res.status === STATUS.OK);
    expect(res.body?.isOk).to.equal(true);
    const links = res.body?.links;
    expect(links?.self?.method).to.equal('PUT');
    expect(links?.self?.href).to.equal(url);
    expect(res.body.result).to.be.undefined;
  }

  async function readSS(ssName: string, dump: SSDump) {
    const url = `${BASE}/${ssName}`;
    const res = await ws.get(url);
    expect(res.status === STATUS.OK);
    expect(res.body?.isOk).to.equal(true);
    const links = res.body?.links;
    expect(links?.self?.method).to.equal('GET');
    expect(links?.self?.href).to.equal(url);
    expect(res.body.result).to.deep.equal(dump);
  }

  //assumes non-destructive evalTestData in topological order
  //with no redundant paren in expr
  function evalTestToSSDump(evalTestData: CellEvalTest[]) : SSDump {
    return evalTestData
      .map(([cellId, {expr, result}]) => [cellId, expr.replace(/\s/g, '')]);
  }
  
});

