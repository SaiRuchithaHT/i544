import { Result, okResult, errResult } from 'cs544-js-utils';

import { SpreadsheetDao } from './spreadsheet-dao.js';
import { Spreadsheet, makeSpreadsheet, Updates } from './spreadsheet.js';

export async function clear(ssDao: SpreadsheetDao)
  : Promise<Result<undefined>> 
{
  return await ssDao.clear();
}

export async function copy(ssDao: SpreadsheetDao,
			   destCellId: string, srcCellId: string)
  : Promise<Result<Updates>>
{
  const ssResult = await getSpreadsheet(ssDao);
  if (!ssResult.isOk) return ssResult;
  const ss = ssResult.val;
  const copyRet = ss.copy(destCellId, srcCellId);
  const srcExpr = ss.query(srcCellId).expr;
  const destExpr = ss.query(destCellId).expr;
  const result = (srcExpr)
    ? await ssDao.setCellExpr(destCellId, destExpr)
    : await ssDao.remove(destCellId);
  if (!result.isOk) return result;
  return copyRet;
}

export async function remove(ssDao: SpreadsheetDao, cellId: string)
  : Promise<Result<Updates>>
{
  const ssResult = await getSpreadsheet(ssDao);
  if (!ssResult.isOk) return ssResult;
  const ss = ssResult.val;
  const delRet = ss.remove(cellId);
  const result = await ssDao.remove(cellId);
  if (!result.isOk) return result;
  return delRet;
}


export async function evaluate(ssDao: SpreadsheetDao,
			       cellId: string, expr: string)
  : Promise<Result<Updates>>
{
  const ssResult = await getSpreadsheet(ssDao);
  if (!ssResult.isOk) return ssResult;
  const ss = ssResult.val;
  const evalResult = ss.eval(cellId, expr);
  if (!evalResult.isOk) return evalResult;
  const setResult = await ssDao.setCellExpr(cellId, expr);
  if (!setResult.isOk) return setResult;
  return evalResult;
}

export async function dump(ssDao: SpreadsheetDao) 
  : Promise<Result<[string, string][]>>
{
  const ssResult = await getSpreadsheet(ssDao);
  if (!ssResult.isOk) return ssResult;
  const ss = ssResult.val;
  return okResult(ss.dump());
}

export async function load(ssDao: SpreadsheetDao,
			   dump: [string, string][])
  : Promise<Result<undefined>>
{
  //validate
  const ssResult = await getSpreadsheet(ssDao);
  if (!ssResult.isOk) return ssResult;
  const ss = ssResult.val;
  ss.clear();
  for (const [cellId, expr] of dump) {
    const evalResult = ss.eval(cellId, expr);
    if (!evalResult.isOk) return evalResult;
  }

  //ideally this should be in some kind of transaction
  ssDao.clear();
  for (const [ cellId, expr ] of dump) {
    const setResult = await ssDao.setCellExpr(cellId, expr);
    if (!setResult.isOk) return setResult;
  }
  
  return okResult(undefined);
}

export async function query(ssDao: SpreadsheetDao, cellId: string)
  : Promise<Result<{ value: number, expr: string}>>
{
  const ssResult = await getSpreadsheet(ssDao);
  if (!ssResult.isOk) return ssResult;
  const ss = ssResult.val;
  return okResult(ss.query(cellId));
}

async function getSpreadsheet(ssDao: SpreadsheetDao)
  : Promise<Result<Spreadsheet>>
{
  const ss = makeSpreadsheet(ssDao.getSpreadsheetName());
  const dataResult = await ssDao.getData();
  if (!dataResult.isOk) return dataResult;
  for (const [cellId, expr] of dataResult.val) {
    ss.eval(cellId, expr);
  }
  return okResult(ss);
}
