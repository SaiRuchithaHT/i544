import { Err, ErrResult, Result, okResult, errResult } from 'cs544-js-utils';

type Updates = { [cellId: string]: number };

export default class SpreadsheetWs {
  private readonly apiUrl: string;

  private constructor(url: string) { this.apiUrl = `${url}/api`; }

  static make(url: string)  {
    return new SpreadsheetWs(url);
  }

  /** return { expr, value } object for cell cellId in 
   *  spreadsheet ssName.
   */
  async query(ssName: string, cellId: string)
  : Promise<Result<{ value: number, expr: string}>>
  {
    const url = makeURL(`${this.apiUrl}/${ssName}/${cellId}`);
    return await doFetchJson('GET', url);
  }

  /** remove formula for cell cellId in spreadsheet ssName.
   *  Return Updates object mapping cellId's to the updated value.
   */
  async remove(ssName: string, cellId: string) : Promise<Result<Updates>>
  {
    const url = makeURL(`${this.apiUrl}/${ssName}/${cellId}`);
    return await doFetchJson('DELETE', url);
  }


  /** copy formula from cell srcCellId in spreadsheet ssName to
   *  cell destCellId.  Update relative references in the formula.
   *  If the formula is empty, then delete cell destCellId.
   *  Return Updates object mapping cellId's to the updated value.
   */
  async copy(ssName: string, destCellId: string, srcCellId: string)
    : Promise<Result<Updates>>
  {
    const url = makeURL(`${this.apiUrl}/${ssName}/${destCellId}`, {srcCellId});
    return await doFetchJson('PATCH', url);
  }

  /** set formula for cell cellId in spreadsheet ssName to expr.
   *  Return Updates object mapping cellId's to the updated value.
   */
  async evaluate(ssName: string, cellId: string, expr: string)
  : Promise<Result<Updates>>
  {
    const url = makeURL(`${this.apiUrl}/${ssName}/${cellId}`, {expr});
    return await doFetchJson('PATCH', url)!;
  }

  /** return list of [cellId, expr] pairs for spreadsheet ssName */
  async dump(ssName: string) : Promise<Result<[string, string][]>>
  {
    const url = makeURL(`${this.apiUrl}/${ssName}`);
    return await doFetchJson('GET', url);
  }

  /** return list of [cellId, expr, value] triples for spreadsheet ssName */
  async dumpWithValues(ssName: string)
    : Promise<Result<[string, string, number][]>>
  {
    const url = makeURL(`${this.apiUrl}/${ssName}`, {doValues: 'true'});
    return await doFetchJson('GET', url);
  }

  /** load spreadsheet ssName with dump of list of [ cellId, expr ]. */
  async load(ssName: string, dump: [string, string][])
  : Promise<Result<undefined>>
  {
    const url = makeURL(`${this.apiUrl}/${ssName}`);
    return await doFetchJson('PUT', url, dump);
  }

  /** clear out all contents of spreadsheet ssName */
  async clear(ssName: string)
  : Promise<Result<undefined>> 
  {
    const url = makeURL(`${this.apiUrl}/${ssName}`);
    return await doFetchJson('DELETE', url);
  }

}

/** A utility function used to extend a url with properly encoded 
 *  query parameters
 */
function makeURL(url: string, queryParams: Record<string, string|number> = {}) {
  const urlObj = new URL(url);
  Object.entries(queryParams).forEach(([k, v]) => {
    urlObj.searchParams.append(k, v.toString())
  });
  return urlObj;
}

/** Return a Result for dispatching HTTP method to url.  If jsonBody
 *  is specified, then it should be sent as the JSONrequest body
 *  (along with a suitable MIME-type).
 *  
 *  This function should convert the response envelope used for
 *  the web services into a Result.  Specifically: 
 *
 *    + The response should return an error Result if there is a fetch
 *      error or if the response JSON contains errors.
 *
 *    + If there are no errors then the function should return the
 *      response result within an ok Result.
 */
async function doFetchJson<T>(method: string, url: URL,
			      jsonBody?: object)
  : Promise<Result<T>> 
{
  //TODO
  return okResult('TODO' as any);
}

