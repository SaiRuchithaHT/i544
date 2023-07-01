import {default as parse, CellRef, Ast } from './expr-parser.js';

import { Result, okResult, errResult } from 'cs544-js-utils';

//factory method
export function makeSpreadsheet(name: string) : Spreadsheet
{
  return new Spreadsheet(name);
}

export type Updates = { [cellId: string]: number };
export class Spreadsheet {

  readonly name: string;
  cells: { [cellId: string]: CellInfo };
  
  constructor(name: string) {
    this.name = name;
    this.cells = {};
  }

  /** Set cell with id cellId to result of evaluating formula
   *  specified by the string expr.  Update all cells which are
   *  directly or indirectly dependent on the base cell cellId.
   *  Return an object mapping the id's of all updated cells to
   *  their updated values.  
   *
   *  Errors must be reported by returning an error Result having its
   *  code options property set to `SYNTAX` for a syntax error and
   *  `CIRCULAR_REF` for a circular reference and message property set
   *  to a suitable error message.
   */
  eval(cellId: string, expr: string) : Result<Updates> {
    const astResult = parse(expr, cellId);
    if (!astResult.isOk) return astResult;
    const ast = astResult.val;
    const validateResult = this.validate(cellId, ast);
    if (!validateResult.isOk) return validateResult;
    const cell = this.getCell(cellId);
    if (cell.ast) this.removeAsDependent(cellId, cell.ast);
    cell.ast = ast; cell.expr = expr.replace(/\s/g, '');
    const updates = this.evalCell(cell);
    return okResult(updates);
  }

  /** return object containing expr and value for cell cellId 
   *  return { value: 0, expr: '' } for an empty cell.
   */
  query(cellId: string): { value: number, expr: string } {
    const cell = this.cells[cellId];
    return { value: cell?.value ?? 0, expr: cell?.expr ?? '', };
  }

  /** Clear contents of this spreadsheet. No undo information recorded. */
  clear() {
    this.cells = {};
  }

  /** Remove all info for cellId from this spreadsheet. Return an
   *  object mapping the id's of all dependent cells to their updated
   *  values.  
   */
  remove(cellId: string) : Result<Updates> {
    const results = {};
    if (this.cells[cellId]) {
      const dependents = this.cells[cellId].dependents;
      delete this.cells[cellId];
      for (const dependent of dependents) {
	const expr = this.cells[dependent].expr
	const evalResult = this.eval(dependent, expr);
	if (!evalResult.isOk) return evalResult;
	Object.assign(results, evalResult.val);
      }
    }
    return okResult(results);
  }

  /** copy formula from srcCellId to destCellId, adjusting any
   *  relative cell references suitably.  Return an object mapping the
   *  id's of all dependent cells to their updated values. Copying
   *  an empty cell is equivalent to deleting the destination cell.
   */
  copy(destCellId: string, srcCellId: string) : Result<Updates> {
    const srcCell = this.cells[srcCellId];
    if (!srcCell?.expr) {
      return this.remove(destCellId);
    }
    else {
      const destCellRef = CellRef.parseRef(destCellId);
      const destExpr = srcCell.ast!.toText(destCellRef);
      return this.eval(destCellId, destExpr);
    }
  }

  /** Return dump of cell values as list of cellId and formula pairs.
   *  Do not include any cell's with empty expr.
   *
   *  Returned list must be sorted by cellId with primary order being
   *  topological (cell A < cell B when B depends on A) and secondary
   *  order being lexicographical (when cells have no dependency
   *  relation). 
   *
   *  Specifically, the cells must be dumped in a non-decreasing depth
   *  order:
   *     
   *    + The depth of a cell with no dependencies is 0.
   *
   *    + The depth of a cell C with direct prerequisite cells
   *      C1, ..., Cn is max(depth(C1), .... depth(Cn)) + 1.
   *
   *  Cells having the same depth must be sorted in lexicographic order
   *  by their IDs.
   *
   *  Note that empty cells must be ignored during the topological
   *  sort.
   */
  dump() : [string, string][] {
    const cells = this.cells;
    const prereqs = this.makePrereqs();
    const toDo = new Set(Object.keys(cells));
    const cellIds0 = Object.entries(prereqs)
      .filter(([cellId, dependents]) => dependents.length === 0)
      .map(([cellId, dependents]) => cellId);
    type CellDepths = { [cellId: string]: number };
    const cellDepths: CellDepths = {};
    cellIds0.forEach(cellId => {
      toDo.delete(cellId);
      cellDepths[cellId] = 0;
    });
    const maxKidsDepth = (cellId: string, cellDepths: CellDepths) => {
      return Math.max(...prereqs[cellId].map(c => cellDepths[c] ?? Infinity));
    };
    while (toDo.size > 0) {
      for (const cellId of toDo) {
	const prereqsDepth = maxKidsDepth(cellId, cellDepths);
	if (prereqsDepth !== Infinity) {
	  cellDepths[cellId] = prereqsDepth + 1;
	  toDo.delete(cellId);
	}
      }
    }
    const sortCellFn = (cellId1: string, cellId2: string) => {
      return (cellDepths[cellId1] - cellDepths[cellId2])
	|| cellId1.localeCompare(cellId2);
    };
    return Object.keys(cells)
      .filter(cellId => cells[cellId].expr !== '')
      .sort(sortCellFn)
      .map(cellId => [cellId, cells[cellId].expr]);
  }

  /** Return object mapping cellId to list containing prerequisites
   *  for cellId (reverse of dependencies).
   */
  private makePrereqs() {
    const prereqs: { [cellId: string] : string[] } =
      Object.fromEntries(Object.keys(this.cells).map(c => [c, []]));
    for (const cell of Object.values(this.cells)) {
      for (const d of cell.dependents) {
	if (prereqs[d])	prereqs[d].push(cell.id);
      }
    }
    return prereqs;
  }

  private validate(baseCellId: string, ast:Ast) : Result<undefined> {
    const prereqs = new Set<string>();
    addAstCells(CellRef.parseRef(baseCellId), ast, prereqs);
    const dependents = [ baseCellId ];
    while (dependents.length > 0) {
      const dependent = dependents.pop() as string;
      if (prereqs.has(dependent)) {
	const msg = `circular ref involving ${dependent}`;
	return errResult(msg, 'CIRCULAR_REF');
      }
      const depCell = this.getCell(dependent);
      depCell.dependents.forEach(cellId => dependents.push(cellId));
    }
    return okResult(undefined);
  }

  private evalCell(cell: CellInfo) : Updates {
    const value = this.evalAst(cell.id, cell.ast);
    cell.setValue(value);
    const vals = { [cell.id]: value };
    for (const dependent of cell.dependents) {
      const depCell = this.getCell(dependent);
      const depCellUpdates = this.evalCell(depCell);
      Object.assign(vals, {...depCellUpdates});
    }
    return vals;
  }

  private evalAst(baseCellId: string, ast?: Ast) : number {
    if (ast === null || ast === undefined) {
      return 0;
    }
    else if (ast.kind === 'num') {
      return ast.value;
    }
    else if (ast.kind === 'ref') {
      const baseCell = CellRef.parseRef(baseCellId);
      const cellId = ast.toText(baseCell);
      const cell = this.getCell(cellId);
      cell.addDependent(baseCellId);
      return cell.value;
    }
    else {
      console.assert(ast.kind === 'app', `unknown ast type ${ast.kind}`);
      const f = FNS[ast.fn];
      console.assert(!!f, `unknown ast fn ${ast.fn}`);
      return f(this.evalAst(baseCellId, ast.kids[0]),
	       ast.kids[1] && this.evalAst(baseCellId, ast.kids[1]));
    }
  }

  private removeAsDependent(baseCellId: string, ast: Ast) {
    if (ast.kind === 'app') {
      ast.kids.forEach(k => this.removeAsDependent(baseCellId, k));
    }
    else if (ast.kind === 'ref') {
      const baseCell = CellRef.parseRef(baseCellId);
      const cellId = ast.toText(baseCell);
      this.getCell(cellId).rmDependent(baseCellId);
    }
  }

  private getCell(cellId: string) {
    const id = cellId.replace(/\$/g, '');
    const cell = this.cells[id];
    return cell ?? (this.cells[id] = new CellInfo(id));
  }
}

function addAstCells(baseCell: CellRef, ast: Ast, cellIds: Set<string>) {
  if (ast.kind === 'app') {
    ast.kids.forEach(k => addAstCells(baseCell, k, cellIds));
  }
  else if (ast.kind === 'ref') {
    const cellId = ast.value.toText(baseCell).replace(/\$/g, '');
    cellIds.add(cellId);
  }
}

class CellInfo {

  readonly id: string;
  value: number;
  expr: string;
  ast?: Ast;
  dependents: Set<string>;
    
  constructor(id: string) {
    this.id = id;
    this.value = 0;
    this.expr = '';
    this.ast = undefined;
    this.dependents = new Set();
  }
  
  setValue(value: number) {
    this.value = value;
  }
  addDependent(cellId: string) {
    this.dependents.add(cellId);
  }
  rmDependent(cellId: string) {
    this.dependents.delete(cellId);
  }
}

const FNS = {
  '+': (a:number, b:number) : number => a + b,
  '-': (a:number, b?:number) : number => b === undefined ? -a : a - b,
  '*': (a:number, b:number) : number => a * b,
  '/': (a:number, b:number) : number => a / b,
  min: (a:number, b:number) : number => Math.min(a, b),
  max: (a:number, b:number) : number => Math.max(a, b),
}
