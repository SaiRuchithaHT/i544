import { default as parse, CellRef, Ast } from './expr-parser.js';

import { Result, okResult, errResult } from 'cs544-js-utils';

//factory method
export default async function makeSpreadsheet(name: string):
  Promise<Result<Spreadsheet>> {
  return okResult(new Spreadsheet(name));
}

type Updates = { [cellId: string]: number };

export class Spreadsheet {

  readonly name: string;
  private cells: { [cellId: string]: Ast };
  private shadowCells: { [cellId: string]: Ast };

  constructor(name: string) {
    this.name = name;
    this.cells = {};
    this.shadowCells = {};
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
  async eval(cellId: string, expr: string): Promise<Result<Updates>> {

    let parsedExpression: Result<Ast>;

    try {
      parsedExpression = parse(expr, cellId);
    } catch (error) {
      const msg = `SYNTAX error ...`;
      return errResult(msg, 'SYNTAX');
    }

    if (!parsedExpression.isOk) {
      const msg = `SYNTAX error ...`;
      return errResult(msg, 'SYNTAX');
    }

    this.shadowCells[cellId] = parsedExpression.val;

    const updatedCells: Updates = {};
    try {
      this.evaluateAndUpdate(cellId, updatedCells, cellId);
    } catch (error) {
      if (error instanceof CircularReferenceError) {
        const msg = `cyclic dependency ...`;
        return errResult(msg, 'CIRCULAR_REF');
      } else {
        throw error;
      }
    }

    const dependents = this.findDependentCells2(cellId);

    for (const dependent of dependents) {
      delete updatedCells[dependent];
    }

    this.coverCell(cellId);
    return okResult(updatedCells);
  }

  // method to check for the dependencies of the basecell to print final result
  private findDependentCells2(baseCellId: string): string[] {
    const dependents: string[] = [];
    for (const cellId in this.shadowCells) {
      if (!(cellId === baseCellId)) {
        const ast = this.shadowCells[cellId];
        if (!this.isDependent(ast, cellId, baseCellId)) {
          dependents.push(cellId);
        }
      }
    }
    return dependents;
  }

  // method to check for the dependencies and updating
  private evaluateAndUpdate(cellId: string, updatedCells: Updates, parentCellId: string): number {

    const cell = this.shadowCells[cellId];

    if (cell === undefined) {
      throw new Error(`Cell '${cellId}' does not exist`);
    }

    const result = this.evaluateFormula(cellId, cell, updatedCells, parentCellId);

    updatedCells[cellId] = result;

    const dependents = this.findDependentCells(cellId);
    for (const dependent of dependents) {
      if ((cellId === parentCellId) && !(cellId === dependent)) {
        parentCellId = dependent;
        this.evaluateAndUpdate(dependent, updatedCells, parentCellId);
      }
    }
    return result;
  }

  // method to evaluate the ast of the cell
  private evaluateFormula(basecellId: string, ast: Ast, updatedCells: Updates, parentCellId: string): number {

    if (ast.kind === 'num') {
      return ast.value;
    } else if (ast.kind === 'ref') {
      const result: CellRef = CellRef.parseRef(basecellId);
      const cellId = ast.toText(result);
      if ((cellId == parentCellId)) {
        this.recoverCell(parentCellId);
        throw new CircularReferenceError(`Circular reference found for cell '${cellId}'`);
      }
      if (!Object.prototype.hasOwnProperty.call(this.shadowCells, cellId)) {
        return 0;
      }
      if (Object.prototype.hasOwnProperty.call(updatedCells, cellId)) {
        return updatedCells[cellId];
      } else {
        return this.evaluateAndUpdate(cellId.toString(), updatedCells, parentCellId);
      }
    } else if (ast.kind === 'app') {
      const { fn, kids } = ast;
      type Op = '+' | '-' | '*' | '/';
      type Fn = 'min' | 'max';
      const operator = fn as Op | Fn;
      const operands = kids;
      const args = operands.map((operand: Ast) => {
        return this.evaluateFormula(basecellId, operand, updatedCells, parentCellId);
      });

      if (operator in FNS) {
        const result = FNS[operator].apply(null, args);
        return result;
      } else {
        throw new Error(`Unknown operator '${operator}'`);
      }

    } else {
      throw new Error(`Invalid ast type '${(ast as any).kind}'`);
    }
  }

  // method to check the dependency between the cells 
  private findDependentCells(baseCellId: string): string[] {
    const dependents: string[] = [];
    for (const cellId in this.shadowCells) {
      const ast = this.shadowCells[cellId];
      if (this.isDependent(ast, cellId, baseCellId)) {
        dependents.push(cellId);
      }
    }
    return dependents;
  }

  // method to check the dependency between the cells
  private isDependent(ast: Ast, cellId: string, baseCellId: string): boolean {
    if (ast.kind === 'ref') {
      const result: CellRef = CellRef.parseRef(cellId);
      return ast.toText(result) === baseCellId;
    } else if (ast.kind === 'app') {
      return ast.kids.some((operand: Ast) => this.isDependent(operand, cellId, baseCellId));
    } else {
      return false;
    }
  }

  // mthod to recover shallow cells from cells after Circular dependency
  public recoverCell(cellId: string): void {
    if (this.cells[cellId]) {
      this.shadowCells[cellId] = this.cells[cellId];
    }
  }

  // method to copy shallow cells value to cells
  public coverCell(cellId: string): void {
    if (this.shadowCells[cellId]) {
      this.cells[cellId] = this.shadowCells[cellId];
    }
  }

}

class CircularReferenceError extends Error { }

const FNS = {
  '+': (a: number, b: number): number => a + b,
  '-': (a: number, b?: number): number => b === undefined ? -a : a - b,
  '*': (a: number, b: number): number => a * b,
  '/': (a: number, b: number): number => a / b,
  min: (a: number, b: number): number => Math.min(a, b),
  max: (a: number, b: number): number => Math.max(a, b),
}
