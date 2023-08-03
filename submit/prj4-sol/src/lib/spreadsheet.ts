import SpreadsheetWs from './ss-ws.js';

import { Result, okResult, errResult } from 'cs544-js-utils';

import { Errors, makeElement } from './utils.js';

const [N_ROWS, N_COLS] = [10, 10];

export default async function make(ws: SpreadsheetWs, ssName: string) {
  return await Spreadsheet.make(ws, ssName);
}


class Spreadsheet {

  private readonly ws: SpreadsheetWs;
  private readonly ssName: string;
  private readonly errors: Errors;
  private currentFocusedCell: string | null = null;
  private copySourceCell: string | null = null;
  private errorOccurred: boolean = false;

  constructor(ws: SpreadsheetWs, ssName: string) {
    this.ws = ws;
    this.ssName = ssName;
    this.errors = new Errors();
    this.currentFocusedCell = null;
    this.copySourceCell = null;
    this.makeEmptySS();
    this.addListeners();
    //TODO: initialize added instance variables
  }

  static async make(ws: SpreadsheetWs, ssName: string) {
    const ss = new Spreadsheet(ws, ssName);
    await ss.load();
    return ss;
  }

  /** add listeners for different events on table elements */
  private addListeners() {
    const clearButton = document.querySelector('#clear');
    clearButton?.addEventListener('click', this.clearSpreadsheet);

    const cells = document.querySelectorAll('.cell');
    cells.forEach(cell => {
      cell.addEventListener('focusin', this.focusCell);
      cell.addEventListener('focusout', this.blurCell);
      cell.addEventListener('copy', this.copyCell);
      cell.addEventListener('paste', this.pasteCell);
    });
  }

  /** listener for a click event on #clear button */
  private readonly clearSpreadsheet = async (ev: Event) => {
    const result = await this.ws.clear(this.ssName);
    if (result.isOk) {
      this.clearCellData();
    } else {
      this.errors.display(result.errors);
    }
  };

  /** clear cell data when the "Clear" button is clicked */
  private clearCellData() {
    const cells = document.querySelectorAll('.cell');
    cells.forEach(cell => {
      cell.textContent = '';
      cell.removeAttribute('data-expr');
      cell.removeAttribute('data-value');
    });
  }

  /** listener for the focus event on a spreadsheet data cell */
  private readonly focusCell = (ev: Event) => {
    const cell = ev.target as HTMLElement;
    this.currentFocusedCell = cell.id;
    const expr = cell.getAttribute('data-expr');
    cell.textContent = expr || cell.getAttribute('data-value') || '';
    this.errors.clear();
  };

  /** listener for a blur event on a spreadsheet data cell */
  private readonly blurCell = async (ev: Event) => {
    const cell = ev.target as HTMLElement;
    const cellId = cell.id;
    const content = cell.textContent?.trim() || '';

    const prevExpr = cell.getAttribute('data-expr');
    const prevValue = cell.getAttribute('data-value');
    cell.dataset.expr = content;

    if (content === '') {
      const result = await this.ws.remove(this.ssName, cellId);
      if (result.isOk) {
        cell.textContent = '';
        cell.removeAttribute('data-expr');
        cell.removeAttribute('data-value');
        this.updateCells(result.val);
      } else {
        this.errors.display(result.errors);
        this.errorOccurred = true;
        cell.setAttribute('data-prev-expr', prevExpr || '');
        cell.setAttribute('data-prev-value', prevValue || '');
      }
    } else {
      const result = await this.ws.evaluate(this.ssName, cellId, content);
      if (result.isOk) {
        cell.textContent = cell.dataset.value || '';
        this.updateCells(result.val);
      } else {
        this.errors.display(result.errors);
        this.errorOccurred = true;

        cell.setAttribute('data-prev-expr', prevExpr || '');
        cell.setAttribute('data-prev-value', prevValue || '');
      }
    }
    this.currentFocusedCell = null;
    if (this.errorOccurred) {
      this.resetCellContent(cell);
      this.errorOccurred = false;
    }
  };

  /** reset cell content to its previous value (if error occurred during evaluation) */
  private resetCellContent(cell: HTMLElement) {
    const prevExpr = cell.getAttribute('data-prev-expr');
    const prevValue = cell.getAttribute('data-prev-value');
    cell.textContent = prevValue || prevExpr || '';
    cell.setAttribute('data-expr', prevExpr || '');
    cell.setAttribute('data-value', prevValue || '');
  }

  /** listener for a copy event on a spreadsheet data cell */
  private readonly copyCell = (ev: Event) => {
    const cell = ev.target as HTMLElement;
    const currentCopySourceCell = document.querySelector('.is-copy-source');

    if (currentCopySourceCell) {
      currentCopySourceCell.classList.remove('is-copy-source');
    }
    this.copySourceCell = cell.id;

    cell.classList.add('is-copy-source');
  };

  /** listener for a paste event on a spreadsheet data cell */
  private readonly pasteCell = async (ev: Event) => {
    ev.preventDefault();
    const destinationcell = ev.target as HTMLElement;
    const destinationCellId = destinationcell.id;

    // Check if we have a source cell from which content is copied
    if (this.copySourceCell) {
      const sourceCellId = this.copySourceCell;
      const result = await this.ws.copy(this.ssName, destinationCellId, sourceCellId);
      if (result.isOk) {
        const updates = result.val;
        const res2 = await this.ws.query(this.ssName, destinationCellId);
        if (res2.isOk) {
          const destExpressionVal = res2.val;
          const destExpression = destExpressionVal.expr;
          destinationcell.setAttribute('data-exp', destExpression);
          this.updateCells(result.val);
          destinationcell.textContent = destExpression;
        }
      } else {
        this.errors.display(result.errors);
      }

      // Clear the copy source cell
      const copySourceCellElement = document.getElementById(this.copySourceCell);
      if (copySourceCellElement) {
        copySourceCellElement.classList.remove('is-copy-source');
        this.copySourceCell = null;
      }
    }
  };

  /** update cells with new values after successful evaluation */
  private updateCells(updates: { [cellId: string]: number }) {
    const cells = document.querySelectorAll('.cell');
    cells.forEach(cell => {
      const cellId = cell.id;
      if (updates.hasOwnProperty(cellId) && (this.currentFocusedCell != cellId)) {
        const value = updates[cellId];
        cell.setAttribute('data-value', value.toString());
        cell.textContent = value.toString();
      }
    });
  }

  /** load existing data from the spreadsheet */
  private async load() {
    const result = await this.ws.dumpWithValues(this.ssName);
    if (result.isOk) {
      const data = result.val;
      const cells = document.querySelectorAll('.cell');
      cells.forEach(cell => {
        const cellId = cell.id;
        const cellData = data.find(item => item[0] === cellId);
        if (cellData) {
          const [_, expr, value] = cellData;
          cell.textContent = value.toString();
          cell.setAttribute('data-expr', expr);
          cell.setAttribute('data-value', value.toString());
        }
      });
    } else {
      this.errors.display(result.errors);
    }
  }


  private makeEmptySS() {
    const ssDiv = document.querySelector('#ss')!;
    ssDiv.innerHTML = '';
    const ssTable = makeElement('table');
    const header = makeElement('tr');
    const clearCell = makeElement('td');
    const clear = makeElement('button', { id: 'clear', type: 'button' }, 'Clear');
    clearCell.append(clear);
    header.append(clearCell);
    const A = 'A'.charCodeAt(0);
    for (let i = 0; i < N_COLS; i++) {
      header.append(makeElement('th', {}, String.fromCharCode(A + i)));
    }
    ssTable.append(header);
    for (let i = 0; i < N_ROWS; i++) {
      const row = makeElement('tr');
      row.append(makeElement('th', {}, (i + 1).toString()));
      const a = 'a'.charCodeAt(0);
      for (let j = 0; j < N_COLS; j++) {
        const colId = String.fromCharCode(a + j);
        const id = colId + (i + 1);
        const cell = makeElement('td', { id, class: 'cell', contentEditable: 'true' });
        row.append(cell);
      }
      ssTable.append(row);
    }
    ssDiv.append(ssTable);
  }

}


