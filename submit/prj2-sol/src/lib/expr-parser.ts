import {
  colSpecToIndex, rowSpecToIndex,
  indexToColSpec, indexToRowSpec,
} from './utils.js';
import { panic, errResult, okResult, Result } from 'cs544-js-utils';

/************************* Top Level Functions *************************/

/*
Default export is 

parseExpr(expr, baseCellRef): Result<Ast>

expr is a string specifying a spreadsheet formula which could be typed
by an end-user into the cell specified by baseCellRef.  All relative
references in expr are WRT baseCellRef.

If an error is detected in `expr` or `baseCellRef`, then the result
is an error-result; otherwise it is an ok-result with val set to the Ast.

An AST is described below.
*/

export default function parseExpr(expr: string, base: CellRef|string = 'A1')
  : Result<Ast> 
{
  let baseCell: CellRef;
  if (typeof base === 'string') {
    const baseCellResult = CellRef.parse(base);
    if (!baseCellResult.isOk) return baseCellResult;
    baseCell = baseCellResult.val;
  }
  else {
    baseCell = base;
  };
  const tokensResult = scan(expr, baseCell);
  if (!tokensResult.isOk) return tokensResult;
  const parser = new ExprParser(tokensResult.val, baseCell);
  return parser.parse();
}

/*

The value returned by parseExpr() is a node of an Abstract Syntax Tree (AST).

An AST node is either an internal node or a leaf node.  Each
AST node is represented as JS object with a kind property
identifying the type of the node:

  kind === 'app':
    An internal node.  Additional properties 'fn' which is one of '+'
    '-', '*', '/','max, or 'min'and 'kids' which is a list of AST's.  It
    represents the application of function 'fn' to 'kids'.

  kind === 'num':
    A leaf node representing a number. It has an additional property
    named 'value' which is a JS number.

  kind === 'ref':
    A leaf node representing a reference to a spreadsheet cell.
    It has an additional property named 'value' which is a CellRef
    object having properties named row and col, each of which
    has the following properties:

      index: the index (0-based) of the row or col in the spreadsheet.

      isAbs: truthy if the index is an absolute index.
             falsy if the index is relative to the cell containing
	     the expr correponsing to the overall AST.

For example, the AST corresponding to the formula 
'((1 + c$2) * $b3)' entered into cell 'a5' is:

AppAst {
  kind: 'app',
  fn: '*',
  kids: [
    AppAst {
      kind: 'app',
      fn: '+',
      kids: [
        NumAst { kind: 'num', value: 1 },
        RefAst {
          kind: 'ref',
          value: CellRef {
            col: { isAbs: false, index: 2 },  //c - a
            row: { isAbs: true, index: 1 }    //$2, but 0-origin
          }
        }
      ],
    },
    RefAst {
      kind: 'ref',
      value: CellRef {
        col: { isAbs: true, index: 1 },       //$b, 0-orign
        row: { isAbs: false, index: -2 }      //3 - 5
      }
    }
  ],
}

Note that this representation a formula's AST makes it possible to copy
a formula from one cell to another by simply copying the AST with the
relative cell references unchanged but still working correctly in the new
cell.

An AST has a toString(baseCell) method which produces a minimally
parenthesized representation of the AST with all relative cell references
made relative to baseCell.

*/

/********************** Abstract Syntax Tree Types *********************/

// All Ast variants have a `kind` field which has a string literal
// type.  This allows TS to use type narrowing on the AST variant
// based on the `kind`.

type Op = keyof typeof OPS;  //an operator like '+', '-', etc.
type Fn = 'max' | 'min';

class RefAst {
  readonly kind: 'ref';
  readonly value: CellRef;
  constructor(value: CellRef) {
    this.kind = 'ref';
    this.value = value;
  }
  prec() { return MAX_PREC; }
  toText(baseCell: CellRef) {
    return this.value.toText(baseCell);
  }
}

class NumAst {
  readonly kind: 'num';
  readonly value: number;
  constructor(value: number) {
    this.kind = 'num';
    this.value = value;
  }
  prec() { return MAX_PREC; }
  toText(baseCell: CellRef) {
    return this.value.toString();
  }
}

class AppAst {
  readonly kind: 'app';
  readonly fn: Op|Fn;
  readonly kids: Ast[];
  constructor(fn: Op|Fn, kids: Ast[]) {
    this.kind = 'app';
    this.fn = fn; this.kids = kids;
  }

  prec() {
    return (this.fn === 'max' || this.fn === 'min')
      ? MAX_PREC
      : OPS[this.fn].prec;
  }

  toText(baseCell: CellRef): string {
    const fn = this.fn;
    if (fn === 'max' || fn === 'min') {
      return fn +
	'(' + this.kids.map(k=>k.toText(baseCell)).join(', ') + ')';
    }
    else {
      const fnInfo = OPS[fn];
      if (fnInfo.assoc === 'left') {
	if (this.kids.length === 1) {
	  console.assert(fn === '-', "'-' is only unary operator");
	  return fn + this.kids[0].toText(baseCell);
	}
	else {
	  console.assert(this.kids.length === 2,
			 'assoc operator must be binary');
	  const p0 = (this.kids[0].prec() < fnInfo.prec);
	  const p1 = (this.kids[1].prec() <= fnInfo.prec);
	  return AppAst.left(p0) +
	    this.kids[0].toText(baseCell) + AppAst.right(p0) + fn +
	    AppAst.left(p1) + this.kids[1].toText(baseCell) +
	    AppAst.right(p1);
	}
      }
      else {
	panic(`operator type ${fnInfo.assoc} not handled`);
      }
    }
  }

  private static left(isParen: boolean) { return isParen ? '(' : ''; }
  private static right(isParen: boolean) { return isParen ? ')' : ''; }
  
}

export type Ast = RefAst | NumAst | AppAst;

/******************************** Parser ******************************/

/*  Crude recursive descent parser:
expr is a spreadsheet formula specified by the following EBNF grammar:

expr
  : term ( ( '+' | '-' ) term )*
  ;
term
  : factor ( ( '*' | '/' ) factor )*
  ;
factor
  : NUMBER
  | '-' factor
  | FN '(' expr ( ',' expr )* ')'
  | cellRef
  | '(' expr ')'
  ;
cellRef
  : '$'? LETTER '$'? DIGITS+ //no intervening whitespace

The above grammar gives the structure of a language over some
vocabulary of symbols (for the spreadsheet, the vocabulary consists
numbers, cell references, function names like max and min, arithmetic
operators like + and * and punctuation symbols like , ( and ).  

The grammar specifies the phrases in the language recognized by the
grammar using rules of the form

phrase
  : alt1
  | alt2
  | ...
  | altn
  ;

The top level phrase in the grammar is expr.

The alternatives alt1, alt2, ..., altN for each rule consists of a 
sequence of symbols of the following kind:

  Vocabulary Symbols:
    Enclosed within single-quotes '...' or an all upper-case identifier;
    the former stand for themselves; the latter are not defined further
    and stand for what is implied by their name.

  Phrase Symbols:
    An identifier starting with a lower-case letter.  Defined by
    a grammar rule.

  Meta Symbols:
    These are part of the grammar notation:

       * postfix operator denoting 0-or-more repetitions of the previous symbol.

       ? postfix operator denoting previous symbol is optional.

       | infix operator denoting alternatives

       ( ) used for grouping symbols

Note that quoted '(' and ')' are vocabulary symbols whereas ( ) are 
meta symbols used for grouping.

For example, the first rule above:

expr
  : term ( ( '+' | '-' ) term )*
  ;

says that an expr consists of 1-or-more term's separated by '+' or
'-'.

*/
class ExprParser {

  private readonly baseCell: CellRef;
  private readonly toks: Token[];
  private readonly nToks: number;
  private tok: Token;       //current lookahead
  private index: number;    //index of next token
  
  constructor(tokens: Token[], baseCell: CellRef) {
    this.baseCell = baseCell;
    this.toks = tokens;
    this.nToks = tokens.length;
    if (this.nToks === 0) panic('no tokens: expect at least EOF token');
    this.tok = this.toks[0];
    this.index = 1;
  }

  private nextTok() {
    if (this.index >= this.nToks) panic(`nextTok() bad index '${this.index}'`);
    this.tok = this.toks[this.index++];
  }

  private peek(kind: Token['kind'], lexeme?: string) {
    return this.tok.kind === kind && (!lexeme || this.tok.lexeme === lexeme);
  }
  
  private consume(kind: Token['kind'], lexeme?: string) {
    if (this.peek(kind, lexeme)) {
      if (this.tok.kind !== 'EOF') this.nextTok();
    }
    else {
      throw `unexpected token at '${this.tok.lexeme}': expected '${kind}'`;
    }
  }

  parse() : Result<Ast> {
    try {
      const e = this.expr();
      if (!this.peek('EOF')) return errResult('expected EOF', 'SYNTAX');
      return okResult(e);
    }
    catch (err) {
      return errResult(err, 'SYNTAX');
    }
  }

  private expr(): Ast {
    let t0 = this.term();
    while (this.peek('+') || this.peek('-')) {
      const op = this.tok.kind as '+'|'-';
      this.nextTok();
      const t1 = this.term();
      t0 = new AppAst(op, [t0, t1]);
    }
    return t0;
  }

  private term() {
    let f0 = this.factor();
    while (this.peek('*') || this.peek('/')) {
      const op = this.tok.kind as '*'|'/';
      this.nextTok();
      const f1 = this.factor();
      f0 = new AppAst(op, [f0, f1]);
    }
    return f0;
  }

  private factor() : Ast {
    let e;
    if (this.peek('(')) {
      this.nextTok();
      e = this.expr();
      this.consume(')');
    }
    else if (this.peek('ref')) {
      e = new RefAst(this.tok.value as CellRef);
      this.nextTok();
    }
    else if (this.peek('-')) {
      this.nextTok();
      const operand = this.factor();
      e = new AppAst('-', [operand]);
    }
    else if (this.peek('fn')) {
      const fn = this.tok.lexeme as 'max'|'min';
      this.nextTok();
      this.consume('(');
      const args: Ast[] = [];
      args.push(this.expr());
      while (this.peek(',')) {
	this.nextTok();
	args.push(this.expr());
      }
      this.consume(')');
      e = new AppAst(fn, args);
    }
    else {
      const t = this.tok;
      this.consume('num');
      e = new NumAst(t.value as number);
    }
    return e;
  }

}



/******************************* Scanner *******************************/

type Token = { kind: string, lexeme: string, value?: number|CellRef };

function scan(str: string, baseCell: CellRef) : Result<Token[]> {
  const tokens: Token[] = [];
  while ((str = str.trimLeft()).length > 0) {
    let tok : Token;
    const c = str[0];
    if (c.match(/\d/)) {
      const [ lexeme ] = str.match(/^\d+(\.\d+)?([eE][-+]?\d+)?/)!;
      tok = { kind: 'num', lexeme, value: Number(lexeme) };
    }
    else if (c.match(/[\w\$]/)) {
      const [ lexeme ] = str.match(/^[\w\$]+/)!;
      if (lexeme === 'max' || lexeme === 'min') {
	tok = { kind: 'fn', lexeme, };
      }
      else {
	const cellRefResult = CellRef.parse(lexeme, baseCell);
	if (!cellRefResult.isOk) return cellRefResult;
	tok = { kind: 'ref', lexeme, value: cellRefResult.val };
      }
    }
    else {
      tok = { kind: c, lexeme: c, };
    }
    str = str.slice(tok.lexeme.length);
    tokens.push(tok);
  } //while
  tokens.push({ kind: 'EOF', lexeme: '<EOF>',  });
  return okResult(tokens);
}

//for testing only
export { scan };


/**************************** Cell Reference ***************************/

class Coord {
  readonly index: number;
  readonly isAbs: boolean;
  constructor(index: number, isAbs: boolean) {
    this.index = index; this.isAbs = isAbs;
  }
};

export class CellRef {
  readonly col: Coord;
  readonly row: Coord;

  constructor(col: Coord, row: Coord) {
    this.col = col; this.row = row;
  }

  /** validate and parse ref-string relative to baseCell */
  static parse(ref: string, baseCell: CellRef=CELL_A1) : Result<CellRef> {
    ref = ref.trim().toLowerCase();
    const match = ref.match(/^(\$?)([a-zA-Z])(\$?)(\d+)$/);
    if (!match) {
      return errResult(`bad cell ref ${ref}`, 'SYNTAX');
    }
    else {
      return okResult(CellRef.parseRef(ref, baseCell));
    }
  }

  /** parse ref-string wrt baseCell; assumes ref is valid */
  static parseRef(ref: string, baseCell: CellRef=CELL_A1) : CellRef {
    const match = ref.match(/^(\$?)([a-zA-Z])(\$?)(\d+)$/)!;
    const [_, isAbsCol, colSpec, isAbsRow, rowSpec ] = match;
    const colIndex =
      colSpecToIndex(colSpec) - (isAbsCol ? 0 : baseCell.col.index ?? 0);
    const rowIndex =
      rowSpecToIndex(rowSpec) - (isAbsRow ? 0 : baseCell.row.index ?? 0);
    const col = new Coord(colIndex, !!isAbsCol);
    const row = new Coord(rowIndex, !!isAbsRow);
    return new CellRef(col, row);
  }
  
  toText(baseCell: CellRef=CELL_A1) {
    let str = '';
    if (this.col.isAbs) {
      str += '$' + indexToColSpec(this.col.index);
    }
    else {
      str += indexToColSpec(this.col.index, baseCell.col.index);
    }
    if (this.row.isAbs) {
      str += '$' + indexToRowSpec(this.row.index);
    }
    else {
      str += indexToRowSpec(this.row.index, baseCell.row.index);
    }
    return str;
  }
  
}


/********************************** Data *******************************/

const CELL_A1 = new CellRef(new Coord(0, false), new Coord(0, false));

const MAX_PREC = 100;

/* Operator Information */
const OPS = {
  '+': {
    fn: '+',
    prec: 10,
    assoc: 'left',
  },
  '-': {
    fn: '-',
    prec: 10,
    assoc: 'left',
  },
  '*': {
    fn: '*',
    prec: 20,
    assoc: 'left',
  },
  '/': {
    fn: '/',
    prec: 20,
    assoc: 'left',
  },
};

