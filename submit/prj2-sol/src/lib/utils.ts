import LIMITS from './limits.js';

export type ObjMap<T> = { [key: string]: T };


export function colSpecToIndex(colSpec : string) : number {
  console.assert(0 < colSpec.length && colSpec.length <= 1,
		 'col coord can have only a single letter');
  const a = 'a'.codePointAt(0)!;
  return colSpec[0].codePointAt(0)! - a;
}

export function indexToColSpec(index: number, baseIndex=0) : string {
  console.assert(0 < LIMITS.MAX_N_COLS,
		 `bad col index ${index}; must be under ${LIMITS.MAX_N_COLS}`);
  const a = 'a'.codePointAt(0)!;
  return String.fromCodePoint(a + baseIndex + index);
}

export function rowSpecToIndex(rowSpec: string) : number {
  const index = Number(rowSpec) - 1;
  const max =  LIMITS.MAX_N_ROWS;
  console.assert(index < max, `bad row spec "${rowSpec}" exceeds ${max}`);
  return index;
}


export function indexToRowSpec(index: number, baseIndex=0) {
  console.assert(index < LIMITS.MAX_N_ROWS,
		 `bad row index ${index}; must be under ${LIMITS.MAX_N_ROWS}`);
  return String(baseIndex + index + 1);
}
