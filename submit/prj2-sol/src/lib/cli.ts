import { errResult, okResult, Result } from 'cs544-js-utils';
import { readJson } from 'cs544-node-utils';

import assert from 'assert';
import Path from 'path';

import {
  clear as clearAct, copy as copyAct,  remove as removeAct,
  evaluate as evaluateAct, dump as dumpAct, load as loadAct, query as queryAct,
} from './ss-services.js';

import { makeSpreadsheetDao, SpreadsheetDao } from './spreadsheet-dao.js';

type Dump = [ string, string ][];

/** handler for load command */
async function loadFile(dao: SpreadsheetDao, fileName: string)
  : Promise<Result<undefined>> 
{
  const dataResult: Result<Dump> = await readJson(fileName);
  if (!dataResult.isOk) return dataResult;
  return loadAct(dao, dataResult.val);
}

class CmdArg {
  readonly name: string;
  readonly kind: string;
  constructor(name: string, kind: 'cellRef' | 'str' ) {
    this.name = name; this.kind = kind;
  }
  toString() { return this.name; }

  check(val: string) : Result<undefined> {
    const isOk =
	  (this.kind === 'cellRef' && val.match(/^[a-zA-Z]+\d+$/)) ||
	  (this.kind === 'str');
    if (!isOk) {
      return errResult(`invalid value '${val}' for ${this.name}; ` +
	`must be a ${this.kind}`);
    }
    return okResult(undefined);
  }
  
};

const CMD_WIDTH = 8;

type Act = (dao: SpreadsheetDao, ...args: string[]) => Promise<Result<any>>;
class Cmd {
  readonly name: string;
  readonly msg: string;
  readonly act: Act;
  readonly args: CmdArg[];
  constructor(name: string, msg: string, act: Act, ...args: CmdArg[]) {
    this.name = name; this.msg = msg; this.act = act; this.args = args;
  }

  toString() {
    const args = this.args.map(a => a.toString()).join(' ');
    return `
    ${this.name} ${args}
      ${this.msg}
    `.replace(/\s+$/, '');
  }

  check(args: string[]) : boolean {
    if (args.length !== this.args.length) {
      console.error(`command ${this.name} needs ${this.args.length} arguments`);
      return false;
    }
    for (const [i, arg] of args.entries()) {
      if (!this.args[i].check(arg)) return false;
    }
    return true;
  }

  async doIt(dao: SpreadsheetDao, args: string[]) {
    return await this.act(dao, ...args);
  }
}

const COMMANDS = Object.fromEntries(
  [
    new Cmd('clear',
	    'clear spreadsheet',
	    clearAct),
    new Cmd('copy',
	    'copy formula from cell SRC_CELL_ID to cell DEST_CELL_ID',
	    copyAct,
	    new CmdArg('DEST_CELL_ID', 'cellRef'),
	    new CmdArg('SRC_CELL_ID', 'cellRef')),
    new Cmd('delete',
	    'delete formula in cell specified by CELL_ID',
	    removeAct,
	    new CmdArg('CELL_ID', 'cellRef')),
    new Cmd('dump',
	    'dump spreadsheet formulas in topological order to stdout',
	    dumpAct),
    new Cmd('eval',
	    'eval formula FORMULA into cell CELL_ID',
	    evaluateAct,
	    new CmdArg('CELL_ID', 'cellRef'),
	    new CmdArg('FORMULA', 'str')),
    new Cmd('load',
	    'load previously dumped data from file FILE into spreadsheet',
	    loadFile,
	    new CmdArg('FILE', 'str')),
    new Cmd('query',
	    'return formula and current value of cell specified by CELL_ID',
	    queryAct,
	    new CmdArg('CELL_ID', 'cellRef')),
  ].map(cmd => [cmd.name, cmd])
);

/** output usage message */
function usage() {
  let msg =
    `usage: ${Path.basename(process.argv[1])} MONGO_DB_URL ` +
    `SPREADSHEET_NAME CMD [ARGS...]\n`;
  msg += 'Command CMD can be';
  Object.values(COMMANDS).forEach( cmd => msg += cmd.toString());
  console.error(msg);
  process.exit(1);
}

/** Top level routine */
export default async function go(args: string[]) {
  if (process.argv.length < 3) {
    usage();
  }
  assert(args.length >= 3);
  const [ mongoDbUrl, spreadsheetName, cmdName, ...cmdArgs ] = args;
  let dao;
  try {
    const daoResult = await makeSpreadsheetDao(mongoDbUrl, spreadsheetName);
    if (!daoResult.isOk) {
      console.error(daoResult.errors[0].message);
    }
    else {
      dao = daoResult.val;
      const cmd = COMMANDS[cmdName];
      if (!cmd) {
	console.error(`invalid command ${cmdName}: must be one of ` +
	  Object.keys(COMMANDS).join('|'));
	usage();
      }
      else if (!cmd.check(cmdArgs)) {
	usage();
      }
      else {
	const result = await cmd.doIt(dao, cmdArgs);
	if (result.isOk) {
	  if (result.val !== undefined) {
	    console.log(JSON.stringify(result.val, null, 2));
	  }
	}
	else {
	  for (const err of result.errors) {
	    console.error(err.message);
	  }
	}
      }
    }
  }
  finally {
    if (dao) await dao.close();
  }
}


