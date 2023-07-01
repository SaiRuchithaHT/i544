import cli from './lib/cli.js';

cli(process.argv.slice(2)).catch(err => console.error(err));

