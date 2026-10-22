'use strict';
// Keep this entry point ASCII-only for Windows CMD compatibility.
const service = require('./server/\u6279\u91cf\u67e5\u8be2\u670d\u52a1.cjs');
if (process.argv.includes('--check')) {
  if (Number(process.versions.node.split('.')[0]) < 18) throw Error('Node.js 18 or later is required.');
  console.log('Startup check passed (Node.js ' + process.version + ').');
} else {
  service.startServer();
}
