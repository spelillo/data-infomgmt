#!/usr/bin/env node
// A tiny static server for local development and the browser tests:
//   npm run serve            -> http://localhost:8080
//   PORT=3000 npm run serve
// The site also works by opening index.html directly; serving it over http
// matches how GitHub Pages hosts it.
'use strict';
const http = require('http');
const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const TYPES = {
  '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8',
  '.json': 'application/json', '.md': 'text/markdown; charset=utf-8', '.svg': 'image/svg+xml', '.png': 'image/png',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', '.csv': 'text/csv',
};

function createServer() {
  return http.createServer((req, res) => {
    let rel = decodeURIComponent(req.url.split('?')[0]).replace(/^\/+/, '');
    if (rel === '' || rel.endsWith('/')) rel += 'index.html';
    const file = path.normalize(path.join(ROOT, rel));
    if (!file.startsWith(ROOT + path.sep) || !fs.existsSync(file) || fs.statSync(file).isDirectory()) {
      res.writeHead(404, { 'content-type': 'text/plain' });
      res.end('Not found');
      return;
    }
    res.writeHead(200, { 'content-type': TYPES[path.extname(file)] || 'application/octet-stream', 'cache-control': 'no-store' });
    fs.createReadStream(file).pipe(res);
  });
}

// Resolves to a listening server; port 0 picks a free port.
function serve(port = 0, host = '127.0.0.1') {
  const server = createServer();
  return new Promise((resolve) => server.listen(port, host, () => resolve(server)));
}

module.exports = { serve, createServer };

if (require.main === module) {
  const port = Number(process.env.PORT) || 8080;
  serve(port, process.env.HOST || '127.0.0.1').then((s) => {
    console.log('Course site: http://' + s.address().address + ':' + s.address().port + '/');
  });
}
