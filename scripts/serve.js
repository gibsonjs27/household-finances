import http from 'node:http';
import { readFile } from 'node:fs/promises';
import { extname, resolve, sep } from 'node:path';
const root = resolve('.');
const types = { '.html':'text/html', '.js':'text/javascript', '.css':'text/css', '.svg':'image/svg+xml' };
http.createServer(async (request, response) => {
  try {
    const pathname = decodeURIComponent(new URL(request.url, 'http://localhost').pathname);
    const relative = pathname.replace(/^\/household-finances\/?/, '').replace(/^\//, '') || 'index.html';
    const path = resolve(root, relative);
    if (!path.startsWith(root + sep) || relative.includes('.git') || !types[extname(path)]) throw new Error('Not found');
    response.setHeader('Content-Type', types[extname(path)]);
    response.end(await readFile(path));
  } catch { response.writeHead(404); response.end('Not found'); }
}).listen(4173, '127.0.0.1', () => console.log('Local: http://127.0.0.1:4173/household-finances/'));
