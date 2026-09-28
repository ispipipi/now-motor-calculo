import { createReadStream, statSync } from 'node:fs';
import { createServer } from 'node:http';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const webRoot = path.dirname(fileURLToPath(import.meta.url));
const port = Number(process.env.PORT || process.argv[2] || 8063);
const rexApiBaseUrl = process.env.REX_API_BASE_URL || 'https://rexmas-api-531866499459.us-central1.run.app';
const rexApiKey = process.env.REX_API_KEY || '';
const mimeTypes = {
  '.css': 'text/css; charset=utf-8',
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.map': 'application/json; charset=utf-8',
  '.svg': 'image/svg+xml',
};

function jsonResponse(response, status, payload) {
  response.writeHead(status, {
    'Content-Type': 'application/json; charset=utf-8',
    'Cache-Control': 'no-store',
    'Access-Control-Allow-Origin': '*',
  });
  response.end(JSON.stringify(payload));
}

function rowsFromPayload(payload) {
  const visit = (value) => {
    if (Array.isArray(value)) {
      if (!value.length) return [];
      if (value.every((item) => item && typeof item === 'object' && !Array.isArray(item))) return value;
      return null;
    }
    if (!value || typeof value !== 'object') return null;
    const columns = value.columns || value.headers;
    const matrix = value.rows;
    if (Array.isArray(columns) && Array.isArray(matrix)) {
      return matrix.map((row) => columns.reduce((record, column, index) => {
        record[column] = Array.isArray(row) ? row[index] ?? '' : row[column] ?? '';
        return record;
      }, {}));
    }
    for (const key of ['records', 'rows', 'data', 'results', 'items', 'report']) {
      const result = visit(value[key]);
      if (result) return result;
    }
    return null;
  };
  return visit(payload) || [];
}

async function handleRexMaster(request, response, requestUrl) {
  if (request.method === 'OPTIONS') {
    response.writeHead(204, { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Methods': 'GET, OPTIONS', 'Access-Control-Allow-Headers': 'Content-Type' });
    response.end();
    return true;
  }
  if (request.method !== 'GET') {
    jsonResponse(response, 405, { error: 'Método no permitido' });
    return true;
  }
  if (!rexApiKey) {
    jsonResponse(response, 503, { error: 'REX_API_KEY no está configurada en el servidor local.' });
    return true;
  }
  const query = new URL(requestUrl, `http://localhost:${port}`).searchParams;
  const empresa = query.get('empresa') || 'NOW';
  const upstream = new URL('/api/v1/reports/saved-report', rexApiBaseUrl);
  upstream.searchParams.set('report_name', 'Mestro NOW');
  upstream.searchParams.set('empresa', empresa);
  upstream.searchParams.set('output_format', 'json');
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 15000);
  try {
    const upstreamResponse = await fetch(upstream, {
      headers: { 'X-API-Key': rexApiKey, Accept: 'application/json' },
      signal: controller.signal,
    });
    const body = await upstreamResponse.text();
    if (!upstreamResponse.ok) {
      jsonResponse(response, 502, { error: `REX+ respondió ${upstreamResponse.status}.`, detail: body.slice(0, 500) });
      return true;
    }
    let payload;
    try {
      payload = JSON.parse(body);
    } catch {
      jsonResponse(response, 502, { error: 'REX+ no devolvió JSON válido.' });
      return true;
    }
    jsonResponse(response, 200, { source: 'rex+', report: 'Mestro NOW', empresa, rows: rowsFromPayload(payload), fetchedAt: new Date().toISOString() });
  } catch (error) {
    jsonResponse(response, 502, { error: error.name === 'AbortError' ? 'REX+ tardó demasiado en responder.' : 'No fue posible conectar con REX+.' });
  } finally {
    clearTimeout(timeout);
  }
  return true;
}

function safePath(requestUrl) {
  const pathname = decodeURIComponent(new URL(requestUrl, `http://localhost:${port}`).pathname);
  const relative = pathname === '/' ? 'index.html' : pathname.replace(/^\/+/, '');
  const target = path.resolve(webRoot, relative);
  return target === webRoot || target.startsWith(`${webRoot}${path.sep}`) ? target : null;
}

const server = createServer(async (request, response) => {
  const requestPath = new URL(request.url || '/', `http://localhost:${port}`).pathname;
  if (requestPath === '/api/rex/master') {
    await handleRexMaster(request, response, request.url || '/api/rex/master');
    return;
  }
  const target = safePath(request.url || '/');
  if (!target) {
    response.writeHead(403, { 'Content-Type': 'text/plain; charset=utf-8' });
    response.end('Acceso denegado');
    return;
  }
  try {
    if (!statSync(target).isFile()) throw new Error('No es un archivo');
    const extension = path.extname(target).toLowerCase();
    response.writeHead(200, {
      'Content-Type': mimeTypes[extension] || 'application/octet-stream',
      'Cache-Control': 'no-cache',
    });
    createReadStream(target).pipe(response);
  } catch {
    response.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
    response.end('Archivo no encontrado');
  }
});

server.listen(port, '127.0.0.1', () => {
  console.log(`NOW plataforma disponible en http://127.0.0.1:${port}`);
});
