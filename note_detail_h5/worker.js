const assetRoot = '/assets/packages/ai_note_feature_note/assets/webview/h5/ainote-share-mobile';
const assetPrefixes = ['/js/', '/css/', '/img/', '/fonts/', '/static/', '/markmap/'];
const apiPrefixes = ['/business/aitablet/', '/aitablet/', '/api/detail', '/files/', '/audio/', '/image/'];
let nextRequestId = 0;

self.addEventListener('install', event => event.waitUntil(self.skipWaiting()));
self.addEventListener('activate', event => event.waitUntil(self.clients.claim()));

self.addEventListener('fetch', event => {
  const url = new URL(event.request.url);
  if (url.origin !== self.location.origin) return;
  if (assetPrefixes.some(prefix => url.pathname.startsWith(prefix))) {
    event.respondWith(fetch(assetRoot + url.pathname + url.search));
    return;
  }
  if (!apiPrefixes.some(prefix => url.pathname.startsWith(prefix))) return;
  event.respondWith(forwardToApp(event));
});

async function forwardToApp(event) {
  const client = await self.clients.get(event.clientId);
  if (!client || !new URL(client.url).pathname.startsWith('/note_detail_h5/')) {
    return new Response('', { status: 403 });
  }
  const channel = new MessageChannel();
  const id = ++nextRequestId;
  const body = event.request.method === 'GET' || event.request.method === 'HEAD'
    ? '' : await event.request.text();
  const result = new Promise(resolve => {
    const timeout = setTimeout(() => resolve({ status: 504, body: '' }), 15000);
    channel.port1.onmessage = message => {
      clearTimeout(timeout);
      resolve(message.data);
    };
  });
  client.postMessage({
    source: 'ai-note-detail-worker', id,
    method: event.request.method,
    path: new URL(event.request.url).pathname,
    query: new URL(event.request.url).search,
    body
  }, [channel.port2]);
  const response = await result;
  return new Response(response.body || '', {
    status: response.status || 200,
    headers: { 'content-type': response.contentType || 'application/json; charset=utf-8' }
  });
}
