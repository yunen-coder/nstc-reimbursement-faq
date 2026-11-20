/** Cloudflare Pages Function: 將前端 API 請求轉送到 Google Apps Script */
export async function onRequest(context) {
  const { request, env, params } = context;
  const path = Array.isArray(params.path) ? params.path.join('/') : String(params.path || '');
  const gasUrl = env.GAS_URL;

  if (!gasUrl) {
    return Response.json({ error: '尚未設定 GAS_URL。請在 Cloudflare Pages 環境變數中填入 Apps Script 網頁應用程式網址。' }, { status: 500 });
  }

  let action;
  let method = 'GET';
  let body;

  if (path === 'data') action = 'data';
  else if (path === 'popular') action = 'popular';
  else if (path === 'view' && request.method === 'POST') {
    action = 'view';
    method = 'POST';
    body = await request.text();
  } else {
    return Response.json({ error: 'API 路徑不存在' }, { status: 404 });
  }

  const target = new URL(gasUrl);
  target.searchParams.set('action', action);

  const upstream = await fetch(target.toString(), {
    method,
    headers: method === 'POST' ? { 'Content-Type': 'application/json' } : undefined,
    body
  });

  const text = await upstream.text();
  return new Response(text, {
    status: upstream.ok ? 200 : upstream.status,
    headers: { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store' }
  });
}
