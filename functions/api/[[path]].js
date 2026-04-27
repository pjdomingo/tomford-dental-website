/**
 * TomFord Dental — Cloudflare Pages Function: API Proxy
 *
 * Proxies all /api requests to the Google Apps Script Web App.
 * The Apps Script URL lives in the SCRIPT_URL environment variable
 * (set in Cloudflare Pages → Settings → Environment Variables).
 *
 * To update the Apps Script URL after a redeploy:
 *   1. Cloudflare Dashboard → tomford-dental-website → Settings → Environment Variables
 *   2. Update SCRIPT_URL to the new Apps Script URL
 *   3. Done — no code changes, no git push needed.
 *
 * Frontend always calls /api — never the Apps Script URL directly.
 */

export async function onRequest(context) {
  const { request, env } = context;
  const scriptUrl = (env.SCRIPT_URL || '').trim();

  // OPTIONS preflight (CORS)
  if (request.method === 'OPTIONS') {
    return new Response(null, {
      headers: {
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type',
      },
    });
  }

  // Guard: env var not set
  if (!scriptUrl) {
    return new Response(
      JSON.stringify({ error: 'SCRIPT_URL environment variable is not configured.' }),
      { status: 500, headers: { 'Content-Type': 'application/json' } }
    );
  }

  // Build target URL — preserve query string from incoming request
  const incoming   = new URL(request.url);
  const targetUrl  = scriptUrl + incoming.search;

  const fetchInit = {
    method:   request.method,
    redirect: 'follow', // Apps Script GET requests redirect to googleusercontent.com
  };

  // Forward POST body as URL-encoded (Apps Script reads e.parameter from either format)
  if (request.method === 'POST') {
    const formData = await request.formData();
    const params   = new URLSearchParams();
    for (const [key, value] of formData.entries()) {
      params.append(key, value);
    }
    fetchInit.body    = params.toString();
    fetchInit.headers = { 'Content-Type': 'application/x-www-form-urlencoded' };
  }

  try {
    const upstream = await fetch(targetUrl, fetchInit);
    const text     = await upstream.text();

    return new Response(text, {
      status: 200,
      headers: {
        'Content-Type':                'application/json',
        'Cache-Control':               'no-store',
        'Access-Control-Allow-Origin': '*',
      },
    });
  } catch (err) {
    return new Response(
      JSON.stringify({ error: 'Proxy error: ' + err.message }),
      { status: 502, headers: { 'Content-Type': 'application/json' } }
    );
  }
}
