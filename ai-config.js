/* ========================================
   Incident Reconstruction Engine — AI/OCR runtime config
   ========================================
   Photo analysis + document OCR ALWAYS use the real vision model via the
   incident AI proxy. There is no mock fallback: if the proxy is unreachable,
   the relevant fields are simply left for manual entry (never fabricated).

   The proxy URL is resolved automatically by environment so it "just works"
   and survives the dev harness "Reset Data" button (which clears localStorage):
     - localhost / 127.0.0.1  -> the local dev proxy (dev/ai-proxy.py)
     - anywhere else (prod)    -> PROD_PROXY_URL (Phase B Cloud Run deployment)

   An optional localStorage override can point at a different proxy/region:
     localStorage.setItem('INCIDENT_AI_OVERRIDE',
       JSON.stringify({ proxyUrl: 'https://my-proxy.example', region: 'ca' })); */
window.INCIDENT_AI = (function () {
  const LOCAL_PROXY_URL = 'http://localhost:8788';
  const PROD_PROXY_URL  = ''; // TODO Phase B: set to the deployed Cloud Run proxy URL

  const host = location.hostname;
  const isLocal = host === 'localhost' || host === '127.0.0.1';

  const cfg = {
    proxyUrl: isLocal ? LOCAL_PROXY_URL : PROD_PROXY_URL,
    region: 'us',
    timeoutMs: 45000
  };

  try {
    const override = JSON.parse(localStorage.getItem('INCIDENT_AI_OVERRIDE') || 'null');
    if (override && typeof override === 'object') Object.assign(cfg, override);
  } catch (e) {
    console.warn('[AI Config] Failed to read override:', e);
  }

  // Direct-to-gateway mode (prototype): the token + gateway base URL are entered
  // at RUNTIME via the in-app AI Setup screen and stored only in this device's
  // localStorage — never committed to the (public) bundle. When present, the
  // add-in calls the GenAI Gateway directly (no proxy needed). Used for the
  // single-user prototype demo; production must use a hosted proxy + service account.
  try {
    const direct = JSON.parse(localStorage.getItem('INCIDENT_AI_DIRECT') || 'null');
    if (direct && direct.token && direct.baseUrl) {
      cfg.direct = { token: direct.token, baseUrl: direct.baseUrl, region: direct.region || cfg.region };
    }
  } catch (e) {
    console.warn('[AI Config] Failed to read direct config:', e);
  }

  console.log('[AI Config] mode:', cfg.direct ? 'direct-to-gateway' : (cfg.proxyUrl ? 'proxy' : 'none — manual entry only'), '| region:', cfg.region);
  return cfg;
})();
