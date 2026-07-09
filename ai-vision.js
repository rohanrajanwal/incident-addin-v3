/* ========================================
   Incident Reconstruction Engine — AI Vision client
   ========================================
   Speaks the incident AI proxy's JSON contract for:
     - analyzePhotos(photos, { party })  -> collision damage assessment
     - extractDocument(type, imageDataUrl) -> document OCR

   The proxy (dev: dev/ai-proxy.py; prod: Cloud Run function) holds the GenAI
   Gateway credential and talks to Claude Sonnet. This module never sees a key.

   It also DEFENSIVELY repairs model output: damage-zone and severity strings
   must match the UI vocabulary exactly, because app.js selects chips via
   [onclick*="<zone>"] and severity via <h4>.textContent === <severity>. Any
   value outside the controlled vocabulary is dropped/nulled so a bad model
   string can never silently break the Review screen. */
(function () {
  // MUST match the chip onclick values + severity <h4> text in index.html.
  const ZONES = [
    'Front Left', 'Front Center', 'Front Right',
    'Left Side', 'Right Side',
    'Rear Left', 'Rear Center', 'Rear Right',
    'Roof', 'Undercarriage / Frame'
  ];
  const SEVERITIES = ['Minor', 'Functional', 'Disabling'];

  const DOC_FIELDS = {
    license:      ['dlName', 'dlNumber', 'dlDob', 'dlAddress'],
    insurance:    ['name', 'policy', 'insurer', 'insDates'],
    registration: ['vin', 'plate', 'makeModel', 'regYear', 'regOwner'],
    police:       ['reportNumber', 'officerName', 'badgeNumber'],
    citation:     ['citationNumber', 'citationViolations']
  };

  const cfg = () => window.INCIDENT_AI || {};

  async function _post(path, body) {
    const c = cfg();
    if (!c.proxyUrl) throw new Error('No AI proxy configured (INCIDENT_AI.proxyUrl)');
    const base = c.proxyUrl.replace(/\/+$/, '');
    const ctrl = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), c.timeoutMs || 45000);
    try {
      const resp = await fetch(base + path, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ region: c.region || 'us', ...body }),
        signal: ctrl.signal
      });
      const text = await resp.text();
      let json = null;
      try { json = JSON.parse(text); } catch (e) { /* not JSON */ }
      if (!resp.ok || !json) {
        throw new Error(`Proxy HTTP ${resp.status}: ${(text || '').slice(0, 200)}`);
      }
      if (json.ok === false) throw new Error(json.error || 'Proxy returned error');
      return json.data != null ? json.data : json;
    } finally {
      clearTimeout(timer);
    }
  }

  function _cleanZones(arr) {
    if (!Array.isArray(arr)) return [];
    const seen = new Set();
    const out = [];
    arr.forEach(z => {
      if (ZONES.includes(z) && !seen.has(z)) { seen.add(z); out.push(z); }
    });
    return out;
  }

  function _cleanSeverity(s) {
    return SEVERITIES.includes(s) ? s : null;
  }

  function _str(v) {
    const s = (v == null ? '' : String(v)).trim();
    if (!s) return null;
    // Models sometimes echo "null"/"n/a"/"not visible" as text — treat as empty.
    if (/^(null|n\/?a|none|not visible|not legible|unknown|unreadable)$/i.test(s)) return null;
    return s;
  }

  // ---- Direct-to-gateway mode (prototype) ----
  // When a runtime token+baseUrl is configured (AI Setup screen -> localStorage),
  // the add-in calls the GenAI Gateway's /chat/completions directly from the
  // browser (proven to work from the *.geotab.com add-in origin). Prompts are
  // ported from dev/ai-proxy.py so output matches the proxy's contract exactly.
  const SEVERITY_DEFS =
    '- Minor: scratches, scuffs, or small dents; vehicle fully driveable.\n' +
    '- Functional: broken glass, hanging bumpers, or light damage; may still be driveable.\n' +
    '- Disabling: structural/frame damage, wheel misalignment, or airbag deployment; not driveable.';

  const DOC_SPECS = {
    license:      ["a driver's license", '{ "dlName": str|null, "dlNumber": str|null, "dlDob": "YYYY-MM-DD"|null, "dlAddress": str|null, "confidence": 0-1 }'],
    insurance:    ['an auto insurance card', '{ "name": "policyholder name"|null, "policy": "policy number"|null, "insurer": "insurance company"|null, "insDates": "coverage period e.g. 01/2026 - 01/2027"|null, "confidence": 0-1 }'],
    registration: ['a vehicle registration document', '{ "vin": "17-char VIN"|null, "plate": "license plate"|null, "makeModel": "year make model"|null, "regYear": "registration year"|null, "regOwner": "registered owner"|null, "confidence": 0-1 }'],
    police:       ['a police accident/incident report', '{ "reportNumber": str|null, "officerName": str|null, "badgeNumber": str|null, "confidence": 0-1 }'],
    citation:     ['a traffic citation / ticket', '{ "citationNumber": str|null, "citationViolations": "violation description"|null, "confidence": 0-1 }']
  };

  function _analyzeMessages(images, party, sceneFrames) {
    const zonesList = ZONES.map(z => '"' + z + '"').join(', ');
    const hasFrames = !!(sceneFrames && sceneFrames.length);
    let role, schema, extra;
    if (party === 'third') {
      role = 'the OTHER (third-party) vehicle involved in the collision';
      schema = '{ "thirdPartyDetected": true|false, "thirdPartyVehicleType": "e.g. Sedan, SUV, Pickup Truck, Box Truck, Motorcycle"|null, "thirdPartyPlate": "license plate"|null, "thirdPartyVIN": "17-char VIN"|null, "damageZones": { "third": [ <zones> ] }, "severityThird": "Minor"|"Functional"|"Disabling"|null, "confidenceScores": { "vehicleType":0-1, "plate":0-1, "vin":0-1, "severity":0-1, "damageZones":0-1 } }';
      extra = 'Read the license plate and VIN ONLY if clearly legible; otherwise null. NEVER guess or fabricate a plate or VIN.';
    } else {
      role = "the USER'S OWN vehicle";
      schema = '{ "damageZones": { "first": [ <zones> ] }, "severityFirst": "Minor"|"Functional"|"Disabling"|null, "weather": "short phrase e.g. Clear, Rain, Snow"|null, "roadConditions": "short phrase e.g. Dry, Wet, Icy"|null, "confidenceScores": { "severity":0-1, "damageZones":0-1 } }';
      extra = 'Infer weather/road conditions only if visible in the scene; otherwise null.';
    }
    const sceneNote = hasFrames
      ? ' Some of the images are frames sampled from a 360° walkaround video of the overall scene — use them for weather/road conditions and any additional visible damage; they may also show surroundings and other vehicles.'
      : '';
    const system =
      'You are an expert vehicle collision damage assessor. You will receive one or more images of ' + role + '. Assess visible damage.' + sceneNote + '\n\n' +
      'Damage zones MUST be chosen from this exact list (verbatim strings): [' + zonesList + ']. Include a zone only if it shows visible damage.\n\n' +
      'Severity levels (choose one overall):\n' + SEVERITY_DEFS + '\n\n' + extra + '\n\n' +
      'Respond with ONLY a JSON object in exactly this schema (no prose, no markdown):\n' + schema;
    const userContent = [];
    if (images.length) {
      userContent.push({ type: 'text', text: 'Close-up photos of the vehicle:' });
      images.forEach(u => userContent.push({ type: 'image_url', image_url: { url: u } }));
    }
    if (hasFrames) {
      userContent.push({ type: 'text', text: 'Frames from the 360° walkaround video of the collision scene:' });
      sceneFrames.forEach(u => userContent.push({ type: 'image_url', image_url: { url: u } }));
    }
    return [{ role: 'system', content: system }, { role: 'user', content: userContent }];
  }

  function _ocrMessages(type, image) {
    const spec = DOC_SPECS[type];
    const system =
      'You are an OCR and data-extraction engine. The image is ' + spec[0] + '. Extract the requested fields ' +
      'verbatim from what is legible. If a field is not present or not clearly legible, return null for it — ' +
      'NEVER guess or fabricate. Respond with ONLY a JSON object in exactly this schema (no prose, no markdown):\n' + spec[1];
    return [
      { role: 'system', content: system },
      { role: 'user', content: [{ type: 'text', text: 'Extract fields from this document:' }, { type: 'image_url', image_url: { url: image } }] }
    ];
  }

  async function _callGatewayDirect(messages) {
    const d = cfg().direct;
    if (!d || !d.token || !d.baseUrl) throw new Error('No direct AI config');
    const url = d.baseUrl.replace(/\/+$/, '') + '/chat/completions';
    const ctrl = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), cfg().timeoutMs || 45000);
    try {
      const resp = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'api-key': d.token },
        // NOTE: no response_format — claude-sonnet via the gateway's Vertex backend
        // rejects OpenAI's response_format:json_object (400). The prompts demand
        // JSON-only and the parser below strips ```json fences, so this is fine.
        body: JSON.stringify({ model: 'claude-sonnet', max_tokens: 1024, temperature: 0, messages }),
        signal: ctrl.signal
      });
      const text = await resp.text();
      if (!resp.ok) throw new Error('Gateway HTTP ' + resp.status + ': ' + (text || '').slice(0, 200));
      const j = JSON.parse(text);
      let content = (j.choices && j.choices[0] && j.choices[0].message && j.choices[0].message.content) || '';
      if (Array.isArray(content)) content = content.map(p => (p && p.text) || '').join('');
      content = String(content).trim().replace(/^```(?:json)?\s*/, '').replace(/\s*```$/, '').trim();
      try { return JSON.parse(content); }
      catch (e) { const m = content.match(/\{[\s\S]*\}/); if (m) return JSON.parse(m[0]); throw new Error('Model did not return JSON'); }
    } finally {
      clearTimeout(timer);
    }
  }

  /**
   * Analyze collision photos for ONE vehicle.
   * @param {Array<string|null>} photos - data URLs (nulls allowed/ignored)
   * @param {{party?: 'first'|'third'}} opts - 'first' = user's vehicle, 'third' = other vehicle
   * @returns object shaped like getMockAIResults() so applyAIResults() works unchanged.
   */
  async function analyzePhotos(photos, opts = {}) {
    const party = opts.party === 'third' ? 'third' : 'first';
    const imgs = (photos || []).filter(Boolean);
    const frames = (opts.sceneFrames || []).filter(Boolean);
    if (!imgs.length && !frames.length) throw new Error('No photos to analyze');

    const d = cfg().direct;
    const raw = (d && d.token)
      ? await _callGatewayDirect(_analyzeMessages(imgs, party, frames))
      : await _post('/analyze-photos', { photos: imgs, party, sceneFrames: frames });
    const dz = (raw && raw.damageZones) || {};

    return {
      thirdPartyDetected: !!raw.thirdPartyDetected,
      damageZones: {
        first: _cleanZones(dz.first),
        third: _cleanZones(dz.third)
      },
      severityFirst: _cleanSeverity(raw.severityFirst),
      severityThird: _cleanSeverity(raw.severityThird),
      thirdPartyVehicleType: _str(raw.thirdPartyVehicleType),
      thirdPartyVIN: _str(raw.thirdPartyVIN),
      thirdPartyPlate: _str(raw.thirdPartyPlate),
      weather: _str(raw.weather),
      roadConditions: _str(raw.roadConditions),
      confidenceScores: (raw.confidenceScores && typeof raw.confidenceScores === 'object')
        ? raw.confidenceScores : {}
    };
  }

  /**
   * OCR a single document image.
   * @param {'license'|'insurance'|'registration'|'police'|'citation'} type
   * @param {string} imageDataUrl
   * @returns object with exactly the field keys the app reads (null when illegible).
   */
  async function extractDocument(type, imageDataUrl) {
    const fields = DOC_FIELDS[type];
    if (!fields) throw new Error('Unknown document type: ' + type);
    if (!imageDataUrl) throw new Error('No image to OCR');

    const d = cfg().direct;
    const raw = (d && d.token)
      ? await _callGatewayDirect(_ocrMessages(type, imageDataUrl))
      : await _post('/ocr-document', { type, image: imageDataUrl });
    const out = {};
    fields.forEach(k => { out[k] = _str(raw ? raw[k] : null); });
    if (raw && raw.confidence != null) out._confidence = raw.confidence;
    return out;
  }

  window.IncidentAIVision = { analyzePhotos, extractDocument, ZONES, SEVERITIES, DOC_FIELDS };
})();
