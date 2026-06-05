/* ========================================
   Incident Reconstruction Engine
   Main Application Logic
   ======================================== */

const app = {
  // How far back the incidents list looks for collision ExceptionEvents.
  // MyGeotab's "Last month" filter is ~calendar-month/looser, so a strict 30-day
  // window can hide events that still appear in MyGeotab (e.g. an event from 31 days ago).
  INCIDENT_LOOKBACK_DAYS: 90,

  // State
  currentScreen: 'incidents',
  api: null,
  state: null,
  reportData: {
    answers: {},
    photosYours: [],
    photosThird: [],
    sceneVideo: null,
    damageZones: { first: [], third: [] },
    severityFirst: null,
    severityThird: null,
    // Flat ocr object keeps backward-compat for formatReportText / populateReview / AddInData submission
    // (which still read ocr.name, ocr.policy, ocr.vin, ocr.plate). New per-doc fields added inline.
    ocr: {
      // From Driver's License
      dlName: '', dlNumber: '', dlDob: '', dlAddress: '',
      // From Insurance Card — name + policy + vin + plate are legacy keys kept for compat
      name: '', policy: '', insurer: '', insDates: '',
      // From Vehicle Registration
      vin: '', plate: '', makeModel: '', regYear: '', regOwner: ''
    },
    // Captured document images (uploaded as MediaFile attachments)
    docLicense: null,
    docInsurance: null,
    docRegistration: null,
    thirdPartyPhone: { countryCode: '+1', number: '' },
    narrative: '',
    occupancy: { yourVehicle: 1, yourInjuries: null, thirdVehicle: 1, thirdInjuries: null },
    witnesses: { hasWitnesses: null, name: '', phone: { countryCode: '+1', number: '' } },
    policeReport: { filed: null, document: null, violations: null, citations: null, citationDoc: null, reportNumber: '', officerName: '', badgeNumber: '', citationNumber: '', citationViolations: '' },
    propertyDamageInfo: { damaged: null, photo: null, propertyName: '', address: '', ownerName: '', ownerPhone: { countryCode: '+1', number: '' } },
    context: {},
    aiResults: null,
    corrections: []
  },
  _sceneVideoBlob: null,
  _selectedExceptionEventId: null,
  _eventsCache: {},

  // ---- Geotab Add-in Lifecycle ----
  initializeAddin() {
    // This is called by the Geotab Drive harness or dev harness
    window.geotab = window.geotab || {};
    window.geotab.addin = window.geotab.addin || {};
    window.geotab.addin.incidentReport = () => ({
      initialize: (api, state, callback) => {
        app.api = api;
        app.state = state;
        app.onInitialized();
        callback();
      },
      focus: (api, state) => {
        app.api = api;
        app.state = state;
        app.onFocus();
      },
      blur: () => {
        app.saveProgress();
      }
    });
  },

  onInitialized() {
    console.log('[Incident Add-in] Initialized');
    this.setupOfflineDetection();
    this.loadSavedProgress();
    this.initDamageZoneClicks();
    this._updateDriverGreeting();
    this.loadIncidents();
  },

  onFocus() {
    console.log('[Incident Add-in] Focused');
    this._updateDriverGreeting();
    // Refresh incidents when app comes into focus (may have new events)
    if (this.currentScreen === 'incidents') {
      this.loadIncidents();
    }
  },

  async _updateDriverGreeting() {
    if (!this.state) return;

    // 1. Prefer driver name (the person logged into the device in Drive)
    const driver = this.state.driver;
    if (driver && driver.name) {
      this._applyGreetingName(driver.name);
      return;
    }

    // 2. Query the User entity from the API — has proper firstName/lastName fields
    if (this.api) {
      try {
        const userName = await new Promise((resolve) => {
          if (typeof this.api.getSession === 'function') {
            try { this.api.getSession((s) => resolve(s?.userName || this.state.userName || null)); }
            catch (e) { resolve(this.state.userName || null); }
          } else {
            resolve(this.state.userName || null);
          }
        });

        if (userName) {
          const users = await new Promise((resolve, reject) => {
            this.api.call('Get', { typeName: 'User', search: { name: userName }, resultsLimit: 1 }, resolve, reject);
          });
          if (users && users.length > 0) {
            const u = users[0];
            const firstName = u.firstName || u.name?.split(' ')[0] || '';
            const lastName = u.lastName || u.name?.split(' ')[1] || '';
            if (firstName) {
              this._applyGreetingName(firstName + (lastName ? ' ' + lastName : ''));
              return;
            }
          }
        }
      } catch (e) {
        console.warn('[Greeting] API user lookup failed:', e);
      }
    }

    // 3. Last resort: parse the email local-part
    const email = this.state.userName || '';
    if (email) {
      const local = email.split('@')[0];
      const part = local.split(/[._-]/)[0];
      if (part) this._applyGreetingName(part.charAt(0).toUpperCase() + part.slice(1).toLowerCase());
    }
  },

  _applyGreetingName(fullName) {
    const parts = fullName.trim().split(' ');
    const firstName = parts[0];
    this.setEl('driverFirstName', firstName);
    const initials = firstName[0] + (parts[1] ? parts[1][0] : firstName[1] || '');
    this.setEl('driverInitials', initials.toUpperCase());
  },

  // ---- Incidents List ----
  async loadIncidents() {
    const currentList = document.getElementById('currentIncidentsList');
    const pastHeader = document.getElementById('pastIncidentsHeader');
    const pastSubtitle = document.getElementById('pastIncidentsSubtitle');
    const pastList = document.getElementById('pastIncidentsList');
    if (!currentList) return;

    // Show loading
    currentList.innerHTML = `
      <div style="text-align:center;padding:28px 0">
        <div class="ai-spinner"></div>
        <p style="margin-top:12px;font-size:13px;color:var(--text-secondary)">Loading incidents…</p>
      </div>`;

    if (!this.api || this.api._isMock) {
      this.renderFallbackIncidents(currentList, pastHeader, pastSubtitle, pastList);
      return;
    }

    try {
      const deviceId = this.state?.device?.id;
      console.log('[Incidents] state.device:', JSON.stringify(this.state?.device || null));
      if (!deviceId) {
        console.warn('[Incidents] No deviceId in state — cannot load incidents');
        this.renderFallbackIncidents(currentList, pastHeader, pastSubtitle, pastList);
        return;
      }

      const now = new Date();
      const lookbackStart = new Date(now - this.INCIDENT_LOOKBACK_DAYS * 24 * 60 * 60 * 1000);

      // Fetch rules, events, and existing reports in parallel
      const [allRules, events, existingReports] = await Promise.all([
        new Promise((resolve) =>
          this.api.call('Get', { typeName: 'Rule' }, resolve, () => resolve([]))
        ),
        new Promise((resolve, reject) =>
          this.api.call('Get', {
            typeName: 'ExceptionEvent',
            search: {
              deviceSearch: { id: deviceId },
              fromDate: lookbackStart.toISOString(),
              toDate: now.toISOString()
            }
          }, resolve, reject)
        ),
        new Promise((resolve) =>
          this.api.call('Get', {
            typeName: 'AddInData',
            search: { addInId: 'aIncidentReport001' }
          }, resolve, () => resolve([]))
        )
      ]);

      // Build a map of ruleId → full rule object (for name lookup)
      const ruleMap = {};
      (allRules || []).forEach(r => { ruleMap[r.id] = r; });

      // Identify collision/incident rules by two signals:
      // 1. System rule IDs — Geotab's built-in collision rules have PascalCase descriptive IDs
      //    like "RuleEnhancedMajorCollisionId", "RulePossibleCollisionId". Custom rules use
      //    base64 GUID IDs (e.g. "aXy123...") which won't match.
      // 2. Exact known rule names (case-insensitive) — for any custom collision rules.
      // Substring matching for "collision" was too broad — caught notification/alert rules on g560.
      const KNOWN_COLLISION_NAMES = ['possible collision', 'major collision', 'minor collision', 'near collision', 'enhanced major collision', 'enhanced minor collision', 'accident level event'];
      const isCollisionRule = (r) => {
        if ((r.id || '').includes('Collision')) return true;
        const n = (r.name || '').toLowerCase().trim();
        return KNOWN_COLLISION_NAMES.includes(n);
      };
      const matchedRules = (allRules || []).filter(isCollisionRule);
      const collisionRuleIds = new Set(matchedRules.map(r => r.id));
      console.log('[Incidents] matched collision rules:', matchedRules.map(r => `${r.name} (${r.id})`));
      console.log('[Incidents] collision rule IDs:', [...collisionRuleIds]);
      console.log('[Incidents] total events fetched:', (events || []).length);
      console.log('[Incidents] sample event rule:', JSON.stringify(events?.[0]?.rule || null));

      // Build set of already-reported exception event IDs
      const reportedIds = new Set(
        (existingReports || [])
          .map(r => r.details?.exceptionEventId)
          .filter(Boolean)
      );

      // Filter to collision rules only, and enrich each event with the full rule name
      // (ExceptionEvent responses only include rule.id, not rule.name)
      const collisionEvents = (events || []).filter(e => collisionRuleIds.has(e.rule?.id));
      collisionEvents.forEach(e => {
        const fullRule = ruleMap[e.rule?.id];
        if (fullRule && e.rule) {
          e.rule.name = fullRule.name;
        }
      });

      // Cache events for context lookup when starting a report
      this._eventsCache = {};
      collisionEvents.forEach(e => { this._eventsCache[e.id] = e; });

      // Sort newest first (by activeFrom timestamp, with event ID as tiebreaker for stability)
      const sorted = collisionEvents.sort((a, b) => {
        const diff = new Date(b.activeFrom) - new Date(a.activeFrom);
        if (diff !== 0) return diff;
        return (b.id || '').localeCompare(a.id || '');
      });
      console.log('[Incidents] sorted events:', sorted.map(e => ({
        id: e.id?.slice(0, 12),
        activeFrom: e.activeFrom,
        ruleName: e.rule?.name
      })));

      const unreported = sorted.filter(e => !reportedIds.has(e.id));
      const reported = sorted.filter(e => reportedIds.has(e.id));

      this.renderIncidentsList(unreported, reported, currentList, pastHeader, pastSubtitle, pastList);
    } catch (err) {
      console.warn('[Incidents] Failed to load from API:', err);
      this.renderFallbackIncidents(currentList, pastHeader, pastSubtitle, pastList);
    }
  },

  renderIncidentsList(unreported, reported, currentList, pastHeader, pastSubtitle, pastList) {
    currentList.innerHTML = '';

    if (unreported.length === 0) {
      currentList.innerHTML = `
        <div style="text-align:center;padding:28px 16px;color:var(--text-secondary)">
          <p style="font-size:14px">No open incidents found in the last ${this.INCIDENT_LOOKBACK_DAYS} days.</p>
          <button class="btn btn-secondary" style="max-width:220px;margin:14px auto 0;display:block;font-size:13px"
            onclick="app.injectTestIncident()">+ Add Test Incident</button>
        </div>`;
    } else {
      unreported.forEach(event => {
        currentList.appendChild(this.buildIncidentCard(event, false));
      });
    }

    if (reported.length > 0) {
      pastHeader.style.display = '';
      pastSubtitle.style.display = '';
      pastList.innerHTML = '';
      reported.forEach(event => {
        pastList.appendChild(this.buildIncidentCard(event, true));
      });
    } else {
      pastHeader.style.display = 'none';
      pastSubtitle.style.display = 'none';
      pastList.innerHTML = '';
    }
  },

  // Classify a Geotab rule into a human-readable event type.
  // Checks both rule.id (Geotab system rule IDs, always present) and rule.name
  // (display name, sometimes null). System IDs like "RuleEnhancedMajorCollisionId"
  // are the most reliable signal.
  _classifyIncident(ruleName, ruleId) {
    const name = (ruleName || '').toLowerCase();
    const id   = (ruleId   || '').toLowerCase();

    if (id.includes('major')   || name.includes('major'))              return { type: 'Major Collision',    cls: 'severity-major' };
    if (id.includes('minor')   || name.includes('minor'))              return { type: 'Minor Collision',    cls: 'severity-minor' };
    if (id.includes('nearmiss') || id.includes('near_miss') ||
        id.includes('near')    || name.includes('near') ||
        name.includes('miss'))                                          return { type: 'Near Collision',     cls: 'severity-near' };
    if (id.includes('possible')  || name.includes('possible'))         return { type: 'Possible Collision', cls: 'severity-possible' };
    if (id.includes('collision') || name.includes('collision'))         return { type: 'Possible Collision', cls: 'severity-possible' };
    if ((id.includes('harsh') && id.includes('brak')) ||
        (name.includes('harsh') && name.includes('brak')))             return { type: 'Harsh Braking',      cls: 'severity-possible' };
    if (id.includes('accident') || name.includes('accident'))          return { type: 'Accident',           cls: 'severity-major' };

    // Fallback: use the rule name itself if meaningful, otherwise null (no subtype)
    const displayName = ruleName && ruleName !== 'Incident' ? ruleName : null;
    return { type: displayName, cls: 'severity-default' };
  },

  // Geotab event IDs are long URL-safe base64 strings; show the first 8 chars
  // (same approach as git short hashes — enough to disambiguate events at a glance).
  _shortEventId(id) {
    if (!id) return '';
    // Strip leading 'a' that Geotab prepends to most entity IDs, then take 8 chars
    const trimmed = id.startsWith('a') ? id.slice(1) : id;
    return trimmed.slice(0, 8);
  },

  buildIncidentCard(event, isCompleted) {
    const severity = this._classifyIncident(event.rule?.name, event.rule?.id);
    const title = severity.type ? `Incident — ${severity.type}` : 'Incident';
    // Round to nearest second before display — MyGeotab rounds milliseconds, JS truncates,
    // causing 1-second discrepancies for events stored at e.g. 20:51:20.710Z (shows :20 in JS, :21 in MYG)
    const rawMs = new Date(event.activeFrom).getTime();
    const date = new Date(Math.round(rawMs / 1000) * 1000);
    const dateStr = date.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
    const timeStr = date.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit', second: '2-digit' });
    const vehicleName = event.device?.name || this.state?.device?.name || '';
    const shortId = this._shortEventId(event.id);

    const badge = isCompleted
      ? '<span class="badge completed">Completed</span>'
      : '<span class="badge new">New</span>';

    const svgIcon = `<svg viewBox="0 0 24 24" stroke-linecap="round" stroke-linejoin="round"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>`;

    const calSvg = `<svg class="meta-svg" viewBox="0 0 24 24" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>`;
    const truckSvg = `<svg class="meta-svg" viewBox="0 0 24 24" stroke-linecap="round" stroke-linejoin="round"><rect x="1" y="3" width="15" height="13"/><polygon points="16 8 20 8 23 11 23 16 16 16 16 8"/><circle cx="5.5" cy="18.5" r="2.5"/><circle cx="18.5" cy="18.5" r="2.5"/></svg>`;
    const hashSvg = `<svg class="meta-svg" viewBox="0 0 24 24" stroke-linecap="round" stroke-linejoin="round"><line x1="4" y1="9" x2="20" y2="9"/><line x1="4" y1="15" x2="20" y2="15"/><line x1="10" y1="3" x2="8" y2="21"/><line x1="16" y1="3" x2="14" y2="21"/></svg>`;

    const continueBtn = isCompleted ? '' : `
      <hr class="card-divider">
      <button class="btn btn-primary" style="width:100%" onclick="app.startReport('${event.id}')">Continue</button>`;

    const div = document.createElement('div');
    div.className = 'incident-card';
    div.innerHTML = `
      <div class="card-header">
        <div class="card-icon collision">${svgIcon}</div>
        <h3 style="flex:1;line-height:1.3">${this._escHtml(title)}</h3>
        ${badge}
      </div>
      <p class="description" style="font-size:12px;color:var(--text-muted);margin:2px 0 8px;font-family:ui-monospace,SFMono-Regular,Menlo,monospace">Event ${shortId}</p>
      <p class="meta-line">${calSvg} ${dateStr} at ${timeStr}</p>
      ${vehicleName ? `<p class="meta-line">${truckSvg} ${this._escHtml(vehicleName)}</p>` : ''}
      ${continueBtn}`;
    return div;
  },

  renderFallbackIncidents(currentList, pastHeader, pastSubtitle, pastList) {
    currentList.innerHTML = `
      <div style="text-align:center;padding:28px 16px;color:var(--text-secondary)">
        <p style="font-size:14px;margin-bottom:12px">Could not load incidents from database.</p>
        <button class="btn btn-primary" style="max-width:220px;margin:0 auto;display:block" onclick="app.startReport(null)">Start New Report</button>
        <button class="btn btn-secondary" style="max-width:220px;margin:10px auto 0;display:block;font-size:13px" onclick="app.injectTestIncident()">+ Add Test Incident</button>
      </div>`;
    pastHeader.style.display = 'none';
    pastSubtitle.style.display = 'none';
    pastList.innerHTML = '';
  },

  startReport(exceptionEventId) {
    this._selectedExceptionEventId = exceptionEventId;
    const event = exceptionEventId ? (this._eventsCache[exceptionEventId] || null) : null;
    this.reportData.context = {}; // reset context for new report
    this._fetchEventContext(event); // async, runs in background
    this.goTo('safety');
  },

  injectTestIncident() {
    // Creates a fake ExceptionEvent for UI/submission testing when no real events exist.
    // ExceptionEvents are system-generated and cannot be created via API directly.
    // This test event pre-seeds realistic context data so the full flow can be validated.
    // Remove or hide this button before any production rollout.
    const eventTime = new Date(Date.now() - 2 * 60 * 60 * 1000); // 2 hours ago
    const deviceName = this.state?.device?.name || 'Test Vehicle';

    const fakeEvent = {
      id: 'aTestIncident001',
      activeFrom: eventTime.toISOString(),
      rule: { name: 'Collision Detected (TEST)' },
      device: { id: this.state?.device?.id || null, name: deviceName }
    };

    // Cache so startReport can find it
    this._eventsCache['aTestIncident001'] = fakeEvent;

    // Pre-seed context with realistic dummy data so context/review screens are populated
    this.reportData.context = {
      eventTime,
      locationStr:    '1145 Eglinton Ave E, Toronto, ON',
      latitude:        43.7085,
      longitude:      -79.3398,
      speedKmh:        52,
      gForce:          2.8
    };

    // Replace the empty-state with a card in the list
    const currentList = document.getElementById('currentIncidentsList');
    if (currentList) {
      currentList.innerHTML = '';
      currentList.appendChild(this.buildIncidentCard(fakeEvent, false));
    }
  },

  _escHtml(str) {
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  },

  // ---- Navigation ----
  goTo(screenId) {
    const current = document.querySelector('.screen.active');
    if (current) current.classList.remove('active');

    const next = document.getElementById('screen-' + screenId);
    if (next) {
      next.classList.add('active');
      this.currentScreen = screenId;
      this.updateHeader(screenId);
      window.scrollTo(0, 0);

      if (screenId === 'review') this.populateReview();
      if (screenId === 'context') { this.populateContextScreen(); }
      if (screenId === 'narrative') this.initNarrativeScreen();
      if (screenId === 'property-damage') this.initPropertyDamageScreen();
    }
  },

  updateHeader(screenId) {
    const titles = {
      'incidents': 'Incidents',
      'safety': 'Incident Reporting',
      'communication': 'Incident Reporting',
      'qualifying': 'Report Incident',
      'photos-yours': 'Report Incident',
      'photos-third': 'Report Incident',
      'damage-first': 'Report Incident',
      'severity-first': 'Report Incident',
      'severity-third': 'Report Incident',
      'damage-third': 'Report Incident',
      'documents': 'Report Incident',
      'narrative': 'Report Incident',
      'police-report': 'Report Incident',
      'property-damage': 'Report Incident',
      'context': 'Report Incident',
      'review': 'Report Incident',
      'success': 'Incident Reporting'
    };
    this.setEl('headerTitle', titles[screenId] || 'Incident Reporting');
  },

  // ---- Qualifying Questions ----
  setAnswer(key, value, btnEl) {
    this.reportData.answers[key] = value;

    // Toggle button selection
    const siblings = btnEl.parentElement.querySelectorAll('.toggle-btn');
    siblings.forEach(b => b.classList.remove('selected'));
    btnEl.classList.add('selected');

    if (key === 'thirdParty') {
      const hitRunGroup = document.getElementById('hitAndRunGroup');
      if (hitRunGroup) hitRunGroup.style.display = value ? '' : 'none';
      if (!value) this.reportData.answers.hitAndRun = null;
    }
  },

  qualifyingContinue() {
    this.goTo('photos-yours');
  },

  // ---- Narrative screen init ----
  initNarrativeScreen() {
    const show = !!this.reportData.answers.thirdParty;
    const g = document.getElementById('thirdOccupancyGroup');
    if (g) g.style.display = show ? '' : 'none';
  },

  narrativeBack() {
    // If no third party, docs was skipped — go back to your photos
    if (this.reportData.answers.thirdParty) {
      this.goTo('documents');
    } else {
      this.goTo('photos-yours');
    }
  },

  narrativeContinue() {
    this.reportData.narrative = document.getElementById('narrativeText').value;
    this.goTo('police-report');
  },

  // ---- Occupancy & witness helpers ----
  adjustOccupancy(party, delta) {
    const field = party === 'yours' ? 'yourVehicle' : 'thirdVehicle';
    const elId = party === 'yours' ? 'yourVehicleCount' : 'thirdVehicleCount';
    const current = this.reportData.occupancy[field];
    const next = Math.max(1, current + delta);
    this.reportData.occupancy[field] = next;
    this.setEl(elId, next);
  },

  setNarrativeAnswer(key, value, btnEl) {
    this.reportData.occupancy[key] = value;
    const siblings = btnEl.parentElement.querySelectorAll('.toggle-btn');
    siblings.forEach(b => b.classList.remove('selected'));
    btnEl.classList.add('selected');
  },

  setWitnessToggle(value, btnEl) {
    this.reportData.witnesses.hasWitnesses = value;
    const siblings = btnEl.parentElement.querySelectorAll('.toggle-btn');
    siblings.forEach(b => b.classList.remove('selected'));
    btnEl.classList.add('selected');
    const fields = document.getElementById('witnessFields');
    if (fields) fields.style.display = value ? '' : 'none';
  },

  // ---- Police Report ----
  setPoliceAnswer(key, value, btnEl) {
    this.reportData.policeReport[key] = value;
    const siblings = btnEl.parentElement.querySelectorAll('.toggle-btn');
    siblings.forEach(b => b.classList.remove('selected'));
    btnEl.classList.add('selected');
    if (key === 'filed') {
      const details = document.getElementById('policeReportDetails');
      if (details) details.style.display = value ? '' : 'none';
    }
    if (key === 'citations') {
      const citGroup = document.getElementById('citationDocGroup');
      if (citGroup) citGroup.style.display = value ? '' : 'none';
    }
  },

  capturePoliceReport() {
    this._captureDocFile('policeDocUpload', 'policeDocPreview', 'app.removePoliceDoc()', (data) => {
      this.reportData.policeReport.document = data;
      // Mock OCR — extract report #, officer, badge from the police report card
      this.reportData.policeReport.reportNumber = 'RPT-2026-04827';
      this.reportData.policeReport.officerName = 'Ofc. M. Williams';
      this.reportData.policeReport.badgeNumber = '4821';
      const r = document.getElementById('ocrResultsPolice');
      if (r) r.style.display = 'block';
      this.setEl('ocrPoliceReport', this.reportData.policeReport.reportNumber);
      this.setEl('ocrPoliceOfficer', this.reportData.policeReport.officerName);
      this.setEl('ocrPoliceBadge', this.reportData.policeReport.badgeNumber);
    });
  },

  removePoliceDoc() {
    this.reportData.policeReport.document = null;
    this.reportData.policeReport.reportNumber = '';
    this.reportData.policeReport.officerName = '';
    this.reportData.policeReport.badgeNumber = '';
    document.getElementById('policeDocUpload').style.display = '';
    document.getElementById('policeDocPreview').style.display = 'none';
    const r = document.getElementById('ocrResultsPolice');
    if (r) r.style.display = 'none';
  },

  captureCitationDoc() {
    this._captureDocFile('citationDocUpload', 'citationDocPreview', 'app.removeCitationDoc()', (data) => {
      this.reportData.policeReport.citationDoc = data;
      // Mock OCR — extract citation # and violations
      this.reportData.policeReport.citationNumber = 'CIT-2026-11432';
      this.reportData.policeReport.citationViolations = 'Following too closely';
      const r = document.getElementById('ocrResultsCitation');
      if (r) r.style.display = 'block';
      this.setEl('ocrCitationNumber', this.reportData.policeReport.citationNumber);
      this.setEl('ocrCitationViolations', this.reportData.policeReport.citationViolations);
    });
  },

  removeCitationDoc() {
    this.reportData.policeReport.citationDoc = null;
    this.reportData.policeReport.citationNumber = '';
    this.reportData.policeReport.citationViolations = '';
    document.getElementById('citationDocUpload').style.display = '';
    document.getElementById('citationDocPreview').style.display = 'none';
    const r = document.getElementById('ocrResultsCitation');
    if (r) r.style.display = 'none';
  },

  // ---- Property Damage ----
  setPropertyAnswer(key, value, btnEl) {
    this.reportData.propertyDamageInfo[key] = value;
    const siblings = btnEl.parentElement.querySelectorAll('.toggle-btn');
    siblings.forEach(b => b.classList.remove('selected'));
    btnEl.classList.add('selected');
    if (key === 'damaged') {
      const details = document.getElementById('propertyDamageDetails');
      if (details) details.style.display = value ? '' : 'none';
    }
  },

  initPropertyDamageScreen() {
    // Pre-fill address from GPS context if available
    const loc = this.reportData.context;
    const addrEl = document.getElementById('propertyAddress');
    if (addrEl && loc.locationStr && !addrEl.value) {
      addrEl.value = loc.locationStr;
      this.reportData.propertyDamageInfo.address = loc.locationStr;
    }
  },

  capturePropertyPhoto() {
    this._captureDocFile('propertyPhotoUpload', 'propertyPhotoPreview', 'app.removePropertyPhoto()', (data) => {
      this.reportData.propertyDamageInfo.photo = data;
    });
  },

  removePropertyPhoto() {
    this.reportData.propertyDamageInfo.photo = null;
    document.getElementById('propertyPhotoUpload').style.display = '';
    document.getElementById('propertyPhotoPreview').style.display = 'none';
  },

  // ---- Optional collapsible sections ----
  toggleSection(sectionId) {
    const body = document.getElementById(sectionId + 'Body');
    const chevron = document.getElementById(sectionId + 'Chevron');
    if (!body) return;
    const isOpen = body.style.display !== 'none';
    body.style.display = isOpen ? 'none' : '';
    if (chevron) chevron.style.transform = isOpen ? '' : 'rotate(180deg)';
  },

  // ---- 360° Scene Video ----
  captureSceneVideo() {
    // NOTE: Live video recording is disabled. Geotab Drive's iOS app is missing
    // NSMicrophoneUsageDescription in its Info.plist — iOS terminates the process
    // the moment anything requests microphone access. Library picker only.
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = 'video/*';
    input.style.cssText = 'position:fixed;top:-9999px;left:-9999px;opacity:0;';
    input.onchange = async (e) => {
      const file = e.target.files[0];
      if (file) {
        // iOS WKWebView issue: File references from <input type="file"> can become invalid
        // after the input is removed from the DOM. Materialize the bytes into a fresh Blob
        // immediately so we have a persistent reference for the submit.
        try {
          const buf = await file.arrayBuffer();
          this._sceneVideoBlob = new Blob([buf], { type: file.type || 'video/mp4' });
          this.reportData.sceneVideo = { name: file.name, size: file.size, type: file.type };
          document.getElementById('sceneVideoSlot').style.display = 'none';
          document.getElementById('sceneVideoPreview').style.display = 'flex';
          this.setEl('sceneVideoName', file.name);
          console.log('[Video] Materialized scene video:', file.name, file.size, 'bytes,', file.type);
        } catch (err) {
          console.error('[Video] Failed to materialize video file:', err);
          alert('Could not attach video — please try again or skip this step.');
        }
      }
      if (input.parentNode) input.parentNode.removeChild(input);
    };
    document.body.appendChild(input);
    input.click();
  },

  removeSceneVideo() {
    this.reportData.sceneVideo = null;
    this._sceneVideoBlob = null;
    document.getElementById('sceneVideoSlot').style.display = '';
    document.getElementById('sceneVideoPreview').style.display = 'none';
  },

  _startInAppVideoRecording() {
    // MediaRecorder with getUserMedia is not supported in iOS WKWebView (Geotab Drive).
    // Fall back to native file input which shows the iOS camera/video picker.
    const supportsMediaRecorder = (
      typeof MediaRecorder !== 'undefined' &&
      typeof navigator.mediaDevices !== 'undefined' &&
      typeof navigator.mediaDevices.getUserMedia === 'function'
    );

    if (!supportsMediaRecorder) {
      this._videoFallbackInput();
      return;
    }

    const overlay = document.createElement('div');
    overlay.className = 'video-record-overlay';
    overlay.innerHTML = `
      <div class="video-record-modal">
        <div class="video-record-header">
          <span>360° Scene Video</span>
          <button class="video-record-close" id="videoCloseBtn">&times;</button>
        </div>
        <div class="video-preview-container">
          <video id="videoPreviewStream" autoplay muted playsinline></video>
          <div class="video-record-indicator" id="videoRecordIndicator" style="display:none">
            <span class="video-rec-dot"></span> REC
          </div>
        </div>
        <div class="video-record-controls">
          <p id="videoRecordStatus" class="video-record-status">Position camera for a 360° walkthrough</p>
          <button class="video-rec-btn" id="videoRecordToggle">
            <span class="video-rec-icon"></span>
            <span id="videoRecordLabel">Start Recording</span>
          </button>
        </div>
      </div>
    `;
    document.body.appendChild(overlay);

    let mediaStream = null;
    let mediaRecorder = null;
    let recordedChunks = [];
    let isRecordingVideo = false;

    const videoEl = document.getElementById('videoPreviewStream');
    const statusEl = document.getElementById('videoRecordStatus');
    const toggleBtn = document.getElementById('videoRecordToggle');
    const labelEl = document.getElementById('videoRecordLabel');
    const indicatorEl = document.getElementById('videoRecordIndicator');

    const cleanup = () => {
      try {
        if (mediaRecorder && mediaRecorder.state !== 'inactive') mediaRecorder.stop();
        if (mediaStream) mediaStream.getTracks().forEach(t => t.stop());
      } catch (e) { /* ignore */ }
      overlay.remove();
    };

    document.getElementById('videoCloseBtn').onclick = cleanup;

    navigator.mediaDevices.getUserMedia({ video: { facingMode: 'environment' }, audio: true })
      .then(stream => {
        mediaStream = stream;
        videoEl.srcObject = stream;
      })
      .catch(err => {
        console.warn('[VideoRecord] Camera access denied:', err);
        overlay.remove();
        // Fall back to native file input
        this._videoFallbackInput();
      });

    toggleBtn.onclick = () => {
      if (!mediaStream) return;
      try {
        if (!isRecordingVideo) {
          recordedChunks = [];
          const mimeType = MediaRecorder.isTypeSupported('video/mp4;codecs=avc1')
            ? 'video/mp4'
            : MediaRecorder.isTypeSupported('video/webm;codecs=vp9')
              ? 'video/webm;codecs=vp9'
              : 'video/webm';

          try {
            mediaRecorder = new MediaRecorder(mediaStream, { mimeType });
          } catch (e) {
            mediaRecorder = new MediaRecorder(mediaStream);
          }

          mediaRecorder.ondataavailable = e => { if (e.data.size > 0) recordedChunks.push(e.data); };
          mediaRecorder.onstop = () => {
            try {
              const finalMime = mediaRecorder.mimeType || 'video/webm';
              const blob = new Blob(recordedChunks, { type: finalMime });
              const ext = finalMime.includes('mp4') ? 'mp4' : 'webm';
              const fileName = `scene_video_${Date.now()}.${ext}`;
              this._sceneVideoBlob = blob;
              this.reportData.sceneVideo = { name: fileName, size: blob.size };
              document.getElementById('sceneVideoSlot').style.display = 'none';
              document.getElementById('sceneVideoPreview').style.display = 'flex';
              this.setEl('sceneVideoName', fileName);
              if (mediaStream) mediaStream.getTracks().forEach(t => t.stop());
            } catch (e) { console.warn('[VideoRecord] onstop error:', e); }
            overlay.remove();
          };

          mediaRecorder.start(100);
          isRecordingVideo = true;
          indicatorEl.style.display = 'flex';
          toggleBtn.classList.add('recording');
          labelEl.textContent = 'Stop & Save';
          statusEl.textContent = 'Recording — walk around the scene slowly';
        } else {
          isRecordingVideo = false;
          statusEl.textContent = 'Processing video…';
          toggleBtn.disabled = true;
          mediaRecorder.stop();
        }
      } catch (e) {
        console.warn('[VideoRecord] Recording error:', e);
        cleanup();
        this._videoFallbackInput();
      }
    };
  },

  _videoFallbackInput() {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = 'video/*';
    input.capture = 'environment'; // opens camera directly in video mode
    input.style.cssText = 'position:fixed;top:-9999px;left:-9999px;opacity:0;';
    input.onchange = (e) => {
      const file = e.target.files[0];
      if (file) {
        this._sceneVideoBlob = file;
        this.reportData.sceneVideo = { name: file.name, size: file.size };
        document.getElementById('sceneVideoSlot').style.display = 'none';
        document.getElementById('sceneVideoPreview').style.display = 'flex';
        this.setEl('sceneVideoName', file.name);
      }
      if (input.parentNode) input.parentNode.removeChild(input);
    };
    document.body.appendChild(input);
    input.click();
  },

  // ---- Shared capture helpers ----
  // ---- Phone number formatting ----
  formatPhoneInput(inputEl, storeCallback) {
    let digits = inputEl.value.replace(/\D/g, '').slice(0, 10);
    let formatted = digits;
    if (digits.length > 6) {
      formatted = `(${digits.slice(0,3)}) ${digits.slice(3,6)}-${digits.slice(6)}`;
    } else if (digits.length > 3) {
      formatted = `(${digits.slice(0,3)}) ${digits.slice(3)}`;
    } else if (digits.length > 0) {
      formatted = `(${digits}`;
    }
    inputEl.value = formatted;
    storeCallback(digits);
  },

  _showDocMenu(onTake, onChoose) {
    const overlay = document.createElement('div');
    overlay.className = 'photo-menu-overlay';
    overlay.onclick = (e) => { if (e.target === overlay) overlay.remove(); };
    const menu = document.createElement('div');
    menu.className = 'photo-menu';
    menu.innerHTML = `
      <button class="photo-menu-btn">Take a Photo</button>
      <button class="photo-menu-btn">Choose from Photos</button>
      <button class="photo-menu-cancel">Cancel</button>
    `;
    const btns = menu.querySelectorAll('.photo-menu-btn');
    btns[0].onclick = () => { overlay.remove(); onTake(); };
    btns[1].onclick = () => { overlay.remove(); onChoose(); };
    menu.querySelector('.photo-menu-cancel').onclick = () => overlay.remove();
    overlay.appendChild(menu);
    document.body.appendChild(overlay);
  },

  _showVideoMenu(onRecord, onChoose) {
    const overlay = document.createElement('div');
    overlay.className = 'photo-menu-overlay';
    overlay.onclick = (e) => { if (e.target === overlay) overlay.remove(); };
    const menu = document.createElement('div');
    menu.className = 'photo-menu';
    menu.innerHTML = `
      <button class="photo-menu-btn">Take a Video</button>
      <button class="photo-menu-btn">Choose from Videos</button>
      <button class="photo-menu-cancel">Cancel</button>
    `;
    const btns = menu.querySelectorAll('.photo-menu-btn');
    btns[0].onclick = () => { overlay.remove(); onRecord(); };
    btns[1].onclick = () => { overlay.remove(); onChoose(); };
    menu.querySelector('.photo-menu-cancel').onclick = () => overlay.remove();
    overlay.appendChild(menu);
    document.body.appendChild(overlay);
  },

  _imageInput(useCamera, onFile) {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = 'image/*';
    // capture="environment" opens the camera directly in photo mode — works in Geotab Drive.
    // Do NOT let users slide to video from here (video mode needs mic, which may not be
    // permitted). For video, use _videoInput which opens the camera directly in video mode.
    if (useCamera) input.capture = 'environment';
    input.onchange = (e) => { if (e.target.files[0]) onFile(e.target.files[0]); };
    input.click();
  },

  _videoInput(useCamera, onFile) {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = 'video/*';
    // capture="environment" with accept="video/*" opens the camera directly in video mode.
    // This avoids the crash caused by sliding from photo mode to video mode mid-session,
    // because iOS handles mic permission upfront when opening in video mode directly.
    if (useCamera) input.capture = 'environment';
    input.onchange = (e) => { if (e.target.files[0]) onFile(e.target.files[0]); };
    input.click();
  },

  // Generic document capture: skips app-native menu, goes straight to OS native picker.
  // iOS shows "Take Photo / Photo Library / Browse" in one step.
  // Android shows camera + gallery in one step.
  // Shows a thumbnail of the chosen image in the preview element.
  _captureDocFile(uploadId, previewId, removeCall, onStore) {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = 'image/*';
    input.style.cssText = 'position:fixed;top:-9999px;left:-9999px;opacity:0;';
    input.onchange = (e) => {
      const file = e.target.files[0];
      if (file) {
        const blobUrl = URL.createObjectURL(file);
        const preview = document.getElementById(previewId);
        if (preview) {
          preview.innerHTML =
            `<img src="${blobUrl}" style="height:56px;width:auto;max-width:calc(100% - 48px);border-radius:6px;object-fit:cover;flex-shrink:0">` +
            `<button class="doc-remove-btn" onclick="event.stopPropagation();${removeCall}">&times;</button>`;
          document.getElementById(uploadId).style.display = 'none';
          preview.style.display = 'flex';
        }
        const reader = new FileReader();
        reader.onload = () => onStore(reader.result);
        reader.readAsDataURL(file);
      }
      if (input.parentNode) input.parentNode.removeChild(input);
    };
    document.body.appendChild(input);
    input.click();
  },

  // ---- Document capture with OCR (license, insurance, registration) ----
  // Unified pipeline: capture photo → store image → run type-specific OCR → render extracted fields.
  _docMap: {
    license:      { key: 'docLicense',      box: 'docLicenseBox',      preview: 'docLicensePreview',      results: 'ocrResultsLicense' },
    insurance:    { key: 'docInsurance',    box: 'docInsuranceBox',    preview: 'docInsurancePreview',    results: 'ocrResultsInsurance' },
    registration: { key: 'docRegistration', box: 'docRegistrationBox', preview: 'docRegistrationPreview', results: 'ocrResultsRegistration' }
  },

  captureDocOCR(type) {
    const m = this._docMap[type];
    if (!m) return;
    this._captureDocFile(m.box, m.preview, `app.removeDocOCR('${type}')`, async (data) => {
      this.reportData[m.key] = data;
      await this._runDocOCR(type, data);
    });
  },

  removeDocOCR(type) {
    const m = this._docMap[type];
    if (!m) return;
    this.reportData[m.key] = null;
    document.getElementById(m.box).style.display = '';
    document.getElementById(m.preview).style.display = 'none';
    const r = document.getElementById(m.results);
    if (r) r.style.display = 'none';
    this._clearDocFields(type);
  },

  async _runDocOCR(type, imageData) {
    // Mock OCR per document type. Replace with real GenAI Gateway call in Phase 3.
    const mock = {
      license: {
        dlName: 'Robert Johnson',
        dlNumber: 'D1234-5678-9012',
        dlDob: '1985-03-22',
        dlAddress: '742 Evergreen Terrace, Springfield, IL'
      },
      insurance: {
        name: 'Robert Johnson',
        policy: 'POL-789456123',
        insurer: 'State Farm',
        insDates: '01/2026 — 01/2027'
      },
      registration: {
        vin: '1FTFW1ET5DFC10042',
        plate: 'ABC 1234',
        makeModel: '2022 Ford F-150 SuperCrew',
        regYear: '2022',
        regOwner: 'Robert Johnson'
      }
    }[type] || {};
    // Merge into flat ocr object (preserves backward compat for formatReportText / AddInData)
    Object.assign(this.reportData.ocr, mock);
    this._renderDocFields(type);
  },

  _renderDocFields(type) {
    const r = document.getElementById(this._docMap[type].results);
    if (r) r.style.display = 'block';
    const o = this.reportData.ocr;
    if (type === 'license') {
      this.setEl('ocrDlName', o.dlName || '—');
      this.setEl('ocrDlNumber', o.dlNumber || '—');
      this.setEl('ocrDlDob', o.dlDob || '—');
      this.setEl('ocrDlAddress', o.dlAddress || '—');
    } else if (type === 'insurance') {
      this.setEl('ocrInsName', o.name || '—');
      this.setEl('ocrInsPolicy', o.policy || '—');
      this.setEl('ocrInsCompany', o.insurer || '—');
      this.setEl('ocrInsDates', o.insDates || '—');
    } else if (type === 'registration') {
      this.setEl('ocrRegVin', o.vin || '—');
      this.setEl('ocrRegPlate', o.plate || '—');
      this.setEl('ocrRegMakeModel', o.makeModel || '—');
      this.setEl('ocrRegOwner', o.regOwner || '—');
    }
  },

  _clearDocFields(type) {
    const o = this.reportData.ocr;
    if (type === 'license') { o.dlName = ''; o.dlNumber = ''; o.dlDob = ''; o.dlAddress = ''; }
    else if (type === 'insurance') { o.name = ''; o.policy = ''; o.insurer = ''; o.insDates = ''; }
    else if (type === 'registration') { o.vin = ''; o.plate = ''; o.makeModel = ''; o.regYear = ''; o.regOwner = ''; }
  },

  // ---- Photo Capture ----
  showPhotoMenu(party, index) {
    // Always use the OS-native file picker (Photo Library / Take Photo / Choose File).
    // Bypasses both our custom menu AND the Drive SDK's "Adding Image..." modal.
    // The OS picker handles camera + library in one menu natively.
    this._triggerPhotoInput(party, index, false);
  },

  capturePhotoFromCamera(party, index) {
    this._triggerPhotoInput(party, index, true);
  },

  capturePhotoFromLibrary(party, index) {
    this._triggerPhotoInput(party, index, false);
  },

  async _triggerPhotoInput(party, index, useCamera) {
    // On Geotab Drive mobile, use the native camera API — avoids iOS permission issues
    if (useCamera && this.api?.mobile?.exists() && typeof this.api.mobile.camera?.takePicture === 'function') {
      try {
        const dataUrl = await this.api.mobile.camera.takePicture();
        if (dataUrl) {
          const processed = await this._processPhotoDataUrl(dataUrl);
          this.renderPhotoSlot(party, index, processed);
          const arr = party === 'yours' ? this.reportData.photosYours : this.reportData.photosThird;
          arr[index] = processed;
        }
      } catch (e) {
        console.warn('[Photo] Native camera failed, falling back to file input:', e);
        this._fileInputPhoto(false, (file) => this._storePhotoFile(file, party, index));
      }
      return;
    }
    // Library picker or desktop fallback
    this._fileInputPhoto(useCamera, (file) => this._storePhotoFile(file, party, index));
  },

  _fileInputPhoto(useCamera, onFile) {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = 'image/*';
    if (useCamera) input.capture = 'environment';
    input.style.cssText = 'position:fixed;top:-9999px;left:-9999px;opacity:0;';
    input.onchange = (e) => {
      if (e.target.files[0]) onFile(e.target.files[0]);
      if (input.parentNode) input.parentNode.removeChild(input);
    };
    document.body.appendChild(input);
    input.click();
  },

  _storePhotoFile(file, party, index) {
    const blobUrl = URL.createObjectURL(file);
    this.renderPhotoSlot(party, index, blobUrl);
    const reader = new FileReader();
    reader.onload = () => {
      const arr = party === 'yours' ? this.reportData.photosYours : this.reportData.photosThird;
      arr[index] = reader.result;
    };
    reader.readAsDataURL(file);
  },

  async _processPhotoDataUrl(dataUrl) {
    // Drive on iOS can return octet-stream — draw to canvas to get a real JPEG data URL
    if (!dataUrl.startsWith('data:image/')) {
      return new Promise((resolve) => {
        const img = new Image();
        img.onload = () => {
          const canvas = document.createElement('canvas');
          canvas.width = img.naturalWidth;
          canvas.height = img.naturalHeight;
          canvas.getContext('2d').drawImage(img, 0, 0);
          resolve(canvas.toDataURL('image/jpeg', 0.85));
        };
        img.onerror = () => resolve(dataUrl);
        img.src = dataUrl;
      });
    }
    return dataUrl;
  },

  renderPhotoSlot(party, index, dataUrl) {
    const gridId = party === 'yours' ? 'photoGridYours' : 'photoGridThird';
    const grid = document.getElementById(gridId);
    const slot = grid.children[index];
    if (!slot) return;
    slot.innerHTML = `
      <img src="${dataUrl}" alt="Photo ${index + 1}">
      <button class="remove-photo" onclick="event.stopPropagation(); app.removePhoto('${party}',${index})">&times;</button>
    `;
  },

  removePhoto(party, index) {
    const arr = party === 'yours' ? this.reportData.photosYours : this.reportData.photosThird;
    arr[index] = null;
    const labelsYours = ['Front View', 'Rear View', 'Left Side', 'Right Side', 'Damage Close-up'];
    const labelsThird = ['Front View', 'Rear View', 'Left Side', 'Right Side', 'Damage Close-up'];
    const labels = party === 'yours' ? labelsYours : labelsThird;
    const gridId = party === 'yours' ? 'photoGridYours' : 'photoGridThird';
    const grid = document.getElementById(gridId);
    const slot = grid.children[index];
    slot.innerHTML = `
      <svg class="camera-icon" viewBox="0 0 24 24" stroke-linecap="round" stroke-linejoin="round"><path d="M23 19a2 2 0 0 1-2 2H3a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h4l2-3h6l2 3h4a2 2 0 0 1 2 2z"/><circle cx="12" cy="13" r="4"/></svg>
      <span class="label">${labels[index]}</span>
    `;
  },

  // ---- AI Photo Analysis ----
  // Track whether we entered a damage screen from the linear flow vs from edit-review,
  // so the Back/Done buttons can route correctly.
  _damageFirstFromReview: false,
  _damageThirdFromReview: false,

  async analyzeYoursAndContinue() {
    const hasPhotos = this.reportData.photosYours.some(p => p);

    if (hasPhotos && navigator.onLine) {
      document.getElementById('aiAnalysisStatusYours').style.display = 'block';

      try {
        const results = await this.callAIAnalysis(this.reportData.photosYours);
        this.reportData.aiResults = results;
        this.applyAIResults(results);
      } catch (err) {
        console.warn('[AI Analysis] Failed, continuing with manual input:', err);
      }

      document.getElementById('aiAnalysisStatusYours').style.display = 'none';
    }

    // Damage + severity are AI-filled in the background and only surfaced if the
    // driver clicks Edit on the Review screen. Skip the damage screen in the main flow.
    if (this.reportData.answers.thirdParty) {
      this.goTo('photos-third');
    } else {
      // No third party → also skip photos-third + documents (which are 3P-specific)
      this.goTo('narrative');
    }
  },

  damageFirstBack() {
    // This screen is only reachable from Review's Edit link — always go back to review
    this.goTo('review');
  },

  damageFirstDone() {
    this.goTo('review');
  },

  async analyzeThirdAndContinue() {
    const hasPhotos = this.reportData.photosThird.some(p => p);

    if (hasPhotos && navigator.onLine) {
      document.getElementById('aiAnalysisStatusThird').style.display = 'block';

      try {
        const results = await this.callAIAnalysis(this.reportData.photosThird);
        if (this.reportData.aiResults) {
          if (results.damageZones && results.damageZones.third) {
            this.reportData.aiResults.damageZones.third = results.damageZones.third;
          }
          if (results.severityThird) {
            this.reportData.aiResults.severityThird = results.severityThird;
          }
        } else {
          this.reportData.aiResults = results;
        }
        this.applyAIResults(this.reportData.aiResults);
      } catch (err) {
        console.warn('[AI Analysis] Failed, continuing with manual input:', err);
      }

      document.getElementById('aiAnalysisStatusThird').style.display = 'none';
    }

    // Same as above — damage/severity edited only via Review. Go straight to docs.
    this.goTo('documents');
  },

  damageThirdBack() {
    this.goTo('review');
  },

  damageThirdDone() {
    this.goTo('review');
  },

  async callAIAnalysis(photos) {
    // In production, this calls your backend API
    // For dev/mock, we simulate a response
    if (this.api && this.api._isMock) {
      return this.getMockAIResults();
    }

    // Real API call would go here:
    // const response = await fetch(AI_BACKEND_URL, { ... });
    // return response.json();
    return this.getMockAIResults();
  },

  getMockAIResults() {
    return new Promise(resolve => {
      setTimeout(() => {
        resolve({
          thirdPartyDetected: true,
          thirdPartyVehicleType: 'Pick Up Truck',
          thirdPartyVIN: '1FTFW1ET5DFC10042',
          thirdPartyPlate: 'ABC 1234',
          damageZones: {
            first: ['Front Left', 'Front Center'],
            third: ['Rear Center', 'Rear Right']
          },
          severityFirst: 'Functional',
          severityThird: 'Minor',
          confidenceScores: {
            vehicleType: 0.92,
            vin: 0.67,
            plate: 0.88,
            severity: 0.81
          }
        });
      }, 2000);
    });
  },

  applyAIResults(results) {
    if (!results) return;

    // Auto-select damage zones
    if (results.damageZones) {
      (results.damageZones.first || []).forEach(zone => {
        const chip = document.querySelector(`#damageChipsFirst .zone-chip[onclick*="${zone}"]`);
        if (chip && !chip.classList.contains('selected')) {
          chip.classList.add('selected');
          this.reportData.damageZones.first.push(zone);
        }
      });
      (results.damageZones.third || []).forEach(zone => {
        const chip = document.querySelector(`#damageChipsThird .zone-chip[onclick*="${zone}"]`);
        if (chip && !chip.classList.contains('selected')) {
          chip.classList.add('selected');
          this.reportData.damageZones.third.push(zone);
        }
      });
    }

    // Auto-select severity (first party) — now lives inside the combined damage-first screen
    if (results.severityFirst) {
      this.reportData.severityFirst = results.severityFirst;
      document.querySelectorAll('#severityOptionsFirst .severity-option').forEach(opt => {
        opt.classList.remove('selected');
        if (opt.querySelector('h4')?.textContent === results.severityFirst) {
          opt.classList.add('selected');
        }
      });
    }

    // Auto-select severity (third party) — now lives inside the combined damage-third screen
    if (results.severityThird) {
      this.reportData.severityThird = results.severityThird;
      document.querySelectorAll('#severityOptionsThird .severity-option').forEach(opt => {
        opt.classList.remove('selected');
        if (opt.querySelector('h4')?.textContent === results.severityThird) {
          opt.classList.add('selected');
        }
      });
    }

    // Auto-select third party vehicle type
    if (results.thirdPartyVehicleType) {
      this.reportData.answers.thirdPartyType = results.thirdPartyVehicleType;
    }

    // Populate OCR-like fields if available
    if (results.thirdPartyVIN) {
      this.reportData.ocr.vin = results.thirdPartyVIN;
    }
    if (results.thirdPartyPlate) {
      this.reportData.ocr.plate = results.thirdPartyPlate;
    }
  },

  // ---- Edit-from-review flow ----
  editFromReview(screenId) {
    // Severity is now part of the combined damage screen — route accordingly
    if (screenId === 'damage-first' || screenId === 'severity-first') {
      this._damageFirstFromReview = true;
      this.goTo('damage-first');
      return;
    }
    if (screenId === 'damage-third' || screenId === 'severity-third') {
      this._damageThirdFromReview = true;
      this.goTo('damage-third');
      return;
    }
    this.goTo(screenId);
  },

  returnToReview() {
    this.goTo('review');
  },

  // ---- Damage Zones ----
  initDamageZoneClicks() {
    document.querySelectorAll('.damage-zone').forEach(zoneEl => {
      zoneEl.addEventListener('click', () => {
        const zoneName = zoneEl.getAttribute('data-zone');
        // data-party tells us which screen this SVG belongs to (defaults to 'first')
        const party = zoneEl.getAttribute('data-party') || 'first';
        const chipContainerId = party === 'third' ? 'damageChipsThird' : 'damageChipsFirst';
        const chip = document.querySelector(`#${chipContainerId} .zone-chip[onclick*="${zoneName}"]`);
        if (chip) {
          this.toggleZone(party, zoneName, chip);
          // Sync SVG highlight on the clicked zone only
          zoneEl.classList.toggle('selected', this.reportData.damageZones[party].includes(zoneName));
        }
      });
    });
  },

  toggleZone(party, zone, chipEl) {
    const zones = this.reportData.damageZones[party];
    const idx = zones.indexOf(zone);

    if (idx > -1) {
      zones.splice(idx, 1);
      chipEl.classList.remove('selected');
    } else {
      zones.push(zone);
      chipEl.classList.add('selected');
    }

    // Sync SVG zone highlight
    const svgZone = document.querySelector(`.damage-zone[data-zone="${zone}"]`);
    if (svgZone) {
      svgZone.classList.toggle('selected', zones.includes(zone));
    }

    // Track if user overrides AI suggestion
    if (this.reportData.aiResults) {
      const aiZones = this.reportData.aiResults.damageZones[party] || [];
      if (aiZones.includes(zone)) {
        this.reportData.corrections.push({
          field: `damageZones.${party}`,
          aiValue: zone,
          userAction: idx > -1 ? 'removed' : 'kept',
          timestamp: new Date().toISOString()
        });
      }
    }
  },

  // ---- Severity ----
  setSeverity(party, level, optionEl) {
    const field = party === 'first' ? 'severityFirst' : 'severityThird';
    const previous = this.reportData[field];
    this.reportData[field] = level;

    // Only deselect siblings within the same screen
    const screen = optionEl.closest('.screen');
    screen.querySelectorAll('.severity-option').forEach(o => o.classList.remove('selected'));
    optionEl.classList.add('selected');

    // Track correction
    const aiField = party === 'first' ? 'severityFirst' : 'severityThird';
    if (this.reportData.aiResults && previous !== level) {
      this.reportData.corrections.push({
        field: field,
        aiValue: this.reportData.aiResults[aiField],
        userValue: level,
        timestamp: new Date().toISOString()
      });
    }
  },

  // ---- Documents navigation ----
  documentsBack() {
    // Damage screens are hidden from main flow — go back to the photos screen
    if (this.reportData.answers.thirdParty) {
      this.goTo('photos-third');
    } else {
      this.goTo('photos-yours');
    }
  },

  // Old captureDocument/removeDocument/processDocumentOCR removed — replaced by captureDocOCR/removeDocOCR/_runDocOCR above.

  // Old renderOCRResults/updateConfBadge removed — replaced by _renderDocFields above.

  // ---- Telemetry & Context ----
  async _fetchEventContext(event) {
    // Test incident — restore pre-seeded dummy context instead of hitting the API
    if (event?.id === 'aTestIncident001') {
      this.reportData.context = {
        eventTime:   new Date(event.activeFrom),
        locationStr: '1145 Eglinton Ave E, Toronto, ON',
        latitude:     43.7085,
        longitude:   -79.3398,
        speedKmh:     52,
        gForce:       2.8
      };
      return;
    }

    const eventTime = event?.activeFrom ? new Date(event.activeFrom) : new Date();
    const deviceId = event?.device?.id || this.state?.device?.id;

    // Always store the event timestamp
    this.reportData.context.eventTime = eventTime;

    if (!this.api || !deviceId) return;

    try {
      // Query LogRecord ±2 min around event time to get GPS + speed
      const from = new Date(eventTime.getTime() - 120000);
      const to   = new Date(eventTime.getTime() + 30000);

      const logs = await new Promise((resolve, reject) =>
        this.api.call('Get', {
          typeName: 'LogRecord',
          search: { deviceSearch: { id: deviceId }, fromDate: from.toISOString(), toDate: to.toISOString() }
        }, resolve, reject)
      );

      if (!logs || logs.length === 0) return;

      // Record closest to the event time
      const target = eventTime.getTime();
      const sorted = logs.slice().sort((a, b) => new Date(a.dateTime) - new Date(b.dateTime));
      const closest = sorted.reduce((prev, cur) =>
        Math.abs(new Date(cur.dateTime) - target) < Math.abs(new Date(prev.dateTime) - target) ? cur : prev
      );

      this.reportData.context.latitude  = closest.latitude;
      this.reportData.context.longitude = closest.longitude;
      this.reportData.context.speedKmh  = closest.speed; // km/h

      // G-Force: ExceptionEvent.activeReason sometimes contains the peak g-force value.
      // Primary: read it from the event directly if available (set in the caller).
      // Fallback: estimate from speed delta between the LogRecord just before and just after the event.
      // This estimate is rough — it depends on GPS polling frequency (~every 1-5s) and accuracy.
      // Real accelerometer data lives in DebugRecord (high-freq ~0.1s) but requires separate queries.
      if (!this.reportData.context.gForce) {
        const before = sorted.filter(l => new Date(l.dateTime) <= eventTime);
        const after  = sorted.filter(l => new Date(l.dateTime) >  eventTime);
        if (before.length && after.length) {
          const b = before[before.length - 1], a = after[0];
          const dtSec = (new Date(a.dateTime) - new Date(b.dateTime)) / 1000;
          if (dtSec > 0) {
            const dvMs = (b.speed - a.speed) / 3.6; // Δv in m/s
            const g = Math.abs(dvMs / (dtSec * 9.81));
            if (g > 0.05) this.reportData.context.gForce = parseFloat(g.toFixed(2));
          }
        }
      }

      // Reverse-geocode coordinates
      try {
        const addrs = await new Promise((resolve, reject) =>
          this.api.call('GetAddresses', {
            coordinates: [{ x: closest.longitude, y: closest.latitude }]
          }, resolve, reject)
        );
        if (addrs && addrs.length) {
          const a = addrs[0];
          const parts = [
            a.streetNumber ? `${a.street || ''} ${a.streetNumber}`.trim() : (a.street || null),
            a.city,
            a.province || a.state
          ].filter(Boolean);
          this.reportData.context.locationStr = parts.length
            ? parts.join(', ')
            : `${closest.latitude.toFixed(4)}, ${closest.longitude.toFixed(4)}`;
        }
      } catch (e) {
        this.reportData.context.locationStr =
          `${closest.latitude.toFixed(4)}°, ${closest.longitude.toFixed(4)}°`;
      }

      // Weather and road conditions are intentionally not fetched here.
      // External weather APIs (e.g. Open-Meteo) are blocked by Geotab's CSP.
      // These fields will be populated by AI analysis of scene photos/video.

    } catch (err) {
      console.warn('[Context] Failed to fetch event context:', err);
    }
  },

populateContextScreen() {
    const ctx = this.reportData.context;

    // Time
    if (ctx.eventTime) {
      const t = new Date(ctx.eventTime);
      this.setEl('ctxTime', t.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' }));
    }

    // Location
    this.setEl('ctxLocation',
      ctx.locationStr ||
      (ctx.latitude != null ? `${ctx.latitude.toFixed(4)}°, ${ctx.longitude.toFixed(4)}°` : '—')
    );

    // Speed at impact
    this.setEl('ctxSpeed',
      ctx.speedKmh != null ? `${(ctx.speedKmh * 0.621371).toFixed(0)} mph` : '—'
    );

    // G-force
    this.setEl('ctxGForce',
      ctx.gForce != null ? `-${ctx.gForce.toFixed(1)}g` : '—'
    );

    // Weather & road
    this.setEl('ctxWeather',        ctx.weather        || '—');
    this.setEl('ctxRoad',           ctx.roadConditions || '—');
  },

  // ---- Review Screen ----
  populateReview() {
    const d = this.reportData;
    const a = d.answers;
    const isThirdParty = !!a.thirdParty;

    // Incident date/time and location from event context
    const ctx = d.context;
    if (ctx.eventTime) {
      const t = new Date(ctx.eventTime);
      this.setEl('revDateTime',
        t.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' }) +
        ' ' + t.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' })
      );
    }
    if (ctx.locationStr) {
      this.setEl('revLocation', ctx.locationStr);
    }

    this.setEl('revThirdParty', isThirdParty ? 'Yes' : 'No');
    // Hit-and-run only shown when thirdParty=Yes
    const harRow = document.getElementById('revRowHitAndRun');
    if (harRow) harRow.style.display = isThirdParty ? '' : 'none';
    this.setEl('revHitAndRun', a.hitAndRun === null || a.hitAndRun === undefined ? '—' : (a.hitAndRun ? 'Yes' : 'No'));
    this.setEl('revDriveable', a.driveable || '—');
    this.setEl('revFirstDamage', d.damageZones.first.join(', ') || '—');
    this.setEl('revSeverityFirst', d.severityFirst || '—');
    this.setEl('revThirdType', a.thirdPartyType || '—');
    this.setEl('revThirdDamage', d.damageZones.third.join(', ') || '—');
    this.setEl('revSeverityThird', d.severityThird || '—');
    this.setEl('revName', d.ocr.name || '—');
    this.setEl('revPolicy', d.ocr.policy || '—');
    this.setEl('revVin', d.ocr.vin || '—');
    this.setEl('revPlate', d.ocr.plate || '—');

    // Occupancy
    this.setEl('revYourVehiclePeople', d.occupancy.yourVehicle);
    this.setEl('revYourInjuries', d.occupancy.yourInjuries === null ? '—' : d.occupancy.yourInjuries ? 'Yes' : 'No');
    this.setEl('revThirdVehiclePeople', d.occupancy.thirdVehicle);
    this.setEl('revThirdInjuries', d.occupancy.thirdInjuries === null ? '—' : d.occupancy.thirdInjuries ? 'Yes' : 'No');

    // Witnesses
    const hasW = d.witnesses.hasWitnesses;
    this.setEl('revWitnesses', hasW === null ? '—' : hasW ? 'Yes' : 'No');
    this.setEl('revWitnessName', hasW && d.witnesses.name ? d.witnesses.name : '—');
    const wp = d.witnesses.phone;
    this.setEl('revWitnessPhone', hasW && wp.number ? `${wp.countryCode} ${wp.number}` : '—');
    ['revRowWitnessName', 'revRowWitnessPhone'].forEach(id => {
      const el = document.getElementById(id);
      if (el) el.style.display = hasW ? '' : 'none';
    });

    // Police report
    const pr = d.policeReport;
    this.setEl('revPoliceFiled', pr.filed === null ? '—' : pr.filed ? 'Yes' : 'No');
    this.setEl('revPoliceDoc', pr.document ? 'Attached ✓' : 'Not provided');
    this.setEl('revPoliceViolations', pr.violations === null ? '—' : pr.violations ? 'Yes' : 'No');
    this.setEl('revPoliceCitations', pr.citations === null ? '—' : pr.citations ? 'Yes' : 'No');
    ['revRowPoliceDoc', 'revRowPoliceViolations', 'revRowPoliceCitations'].forEach(id => {
      const el = document.getElementById(id);
      if (el) el.style.display = pr.filed ? '' : 'none';
    });

    // Property damage
    const pd = d.propertyDamageInfo;
    this.setEl('revPropertyDamaged', pd.damaged === null ? '—' : pd.damaged ? 'Yes' : 'No');
    this.setEl('revPropertyName', pd.propertyName || '—');
    this.setEl('revPropertyAddress', pd.address || '—');
    ['revRowPropertyName', 'revRowPropertyAddress'].forEach(id => {
      const el = document.getElementById(id);
      if (el) el.style.display = pd.damaged ? '' : 'none';
    });

    // Narrative preview
    const narrative = (d.narrative || '').trim();
    this.setEl('revNarrative', narrative || '— (none provided)');

    // Documents captured summary
    const docs = [];
    if (d.docLicense) docs.push("Driver's License");
    if (d.docInsurance) docs.push('Insurance Card');
    if (d.docRegistration) docs.push('Vehicle Registration');
    if (d.policeReport.document) docs.push('Police Report');
    if (d.policeReport.citationDoc) docs.push('Citation');
    if (d.propertyDamageInfo.photo) docs.push('Property Damage Photo');
    if (d.sceneVideo) docs.push('Scene Video');
    const docsSection = document.getElementById('revSectionDocs');
    if (docsSection) docsSection.style.display = docs.length ? '' : 'none';
    const docsEl = document.getElementById('revDocsCaptured');
    if (docsEl) {
      docsEl.innerHTML = docs.length
        ? docs.map(d => `<div style="padding:1px 0">✓ ${d}</div>`).join('')
        : '—';
    }

    // Show/hide third-party rows
    const thirdRows = ['revRowThirdDamage', 'revRowThirdSeverity', 'revRowThirdType'];
    thirdRows.forEach(id => {
      const el = document.getElementById(id);
      if (el) el.style.display = isThirdParty ? '' : 'none';
    });
    const thirdInfoSection = document.getElementById('revSectionThirdInfo');
    if (thirdInfoSection) thirdInfoSection.style.display = isThirdParty ? '' : 'none';
    ['revRowThirdVehiclePeople', 'revRowThirdInjuries'].forEach(id => {
      const el = document.getElementById(id);
      if (el) el.style.display = isThirdParty ? '' : 'none';
    });
  },

  // ---- Submit ----
  async submitReport() {
    this.reportData.narrative = (document.getElementById('narrativeText')?.value) || this.reportData.narrative;

    const submitBtn = document.querySelector('#screen-review .btn-success');
    const statusEl = document.getElementById('submitStatus');
    if (submitBtn) { submitBtn.disabled = true; submitBtn.textContent = 'Submitting…'; }
    if (statusEl) statusEl.style.display = '';

    try {
      let receipt = null;
      if (this.api && !this.api._isMock) {
        receipt = await this.submitToGeotab();
      } else {
        // Demo mode — simulate delay
        this.setEl('submitStatus', 'Saving report (demo mode)…');
        await new Promise(r => setTimeout(r, 1500));
        console.log('[Submit] Mock submission data:', JSON.stringify(this.reportData, null, 2));
      }

      this.goTo('success');
      this._showSubmitReceipt(receipt);
      this.clearProgress();
    } catch (err) {
      console.error('[Submit] Failed:', err);
      if (submitBtn) { submitBtn.disabled = false; submitBtn.textContent = 'Submit Report'; }
      if (statusEl) {
        const msg = err?.message || err?.name || (typeof err === 'string' ? err : JSON.stringify(err));
        statusEl.textContent = `Submission failed: ${msg || 'Unknown error'}`;
        statusEl.style.color = 'var(--error, #c62828)';
      }
      await this.saveOffline(this.reportData);
    }
  },

  _showSubmitReceipt(receipt) {
    const el = document.getElementById('submitReceipt');
    if (!el) return;

    if (!receipt) {
      el.innerHTML = '<span style="color:var(--text-muted)">Demo mode — no data sent to database.</span>';
      el.style.display = '';
      return;
    }

    const {
      addInDataId, addInDataError,
      mediaFileIds = [], uploadErrors = [], binaryErrors = [], attachmentErrors = [],
      commentOk, commentError,
      sceneVideoStatus, sceneVideoError,
      exceptionEventId, mediaFileSupported, mediaFileSample, deviceId
    } = receipt;
    const lines = [];

    lines.push(`<strong>Device ID:</strong> ${deviceId || '?'}`);
    lines.push(`<strong>Exception Event:</strong> ${exceptionEventId || '<span style="color:var(--text-muted)">None</span>'}`);
    const mfStatus = mediaFileSupported === true ? `✓ supported${mediaFileSample ? ' (sample solutionId: ' + (mediaFileSample.solutionId || 'none') + ')' : ' (no existing files)'}` : mediaFileSupported === false ? '<span style="color:var(--error)">✗ NOT supported in this database</span>' : '<span style="color:var(--text-muted)">not checked</span>';
    lines.push(`<strong>MediaFile entity:</strong> ${mfStatus}`);

    if (addInDataId) {
      lines.push(`<strong>AddInData:</strong> saved ✓`);
    } else {
      const errDetail = addInDataError ? ` — ${addInDataError}` : '';
      lines.push(`<strong>AddInData:</strong> <span style="color:var(--text-muted)">Skipped${errDetail}</span>`);
    }

    // Per-step counts: MediaFile entity created vs binary uploaded vs attached to exception
    const totalEntities = mediaFileIds.length;
    const binaryOkCount = mediaFileIds.filter(f => f.binaryOk).length;
    const attachOkCount = totalEntities - attachmentErrors.length;
    if (totalEntities > 0) {
      lines.push(`<strong>MediaFile records created:</strong> ${totalEntities}`);
      lines.push(`<strong>Binary bytes uploaded:</strong> ${binaryOkCount} / ${totalEntities}${binaryOkCount === totalEntities ? ' ✓' : ' <span style="color:var(--error)">⚠</span>'}`);
      if (exceptionEventId) {
        lines.push(`<strong>Attached to exception event:</strong> ${attachOkCount} / ${totalEntities}${attachOkCount === totalEntities ? ' ✓' : ' <span style="color:var(--error)">⚠</span>'}`);
      }
      lines.push(`<strong>Files:</strong>`);
      mediaFileIds.forEach(f => lines.push(`&nbsp;&nbsp;• ${f.name}${f.binaryOk ? '' : ' <span style="color:var(--error)">(no binary)</span>'}`));
    } else {
      lines.push(`<strong>Files uploaded:</strong> None`);
    }

    if (sceneVideoStatus && sceneVideoStatus !== 'none') {
      const sv = sceneVideoStatus === 'success'
        ? 'uploaded ✓'
        : sceneVideoStatus === 'failed'
          ? `<span style="color:var(--error)">failed — ${sceneVideoError || 'unknown'}</span>`
          : 'attempted';
      lines.push(`<strong>Scene video:</strong> ${sv}`);
    }

    if (exceptionEventId) {
      lines.push(`<strong>Comment on exception:</strong> ${commentOk ? 'added ✓' : `<span style="color:var(--error)">failed — ${commentError || 'unknown'}</span>`}`);
    }

    if (uploadErrors.length > 0) {
      lines.push(`<strong style="color:var(--error)">MediaFile.Add errors (${uploadErrors.length}):</strong>`);
      uploadErrors.slice(0, 5).forEach(e => lines.push(`<span style="font-size:13px;color:var(--error);word-break:break-all;display:block;padding:4px 0">&bull; ${e}</span>`));
    }
    if (binaryErrors.length > 0) {
      lines.push(`<strong style="color:var(--error)">Binary upload errors (${binaryErrors.length}):</strong>`);
      binaryErrors.slice(0, 5).forEach(e => lines.push(`<span style="font-size:13px;color:var(--error);word-break:break-all;display:block;padding:4px 0">&bull; ${e}</span>`));
    }
    if (attachmentErrors.length > 0) {
      lines.push(`<strong style="color:var(--error)">ExceptionEventAttachment errors (${attachmentErrors.length}):</strong>`);
      attachmentErrors.slice(0, 5).forEach(e => lines.push(`<span style="font-size:13px;color:var(--error);word-break:break-all;display:block;padding:4px 0">&bull; ${e}</span>`));
    }

    el.innerHTML = lines.join('<br>');
    el.style.display = '';
  },

  async submitToGeotab() {
    const deviceId = this.state.device.id;
    const driverId = this.state.driver?.id || null;
    const dateTime = new Date().toISOString();
    const server = document.location.hostname || this.state?.server || 'my.geotab.com';
    const credentials = await this._getApiCredentials();

    // 1. Find the relevant exception event (full object for activeFrom timestamp)
    this.setEl('submitStatus', 'Locating incident event…');
    const exceptionEvent = await this.getExceptionEvent();
    const exceptionEventId = exceptionEvent?.id || null;
    const exceptionDateTime = exceptionEvent?.activeFrom || dateTime;
    console.log('[Submit] Exception event:', exceptionEventId, 'at', exceptionDateTime);

    // 1b. Preflight: verify MediaFile entity type is accessible in this database
    this.setEl('submitStatus', 'Checking media support…');
    let mediaFileSupported = false;
    let mediaFileSample = null;
    try {
      const mfResult = await new Promise((resolve, reject) =>
        this.api.call('Get', {
          typeName: 'MediaFile',
          search: { deviceSearch: { id: deviceId } },
          resultsLimit: 1
        }, resolve, reject)
      );
      mediaFileSupported = true;
      mediaFileSample = mfResult?.[0] || null;
      console.log('[Submit] MediaFile Get OK, count:', mfResult?.length, 'sample:', JSON.stringify(mediaFileSample));
    } catch (e) {
      console.error('[Submit] MediaFile Get FAILED:', JSON.stringify(e));
      mediaFileSupported = false;
    }

    // 2. Collect every photo / document to upload — names are descriptive PascalCase for organization
    const yourLabels  = ['UserVehicle_FrontView', 'UserVehicle_RearView', 'UserVehicle_LeftSide', 'UserVehicle_RightSide', 'UserVehicle_DamageCloseup'];
    const thirdLabels = ['ThirdParty_FrontView',  'ThirdParty_RearView',  'ThirdParty_LeftSide',  'ThirdParty_RightSide',  'ThirdParty_DamageCloseup'];
    const uploads = [
      ...this.reportData.photosYours.map((d, i) => d ? { data: d, name: yourLabels[i]  } : null),
      ...this.reportData.photosThird.map((d, i) => d ? { data: d, name: thirdLabels[i] } : null),
      this.reportData.docLicense               && { data: this.reportData.docLicense,                  name: 'ThirdParty_DriversLicense' },
      this.reportData.docInsurance             && { data: this.reportData.docInsurance,                name: 'ThirdParty_InsuranceCard' },
      this.reportData.docRegistration           && { data: this.reportData.docRegistration,             name: 'ThirdParty_Registration' },
      this.reportData.policeReport.document    && { data: this.reportData.policeReport.document,       name: 'PoliceReport_Document' },
      this.reportData.policeReport.citationDoc && { data: this.reportData.policeReport.citationDoc,    name: 'PoliceReport_Citation' },
      this.reportData.propertyDamageInfo.photo && { data: this.reportData.propertyDamageInfo.photo,    name: 'PropertyDamage_Photo' },
    ].filter(Boolean);

    // 3. Upload each file as MediaFile + binary + ExceptionEventAttachment.
    // Track each step independently so the receipt can show exactly what succeeded/failed.
    const mediaFileIds = [];          // entity records that exist server-side
    const uploadErrors = [];          // MediaFile.Add throws
    const binaryErrors = [];          // binary POST failed (entity exists but no bytes)
    const attachmentErrors = [];      // ExceptionEventAttachment failed (binary exists but not linked to event)
    for (let i = 0; i < uploads.length; i++) {
      const item = uploads[i];
      this.setEl('submitStatus', `Uploading ${i + 1} of ${uploads.length}: ${item.name}…`);
      try {
        const result = await this.uploadMediaFile(
          item.data, item.name, deviceId, driverId,
          dateTime, exceptionEventId, server, credentials
        );
        if (result?.id) {
          mediaFileIds.push({ id: result.id, name: item.name, binaryOk: result.binaryOk });
          if (!result.binaryOk) binaryErrors.push(`${item.name}: ${result.binaryError || 'unknown'}`);
          if (exceptionEventId) {
            const att = await this._attachMediaToException(result.id, exceptionEventId);
            if (!att.ok) attachmentErrors.push(`${item.name}: ${att.error}`);
          }
        }
      } catch (e) {
        const msg = (e?.message || e?.name || String(e) || 'unknown') + (e ? ' | raw: ' + JSON.stringify(e) : '');
        console.error('[Submit] MediaFile.Add failed for', item.name, JSON.stringify(e), e);
        uploadErrors.push(`${item.name}: ${msg}`);
      }
    }

    // 3b. Upload scene video blob if captured.
    // Track all three possible states distinctly so the receipt shows truth:
    //   - sceneVideoStatus: 'none' (user never attached), 'attempted', 'success', or 'failed'
    let sceneVideoStatus = 'none';
    let sceneVideoError = null;
    if (this.reportData.sceneVideo) {
      sceneVideoStatus = 'attempted';
      if (!this._sceneVideoBlob) {
        sceneVideoError = 'Video blob missing at submit time (likely iOS File reference invalidation).';
        sceneVideoStatus = 'failed';
        binaryErrors.push(`SceneVideo: ${sceneVideoError}`);
        console.error('[Submit] Scene video metadata exists but blob is null — possible iOS File invalidation');
      } else {
        this.setEl('submitStatus', 'Uploading scene video…');
        try {
          const result = await this.uploadVideoFile(
            this._sceneVideoBlob, 'SceneVideo',
            deviceId, driverId, dateTime, exceptionEventId, server, credentials
          );
          if (result?.id) {
            mediaFileIds.push({ id: result.id, name: 'SceneVideo', binaryOk: result.binaryOk });
            if (!result.binaryOk) {
              binaryErrors.push(`SceneVideo: ${result.binaryError || 'unknown'}`);
              sceneVideoStatus = 'failed';
              sceneVideoError = result.binaryError;
            } else {
              sceneVideoStatus = 'success';
            }
            if (exceptionEventId) {
              const att = await this._attachMediaToException(result.id, exceptionEventId);
              if (!att.ok) attachmentErrors.push(`SceneVideo: ${att.error}`);
            }
          } else {
            sceneVideoStatus = 'failed';
            sceneVideoError = 'MediaFile.Add returned no ID';
            binaryErrors.push(`SceneVideo: ${sceneVideoError}`);
          }
        } catch (e) {
          sceneVideoStatus = 'failed';
          sceneVideoError = (e?.message || String(e)).slice(0, 200);
          binaryErrors.push(`SceneVideo: ${sceneVideoError}`);
          console.warn('[Submit] Scene video upload failed:', e);
        }
      }
    }

    // 4. Save full structured report as AddInData
    this.setEl('submitStatus', 'Saving report…');
    const d = this.reportData;
    const reportText = this.formatReportText();
    // Strip raw image data from OCR — only keep extracted text fields (image already uploaded as MediaFile)
    const { documentImage: _omit, ...ocrTextFields } = d.ocr || {};
    // Look up the actual registered add-in ID — AddInData requires a valid registered AddIn id
    let resolvedAddInId = this._cachedAddInId || null;
    if (!resolvedAddInId) {
      try {
        const addIns = await new Promise((resolve) =>
          this.api.call('Get', { typeName: 'AddIn' }, resolve, () => resolve([]))
        );
        const mine = (addIns || []).find(a =>
          // Match on v3 URL first, v2 URL for backward compat, or add-in name
          ['incident-addin-v3', 'incident-addin-v2', 'Incident Reconstruction Engine'].some(
            term => JSON.stringify(a.configuration || a).includes(term)
          )
        );
        resolvedAddInId = mine?.id || null;
        if (resolvedAddInId) this._cachedAddInId = resolvedAddInId;
        console.log('[Submit] Resolved AddIn id:', resolvedAddInId, 'from', (addIns || []).length, 'add-ins');
      } catch (e) {
        console.warn('[Submit] Could not look up AddIn id:', e);
      }
    }

    let addInDataId = null;
    let addInDataError = null;
    if (!resolvedAddInId) {
      addInDataError = 'Add-in not found in database — AddInData skipped';
      console.warn('[Submit]', addInDataError);
    } else {
    try {
      addInDataId = await new Promise((resolve, reject) =>
        this.api.call('Add', {
          typeName: 'AddInData',
          entity: {
            addInId: resolvedAddInId,
            details: {
              exceptionEventId: exceptionEventId || null,
              submittedAt: dateTime,
              device: { id: deviceId, name: this.state.device.name },
              driver: driverId ? { id: driverId, name: this.state.driver.name } : null,
              mediaFileIds,
              reportText,
              incident: {
                answers: d.answers,
                narrative: d.narrative,
                occupancy: d.occupancy,
                witnesses: { hasWitnesses: d.witnesses.hasWitnesses, name: d.witnesses.name, phone: d.witnesses.phone },
                policeReport: { filed: d.policeReport.filed, violations: d.policeReport.violations, citations: d.policeReport.citations },
                propertyDamage: {
                  damaged: d.propertyDamageInfo.damaged,
                  propertyName: d.propertyDamageInfo.propertyName,
                  address: d.propertyDamageInfo.address,
                  ownerName: d.propertyDamageInfo.ownerName,
                },
                thirdPartyOcr: ocrTextFields,
                thirdPartyPhone: d.thirdPartyPhone,
                damageZones: d.damageZones,
                severityFirst: d.severityFirst,
                severityThird: d.severityThird,
              }
            }
          }
        }, resolve, reject)
      );
    } catch (e) {
      addInDataError = e?.message || e?.name || JSON.stringify(e);
      console.error('[Submit] AddInData Add failed:', e);
    }
    } // end if (resolvedAddInId)

    // 5. Add report text as an ExceptionEventComment so it's visible on the exception page.
    // (Old code used Set ExceptionEvent { comment } — that field doesn't exist on the entity;
    // the correct approach is Add ExceptionEventComment, matching the official collision-form.)
    let commentOk = false;
    let commentError = null;
    if (exceptionEventId) {
      this.setEl('submitStatus', 'Adding comment to exception…');
      try {
        await new Promise((resolve, reject) =>
          this.api.call('Add', {
            typeName: 'ExceptionEventComment',
            entity: {
              exceptionEvent: { id: exceptionEventId },
              text: reportText
            }
          }, resolve, reject)
        );
        commentOk = true;
      } catch (e) {
        commentError = (e?.message || e?.name || JSON.stringify(e) || 'unknown').slice(0, 300);
        console.warn('[Submit] ExceptionEventComment add failed:', e);
      }
    }

    return {
      addInDataId, addInDataError,
      mediaFileIds, uploadErrors, binaryErrors, attachmentErrors,
      commentOk, commentError,
      sceneVideoStatus, sceneVideoError,
      exceptionEventId, mediaFileSupported, mediaFileSample, deviceId
    };
  },

  // ---- Submission helpers ----

  async getExceptionEvent() {
    // 1. User selected a specific event from the incidents list — pull full object from cache
    if (this._selectedExceptionEventId) {
      const cached = this._eventsCache?.[this._selectedExceptionEventId];
      if (cached) return cached;
      // Cache miss — fetch full event by id
      try {
        const events = await new Promise((resolve, reject) =>
          this.api.call('Get', {
            typeName: 'ExceptionEvent',
            search: { id: this._selectedExceptionEventId }
          }, resolve, reject)
        );
        if (events?.[0]) return events[0];
      } catch (e) { /* fall through */ }
      return { id: this._selectedExceptionEventId, activeFrom: null };
    }

    // 2. Drive SDK injects exception event via state
    if (this.state?.exceptionEvent?.id) return this.state.exceptionEvent;

    // 3. Fallback: most recent collision exception for this device in the last 2 hours
    try {
      const now = new Date();
      const twoHoursAgo = new Date(now - 2 * 60 * 60 * 1000);
      const events = await new Promise((resolve, reject) =>
        this.api.call('Get', {
          typeName: 'ExceptionEvent',
          search: {
            deviceSearch: { id: this.state.device.id },
            fromDate: twoHoursAgo.toISOString(),
            toDate: now.toISOString(),
          }
        }, resolve, reject)
      );
      if (events?.length > 0) {
        events.sort((a, b) => new Date(b.activeFrom) - new Date(a.activeFrom));
        return events[0];
      }
    } catch (e) {
      console.warn('[Submit] Could not find exception event:', e);
    }
    return null;
  },

  async uploadMediaFile(base64DataUrl, name, deviceId, driverId, eventDateTime, exceptionEventId, server, credentials) {
    // Resize/compress before upload
    const resized = await this._resizeImage(base64DataUrl);

    const fileName = this._descriptiveFileName(name, 'jpg', exceptionEventId);

    // Step 1: Create the MediaFile entity (metadata only — minimal fields per official pattern)
    // SolutionId must be a valid Geotab-format ID (base64-encoded GUID), not an arbitrary string.
    // Server's .NET deserializer crashes with GenericException if it can't parse SolutionId as a GUID.
    // Borrowing the official collision-form add-in's SolutionId until we register our own.
    const entity = {
      name: fileName,
      SolutionId: 'aYnBQxCQMv0-lyIH3F8689Q',
    };
    console.log('[Submit] MediaFile entity:', JSON.stringify(entity));
    const entityId = await new Promise((resolve, reject) =>
      this.api.call('Add', { typeName: 'MediaFile', entity }, resolve, reject)
    );

    // Step 2: Upload the binary via multipart POST to /apiv1/
    // The JSON-RPC params go in a form field named 'JSON-RPC' (URL-encoded).
    // The file goes in a form field named exactly after the entity's name.
    if (!entityId || !credentials || !server) {
      console.warn('[Submit] Skipping binary upload — missing credentials or server');
      return entityId;
    }

    let binaryOk = false;
    let binaryError = null;
    try {
      const base64 = resized.split(',')[1];
      const mimeType = resized.match(/data:([^;]+);/)?.[1] || 'image/jpeg';
      const byteChars = atob(base64);
      const byteArray = new Uint8Array(byteChars.length);
      for (let i = 0; i < byteChars.length; i++) byteArray[i] = byteChars.charCodeAt(i);
      const blob = new Blob([byteArray], { type: mimeType });

      const host = server.includes('://') ? new URL(server).hostname : server;
      const params = { method: 'UploadMediaFile', params: { credentials, mediaFile: { id: entityId } } };
      const formData = new FormData();
      formData.append('JSON-RPC', encodeURIComponent(JSON.stringify(params)));
      formData.append(fileName, blob, fileName);

      const resp = await fetch(`https://${host}/apiv1`, { method: 'POST', body: formData });
      const text = await resp.text();
      let json = null;
      try { json = JSON.parse(text); } catch (e) { /* not JSON */ }
      if (!resp.ok) {
        binaryError = `HTTP ${resp.status}: ${text.slice(0, 300)}`;
      } else if (json?.error) {
        binaryError = JSON.stringify(json.error).slice(0, 500);
      } else {
        binaryOk = true;
      }
      if (binaryError) console.warn('[Submit] UploadMediaFile binary error:', binaryError);
    } catch (e) {
      binaryError = (e?.message || String(e)).slice(0, 300);
      console.warn('[Submit] Binary upload threw:', e);
    }

    return { id: entityId, binaryOk, binaryError };
  },

  async uploadVideoFile(blob, name, deviceId, driverId, eventDateTime, exceptionEventId, server, credentials) {
    // Match file extension + Content-Type to the blob's actual format. The server validates that
    // filename extension and the multipart Content-Type agree — mismatch produces
    // "ArgumentException: .mp4 requires content type video/mp4". Sources of mismatch:
    //   - iOS Camera app records .mov (video/quicktime)
    //   - Android picker may return video/webm or video/3gpp
    //   - iOS Drive sometimes returns application/octet-stream with no real type info
    const sourceType = (blob?.type || '').toLowerCase();
    let ext = 'mp4', mime = 'video/mp4';
    if (sourceType.includes('quicktime') || sourceType.includes('mov'))      { ext = 'mov';  mime = 'video/quicktime'; }
    else if (sourceType.includes('webm'))                                    { ext = 'webm'; mime = 'video/webm'; }
    else if (sourceType.includes('3gpp') || sourceType.includes('3gp'))      { ext = '3gp';  mime = 'video/3gpp'; }
    else if (sourceType.includes('mp4'))                                     { ext = 'mp4';  mime = 'video/mp4'; }
    // If unknown/octet-stream, assume mp4 — most common on iOS Drive native uploads.

    // Ensure the blob actually carries the matching Content-Type. FormData uses blob.type as the
    // multipart part's Content-Type header. If blob.type is empty or wrong, the upload is rejected.
    const typedBlob = blob.type === mime ? blob : new Blob([blob], { type: mime });

    const fileName = this._descriptiveFileName(name, ext, exceptionEventId);

    const entityId = await new Promise((resolve, reject) =>
      this.api.call('Add', {
        typeName: 'MediaFile',
        entity: { name: fileName, SolutionId: 'aYnBQxCQMv0-lyIH3F8689Q' }
      }, resolve, reject)
    );

    if (!entityId || !credentials || !server) {
      console.warn('[Submit] Skipping video binary upload — missing credentials or server');
      return { id: entityId, binaryOk: false, binaryError: 'missing credentials or server' };
    }

    let binaryOk = false;
    let binaryError = null;
    try {
      const host = server.includes('://') ? new URL(server).hostname : server;
      const params = { method: 'UploadMediaFile', params: { credentials, mediaFile: { id: entityId } } };
      const formData = new FormData();
      formData.append('JSON-RPC', encodeURIComponent(JSON.stringify(params)));
      formData.append(fileName, typedBlob, fileName);

      const resp = await fetch(`https://${host}/apiv1`, { method: 'POST', body: formData });
      const text = await resp.text();
      let json = null;
      try { json = JSON.parse(text); } catch (e) { /* not JSON */ }
      if (!resp.ok) binaryError = `HTTP ${resp.status}: ${text.slice(0, 300)}`;
      else if (json?.error) binaryError = JSON.stringify(json.error).slice(0, 500);
      else binaryOk = true;
      if (binaryError) console.warn('[Submit] UploadMediaFile video binary error:', binaryError);
    } catch (e) {
      binaryError = (e?.message || String(e)).slice(0, 300);
      console.warn('[Submit] Video binary upload threw:', e);
    }

    return { id: entityId, binaryOk, binaryError };
  },

  async _attachMediaToException(mediaFileId, exceptionEventId) {
    try {
      const attachmentId = await new Promise((resolve, reject) =>
        this.api.call('Add', {
          typeName: 'ExceptionEventAttachment',
          entity: {
            exceptionEvent: { id: exceptionEventId },
            mediaFileAttachment: { id: mediaFileId }
          }
        }, resolve, reject)
      );
      return { ok: true, attachmentId };
    } catch (e) {
      const msg = (e?.message || e?.name || JSON.stringify(e) || 'unknown').slice(0, 300);
      console.warn('[Submit] ExceptionEventAttachment failed for', mediaFileId, e);
      return { ok: false, error: msg };
    }
  },

  async _getApiCredentials() {
    // Returns the inner { userName, sessionId, database } credentials object.
    // api.getSession() in Drive returns the full session: { credentials: { ... }, server, ... }.
    // UploadMediaFile's params expect only the inner credentials, NOT the wrapping session.
    // (Wrapping the whole session would double-nest credentials and the server rejects it.)
    const unwrap = (s) => (s && s.credentials) ? s.credentials : s;

    // 1. Use the documented api.getSession() if available (standard Geotab JS API)
    if (typeof this.api?.getSession === 'function') {
      return new Promise((resolve) => {
        try {
          this.api.getSession((session) => resolve(unwrap(session) || null));
        } catch (e) {
          resolve(null);
        }
      });
    }
    // 2. Fall back to internal property inspection
    for (const src of [this.api, this.api?._api, this.api?._rpc]) {
      if (src?._credentials) return unwrap(src._credentials);
      if (src?.credentials) return unwrap(src.credentials);
    }
    // 3. Build from state if available (Drive SDK sometimes exposes these)
    if (this.state?.sessionId) {
      return {
        sessionId: this.state.sessionId,
        userName: this.state.userName,
        database: this.state.database
      };
    }
    return null;
  },

  // Build a human-readable MediaFile name like "UserVehicle_FrontView_3Fa9Kp2x.jpg".
  // The suffix is the short exception event id (same 8 chars shown on the incident card)
  // so a vehicle's files group together and trace back to the event; falls back to a
  // random token for manual/test reports that have no exception event.
  _descriptiveFileName(label, ext, exceptionEventId) {
    const safeLabel = ((label || '').replace(/[^A-Za-z0-9_]/g, '').slice(0, 60)) || 'IncidentMedia';
    const suffix = (exceptionEventId && this._shortEventId(exceptionEventId)) || this._randomToken(8);
    return `${safeLabel}_${suffix}.${ext}`;
  },

  _randomToken(len) {
    const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    let s = '';
    for (let i = 0; i < len; i++) s += chars[Math.floor(Math.random() * chars.length)];
    return s;
  },

  async _resizeImage(base64DataUrl, maxWidth = 1280, quality = 0.75) {
    return new Promise((resolve) => {
      // Only resize actual images (skip PDFs or unknown types)
      if (!base64DataUrl || !base64DataUrl.startsWith('data:image/')) {
        return resolve(base64DataUrl);
      }
      const img = new Image();
      img.onload = () => {
        const scale = Math.min(1, maxWidth / img.width);
        const canvas = document.createElement('canvas');
        canvas.width = Math.round(img.width * scale);
        canvas.height = Math.round(img.height * scale);
        canvas.getContext('2d').drawImage(img, 0, 0, canvas.width, canvas.height);
        resolve(canvas.toDataURL('image/jpeg', quality));
      };
      img.onerror = () => resolve(base64DataUrl);
      img.src = base64DataUrl;
    });
  },

  formatReportText() {
    const d = this.reportData;
    const ctx = d.context || {};
    const wp = d.witnesses.phone || {};
    const op = d.propertyDamageInfo.ownerPhone || {};
    const yn = (v) => v === null || v === undefined ? 'Not specified' : v ? 'Yes' : 'No';
    const phone = (p) => p?.number ? `${p.countryCode || '+1'} ${p.number}` : 'N/A';
    const severityDesc = { Minor: 'Minor (scratches, scuffs, or small dents — vehicle fully driveable)', Functional: 'Functional (broken glass, hanging bumpers, or light damage — may still be driveable)', Disabling: 'Disabling (structural/frame damage, wheel misalignment, or airbag deployment — not driveable)' };
    const fmtDate = (dt) => {
      if (!dt) return null;
      const t = dt instanceof Date ? dt : new Date(dt);
      if (isNaN(t)) return null;
      return t.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' }) + ' ' +
             t.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' });
    };
    const now = new Date();
    const ts = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}-${String(now.getDate()).padStart(2,'0')} ${String(now.getHours()).padStart(2,'0')}:${String(now.getMinutes()).padStart(2,'0')}`;

    const lines = [];
    lines.push('INCIDENT REPORT');
    lines.push(`Submitted ${ts} by ${this.state?.driver?.name || 'Unknown'}`);
    lines.push(`Vehicle: ${this.state?.device?.name || 'Unknown'}`);
    lines.push('');

    // Incident details (telematics data — auto-filled, not entered by driver)
    lines.push('— Incident Details (from telematics) —');
    lines.push(`Date / Time: ${fmtDate(ctx.eventTime) || '—'}`);
    lines.push(`Location: ${ctx.locationStr || '—'}`);
    if (ctx.latitude && ctx.longitude) lines.push(`GPS: ${ctx.latitude.toFixed(5)}, ${ctx.longitude.toFixed(5)}`);
    const speedMph = ctx.speedKmh != null ? Math.round(ctx.speedKmh * 0.621371) : null;
    lines.push(`Speed at event: ${speedMph != null ? speedMph + ' mph (' + Math.round(ctx.speedKmh) + ' km/h)' : '—'}`);
    lines.push(`G-force: ${ctx.gForce != null ? ctx.gForce.toFixed(2) + ' G' : '—'}`);
    lines.push(`Weather: ${ctx.weather || '— (pending AI analysis)'}`);
    lines.push(`Road conditions: ${ctx.road || '— (pending AI analysis)'}`);
    lines.push('');

    lines.push('— Qualifying Questions —');
    lines.push('Q: Does the collision involve a third party driver?');
    lines.push(`A: ${yn(d.answers.thirdParty)}`);
    if (d.answers.thirdParty) {
      lines.push('Q: Was this a hit-and-run?');
      lines.push(`A: ${yn(d.answers.hitAndRun)}`);
    }
    lines.push(`Vehicle driveable: ${d.answers.driveable || 'Not specified'}`);
    lines.push('');

    lines.push('— Damage: Your Vehicle —');
    lines.push('Q: Which areas are damaged?');
    lines.push(`A: ${d.damageZones.first.length ? d.damageZones.first.join(', ') : 'None selected'}`);
    lines.push('Q: How severe is the damage?');
    lines.push(`A: ${severityDesc[d.severityFirst] || d.severityFirst || 'Not specified'}`);
    lines.push('');

    if (d.answers.thirdParty) {
      lines.push('— Damage: 3rd Party Vehicle —');
      if (d.answers.thirdPartyType) {
        lines.push('Q: What type of vehicle?');
        lines.push(`A: ${d.answers.thirdPartyType}`);
      }
      lines.push('Q: Which areas are damaged?');
      lines.push(`A: ${d.damageZones.third.length ? d.damageZones.third.join(', ') : 'None selected'}`);
      lines.push('Q: How severe is the damage?');
      lines.push(`A: ${severityDesc[d.severityThird] || d.severityThird || 'Not specified'}`);
      lines.push('');

      const docs = [];
      if (d.docLicense) docs.push("Driver's License");
      if (d.docInsurance) docs.push('Insurance Card');
      if (d.docRegistration) docs.push('Vehicle Registration');
      lines.push('— 3rd Party Driver Info —');
      lines.push(`Name: ${d.ocr.name || '—'}`);
      lines.push(`Insurance policy #: ${d.ocr.policy || '—'}`);
      lines.push(`Vehicle VIN: ${d.ocr.vin || '—'}`);
      lines.push(`License plate: ${d.ocr.plate || '—'}`);
      lines.push(`Phone: ${d.thirdPartyPhone?.number ? phone(d.thirdPartyPhone) : '—'}`);
      lines.push(`Documents captured: ${docs.length ? docs.join(', ') : 'None'}`);
      lines.push('');
    }

    lines.push('— Incident Description —');
    lines.push(d.narrative || 'No description provided');
    lines.push('');

    lines.push('— Occupancy & Injuries —');
    const yourInj = d.occupancy.yourInjuries === null ? '' : d.occupancy.yourInjuries ? ' — injuries reported' : ' — no injuries';
    lines.push(`Your vehicle: ${d.occupancy.yourVehicle || 1} person(s)${yourInj}`);
    if (d.answers.thirdParty) {
      const thirdInj = d.occupancy.thirdInjuries === null ? '' : d.occupancy.thirdInjuries ? ' — injuries reported' : ' — no injuries';
      lines.push(`Third party vehicle: ${d.occupancy.thirdVehicle || 1} person(s)${thirdInj}`);
    }
    lines.push('');

    lines.push('— Witnesses —');
    if (d.witnesses.hasWitnesses) {
      lines.push(`Witnesses present: Yes`);
      if (d.witnesses.name) lines.push(`Name: ${d.witnesses.name}`);
      if (wp.number) lines.push(`Phone: ${phone(wp)}`);
    } else {
      lines.push(`Witnesses present: ${yn(d.witnesses.hasWitnesses)}`);
    }
    lines.push('');

    lines.push('— Police Report —');
    if (d.policeReport.filed) {
      lines.push(`Report filed: Yes${d.policeReport.document ? ' (document attached)' : ''}`);
      if (d.policeReport.reportNumber) lines.push(`Report #: ${d.policeReport.reportNumber}`);
      if (d.policeReport.officerName) lines.push(`Officer: ${d.policeReport.officerName}${d.policeReport.badgeNumber ? ' (Badge #' + d.policeReport.badgeNumber + ')' : ''}`);
      lines.push(`Violations cited: ${yn(d.policeReport.violations)}`);
      lines.push(`Citation issued: ${yn(d.policeReport.citations)}${d.policeReport.citationDoc ? ' (document attached)' : ''}`);
      if (d.policeReport.citationNumber) lines.push(`Citation #: ${d.policeReport.citationNumber}`);
      if (d.policeReport.citationViolations) lines.push(`Violation details: ${d.policeReport.citationViolations}`);
    } else {
      lines.push(`Report filed: ${yn(d.policeReport.filed)}`);
    }
    lines.push('');

    lines.push('— Property Damage —');
    if (d.answers.propertyDamage || d.propertyDamageInfo.damaged) {
      lines.push(`Property damaged: Yes`);
      lines.push(`Property: ${d.propertyDamageInfo.propertyName || '—'}`);
      lines.push(`Address: ${d.propertyDamageInfo.address || '—'}`);
      lines.push(`Owner: ${d.propertyDamageInfo.ownerName || '—'}${op.number ? ', ' + phone(op) : ''}`);
      if (d.propertyDamageInfo.photo) lines.push('Photo of property damage attached.');
    } else {
      lines.push(`Property damaged: ${yn(d.propertyDamageInfo.damaged)}`);
    }
    lines.push('');

    const photoCountYours = (d.photosYours || []).filter(Boolean).length;
    const photoCountThird = (d.photosThird || []).filter(Boolean).length;
    const docCount = [d.docLicense, d.docInsurance, d.docRegistration].filter(Boolean).length;
    const policeCount = [d.policeReport.document, d.policeReport.citationDoc].filter(Boolean).length;
    lines.push('— Attachments —');
    lines.push(`Your vehicle photos: ${photoCountYours}`);
    if (d.answers.thirdParty) lines.push(`Third party vehicle photos: ${photoCountThird}`);
    if (docCount) lines.push(`Third party documents: ${docCount}`);
    if (policeCount) lines.push(`Police report documents: ${policeCount}`);
    if (d.propertyDamageInfo.photo) lines.push('Property damage photo: 1');
    if (d.sceneVideo) lines.push('Scene video: 1');
    lines.push('(see attached files on this event)');

    return lines.join('\n');
  },

  // ---- Offline / Persistence ----
  setupOfflineDetection() {
    const banner = document.getElementById('offlineBanner');
    window.addEventListener('online', () => {
      banner.classList.remove('visible');
      this.processPendingReports();
    });
    window.addEventListener('offline', () => {
      banner.classList.add('visible');
    });
    if (!navigator.onLine) banner.classList.add('visible');
  },

  saveProgress() {
    try {
      localStorage.setItem('incident-progress', JSON.stringify({
        screen: this.currentScreen,
        data: this.reportData,
        timestamp: Date.now()
      }));
    } catch (e) {
      console.warn('[Save] localStorage failed:', e);
    }
  },

  loadSavedProgress() {
    try {
      const saved = localStorage.getItem('incident-progress');
      if (saved) {
        const parsed = JSON.parse(saved);
        // Only restore if less than 24 hours old
        if (Date.now() - parsed.timestamp < 86400000) {
          this.reportData = { ...this.reportData, ...parsed.data };
          console.log('[Restore] Loaded saved progress');
        }
      }
    } catch (e) {
      console.warn('[Restore] Failed:', e);
    }
  },

  clearProgress() {
    localStorage.removeItem('incident-progress');
  },

  async saveOffline(reportData) {
    try {
      const pending = JSON.parse(localStorage.getItem('incident-pending') || '[]');
      pending.push({
        id: crypto.randomUUID(),
        data: reportData,
        timestamp: Date.now()
      });
      localStorage.setItem('incident-pending', JSON.stringify(pending));
    } catch (e) {
      console.warn('[Offline Save] Failed:', e);
    }
  },

  async processPendingReports() {
    try {
      const pending = JSON.parse(localStorage.getItem('incident-pending') || '[]');
      if (!pending.length || !this.api) return;

      for (const report of pending) {
        try {
          this.reportData = report.data;
          await this.submitToGeotab();
        } catch (e) {
          console.warn('[Pending] Failed to submit:', e);
          return; // Stop on first failure, try again later
        }
      }
      localStorage.removeItem('incident-pending');
    } catch (e) {
      console.warn('[Pending] Process failed:', e);
    }
  },

  // ---- Helpers ----
  setEl(id, text) {
    const el = document.getElementById(id);
    if (el) el.textContent = text;
  }
};

// Initialize the add-in registration
app.initializeAddin();

// If running outside Geotab Drive (dev mode), auto-start
if (!window._geotabDrive) {
  document.addEventListener('DOMContentLoaded', () => {
    // Check if dev harness will call initialize
    setTimeout(() => {
      if (!app.api) {
        console.log('[Dev] No API injected, running standalone');
        app.onInitialized();
      }
    }, 500);
  });
}
