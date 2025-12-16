// ==UserScript==
// @name         Genesys Queue Monitor (v2.2 - Dashboard Widget)
// @namespace    http://tampermonkey.net/
// @version      2.2
// @description  Dashboard flottant Genesys Cloud basé sur le widget AGENT_STATUS
// @author       Miloud Mostefa-Hanchour
// @match        https://apps.mypurecloud.de/*
// @match        https://apps.mypurecloud.de/directory/*
// @grant        none
// @run-at       document-idle
// @noframes
// ==/UserScript==

(function() {
  'use strict';

  console.log('[QM] userscript loaded on', window.location.href);

  // Anti-iframe + anti double instance
  if (window.top !== window.self) return;
  if (window.__QUEUE_MONITOR_REPORT_RUNNING__) return;
  window.__QUEUE_MONITOR_REPORT_RUNNING__ = true;

  // ===================== CONFIG =====================
  const CONFIG = {
    REFRESH_INTERVAL: 1500,
    REFRESH_INTERVAL_HIDDEN: 4000,
    HIDDEN_BACKOFF_AFTER_MS: 60_000,
    TICK_INTERVAL: 1000,
    DEBUG: false,
    STORAGE_PREFIX: 'queueMonitorReport_',

    MENU_DEBOUNCE: 500,
    CALL_ALERT_SECS: 10 * 60,
    CHAT_ALERT_SECS: 10 * 60,
    CALL_CLEAR_TICKS: 3,

    PANEL_MIN_W: 480,
    PANEL_MIN_H: 560,
    ALERTS_ENABLED: true,
    NOTIFICATIONS_ENABLED: true,
    SHOW_CHAT_MULTIPLIER: false
  };

  const log = {
    debug: (...a) => CONFIG.DEBUG && console.log('[QM-R]', ...a),
    info:  (...a) => console.log('[QM-R INFO]', ...a),
    warn:  (...a) => console.warn('[QM-R WARN]', ...a),
    error: (...a) => console.error('[QM-R ERROR]', ...a),
  };

  // ===================== PERSISTENCE =====================
  class Storage {
    static get(key, def = null) {
      try {
        const v = localStorage.getItem(CONFIG.STORAGE_PREFIX + key);
        return v ? JSON.parse(v) : def;
      } catch (e) {
        log.error('Storage get', e);
        return def;
      }
    }
    static set(key, val) {
      try {
        localStorage.setItem(CONFIG.STORAGE_PREFIX + key, JSON.stringify(val));
      } catch (e) {
        log.error('Storage set', e);
      }
    }
    static clear() {
      Object.keys(localStorage)
        .filter(k => k.startsWith(CONFIG.STORAGE_PREFIX))
        .forEach(k => localStorage.removeItem(k));
    }
  }

  let masterSlotList       = Storage.get('masterSlotList', []);
  let connectedUsers       = Storage.get('connectedUsers', []);
  let lastGroup            = Storage.get('lastGroup', {});
  let slotCounter          = Storage.get('slotCounter', 1);
  let lastStatusKey        = Storage.get('lastStatusKey', {});
  let statusStartAt        = Storage.get('statusStartAt', {});
  let callStartAt          = Storage.get('callStartAt', {});
  let chatStartAt          = Storage.get('chatStartAt', {});
  let lastInCall           = Storage.get('lastInCall', {});
  let callOffStreak        = Storage.get('callOffStreak', {});
  let lastInChat           = Storage.get('lastInChat', {});
  let chatOffStreak        = Storage.get('chatOffStreak', {});
  let muted                = Storage.get('muted', false);
  let snoozeUntil          = Storage.get('snoozeUntil', 0);
  let historyByDay         = Storage.get('historyByDay', {});
  let dailyAgg             = Storage.get('dailyAgg', {});
  let lastProhibSubtype    = Storage.get('lastProhibSubtype', {});

  let favorites            = new Set(Storage.get('favorites', []));
  let sectionVisibility    = Storage.get('sectionVisibility', null) || {};
  let searchFilter         = Storage.get('searchFilter', '');
  let minimized            = Storage.get('minimized', false);
  let panelPos             = Storage.get('panelPos', null);
  let panelSize            = Storage.get('panelSize', null);
  let sortCallsOrder       = Storage.get('sortCallsOrder', 'desc');
  let sortStatusOrder      = Storage.get('sortStatusOrder', '');

  const save = () => {
    Storage.set('masterSlotList', masterSlotList);
    Storage.set('connectedUsers', connectedUsers);
    Storage.set('lastGroup', lastGroup);
    Storage.set('slotCounter', slotCounter);
    Storage.set('lastStatusKey', lastStatusKey);
    Storage.set('statusStartAt', statusStartAt);
    Storage.set('callStartAt', callStartAt);
    Storage.set('chatStartAt', chatStartAt);
    Storage.set('lastInCall', lastInCall);
    Storage.set('callOffStreak', callOffStreak);
    Storage.set('lastInChat', lastInChat);
    Storage.set('chatOffStreak', chatOffStreak);
    Storage.set('favorites', [...favorites]);
    Storage.set('sectionVisibility', sectionVisibility);
    Storage.set('searchFilter', searchFilter);
    Storage.set('minimized', minimized);
    Storage.set('panelPos', panelPos);
    Storage.set('panelSize', panelSize);
    Storage.set('sortCallsOrder', sortCallsOrder);
    Storage.set('sortStatusOrder', sortStatusOrder);
    Storage.set('muted', muted);
    Storage.set('snoozeUntil', snoozeUntil);
    Storage.set('historyByDay', historyByDay);
    Storage.set('dailyAgg', dailyAgg);
    Storage.set('lastProhibSubtype', lastProhibSubtype);
  };

  // ===================== HELPERS =====================
  const txt        = el => (el && (el.innerText || el.textContent) || '').trim();
  const nowIso     = () => new Date().toISOString();
  const todayKey   = () => new Date().toISOString().slice(0, 10);
  const msSince    = iso => (iso ? (Date.now() - new Date(iso).getTime()) : 0);
  const clamp      = (n, min, max) => Math.max(min, Math.min(max, n));
  const escapeHtml = s => (s || '').replace(/[&<>"']/g, m => ({
    '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;','\'':'&#39;'
  }[m]));
  const norm       = s => (s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
  const toHMS      = ms => {
    const total = Math.floor(ms / 1000);
    const h = Math.floor(total / 3600);
    const m = Math.floor((total % 3600) / 60);
    const s = total % 60;
    return String(h).padStart(2, '0') + ':' +
           String(m).padStart(2, '0') + ':' +
           String(s).padStart(2, '0');
  };

  // Document + iframes (même origine)
  function getAllDocs() {
    const docs = [document];
    const iframes = document.querySelectorAll('iframe');
    iframes.forEach(f => {
      try {
        const doc = f.contentDocument || (f.contentWindow && f.contentWindow.document);
        if (doc && doc.documentElement) docs.push(doc);
      } catch (e) {
        // cross origin → ignore
      }
    });
    return docs;
  }

  // ===================== SELECTEURS DE CANAUX =====================
  const PHONE_SELECTORS = [
    'gux-icon[icon-name*="phone"]',
    'gux-icon[icon-name*="call"]',
    'svg[aria-label="Voix"]',
    'svg[aria-label="voix"]',
    'svg[aria-label="Voice"]',
    'svg[aria-label="voice"]',
    'svg[aria-label="Appel"]',
    'svg[aria-label="appel"]',
    'svg[aria-label*="Voix"]',
    'svg[aria-label*="voix"]',
    'svg[aria-label*="Voice"]',
    'svg[aria-label*="voice"]',
    'svg[aria-label*="Appel"]',
    'svg[aria-label*="appel"]'
  ];
  const CHAT_SELECTORS = [
    'gux-icon[icon-name*="chat"]',
    'gux-icon[icon-name*="message"]'
  ];

  const EMAIL_SELECTORS = [
    'gux-icon[icon-name*="email"]',
    'gux-icon[icon-name*="mail"]'
  ];
  const SMS_SELECTORS = [
    'gux-icon[icon-name*="sms"]',
    'gux-icon[icon-name*="text"]'
  ];
  const TASK_SELECTORS = [
    'gux-icon[icon-name*="task"]',
    'gux-icon[icon-name*="work"]',
    'gux-icon[icon-name*="workitem"]',
    '[class*="workitem"]',
    '[class*="work-item"]'
  ];

  // ===================== STATUTS (depuis widget) =====================
  function getStatusClassFromWidget(statusElement, explicitText) {
    const text = (explicitText || (statusElement ? statusElement.textContent : '') || '')
      .trim()
      .toLowerCase();
    const classes = (statusElement && statusElement.className
      ? String(statusElement.className).toLowerCase()
      : '');

    // CORRECTION : "Occupé" est différent de "En appel"
    if (text.includes('occupé') || text.includes('busy')) return 'Busy';
    if (text.includes('interaction') || text.includes('en cours d\'interaction') || text.includes('en cours de communication')) return 'On Call';
    if (text.includes('tâche associée') || text.includes('associated task') || text.includes('tâche')) return 'Associated Task';
    if (classes.includes('meal') || text.includes('repas')) return 'Meal';
    if (classes.includes('break') || text.includes('pause')) return 'Break';
    if (classes.includes('meeting') || text.includes('réunion')) return 'Meeting';
    if (classes.includes('training') || text.includes('formation')) return 'Training';
    if (classes.includes('available') || text.includes('disponible')) return 'Available';
    if (classes.includes('away') || text.includes('absent')) return 'Away';
    if (classes.includes('on_queue') || classes.includes('on-queue') ||
        text.includes('file d\'attente') || text.includes('en file') || text.includes('queue'))
      return 'On Queue';
    if (classes.includes('on_call') || classes.includes('on-call') || text.includes('en appel'))
      return 'On Call';
    // NOUVEAU : Hors ligne détecté
    if (classes.includes('offline') || text.includes('hors ligne') || text.includes('déconnecté'))
      return 'Offline';
    if (classes.includes('idle') || text.includes('non occupé') || text.includes('inactif'))
      return 'Idle';
    if (text.includes('sans réponse') || text.includes('not responding'))
      return 'Not Responding';

    return 'default';
  }

  function getStatusClassFromDot(dot, explicitText) {
    const t = (explicitText || '').toLowerCase();

    // 0) Le texte domine
    // CORRECTION : "Occupé" ≠ "En appel"
    if (/\b(occupé|busy)\b/.test(t)) return 'Busy';
    if (/\b(interaction|interacting|en cours d'interaction|en cours de communication)\b/.test(t)) return 'On Call';
    if (/\b(file d'attente|en\s*file|on\s*queue)\b/.test(t)) return 'On Queue';
    if (/\b(tâche associée|tache associee|associated task|work item)\b/.test(t)) return 'Associated Task';
    if (/\b(non occupé|inactif|idle)\b/.test(t)) return 'Idle';
    // NOUVEAU : Hors ligne
    if (/\b(hors ligne|offline|déconnecté)\b/.test(t)) return 'Offline';

    // 1) Sinon, on regarde les classes du rond
    if (!dot) return getStatusClassFromWidget(null, explicitText);
    const classes = String(dot.className).toLowerCase();

    if (classes.includes('busy')) return 'Busy'; // CORRECTION : "busy" = "Occupé", pas "En call"
    if (classes.includes('interacting') || classes.includes('on_call')) return 'On Call';
    if (classes.includes('on_queue')) return 'On Queue';
    if (classes.includes('available')) return 'Available';
    if (classes.includes('away')) return 'Away';
    if (classes.includes('offline')) return 'Offline'; // NOUVEAU
    if (classes.includes('idle')) return 'Idle';
    if (classes.includes('break')) return 'Break';

    // 2) Fallback : heuristique texte
    return getStatusClassFromWidget(null, explicitText);
  }

  // ===================== STATUTS PROHIBÉS =====================
  function isRonaStatus(s) {
    const n = norm(s);
    return n.includes('sans reponse') ||
           n.includes('rona') ||
           n.includes('no answer') ||
           n.includes('ring no answer') ||
           n.includes('not answered');
  }

  function isPostcallStatus(s) {
    const n = norm(s);
    return n.includes('travail apres appel') ||
           n.includes('postcall') ||
           n.includes('after call work') ||
           n.includes('acw') ||
           n.includes('wrap');
  }

  function prohibSubtypeFromStatus(s) {
    if (isRonaStatus(s)) return 'RONA';
    if (isPostcallStatus(s)) return 'Postcall';
    return '';
  }

  // ===================== SECTIONS (AJOUT DE "OCCUPÉ" et EXCLUSION "HORS LIGNE") =====================
  const SECTION_DEFS = [
    { key: 'favoris',          label: 'Favoris' },
    { key: 'prohib',           label: 'Statut prohibé' },
    { key: 'disponible',       label: 'Disponible' },
    { key: 'queue_free',       label: "En file d'attente (sans call)" },
    { key: 'en_call',          label: 'En call' },
    { key: 'en_chat',          label: 'En chat' },
    { key: 'tache',            label: 'Tâche associée' },
    { key: 'occupe',           label: 'Occupé' }, // NOUVELLE SECTION
    { key: 'non_telecontact',  label: 'Non télécontact' },
    { key: 'travaux',          label: 'Travaux payants' },
    { key: 'pause',            label: 'Pause' },
    { key: 'repas',            label: 'Repas' },
    { key: 'reunion',          label: 'Réunion' },
    { key: 'formation',        label: 'Formation' },
    { key: 'interaction_hf',   label: 'En interaction (hors file)' },
    { key: 'autre',            label: 'Autre' }
  ];

  if (Object.keys(sectionVisibility).length === 0) {
    sectionVisibility = Object.fromEntries(SECTION_DEFS.map(s => [s.key, true]));
  } else {
    SECTION_DEFS.forEach(s => {
      if (typeof sectionVisibility[s.key] === 'undefined') {
        sectionVisibility[s.key] = true;
      }
    });
  }

  // ===================== PERF / BACKOFF =====================
  let hiddenSince = 0;
  document.addEventListener('visibilitychange', () => {
    hiddenSince = document.hidden ? Date.now() : 0;
  });

  function effectiveRefreshMs() {
    if (!document.hidden) return CONFIG.REFRESH_INTERVAL;
    const hiddenFor = Date.now() - (hiddenSince || Date.now());
    return hiddenFor >= CONFIG.HIDDEN_BACKOFF_AFTER_MS
      ? CONFIG.REFRESH_INTERVAL_HIDDEN
      : CONFIG.REFRESH_INTERVAL;
  }

  // ===================== HISTORIQUE =====================
  function histAdd(name, type, payload) {
    if (!name) return;
    const day = todayKey();
    if (!historyByDay[day]) historyByDay[day] = {};
    if (!historyByDay[day][name]) historyByDay[day][name] = [];
    historyByDay[day][name].push(Object.assign({ ts: Date.now(), type }, payload || {}));
    const arr = historyByDay[day][name];
    if (arr.length > 200) arr.splice(0, arr.length - 200);
    save();
  }

  function aggAddMs(name, field, ms) {
    if (!name || !field) return;
    const day = todayKey();
    if (!dailyAgg[day]) dailyAgg[day] = {};
    if (!dailyAgg[day][name]) {
      dailyAgg[day][name] = { callMs: 0, chatMs: 0, postcallMs: 0, ronaMs: 0 };
    }
    const capped = clamp(ms, 0, 8 * 60 * 60 * 1000);
    dailyAgg[day][name][field] =
      (dailyAgg[day][name][field] || 0) + capped;
    save();
  }

  function histPreview(name) {
    const day = todayKey();
    const arr = (historyByDay[day] && historyByDay[day][name]) ? historyByDay[day][name] : [];
    const last = arr.slice(-5).reverse().map(ev => {
      const t = new Date(ev.ts).toLocaleTimeString();
      if (ev.type === 'call_end')   return t + ' • Fin appel (' + toHMS(ev.durMs) + ')';
      if (ev.type === 'chat_end')   return t + ' • Fin chat (' + toHMS(ev.durMs) + ')';
      if (ev.type === 'status')     return t + ' • Statut → ' + (ev.to || '');
      if (ev.type === 'prohib_on')  return t + ' • Prohibé (' + (ev.sub || '') + ')';
      if (ev.type === 'prohib_off') return t + ' • Fin prohibé (' + (ev.sub || '') + ', ' + toHMS(ev.durMs) + ')';
      return t + ' • ' + ev.type;
    }).join('\n');
    return last || 'Aucun historique';
  }

  // ===================== SLOTS =====================
  class SlotManager {
    static assignSlot(name) {
      if (!name) return null;
      const ex = masterSlotList.find(i => i.name === name);
      if (ex) {
        ex.lastSeen = Date.now();
        return ex.slot;
      }
      const slot = slotCounter++;
      masterSlotList.push({
        name,
        slot,
        assignedAt: nowIso(),
        totalCalls: 0,
        lastCallAt: null,
        lastSeen: Date.now()
      });
      save();
      return slot;
    }

    static reorderAfterCall(name) {
      const idx = masterSlotList.findIndex(i => i.name === name);
      if (idx === -1) return;
      const it = masterSlotList[idx];
      it.totalCalls = (it.totalCalls || 0) + 1;
      it.lastCallAt = nowIso();
      const [m] = masterSlotList.splice(idx, 1);
      masterSlotList.push(m);
      masterSlotList.forEach((row, i) => { row.slot = i + 1; });
      save();
    }

    static getSlot(name) {
      const it = masterSlotList.find(i => i.name === name);
      if (!it) return null;
      it.lastSeen = Date.now();
      return it.slot;
    }

    static getAllBySlot() {
      return [...masterSlotList].sort((a, b) => a.slot - b.slot);
    }

    static cleanup(activeDays = 7) {
      const cutoff = Date.now() - activeDays * 24 * 60 * 60 * 1000;
      const before = masterSlotList.length;
      masterSlotList = masterSlotList.filter(i => !i.lastSeen || i.lastSeen > cutoff);
      if (masterSlotList.length !== before) {
        log.info('Nettoyage slots: ' + (before - masterSlotList.length) + ' supprimé(s)');
        save();
      }
    }
  }

  // ===================== CONNECTED USERS =====================
  class ConnectedUsersManager {
    static updateConnectedUsers(agents) {
      // FILTRE : Ne garde que les agents qui ne sont PAS "Hors ligne"
      const onlineAgents = agents.filter(agent => {
        const statusText = agent.status || '';
        const statusClass = agent.statusClass || '';
        const isOffline = statusText.includes('Hors ligne') || 
                         statusClass.includes('Offline') ||
                         statusText.includes('Hors ligne');
        return !isOffline;
      });
      
      const names = onlineAgents.map(a => a.name).filter(Boolean);
      const unique = Array.from(new Set(names));
      unique.forEach(n => SlotManager.assignSlot(n));

      connectedUsers = unique;
      save();
    }

    static getConnectedBySlot() {
      return connectedUsers
        .map(name => ({
          name,
          slot: SlotManager.getSlot(name) || 999
        }))
        .sort((a, b) => a.slot - b.slot);
    }
  }

  // ===================== PARSEUR DE DURÉES =====================
  const CALL_DURATION_SELECTORS = [
    '.time-duration',
    '.call-duration',
    '.interaction-duration',
    '.duration',
    '[class*="duration"]',
    '[data-test*="duration"]',
    '.timer',
    '[class*="timer"]'
  ];

  const TIME_RE = /(\d{1,2}):([0-5]\d)(?::([0-5]\d))?/;

  function parseDurationSecFromString(text) {
    if (!text) return null;
    const s = String(text).replace(/[\u202F\u00A0]/g, ' ').trim().toLowerCase();

    // Cas "jours"
    const joursMatch = s.match(/(\d+)\s*jour/);
    if (joursMatch) {
      const jours = parseInt(joursMatch[1], 10);
      if (!Number.isNaN(jours)) {
        return jours * 24 * 3600;
      }
    }

    // hh:mm[:ss] ou mm:ss
    const m = s.match(TIME_RE);
    if (m) {
      const hasSec = m[3] != null;
      const h = hasSec ? parseInt(m[1], 10) : 0;
      const mm = hasSec ? parseInt(m[2], 10) : parseInt(m[1], 10);
      const ss = hasSec ? parseInt(m[3], 10) : parseInt(m[2], 10);
      if ([h, mm, ss].some(Number.isNaN)) return null;
      return h * 3600 + mm * 60 + ss;
    }

    // formats verbaux 1h 02m 03s / 4m 46s etc.
    let total = 0;
    let found = false;
    const unitRe = /(\d+)\s*(h|hr|hrs|hour|hours|heure|heures|m|min|mn|minute|minutes|s|sec|secs|second|seconds|seconde|secondes)\b/g;
    let u;
    while ((u = unitRe.exec(s)) !== null) {
      const n = parseInt(u[1], 10);
      if (Number.isNaN(n)) continue;
      const unit = u[2];
      if (/^(h|hr|hrs|hour|hours|heure|heures)$/.test(unit)) total += n * 3600;
      else if (/^(m|min|mn|minute|minutes)$/.test(unit)) total += n * 60;
      else total += n;
      found = true;
    }
    if (found) return total;

    // compacts 1h2m3s
    const compactRe = /(?:(\d+)h)?\s*(?:(\d+)m|(\d+)min|(\d+)mn)?\s*(?:(\d+)s|(\d+)sec)?/;
    const c = s.match(compactRe);
    if (c && c[0].trim() !== '') {
      const hh = c[1] ? parseInt(c[1], 10) : 0;
      const mm = c[2] ? parseInt(c[2], 10)
        : c[3] ? parseInt(c[3], 10)
        : c[4] ? parseInt(c[4], 10)
        : 0;
      const ss = c[5] ? parseInt(c[5], 10)
        : c[6] ? parseInt(c[6], 10)
        : 0;
      if ([hh, mm, ss].some(Number.isNaN)) return null;
      if (hh || mm || ss) return hh * 3600 + mm * 60 + ss;
    }

    return null;
  }

  // Lit une durée d'interaction depuis le widget
  function extractCallDurationFromRow(row) {
    if (!row) return null;

    // Cherche la cellule de durée unifiée
    const durationCell = row.querySelector('.column-unifiedDuration') || row;
    const timeInStatus = durationCell.querySelector('.time-in-status');
    
    if (timeInStatus) {
      const sec = parseDurationSecFromString(timeInStatus.textContent || '');
      if (sec != null) return sec;
    }

    // Fallback: chercher dans toute la cellule
    const bulk = (durationCell.innerText || durationCell.textContent || '').slice(0, 1500);
    const sec2 = parseDurationSecFromString(bulk);
    return sec2 != null ? sec2 : null;
  }

  // Durée d'appel pour un agent en "Tâche associée"
  function extractTaskCallDurationFromRow(row) {
    if (!row) return null;
    return extractCallDurationFromRow(row);
  }

  // Durée de statut depuis le widget
  function extractStatusDurationFromRow(row) {
    if (!row) return null;

    // Cherche la cellule de durée unifiée
    const durationCell = row.querySelector('.column-unifiedDuration') || row;
    const timeInStatus = durationCell.querySelector('.time-in-status');
    
    if (timeInStatus) {
      const sec = parseDurationSecFromString(timeInStatus.textContent || '');
      if (sec != null) return sec;
    }

    // Fallback
    const bulk = (durationCell.innerText || durationCell.textContent || '').slice(0, 300);
    const sec2 = parseDurationSecFromString(bulk);
    return sec2 != null ? sec2 : null;
  }

  // === CHAT: durées des chats
  function extractChatDurationsFromAgentRow(row) {
    const out = [];
    if (!row) return out;
    
    // Dans le widget AGENT_STATUS, on a généralement qu'une durée (celle du statut)
    const durationCell = row.querySelector('.column-unifiedDuration') || row;
    const timeInStatus = durationCell.querySelector('.time-in-status');
    
    if (timeInStatus) {
      const sec = parseDurationSecFromString(timeInStatus.textContent || '');
      if (sec != null) out.push(sec);
    }
    
    return out.slice(0, 2);
  }

  // ===================== DÉTECTION AGENTS DEPUIS WIDGET =====================
  function findAgentsInWidget() {
    const agents = [];
    const seen = new Set();

    getAllDocs().forEach(doc => {
      // Cherche le widget AGENT_STATUS
      const widget = doc.querySelector('.widget-type-AGENT_STATUS');
      if (!widget) return;

      // Trouve toutes les lignes d'agents dans la table
      const rows = widget.querySelectorAll('tbody tr[data-row-id]');
      rows.forEach(row => {
        try {
          // Nom de l'agent
          const nameEl = row.querySelector('.column-agent a');
          if (!nameEl) return;
          const name = (nameEl.textContent || '').trim();
          if (!name || seen.has(name)) return;

          // Statut texte et dot
          const statusCell = row.querySelector('.column-statusAndPresence');
          const dot = statusCell ? statusCell.querySelector('.entity-v3-presence-indicator-dot') : null;
          
          const additionalLabel = statusCell ? statusCell.querySelector('.additional-label') : null;
          const statusText = additionalLabel ? (additionalLabel.textContent || '').trim() : '';
          
          // CORRECTION CRITIQUE : Détection spécifique de "Tâche associée"
          const isTacheAssociee = statusText.includes('Tâche associée') || 
                                  statusText.includes('Tâche') || 
                                  statusText.includes('Associated Task') ||
                                  statusText.includes('Work item');
          
          const statusClass = getStatusClassFromDot(dot, statusText);

          // CORRECTION : "Occupé" n'est PAS un appel
          const isOccupied = statusText.includes('Occupé') || statusClass === 'Busy';
          
          // "En cours d'interaction" ou "En cours de communication" = vrai appel
          const isInteracting = !isTacheAssociee && !isOccupied &&
                               (statusText.includes('En cours d\'interaction') ||
                                statusText.includes('En cours de communication'));

          // NOUVEAU : Vérifie si l'agent est "Hors ligne"
          const isOffline = statusText.includes('Hors ligne') || 
                           statusClass.includes('Offline');

          const agent = {
            id: 'widget_' + norm(name),
            name,
            status: statusText || statusClass || '',
            statusClass: statusClass || 'default',
            activityCount: isInteracting ? 1 : 0,
            onQueue: statusText.includes('En file d\'attente') || statusClass === 'On Queue',
            channel: {
              call: isInteracting && !isTacheAssociee && !isOccupied,  // Pas de call si AT ou Occupé
              chat: false,
              email: false,
              sms: false,
              task: isTacheAssociee,
              chatCount: 0
            },
            rowElement: row,
            timestamp: Date.now(),
            // Champs pour détection spéciale
            isTacheAssociee: isTacheAssociee,
            isOccupied: isOccupied,
            isOffline: isOffline  // NOUVEAU
          };

          // FILTRE : Ne pas ajouter les agents "Hors ligne"
          if (!isOffline) {
            seen.add(name);
            agents.push(agent);
          }
        } catch (e) {
          log.error('Erreur extraction agent widget:', e);
        }
      });
    });

    return agents;
  }

  // ===================== MAPPAGE STATUT → SECTION =====================
  function deriveStatusKey(agent) {
    const raw = agent.status || agent.statusClass || '';
    const st = raw.toLowerCase();
    const noacc = norm(raw);
    const ch = agent.channel || {};
    const onQ = !!agent.onQueue;
    const cnt = agent.activityCount || 0;
    
    // NOUVEAU : Vérifie d'abord si l'agent est "Occupé"
    if (agent.isOccupied) return 'occupe';
    
    // Vérifie ensuite si l'agent a explicitement le flag isTacheAssociee
    if (agent.isTacheAssociee) return 'tache';

    // Prohibés en premier
    if (isRonaStatus(raw) || isPostcallStatus(raw)) return 'prohib';

    // Statuts textuels qui doivent DOMINER l'auto-détection
    if (
      st.includes('tâche') || st.includes('tache') ||
      noacc.includes('tache associee') ||
      st.includes('work item') || st.includes('workitem') ||
      st.includes('associated task')
    ) return 'tache';

    // CORRECTION : "Occupé" va dans la nouvelle section
    if (st.includes('occupé') || st.includes('busy')) return 'occupe';

    if (
      st.includes('non télé') || st.includes('non-télé') ||
      noacc.includes('non telecontact') || noacc.includes('non-telecontact') ||
      noacc.includes('non tele contact') || noacc.includes('non telec')
    ) return 'non_telecontact';

    if (st.includes('pause') || st.includes('break')) return 'pause';
    if (st.includes('repas') || st.includes('meal')) return 'repas';
    if (st.includes('réunion') || st.includes('reunion') || st.includes('meeting')) return 'reunion';
    if (st.includes('formation') || st.includes('training')) return 'formation';
    if (st.includes('travaux pay')) return 'travaux';

    // Digital / voix
    if (ch.chat) return 'en_chat';

    // Voix uniquement si Genesys nous indique explicitement un appel
    // ET que ce n'est PAS une tâche associée OU Occupé
    if (!agent.isTacheAssociee && !agent.isOccupied && 
        (ch.call || agent.statusClass === 'On Call' || 
        st.includes('interaction') || st.includes('en cours d\'interaction'))) return 'en_call';

    // File sans interaction
    if (onQ && cnt === 0) return 'queue_free';

    // Disponibles / interactions hors file
    if (!onQ && cnt > 0) return 'interaction_hf';
    if (!onQ && (st.includes('available') || st.includes('disponible'))) return 'disponible';
    if (st.includes('non occupé') || st.includes('inactif') || st.includes('idle')) return 'disponible';

    return 'autre';
  }

  // ===================== TIMERS STATUT / CALL / CHAT =====================
  function ensureStatusTimerOnly(key, name) {
    if (!name) return;
    const prevKey = lastStatusKey[name];

    if (prevKey === 'prohib' && key !== 'prohib') {
      const start = statusStartAt[name];
      const dur = msSince(start);
      const sub = lastProhibSubtype[name] || '';
      if (sub === 'RONA') aggAddMs(name, 'ronaMs', dur);
      else if (sub === 'Postcall') aggAddMs(name, 'postcallMs', dur);
      histAdd(name, 'prohib_off', { sub, durMs: dur });
      delete lastProhibSubtype[name];
    }

    if (key === 'prohib' && prevKey !== 'prohib') {
      const sub = prohibSubtypeFromStatus((key && key.status) || '');
      lastProhibSubtype[name] = sub;
      histAdd(name, 'prohib_on', { sub });
    }

    if (lastStatusKey[name] !== key) {
      lastStatusKey[name] = key;
      statusStartAt[name] = nowIso();
      histAdd(name, 'status', { to: key });
    }

    save();
  }

  function updateCallState(name, inCallConnected, genesysSec) {
    const prev = !!lastInCall[name];

    if (inCallConnected) {
      callOffStreak[name] = 0;
      if (!prev) {
        const startIso = genesysSec != null
          ? new Date(Date.now() - genesysSec * 1000).toISOString()
          : nowIso();
        callStartAt[name] = startIso;
        lastInCall[name] = true;
      } else if (genesysSec != null && callStartAt[name]) {
        const currentMs = msSince(callStartAt[name]);
        const desiredMs = genesysSec * 1000;
        const diff = desiredMs - currentMs;
        if (Math.abs(diff) > 5000) {
          callStartAt[name] = new Date(Date.now() - desiredMs).toISOString();
        }
      }
    } else {
      if (prev) {
        callOffStreak[name] = (callOffStreak[name] || 0) + 1;
        if (callOffStreak[name] >= CONFIG.CALL_CLEAR_TICKS) {
          const dur = msSince(callStartAt[name] || nowIso());
          aggAddMs(name, 'callMs', dur);
          histAdd(name, 'call_end', { durMs: dur });
          delete callStartAt[name];
          lastInCall[name] = false;
          callOffStreak[name] = 0;
          SlotManager.reorderAfterCall(name);
        }
      } else {
        callOffStreak[name] = 0;
      }
    }

    save();
  }

  function updateChatState(name, chatCount) {
    if (!name) return;
    const prevCount = (lastInChat[name] && lastInChat[name].length) || 0;

    if (!chatStartAt[name]) chatStartAt[name] = [];
    if (!lastInChat[name]) lastInChat[name] = [];
    if (!chatOffStreak[name]) chatOffStreak[name] = [];

    // Ouvre les nouveaux chats
    for (let i = 0; i < chatCount; i++) {
      if (!lastInChat[name][i]) {
        chatStartAt[name][i] = nowIso();
        lastInChat[name][i] = true;
        chatOffStreak[name][i] = 0;
      }
    }

    // Ferme les chats en trop
    for (let i = chatCount; i < prevCount; i++) {
      if (lastInChat[name][i]) {
        chatOffStreak[name][i] = (chatOffStreak[name][i] || 0) + 1;
        if (chatOffStreak[name][i] >= CONFIG.CALL_CLEAR_TICKS) {
          const dur = msSince(chatStartAt[name][i] || nowIso());
          aggAddMs(name, 'chatMs', dur);
          histAdd(name, 'chat_end', { durMs: dur });
          lastInChat[name][i] = false;
          chatStartAt[name][i] = null;
          chatOffStreak[name][i] = 0;
        }
      }
    }

    // Nettoyage
    if (chatStartAt[name]) {
      chatStartAt[name] = chatStartAt[name].slice(0, chatCount);
    }
    if (lastInChat[name]) {
      lastInChat[name] = lastInChat[name].slice(0, chatCount);
    }
    if (chatOffStreak[name]) {
      chatOffStreak[name] = chatOffStreak[name].slice(0, chatCount);
    }

    save();
  }

  // ===================== KPIs =====================
  function computeKpis(agents) {
    const res = {
      connected: agents.length,
      onQueue: agents.filter(a => a.onQueue).length,
      longestCall: { name: '', ms: 0 },
      longestChat: { name: '', ms: 0 },
      prohibCount: 0,
      occupeCount: 0
    };

    for (const a of agents) {
      if (lastInCall[a.name] && callStartAt[a.name]) {
        const ms = msSince(callStartAt[a.name]);
        if (ms > res.longestCall.ms) {
          res.longestCall = { name: a.name, ms };
        }
      }
      if (lastInChat[a.name] && Array.isArray(chatStartAt[a.name]) && chatStartAt[a.name][0]) {
        const ms = msSince(chatStartAt[a.name][0]);
        if (ms > res.longestChat.ms) {
          res.longestChat = { name: a.name, ms };
        }
      }
      const key = deriveStatusKey(a);
      if (key === 'prohib') res.prohibCount++;
      if (key === 'occupe') res.occupeCount++;
    }

    return res;
  }

  // ====== ALERTES / NOTIFS / AUDIO ======
  let callAlerted   = {};
  let chatAlerted   = {};
  let prohibAlerted = {};
  let activeAlerts  = { total: 0 };

  function canNotify() {
    if (muted) return false;
    if (Date.now() < (snoozeUntil || 0)) return false;
    return CONFIG.ALERTS_ENABLED;
  }

  function requestNotifyPermission() {
    try {
      if (CONFIG.NOTIFICATIONS_ENABLED &&
          'Notification' in window &&
          Notification.permission === 'default') {
        Notification.requestPermission(() => {});
      }
    } catch (e) {}
  }

  let audioCtx = null;
  let audioReady = false;

  function primeAudio() {
    try {
      audioCtx = audioCtx || new (window.AudioContext || window.webkitAudioContext)();
      audioReady = true;
    } catch (e) {
      audioReady = false;
    }
  }

  function beep(freq, durMs) {
    if (!canNotify() || !audioReady) return;
    const f = freq || 880;
    const d = (durMs || 180) / 1000;
    try {
      const now = audioCtx.currentTime;
      const osc = audioCtx.createOscillator();
      const gain = audioCtx.createGain();
      osc.connect(gain);
      gain.connect(audioCtx.destination);
      osc.type = 'sine';
      osc.frequency.value = f;
      gain.gain.setValueAtTime(0.0001, now);
      gain.gain.exponentialRampToValueAtTime(0.2, now + 0.02);
      gain.gain.exponentialRampToValueAtTime(0.0001, now + d);
      osc.start(now);
      osc.stop(now + d + 0.02);
    } catch (e) {}
  }

  function notify(title, body) {
    if (!canNotify() || !CONFIG.NOTIFICATIONS_ENABLED) return;
    try {
      if ('Notification' in window && Notification.permission === 'granted') {
        new Notification(title, { body });
      }
    } catch (e) {}
  }

  // ====== STYLES UI ======
  function injectPulseStyle() {
    if (document.getElementById('qm-pulse-style')) return;
    const reduce = matchMedia('(prefers-reduced-motion: reduce)').matches;
    const st = document.createElement('style');
    st.id = 'qm-pulse-style';
    st.textContent = reduce
      ? '.qm-card.qm-pulse{}'
      : '@keyframes qmPulse{0%{box-shadow:0 0 0 0 rgba(239,68,68,.6)}70%{box-shadow:0 0 0 10px rgba(239,68,68,0)}100%{box-shadow:0 0 0 0 rgba(239,68,68,0)}} .qm-card.qm-pulse{animation:qmPulse 1.8s infinite}';
    document.head.appendChild(st);
  }

  function injectThemeStyles() {
    if (document.getElementById('qm-theme-style')) return;
    const st = document.createElement('style');
    st.id = 'qm-theme-style';
    st.textContent = `
      @media (prefers-color-scheme: dark){
        #queue-monitor-panel{
          background:linear-gradient(135deg,#0f172a 0%,#111827 100%) !important;
          color:#e5e7eb !important;
          border-color:#1f2937 !important;
        }
        #queue-monitor-panel .qm-card{
          color:#e5e7eb !important;
        }
        #qm-sections-menu{
          background:#111827 !important;
          color:#e5e7eb !important;
          border-color:#1f2937 !important;
        }
        #queue-monitor-panel input,
        #queue-monitor-panel select{
          background:#0b1220 !important;
          color:#e5e7eb !important;
        }
        #queue-monitor-panel option{
          color:#e5e7eb !important;
          background:#0b1220 !important;
        }
        #qm-snooze{
          background:#111827 !important;
          color:#e5e7eb !important;
          border-color:#1f2937 !important;
        }
        #qm-snooze button{
          background:#0b1220 !important;
          color:#e5e7eb !important;
        }
        #qm-snooze button:hover{
          background:#111827 !important;
        }
      }
    `;
    document.head.appendChild(st);
  }

  function injectUiFixStyles() {
    if (document.getElementById('qm-ui-fix-style')) return;
    const st = document.createElement('style');
    st.id = 'qm-ui-fix-style';
    st.textContent = `
      #queue-monitor-panel input,
      #queue-monitor-panel select{
        color:#111 !important;
        background:#fff !important;
      }
      #queue-monitor-panel option{
        color:#111 !important;
      }
      #queue-monitor-panel input::placeholder{
        color:#6b7280 !important;
        opacity:1;
      }
      #queue-monitor-panel .kpi-chip{
        display:inline-flex;
        align-items:center;
        gap:6px;
        padding:2px 8px;
        border-radius:999px;
        font-size:12px;
        margin-right:6px;
        background:rgba(0,0,0,.06);
      }
      #queue-monitor-panel .bell{
        cursor:pointer;
        border:none;
        background:rgba(255,255,255,.25);
        color:white;
        border-radius:6px;
        padding:0 8px;
        height:26px;
        display:flex;
        align-items:center;
        gap:6px;
      }
      #queue-monitor-panel .bell.muted{
        background:rgba(239,68,68,.85);
      }
      #queue-monitor-panel .snooze{
        position:absolute;
        top:48px;
        right:10px;
        background:#ffffff;
        color:#111;
        border:1px solid #e5e7eb;
        border-radius:8px;
        box-shadow:0 12px 28px rgba(0,0,0,.20);
        padding:8px;
        z-index:1000002;
        min-width:200px;
        display:none;
      }
      #queue-monitor-panel .snooze button{
        width:100%;
        border:none;
        padding:8px 10px;
        border-radius:6px;
        background:#f3f4f6;
        color:#111;
        font-weight:600;
        margin:6px 0;
        cursor:pointer;
      }
      #queue-monitor-panel .snooze button:hover{
        background:#e5e7eb;
      }
      #panel-header{ pointer-events:auto !important; }
      #qm-min{ pointer-events:auto !important; }
      #queue-monitor-panel{
        display:flex !important;
        flex-direction:column !important;
      }
      #queue-monitor-panel #panel-content{
        flex:1 1 auto !important;
        min-height:0 !important;
        height:auto !important;
        overflow-y:auto !important;
        -webkit-overflow-scrolling:touch !important;
      }
      #queue-monitor-panel.qm-hidden{
        display:none !important;
      }
      #queue-monitor-panel #panel-content{
        display:grid;
        grid-template-columns:repeat(auto-fit, minmax(500px, 1fr));
        gap:12px;
      }
      #queue-monitor-panel .qm-card{
        min-width:280px;
        max-width:600px;
        box-sizing:border-box;
      }
      .qm-card .timers{
        display:flex;
        align-items:center;
        gap:10px;
      }
      .qm-card.multi-chat .timers{
        flex-direction:column;
        align-items:flex-end;
        gap:4px;
      }
      #queue-monitor-panel .qm-card .agent-name,
      #queue-monitor-panel .qm-card .agent-name *{
        white-space:nowrap !important;
        overflow:visible !important;
        text-overflow:unset !important;
        display:inline-block !important;
        width:auto !important;
        max-width:none !important;
      }
    `;
    document.head.appendChild(st);
  }

  // ====== PANEL & MINIMIZE ======
  let panel = null;
  let minimizedBtn = null;
  let sectionsOpenRuntime = false;
  let selectActiveRuntime = false;
  let searchActiveRuntime = false;
  let snoozeOpenRuntime = false;
  let sectionsDocClickHandlerBound = false;

  function createMinimizedButton() {
    if (minimizedBtn) minimizedBtn.remove();
    minimizedBtn = document.createElement('button');
    minimizedBtn.id = 'queue-monitor-mini';
    minimizedBtn.textContent = 'QM';
    Object.assign(minimizedBtn.style, {
      position: 'fixed',
      bottom: '14px',
      right: '14px',
      padding: '8px 10px',
      borderRadius: '999px',
      border: '1px solid #d1d5db',
      background: '#ffffff',
      boxShadow: '0 6px 18px rgba(0,0,0,0.12)',
      cursor: 'pointer',
      zIndex: 1000000,
      display: minimized ? 'block' : 'none'
    });
    minimizedBtn.onclick = () => {
      minimized = false;
      save();
      if (panel) panel.classList.remove('qm-hidden');
      minimizedBtn.style.display = 'none';
      setTimeout(() => {
        const pc = panel && panel.querySelector('#panel-content');
        if (pc) {
          pc.style.overflowY = 'auto';
          pc.style.webkitOverflowScrolling = 'touch';
          pc.style.flex = '1 1 auto';
          pc.style.height = 'auto';
          pc.style.minHeight = '0';
        }
      }, 0);
    };
    document.body.appendChild(minimizedBtn);
  }

  function createPanel() {
    const existing = document.getElementById('queue-monitor-panel');
    if (existing) existing.remove();

    panel = document.createElement('div');
    panel.id = 'queue-monitor-panel';
    Object.assign(panel.style, {
      position: 'fixed',
      top: (panelPos && panelPos.top) || '10px',
      left: (panelPos && panelPos.left) || 'auto',
      right: (panelPos && panelPos.left) ? 'auto' : '10px',
      background: 'linear-gradient(135deg,#fff 0%,#f8f9fa 100%)',
      border: '1px solid #e1e5e9',
      borderRadius: '12px',
      padding: '0',
      fontFamily: 'system-ui, sans-serif',
      fontSize: '13px',
      zIndex: 999999,
      minWidth: CONFIG.PANEL_MIN_W + 'px',
      minHeight: CONFIG.PANEL_MIN_H + 'px',
      width: (panelSize && panelSize.w) ? panelSize.w + 'px' : '560px',
      height: (panelSize && panelSize.h) ? panelSize.h + 'px' : '700px',
      boxShadow: '0 8px 24px rgba(0,0,0,0.12)',
      resize: 'both',
      overflow: 'hidden',
      display: minimized ? 'none' : 'flex',
      flexDirection: 'column',
      isolation: 'isolate'
    });

    document.body.appendChild(panel);
    createMinimizedButton();

    let isDragging = false;
    let offsetX = 0;
    let offsetY = 0;

    function setupDrag(handle) {
      if (!handle) return;
      handle.style.cursor = 'move';
      handle.addEventListener('mousedown', e => {
        if (e.target.closest('input,select,button,[role="menu"],.qm-nodrag,#qm-min,#qm-bell,#qm-sections,#q-slots,#q-reset')) return;
        isDragging = true;
        offsetX = e.clientX - panel.offsetLeft;
        offsetY = e.clientY - panel.offsetTop;
        e.preventDefault();
      });
    }

    document.addEventListener('mouseup', () => { isDragging = false; });
    document.addEventListener('mousemove', e => {
      if (!isDragging) return;
      const x = e.clientX - offsetX;
      const y = e.clientY - offsetY;
      panel.style.left = x + 'px';
      panel.style.top = y + 'px';
      panel.style.right = 'auto';
      panelPos = { left: panel.style.left, top: panel.style.top };
      save();
    });

    panel._setupDragHandler = setupDrag;
    return panel;
  }

  // ====== HEADER & RENDERING ======
  function headerHtml(kpis) {
    const alertsCount = activeAlerts.total || 0;
    const bellCls =
      'bell qm-nodrag ' +
      ((muted || (Date.now() < (snoozeUntil || 0))) ? 'muted' : '');
    const snoozeTxt =
      (Date.now() < (snoozeUntil || 0))
        ? ' (⏱️ ' + Math.ceil((snoozeUntil - Date.now()) / 60000) + 'm)'
        : '';

    return (
      '<div id="panel-header" style="background:linear-gradient(90deg,#4f46e5,#7c3aed);' +
        'color:white;padding:10px 12px;border-radius:12px 12px 0 0;display:flex;' +
        'gap:8px;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:1;">' +
        '<div style="display:flex;align-items:center;gap:8px;min-width:0;">' +
          '<span>📊</span>' +
          '<div style="font-weight:700;white-space:nowrap;">Queue Monitor v2.2 (Widget)</div>' +
          '<div class="qm-nodrag" style="margin-left:8px;display:flex;gap:6px;align-items:center;">' +
            '<input id="qm-search" class="qm-nodrag" type="text" placeholder="Rechercher agent…" ' +
              'value="' + escapeHtml(searchFilter) + '" ' +
              'style="height:26px;border-radius:6px;border:none;padding:0 8px;outline:none;width:200px;color:#111;background:#fff;"/>' +
            '<button id="qm-sections" type="button" class="qm-nodrag" ' +
              'style="height:26px;border:none;border-radius:6px;padding:0 8px;' +
                     'background:rgba(255,255,255,.25);color:white;cursor:pointer;">Sections ▾</button>' +
            '<button id="q-slots" type="button" class="qm-nodrag" title="Liste des slots" ' +
              'style="height:26px;border:none;border-radius:6px;padding:0 8px;' +
                     'background:rgba(255,255,255,.25);color:white;cursor:pointer;">📋</button>' +
            '<button id="qm-min" type="button" class="qm-nodrag" title="Minimiser" ' +
              'style="height:26px;border:none;border-radius:6px;padding:0 8px;' +
                     'background:rgba(255,255,255,.25);color:white;cursor:pointer;">—</button>' +
          '</div>' +
          '<div class="qm-nodrag" style="display:flex;gap:6px;align-items:center;">' +
            '<label style="display:flex;align-items:center;gap:6px;font-size:12px;">' +
              '<span>Tri appel</span>' +
              '<select id="qm-sort-calls" class="qm-nodrag" ' +
                'style="height:26px;border-radius:6px;border:none;padding:0 6px;color:#111;background:#fff;">' +
                '<option value="">Aucun</option>' +
                '<option value="desc"' + (sortCallsOrder === 'desc' ? ' selected' : '') + '>Long → court</option>' +
                '<option value="asc"'  + (sortCallsOrder === 'asc'  ? ' selected' : '') + '>Court → long</option>' +
              '</select>' +
            '</label>' +
            '<label style="display:flex;align-items:center;gap:6px;font-size:12px;">' +
              '<span>Tri statut</span>' +
              '<select id="qm-sort-status" class="qm-nodrag" ' +
                'style="height:26px;border-radius:6px;border:none;padding:0 6px;color:#111;background:#fff;">' +
                '<option value="">Aucun</option>' +
                '<option value="desc"' + (sortStatusOrder === 'desc' ? ' selected' : '') + '>Long → court</option>' +
                '<option value="asc"'  + (sortStatusOrder === 'asc'  ? ' selected' : '') + '>Court → long</option>' +
              '</select>' +
            '</label>' +
            '<button id="q-reset" type="button" class="qm-nodrag" title="Réinitialiser" ' +
              'style="height:26px;border:none;border-radius:6px;padding:0 8px;' +
                     'background:rgba(255,255,255,.25);color:white;cursor:pointer;">🔄</button>' +
            '<button id="qm-bell" type="button" class="' + bellCls + '" title="Alertes' + snoozeTxt + '">' +
              (alertsCount ? '🔔 ' + alertsCount : '🔔') +
            '</button>' +
          '</div>' +
        '</div>' +
        '<div id="qm-sections-menu" role="menu" ' +
          'style="display:none;position:absolute;top:48px;left:10px;background:white;color:#111;' +
                 'border:1px solid #e5e7eb;border-radius:8px;box-shadow:0 8px 22px rgba(0,0,0,.12);' +
                 'padding:8px;z-index:1000000;min-width:240px;">' +
          SECTION_DEFS.map(s =>
            '<label style="display:flex;align-items:center;gap:8px;padding:4px 6px;cursor:pointer;">' +
              '<input data-section="' + s.key + '" type="checkbox" ' +
                (sectionVisibility[s.key] !== false ? 'checked' : '') + '/>' +
              '<span>' + s.label + '</span>' +
            '</label>'
          ).join('') +
        '</div>' +
        '<div id="qm-snooze" class="snooze">' +
          '<button data-snooze="0">Désactiver les alertes (toggle)</button>' +
          '<button data-snooze="10">Snooze 10 min</button>' +
          '<button data-snooze="30">Snooze 30 min</button>' +
          '<button data-snooze="60">Snooze 60 min</button>' +
        '</div>' +
      '</div>' +
      '<div id="panel-kpis" style="padding:6px 12px;display:flex;flex-wrap:wrap;gap:6px;' +
             'align-items:center;border-bottom:1px solid rgba(0,0,0,.05);">' +
        '<span class="kpi-chip">👥 ' + kpis.connected + ' agents</span>' +
        '<span class="kpi-chip">🟢 On Queue ' + kpis.onQueue + '</span>' +
        '<span class="kpi-chip">🔵 Longest call ' +
          (kpis.longestCall.name ? (kpis.longestCall.name + ' ' + toHMS(kpis.longestCall.ms)) : '—') +
        '</span>' +
        '<span class="kpi-chip">💬 Longest chat ' +
          (kpis.longestChat.name ? (kpis.longestChat.name + ' ' + toHMS(kpis.longestChat.ms)) : '—') +
        '</span>' +
        '<span class="kpi-chip">⛔ Prohibés ' + kpis.prohibCount + '</span>' +
        '<span class="kpi-chip">🟤 Occupés ' + kpis.occupeCount + '</span>' +
      '</div>'
    );
  }

  // Palette de couleurs par section (AJOUT DE "OCCUPÉ")
  const SECTION_STYLE = {
    default:         { bg:'#fefce8', border:'#f59e0b', dot:'🟡' },
    en_call:         { bg:'#dbeafe', border:'#3b82f6', dot:'🔵' },
    queue_free:      { bg:'#dcfce7', border:'#10b981', dot:'🟢' },
    interaction_hf:  { bg:'#fee2e2', border:'#ef4444', dot:'🔴' },
    disponible:      { bg:'#f3f4f6', border:'#9ca3af', dot:'⚪' },
    prohib:          { bg:'#fee2e2', border:'#dc2626', dot:'⛔' },
    en_chat:         { bg:'#ede9fe', border:'#7c3aed', dot:'🟣' },
    tache:           { bg:'#fff7ed', border:'#fb923c', dot:'🟠' },
    occupe:          { bg:'#ffedd5', border:'#f97316', dot:'🟤' }, // NOUVELLE COULEUR POUR OCCUPÉ
    non_telecontact: { bg:'#ffe4e6', border:'#f43f5e', dot:'⚫' },
    travaux:         { bg:'#f5f5f4', border:'#92400e', dot:'🟤' },
    pause:           { bg:'#f1f5f9', border:'#64748b', dot:'☕' },
    repas:           { bg:'#fff1f2', border:'#f43f5e', dot:'🍽️' },
    reunion:         { bg:'#ecfeff', border:'#06b6d4', dot:'📅' },
    formation:       { bg:'#fce7f3', border:'#db2777', dot:'📚' },
    autre:           { bg:'#fefce8', border:'#f59e0b', dot:'🟡' }
  };

  function renderAgentCard(agent, showSlot, statusKey) {
    const isFav = favorites.has(agent.name);
    const key = statusKey || deriveStatusKey(agent);
    const style = SECTION_STYLE[key] || SECTION_STYLE.default;
    const slot = SlotManager.getSlot(agent.name);
    const slotBadge = showSlot && slot
      ? '<span style="background:#6b7280;color:white;padding:1px 6px;border-radius:10px;font-size:10px;margin-left:4px;">S' + slot + '</span>'
      : '';
    const statusLabel = agent.status || agent.statusClass || '';

    const callMs = lastInCall[agent.name] && callStartAt[agent.name]
      ? msSince(callStartAt[agent.name])
      : 0;
    const statusMs = statusStartAt[agent.name]
      ? msSince(statusStartAt[agent.name])
      : 0;

    const chatStarts = Array.isArray(chatStartAt[agent.name]) ? chatStartAt[agent.name] : [];
    const chatTimers = chatStarts.filter(Boolean).length;
    const multiClass = chatTimers > 1 ? 'multi-chat' : '';

    const alertPulse =
      (callMs >= CONFIG.CALL_ALERT_SECS * 1000) ? 'qm-pulse' : '';

    let nameExtra = '';
    if (key === 'prohib') {
      const sub = prohibSubtypeFromStatus(agent.status || '') || 'Prohibé';
      nameExtra = '<span style="color:#b91c1c;font-weight:800;margin-left:6px;">(' + sub + ')</span>';
      lastProhibSubtype[agent.name] = sub;
    }

    const timersHtml = [];

    // Timer statut
    timersHtml.push(
      '<div class="qm-timer" data-type="status" data-name="' + escapeHtml(agent.name) + '" ' +
        'style="background:#eef2ff;padding:2px 8px;border-radius:12px;font-weight:600;' +
               'min-width:92px;text-align:center;' +
               (key === 'prohib' ? 'border:2px solid #dc2626;' : '') + '">' +
        '⏳ ' + toHMS(statusMs) +
      '</div>'
    );

    // Timer appel si actif
    if (lastInCall[agent.name] && callStartAt[agent.name]) {
      timersHtml.push(
        '<div class="qm-timer" data-type="call" data-name="' + escapeHtml(agent.name) + '" ' +
          'style="background:#e5e7eb;padding:2px 8px;border-radius:12px;font-weight:600;' +
                 'min-width:92px;text-align:center;">' +
          '📞 ' + toHMS(callMs) +
        '</div>'
      );
    }

    // Timers chats
    chatStarts.forEach((start, idx) => {
      if (!start) return;
      const ms = msSince(start);
      timersHtml.push(
        '<div class="qm-timer" data-type="chat" data-name="' + escapeHtml(agent.name) + '" data-idx="' + idx + '" ' +
          'style="background:#e6f4ff;padding:2px 8px;border-radius:12px;font-weight:600;' +
                 'min-width:92px;text-align:center;">' +
          '💬 ' + toHMS(ms) +
        '</div>'
      );
    });

    // Compteur activité
    const counterVal = agent.activityCount || 0;
    const counterHtml =
      '<div style="background:rgba(0,0,0,0.08);padding:2px 8px;border-radius:12px;' +
             'font-weight:bold;min-width:20px;text-align:center;">' +
        counterVal +
      '</div>';

    return (
      '<div class="qm-card ' + alertPulse + ' ' + multiClass + '" ' +
            'data-name="' + escapeHtml(agent.name) + '" ' +
            'title="' + escapeHtml(histPreview(agent.name)) + '" ' +
            'style="background:' + style.bg + ';margin:4px 0;padding:8px 12px;border-radius:8px;' +
                   'border-left:4px solid ' + style.border + ';display:flex;justify-content:space-between;' +
                   'align-items:center;gap:8px;">' +
        '<div style="display:flex;align-items:center;gap:8px;min-width:0;">' +
          '<button class="qm-fav" data-name="' + escapeHtml(agent.name) + '" ' +
                  'style="border:none;background:transparent;cursor:pointer;">' +
            (isFav ? '★' : '☆') +
          '</button>' +
          '<span>' + style.dot + '</span>' +
          '<span class="agent-name" style="' +
            (key === 'prohib' ? 'color:#b91c1c;font-weight:700;' : 'font-weight:600;') +
            'white-space:nowrap;max-width:200px;overflow:hidden;text-overflow:ellipsis;">' +
            escapeHtml(agent.name) +
          '</span>' +
          nameExtra +
          '<span style="color:#6b7280;font-size:12px;">' + escapeHtml(statusLabel) + '</span>' +
          slotBadge +
        '</div>' +
        '<div class="timers">' +
          timersHtml.join('') +
          counterHtml +
        '</div>' +
      '</div>'
    );
  }

  function renderSection(title, list, emoji, key, showSlot) {
    if (!list || !list.length || sectionVisibility[key] === false) return '';
    const favs = list.filter(a => favorites.has(a.name));
    const others = list.filter(a => !favorites.has(a.name));
    const final = favs.concat(others);
    return (
      '<div style="margin:12px 0;">' +
        '<div style="font-weight:700;color:#374151;margin-bottom:6px;display:flex;align-items:center;gap:6px;">' +
          '<span>' + emoji + '</span>' +
          '<span>' + title + '</span>' +
          '<span style="background:#f3f4f6;padding:2px 6px;border-radius:10px;font-size:11px;font-weight:normal;">' +
            final.length +
          '</span>' +
        '</div>' +
        final.map(a => renderAgentCard(a, showSlot, key)).join('') +
      '</div>'
    );
  }

  function renderSlotList() {
    const all = SlotManager.getAllBySlot();
    if (!all.length) return '';
    return (
      '<div style="margin:12px 0;padding:12px;background:#f8f9ff;' +
           'border:1px solid #e0e7ff;border-radius:8px;">' +
        '<div style="font-weight:700;color:#4338ca;margin-bottom:8px;display:flex;align-items:center;gap:6px;">' +
          '📋 <span>Liste des Slots (' + all.length + ')</span>' +
        '</div>' +
        '<div style="max-height:200px;overflow-y:auto;">' +
          all.map(it =>
            '<div style="display:flex;justify-content:space-between;align-items:center;' +
                        'padding:4px 8px;margin:2px 0;background:white;border-radius:4px;font-size:11px;">' +
              '<div><strong>Slot ' + it.slot + '</strong>: ' + escapeHtml(it.name) + '</div>' +
              '<div style="color:#6b7280;">' + (it.totalCalls || 0) + ' appels</div>' +
            '</div>'
          ).join('') +
        '</div>' +
      '</div>'
    );
  }

  // ====== PRÉSENCES RAPIDES ======
  let presenceOpen = false;
  let presenceInterval = null;

  function normalizeNameSimple(s) {
    return (s || '')
      .normalize('NFD')
      .replace(/\p{Diacritic}/gu, '')
      .toLowerCase()
      .trim();
  }

  function addPresenceButton() {
    const header = document.getElementById('panel-header');
    if (!header || document.getElementById('presence-btn')) return;
    const btn = document.createElement('button');
    btn.id = 'presence-btn';
    btn.textContent = '👥';
    btn.title = 'Présences rapides';
    btn.style.marginLeft = '8px';
    btn.style.border = 'none';
    btn.style.borderRadius = '6px';
    btn.style.padding = '0 8px';
    btn.style.height = '26px';
    btn.style.cursor = 'pointer';
    btn.style.background = 'rgba(255,255,255,.25)';
    btn.style.color = 'white';
    btn.onclick = togglePresenceTable;
    header.appendChild(btn);
  }

  function togglePresenceTable() {
    presenceOpen = !presenceOpen;
    const existing = document.getElementById('presence-table');
    if (presenceOpen) {
      if (!existing) renderPresenceTable();
    } else {
      if (existing) existing.remove();
      localStorage.removeItem('presenceInput');
      if (presenceInterval) {
        clearInterval(presenceInterval);
        presenceInterval = null;
      }
    }
  }

  function renderPresenceTable() {
    let div = document.getElementById('presence-table');
    if (div) return;

    div = document.createElement('div');
    div.id = 'presence-table';
    div.className = 'qm-card';
    Object.assign(div.style, {
      position: 'fixed',
      right: '20px',
      bottom: '120px',
      width: '320px',
      zIndex: 1000003,
      background: '#fff',
      border: '1px solid #ccc',
      padding: '10px',
      borderRadius: '8px',
      boxShadow: '0 2px 6px rgba(0,0,0,0.2)'
    });

    div.innerHTML =
      '<div id="presence-header" style="display:flex;justify-content:space-between;align-items:center;cursor:move;">' +
        '<h3 style="margin:0;">Présences rapides</h3>' +
        '<button id="presence-close" style="border:none;background:none;font-size:16px;cursor:pointer;">❌</button>' +
      '</div>' +
      '<textarea id="presence-input" rows="4" style="width:100%;resize:vertical;"></textarea>' +
      '<button id="presence-validate" style="margin-top:6px;">Valider</button>' +
      '<div id="presence-result" style="margin-top:10px;max-height:200px;overflow:auto;"></div>';

    document.body.appendChild(div);

    const closeBtn = document.getElementById('presence-close');
    closeBtn.onclick = () => {
      div.remove();
      presenceOpen = false;
      localStorage.removeItem('presenceInput');
      if (presenceInterval) {
        clearInterval(presenceInterval);
        presenceInterval = null;
      }
    };

    const input = document.getElementById('presence-input');
    input.value = localStorage.getItem('presenceInput') || '';
    input.addEventListener('input', e => {
      localStorage.setItem('presenceInput', e.target.value);
    });

    const stop = e => e.stopPropagation();
    [
      'keydown','keypress','keyup','input','compositionstart','compositionend',
      'paste','cut','copy','mousedown','click','wheel','contextmenu'
    ].forEach(evt => input.addEventListener(evt, stop, true));

    setTimeout(() => {
      input.focus();
      try {
        input.selectionStart = input.value.length;
        input.selectionEnd = input.value.length;
      } catch (e) {}
    }, 0);

    function getNamesFromTextarea() {
      return input.value
        .split(/[,\n]+/)
        .map(n => n.trim())
        .filter(n => n);
    }

    function updatePresenceResult(names) {
      const resultDiv = document.getElementById('presence-result');
      if (!resultDiv) return;

      const currentAgents = connectedUsers || [];
      let html = "<table style='width:100%;border-collapse:collapse;'><tr><th>Nom</th><th>Statut</th></tr>";
      names.forEach(name => {
        const target = currentAgents.find(
          aName => normalizeNameSimple(aName).includes(normalizeNameSimple(name))
        );
        const status = target
          ? "<span style='font-weight:bold;color:green;'>🟢 Connecté</span>"
          : "<span style='font-weight:bold;color:red;'>🔴 Absent</span>";
        html += "<tr><td style='padding:6px;'>" + escapeHtml(name) +
                "</td><td style='padding:6px;'>" + status + "</td></tr>";
      });
      html += "</table>";
      resultDiv.innerHTML = html;

      const savedPos = JSON.parse(localStorage.getItem('presencePos') || 'null');
      if (savedPos && div) {
        div.style.left = savedPos.left || div.style.left;
        div.style.top = savedPos.top || div.style.top;
        div.style.right = 'auto';
        div.style.bottom = 'auto';
      }
    }

    document.getElementById('presence-validate').onclick = () => {
      updatePresenceResult(getNamesFromTextarea());
    };

    updatePresenceResult(getNamesFromTextarea());

    if (presenceInterval) clearInterval(presenceInterval);
    presenceInterval = setInterval(() => {
      if (document.getElementById('presence-table')) {
        updatePresenceResult(getNamesFromTextarea());
      }
    }, 2000);

    input.addEventListener('input', () => {
      updatePresenceResult(getNamesFromTextarea());
    });

    const header = document.getElementById('presence-header');
    let offsetX = 0;
    let offsetY = 0;
    let isDown = false;

    header.addEventListener('mousedown', e => {
      isDown = true;
      offsetX = e.clientX - div.offsetLeft;
      offsetY = e.clientY - div.offsetTop;
      div.style.userSelect = 'none';
      e.preventDefault();
    });

    document.addEventListener('mouseup', () => {
      if (!isDown) return;
      isDown = false;
      div.style.userSelect = 'auto';
      try {
        localStorage.setItem(
          'presencePos',
          JSON.stringify({
            left: div.style.left || (div.offsetLeft + 'px'),
            top: div.style.top || (div.offsetTop + 'px')
          })
        );
      } catch (e) {}
    });

    document.addEventListener('mousemove', e => {
      if (!isDown) return;
      div.style.left = (e.clientX - offsetX) + 'px';
      div.style.top = (e.clientY - offsetY) + 'px';
      div.style.right = 'auto';
      div.style.bottom = 'auto';
    });
  }

  // ====== UPDATE UI PRINCIPALE ======
  function computeKpisFromGroups(groups) {
    const seen = new Set();
    const agents = [];
    Object.values(groups).forEach(arr => {
      (arr || []).forEach(a => {
        if (!seen.has(a.name)) {
          seen.add(a.name);
          agents.push(a);
        }
      });
    });
    return computeKpis(agents);
  }

  function updateUI(groups) {
    if (!panel || !document.body.contains(panel)) {
      panel = createPanel();
    }

    injectPulseStyle();
    injectUiFixStyles();
    injectThemeStyles();

    const kpis = computeKpisFromGroups(groups);
    const header = headerHtml(kpis);

    const connectedTotal = Object.values(groups)
      .reduce((acc, arr) => acc + (arr ? arr.length : 0), 0);

    const content =
      '<div id="panel-content" style="padding:12px;overflow-y:auto;flex:1 1 auto;height:auto;min-height:0;' +
        '-webkit-overflow-scrolling:touch;">' +
        renderSection('Favoris', groups.favoris || [], '⭐', 'favoris', true) +
        renderSection('Statut prohibé', groups.prohib || [], '⛔', 'prohib', true) +
        renderSection('Disponible', groups.disponible || [], '⚪', 'disponible', true) +
        renderSection("En file d'attente", groups.queue_free || [], '🟢', 'queue_free', true) +
        renderSection('En call', groups.en_call || [], '🔵', 'en_call', true) +
        renderSection('En chat', groups.en_chat || [], '💬', 'en_chat', true) +
        renderSection('Tâche associée', groups.tache || [], '📌', 'tache', true) +
        renderSection('Occupé', groups.occupe || [], '🟤', 'occupe', true) + // NOUVELLE SECTION
        renderSection('Non télécontact', groups.non_telecontact || [], '🚫', 'non_telecontact', true) +
        renderSection('Travaux payants', groups.travaux || [], '💼', 'travaux', true) +
        renderSection('Pause', groups.pause || [], '☕', 'pause', true) +
        renderSection('Repas', groups.repas || [], '🍽️', 'repas', true) +
        renderSection('Réunion', groups.reunion || [], '📅', 'reunion', true) +
        renderSection('Formation', groups.formation || [], '📚', 'formation', true) +
        renderSection('En interaction (hors file)', groups.interaction_hf || [], '🔴', 'interaction_hf', true) +
        renderSection('Autre', groups.autre || [], '🟡', 'autre', true) +
        (panel._showSlots ? renderSlotList() : '') +
        '<div style="margin-top:12px;padding:8px;background:#f8fafc;border-radius:6px;font-size:11px;color:#64748b;">' +
          '<span>Connectés: ' + connectedTotal + ' agents</span> • ' +
          '<span>Slots total: ' + masterSlotList.length + '</span> • ' +
          '<span>Maj: ' + new Date().toLocaleTimeString() + '</span>' +
        '</div>' +
      '</div>';

    const prev = panel.querySelector('#panel-content');
    const prevScroll = prev ? prev.scrollTop : 0;
    const sectionsWasOpen = sectionsOpenRuntime;

    panel.innerHTML = header + content;

    const cur = panel.querySelector('#panel-content');
    if (cur) cur.scrollTop = prevScroll;

    const headerEl = panel.querySelector('#panel-header');
    if (headerEl && panel._setupDragHandler) {
      panel._setupDragHandler(headerEl);
    }

    // Poignée de drag bas
    (function addBottomDragHandle() {
      const bh = document.createElement('div');
      bh.id = 'qm-bottom-drag';
      Object.assign(bh.style, {
        position: 'absolute',
        left: '50%',
        transform: 'translateX(-50%)',
        bottom: '0px',
        width: '140px',
        height: '14px',
        cursor: 'move',
        zIndex: 1000001,
        borderTop: '1px dashed rgba(0,0,0,.15)',
        background: 'linear-gradient(180deg, rgba(0,0,0,0.05), rgba(0,0,0,0.02))',
        borderRadius: '8px 8px 0 0'
      });
      bh.addEventListener('mousedown', e => e.preventDefault(), { capture: true });
      panel.appendChild(bh);
      if (panel._setupDragHandler) panel._setupDragHandler(bh);
    })();

    // Boutons et interactions
    wireHeaderInteractions(groups, sectionsWasOpen);
    addPresenceButton();
    if (presenceOpen) {
      renderPresenceTable();
    }
  }

  // Wire header (séparé pour clarté)
  function wireHeaderInteractions(groups, sectionsWasOpen) {
    const resetBtn = panel.querySelector('#q-reset');
    if (resetBtn) {
      resetBtn.onclick = doReset;
    }

    const slotsBtn = panel.querySelector('#q-slots');
    if (slotsBtn) {
      slotsBtn.onclick = () => {
        panel._showSlots = !panel._showSlots;
        slotsBtn.style.background = panel._showSlots
          ? 'rgba(255,255,255,0.4)'
          : 'rgba(255,255,255,0.25)';
        updateUI(groups);
      };
    }

    // Bell / Snooze
    const bell = panel.querySelector('#qm-bell');
    const snooze = panel.querySelector('#qm-snooze');
    if (snooze) {
      snooze.style.display = snoozeOpenRuntime ? 'block' : 'none';
    }

    if (bell) {
      bell.onclick = e => {
        e.stopPropagation();
        snoozeOpenRuntime = !snoozeOpenRuntime;
        if (snooze) snooze.style.display = snoozeOpenRuntime ? 'block' : 'none';
      };
    }

    if (snooze) {
      snooze.addEventListener('click', e => e.stopPropagation());
      snooze.querySelectorAll('button[data-snooze]').forEach(btn => {
        btn.onclick = e => {
          const mins = parseInt(e.currentTarget.getAttribute('data-snooze'), 10);
          if (mins === 0) {
            muted = !muted;
            snoozeUntil = 0;
          } else {
            muted = false;
            snoozeUntil = Date.now() + mins * 60000;
          }
          save();
          updateUI(groups);
        };
      });

      if (!window.__qmSnoozeDocHandlerBound) {
        window.__qmSnoozeDocHandlerBound = true;
        document.addEventListener('click', e => {
          if (!snoozeOpenRuntime) return;
          const snoozeEl = document.getElementById('qm-snooze');
          const bellEl = document.getElementById('qm-bell');
          if (snoozeEl &&
              !snoozeEl.contains(e.target) &&
              e.target !== bellEl) {
            snoozeOpenRuntime = false;
            snoozeEl.style.display = 'none';
          }
        });
      }
    }

    // Sections menu
    const sectionsBtn = panel.querySelector('#qm-sections');
    const sectionsMenu = panel.querySelector('#qm-sections-menu');
    if (sectionsBtn && sectionsMenu) {
      sectionsMenu.style.display = sectionsWasOpen ? 'block' : 'none';
      sectionsOpenRuntime = sectionsWasOpen;

      sectionsBtn.onclick = e => {
        e.stopPropagation();
        const open = sectionsMenu.style.display === 'none';
        sectionsMenu.style.display = open ? 'block' : 'none';
        sectionsOpenRuntime = open;
      };

      sectionsMenu.addEventListener('click', e => e.stopPropagation(), { passive: true });

      if (!sectionsDocClickHandlerBound) {
        sectionsDocClickHandlerBound = true;
        document.addEventListener('click', e => {
          const menuEl = document.getElementById('qm-sections-menu');
          const btnEl = document.getElementById('qm-sections');
          if (!menuEl || !btnEl) return;
          if (menuEl.style.display === 'block' &&
              !menuEl.contains(e.target) &&
              e.target !== btnEl) {
            menuEl.style.display = 'none';
            sectionsOpenRuntime = false;
          }
        });
      }

      sectionsMenu.querySelectorAll('input[data-section]').forEach(cb => {
        cb.onchange = () => {
          const key = cb.getAttribute('data-section');
          sectionVisibility[key] = cb.checked;
          save();
          sectionsOpenRuntime = true;
          updateUI(groups);
        };
      });
    }

    // Search
    const searchEl = panel.querySelector('#qm-search');
    if (searchEl) {
      searchEl.onfocus = () => { searchActiveRuntime = true; };
      searchEl.onblur = () => { searchActiveRuntime = false; };
      let searchDebounceTimer = null;
      searchEl.oninput = () => {
        clearTimeout(searchDebounceTimer);
        searchDebounceTimer = setTimeout(() => {
          searchFilter = searchEl.value.trim();
          save();
        }, 200);
      };
      searchEl.onkeydown = e => {
        if (e.key === 'Enter') {
          searchFilter = searchEl.value.trim();
          save();
        }
      };
      searchEl.addEventListener('pointerdown', primeAudio, { once: true });
      searchEl.addEventListener('pointerdown', requestNotifyPermission, { once: true });
    }

    // Minimiser
    const minBtn = panel.querySelector('#qm-min');
    if (minBtn) {
      const minimize = e => {
        e.preventDefault();
        e.stopPropagation();
        if (minBtn.dataset._done === '1') return;
        minBtn.dataset._done = '1';
        minimized = true;
        snoozeOpenRuntime = false;
        save();
        panel.classList.add('qm-hidden');
        if (minimizedBtn) minimizedBtn.style.display = 'block';
        setTimeout(() => { minBtn.dataset._done = ''; }, 300);
      };
      ['pointerdown', 'mousedown', 'click', 'pointerup'].forEach(ev => {
        minBtn.addEventListener(ev, minimize, { capture: true, passive: false });
      });
    }

    // Tri sélecteurs
    const callsSel = panel.querySelector('#qm-sort-calls');
    if (callsSel) {
      callsSel.onmousedown = () => { selectActiveRuntime = true; };
      callsSel.onchange = () => {
        sortCallsOrder = callsSel.value;
        save();
      };
      callsSel.onblur = () => { selectActiveRuntime = false; };
    }

    const statusSel = panel.querySelector('#qm-sort-status');
    if (statusSel) {
      statusSel.onmousedown = () => { selectActiveRuntime = true; };
      statusSel.onchange = () => {
        sortStatusOrder = statusSel.value;
        save();
      };
      statusSel.onblur = () => { selectActiveRuntime = false; };
    }

    // Favoris
    panel.querySelectorAll('button.qm-fav').forEach(btn => {
      btn.onclick = e => {
        const name = e.currentTarget.getAttribute('data-name');
        if (favorites.has(name)) favorites.delete(name);
        else favorites.add(name);
        Storage.set('favorites', [...favorites]);
        updateUI(groups);
      };
    });
  }

  // ====== UI BUSY / MENUS ======
  function isQmUiBusy() {
    return sectionsOpenRuntime || selectActiveRuntime || searchActiveRuntime;
  }

  function isMenuOpen() {
    const sels = [
      '.menu.show-dropdown',
      '.dropdown.show',
      '.user-menu[aria-expanded="true"]',
      '.presence-sidebar[style*="display: block"]',
      '[class*="menu"][style*="display: block"]',
      '.show-dropdown',
      '[aria-expanded="true"]:not(button[role="radio"])',
      '.menus[style*="display: block"]'
    ];
    return sels.some(sel =>
      Array.from(document.querySelectorAll(sel)).some(el => {
        const s = getComputedStyle(el);
        return s.display !== 'none' && s.visibility !== 'hidden' && s.opacity !== '0';
      })
    );
  }

  // ====== RESET COMPLET ======
  function doReset() {
    if (!confirm('Réinitialiser toutes les données (slots, timers, favoris, préférences) ?')) return;

    masterSlotList = [];
    connectedUsers = [];
    lastGroup = {};
    slotCounter = 1;
    lastStatusKey = {};
    statusStartAt = {};
    callStartAt = {};
    chatStartAt = {};
    lastInCall = {};
    callOffStreak = {};
    lastInChat = {};
    chatOffStreak = {};
    favorites = new Set();
    lastProhibSubtype = {};
    sectionVisibility = Object.fromEntries(SECTION_DEFS.map(s => [s.key, true]));
    searchFilter = '';
    minimized = false;
    panelPos = null;
    panelSize = null;
    sortCallsOrder = 'desc';
    sortStatusOrder = '';
    muted = false;
    snoozeUntil = 0;
    historyByDay = {};
    dailyAgg = {};
    callAlerted = {};
    chatAlerted = {};
    prohibAlerted = {};
    activeAlerts = { total: 0 };

    Storage.clear();
    save();
    if (panel) panel.remove();
    panel = createPanel();
    processAgents();
  }

  // ====== DEBOUNCE PROCESS ======
  let menuDebounceTimer = null;

  function debouncedProcessAgents() {
    clearTimeout(menuDebounceTimer);
    menuDebounceTimer = setTimeout(() => {
      if (!isMenuOpen() && !isQmUiBusy()) {
        processAgents();
      } else {
        debouncedProcessAgents();
      }
    }, CONFIG.MENU_DEBOUNCE);
  }

  // ====== COEUR: PROCESS AGENTS (WIDGET) ======
  function processAgents() {
    try {
      if (isMenuOpen() || isQmUiBusy()) {
        debouncedProcessAgents();
        return;
      }

      const agentsRaw = findAgentsInWidget();
      let agents = agentsRaw;

      if (searchFilter) {
        const f = norm(searchFilter);
        agents = agents.filter(a => norm(a.name).includes(f));
      }

      // CORRECTION : ConnectedUsersManager.updateConnectedUsers() filtre déjà les "Hors ligne"
      ConnectedUsersManager.updateConnectedUsers(agents);

      agents.forEach(agent => {
        const key = deriveStatusKey(agent);
        ensureStatusTimerOnly(key, agent.name);
        
        // CORRECTION : Arrêt des timers d'appel si on passe en AT ou Occupé
        if ((key === 'tache' || key === 'occupe') && lastInCall[agent.name]) {
          // Si l'agent passe en AT ou Occupé alors qu'il était en call, on arrête le timer d'appel
          const dur = msSince(callStartAt[agent.name] || nowIso());
          aggAddMs(agent.name, 'callMs', dur);
          histAdd(agent.name, 'call_end', { durMs: dur });
          delete callStartAt[agent.name];
          lastInCall[agent.name] = false;
          callOffStreak[agent.name] = 0;
          log.debug('Appel arrêté pour', agent.name, 'passage en ' + key);
        }
        
        // Aligne le timer de statut
        const stSec = extractStatusDurationFromRow(agent.rowElement);
        if (stSec != null) {
          const desiredMs = stSec * 1000;
          const curIso = statusStartAt[agent.name];
          const drift = curIso ? desiredMs - msSince(curIso) : Infinity;
          if (!curIso || Math.abs(drift) > 5000) {
            statusStartAt[agent.name] = new Date(Date.now() - desiredMs).toISOString();
          }
        }

        // 1) Appels - UNIQUEMENT si ce n'est pas une AT ou Occupé
        let inCall = false;
        let genesysSec = null;
        
        // Seulement pour les statuts qui peuvent vraiment avoir des appels
        // ET qui ne sont pas en AT ou Occupé
        if (key !== 'tache' && key !== 'occupe' && key !== 'prohib') {
          const isInteracting = agent.statusClass === 'On Call' || 
                               agent.status.includes('En cours d\'interaction') ||
                               agent.status.includes('En cours de communication');
          
          if (isInteracting) {
            inCall = true;
            genesysSec = extractCallDurationFromRow(agent.rowElement);
          }
        }

        updateCallState(agent.name, inCall, genesysSec);

        // 2) Chats
        let chatCount = 0;
        let chatDurSecs = [];
        if (key === 'en_chat') {
          chatDurSecs = extractChatDurationsFromAgentRow(agent.rowElement) || [];
          if (chatDurSecs.length) {
            chatCount = Math.min(2, chatDurSecs.length);
          }
        }
        updateChatState(agent.name, chatCount);

        if (chatDurSecs.length) {
          if (!Array.isArray(chatStartAt[agent.name])) chatStartAt[agent.name] = [];
          chatDurSecs.slice(0, chatCount).forEach((sec, i) => {
            const desiredMs = sec * 1000;
            const curIso = chatStartAt[agent.name][i];
            const diff = curIso ? desiredMs - msSince(curIso) : Infinity;
            if (!curIso || Math.abs(diff) > 5000) {
              chatStartAt[agent.name][i] = new Date(Date.now() - desiredMs).toISOString();
            }
            lastInChat[agent.name][i] = true;
            chatOffStreak[agent.name][i] = 0;
          });
        }

        lastGroup[agent.name] = key;
      });

      save();

      // Construction des groupes par slot
      const bySlot = ConnectedUsersManager.getConnectedBySlot();
      const map = new Map(agents.map(a => [a.name, a]));

      const groups = {
        favoris: [],
        queue_free: [],
        en_call: [],
        en_chat: [],
        tache: [],
        occupe: [], // NOUVELLE SECTION
        non_telecontact: [],
        travaux: [],
        pause: [],
        repas: [],
        reunion: [],
        formation: [],
        interaction_hf: [],
        disponible: [],
        autre: [],
        prohib: []
      };

      bySlot.forEach(({ name }) => {
        const agent = map.get(name);
        if (!agent) return;
        const key = deriveStatusKey(agent) || 'autre';
        const target = groups[key] || groups.autre;
        target.push(agent);
        if (favorites.has(name)) {
          groups.favoris.push(agent);
        }
      });

      // Tri des groupes
      Object.keys(groups).forEach(k => {
        const arr = groups[k];
        if (!arr || !arr.length) return;

        // 1) En call : le tri par durée d'appel prime si activé
        if (k === 'en_call' && sortCallsOrder) {
          arr.sort((a, b) => {
            const da = callStartAt[a.name] ? msSince(callStartAt[a.name]) : 0;
            const db = callStartAt[b.name] ? msSince(callStartAt[b.name]) : 0;
            const cmp = (sortCallsOrder === 'desc') ? (db - da) : (da - db);
            if (cmp !== 0) return cmp;
            // tie-breaker stable par slot
            const sa = SlotManager.getSlot(a.name) || 999;
            const sb = SlotManager.getSlot(b.name) || 999;
            return sa - sb;
          });
          return;
        }

        // 2) Sinon, applique le tri statut si demandé…
        if (sortStatusOrder) {
          arr.sort((a, b) => {
            const da = statusStartAt[a.name] ? msSince(statusStartAt[a.name]) : 0;
            const db = statusStartAt[b.name] ? msSince(statusStartAt[b.name]) : 0;
            const cmp = (sortStatusOrder === 'desc') ? (db - da) : (da - db);
            if (cmp !== 0) return cmp;
            const sa = SlotManager.getSlot(a.name) || 999;
            const sb = SlotManager.getSlot(b.name) || 999;
            return sa - sb;
          });
          return;
        }

        // 3) …sinon tri par slot par défaut
        arr.sort((a, b) => {
          const sa = SlotManager.getSlot(a.name) || 999;
          const sb = SlotManager.getSlot(b.name) || 999;
          return sa - sb;
        });
      });

      // ====== ALERTES ======
      activeAlerts = { total: 0 };
      const seen = new Set();
      agents.forEach(a => {
        if (seen.has(a.name)) return;
        seen.add(a.name);
        const key = deriveStatusKey(a);

        // Prohibé
        if (key === 'prohib') {
          activeAlerts.total++;
          if (!prohibAlerted[a.name]) {
            prohibAlerted[a.name] = true;
            beep(520, 200);
            notify('Statut prohibé', a.name + ' en ' + (prohibSubtypeFromStatus(a.status || '') || 'prohibé'));
          }
        } else {
          prohibAlerted[a.name] = false;
        }

        // Appel long
        if (lastInCall[a.name] && callStartAt[a.name]) {
          const ms = msSince(callStartAt[a.name]);
          if (ms >= CONFIG.CALL_ALERT_SECS * 1000) {
            activeAlerts.total++;
            if (!callAlerted[a.name]) {
              callAlerted[a.name] = true;
              beep(880, 200);
              notify('Appel long', a.name + ' 📞 ' + toHMS(ms));
            }
          } else {
            callAlerted[a.name] = false;
          }
        } else {
          callAlerted[a.name] = false;
        }

        // Chat long (1er chat actif)
        if (lastInChat[a.name] && Array.isArray(chatStartAt[a.name]) && chatStartAt[a.name][0]) {
          const ms = msSince(chatStartAt[a.name][0]);
          if (ms >= CONFIG.CHAT_ALERT_SECS * 1000) {
            activeAlerts.total++;
            if (!chatAlerted[a.name]) {
              chatAlerted[a.name] = true;
              beep(700, 200);
              notify('Chat long', a.name + ' 💬 ' + toHMS(ms));
            }
          } else {
            chatAlerted[a.name] = false;
          }
        } else {
          chatAlerted[a.name] = false;
        }
      });

      updateUI(groups);
    } catch (e) {
      console.error('[QM ERROR] processAgents', e);
      if (panel) {
        panel.innerHTML =
          '<div style="background:linear-gradient(90deg,#ef4444,#dc2626);color:white;' +
          'padding:12px 16px;border-radius:12px;">Erreur Queue Monitor – voir console</div>';
      }
    }
  }

  // ====== OBSERVERS ======
  function setupObservers() {
    try {
      const obs = new MutationObserver(muts => {
        let need = false;
        for (const m of muts) {
          if (m.type === 'childList') {
            if ([...m.addedNodes, ...m.removedNodes].some(
              n => n.nodeType === 1 && (n.matches('.widget-type-AGENT_STATUS') || 
                    n.querySelector?.('.widget-type-AGENT_STATUS') ||
                    n.matches('tbody tr[data-row-id]') ||
                    n.querySelector?.('tbody tr[data-row-id]'))
            )) {
              need = true;
              break;
            }
          }
          if (m.type === 'attributes') {
            const t = m.target;
            if (t.matches &&
                (t.matches('.widget-type-AGENT_STATUS') ||
                 t.matches('tbody tr[data-row-id]') ||
                 t.matches('.entity-v3-presence-indicator-dot') ||
                 t.matches('.additional-label') ||
                 t.matches('.time-in-status'))) {
              need = true;
              break;
            }
          }
        }
        if (need && !isMenuOpen() && !isQmUiBusy()) {
          setTimeout(processAgents, 150);
        }
      });

      obs.observe(document.body, {
        childList: true,
        subtree: true,
        attributes: true,
        attributeFilter: ['class', 'style', 'aria-label', 'aria-expanded', 'data-state']
      });
    } catch (e) {
      console.error('[QM ERROR] setupObservers', e);
    }
  }

  // ====== VISUAL TICK (timers UI) ======
  function visualTick() {
    try {
      document.querySelectorAll('#queue-monitor-panel .qm-timer').forEach(el => {
        const name = el.getAttribute('data-name');
        const type = el.getAttribute('data-type');
        const idxAttr = el.getAttribute('data-idx');
        const idx = idxAttr != null ? parseInt(idxAttr, 10) : null;

        let ms = 0;
        let icon = '⏳';

        if (!name) return;

        if (type === 'call' && callStartAt[name]) {
          ms = msSince(callStartAt[name]);
          icon = '📞';
        } else if (type === 'chat' && Array.isArray(chatStartAt[name]) && idx != null && chatStartAt[name][idx]) {
          ms = msSince(chatStartAt[name][idx]);
          icon = '💬';
        } else if (type === 'status' && statusStartAt[name]) {
          ms = msSince(statusStartAt[name]);
          icon = '⏳';
        } else {
          return;
        }

        el.textContent = icon + ' ' + toHMS(ms);

        if (type === 'call') {
          const card = el.closest('.qm-card');
          if (card) {
            if (ms >= CONFIG.CALL_ALERT_SECS * 1000) card.classList.add('qm-pulse');
            else card.classList.remove('qm-pulse');
          }
        }
      });
    } catch (e) {
      console.error('[QM ERROR] visualTick', e);
    }
  }

  // ====== HOTKEYS & WATCHDOG ======
  function hotkeys(e) {
    if (!(e.shiftKey && e.altKey)) return;
    if (e.code === 'KeyQ') {
      minimized = false;
      panelPos = null;
      panelSize = null;
      save();
      if (panel) panel.remove();
      createPanel();
      processAgents();
      if (minimizedBtn) minimizedBtn.style.display = 'none';
    } else if (e.code === 'KeyR') {
      if (confirm('Queue Monitor – Reset complet ?')) {
        Storage.clear();
        location.reload();
      }
    }
  }

  function watchdog() {
    setTimeout(() => {
      if (!document.getElementById('queue-monitor-panel')) {
        try {
          createPanel();
          processAgents();
        } catch (e) {
          console.error('[QM ERROR] Watchdog', e);
        }
      }
    }, 1500);
  }

  // ====== LOOP LOGIQUE AUTO-ADAPTATIVE ======
  let logicTimer = null;

  function scheduleLogicLoop() {
    clearTimeout(logicTimer);
    logicTimer = setTimeout(() => {
      if (!document.hidden && !isMenuOpen() && !isQmUiBusy()) {
        processAgents();
      }
      scheduleLogicLoop();
    }, effectiveRefreshMs());
  }

  // ====== INIT ======
  function init() {
    const host = window.location.hostname;

    if (!/mypurecloud|pure\.cloud/i.test(host)) {
      console.error('[QM ERROR] Not on Genesys Cloud host:', host);
      return;
    }

    log.info('[QM] init (widget) on', host);

    panel = createPanel();
    injectPulseStyle();
    injectUiFixStyles();
    injectThemeStyles();

    setupObservers();
    scheduleLogicLoop();
    setInterval(visualTick, CONFIG.TICK_INTERVAL);

    setTimeout(() => {
      if (!isMenuOpen() && !isQmUiBusy()) processAgents();
    }, 800);

    window.QueueMonitor = {
      refresh: () => {
        if (!isMenuOpen() && !isQmUiBusy()) processAgents();
        else debouncedProcessAgents();
      },
      debug: () => ({
        masterSlotList,
        connectedUsers,
        lastGroup,
        slotCounter,
        lastStatusKey,
        statusStartAt,
        callStartAt,
        chatStartAt,
        lastInCall,
        lastInChat,
        callOffStreak,
        chatOffStreak,
        favorites: [...favorites],
        sectionVisibility,
        searchFilter,
        sortCallsOrder,
        sortStatusOrder,
        muted,
        snoozeUntil,
        dailyAgg,
        historyByDay
      })
    };

    const ro = new ResizeObserver(entries => {
      for (const e of entries) {
        const r = e.contentRect;
        panelSize = {
          w: Math.max(r.width, CONFIG.PANEL_MIN_W),
          h: Math.max(r.height, CONFIG.PANEL_MIN_H)
        };
        save();
      }
    });
    ro.observe(panel);

    if (minimized) {
      panel.classList.add('qm-hidden');
      if (minimizedBtn) minimizedBtn.style.display = 'block';
    }

    document.addEventListener('keydown', hotkeys);
    requestNotifyPermission();
    panel.addEventListener('pointerdown', primeAudio, { once: true });

    watchdog();
    log.info('[QM] running (widget, slots + timers + alertes actifs)');
  }

  // ====== INIT LANCEMENT ======
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    setTimeout(init, 300);
  }

})();
