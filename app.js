/* =====================================================================
   AYAEL ENTREPRISE — Application de gestion (back-office)
   Une seule page. Les données sont dans Supabase (version d'équipe)
   ou dans une base locale du navigateur (mode démo).
   ===================================================================== */
'use strict';
(() => {
  const PGLITE_URL = 'https://cdn.jsdelivr.net/npm/@electric-sql/pglite@0.2.17/dist/index.js';
  const XLSX_URL = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
  const DEMO_SCHEMA_VERSION = '10';
  const LOGIN_DOMAIN = 'ayael.app';

  const $app = document.getElementById('app');
  const $sheets = document.getElementById('sheet-root');
  const $toasts = document.getElementById('toast-root');

  const S = {
    api: null, me: null, accounts: [], products: [], team: [], settings: {},
    bal: new Map(), ops: new Map(), cards: new Map(), clients: new Map(),
    view: null, sheetOpen: 0, pendingRefresh: false, subscribed: false,
    rc: { gtpId: Number(store('rc-gtp')) || null, q: '' },
    cl: { q: '' },
    jf: { period: 'today', from: null, to: null, kind: '', user: '', account: '', q: '' },
    rp: { period: 'month', from: null, to: null },
    cf: { q: '', bank: '' },
  };

  // ---------------------------------------------------------------- utilitaires
  function store(k, v) {
    try {
      if (v === undefined) return localStorage.getItem('ayael-' + k);
      if (v === null) localStorage.removeItem('ayael-' + k); else localStorage.setItem('ayael-' + k, v);
    } catch (e) { return null; }
  }
  const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const nf = new Intl.NumberFormat('fr-FR', { maximumFractionDigits: 0 });
  const fmtN = n => nf.format(Math.round(Number(n) || 0)).replace(/[  ]/g, ' ');
  const fmt = n => fmtN(n) + ' F';
  const fmtSigned = n => (n > 0 ? '+' : n < 0 ? '−' : '') + fmt(Math.abs(n));
  function parseAmt(v) {
    const s = String(v ?? '').trim();
    const neg = /^[-−]/.test(s);
    const d = s.replace(/\D/g, '');
    return d === '' ? null : (neg ? -1 : 1) * Number(d);
  }
  const isoDay = d => { const z = new Date(d.getTime() - d.getTimezoneOffset() * 60000); return z.toISOString().slice(0, 10); };
  const todayIso = () => isoDay(new Date());
  const fmtDate = s => { if (!s) return ''; const [y, m, d] = String(s).slice(0, 10).split('-'); return `${d}/${m}/${y}`; };
  const fmtTime = iso => iso ? new Date(iso).toLocaleTimeString('fr-FR', { hour: '2-digit', minute: '2-digit' }) : '';
  const fmtDT = iso => iso ? fmtDate(isoDay(new Date(iso))) + ' à ' + fmtTime(iso) : '';
  function ago(iso) {
    if (!iso) return '';
    const s = (Date.now() - new Date(iso).getTime()) / 1000;
    if (s < 60) return 'à l’instant';
    if (s < 3600) return `il y a ${Math.floor(s / 60)} min`;
    if (s < 86400) return `il y a ${Math.floor(s / 3600)} h`;
    if (s < 172800) return 'hier';
    return `il y a ${Math.floor(s / 86400)} j`;
  }
  const debounce = (f, ms) => { let t; return (...a) => { clearTimeout(t); t = setTimeout(() => f(...a), ms); }; };
  const isDir = () => S.me && S.me.role === 'gerante';                 // direction (gérante de l'agence principale)
  const isMgr = () => S.me && S.me.role !== 'agent';                   // direction ou gérant d'agence
  const roleLabel = r => ({ gerante: 'Direction', gerant: 'Gérant(e) d’agence', agent: 'Agent' }[r] || r);
  const acc = id => S.accounts.find(a => a.id === Number(id));
  const activeCards = c => ((c && c.cards) || []).filter(k => k.active !== false);
  // Compte utilisable par moi : de mon agence, ou commun à toutes les agences.
  const myAcc = a => a.agence_id == null || a.agence_id === S.me.agence_id;
  // Compte électronique à proposer pour un produit (ex. « Dépôt Mixx » à Noépé → le Mixx de Noépé).
  function counterFor(p, elecs) {
    if (!p.counter_account_id) return null;
    const own = elecs.find(a => a.id === p.counter_account_id);
    if (own) return own.id;
    const words = norm(p.name).split(/\s+/).filter(w => w.length > 2 && !/^(DEPOT|RETRAIT|ENVOI|VENTE|ACHAT)$/.test(w));
    const match = elecs.filter(a => words.some(w => norm(a.name).includes(w)));
    return (match.find(a => a.agence_id === S.me.agence_id) || match[0] || {}).id || null;
  }
  const initials = n => String(n || '?').replace(/\(.*?\)/g, ' ').split(/[^A-Za-zÀ-ÿ0-9]+/).filter(Boolean).slice(0, 2).map(w => w[0]).join('').toUpperCase() || '?';
  // Pour la direction, les écrans de suivi montrent l'agence choisie en haut (vide = toutes).
  const AG_RPC = new Set(['dashboard', 'list_operations', 'list_closings', 'report']);
  const rpc = (n, p) => {
    p = p || {};
    if (AG_RPC.has(n) && isDir() && S.agence && p.agence_id === undefined) p = Object.assign({ agence_id: S.agence }, p);
    return S.api.rpc(n, p);
  };
  function period(key) {
    const d = new Date(), t = isoDay(d);
    switch (key) {
      case 'today': return [t, t];
      case 'yesterday': { const y = new Date(d); y.setDate(d.getDate() - 1); const s = isoDay(y); return [s, s]; }
      case '7d': { const y = new Date(d); y.setDate(d.getDate() - 6); return [isoDay(y), t]; }
      case 'month': return [t.slice(0, 8) + '01', t];
      case 'lastmonth': return [isoDay(new Date(d.getFullYear(), d.getMonth() - 1, 1)), isoDay(new Date(d.getFullYear(), d.getMonth(), 0))];
      case 'year': return [t.slice(0, 4) + '-01-01', t];
      default: return [t, t];
    }
  }
  const PERIODS = [['today', 'Aujourd’hui'], ['yesterday', 'Hier'], ['7d', '7 jours'], ['month', 'Ce mois'], ['lastmonth', 'Mois dernier'], ['year', 'Cette année'], ['custom', 'Dates…']];
  const KIND = { caisse: 'Caisse', recharge: 'Recharge', approvisionnement: 'Approvisionnement', depense: 'Dépense', ajustement: 'Ajustement', transfert: 'Transfert', cloture: 'Remise de caisse' };
  const ACCOUNT_KIND = { gtp: 'Compte GTP', electronique: 'Compte électronique', caisse: 'Caisse d’agent', coffre: 'Coffre', banque: 'Banque' };
  const DEP_CATS = ['Transport', 'Repas', 'Fournitures', 'Électricité / eau', 'Loyer', 'Salaire / avance', 'Autre'];

  function commission(a, amt) {
    if (!a || !Number(a.comm_rate) || !amt || amt <= 0) return 0;
    let v = amt * Number(a.comm_rate) / 100 / (Number(a.comm_divisor) || 1);
    v = a.comm_rounding === 'ceil' ? Math.ceil(v) : a.comm_rounding === 'floor' ? Math.floor(v) : Math.round(v);
    return Math.max(v, Number(a.comm_min) || 0);
  }
  function commRule(a) {
    if (!a || !Number(a.comm_rate)) return 'Pas de commission automatique';
    const r = String(a.comm_rate).replace('.', ',');
    const dv = String(a.comm_divisor).replace('.', ',');
    return `${r} % ÷ ${dv}${Number(a.comm_min) ? `, minimum ${fmt(a.comm_min)}` : ''}`;
  }

  // ---------------------------------------------------------------- icônes
  const I = {
    home: '<path d="M3 10.5 12 3l9 7.5V20a1 1 0 0 1-1 1h-5v-6H9v6H4a1 1 0 0 1-1-1z"/>',
    cash: '<rect x="2" y="6" width="20" height="12" rx="2"/><circle cx="12" cy="12" r="2.5"/><path d="M6 12h.01M18 12h.01"/>',
    card: '<rect x="2" y="5" width="20" height="14" rx="2"/><path d="M2 10h20M6 15h4"/>',
    list: '<path d="M8 6h13M8 12h13M8 18h13M3 6h.01M3 12h.01M3 18h.01"/>',
    lock: '<rect x="4" y="11" width="16" height="10" rx="2"/><path d="M8 11V7a4 4 0 0 1 8 0v4"/>',
    plus: '<path d="M12 5v14M5 12h14"/>',
    grid: '<rect x="3" y="3" width="7" height="7" rx="1.5"/><rect x="14" y="3" width="7" height="7" rx="1.5"/><rect x="3" y="14" width="7" height="7" rx="1.5"/><rect x="14" y="14" width="7" height="7" rx="1.5"/>',
    search: '<circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/>',
    in: '<path d="M12 19V5M5 12l7 7 7-7"/>',
    out: '<path d="M12 5v14M5 12l7-7 7 7"/>',
    swap: '<path d="M7 7h13l-4-4M17 17H4l4 4"/>',
    adjust: '<path d="M4 6h10M18 6h2M4 12h4M12 12h8M4 18h12M20 18h0"/><circle cx="16" cy="6" r="2"/><circle cx="10" cy="12" r="2"/><circle cx="18" cy="18" r="2"/>',
    truck: '<path d="M3 7h11v9H3zM14 10h4l3 3v3h-7"/><circle cx="7" cy="18" r="1.8"/><circle cx="17" cy="18" r="1.8"/>',
    chart: '<path d="M4 20V10M10 20V4M16 20v-7M22 20H2"/>',
    users: '<circle cx="9" cy="8" r="3.5"/><path d="M2 20c0-3.5 3-6 7-6s7 2.5 7 6"/><path d="M16 4.5a3.5 3.5 0 0 1 0 7M18 14c2.5.6 4 2.8 4 6"/>',
    bank: '<path d="M3 10 12 4l9 6M5 10v8M9.5 10v8M14.5 10v8M19 10v8M3 20h18"/>',
    tag: '<path d="M3 12V4h8l10 10-8 8z"/><circle cx="7.5" cy="8.5" r="1.5"/>',
    upload: '<path d="M12 16V4M7 9l5-5 5 5M4 20h16"/>',
    download: '<path d="M12 4v12M7 11l5 5 5-5M4 20h16"/>',
    user: '<circle cx="12" cy="8" r="4"/><path d="M4 21c0-4 3.5-7 8-7s8 3 8 7"/>',
    print: '<path d="M6 9V3h12v6M6 18H4v-7h16v7h-2M7 14h10v7H7z"/>',
    chev: '<path d="m9 6 6 6-6 6"/>',
    receipt: '<path d="M5 3h14v18l-3-2-2 2-2-2-2 2-2-2-3 2z"/><path d="M9 8h6M9 12h6M9 16h3"/>',
    chat: '<path d="M21 12a8.5 8.5 0 0 1-12.6 7.4L3 21l1.6-5.2A8.5 8.5 0 1 1 21 12z"/>',
  };
  const icon = n => `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${I[n] || ''}</svg>`;

  // ---------------------------------------------------------------- toasts, feuilles, confirmations
  function toast(msg, type = '') {
    const el = document.createElement('div');
    el.className = 'toast ' + type;
    el.setAttribute('role', 'status');
    el.textContent = msg;
    $toasts.append(el);
    setTimeout(() => el.remove(), type === 'bad' ? 6000 : 3200);
  }

  function bindMoney(root) {
    root.querySelectorAll('input.money').forEach(inp => {
      if (inp.dataset.bound) return;
      inp.dataset.bound = '1';
      inp.addEventListener('input', () => {
        const n = parseAmt(inp.value);
        const lone = /^\s*[-−]\s*$/.test(inp.value) && inp.dataset.neg;
        inp.value = lone ? '-' : n == null ? '' : (n < 0 && inp.dataset.neg ? '-' : '') + fmtN(Math.abs(n));
      });
    });
  }

  function sheet({ title, body, foot = '', onMount, wide = false }) {
    const bd = document.createElement('div');
    bd.className = 'sheet-backdrop';
    bd.innerHTML = `<div class="sheet${wide ? ' wide' : ''}" role="dialog" aria-modal="true" aria-label="${esc(title)}">
      <header><h3>${esc(title)}</h3><button class="close" type="button" aria-label="Fermer">×</button></header>
      <div class="body">${body}</div>${foot ? `<div class="foot">${foot}</div>` : ''}</div>`;
    $sheets.append(bd);
    S.sheetOpen++;
    const el = bd.querySelector('.sheet');
    let closed = false;
    const ctx = {
      el, bd, q: s => el.querySelector(s), qa: s => [...el.querySelectorAll(s)],
      set onKey(f) { bd._onkey = f; },
      close() {
        if (closed) return; closed = true;
        bd.remove(); S.sheetOpen--;
        if (S.sheetOpen === 0 && S.pendingRefresh) { S.pendingRefresh = false; refresh(); }
        ctx.onClose && ctx.onClose();
      },
    };
    bd.addEventListener('mousedown', e => { if (e.target === bd) ctx.close(); });
    el.querySelector('.close').onclick = () => ctx.close();
    bindMoney(el);
    onMount && onMount(ctx);
    const f = el.querySelector('.autofocus');
    if (f) setTimeout(() => f.focus(), 80);
    return ctx;
  }
  // Raccourcis clavier (ordinateur)
  document.addEventListener('keydown', e => {
    const top = $sheets.lastElementChild;
    if (e.key === 'Escape') {
      const combo = top && top.querySelector('.combo .drop:not(.hidden)');
      if (combo) { combo.classList.add('hidden'); return; }
      if (top) top.querySelector('.close').click();
      return;
    }
    if (top) { if (top._onkey) top._onkey(e); return; }
    const typing = /^(INPUT|TEXTAREA|SELECT)$/.test(e.target.tagName);
    if (typing || e.ctrlKey || e.metaKey || e.altKey || !S.me || !S.me.active || !document.getElementById('main')) return;
    if (e.key === 'n' || e.key === 'N' || e.key === 'F2') { e.preventDefault(); sheetVente(); }
    else if (e.key === '/') { const s = document.querySelector('main .search input'); if (s) { e.preventDefault(); s.focus(); } }
  });

  function ask(title, html, okLabel = 'Confirmer', danger = false) {
    return new Promise(resolve => {
      let done = false;
      const s = sheet({
        title, body: `<div>${html}</div>`,
        foot: `<button class="btn" data-x="no">Annuler</button><button class="btn ${danger ? 'danger' : 'primary'}" data-x="yes">${esc(okLabel)}</button>`,
        onMount(c) {
          c.q('[data-x=no]').onclick = () => { done = true; resolve(false); c.close(); };
          c.q('[data-x=yes]').onclick = () => { done = true; resolve(true); c.close(); };
        },
      });
      s.onClose = () => { if (!done) resolve(false); };
    });
  }

  function askText(title, label, { okLabel = 'Valider', danger = false, required = true, placeholder = '' } = {}) {
    return new Promise(resolve => {
      let done = false;
      const s = sheet({
        title,
        body: `<label class="field"><span>${esc(label)}</span><textarea class="input autofocus" name="t" placeholder="${esc(placeholder)}"></textarea></label>`,
        foot: `<button class="btn" data-x="no">Retour</button><button class="btn ${danger ? 'danger' : 'primary'}" data-x="yes">${esc(okLabel)}</button>`,
        onMount(c) {
          c.q('[data-x=no]').onclick = () => { done = true; resolve(null); c.close(); };
          c.q('[data-x=yes]').onclick = () => {
            const v = c.q('[name=t]').value.trim();
            if (required && !v) { toast('Ce champ est obligatoire.', 'bad'); return; }
            done = true; resolve(v); c.close();
          };
        },
      });
      s.onClose = () => { if (!done) resolve(null); };
    });
  }

  async function busy(btn, fn) {
    if (!btn || btn.disabled) return;
    const old = btn.innerHTML;
    btn.disabled = true; btn.textContent = 'Patientez…';
    try { return await fn(); }
    catch (e) { toast(e.message || String(e), 'bad'); }
    finally { if (btn.isConnected) { btn.disabled = false; btn.innerHTML = old; } }
  }

  const field = (label, inner, help = '') => `<label class="field"><span>${label}</span>${inner}${help ? `<small>${help}</small>` : ''}</label>`;
  // attrs peut contenir « autofocus » (champ sélectionné à l'ouverture) et d'autres attributs.
  const moneyInput = (name, attrs = '') => {
    const auto = /autofocus/.test(attrs);
    const rest = attrs.replace(/class="[^"]*"/g, '').replace(/autofocus/g, '');
    return `<input class="input money${auto ? ' autofocus' : ''}" name="${name}" inputmode="numeric" autocomplete="off" placeholder="0" ${rest}>`;
  };
  const options = (list, sel, empty) => (empty !== undefined ? `<option value="">${esc(empty)}</option>` : '') +
    list.map(([v, l]) => `<option value="${esc(v)}" ${String(v) === String(sel ?? '') ? 'selected' : ''}>${esc(l)}</option>`).join('');
  const section = (title, inner, extra = '') => `<div class="section"><h2><span>${title}</span>${extra}</h2>${inner}</div>`;

  function chipGroup(ctx, name, onChange) {
    ctx.qa(`[data-chip="${name}"]`).forEach(b => b.addEventListener('click', () => {
      ctx.qa(`[data-chip="${name}"]`).forEach(x => x.classList.toggle('on', x === b));
      onChange && onChange(b.dataset.v);
    }));
  }
  const chipVal = (ctx, name) => { const b = ctx.q(`[data-chip="${name}"].on`); return b ? b.dataset.v : null; };

  // ---------------------------------------------------------------- accès aux données
  function translateAuth(msg) {
    const m = String(msg || '');
    if (/Invalid login/i.test(m)) return 'Identifiant ou mot de passe incorrect.';
    if (/already registered|already exists/i.test(m)) return 'Cet identifiant est déjà utilisé.';
    if (/at least 6/i.test(m)) return 'Le mot de passe doit contenir au moins 6 caractères.';
    if (/Email not confirmed/i.test(m)) return 'Compte non confirmé : la gérante doit désactiver la confirmation par e-mail dans Supabase.';
    if (/Failed to fetch|NetworkError|Load failed/i.test(m)) return 'Pas de connexion Internet.';
    if (/rate limit/i.test(m)) return 'Trop de tentatives. Réessayez dans quelques minutes.';
    return m;
  }
  const loginToEmail = l => { const v = String(l || '').trim().toLowerCase(); return v.includes('@') ? v : `${v.replace(/\s+/g, '.')}@${LOGIN_DOMAIN}`; };

  // « ?demo » dans l'adresse : version d'essai, sans toucher à la base de l'équipe.
  const FORCE_DEMO = new URLSearchParams(location.search).has('demo');
  function getConfig() {
    if (FORCE_DEMO) return null;
    const c = window.AYAEL_CONFIG || {};
    if (c.supabaseUrl && c.supabaseAnonKey) return { url: c.supabaseUrl, key: c.supabaseAnonKey, fromFile: true };
    try { const s = JSON.parse(store('cfg') || 'null'); if (s && s.url && s.key) return s; } catch (e) { /* rien */ }
    return null;
  }

  function makeSupabase(cfg) {
    if (!window.supabase) throw new Error('Bibliothèque Supabase non chargée (connexion Internet ?).');
    const sb = window.supabase.createClient(cfg.url, cfg.key, { auth: { persistSession: true, autoRefreshToken: true } });
    return {
      mode: 'supabase',
      async rpc(name, p) {
        const { data, error } = await sb.rpc(name, { p });
        if (error) throw new Error(translateAuth(error.message));
        return data;
      },
      async session() { const { data } = await sb.auth.getSession(); return data.session; },
      async signIn(login, pwd) {
        const { error } = await sb.auth.signInWithPassword({ email: loginToEmail(login), password: pwd });
        if (error) throw new Error(translateAuth(error.message));
      },
      async signUp(name, login, pwd) {
        const { data, error } = await sb.auth.signUp({ email: loginToEmail(login), password: pwd, options: { data: { full_name: name } } });
        if (error) throw new Error(translateAuth(error.message));
        if (!data.session) throw new Error('Compte créé, mais la confirmation par e-mail est activée dans Supabase. La gérante doit la désactiver (voir le guide).');
      },
      async signOut() { await sb.auth.signOut(); },
      async changePassword(pwd) { const { error } = await sb.auth.updateUser({ password: pwd }); if (error) throw new Error(translateAuth(error.message)); },
      subscribe(cb) {
        sb.channel('ayael-changes')
          .on('postgres_changes', { event: 'INSERT', schema: 'public', table: 'change_log' }, () => cb())
          .subscribe(status => setLive(status === 'SUBSCRIBED'));
      },
    };
  }

  async function makeDemo() {
    const { PGlite } = await import(PGLITE_URL);
    const db = new PGlite('idb://ayael-demo');
    await db.waitReady;
    const txt = async f => { const r = await fetch('sql/' + f, { cache: 'no-store' }); if (!r.ok) throw new Error('Fichier manquant : sql/' + f); return r.text(); };
    const exists = (await db.query("select to_regclass('public.operations') is not null as ok")).rows[0].ok;
    if (!exists) {
      await db.exec(await txt('demo-auth.sql'));
      await db.exec(await txt('schema.sql'));
      await db.exec(await txt('demo-seed.sql'));
      store('demo-schema', DEMO_SCHEMA_VERSION);
    } else if (store('demo-schema') !== DEMO_SCHEMA_VERSION) {
      await db.exec(await txt('demo-auth.sql'));
      await db.exec(await txt('schema.sql'));
      store('demo-schema', DEMO_SCHEMA_VERSION);
    }
    let uid = store('demo-uid') || '';
    const listeners = [];
    let queue = Promise.resolve();
    const run = fn => (queue = queue.then(fn, fn));
    return {
      mode: 'demo',
      rpc(name, p) {
        if (!/^[a-z_]+$/.test(name)) return Promise.reject(new Error('Fonction inconnue'));
        return run(async () => {
          await db.query("select set_config('app.uid', $1, false)", [uid]);
          let r;
          try { r = (await db.query(`select public.${name}($1::jsonb) as r`, [JSON.stringify(p || {})])).rows[0].r; }
          catch (e) { throw new Error(e.message); }
          if (/^(op_|save_|closing_submit|closing_validate|import_)/.test(name)) setTimeout(() => listeners.forEach(f => f()), 0);
          return r;
        });
      },
      async session() { return uid ? { user: { id: uid } } : null; },
      async users() { return (await db.query('select id, full_name, role, active from profiles order by role desc, full_name')).rows; },
      async signInAs(id) { uid = id; store('demo-uid', id); },
      async signUpDemo(name) {
        const r = await db.query('insert into auth.users(email, raw_user_meta_data) values ($1, $2) returning id',
          [`demo${Date.now()}@${LOGIN_DOMAIN}`, JSON.stringify({ full_name: name })]);
        return r.rows[0].id;
      },
      async signOut() { uid = ''; store('demo-uid', null); },
      async reset() {
        await db.close();
        await new Promise(res => { const r = indexedDB.deleteDatabase('/pglite/ayael-demo'); r.onsuccess = r.onerror = r.onblocked = () => res(); });
        store('demo-uid', null); store('demo-schema', null);
      },
      subscribe(cb) { listeners.push(cb); setLive(true); },
    };
  }

  // ---------------------------------------------------------------- démarrage
  async function boot() {
    if ('serviceWorker' in navigator && (location.protocol === 'https:' || location.hostname === 'localhost')) {
      navigator.serviceWorker.register('sw.js').catch(() => {});
    }
    window.addEventListener('online', renderBanners);
    window.addEventListener('offline', renderBanners);
    window.addEventListener('hashchange', route);
    try {
      const cfg = getConfig();
      if (cfg) S.api = makeSupabase(cfg);
      else if (FORCE_DEMO || store('mode') === 'demo') { splash('Préparation du mode démo…'); S.api = await makeDemo(); }
      else return showSetup();
      await startSession();
    } catch (e) {
      $app.innerHTML = `<div class="login"><img class="logo" src="icons/logo.png" alt="Ayael Entreprise">
        <div class="alert bad">${esc(e.message || e)}</div>
        <button class="btn block" onclick="location.reload()">Réessayer</button>
        ${store('mode') === 'demo' ? '<button class="btn ghost block" id="quitdemo">Quitter le mode démo</button>' : ''}</div>`;
      const q = document.getElementById('quitdemo');
      if (q) q.onclick = () => { store('mode', null); location.href = location.pathname; };
    }
  }

  function splash(msg) { $app.innerHTML = `<div class="splash"><img src="icons/logo.png" alt="Ayael Entreprise"><p class="muted">${esc(msg)}</p></div>`; }

  async function startSession() {
    const sess = await S.api.session();
    if (!sess) return showLogin();
    await loadBoot();
    if (!S.me) return showPending('Votre profil n’a pas été trouvé. Vérifiez que le schéma de la base a bien été installé (voir le guide).');
    if (!S.me.active) return showPending();
    if (!S.subscribed) { S.subscribed = true; S.api.subscribe(onRemoteChange); }
    renderShell();
    route();
  }

  async function loadBoot() {
    const b = await rpc('app_bootstrap');
    S.me = b.me; S.accounts = b.accounts || []; S.products = b.products || []; S.team = b.team || []; S.settings = b.settings || {};
    S.agences = b.agences || [];
    if (S.me && S.me.role === 'gerante') {
      const saved = Number(store('agence')) || null;
      S.agence = S.agences.some(g => g.id === saved) ? saved : null;
    } else S.agence = null;
  }

  const onRemoteChange = debounce(async () => {
    try { await loadBoot(); } catch (e) { return; }
    if (S.me && !S.me.active) return location.reload();
    renderTabbar();
    refresh();
  }, 500);

  function refresh() {
    if (!S.view) return;
    if (S.sheetOpen > 0) { S.pendingRefresh = true; return; }
    const v = VIEWS[S.view];
    if (v && v.live) (v.refresh || v.render)(document.getElementById('main')).catch(() => {});
  }

  // ---------------------------------------------------------------- écrans d'accueil
  function showSetup() {
    $app.innerHTML = `<div class="login">
      <img class="logo" src="icons/logo.png" alt="Ayael Entreprise">
      <h1>Bienvenue</h1>
      <p class="muted">Cette application n’est pas encore reliée à la base de données de l’équipe.</p>
      <button class="btn primary block" id="demo">Essayer en mode démo</button>
      <p class="small muted" style="margin:8px 0 24px">Données fictives, enregistrées uniquement sur cet appareil.</p>
      <details class="card pad"><summary style="font-weight:650;cursor:pointer">Relier la base Supabase de l’équipe</summary>
        <div style="margin-top:14px">
          ${field('Project URL', '<input class="input" id="cu" placeholder="https://xxxx.supabase.co" autocomplete="off">')}
          ${field('Clé « anon public »', '<input class="input" id="ck" placeholder="eyJ…" autocomplete="off">', 'Supabase › Project Settings › API. Il est préférable de la mettre dans config.js avant la publication.')}
          <button class="btn block" id="save">Enregistrer</button>
        </div></details></div>`;
    document.getElementById('demo').onclick = () => { store('mode', 'demo'); location.reload(); };
    document.getElementById('save').onclick = () => {
      const url = document.getElementById('cu').value.trim(), key = document.getElementById('ck').value.trim();
      if (!/^https:\/\/.+/.test(url) || key.length < 20) return toast('Vérifiez l’adresse et la clé.', 'bad');
      store('cfg', JSON.stringify({ url, key })); location.reload();
    };
  }

  async function showLogin() {
    if (S.api.mode === 'demo') {
      const users = await S.api.users();
      $app.innerHTML = `<div class="login">
        <img class="logo" src="icons/logo.png" alt="Ayael Entreprise">
        <div class="alert warn">Mode démo : choisissez qui vous êtes pour essayer l’application.</div>
        <div class="list" style="margin-bottom:16px">${users.map(u => `<button class="row" data-uid="${u.id}">
          <div class="ico ${u.role === 'gerante' ? 'rc' : 'neu'}">${esc(initials(u.full_name))}</div>
          <div class="main"><div class="t">${esc(u.full_name)}</div><div class="s">${u.role === 'gerante' ? 'Gérante' : 'Agent'}${u.active ? '' : ' · en attente d’activation'}</div></div>${icon('chev')}</button>`).join('')}</div>
        <button class="btn block" id="newu">Créer un compte agent (démo)</button>
        <button class="btn ghost block" id="quit">Quitter le mode démo</button></div>`;
      $app.querySelectorAll('[data-uid]').forEach(b => b.onclick = async () => { await S.api.signInAs(b.dataset.uid); startSession(); });
      document.getElementById('newu').onclick = async () => {
        const n = await askText('Nouvel agent', 'Nom complet', { okLabel: 'Créer' });
        if (!n) return;
        const id = await S.api.signUpDemo(n); await S.api.signInAs(id); startSession();
      };
      document.getElementById('quit').onclick = () => { store('mode', null); location.href = location.pathname; };
      return;
    }
    let signup = false;
    const draw = () => {
      $app.innerHTML = `<div class="login">
        <img class="logo" src="icons/logo.png" alt="Ayael Entreprise">
        <h1>${signup ? 'Créer mon compte' : 'Connexion'}</h1>
        <p class="muted">${signup ? 'Après la création, la gérante doit activer votre compte.' : 'Registre de caisse et de recharges'}</p>
        <form id="f" autocomplete="on">
          ${signup ? field('Nom complet', '<input class="input" name="name" required autocomplete="name">') : ''}
          ${field('Identifiant', '<input class="input" name="login" required autocomplete="username" autocapitalize="none" spellcheck="false" placeholder="ex. awa">', signup ? 'Lettres et chiffres, sans espace. Il servira à vous connecter.' : '')}
          ${field('Mot de passe', `<input class="input" name="pwd" type="password" required minlength="6" autocomplete="${signup ? 'new-password' : 'current-password'}">`)}
          ${signup ? field('Confirmer le mot de passe', '<input class="input" name="pwd2" type="password" required minlength="6" autocomplete="new-password">') : ''}
          <button class="btn primary block" type="submit">${signup ? 'Créer le compte' : 'Se connecter'}</button>
        </form>
        <button class="btn ghost block" id="sw" style="margin-top:8px">${signup ? 'J’ai déjà un compte' : 'Créer un compte'}</button></div>`;
      document.getElementById('sw').onclick = () => { signup = !signup; draw(); };
      const f = document.getElementById('f');
      f.onsubmit = e => {
        e.preventDefault();
        const v = Object.fromEntries(new FormData(f));
        busy(f.querySelector('[type=submit]'), async () => {
          if (signup) {
            if (!/^[a-zA-Z0-9._-]{2,}$/.test(v.login.trim())) throw new Error('Identifiant : lettres, chiffres, point ou tiret, sans espace.');
            if (v.pwd !== v.pwd2) throw new Error('Les deux mots de passe sont différents.');
            await S.api.signUp(v.name.trim(), v.login, v.pwd);
          } else {
            await S.api.signIn(v.login, v.pwd);
          }
          await startSession();
        });
      };
    };
    draw();
  }

  function showPending(msg) {
    $app.innerHTML = `<div class="login">
      <img class="logo" src="icons/logo.png" alt="Ayael Entreprise">
      <h1>Compte en attente</h1>
      <p class="muted">${esc(msg || 'Votre compte est créé. La gérante doit maintenant l’activer et choisir vos droits (caisse, recharges).')}</p>
      <button class="btn primary block" id="r">Actualiser</button>
      <button class="btn ghost block" id="o">Se déconnecter</button></div>`;
    document.getElementById('r').onclick = () => startSession();
    document.getElementById('o').onclick = async () => { await S.api.signOut(); location.reload(); };
  }

  // ---------------------------------------------------------------- coquille
  // Barre du bas (téléphone) : le « + » central ouvre une nouvelle opération.
  function tabs() {
    if (isMgr()) return [['tableau', 'Tableau', 'home'], ['journal', 'Journal', 'list'], ['+'], ['clotures', 'Clôtures', 'lock'], ['plus', 'Plus', 'grid']];
    return [['accueil', 'Accueil', 'home'], ['journal', 'Journal', 'list'], ['+'], ['clients', 'Clients', 'users'], ['cloture', 'Clôture', 'lock']];
  }
  // Menu latéral (ordinateur)
  function navItems() {
    if (isDir()) return [['tableau', 'Tableau de bord', 'home'], ['journal', 'Journal', 'list'], ['clients', 'Clients', 'users'], ['clotures', 'Clôtures', 'lock'],
      ['commissions', 'Commissions à vérifier', 'card'], ['rapports', 'Rapports', 'chart'], 'Réglages', ['agences', 'Agences', 'bank'], ['equipe', 'Équipe', 'user'],
      ['comptes', 'Comptes et commissions', 'bank'], ['produits', 'Produits de caisse', 'tag'], ['cartes', 'Cartes clients', 'card'], ['doublons', 'Clients en double', 'users'],
      ['import', 'Importer depuis Excel', 'upload'], ['plus', 'Autres réglages', 'grid']];
    if (isMgr()) return [['tableau', 'Tableau de bord', 'home'], ['journal', 'Journal', 'list'], ['clients', 'Clients', 'users'], ['clotures', 'Clôtures', 'lock'],
      ['commissions', 'Commissions à vérifier', 'card'], ['rapports', 'Rapports', 'chart'], ['cartes', 'Cartes clients', 'card'], ['plus', 'Plus', 'grid']];
    return [['accueil', 'Accueil', 'home'], ['journal', 'Journal', 'list'], ['clients', 'Clients', 'users'],
      ...(S.me.can_recharge ? [['commissions', 'Commissions à vérifier', 'card']] : []), ['cloture', 'Clôture du soir', 'lock']];
  }
  const firstView = () => isMgr() ? 'tableau' : 'accueil';

  function renderShell() {
    $app.innerHTML = `
      <aside class="sidebar">
        <img class="logo" src="icons/logo.png" alt="Ayael Entreprise">
        <button class="hero-cta" data-do="vente">${icon('plus')}Nouvelle opération <span class="kbd">N</span></button>
        <nav id="sidenav" aria-label="Menu"></nav>
        <button class="who" data-do="me"><span class="avatar">${esc(initials(S.me.full_name))}</span>
          <span><span class="t" style="display:block">${esc(S.me.full_name)}</span><span class="s">${esc(roleLabel(S.me.role))}${S.me.agence ? ' · ' + esc(S.me.agence) : ''}</span></span></button>
      </aside>
      <div class="shell">
        <header class="topbar">
          <img src="icons/logo.png" alt="Ayael Entreprise">
          <div class="title" id="vtitle"></div>
          ${agenceSwitch()}
          <button class="btn primary small newop" data-do="vente">${icon('plus')} Nouvelle opération</button>
          <span class="live" id="live">…</span>
          <button class="avatar" data-do="me" aria-label="Mon compte">${esc(initials(S.me.full_name))}</button>
        </header>
        <div id="banners"></div>
        <main id="main"></main>
      </div>
      <nav class="tabbar" id="tabbar" aria-label="Navigation"></nav>`;
    renderBanners();
    renderTabbar();
    setLive(S.liveOn);
    const sw = document.getElementById('agsel');
    if (sw) sw.onchange = () => { S.agence = sw.value ? Number(sw.value) : null; store('agence', sw.value || null); route(); };
  }

  // Choix de l'agence affichée (direction) ou nom de l'agence (gérant, agent).
  function agenceSwitch() {
    const ags = S.agences || [];
    if (isDir() && ags.length > 1) {
      return `<select class="input agsel" id="agsel" aria-label="Agence affichée">${options(ags.map(g => [g.id, g.name]), S.agence || '', 'Toutes les agences')}</select>`;
    }
    return ags.length > 1 || (S.me.agence && !S.me.is_main_agence) ? `<span class="tag brand agtag">${esc(S.me.agence || '')}</span>` : '';
  }
  const agName = id => ((S.agences || []).find(g => g.id === Number(id)) || {}).name || '';

  function renderTabbar() {
    const badge = k => k === 'clotures' && S.pendingClosings ? `<span class="badge">${S.pendingClosings}</span>` : '';
    const nav = document.getElementById('tabbar');
    if (nav) nav.innerHTML = tabs().map(([k, l, ic]) => k === '+'
      ? `<a href="#" data-do="vente" class="plus" aria-label="Nouvelle opération"><span class="round">${icon('plus')}</span><span>Opération</span></a>`
      : `<a href="#/${k}" class="${S.view === k ? 'active' : ''}">${icon(ic)}<span>${l}</span>${badge(k)}</a>`).join('');
    const side = document.getElementById('sidenav');
    if (side) side.innerHTML = navItems().map(it => typeof it === 'string' ? `<div class="sep">${esc(it)}</div>`
      : `<a href="#/${it[0]}" class="${S.view === it[0] ? 'active' : ''}">${icon(it[2])}<span>${esc(it[1])}</span>${badge(it[0])}</a>`).join('');
  }

  function renderBanners() {
    const b = document.getElementById('banners');
    if (!b) return;
    b.innerHTML = (S.api && S.api.mode === 'demo' ? '<div class="banner demo">MODE DÉMO — données d’essai, uniquement sur cet appareil</div>' : '') +
      (navigator.onLine ? '' : '<div class="banner offline">Pas de connexion Internet : les opérations ne peuvent pas être enregistrées.</div>');
  }

  function setLive(on) {
    S.liveOn = on;
    const el = document.getElementById('live');
    if (!el) return;
    el.classList.toggle('on', !!on);
    el.textContent = on ? 'En direct' : 'Hors ligne';
  }

  async function route() {
    if (!S.me || !S.me.active || !document.getElementById('main')) return;
    let key = (location.hash.match(/^#\/([a-z]+)/) || [])[1];
    const allowed = k => VIEWS[k] && (!VIEWS[k].mgr || isMgr()) && (!VIEWS[k].dir || isDir()) && (!VIEWS[k].need || isMgr() || S.me[VIEWS[k].need]);
    if (!key || !allowed(key)) {
      key = firstView();
      if (location.hash !== '#/' + key) { history.replaceState(null, '', '#/' + key); }
    }
    S.view = key;
    const v = VIEWS[key];
    document.getElementById('vtitle').textContent = v.title;
    document.title = `${v.title} · Ayael`;
    const main = document.getElementById('main');
    renderTabbar();
    main.innerHTML = '<div class="spinner" aria-label="Chargement"></div>';
    window.scrollTo(0, 0);
    try { await v.render(main); }
    catch (e) { main.innerHTML = `<div class="alert bad">${esc(e.message || e)}</div><button class="btn block" data-go="${key}">Réessayer</button>`; }
  }

  // ---------------------------------------------------------------- affichage des opérations et soldes
  function opInfo(o) {
    const r = { ico: 'neu', sym: '•', title: KIND[o.kind] || o.kind, sub: '', sign: 0 };
    switch (o.kind) {
      case 'caisse': {
        const inn = o.direction === 'entree';
        Object.assign(r, { ico: inn ? 'in' : 'out', sym: inn ? icon('in') : icon('out'), title: o.product || 'Caisse', sign: inn ? 1 : -1,
          sub: [o.client_name, o.details, o.account2].filter(Boolean).join(' · ') });
        break;
      }
      case 'recharge':
        Object.assign(r, { ico: 'rc', sym: icon('card'), title: o.client_name || o.holder_name || 'Recharge', sign: 1,
          sub: `${o.card_number || ''} · ${o.account || ''} · com. ${fmt(o.commission)} · ${/^Caisse/.test(o.account2 || '') ? 'espèces' : (o.account2 || '')}` });
        break;
      case 'depense':
        Object.assign(r, { ico: 'out', sym: icon('out'), title: 'Dépense : ' + (o.details || ''), sign: -1, sub: [o.category, o.account].filter(Boolean).join(' · ') });
        break;
      case 'approvisionnement':
        Object.assign(r, { ico: 'rc', sym: icon('truck'), title: 'Approvisionnement ' + (o.account || ''), sign: 1, sub: [o.account2 ? 'depuis ' + o.account2 : '', o.details].filter(Boolean).join(' · ') });
        break;
      case 'transfert':
        Object.assign(r, { sym: icon('swap'), title: `${o.account} → ${o.account2}`, sub: o.details || 'Transfert' });
        break;
      case 'ajustement':
        Object.assign(r, { sym: icon('adjust'), title: 'Ajustement ' + (o.account || ''), sign: Math.sign(o.amount), sub: o.details || '' });
        break;
      case 'cloture':
        Object.assign(r, { sym: icon('lock'), title: 'Remise de caisse', sub: o.details || '' });
        break;
    }
    return r;
  }

  function opRow(o, showAgent) {
    S.ops.set(o.id, o);
    const i = opInfo(o);
    const amt = o.kind === 'ajustement' ? Math.abs(o.amount) : o.amount;
    const cls = o.status !== 'ok' ? '' : i.sign > 0 ? 'pos' : i.sign < 0 ? 'neg' : '';
    const flag = o.status !== 'ok' ? '<span class="tag red">Annulée</span>' : o.flags ? '<span class="tag amber">À vérifier</span>' : '';
    return `<button class="row ${o.status !== 'ok' ? 'cancelled' : ''}" data-op="${o.id}">
      <div class="ico ${i.ico}">${i.sym}</div>
      <div class="main"><div class="t">${esc(i.title)}</div>
        <div class="s">${fmtTime(o.created_at)}${showAgent && o.agent ? ' · ' + esc(o.agent) : ''}${i.sub ? ' · ' + esc(i.sub) : ''}</div></div>
      <div class="end"><div class="amt num ${cls}">${i.sign > 0 ? '+' : i.sign < 0 ? '−' : ''}${fmt(amt)}</div>${flag}</div></button>`;
  }

  const opList = (list, showAgent = isMgr(), empty = 'Aucune opération.') =>
    list && list.length ? `<div class="list">${list.map(o => opRow(o, showAgent)).join('')}</div>` : `<div class="list"><div class="empty">${empty}</div></div>`;

  function balRow(b, clickable) {
    const low = Number(b.low_alert) > 0 && b.balance < b.low_alert;
    const neg = b.balance < 0;
    const sub = b.kind === 'caisse' ? `${b.today_ops} opération(s) aujourd’hui${b.last_at ? ' · ' + ago(b.last_at) : ''}`
      : b.kind === 'gtp' ? `${b.bank}${b.last_at ? ' · dernière opération ' + ago(b.last_at) : ''}` : (ACCOUNT_KIND[b.kind] || '');
    const tagName = clickable ? 'button' : 'div';
    const agTxt = (S.agences || []).length > 1 && isMgr() ? (b.agence ? b.agence : 'Commun') + ' · ' : '';
    return `<${tagName} class="row" ${clickable ? `data-acc="${b.id}"` : ''}>
      <div class="ico ${b.kind === 'gtp' ? 'rc' : 'neu'}">${b.kind === 'gtp' ? esc((b.bank || '?').slice(0, 3)) : b.kind === 'caisse' ? esc(initials(b.agent)) : icon(b.kind === 'electronique' ? 'cash' : 'bank')}</div>
      <div class="main"><div class="t">${esc(b.name)}</div><div class="s">${esc(agTxt + sub)}</div></div>
      <div class="end"><div class="amt num ${neg ? 'neg' : low ? 'warn' : ''}">${fmt(b.balance)}</div>${neg ? '<span class="tag red">Négatif</span>' : low ? '<span class="tag amber">Bas</span>' : ''}</div></${tagName}>`;
  }
  const balList = (list, clickable) => list.length ? `<div class="list">${list.map(b => balRow(b, clickable)).join('')}</div>` : '<div class="list"><div class="empty">Aucun compte.</div></div>';

  function rememberBalances(d) { (d.balances || []).forEach(b => S.bal.set(b.id, Number(b.balance))); S.pendingClosings = isMgr() ? d.pending_closings : 0; renderTabbar(); }

  // ---------------------------------------------------------------- écrans agent
  async function vAccueil(m) {
    const d = await rpc('dashboard');
    rememberBalances(d);
    const mine = d.balances.find(b => b.id === S.me.caisse_id) || { balance: 0 };
    const gtp = d.balances.filter(b => b.kind === 'gtp');
    const el = d.balances.filter(b => b.kind === 'electronique');
    m.innerHTML = `
      <div class="cols">
        <div>
          <div class="hero"><div class="label">Ma caisse (espèces)</div><div class="big num">${fmt(mine.balance)}</div>
            <div class="sub">${d.today.ops_n} opération(s) aujourd’hui${d.pending_closings ? ' · clôture en attente de validation' : ''}</div></div>
          <button class="hero-cta only-mobile" data-do="vente">${icon('plus')}Nouvelle opération</button>
          ${d.commissions_to_check ? `<div class="alert warn"><span>${d.commissions_to_check} commission(s) de recharge à vérifier sur le rapport du GTP. <a href="#/commissions">Vérifier</a></span></div>` : ''}
          <div class="actions">
            <button class="action" data-do="depense">${icon('out')}Dépense</button>
            <a class="action" href="#/clients">${icon('users')}Clients</a>
            <a class="action" href="#/cloture">${icon('lock')}Clôture</a>
          </div>
          ${section('Mes dernières opérations', opList(d.recent, false, 'Aucune opération pour le moment.'))}
        </div>
        <div>
          ${S.me.can_recharge ? section('Soldes GTP', balList(gtp)) : ''}
          ${section('Comptes électroniques', balList(el))}
        </div>
      </div>`;
  }

  function sheetDepense() {
    const accs = isMgr() ? S.accounts.filter(a => a.active && a.kind !== 'gtp' && myAcc(a) && (a.kind !== 'caisse' || a.agent_id === S.me.id)) : [];
    sheet({
      title: 'Dépense / sortie',
      body: `${field('Montant', moneyInput('amount', 'class="input money autofocus"'))}
        ${field('Motif', '<input class="input" name="details" placeholder="ex. Transport, achat de papier…" autocomplete="off">')}
        <div class="field"><span>Catégorie</span><div class="chips wrap">${DEP_CATS.map((x, i) => `<button type="button" class="chip ${i === DEP_CATS.length - 1 ? 'on' : ''}" data-chip="cat" data-v="${esc(x)}">${esc(x)}</button>`).join('')}</div></div>
        ${accs.length ? field('Payé depuis', `<select class="input" name="acc">${options(accs.map(a => [a.id, `${a.name} (${fmt(S.bal.get(a.id) ?? 0)})`]), S.me.caisse_id)}</select>`) : '<p class="small muted">Payé depuis votre caisse.</p>'}`,
      foot: `<button class="btn primary" data-x="ok">Enregistrer la dépense</button>`,
      onMount(c) {
        chipGroup(c, 'cat');
        c.q('[data-x=ok]').onclick = e => busy(e.currentTarget, async () => {
          const payload = { amount: parseAmt(c.q('[name=amount]').value), details: c.q('[name=details]').value, category: chipVal(c, 'cat') };
          if (c.q('[name=acc]')) payload.account_id = Number(c.q('[name=acc]').value);
          const r = await rpc('op_depense', payload);
          c.close(); toast(`Dépense enregistrée. Solde : ${fmt(r.balance)}`, 'ok'); refreshNow();
        });
      },
    });
  }

  // ---------------------------------------------------------------- nouvelle opération (vente)
  // Un seul formulaire : client → produit(s) → paiement → reçu.
  function sheetVente(preset = {}) {
    const canC = isMgr() || S.me.can_caisse, canR = isMgr() || S.me.can_recharge;
    if (!canC && !canR) return toast('Vous n’avez aucun droit de saisie. Demandez à la gérante.', 'bad');
    const st = { client: preset.client || null, lines: [], pay: S.me.caisse_id, sel: '', cardId: null, gtpId: null, commEdited: false, passName: '' };
    if (preset.card && canR) { st.sel = 'R'; st.cardId = preset.card.id; }
    const prods = S.products.filter(p => p.active);
    const pays = [[S.me.caisse_id, 'Espèces'], ...S.accounts.filter(a => a.active && myAcc(a) && (a.kind === 'electronique' || a.kind === 'banque'))
      .map(a => [a.id, a.kind === 'banque' ? 'Versement bancaire' : a.name])];
    const cats = {};
    prods.forEach(p => { const k = p.category || 'Autres'; (cats[k] = cats[k] || []).push(p); });
    const prodOptions = `<option value="">— Choisir le produit ou le service —</option>
      ${canR ? '<optgroup label="Cartes prépayées"><option value="R">Recharge de carte prépayée</option></optgroup>' : ''}
      ${canC ? Object.entries(cats).map(([k, list]) => `<optgroup label="${esc(k)}">${list.map(p => `<option value="${p.id}">${esc(p.name)}${p.direction === 'sortie' ? ' (on remet l’argent au client)' : ''}</option>`).join('')}</optgroup>`).join('') : ''}`;

    const c = sheet({
      title: 'Nouvelle opération', wide: true,
      body: `<div class="vente">
        <div>
          <div class="field"><span>1. Client</span><div id="vclient"></div></div>
          <label class="field"><span>2. Produit ou service</span><select class="input" id="vprod">${prodOptions}</select></label>
          <div id="vfields"></div>
        </div>
        <div>
          <div class="field"><span>Produits de l’opération</span><div class="list" id="vlines"></div></div>
          <div id="vtot"></div>
          <div id="vpay"></div>
        </div></div>`,
      foot: `<span class="hint only-desktop">Entrée : ajouter · Ctrl+Entrée : valider</span>
        <button class="btn" data-x="cancel">Annuler</button>
        <button class="btn primary" data-x="ok">${icon('receipt')} Valider<span class="desk-inline"> et faire le reçu</span></button>`,
    });
    const $ = s => c.q(s);
    c.q('[data-x=cancel]').onclick = () => c.close();
    c.onKey = e => { if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); $('[data-x=ok]').click(); } };

    // --- client
    function drawClient() {
      const box = $('#vclient');
      if (st.client) {
        const cl = st.client;
        box.innerHTML = `<div class="clientcard"><div class="ico rc">${esc(initials(cl.name))}</div>
          <div class="main"><div class="t">${esc(cl.name)}</div><div class="s">${esc([cl.phone, activeCards(cl).length ? activeCards(cl).length + ' carte(s) active(s)' : 'aucune carte active'].filter(Boolean).join(' · '))}</div></div>
          <button class="btn small" type="button" id="vchg">Changer</button></div>`;
        $('#vchg').onclick = () => { st.client = null; drawClient(); drawFields(); $('#vcq').focus(); };
        return;
      }
      box.innerHTML = `<div class="combo"><div class="search" style="margin:0">${icon('search')}
          <input class="input" id="vcq" placeholder="Nom, téléphone ou n° de carte" autocomplete="off" value="${esc(st.passName)}"></div>
          <div class="drop hidden" id="vdrop" role="listbox"></div></div>
        <small class="muted small" style="display:block;margin-top:4px">Laissez vide pour un client de passage.</small>`;
      const inp = $('#vcq'), drop = $('#vdrop');
      let opts = [], hl = 0;
      const show = () => {
        drop.innerHTML = opts.map((o, i) => `<button type="button" class="opt ${i === hl ? 'hl' : ''}" data-i="${i}"><div class="t">${o.t}</div>${o.s ? `<div class="s">${o.s}</div>` : ''}</button>`).join('');
        drop.classList.toggle('hidden', !opts.length);
        drop.querySelectorAll('.opt').forEach(b => b.onmousedown = ev => { ev.preventDefault(); pick(Number(b.dataset.i)); });
      };
      const pick = i => { const o = opts[i]; if (o) o.go(); };
      const search = debounce(async () => {
        const q = inp.value.trim(); st.passName = q;
        if (!q) { opts = []; show(); return; }
        const res = await rpc('search_clients', { q, limit: 8 });
        if (inp.value.trim() !== q) return;
        res.forEach(x => S.clients.set(x.id, x));
        opts = res.map(x => ({ t: esc(x.name), s: esc([x.phone, ...activeCards(x).map(k => `${k.bank} ${k.card_number}`)].filter(Boolean).join(' · ')) || 'pas de carte active',
          go: () => { st.client = x; st.passName = ''; drawClient(); drawFields(); $('#vprod').focus(); } }));
        opts.push({ t: `＋ Créer le client « ${esc(q.toUpperCase())} »`, s: 'Nouvelle fiche client (nom, téléphone)',
          go: () => sheetClientEdit({ name: q }, saved => { st.client = saved; drawClient(); drawFields(); $('#vprod').focus(); }) });
        opts.push({ t: 'Client de passage', s: `Le nom « ${esc(q)} » figurera seulement sur le reçu`, go: () => { drop.classList.add('hidden'); $('#vprod').focus(); } });
        hl = 0; show();
      }, 200);
      inp.addEventListener('input', search);
      inp.addEventListener('focus', () => { if (opts.length) show(); });
      inp.addEventListener('blur', () => setTimeout(() => drop.classList.add('hidden'), 150));
      inp.addEventListener('keydown', e => {
        if (drop.classList.contains('hidden') || !opts.length) { if (e.key === 'Enter' && !e.ctrlKey) { e.preventDefault(); $('#vprod').focus(); } return; }
        if (e.key === 'ArrowDown') { e.preventDefault(); hl = (hl + 1) % opts.length; show(); }
        else if (e.key === 'ArrowUp') { e.preventDefault(); hl = (hl - 1 + opts.length) % opts.length; show(); }
        else if (e.key === 'Enter' && !e.ctrlKey) { e.preventDefault(); pick(hl); }
      });
    }

    // --- champs du produit choisi
    const pendingGtp = gid => st.lines.filter(l => l.type === 'recharge' && l.gtp_id === gid).reduce((s, l) => s - l.amount + l.commission, 0);
    function drawFields() {
      const box = $('#vfields');
      $('#vprod').value = st.sel;
      if (!st.sel) { box.innerHTML = ''; return; }
      const addBtn = `<button class="btn block" type="button" id="vadd">${icon('plus')} Ajouter à l’opération <span class="kbd only-desktop">Entrée</span></button>`;
      if (st.sel === 'R') {
        if (!st.client) { box.innerHTML = '<div class="alert info">Pour une recharge, choisissez d’abord le client (ou créez sa fiche) dans la case « Client ».</div>'; return; }
        const cards = activeCards(st.client);
        if (!cards.length) {
          box.innerHTML = `<div class="alert info">Ce client n’a pas encore de carte enregistrée.</div><button class="btn block" type="button" id="vnewcard">${icon('plus')} Ajouter sa carte</button>`;
          $('#vnewcard').onclick = addCard; return;
        }
        if (!cards.find(k => k.id === st.cardId)) st.cardId = cards[0].id;
        const card = cards.find(k => k.id === st.cardId);
        const gtps = S.accounts.filter(a => a.kind === 'gtp' && a.active && a.bank === card.bank);
        const list = gtps.length ? gtps : S.accounts.filter(a => a.kind === 'gtp' && a.active);
        if (!list.find(a => a.id === st.gtpId)) st.gtpId = (list.find(a => a.id === card.gtp_id) || list.find(a => a.id === S.rc.gtpId) || list[0] || {}).id;
        box.innerHTML = `
          <label class="field"><span>Carte</span><div style="display:flex;gap:8px"><select class="input" id="vcard">${options(cards.map(k => [k.id, `${k.card_number} · ${k.bank}${k.holder_name !== st.client.name ? ' · ' + k.holder_name : ''}`]), st.cardId)}</select>
            <button class="btn small" type="button" id="vnewcard" title="Ajouter une carte" style="min-height:48px">${icon('plus')}</button></div></label>
          ${field('Compte GTP', `<select class="input" id="vgtp">${options(list.map(a => [a.id, `${a.name} (${fmt((S.bal.get(a.id) ?? 0) + pendingGtp(a.id))})`]), st.gtpId)}</select>`)}
          <div class="grid2">${field('Montant rechargé', moneyInput('vamt', 'autofocus'))}
            <label class="field"><span>Commission</span>${moneyInput('vcom')}</label></div>
          <div class="small muted" id="vcinfo" style="margin:-8px 0 12px"></div>
          ${addBtn}`;
        bindMoney(box);
        st.commEdited = false;
        const upd = () => {
          const g = acc(st.gtpId), amt = parseAmt($('[name=vamt]').value) || 0, auto = commission(g, amt);
          if (!st.commEdited) $('[name=vcom]').value = amt ? fmtN(auto) : '';
          const com = parseAmt($('[name=vcom]').value) || 0;
          const after = (S.bal.get(st.gtpId) ?? 0) + pendingGtp(st.gtpId) - amt + com;
          $('#vcinfo').innerHTML = `Commission auto : ${esc(commRule(g))}${st.commEdited && com !== auto ? ` — <b>modifiée</b> (auto ${fmt(auto)})` : ''}<br>Solde GTP après : <b class="${after < 0 ? 'neg' : ''}">${fmt(after)}</b>`;
        };
        $('[name=vamt]').addEventListener('input', upd);
        $('[name=vcom]').addEventListener('input', () => { st.commEdited = true; upd(); });
        $('#vcard').onchange = e => { st.cardId = Number(e.target.value); st.gtpId = null; drawFields(); };
        $('#vgtp').onchange = e => { st.gtpId = Number(e.target.value); S.rc.gtpId = st.gtpId; store('rc-gtp', st.gtpId); upd(); };
        $('#vnewcard').onclick = addCard;
        upd();
      } else {
        const p = S.products.find(x => x.id === Number(st.sel));
        const elecs = S.accounts.filter(a => a.active && myAcc(a) && (a.kind === 'electronique' || a.kind === 'banque'));
        const cnt = counterFor(p, elecs);
        box.innerHTML = `
          ${field(p.direction === 'entree' ? 'Montant' : 'Montant remis au client', moneyInput('vamt', 'autofocus'))}
          ${p.counter_account_id ? field('Compte électronique utilisé', `<select class="input" id="vcounter">${cnt ? '' : '<option value="">— Choisir le compte —</option>'}${options(elecs.map(a => [a.id, a.name]), cnt || '')}</select>`) : ''}
          ${field('Détails (facultatif)', `<input class="input" id="vdet" autocomplete="off" placeholder="N° de téléphone, référence…" value="${p.counter_account_id && st.client && st.client.phone ? esc(st.client.phone) : ''}">`)}
          ${addBtn}`;
        bindMoney(box);
      }
      box.querySelectorAll('input').forEach(i => i.addEventListener('keydown', e => { if (e.key === 'Enter' && !e.ctrlKey && !e.metaKey) { e.preventDefault(); addLine(); } }));
      const ab = $('#vadd'); if (ab) ab.onclick = () => addLine();
      const f = box.querySelector('.autofocus'); if (f) setTimeout(() => f.focus(), 30);
    }
    function addCard() {
      sheetCard(null, { client: st.client, onSaved: card => {
        st.client.cards = [...(st.client.cards || []), card]; st.cardId = card.id; st.gtpId = null; st.sel = 'R'; drawClient(); drawFields();
      } });
    }
    // Ajoute la ligne en cours. Renvoie false si elle est incomplète.
    function addLine(silent) {
      if (!st.sel) return true;
      const amt = parseAmt(($('[name=vamt]') || {}).value);
      if (st.sel === 'R') {
        if (!$('[name=vamt]')) { if (!silent) toast('Choisissez le client et sa carte.', 'bad'); return false; }
        if (!amt || amt <= 0) { if (!silent) toast('Indiquez le montant rechargé.', 'bad'); return false; }
        const card = st.client.cards.find(k => k.id === st.cardId), g = acc(st.gtpId);
        const com = parseAmt($('[name=vcom]').value) ?? commission(g, amt);
        st.lines.push({ type: 'recharge', card_id: card.id, gtp_id: g.id, amount: amt, commission: com, sign: 1,
          label: `Recharge carte ${card.bank} ${card.card_number}`, sub: `${g.name} · commission ${fmt(com)}` });
      } else {
        if (!amt || amt <= 0) { if (!silent) toast('Indiquez le montant.', 'bad'); return false; }
        const p = S.products.find(x => x.id === Number(st.sel));
        const cnt = $('#vcounter') ? Number($('#vcounter').value) : null;
        if ($('#vcounter') && !cnt) { if (!silent) toast('Choisissez le compte électronique utilisé.', 'bad'); return false; }
        const det = $('#vdet').value.trim();
        st.lines.push({ type: 'caisse', product_id: p.id, amount: amt, details: det, counter_account_id: cnt, sign: p.direction === 'entree' ? 1 : -1,
          label: p.name, sub: [det, cnt ? (acc(cnt) || {}).name : ''].filter(Boolean).join(' · ') });
      }
      st.sel = ''; drawFields(); drawLines();
      $('#vprod').focus();
      return true;
    }
    const total = () => st.lines.reduce((s, l) => s + l.sign * l.amount, 0);
    function drawLines() {
      $('#vlines').innerHTML = st.lines.length ? st.lines.map((l, i) => `<div class="vline">
          <div class="main"><div class="t">${esc(l.label)}</div>${l.sub ? `<div class="s">${esc(l.sub)}</div>` : ''}</div>
          <div class="num ${l.sign < 0 ? 'neg' : ''}" style="font-weight:650">${l.sign < 0 ? '−' : ''}${fmt(l.amount)}</div>
          <button class="x" type="button" data-rm="${i}" aria-label="Retirer">×</button></div>`).join('')
        : '<div class="empty">Choisissez un produit, indiquez le montant, puis « Ajouter ».</div>';
      c.qa('[data-rm]').forEach(b => b.onclick = () => { st.lines.splice(Number(b.dataset.rm), 1); drawLines(); });
      const t = total();
      $('#vtot').innerHTML = st.lines.length ? `<div class="vtotal"><span>${t >= 0 ? 'Total à payer' : 'Total à remettre au client'}</span><span class="num">${fmt(Math.abs(t))}</span></div>` : '';
      const showPay = st.lines.some(l => l.sign > 0) && t > 0;
      $('#vpay').innerHTML = showPay ? `
        <div class="field"><span>Payé en</span><div class="chips wrap">${pays.map(p => `<button type="button" class="chip ${p[0] === st.pay ? 'on' : ''}" data-chip="vpay" data-v="${p[0]}">${esc(p[1])}</button>`).join('')}</div></div>
        <div class="grid2 ${st.pay === S.me.caisse_id ? '' : 'hidden'}" id="vcash">
          ${field('Reçu du client', moneyInput('vrec'))}
          <div class="field"><span>Monnaie à rendre</span><div class="change num" id="vchange">—</div></div></div>` : '';
      if (showPay) {
        bindMoney($('#vpay'));
        chipGroup(c, 'vpay', v => { st.pay = Number(v); $('#vcash').classList.toggle('hidden', st.pay !== S.me.caisse_id); });
        const r = $('[name=vrec]');
        r.addEventListener('input', () => {
          const v = parseAmt(r.value), ch = $('#vchange');
          ch.textContent = v == null ? '—' : v < t ? `Manque ${fmt(t - v)}` : fmt(v - t);
          ch.className = 'change num ' + (v != null && v < t ? 'neg' : 'pos');
        });
        r.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); $('[data-x=ok]').click(); } });
      }
    }

    $('#vprod').onchange = e => { st.sel = e.target.value; drawFields(); };
    $('[data-x=ok]').onclick = e => busy(e.currentTarget, async () => {
      if (st.sel && $('[name=vamt]') && parseAmt($('[name=vamt]').value)) { if (!addLine()) return; }
      if (!st.lines.length) throw new Error('Ajoutez au moins un produit.');
      const t = total();
      const rec = $('[name=vrec]') && st.pay === S.me.caisse_id ? parseAmt($('[name=vrec]').value) : null;
      if (rec != null && t > 0 && rec < t) throw new Error(`Le client a donné ${fmt(rec)} : il manque ${fmt(t - rec)}.`);
      const payload = { client_id: st.client ? st.client.id : null, client_name: st.client ? null : (st.passName || null),
        pay_account_id: st.pay, received: rec,
        lines: st.lines.map(l => l.type === 'recharge'
          ? { type: 'recharge', card_id: l.card_id, gtp_id: l.gtp_id, amount: l.amount, commission: l.commission }
          : { type: 'caisse', product_id: l.product_id, amount: l.amount, details: l.details, counter_account_id: l.counter_account_id }) };
      for (let i = 0; i < 3; i++) {
        const r = await rpc('op_vente', payload);
        if (r.ok) {
          const client = st.client;
          c.close();
          toast(`Opération enregistrée : ${fmt(Math.abs(r.total))}.`, 'ok');
          refreshNow();
          sheetReceipt(r.receipt_id, { client });
          return;
        }
        const dups = r.warnings.filter(w => w.type === 'duplicate'), ins = r.warnings.filter(w => w.type === 'insufficient');
        if (ins.length) {
          const msg = ins.map(w => `${w.account} : solde ${fmt(w.balance)}, il manquerait ${fmt(-w.after)}`).join('<br>');
          if (!r.can_force_balance) throw new Error(`Solde GTP insuffisant (${ins.map(w => w.account).join(', ')}). Prévenez la gérante pour un approvisionnement.`);
          if (!await ask('Solde GTP insuffisant', `<p>${msg}</p><p>Enregistrer quand même ? Le solde deviendra négatif et l’opération sera signalée.</p>`, 'Enregistrer quand même', true)) return;
          payload.force_balance = true;
        }
        if (dups.length) {
          const msg = dups.map(w => `<p>Une recharge de <b>${fmt(w.amount)}</b> sur la carte <b>${esc(w.card)}</b> (${esc(w.holder)}) a déjà été enregistrée ${esc(ago(w.created_at))} par ${esc(w.agent || '?')}.</p>`).join('');
          if (!await ask('Doublon possible', msg + '<p>S’agit-il bien d’une nouvelle recharge ?</p>', 'Oui, c’est une nouvelle recharge')) return;
          payload.force_duplicate = true;
        }
      }
    });

    drawClient(); drawFields(); drawLines();
    if (preset.productId) { st.sel = String(preset.productId); drawFields(); }
    if (!st.client && $('#vcq')) $('#vcq').focus(); else if (!st.sel) $('#vprod').focus();
  }

  // ---------------------------------------------------------------- reçu
  const pad = n => String(n).padStart(6, '0');
  function company() { return Object.assign({ name: 'AYAEL ENTREPRISE', address: '', phone: '', title: 'REÇU', footer: '', format: '80mm' }, S.settings.company || {}); }
  function receiptLine(o) {
    const label = o.kind === 'recharge' ? `Recharge carte N° ${o.card_number}` : (o.product || KIND[o.kind]);
    const sub = o.kind === 'recharge' ? (o.holder_name && o.holder_name !== o.client_name ? o.holder_name : '') : (o.details || '');
    const neg = o.kind === 'caisse' && o.direction === 'sortie';
    return `<div class="${o.status !== 'ok' ? 'x' : ''}"><div class="l"><span>${esc(label)}</span><span>${neg ? '-' : ''}${fmt(o.amount)}</span></div>
      ${sub ? `<div style="font-size:11.5px">${esc(sub)}</div>` : ''}${o.status !== 'ok' ? '<div>*** ANNULÉ ***</div>' : ''}</div>`;
  }
  function receiptHtml(r) {
    const co = company();
    const ok = r.lines.filter(o => o.status === 'ok');
    const tot = ok.reduce((s, o) => s + (o.kind === 'caisse' && o.direction === 'sortie' ? -1 : 1) * Number(o.amount), 0);
    const payName = r.pay_kind === 'caisse' ? 'Espèces' : (r.pay_kind === 'banque' ? 'Versement bancaire' : r.pay_account);
    return `<div class="c"><img src="icons/logo.png" alt="">
        <div class="co">${esc(co.name)}</div>${co.address ? `<div>${esc(co.address)}</div>` : ''}${co.phone ? `<div>Tél : ${esc(co.phone)}</div>` : ''}</div>
      <hr><div class="c title">${esc(co.title || 'REÇU')} N° ${pad(r.id)}</div>
      <div class="c">${fmtDT(r.created_at)}</div>
      ${r.client_name ? `<div>Client : ${esc(r.client_name)}${r.client_phone ? ` (${esc(r.client_phone)})` : ''}</div>` : ''}
      <div>Servi par : ${esc(r.agent || '')}</div>
      <hr>${r.lines.map(receiptLine).join('')}<hr>
      <div class="l big"><span>${tot >= 0 ? 'TOTAL' : 'TOTAL REMIS'}</span><span>${fmt(Math.abs(tot))}</span></div>
      ${tot > 0 ? `<div class="l"><span>Payé en</span><span>${esc(payName || '')}</span></div>` : ''}
      ${r.received != null && tot > 0 ? `<div class="l"><span>Reçu</span><span>${fmt(r.received)}</span></div><div class="l"><span>Rendu</span><span>${fmt(r.received - tot)}</span></div>` : ''}
      <hr>${co.footer ? `<div class="c">${esc(co.footer)}</div>` : ''}`;
  }
  function printReceipt(r) {
    const co = company();
    const root = document.getElementById('print-root');
    const fmtCls = co.format === 'A4' ? 'a4' : co.format === '58mm' ? 'w58' : '';
    root.className = fmtCls;
    root.innerHTML = `<div class="ticket" style="${co.format === '58mm' ? 'font-size:11px' : ''}">${receiptHtml(r)}</div>`;
    let st = document.getElementById('pagestyle');
    if (!st) { st = document.createElement('style'); st.id = 'pagestyle'; document.head.append(st); }
    st.textContent = co.format === 'A4' ? '@page { size: A4; margin: 15mm; }' : `@page { size: ${co.format === '58mm' ? '58mm' : '80mm'} auto; margin: 3mm 0; }`;
    const img = root.querySelector('img');
    const go = () => setTimeout(() => window.print(), 50);
    if (img && !img.complete) { img.onload = go; img.onerror = go; } else go();
  }
  function whatsappReceipt(r) {
    const co = company();
    const ok = r.lines.filter(o => o.status === 'ok');
    const tot = ok.reduce((s, o) => s + (o.kind === 'caisse' && o.direction === 'sortie' ? -1 : 1) * Number(o.amount), 0);
    const txt = [`*${co.name}*`, `${co.title || 'Reçu'} N° ${pad(r.id)} — ${fmtDT(r.created_at)}`, r.client_name ? `Client : ${r.client_name}` : '', '',
      ...ok.map(o => `• ${o.kind === 'recharge' ? 'Recharge carte N° ' + o.card_number : o.product}${o.details && o.kind === 'caisse' ? ' (' + o.details + ')' : ''} : ${fmt(o.amount)}`),
      '', `*Total : ${fmt(Math.abs(tot))}*`, co.footer || ''].filter((x, i, a) => x !== '' || a[i - 1] !== '').join('\n').replace(/ /g, ' ');
    let ph = String(r.client_phone || '').replace(/\D/g, '');
    if (ph.length === 8) ph = '228' + ph;
    window.open(`https://wa.me/${ph}?text=${encodeURIComponent(txt)}`, '_blank', 'noopener');
  }
  async function sheetReceipt(id, opts = {}) {
    const r = await rpc('get_receipt', { id });
    sheet({
      title: `Reçu N° ${pad(r.id)}`,
      body: `<div class="ticket">${receiptHtml(r)}</div>`,
      foot: `${opts.client !== undefined ? '<button class="btn" data-x="again">Nouvelle opération</button>' : ''}
        <button class="btn" data-x="wa">${icon('chat')} WhatsApp</button>
        <button class="btn primary" data-x="print">${icon('print')} Imprimer <span class="kbd only-desktop">P</span></button>`,
      onMount(c) {
        c.q('[data-x=print]').onclick = () => printReceipt(r);
        c.q('[data-x=wa]').onclick = () => whatsappReceipt(r);
        const ag = c.q('[data-x=again]');
        if (ag) ag.onclick = () => { c.close(); sheetVente(opts.client ? { client: opts.client } : {}); };
        c.onKey = e => {
          if ((e.key === 'p' || e.key === 'P') && !/^(INPUT|TEXTAREA)$/.test(e.target.tagName)) { e.preventDefault(); printReceipt(r); }
          if ((e.key === 'n' || e.key === 'N') && ag) { e.preventDefault(); ag.click(); }
        };
      },
    });
  }

  // ---------------------------------------------------------------- clients
  async function vClients(m) {
    m.innerHTML = `
      <div class="search">${icon('search')}<input class="input" id="clq" placeholder="Nom, téléphone ou n° de carte" value="${esc(S.cl.q)}" autocomplete="off"></div>
      <div id="clres"><div class="spinner"></div></div>
      <button class="btn block" id="clnew" style="margin-top:12px">${icon('plus')} Nouveau client</button>`;
    const load = async () => {
      const res = await rpc('search_clients', { q: S.cl.q, limit: 60 });
      const box = m.querySelector('#clres'); if (!box) return;
      res.forEach(x => S.clients.set(x.id, x));
      box.innerHTML = res.length ? `<div class="list">${res.map(x => `<button class="row" data-client="${x.id}">
          <div class="ico rc">${esc(initials(x.name))}</div>
          <div class="main"><div class="t">${esc(x.name)}</div>
            <div class="s">${esc([x.phone, activeCards(x).map(k => `${k.bank} ${k.card_number}`).join(', ') || 'pas de carte active'].filter(Boolean).join(' · '))}${x.last_at ? ' · dernière visite ' + ago(x.last_at) : ''}</div></div>${icon('chev')}</button>`).join('')}</div>
          ${res.length >= 60 ? '<p class="small muted">60 premiers résultats : précisez la recherche.</p>' : ''}`
        : '<div class="list"><div class="empty">Aucun client trouvé.</div></div>';
    };
    const q = m.querySelector('#clq');
    q.addEventListener('input', debounce(() => { S.cl.q = q.value.trim(); load(); }, 250));
    m.querySelector('#clnew').onclick = () => sheetClientEdit({ name: S.cl.q }, saved => sheetClient(saved));
    await load();
  }

  async function sheetClient(cl) {
    const hist = await rpc('list_operations', { client_id: cl.id, from: '2000-01-01', limit: 40 });
    sheet({
      title: cl.name, wide: true,
      body: `<div class="cols"><div>
          <dl class="kv"><dt>Téléphone</dt><dd>${esc(cl.phone || '—')}</dd>${cl.note ? `<dt>Note</dt><dd>${esc(cl.note)}</dd>` : ''}<dt>Opérations</dt><dd>${hist.filter(o => o.status === 'ok').length}${hist.length >= 40 ? '+' : ''}</dd></dl>
          ${section('Cartes', `<div class="list">${(cl.cards || []).map(k => `<div class="row"><div class="ico rc">${esc(k.bank.slice(0, 3))}</div><div class="main"><div class="t" style="${k.active === false ? 'text-decoration:line-through;color:var(--muted)' : ''}">${esc(k.card_number)}${k.active === false ? ' <span class="tag">inactive</span>' : ''}</div><div class="s">${esc(k.bank)}${k.holder_name !== cl.name ? ' · titulaire ' + esc(k.holder_name) : ''}${k.note ? ' · ' + esc(k.note) : ''}</div></div></div>`).join('') || '<div class="empty">Aucune carte.</div>'}</div>
            <button class="btn block small" data-x="card" style="margin-top:8px">${icon('plus')} Ajouter une carte</button>`)}
        </div><div>${section('Historique', opList(hist, isMgr(), 'Aucune opération.'))}</div></div>`,
      foot: `<button class="btn" data-x="edit">Modifier la fiche</button><button class="btn primary" data-x="op">${icon('plus')} Nouvelle opération</button>`,
      onMount(c) {
        c.q('[data-x=edit]').onclick = () => { c.close(); sheetClientEdit(cl, saved => sheetClient(saved)); };
        c.q('[data-x=op]').onclick = () => { c.close(); sheetVente({ client: cl }); };
        c.q('[data-x=card]').onclick = () => { c.close(); sheetCard(null, { client: cl, onSaved: async () => { const r = await rpc('search_clients', { q: cl.phone || cl.name, limit: 5 }); sheetClient(r.find(x => x.id === cl.id) || cl); } }); };
      },
    });
  }

  function sheetClientEdit(cl, then) {
    const isNew = !cl.id;
    sheet({
      title: isNew ? 'Nouveau client' : 'Modifier le client',
      body: `${field('Nom complet', `<input class="input autofocus" name="name" autocomplete="off" style="text-transform:uppercase" value="${esc(cl.name || '')}">`)}
        ${field('Téléphone', `<input class="input" name="phone" inputmode="tel" autocomplete="off" value="${esc(cl.phone || '')}" placeholder="ex. 90 11 22 33">`, 'Sert à retrouver le client et à lui envoyer son reçu par WhatsApp.')}
        ${field('Note (facultatif)', `<input class="input" name="note" autocomplete="off" value="${esc(cl.note || '')}">`)}`,
      foot: '<button class="btn primary" data-x="ok">Enregistrer</button>',
      onMount(c) {
        const save = e => busy(c.q('[data-x=ok]'), async () => {
          const r = await rpc('save_client', { id: cl.id || null, name: c.q('[name=name]').value, phone: c.q('[name=phone]').value, note: c.q('[name=note]').value });
          c.close(); toast(isNew ? 'Client créé.' : 'Fiche enregistrée.', 'ok');
          S.clients.set(r.id, r);
          if (then) then(r); else refreshNow();
        });
        c.q('[data-x=ok]').onclick = save;
        c.onKey = e => { if (e.key === 'Enter' && e.target.tagName === 'INPUT') { e.preventDefault(); save(); } };
      },
    });
  }

  function sheetCompany() {
    const co = company();
    sheet({
      title: 'Reçu et coordonnées',
      body: `${field('Nom de l’entreprise', `<input class="input" name="name" value="${esc(co.name)}">`)}
        ${field('Adresse', `<input class="input" name="address" value="${esc(co.address)}" placeholder="Quartier, ville">`)}
        ${field('Téléphone', `<input class="input" name="phone" value="${esc(co.phone)}">`)}
        ${field('Titre du document', `<input class="input" name="title" value="${esc(co.title)}">`, 'ex. REÇU, TICKET, FACTURE')}
        ${field('Message en bas du reçu', `<input class="input" name="footer" value="${esc(co.footer)}">`)}
        ${field('Format d’impression', `<select class="input" name="format">${options([['80mm', 'Ticket 80 mm (imprimante de caisse)'], ['58mm', 'Ticket 58 mm (petite imprimante)'], ['A4', 'Page A4 (imprimante de bureau)']], co.format)}</select>`)}`,
      foot: '<button class="btn primary" data-x="ok">Enregistrer</button>',
      onMount(c) {
        c.q('[data-x=ok]').onclick = e => busy(e.currentTarget, async () => {
          const v = {}; ['name', 'address', 'phone', 'title', 'footer', 'format'].forEach(k => { v[k] = c.q(`[name=${k}]`).value.trim(); });
          await rpc('save_settings', { key: 'company', value: v });
          S.settings.company = v; c.close(); toast('Réglages du reçu enregistrés.', 'ok');
        });
      },
    });
  }

  // Fiche carte. opts : { client, onSaved(card), onRecharge(card) }
  function sheetCard(card, opts = {}) {
    const banks = [...new Set(S.accounts.filter(a => a.kind === 'gtp').map(a => a.bank).filter(Boolean))];
    const g = acc(S.rc.gtpId);
    const c0 = card || { bank: g ? g.bank : banks[0], holder_name: opts.client ? opts.client.name : '', phone: opts.client ? opts.client.phone : '' };
    const then = opts.onRecharge;
    sheet({
      title: card ? 'Carte client' : 'Nouvelle carte client',
      body: `${field('N° de carte', `<input class="input ${card ? '' : 'autofocus'}" name="card_number" inputmode="numeric" autocomplete="off" value="${esc(c0.card_number || '')}">`)}
        ${field('Nom du titulaire', `<input class="input" name="holder_name" autocomplete="off" style="text-transform:uppercase" value="${esc(c0.holder_name || '')}">`)}
        ${field('Banque', `<select class="input" name="bank">${options(banks.map(b => [b, b]), c0.bank)}</select>`)}
        ${field('Téléphone (facultatif)', `<input class="input" name="phone" inputmode="tel" autocomplete="off" value="${esc(c0.phone || '')}">`)}
        ${field('Note (facultatif)', `<input class="input" name="note" autocomplete="off" value="${esc(c0.note || '')}" placeholder="ex. 4 derniers chiffres">`)}
        ${card && isDir() ? `<label class="check"><input type="checkbox" name="active" ${card.active ? 'checked' : ''}> Carte active</label>` : ''}
        ${card && card.n_recharges != null ? `<p class="small muted">${card.n_recharges} recharge(s) enregistrée(s) dans l’application.</p>` : ''}`,
      foot: `${card && then ? '<button class="btn" data-x="rc">Recharger</button>' : ''}<button class="btn primary" data-x="ok">Enregistrer</button>`,
      onMount(c) {
        const rcb = c.q('[data-x=rc]');
        if (rcb) rcb.onclick = () => { c.close(); then(card); };
        c.q('[data-x=ok]').onclick = e => busy(e.currentTarget, async () => {
          const p = { id: card && card.id, card_number: c.q('[name=card_number]').value, holder_name: c.q('[name=holder_name]').value,
            bank: c.q('[name=bank]').value, phone: c.q('[name=phone]').value, note: c.q('[name=note]').value,
            client_id: opts.client ? opts.client.id : (card ? card.client_id : null) };
          if (c.q('[name=active]')) p.active = c.q('[name=active]').checked;
          const r = await rpc('save_card', p);
          c.close(); toast('Carte enregistrée.', 'ok');
          if (opts.onSaved) opts.onSaved(r); else refreshNow();
        });
      },
    });
  }

  // ---------------------------------------------------------------- journal
  async function vJournal(m) {
    const f = S.jf;
    const agents = isMgr() ? S.team.map(t => [t.id, t.full_name]) : [];
    m.innerHTML = `
      <div class="chips" id="per">${PERIODS.map(([k, l]) => `<button class="chip ${f.period === k ? 'on' : ''}" data-p="${k}">${l}</button>`).join('')}</div>
      <div class="filters ${f.period === 'custom' ? '' : 'hidden'}" id="dates">
        <input class="input" type="date" id="jfrom" value="${esc(f.from || todayIso())}" aria-label="Du">
        <input class="input" type="date" id="jto" value="${esc(f.to || todayIso())}" aria-label="Au"></div>
      <div class="filters">
        <select class="input" id="jkind" aria-label="Type">${options(Object.entries(KIND), f.kind, 'Tous les types')}</select>
        ${isMgr() ? `<select class="input" id="juser" aria-label="Agent">${options(agents, f.user, 'Toute l’équipe')}</select>` : `<span></span>`}
      </div>
      ${isMgr() ? `<div class="filters"><select class="input" id="jacc" aria-label="Compte">${options(S.accounts.map(a => [a.id, a.name]), f.account, 'Tous les comptes')}</select><span></span></div>` : ''}
      <div class="search">${icon('search')}<input class="input" id="jq" placeholder="Chercher : client, carte, détail…" value="${esc(f.q)}" autocomplete="off"></div>
      <div id="jres"><div class="spinner"></div></div>`;
    m.querySelectorAll('[data-p]').forEach(b => b.onclick = () => {
      f.period = b.dataset.p;
      m.querySelectorAll('[data-p]').forEach(x => x.classList.toggle('on', x === b));
      m.querySelector('#dates').classList.toggle('hidden', f.period !== 'custom');
      loadJournal();
    });
    const on = (id, k) => { const el = m.querySelector(id); if (el) el.onchange = () => { f[k] = el.value; loadJournal(); }; };
    on('#jfrom', 'from'); on('#jto', 'to'); on('#jkind', 'kind'); on('#juser', 'user'); on('#jacc', 'account');
    const q = m.querySelector('#jq');
    q.addEventListener('input', debounce(() => { f.q = q.value.trim(); loadJournal(); }, 300));
    await loadJournal();
  }

  function journalRange() {
    const f = S.jf;
    if (f.period === 'custom') return [f.from || todayIso(), f.to || f.from || todayIso()];
    return period(f.period);
  }

  async function loadJournal() {
    const box = document.getElementById('jres');
    if (!box) return;
    const f = S.jf, [from, to] = journalRange();
    const list = await rpc('list_operations', { from, to, kind: f.kind || null, user_id: f.user || null, account_id: f.account || null, q: f.q || null, limit: 1000 });
    if (!document.getElementById('jres')) return;
    const ok = list.filter(o => o.status === 'ok');
    const rech = ok.filter(o => o.kind === 'recharge');
    const sumIn = ok.reduce((s, o) => s + (opInfo(o).sign > 0 ? Number(o.amount) : 0), 0);
    const sumOut = ok.reduce((s, o) => s + (opInfo(o).sign < 0 ? Math.abs(Number(o.amount)) : 0), 0);
    const byDay = {};
    list.forEach(o => { (byDay[o.op_day] = byDay[o.op_day] || []).push(o); });
    box.innerHTML = `
      <div class="kpis">
        <div class="kpi"><div class="label">Opérations</div><div class="value num">${ok.length}</div><div class="sub">${list.length - ok.length} annulée(s)</div></div>
        <div class="kpi"><div class="label">Entrées</div><div class="value num pos">${fmt(sumIn)}</div></div>
        <div class="kpi"><div class="label">Sorties</div><div class="value num neg">${fmt(sumOut)}</div></div>
        <div class="kpi"><div class="label">Recharges</div><div class="value num">${fmt(rech.reduce((s, o) => s + Number(o.amount), 0))}</div><div class="sub">${rech.length} · com. ${fmt(rech.reduce((s, o) => s + Number(o.commission), 0))}</div></div>
      </div>
      ${isMgr() ? `<button class="btn block small" id="jexp" style="margin-bottom:8px">${icon('download')} Exporter cette liste en Excel</button>` : ''}
      ${list.length ? `<div class="only-mobile">${Object.keys(byDay).sort().reverse().map(day => `<div class="daytitle"><span>${fmtDate(day)}</span><span>${byDay[day].length} op.</span></div>${opList(byDay[day], isMgr())}`).join('')}</div>
        <div class="only-desktop">${opTable(list)}</div>`
        : '<div class="list"><div class="empty">Aucune opération sur cette période.</div></div>'}
      ${list.length >= 1000 ? '<p class="small muted">Seules les 1 000 dernières opérations sont affichées. Réduisez la période ou utilisez l’export.</p>' : ''}`;
    const ex = box.querySelector('#jexp');
    if (ex) ex.onclick = () => busy(ex, () => exportExcel(from, to, { kind: f.kind || null, user_id: f.user || null, account_id: f.account || null, q: f.q || null }));
  }

  // Journal sous forme de tableau (ordinateur)
  function opTable(list) {
    const multiDay = new Set(list.map(o => o.op_day)).size > 1;
    return `<div class="dtable scrollx"><table><thead><tr>
        <th>${multiDay ? 'Date' : 'Heure'}</th><th>Type</th><th>Client / produit</th><th>Détails</th><th>Compte</th>${isMgr() ? '<th>Agent</th>' : ''}<th class="r">Montant</th><th class="r">Com.</th><th>Reçu</th></tr></thead><tbody>
      ${list.map(o => {
        S.ops.set(o.id, o);
        const i = opInfo(o);
        const amt = o.kind === 'ajustement' ? Math.abs(o.amount) : o.amount;
        const cls = o.status !== 'ok' ? '' : i.sign > 0 ? 'pos' : i.sign < 0 ? 'neg' : '';
        const who = o.kind === 'recharge' ? (o.client_name || o.holder_name) : (o.client_name || '');
        const what = o.kind === 'recharge' ? `Recharge ${o.card_number}` : o.kind === 'caisse' ? o.product : i.title;
        return `<tr class="${o.status !== 'ok' ? 'cancelled' : ''}" data-op="${o.id}">
          <td class="num">${multiDay ? fmtDate(o.op_day) + ' ' : ''}${fmtTime(o.created_at)}</td>
          <td>${esc(KIND[o.kind])}${o.status !== 'ok' ? ' <span class="tag red">Annulée</span>' : o.flags ? ' <span class="tag amber">À vérifier</span>' : ''}</td>
          <td><div style="font-weight:600">${esc(who || what)}</div>${who ? `<div class="sub">${esc(what)}</div>` : ''}</td>
          <td class="sub">${esc(o.details || '')}</td>
          <td class="sub">${esc([o.account, o.account2 && !/^Caisse/.test(o.account2) ? o.account2 : ''].filter(Boolean).join(' → '))}</td>
          ${isMgr() ? `<td class="sub">${esc(o.agent || '')}</td>` : ''}
          <td class="r num amt ${cls}" style="font-weight:650">${i.sign > 0 ? '+' : i.sign < 0 ? '−' : ''}${fmt(amt)}</td>
          <td class="r num sub">${o.kind === 'recharge' ? fmtN(o.commission) : ''}</td>
          <td class="sub">${o.receipt_id ? 'N° ' + pad(o.receipt_id) : ''}</td></tr>`;
      }).join('')}</tbody></table></div>`;
  }

  function opDetail(id) {
    const o = S.ops.get(id);
    if (!o) return;
    const i = opInfo(o);
    const age = (Date.now() - new Date(o.created_at).getTime()) / 60000;
    const canCancel = o.status === 'ok' && (isDir() || (isMgr() && o.agence_id === S.me.agence_id)
      || (o.user_id === S.me.id && age < 30 && !o.closing_id && ['caisse', 'recharge', 'depense'].includes(o.kind)));
    const label2 = { caisse: 'Compte électronique', recharge: 'Payé en', approvisionnement: 'Source', transfert: 'Vers', cloture: 'Vers' }[o.kind] || 'Compte 2';
    const rows = [
      ['Type', KIND[o.kind]], ['Date', fmtDT(o.created_at)], ['Saisi par', o.agent],
      o.client_name && ['Client', o.client_name],
      o.product && ['Produit', o.product], o.card_number && ['Carte', `${o.card_number} — ${o.holder_name}`],
      ['Compte', o.account], o.account2 && [label2, o.kind === 'recharge' && o.account_kind !== 'caisse' && /^Caisse/.test(o.account2) ? 'Espèces' : o.account2],
      ['Montant', fmt(o.kind === 'ajustement' ? o.amount : o.amount)],
      o.received != null && ['Reçu du client', fmt(o.received)], o.change_given != null && ['Monnaie rendue', fmt(o.change_given)],
      o.kind === 'recharge' && ['Commission', fmt(o.commission) + (o.commission_status === 'verifiee' ? ' (vérifiée)' : ' (estimée, à vérifier)')
        + (o.commission_before != null && Number(o.commission_before) !== Number(o.commission) ? ` · avant correction : ${fmt(o.commission_before)}` : '')],
      o.category && ['Catégorie', o.category], o.details && ['Détails', o.details], o.flags && ['Signalement', o.flags],
      o.status !== 'ok' && ['Annulée', `${fmtDT(o.cancelled_at)} par ${o.cancelled_by_name || '?'}`], o.cancel_reason && ['Motif', o.cancel_reason],
      ['N°', '#' + o.id],
    ].filter(Boolean);
    sheet({
      title: i.title,
      body: `<dl class="kv">${rows.map(([k, v]) => `<dt>${esc(k)}</dt><dd>${esc(v ?? '')}</dd>`).join('')}</dl>
        ${o.status === 'ok' && !canCancel && !isMgr() ? '<p class="small muted" style="margin-top:16px">Pour annuler cette opération, demandez à votre gérant(e).</p>' : ''}`,
      foot: (canCancel ? '<button class="btn danger" data-x="cancel">Annuler l’opération</button>' : '') +
        (o.receipt_id ? `<button class="btn" data-x="rec">${icon('receipt')} Voir le reçu N° ${pad(o.receipt_id)}</button>` : '') +
        (canVerify(o) ? '<button class="btn" data-x="com">Vérifier la commission</button>' : ''),
      onMount(c) {
        const cb = c.q('[data-x=com]');
        if (cb) cb.onclick = () => { c.close(); sheetVerifyCommission(o); };
        const rb = c.q('[data-x=rec]');
        if (rb) rb.onclick = () => { c.close(); sheetReceipt(o.receipt_id); };
        const b = c.q('[data-x=cancel]');
        if (b) b.onclick = async () => {
          const reason = await askText('Annuler l’opération', 'Motif de l’annulation (obligatoire)', { okLabel: 'Annuler l’opération', danger: true, placeholder: 'ex. erreur de montant, doublon…' });
          if (!reason) return;
          await busy(b, async () => { await rpc('op_annuler', { id: o.id, reason }); c.close(); toast('Opération annulée (elle reste visible, barrée).', 'ok'); refreshNow(); });
        };
      },
    });
  }

  // ---------------------------------------------------------------- commissions à vérifier
  const canVerify = o => o.kind === 'recharge' && o.status === 'ok'
    && (isDir() || (isMgr() && o.agence_id === S.me.agence_id) || (S.me.can_recharge && o.user_id === S.me.id));

  function sheetVerifyCommission(o) {
    sheet({
      title: 'Vérifier la commission',
      body: `<div class="calc"><div style="font-weight:700">${esc(o.client_name || o.holder_name || '')}</div>
          <div class="small muted">Carte ${esc(o.card_number || '')} · ${esc(o.account || '')} · ${fmtDT(o.created_at)} · montant ${fmt(o.amount)}</div></div>
        ${field('Commission reversée (lue sur le rapport du GTP)', moneyInput('com', `autofocus value="${fmtN(o.commission)}"`),
          `Commission enregistrée : ${fmt(o.commission)}${o.commission_status === 'verifiee' ? ' (déjà vérifiée)' : ' (estimée)'}. Si vous changez le montant, le solde GTP est corrigé de la différence.`)}`,
      foot: '<button class="btn primary" data-x="ok">Confirmer</button>',
      onMount(c) {
        const go = () => busy(c.q('[data-x=ok]'), async () => {
          const v = parseAmt(c.q('[name=com]').value);
          if (v == null || v < 0) throw new Error('Indiquez la commission.');
          const r = await rpc('verify_commission', { id: o.id, commission: v });
          c.close();
          toast(r.diff ? `Commission corrigée (${fmtSigned(r.diff)}). Solde GTP : ${fmt(r.balance)}` : 'Commission confirmée.', 'ok');
          refreshNow();
        });
        c.q('[data-x=ok]').onclick = go;
        c.onKey = e => { if (e.key === 'Enter' && e.target.tagName === 'INPUT') { e.preventDefault(); go(); } };
      },
    });
  }

  async function vCommissions(m) {
    const f = S.cm || (S.cm = { period: 'today', from: null, to: null, account: '', status: 'estimee' });
    const gtps = S.accounts.filter(a => a.kind === 'gtp');
    m.innerHTML = `
      <div class="alert info"><span>À la saisie, la commission est <b>estimée</b>. Ouvrez le rapport du GTP, puis pour chaque recharge tapez la <b>commission reversée</b> et touchez ✓. Si le montant change, le solde GTP est corrigé automatiquement.</span></div>
      <div class="chips">${PERIODS.filter(p => p[0] !== 'year').map(([k, l]) => `<button class="chip ${f.period === k ? 'on' : ''}" data-p="${k}">${l}</button>`).join('')}</div>
      <div class="filters ${f.period === 'custom' ? '' : 'hidden'}" id="cmdates">
        <input class="input" type="date" id="cmfrom" value="${esc(f.from || todayIso())}" aria-label="Du"><input class="input" type="date" id="cmto" value="${esc(f.to || todayIso())}" aria-label="Au"></div>
      <div class="filters">
        <select class="input" id="cmacc" aria-label="Compte GTP">${options(gtps.map(a => [a.id, a.name]), f.account, 'Tous les comptes GTP')}</select>
        <select class="input" id="cmst" aria-label="État">${options([['estimee', 'À vérifier'], ['verifiee', 'Déjà vérifiées'], ['', 'Toutes']], f.status)}</select>
      </div>
      <div id="cmres"><div class="spinner"></div></div>`;
    m.querySelectorAll('[data-p]').forEach(b => b.onclick = () => { f.period = b.dataset.p; route(); });
    const on = (id, k) => { const el = m.querySelector(id); el.onchange = () => { f[k] = el.value; load(); }; };
    on('#cmfrom', 'from'); on('#cmto', 'to'); on('#cmacc', 'account'); on('#cmst', 'status');
    const load = async () => {
      const [from, to] = f.period === 'custom' ? [f.from || todayIso(), f.to || f.from || todayIso()] : period(f.period);
      const list = (await rpc('list_operations', { from, to, kind: 'recharge', account_id: f.account || null,
        commission_status: f.status || null, include_cancelled: false, limit: 2000 })).reverse();   // dans l'ordre du rapport GTP
      const box = m.querySelector('#cmres'); if (!box) return;
      list.forEach(o => S.ops.set(o.id, o));
      const mine = list.filter(canVerify);
      const tot = list.reduce((s, o) => s + Number(o.commission), 0);
      box.innerHTML = list.length ? `
        <p class="small muted" style="margin:0 4px 8px">${list.length} recharge(s) · commissions ${fmt(tot)}</p>
        <div class="dtable scrollx"><table><thead><tr><th>${new Set(list.map(o => o.op_day)).size > 1 ? 'Date' : 'Heure'}</th><th>Client · carte</th><th>Compte</th><th class="r">Montant</th><th class="r">Commission reversée</th><th></th></tr></thead><tbody>
          ${list.map(o => `<tr style="cursor:default" data-row="${o.id}">
            <td class="num">${new Set(list.map(x => x.op_day)).size > 1 ? fmtDate(o.op_day) + ' ' : ''}${fmtTime(o.created_at)}</td>
            <td><div style="font-weight:600">${esc(o.client_name || o.holder_name || '')}</div><div class="sub">${esc(o.card_number || '')}${isMgr() && o.agent ? ' · ' + esc(o.agent) : ''}</div></td>
            <td class="sub">${esc(o.account || '')}</td>
            <td class="r num">${fmtN(o.amount)}</td>
            <td class="r">${canVerify(o) ? `<input class="input money" data-com="${o.id}" inputmode="numeric" value="${fmtN(o.commission)}" style="width:110px;min-height:40px;font-size:16px;text-align:right;padding:6px 10px">` : `<span class="num">${fmtN(o.commission)}</span>`}</td>
            <td>${o.commission_status === 'verifiee' ? '<span class="tag green">vérifiée</span>' : canVerify(o) ? `<button class="btn small primary" data-ok="${o.id}" title="Confirmer">✓</button>` : '<span class="tag amber">à vérifier</span>'}</td></tr>`).join('')}
        </tbody></table></div>
        ${mine.some(o => o.commission_status !== 'verifiee') ? `<button class="btn primary block" id="cmall" style="margin-top:12px">Confirmer toutes les lignes affichées (${mine.filter(o => o.commission_status !== 'verifiee').length})</button>
          <p class="small muted">Vérifiez d’abord chaque montant avec le rapport du GTP : « Confirmer tout » enregistre les montants tels qu’ils sont affichés.</p>` : ''}`
        : `<div class="list"><div class="empty">${f.status === 'estimee' ? 'Aucune commission à vérifier sur cette période. 👍' : 'Aucune recharge sur cette période.'}</div></div>`;
      bindMoney(box);
      const verify = async id => {
        const v = parseAmt(box.querySelector(`[data-com="${id}"]`).value);
        if (v == null || v < 0) throw new Error('Commission invalide.');
        return rpc('verify_commission', { id, commission: v });
      };
      box.querySelectorAll('[data-ok]').forEach(b => b.onclick = () => busy(b, async () => {
        const r = await verify(Number(b.dataset.ok));
        toast(r.diff ? `Corrigée (${fmtSigned(r.diff)}).` : 'Confirmée.', 'ok'); load();
      }));
      box.querySelectorAll('[data-com]').forEach(i => i.addEventListener('keydown', e => {
        if (e.key === 'Enter') { e.preventDefault(); const b = box.querySelector(`[data-ok="${i.dataset.com}"]`); if (b) b.click(); }
      }));
      const all = box.querySelector('#cmall');
      if (all) all.onclick = () => busy(all, async () => {
        const todo = mine.filter(o => o.commission_status !== 'verifiee');
        let n = 0, diff = 0;
        for (const o of todo) { const r = await verify(o.id); n++; diff += Number(r.diff) || 0; all.textContent = `Confirmation… ${n} / ${todo.length}`; }
        toast(`${n} commission(s) confirmée(s)${diff ? `, écart total ${fmtSigned(diff)}` : ''}.`, 'ok'); load();
      });
    };
    await load();
  }

  // ---------------------------------------------------------------- clôture (agent)
  async function vCloture(m) {
    const [pv, hist] = await Promise.all([rpc('closing_preview'), rpc('list_closings', { limit: 15 })]);
    const histHtml = section('Mes clôtures', hist.length ? `<div class="list">${hist.map(closingRow).join('')}</div>` : '<div class="list"><div class="empty">Aucune clôture pour le moment.</div></div>');
    if (pv.pending) {
      m.innerHTML = `<div class="alert info">Votre clôture est envoyée. Elle attend la validation de la gérante.</div>${histHtml}`;
      return;
    }
    const sum = pv.summary || [];
    m.innerHTML = `
      <div class="hero"><div class="label">Attendu dans ma caisse</div><div class="big num">${fmt(pv.expected)}</div>
        <div class="sub">${pv.last && pv.last.id ? 'Depuis la clôture du ' + fmtDT(pv.last.created_at) : 'Depuis le début'}</div></div>
      ${sum.length ? section('Mouvements de ma caisse', `<div class="card scrollx"><table class="tbl"><thead><tr><th>Opération</th><th class="r">Nb</th><th class="r">Entrées</th><th class="r">Sorties</th></tr></thead><tbody>
        ${sum.map(s => `<tr><td>${esc(s.label)}</td><td class="r num">${s.n}</td><td class="r num">${s.in ? fmtN(s.in) : ''}</td><td class="r num">${s.out ? fmtN(s.out) : ''}</td></tr>`).join('')}</tbody></table></div>`) : ''}
      <div class="card pad section">
        ${field('Espèces comptées dans la caisse', moneyInput('counted'))}
        <div id="gap"></div>
        ${field('Montant remis à la gérante', moneyInput('handed'), 'Le reste est gardé comme fond de caisse pour demain.')}
        <div id="keep" class="small muted" style="margin:-6px 0 14px"></div>
        ${field('Note (facultatif)', '<textarea class="input" name="note" placeholder="Explication d’un écart, remarque…"></textarea>')}
        <button class="btn primary block" id="send">Envoyer la clôture à la gérante</button>
      </div>
      ${histHtml}`;
    bindMoney(m);
    const cI = m.querySelector('[name=counted]'), hI = m.querySelector('[name=handed]');
    let handedEdited = false;
    const upd = () => {
      const cnt = parseAmt(cI.value);
      if (!handedEdited) hI.value = cnt == null ? '' : fmtN(cnt);
      const gap = cnt == null ? null : cnt - pv.expected;
      m.querySelector('#gap').innerHTML = gap == null ? '' : gap === 0 ? '<div class="alert ok">Caisse juste.</div>'
        : gap < 0 ? `<div class="alert bad">Manque de ${fmt(-gap)} par rapport à l’attendu.</div>` : `<div class="alert warn">Surplus de ${fmt(gap)} par rapport à l’attendu.</div>`;
      const h = parseAmt(hI.value) || 0;
      m.querySelector('#keep').textContent = cnt == null ? '' : `Fond de caisse gardé : ${fmt(cnt - h)}`;
    };
    cI.addEventListener('input', upd);
    hI.addEventListener('input', () => { handedEdited = true; upd(); });
    m.querySelector('#send').onclick = e => {
      const cnt = parseAmt(cI.value), h = parseAmt(hI.value) || 0;
      if (cnt == null) return toast('Indiquez le montant compté.', 'bad');
      const gap = cnt - pv.expected;
      ask('Envoyer la clôture ?', `<dl class="kv"><dt>Attendu</dt><dd>${fmt(pv.expected)}</dd><dt>Compté</dt><dd>${fmt(cnt)}</dd><dt>Écart</dt><dd class="${gap < 0 ? 'neg' : gap > 0 ? 'warn' : 'pos'}">${fmtSigned(gap)}</dd><dt>Remis à la gérante</dt><dd>${fmt(h)}</dd><dt>Gardé en caisse</dt><dd>${fmt(cnt - h)}</dd></dl>`, 'Envoyer')
        .then(ok => ok && busy(e.target, async () => {
          await rpc('closing_submit', { counted: cnt, handed: h, note: m.querySelector('[name=note]').value });
          toast('Clôture envoyée à la gérante.', 'ok'); route();
        }));
    };
  }

  function closingRow(c) {
    const st = { soumise: '<span class="tag amber">En attente</span>', validee: '<span class="tag green">Validée</span>', rejetee: '<span class="tag red">Refusée</span>' }[c.status];
    return `<div class="row"><div class="ico neu">${icon('lock')}</div>
      <div class="main"><div class="t">${isMgr() ? esc(c.agent) + (isDir() && !S.agence && c.agence && (S.agences || []).length > 1 ? ' (' + esc(c.agence) + ')' : '') + ' · ' : ''}${fmtDT(c.created_at)}</div>
        <div class="s">Compté ${fmt(c.counted)} · écart <span class="${c.gap < 0 ? 'neg' : c.gap > 0 ? 'warn' : ''}">${fmtSigned(c.gap)}</span> · remis ${fmt(c.handed)}${c.manager_note ? ' · ' + esc(c.manager_note) : ''}</div></div>
      <div class="end">${st}</div></div>`;
  }

  // ---------------------------------------------------------------- écrans gérante
  async function vTableau(m) {
    const d = await rpc('dashboard');
    rememberBalances(d);
    const B = d.balances;
    const gtp = B.filter(b => b.kind === 'gtp'), el = B.filter(b => ['electronique', 'banque'].includes(b.kind));
    const cais = B.filter(b => b.kind === 'caisse'), cof = B.filter(b => b.kind === 'coffre');
    const alerts = [];
    B.filter(b => b.balance < 0).forEach(b => alerts.push(`<div class="alert bad"><span><b>${esc(b.name)}</b> est négatif : ${fmt(b.balance)}.</span></div>`));
    B.filter(b => b.balance >= 0 && Number(b.low_alert) > 0 && b.balance < b.low_alert).forEach(b => alerts.push(`<div class="alert warn"><span><b>${esc(b.name)}</b> est bas : ${fmt(b.balance)} (seuil ${fmt(b.low_alert)}).</span></div>`));
    if (d.pending_closings) alerts.push(`<div class="alert info"><span>${d.pending_closings} clôture(s) à valider. <a href="#/clotures">Voir</a></span></div>`);
    if (d.commissions_to_check) alerts.push(`<div class="alert warn"><span>${d.commissions_to_check} commission(s) de recharge à vérifier sur le rapport du GTP. <a href="#/commissions">Vérifier</a></span></div>`);
    const sum = l => l.reduce((s, b) => s + Number(b.balance), 0);
    m.innerHTML = `
      ${alerts.join('')}
      <div class="kpis">
        <div class="kpi"><div class="label">Recharges du jour</div><div class="value num">${fmt(d.today.recharges_amount)}</div><div class="sub">${d.today.recharges_n} recharge(s)</div></div>
        <div class="kpi"><div class="label">Commissions du jour</div><div class="value num">${fmt(d.today.commissions)}</div></div>
        <div class="kpi"><div class="label">Commissions du mois</div><div class="value num">${fmt(d.month_commissions)}</div></div>
        <div class="kpi"><div class="label">Opérations du jour</div><div class="value num">${d.today.ops_n}</div><div class="sub">dépenses ${fmt(d.today.depenses)}</div></div>
      </div>
      <button class="hero-cta only-mobile" data-do="vente">${icon('plus')}Nouvelle opération</button>
      <div class="actions four">
        ${isDir() ? `<button class="action" data-do="appro">${icon('truck')}Approvisionner un compte</button>`
          : `<button class="action" data-do="reversement">${icon('swap')}Reverser à la principale</button>`}
        <button class="action" data-do="transfert">${icon('swap')}Transfert entre comptes</button>
        <button class="action" data-do="ajust">${icon('adjust')}Ajuster un solde</button>
        <button class="action" data-do="depense">${icon('out')}Dépense</button>
      </div>
      ${(d.liaisons || []).length ? section('Entre agences', liaisonList(d.liaisons)) : ''}
      <div class="cols">
        <div>
          ${section('Soldes GTP', balList(gtp, true), `<span class="num">${fmt(sum(gtp))}</span>`)}
          ${section('Comptes électroniques et banque', balList(el, true), `<span class="num">${fmt(sum(el))}</span>`)}
          ${section('Coffre', balList(cof, true))}
        </div>
        <div>
          ${section('Caisses des agents', balList(cais, true), `<span class="num">${fmt(sum(cais))}</span>`)}
          ${d.to_check && d.to_check.length ? section('À vérifier (7 derniers jours)', opList(d.to_check, true)) : ''}
          ${section('En direct', opList(d.recent, true, 'Aucune opération récente.'), '<a href="#/journal" class="small" style="text-transform:none;color:var(--brand-ink)">Tout le journal</a>')}
        </div>
      </div>`;
  }

  const balLabel = a => `${a.name} (${fmt(S.bal.get(a.id) ?? 0)})`;

  // ---------------------------------------------------------------- entre agences
  const mainAgence = () => (S.agences || []).find(g => g.is_main) || {};
  function liaisonText(l) {
    const b = Number(l.balance), n = mainAgence().name || 'Agence principale';
    const main = /^agence\b/i.test(n) ? 'l’' + n.charAt(0).toLowerCase() + n.slice(1) : n;   // « l’agence principale »
    const Main = main.charAt(0).toUpperCase() + main.slice(1);
    if (Math.abs(b) < 1) return { t: `${l.agence} et ${main} sont à jour`, cls: 'pos' };
    return b > 0 ? { t: `${l.agence} doit ${fmt(b)} à ${main}`, cls: 'warn' } : { t: `${Main} doit ${fmt(-b)} à ${l.agence}`, cls: 'neg' };
  }
  const liaisonList = list => `<div class="list">${list.map(l => {
    const x = liaisonText(l);
    return `<button class="row" data-liaison="${l.agence_id}"><div class="ico neu">${icon('swap')}</div>
      <div class="main"><div class="t">${esc(x.t)}</div><div class="s">Recharges, retraits RIA / Western Union, recharges de comptes… · toucher pour le détail</div></div>
      <div class="end"><div class="amt num ${x.cls}">${fmt(Math.abs(l.balance))}</div></div></button>`;
  }).join('')}</div>`;

  async function sheetLiaison(agId) {
    const r = await rpc('liaison_detail', { agence_id: agId });
    const x = liaisonText({ agence: r.agence, balance: r.balance });
    sheet({
      title: `${r.agence} ↔ ${mainAgence().name || 'agence principale'}`, wide: true,
      body: `<div class="alert ${Math.abs(r.balance) < 1 ? 'ok' : 'warn'}"><span><b>${esc(x.t)}</b></span></div>
        <p class="small muted">Montants positifs : ${esc(r.agence)} doit cet argent à la principale (ex. une recharge encaissée à ${esc(r.agence)}). Négatifs : la principale le doit à ${esc(r.agence)} (ex. un retrait RIA payé par ${esc(r.agence)}).</p>
        ${r.lines.length ? `<div class="dtable scrollx"><table><thead><tr><th>Date</th><th>Opération</th><th>Détail</th><th class="r">Montant</th></tr></thead><tbody>
          ${r.lines.map(o => { S.ops.set(o.id, o); const i = opInfo(o); return `<tr data-op="${o.id}">
            <td class="num">${fmtDate(o.op_day)} ${fmtTime(o.created_at)}</td><td>${esc(i.title)}</td><td class="sub">${esc([o.kind === 'recharge' ? 'carte ' + (o.card_number || '') : o.client_name, o.details, o.agent].filter(Boolean).join(' · '))}</td>
            <td class="r num ${o.liaison > 0 ? 'warn' : 'neg'}" style="font-weight:650">${fmtSigned(o.liaison)}</td></tr>`; }).join('')}</tbody></table></div>`
          : '<div class="list"><div class="empty">Aucune opération entre les deux agences.</div></div>'}`,
      foot: `<button class="btn primary" data-x="rev">${icon('swap')} Enregistrer un reversement</button>`,
      onMount(c) { c.q('[data-x=rev]').onclick = () => { c.close(); sheetReversement(agId); }; },
    });
  }

  // Reversement : l'argent remis entre une agence et l'agence principale (ou l'inverse).
  function sheetReversement(agId) {
    const main = mainAgence();
    const ag = isDir() ? (agId || ((S.agences || []).find(g => !g.is_main) || {}).id) : S.me.agence_id;
    if (!ag || ag === main.id) return toast('Aucune agence secondaire.', 'bad');
    const accOf = id => S.accounts.filter(a => a.active && a.agence_id === id && ['coffre', 'caisse', 'electronique', 'banque'].includes(a.kind) && (a.kind !== 'caisse' || isDir() || a.agent_id === S.me.id));
    const mainAccs = S.accounts.filter(a => a.active && (a.agence_id === main.id || (a.agence_id == null && a.kind !== 'gtp')) && (a.kind !== 'caisse' || a.agent_id === S.me.id || isDir()));
    const defCof = list => (list.find(a => a.kind === 'coffre') || list[0] || {}).id;
    sheet({
      title: 'Reversement entre agences',
      body: `${isDir() ? field('Agence', `<select class="input" name="ag">${options((S.agences || []).filter(g => !g.is_main).map(g => [g.id, g.name]), ag)}</select>`) : ''}
        <div class="field"><span>Sens</span><div class="chips wrap">
          <button type="button" class="chip on" data-chip="dir" data-v="up">${esc(agName(ag) || 'Agence')} → ${esc(main.name || 'Principale')}</button>
          ${isDir() ? `<button type="button" class="chip" data-chip="dir" data-v="down">${esc(main.name || 'Principale')} → ${esc(agName(ag) || 'Agence')}</button>` : ''}</div></div>
        <div id="revacc"></div>
        ${field('Montant', moneyInput('amount', 'autofocus'))}
        ${field('Détails', '<input class="input" name="details" autocomplete="off" placeholder="ex. remise du soir, versement Wave…">')}`,
      foot: '<button class="btn primary" data-x="ok">Enregistrer le reversement</button>',
      onMount(c) {
        let cur = ag;
        const draw = () => {
          const up = chipVal(c, 'dir') !== 'down';
          const from = up ? accOf(cur) : mainAccs, to = up ? mainAccs : accOf(cur);
          c.q('#revacc').innerHTML = `<div class="grid2">${field('Argent pris dans', `<select class="input" name="from">${options(from.map(a => [a.id, balLabel(a)]), defCof(from))}</select>`)}
            ${field('Argent remis dans', `<select class="input" name="to">${options(to.map(a => [a.id, a.name]), defCof(to))}</select>`)}</div>`;
        };
        chipGroup(c, 'dir', draw);
        const s = c.q('[name=ag]'); if (s) s.onchange = () => { cur = Number(s.value); draw(); };
        draw();
        c.q('[data-x=ok]').onclick = e => busy(e.currentTarget, async () => {
          const up = chipVal(c, 'dir') !== 'down';
          await rpc('op_transfert', { from_id: Number(c.q('[name=from]').value), to_id: Number(c.q('[name=to]').value), amount: parseAmt(c.q('[name=amount]').value),
            details: `Reversement ${up ? agName(cur) + ' → ' + (main.name || 'principale') : (main.name || 'principale') + ' → ' + agName(cur)}${c.q('[name=details]').value ? ' · ' + c.q('[name=details]').value : ''}` });
          c.close(); toast('Reversement enregistré.', 'ok'); refreshNow();
        });
      },
    });
  }

  // Comptes qu'un gérant peut utiliser : ceux de son agence (+ destination : l'agence principale).
  const ownAgencyAccs = () => S.accounts.filter(a => a.active && (isDir() || a.agence_id === S.me.agence_id));

  // Approvisionner un compte : GTP ou compte électronique (Mixx, Flooz, Tmoney, Wave…).
  function sheetAppro(accId) {
    const label = a => a.agence_id && (S.agences || []).length > 1 ? `${balLabel(a)} · ${agName(a.agence_id)}` : balLabel(a);
    const gtps = S.accounts.filter(a => a.kind === 'gtp' && a.active);
    const elecs = S.accounts.filter(a => a.kind === 'electronique' && a.active);
    const srcs = S.accounts.filter(a => a.active && a.kind !== 'gtp');
    const grp = (t, list) => list.length ? `<optgroup label="${esc(t)}">${list.map(a => `<option value="${a.id}" ${a.id === Number(accId) ? 'selected' : ''}>${esc(label(a))}</option>`).join('')}</optgroup>` : '';
    sheet({
      title: 'Approvisionner un compte',
      body: `${field('Compte à approvisionner', `<select class="input" name="gtp">${grp('Comptes GTP', gtps)}${grp('Comptes électroniques (Mixx, Flooz…)', elecs)}</select>`)}
        ${field('Montant', moneyInput('amount', 'class="input money autofocus"'))}
        ${field('Argent pris dans', `<select class="input" name="src">${options(srcs.map(a => [a.id, label(a)]), '', '— Extérieur (non suivi dans l’application) —')}</select>`, 'Choisissez le coffre, une caisse ou la banque si l’argent sort d’un compte suivi ici. « Extérieur » : l’argent vient d’ailleurs (ex. versement à la banque, achat de crédit Mixx).')}
        ${field('Détails', '<input class="input" name="details" placeholder="ex. versement GTCO du jour, achat de float Mixx…" autocomplete="off">')}`,
      foot: '<button class="btn primary" data-x="ok">Enregistrer</button>',
      onMount(c) {
        c.q('[data-x=ok]').onclick = e => busy(e.currentTarget, async () => {
          const r = await rpc('op_approvisionnement', { gtp_id: Number(c.q('[name=gtp]').value), amount: parseAmt(c.q('[name=amount]').value),
            source_id: c.q('[name=src]').value || null, details: c.q('[name=details]').value });
          c.close(); toast(`Approvisionnement enregistré. Nouveau solde : ${fmt(r.balance)}`, 'ok'); refreshNow();
        });
      },
    });
  }

  function sheetTransfert(fromId) {
    const accs = ownAgencyAccs();
    const dest = isDir() ? accs : S.accounts.filter(a => a.active && (a.agence_id === S.me.agence_id || a.agence_id === mainAgence().id));
    sheet({
      title: 'Transfert entre comptes',
      body: `${field('De', `<select class="input" name="from">${options(accs.map(a => [a.id, balLabel(a)]), fromId || '')}</select>`)}
        ${field('Vers', `<select class="input" name="to">${options(dest.map(a => [a.id, a.agence_id && a.agence_id !== S.me.agence_id ? `${a.name} (${agName(a.agence_id)})` : balLabel(a)]), '')}</select>`)}
        ${field('Montant', moneyInput('amount', 'class="input money autofocus"'))}
        ${field('Détails', '<input class="input" name="details" placeholder="ex. versement en banque, fond de caisse…" autocomplete="off">')}`,
      foot: '<button class="btn primary" data-x="ok">Enregistrer</button>',
      onMount(c) {
        c.q('[data-x=ok]').onclick = e => busy(e.currentTarget, async () => {
          await rpc('op_transfert', { from_id: Number(c.q('[name=from]').value), to_id: Number(c.q('[name=to]').value),
            amount: parseAmt(c.q('[name=amount]').value), details: c.q('[name=details]').value });
          c.close(); toast('Transfert enregistré.', 'ok'); refreshNow();
        });
      },
    });
  }

  function sheetAjust(accId) {
    const accs = S.accounts.filter(a => (a.active || a.id === accId) && (isDir() || a.agence_id === S.me.agence_id));
    sheet({
      title: 'Ajuster un solde',
      body: `<div class="alert warn">Un ajustement corrige un solde sans opération réelle (erreur, solde d’ouverture…). Il reste visible dans le journal.</div>
        ${field('Compte', `<select class="input" name="acc">${options(accs.map(a => [a.id, balLabel(a)]), accId || '')}</select>`)}
        <div class="field"><span>Type d’ajustement</span><div class="chips wrap">
          <button type="button" class="chip on" data-chip="mode" data-v="target">Fixer le solde exact</button>
          <button type="button" class="chip" data-chip="mode" data-v="plus">Ajouter</button>
          <button type="button" class="chip" data-chip="mode" data-v="minus">Retirer</button></div></div>
        <label class="field"><span id="alab">Nouveau solde</span>${moneyInput('amount', 'data-neg="1"')}<small id="ahelp">Saisissez « - » devant le montant pour un solde négatif.</small></label>
        ${field('Motif (obligatoire)', '<input class="input" name="details" autocomplete="off" placeholder="ex. solde d’ouverture, erreur de saisie du 12/10…">')}`,
      foot: '<button class="btn primary" data-x="ok">Enregistrer l’ajustement</button>',
      onMount(c) {
        chipGroup(c, 'mode', v => { c.q('#alab').textContent = v === 'target' ? 'Nouveau solde' : v === 'plus' ? 'Montant à ajouter' : 'Montant à retirer'; c.q('#ahelp').classList.toggle('hidden', v !== 'target'); });
        c.q('[data-x=ok]').onclick = e => busy(e.currentTarget, async () => {
          const mode = chipVal(c, 'mode'), v = parseAmt(c.q('[name=amount]').value);
          if (v == null) throw new Error('Indiquez un montant.');
          const p = { account_id: Number(c.q('[name=acc]').value), details: c.q('[name=details]').value };
          if (mode === 'target') p.target = v; else p.amount = mode === 'plus' ? Math.abs(v) : -Math.abs(v);
          const r = await rpc('op_ajustement', p);
          c.close(); toast(r.unchanged ? 'Le solde était déjà à ce montant.' : `Ajustement enregistré. Nouveau solde : ${fmt(r.balance)}`, 'ok'); refreshNow();
        });
      },
    });
  }

  function sheetAccountView(id) {
    const a = acc(id);
    if (!a) return;
    const bal = S.bal.get(a.id) ?? 0;
    const mine = isDir() || a.agence_id === S.me.agence_id;
    sheet({
      title: a.name,
      body: `<div class="hero" style="margin-bottom:14px"><div class="label">${esc(ACCOUNT_KIND[a.kind] || '')}${a.bank ? ' · ' + esc(a.bank) : ''}${(S.agences || []).length > 1 ? ' · ' + esc(a.agence_id ? agName(a.agence_id) : 'commun à toutes les agences') : ''}</div><div class="big num">${fmt(bal)}</div></div>
        ${a.kind === 'gtp' ? `<p class="small muted">Commission : ${esc(commRule(a))}</p>` : ''}
        ${a.note ? `<p class="small muted">${esc(a.note)}</p>` : ''}
        <div class="list">
          <button class="row" data-x="ops"><div class="ico neu">${icon('list')}</div><div class="main"><div class="t">Voir les opérations du mois</div></div>${icon('chev')}</button>
          ${['gtp', 'electronique'].includes(a.kind) && isDir() ? `<button class="row" data-x="appro"><div class="ico rc">${icon('truck')}</div><div class="main"><div class="t">Approvisionner</div></div>${icon('chev')}</button>` : ''}
          ${mine ? `<button class="row" data-x="tr"><div class="ico neu">${icon('swap')}</div><div class="main"><div class="t">Transférer vers un autre compte</div></div>${icon('chev')}</button>
          <button class="row" data-x="adj"><div class="ico neu">${icon('adjust')}</div><div class="main"><div class="t">Ajuster le solde</div></div>${icon('chev')}</button>` : ''}
          ${isDir() ? `<button class="row" data-x="edit"><div class="ico neu">${icon('tag')}</div><div class="main"><div class="t">Modifier le compte${a.kind === 'gtp' ? ' et la commission' : ''}</div></div>${icon('chev')}</button>` : ''}
        </div>
        ${!mine ? `<p class="small muted">Compte commun : seule la direction peut le modifier ou l’ajuster.</p>` : ''}`,
      onMount(c) {
        c.q('[data-x=ops]').onclick = () => { c.close(); Object.assign(S.jf, { period: 'month', account: String(a.id), kind: '', user: '', q: '' }); location.hash = '#/journal'; if (S.view === 'journal') route(); };
        const ap = c.q('[data-x=appro]'); if (ap) ap.onclick = () => { c.close(); sheetAppro(a.id); };
        const tr = c.q('[data-x=tr]'); if (tr) tr.onclick = () => { c.close(); sheetTransfert(a.id); };
        const adj = c.q('[data-x=adj]'); if (adj) adj.onclick = () => { c.close(); sheetAjust(a.id); };
        const ed = c.q('[data-x=edit]'); if (ed) ed.onclick = () => { c.close(); sheetAccountEdit(a); };
      },
    });
  }

  async function vClotures(m) {
    const [pend, hist] = await Promise.all([rpc('list_closings', { status: 'soumise' }), rpc('list_closings', { limit: 40 })]);
    S.pendingClosings = pend.length; renderTabbar();
    m.innerHTML = `
      ${section('À valider', pend.length ? pend.map(c => `
        <div class="card pad" style="margin-bottom:10px">
          <div style="display:flex;justify-content:space-between;gap:8px;margin-bottom:10px"><b>${esc(c.agent)}${(S.agences || []).length > 1 && c.agence ? ` <span class="tag">${esc(c.agence)}</span>` : ''}</b><span class="small muted">${fmtDT(c.created_at)}</span></div>
          <dl class="kv">
            <dt>Attendu</dt><dd class="num">${fmt(c.expected)}</dd>
            <dt>Compté</dt><dd class="num">${fmt(c.counted)}</dd>
            <dt>Écart</dt><dd class="num ${c.gap < 0 ? 'neg' : c.gap > 0 ? 'warn' : 'pos'}">${c.gap === 0 ? 'Caisse juste' : fmtSigned(c.gap)}</dd>
            <dt>Remis (au coffre)</dt><dd class="num">${fmt(c.handed)}</dd>
            <dt>Gardé en caisse</dt><dd class="num">${fmt(c.counted - c.handed)}</dd>
            ${c.note ? `<dt>Note</dt><dd>${esc(c.note)}</dd>` : ''}
          </dl>
          <div class="grid2" style="margin-top:14px"><button class="btn danger" data-rej="${c.id}">Refuser</button><button class="btn primary" data-val="${c.id}">Valider</button></div>
        </div>`).join('') : '<div class="list"><div class="empty">Aucune clôture en attente.</div></div>')}
      ${section('Historique', hist.filter(c => c.status !== 'soumise').length ? `<div class="list">${hist.filter(c => c.status !== 'soumise').map(closingRow).join('')}</div>` : '<div class="list"><div class="empty">Aucune clôture traitée.</div></div>')}`;
    m.querySelectorAll('[data-val]').forEach(b => b.onclick = async () => {
      const c = pend.find(x => x.id === Number(b.dataset.val));
      const ok = await ask('Valider la clôture', `<p>${fmt(c.handed)} passeront de la caisse de ${esc(c.agent)} au coffre${c.agence && (S.agences || []).length > 1 ? ' de ' + esc(c.agence) : ''}.</p>${c.gap ? `<p>L’écart de ${fmtSigned(c.gap)} sera enregistré comme ajustement de sa caisse.</p>` : ''}`, 'Valider');
      if (ok) busy(b, async () => { await rpc('closing_validate', { id: c.id, accept: true }); toast('Clôture validée.', 'ok'); route(); });
    });
    m.querySelectorAll('[data-rej]').forEach(b => b.onclick = async () => {
      const note = await askText('Refuser la clôture', 'Motif (l’agent devra refaire sa clôture)', { okLabel: 'Refuser', danger: true });
      if (note) busy(b, async () => { await rpc('closing_validate', { id: Number(b.dataset.rej), accept: false, note }); toast('Clôture refusée.', 'ok'); route(); });
    });
  }

  function vPlus(m) {
    const item = (href, ic, t, s) => `<a class="row" href="${href}"><div class="ico neu">${icon(ic)}</div><div class="main"><div class="t">${t}</div><div class="s">${s}</div></div>${icon('chev')}</a>`;
    m.innerHTML = `
      ${section('Suivi', `<div class="list">
        ${item('#/clients', 'users', 'Clients', 'Fiches clients, cartes, historique')}
        ${isDir() ? item('#/doublons', 'users', 'Clients en double', 'Repérer et fusionner les fiches d’une même personne') : ''}
        ${item('#/commissions', 'card', 'Commissions à vérifier', 'Pointer les commissions avec le rapport du GTP')}
        ${item('#/rapports', 'chart', 'Rapports', 'Par agence, par produit, par GTP, par agent, commissions par mois')}
        ${item('#/cartes', 'card', 'Cartes clients', 'Rechercher, corriger, ajouter')}
        <button class="row" data-do="export"><div class="ico neu">${icon('download')}</div><div class="main"><div class="t">Exporter en Excel</div><div class="s">Sauvegarde de toutes les opérations d’une période</div></div>${icon('chev')}</button>
      </div>`)}
      ${isDir() ? section('Réglages', `<div class="list">
        ${item('#/agences', 'bank', 'Agences', 'Créer une agence, soldes entre agences')}
        ${item('#/equipe', 'users', 'Équipe', 'Activer les agents, choisir leur agence et leurs droits')}
        ${item('#/comptes', 'bank', 'Comptes et commissions', 'GTP, comptes électroniques, coffre, seuils d’alerte')}
        ${item('#/produits', 'tag', 'Produits de caisse', 'Boutons des opérations de caisse')}
        ${item('#/import', 'upload', 'Importer depuis Excel', 'Cartes clients et soldes de départ')}
        <button class="row" data-do="company"><div class="ico neu">${icon('receipt')}</div><div class="main"><div class="t">Reçu et coordonnées</div><div class="s">Adresse, téléphone, message, format d’impression</div></div>${icon('chev')}</button>
      </div>`) : ''}
      ${section('Mon compte', `<div class="list"><button class="row" data-do="me"><div class="ico rc">${esc(initials(S.me.full_name))}</div><div class="main"><div class="t">${esc(S.me.full_name)}</div><div class="s">${esc(roleLabel(S.me.role))} · mot de passe, déconnexion</div></div>${icon('chev')}</button></div>`)}`;
    return Promise.resolve();
  }

  // ---------------------------------------------------------------- rapports
  async function vRapports(m) {
    const p = S.rp;
    const [from, to] = p.period === 'custom' ? [p.from || todayIso(), p.to || todayIso()] : period(p.period);
    const r = await rpc('report', { from, to });
    const tbl = (head, rows, total) => `<div class="card scrollx"><table class="tbl"><thead><tr>${head.map((h, i) => `<th class="${i ? 'r' : ''}">${h}</th>`).join('')}</tr></thead><tbody>
      ${rows.length ? rows.map(cols => `<tr>${cols.map((c, i) => `<td class="${i ? 'r num' : ''}">${c}</td>`).join('')}</tr>`).join('') : `<tr><td colspan="${head.length}" class="muted">Rien sur cette période.</td></tr>`}
      ${total && rows.length ? `<tr class="total">${total.map((c, i) => `<td class="${i ? 'r num' : ''}">${c}</td>`).join('')}</tr>` : ''}</tbody></table></div>`;
    const s = (l, k) => l.reduce((a, x) => a + Number(x[k] || 0), 0);
    const monthName = ym => { const [y, mo] = ym.split('-'); return new Date(+y, +mo - 1, 1).toLocaleDateString('fr-FR', { month: 'long', year: 'numeric' }); };
    m.innerHTML = `
      <div class="chips">${PERIODS.filter(x => x[0] !== '7d').map(([k, l]) => `<button class="chip ${p.period === k ? 'on' : ''}" data-p="${k}">${l}</button>`).join('')}</div>
      <div class="filters ${p.period === 'custom' ? '' : 'hidden'}">
        <input class="input" type="date" id="rfrom" value="${esc(from)}" aria-label="Du"><input class="input" type="date" id="rto" value="${esc(to)}" aria-label="Au"></div>
      <p class="small muted" style="margin:0 4px 14px">Du ${fmtDate(from)} au ${fmtDate(to)}${r.agence_id ? ' · agence ' + esc(agName(r.agence_id)) : (S.agences || []).length > 1 ? ' · toutes les agences' : ''}</p>
      ${(r.by_agence || []).length > 1 ? section('Par agence', tbl(['Agence', 'Opér.', 'Recharges', 'Rechargé', 'Commission', 'Caisse', 'Dépenses', 'Doit à la principale'],
        r.by_agence.map(x => [esc(x.agence), x.n, x.recharges_n, fmtN(x.recharges_amount), fmtN(x.commission), x.caisse_n, fmtN(x.depenses), x.liaison == null ? '—' : fmtN(x.liaison)]),
        ['Total', s(r.by_agence, 'n'), s(r.by_agence, 'recharges_n'), fmtN(s(r.by_agence, 'recharges_amount')), fmtN(s(r.by_agence, 'commission')), s(r.by_agence, 'caisse_n'), fmtN(s(r.by_agence, 'depenses')), ''])) : ''}
      ${section('Recharges par compte GTP', tbl(['Compte', 'Nb', 'Rechargé', 'Commission', 'Approvisionné'],
        r.by_gtp.map(x => [esc(x.name), x.n, fmtN(x.amount), fmtN(x.commission), fmtN(x.appro)]),
        ['Total', s(r.by_gtp, 'n'), fmtN(s(r.by_gtp, 'amount')), fmtN(s(r.by_gtp, 'commission')), fmtN(s(r.by_gtp, 'appro'))]))}
      ${section('Commissions par mois', tbl(['Mois', 'Compte', 'Nb', 'Rechargé', 'Commission'],
        r.by_month.map(x => [monthName(x.month), esc(x.name), x.n, fmtN(x.amount), fmtN(x.commission) + (Number(x.to_check) ? ` <span class="tag amber">${x.to_check} à vérifier</span>` : '')]),
        ['Total', '', s(r.by_month, 'n'), fmtN(s(r.by_month, 'amount')), fmtN(s(r.by_month, 'commission'))]))}
      ${section('Caisse par produit', tbl(['Produit', 'Nb', 'Montant'],
        r.by_product.map(x => [`${esc(x.name)} <span class="tag ${x.direction === 'entree' ? 'green' : 'red'}">${x.direction === 'entree' ? 'Entrée' : 'Sortie'}</span>`, x.n, fmtN(x.total)])))}
      ${section('Par agent', tbl(['Agent', 'Opér.', 'Recharges', 'Rechargé', 'Commission', 'Caisse', 'Dépenses', 'Annulées'],
        r.by_agent.map(x => [esc(x.agent), x.n, x.recharges_n, fmtN(x.recharges_amount), fmtN(x.commission), x.caisse_n, fmtN(x.depenses), x.annulees])))}
      ${section('Dépenses par catégorie', tbl(['Catégorie', 'Nb', 'Total'], r.depenses.map(x => [esc(x.category), x.n, fmtN(x.total)]), ['Total', s(r.depenses, 'n'), fmtN(s(r.depenses, 'total'))]))}
      ${section('Clôtures validées', tbl(['Agent', 'Nb', 'Écarts', 'Remis'], r.closings.map(x => [esc(x.agent), x.n, `<span class="${x.gap < 0 ? 'neg' : ''}">${fmtN(x.gap)}</span>`, fmtN(x.handed)])))}
      <button class="btn block" id="rexp">${icon('download')} Exporter cette période en Excel</button>`;
    m.querySelectorAll('[data-p]').forEach(b => b.onclick = () => { p.period = b.dataset.p; if (p.period === 'custom') { p.from = from; p.to = to; } route(); });
    const rf = m.querySelector('#rfrom'), rt = m.querySelector('#rto');
    rf.onchange = () => { p.from = rf.value; route(); }; rt.onchange = () => { p.to = rt.value; route(); };
    m.querySelector('#rexp').onclick = e => busy(e.currentTarget, () => exportExcel(from, to));
  }

  // ---------------------------------------------------------------- cartes clients
  // ---------------------------------------------------------------- clients en double
  function lev(a, b) {
    if (a === b) return 0;
    let prev = Array.from({ length: b.length + 1 }, (_, i) => i);
    for (let i = 1; i <= a.length; i++) {
      const cur = [i];
      for (let j = 1; j <= b.length; j++) cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
      prev = cur;
    }
    return prev[b.length];
  }
  // Ressemblance de deux noms, de 0 (rien à voir) à 1 (identiques, quel que soit l'ordre des mots).
  function nameSimilarity(a, b) {
    const n = s => norm(s).replace(/[^A-Z0-9]+/g, ' ').trim().split(' ').filter(Boolean).sort().join(' ');
    const x = n(a), y = n(b);
    if (!x || !y) return 0;
    if (x === y) return 1;
    return 1 - lev(x, y) / Math.max(x.length, y.length);
  }
  const bestKeep = list => list.slice().sort((a, b) => (b.cards.length + Number(b.n_ops)) - (a.cards.length + Number(a.n_ops)) || b.name.length - a.name.length || a.id - b.id)[0];

  async function vDoublons(m) {
    const raw = await rpc('find_duplicates', { limit: 1000 });
    const ignored = new Set(S.settings.dup_ignored || []);
    const seen = new Set(), groups = [];
    for (const g of raw) {
      const sig = g.clients.map(c => c.id).sort((a, b) => a - b).join('-');
      if (seen.has(sig) || ignored.has(sig)) continue;
      seen.add(sig);
      let sim = 1;
      g.clients.forEach(c => g.clients.forEach(d => { if (c.id < d.id) sim = Math.min(sim, nameSimilarity(c.name, d.name)); }));
      // Téléphones différents = peut-être deux personnes qui portent le même nom.
      const conflict = new Set(g.clients.map(c => String(c.phone || '').replace(/\D/g, '')).filter(Boolean)).size > 1;
      groups.push({ sig, key: g.key, clients: g.clients, sim, conflict, evident: sim >= 0.85 && !conflict });
    }
    groups.sort((a, b) => (b.evident - a.evident) || (a.conflict - b.conflict) || (b.sim - a.sim));
    let shown = 40;
    const groupHtml = g => {
      const keep = bestKeep(g.clients);
      const tag = g.evident ? '<span class="tag green">Même personne très probable</span>'
        : g.conflict && g.sim >= 0.85 ? '<span class="tag amber">Même nom, mais téléphones différents : vérifiez</span>'
        : g.sim >= 0.6 ? '<span class="tag amber">Noms proches : vérifiez</span>' : '<span class="tag red">Noms différents : sans doute pas la même personne</span>';
      const preCheck = g.evident || (g.sim >= 0.6 && !g.conflict);
      const why = g.key.startsWith('tel:') ? `même téléphone (${esc(g.clients.find(c => c.phone)?.phone || '')})` : 'même nom, mots dans un autre ordre';
      return `<div class="card pad" style="margin-bottom:10px" data-group="${g.sig}">
        <div style="display:flex;justify-content:space-between;gap:8px;flex-wrap:wrap;margin-bottom:8px">${tag}<span class="small muted">${why}</span></div>
        <div class="dtable scrollx"><table><thead><tr><th>Garder</th><th>Fusionner</th><th>Nom</th><th>Téléphone</th><th>Cartes</th><th class="r">Opér.</th></tr></thead><tbody>
          ${g.clients.map(c => `<tr style="cursor:default">
            <td><input type="radio" name="k-${g.sig}" value="${c.id}" ${c.id === keep.id ? 'checked' : ''} style="width:20px;height:20px;accent-color:var(--brand)"></td>
            <td><input type="checkbox" data-merge="${c.id}" ${c.id !== keep.id && preCheck ? 'checked' : ''} ${c.id === keep.id ? 'disabled' : ''} style="width:20px;height:20px;accent-color:var(--brand)"></td>
            <td style="font-weight:600">${esc(c.name)}<div class="sub">fiche n° ${c.id}</div></td>
            <td class="num">${esc(c.phone || '')}</td>
            <td class="sub">${c.cards.map(k => `<span style="${k.active ? '' : 'text-decoration:line-through'}">${esc(k.bank)} ${esc(k.card_number)}</span>`).join('<br>') || '—'}</td>
            <td class="r num">${c.n_ops}</td></tr>`).join('')}</tbody></table></div>
        <div class="grid2" style="margin-top:10px"><button class="btn small" data-ignore="${g.sig}">Pas la même personne</button><button class="btn small primary" data-domerge="${g.sig}">Fusionner</button></div></div>`;
    };
    const draw = () => {
      const left = groups.filter(g => !g.done);
      m.innerHTML = `
        <div class="alert info"><span>Fiches qui ont le même téléphone, ou le même nom écrit dans un autre ordre. Pour chaque groupe : choisissez la fiche à <b>garder</b>, cochez celles à <b>fusionner</b> dedans. Leurs cartes et leur historique passent sur la fiche gardée ; les autres fiches sont désactivées (rien n’est effacé).</span></div>
        ${left.length ? `<p><b>${left.length}</b> groupe(s) à vérifier, dont <b>${left.filter(g => g.evident).length}</b> où il s’agit très probablement de la même personne (noms presque identiques, téléphones qui concordent).</p>
          ${left.some(g => g.evident) ? `<button class="btn primary block" id="dauto" style="margin-bottom:16px">Fusionner d’un coup les ${left.filter(g => g.evident).length} groupes évidents</button>` : ''}
          ${left.slice(0, shown).map(groupHtml).join('')}
          ${left.length > shown ? `<button class="btn block" id="dmore">Afficher la suite (${left.length - shown} groupes)</button>` : ''}`
          : '<div class="list"><div class="empty">Aucun doublon à vérifier. 👍</div></div>'}`;
      m.querySelectorAll('[data-group]').forEach(el => {
        const sig = el.dataset.group, g = groups.find(x => x.sig === sig);
        el.querySelectorAll('input[type=radio]').forEach(r => r.onchange = () => {
          el.querySelectorAll('[data-merge]').forEach(cb => { const isKeep = cb.dataset.merge === r.value; cb.disabled = isKeep; if (isKeep) cb.checked = false; });
        });
        el.querySelector('[data-domerge]').onclick = e => busy(e.currentTarget, async () => {
          const keepId = Number(el.querySelector('input[type=radio]:checked').value);
          const ids = [...el.querySelectorAll('[data-merge]:checked')].map(cb => Number(cb.dataset.merge)).filter(id => id !== keepId);
          if (!ids.length) throw new Error('Cochez au moins une fiche à fusionner.');
          await rpc('merge_clients', { keep_id: keepId, merge_ids: ids });
          toast(`${ids.length + 1} fiches regroupées.`, 'ok');
          route(); // recharge la liste : une fiche peut figurer dans plusieurs groupes
        });
        el.querySelector('[data-ignore]').onclick = e => busy(e.currentTarget, async () => {
          const list = [...(S.settings.dup_ignored || []), sig];
          await rpc('save_settings', { key: 'dup_ignored', value: list });
          S.settings.dup_ignored = list; g.done = true; draw();
        });
      });
      const more = m.querySelector('#dmore'); if (more) more.onclick = () => { shown += 40; draw(); };
      const auto = m.querySelector('#dauto');
      if (auto) auto.onclick = async () => {
        const todo = groups.filter(g => !g.done && g.evident);
        if (!await ask('Fusionner les groupes évidents ?', `<p>${todo.length} groupes « même personne très probable » (noms presque identiques, téléphones qui concordent) vont être regroupés. Dans chaque groupe, la fiche gardée est celle qui a le plus de cartes et d’opérations.</p><p>Les autres groupes ne sont pas touchés : vous les vérifierez un par un.</p>`, 'Fusionner')) return;
        busy(auto, async () => {
          let n = 0;
          const moved = new Map();                       // fiche fusionnée -> fiche gardée
          const resolve = id => { while (moved.has(id)) id = moved.get(id); return id; };
          for (const g of todo) {
            const keepId = resolve(bestKeep(g.clients).id);
            const ids = [...new Set(g.clients.map(c => resolve(c.id)))].filter(id => id !== keepId);
            g.done = true;
            if (!ids.length) continue;
            await rpc('merge_clients', { keep_id: keepId, merge_ids: ids });
            ids.forEach(id => moved.set(id, keepId));
            n++;
            auto.textContent = `Fusion en cours… ${n} / ${todo.length}`;
          }
          toast(`${n} groupes fusionnés.`, 'ok');
          route();
        });
      };
    };
    draw();
  }

  async function vCartes(m) {
    const banks = [...new Set(S.accounts.filter(a => a.kind === 'gtp').map(a => a.bank).filter(Boolean))];
    m.innerHTML = `
      <div class="chips">${[['', 'Toutes'], ...banks.map(b => [b, b])].map(([k, l]) => `<button class="chip ${S.cf.bank === k ? 'on' : ''}" data-b="${k}">${l}</button>`).join('')}</div>
      <div class="search">${icon('search')}<input class="input" id="cq" placeholder="N° de carte, nom, téléphone" value="${esc(S.cf.q)}" autocomplete="off"></div>
      <div id="cres"><div class="spinner"></div></div>
      <button class="btn block" data-do="newcard" style="margin-top:12px">${icon('plus')} Nouvelle carte client</button>`;
    const load = async () => {
      const res = await rpc('search_cards', { q: S.cf.q, bank: S.cf.bank || null, limit: 100 });
      const box = m.querySelector('#cres'); if (!box) return;
      res.forEach(c => S.cards.set(c.id, c));
      box.innerHTML = res.length ? `<div class="list">${res.map(c => `<button class="row" data-cardedit="${c.id}">
          <div class="ico rc">${esc(c.bank.slice(0, 3))}</div>
          <div class="main"><div class="t">${esc(c.holder_name)}${c.active ? '' : ' <span class="tag">inactive</span>'}</div>
            <div class="s">${esc(c.card_number)} · ${c.n_recharges} recharge(s)${c.last_recharge ? ' · dernière ' + ago(c.last_recharge.at) : ''}</div></div>${icon('chev')}</button>`).join('')}</div>
          ${res.length >= 100 ? '<p class="small muted">100 premiers résultats : précisez la recherche.</p>' : ''}`
        : '<div class="list"><div class="empty">Aucune carte trouvée.</div></div>';
    };
    m.querySelectorAll('[data-b]').forEach(b => b.onclick = () => { S.cf.bank = b.dataset.b; m.querySelectorAll('[data-b]').forEach(x => x.classList.toggle('on', x === b)); load(); });
    const q = m.querySelector('#cq');
    q.addEventListener('input', debounce(() => { S.cf.q = q.value.trim(); load(); }, 250));
    await load();
  }

  // ---------------------------------------------------------------- réglages
  async function vEquipe(m) {
    await loadBoot();
    m.innerHTML = `
      <div class="alert info"><span>Pour ajouter quelqu’un : sur son téléphone ou son ordinateur, il ouvre l’application et touche « Créer un compte ». Il apparaît ensuite ici : choisissez son agence, son rôle et ses droits, puis activez-le.</span></div>
      ${(S.agences || []).map(g => {
        const list = S.team.filter(t => (t.agence_id || mainAgence().id) === g.id);
        return list.length ? section(esc(g.name), `<div class="list">${list.map(t => `<button class="row" data-prof="${t.id}">
          <div class="ico ${t.role !== 'agent' ? 'rc' : 'neu'}">${esc(initials(t.full_name))}</div>
          <div class="main"><div class="t">${esc(t.full_name)} ${t.active ? '' : '<span class="tag amber">À activer</span>'}</div>
            <div class="s">${esc(roleLabel(t.role))} · identifiant « ${esc(t.login || '')} »${t.role === 'agent' ? ` · ${[t.can_caisse && 'caisse', t.can_recharge && 'recharges'].filter(Boolean).join(' + ') || 'aucun droit'}` : ''}</div></div>${icon('chev')}</button>`).join('')}</div>`) : '';
      }).join('')}`;
    m.querySelectorAll('[data-prof]').forEach(b => b.onclick = () => sheetProfile(S.team.find(t => t.id === b.dataset.prof)));
  }

  function sheetProfile(t) {
    sheet({
      title: t.full_name || 'Utilisateur',
      body: `${field('Nom affiché', `<input class="input" name="full_name" value="${esc(t.full_name)}">`)}
        ${field('Agence', `<select class="input" name="agence">${options((S.agences || []).map(g => [g.id, g.name]), t.agence_id || mainAgence().id)}</select>`)}
        ${field('Rôle', `<select class="input" name="role">${options([['agent', 'Agent'], ['gerant', 'Gérant(e) d’agence (gère son agence)'], ['gerante', 'Direction (voit toutes les agences, règle tout)']], t.role)}</select>`)}
        <label class="check"><input type="checkbox" name="active" ${t.active ? 'checked' : ''}> Compte actif (peut se connecter)</label>
        <label class="check"><input type="checkbox" name="can_caisse" ${t.can_caisse ? 'checked' : ''}> Droit : opérations de caisse</label>
        <label class="check"><input type="checkbox" name="can_recharge" ${t.can_recharge ? 'checked' : ''}> Droit : recharges de cartes</label>
        <p class="small muted">Identifiant de connexion : « ${esc(t.login || '')} ». Désactiver un compte bloque la connexion sans effacer son historique.</p>`,
      foot: '<button class="btn primary" data-x="ok">Enregistrer</button>',
      onMount(c) {
        c.q('[data-x=ok]').onclick = e => busy(e.currentTarget, async () => {
          await rpc('save_profile', { id: t.id, full_name: c.q('[name=full_name]').value, role: c.q('[name=role]').value, agence_id: Number(c.q('[name=agence]').value),
            active: c.q('[name=active]').checked, can_caisse: c.q('[name=can_caisse]').checked, can_recharge: c.q('[name=can_recharge]').checked });
          c.close(); toast('Enregistré.', 'ok'); await loadBoot(); route();
        });
      },
    });
  }

  async function vComptes(m) {
    await loadBoot();
    const groups = [['gtp', 'Comptes GTP'], ['electronique', 'Comptes électroniques'], ['banque', 'Banque'], ['coffre', 'Coffre'], ['caisse', 'Caisses des agents']];
    m.innerHTML = groups.map(([k, l]) => {
      const list = S.accounts.filter(a => a.kind === k);
      if (!list.length) return '';
      return section(l, `<div class="list">${list.map(a => `<button class="row" data-accedit="${a.id}">
        <div class="ico ${k === 'gtp' ? 'rc' : 'neu'}">${k === 'gtp' ? esc((a.bank || '?').slice(0, 3)) : icon(k === 'caisse' ? 'cash' : 'bank')}</div>
        <div class="main"><div class="t">${esc(a.name)} ${a.active ? '' : '<span class="tag">désactivé</span>'} ${(S.agences || []).length > 1 ? `<span class="tag ${a.agence_id ? '' : 'brand'}">${esc(a.agence_id ? agName(a.agence_id) : 'Commun')}</span>` : ''}</div>
          <div class="s">${k === 'gtp' ? esc(commRule(a)) : esc(ACCOUNT_KIND[k])}${Number(a.low_alert) ? ' · alerte sous ' + fmt(a.low_alert) : ''}${a.note ? ' · ' + esc(a.note) : ''}</div></div>${icon('chev')}</button>`).join('')}</div>`);
    }).join('') + `<button class="btn block" id="newacc">${icon('plus')} Nouveau compte</button>`;
    m.querySelectorAll('[data-accedit]').forEach(b => b.onclick = () => sheetAccountEdit(acc(b.dataset.accedit)));
    m.querySelector('#newacc').onclick = () => sheetAccountEdit(null);
  }

  function sheetAccountEdit(a) {
    const x = a || { kind: 'gtp', bank: '', comm_rate: 0, comm_divisor: 1.1, comm_min: 500, comm_rounding: 'round', low_alert: 0, active: true, sort: 100 };
    sheet({
      title: a ? 'Modifier le compte' : 'Nouveau compte',
      body: `${field('Nom', `<input class="input" name="name" value="${esc(x.name || '')}">`)}
        ${field('Type', `<select class="input" name="kind" ${a && a.kind === 'caisse' ? 'disabled' : ''}>${options(Object.entries(ACCOUNT_KIND).filter(([k]) => k !== 'caisse' || x.kind === 'caisse'), x.kind)}</select>`)}
        ${field('Agence', `<select class="input" name="agence" ${a && a.kind === 'caisse' ? 'disabled' : ''}>${options([['', 'Commun à toutes les agences'], ...(S.agences || []).map(g => [g.id, g.name])], x.agence_id || '')}</select>`,
          'Commun : utilisable par toutes les agences (GTP, Wave, Orange Money, RIA, Western Union…). Sinon : propre à une agence (son Mixx, son Flooz, son coffre).')}
        <div id="gtpf">
          ${field('Banque', `<input class="input" name="bank" value="${esc(x.bank || '')}" placeholder="GTCO, UBA, ECOBANK, BNIFIC" style="text-transform:uppercase">`, 'Les cartes sont rattachées à la banque.')}
          <div class="grid2">${field('Commission (%)', `<input class="input" name="comm_rate" inputmode="decimal" value="${esc(String(x.comm_rate).replace('.', ','))}">`)}
            ${field('Divisé par', `<input class="input" name="comm_divisor" inputmode="decimal" value="${esc(String(x.comm_divisor).replace('.', ','))}">`)}</div>
          <div class="grid2">${field('Minimum (F)', `<input class="input money" name="comm_min" inputmode="numeric" value="${fmtN(x.comm_min)}">`)}
            ${field('Arrondi', `<select class="input" name="comm_rounding">${options([['round', 'Au plus proche'], ['ceil', 'Au-dessus'], ['floor', 'En dessous']], x.comm_rounding)}</select>`)}</div>
          <div class="calc small" id="cex"></div>
        </div>
        ${field('Alerte si le solde passe sous (F)', `<input class="input money" name="low_alert" inputmode="numeric" value="${Number(x.low_alert) ? fmtN(x.low_alert) : ''}" placeholder="0 = pas d’alerte">`)}
        <label class="check"><input type="checkbox" name="allow_negative" ${x.allow_negative ? 'checked' : ''}> Autoriser un solde négatif</label>
        <label class="check"><input type="checkbox" name="active" ${x.active ? 'checked' : ''}> Compte actif</label>
        ${field('Ordre d’affichage', `<input class="input" name="sort" inputmode="numeric" value="${esc(x.sort)}">`)}
        ${field('Note', `<input class="input" name="note" value="${esc(x.note || '')}">`)}`,
      foot: '<button class="btn primary" data-x="ok">Enregistrer</button>',
      onMount(c) {
        const dec = n => Number(String(c.q(`[name=${n}]`).value).replace(',', '.')) || 0;
        const read = () => ({ id: a && a.id, name: c.q('[name=name]').value, kind: c.q('[name=kind]').value, bank: c.q('[name=bank]').value,
          comm_rate: dec('comm_rate'), comm_divisor: dec('comm_divisor') || 1, comm_min: parseAmt(c.q('[name=comm_min]').value) || 0,
          comm_rounding: c.q('[name=comm_rounding]').value, low_alert: parseAmt(c.q('[name=low_alert]').value) || 0,
          allow_negative: c.q('[name=allow_negative]').checked, active: c.q('[name=active]').checked,
          sort: Number(c.q('[name=sort]').value) || 100, note: c.q('[name=note]').value,
          agence_id: c.q('[name=agence]').value ? Number(c.q('[name=agence]').value) : null });
        const upd = () => {
          const v = read();
          c.q('#gtpf').classList.toggle('hidden', v.kind !== 'gtp');
          c.q('#cex').innerHTML = [10000, 68000, 100000, 500000].map(n => `<div class="line"><span>${fmt(n)}</span><span class="num">commission ${fmt(commission(v, n))}</span></div>`).join('');
        };
        c.qa('input,select').forEach(i => i.addEventListener('input', upd));
        upd();
        c.q('[data-x=ok]').onclick = e => busy(e.currentTarget, async () => {
          await rpc('save_account', read()); c.close(); toast('Compte enregistré.', 'ok'); await loadBoot(); route();
        });
      },
    });
  }

  // ---------------------------------------------------------------- agences (direction)
  async function vAgences(m) {
    await loadBoot();
    const d = await rpc('dashboard', { agence_id: null });
    const lia = new Map((d.liaisons || []).map(l => [l.agence_id, l]));
    m.innerHTML = `
      <div class="alert info"><span>L’agence principale dirige. Les autres agences ont leur coffre, leurs caisses et leurs comptes propres (Mixx, Flooz…) ; les comptes communs (GTP, Wave, Orange Money, RIA, Western Union) appartiennent à l’agence principale. Chaque agence a un solde avec la principale, mis à jour automatiquement.</span></div>
      <div class="list">${(S.agences || []).map(g => {
        const n = S.team.filter(t => (t.agence_id || mainAgence().id) === g.id && t.active).length;
        const l = lia.get(g.id), x = l ? liaisonText(l) : null;
        return `<button class="row" data-agedit="${g.id}">
          <div class="ico ${g.is_main ? 'rc' : 'neu'}">${esc(initials(g.name))}</div>
          <div class="main"><div class="t">${esc(g.name)} ${g.is_main ? '<span class="tag brand">Principale</span>' : ''} ${g.active ? '' : '<span class="tag">fermée</span>'}</div>
            <div class="s">${n} personne(s) active(s)${x ? ' · ' + esc(x.t) : ''}</div></div>${icon('chev')}</button>`;
      }).join('')}</div>
      <button class="btn block" id="newag" style="margin-top:12px">${icon('plus')} Nouvelle agence</button>`;
    m.querySelectorAll('[data-agedit]').forEach(b => b.onclick = () => sheetAgence((S.agences || []).find(g => g.id === Number(b.dataset.agedit))));
    m.querySelector('#newag').onclick = () => sheetAgence(null);
  }

  function sheetAgence(g) {
    sheet({
      title: g ? g.name : 'Nouvelle agence',
      body: `${field('Nom de l’agence', `<input class="input autofocus" name="name" value="${esc(g ? g.name : '')}" placeholder="ex. Noépé">`)}
        ${g && !g.is_main ? `<label class="check"><input type="checkbox" name="active" ${g.active ? 'checked' : ''}> Agence ouverte</label>` : ''}
        ${g ? '' : '<p class="small muted">Un coffre est créé automatiquement pour la nouvelle agence. Ensuite : créez ses comptes propres (Mixx, Flooz…) dans « Comptes », et rattachez son gérant et ses agents dans « Équipe ».</p>'}
        ${g && !g.is_main ? '<button class="btn block" data-x="lia" type="button" style="margin-top:8px">Voir le solde avec l’agence principale</button>' : ''}`,
      foot: '<button class="btn primary" data-x="ok">Enregistrer</button>',
      onMount(c) {
        const lb = c.q('[data-x=lia]'); if (lb) lb.onclick = () => { c.close(); sheetLiaison(g.id); };
        c.q('[data-x=ok]').onclick = e => busy(e.currentTarget, async () => {
          const p = { id: g ? g.id : null, name: c.q('[name=name]').value };
          if (c.q('[name=active]')) p.active = c.q('[name=active]').checked;
          await rpc('save_agence', p);
          c.close(); toast(g ? 'Agence enregistrée.' : 'Agence créée.', 'ok'); await loadBoot(); renderShell(); route();
        });
      },
    });
  }

  async function vProduits(m) {
    await loadBoot();
    m.innerHTML = `<div class="list">${S.products.map(p => `<button class="row" data-prodedit="${p.id}">
        <div class="ico ${p.direction === 'entree' ? 'in' : 'out'}">${icon(p.direction === 'entree' ? 'in' : 'out')}</div>
        <div class="main"><div class="t">${esc(p.name)} ${p.active ? '' : '<span class="tag">masqué</span>'}</div>
          <div class="s">${p.direction === 'entree' ? 'Entrée' : 'Sortie'}${p.counter_account_id ? ' · compte lié : ' + esc((acc(p.counter_account_id) || {}).name || '?') : ''}${p.category ? ' · ' + esc(p.category) : ''}</div></div>${icon('chev')}</button>`).join('')}</div>
      <button class="btn block" id="newp" style="margin-top:12px">${icon('plus')} Nouveau produit</button>`;
    m.querySelectorAll('[data-prodedit]').forEach(b => b.onclick = () => sheetProduct(S.products.find(p => p.id === Number(b.dataset.prodedit))));
    m.querySelector('#newp').onclick = () => sheetProduct(null);
  }

  function sheetProduct(p) {
    const x = p || { direction: 'entree', active: true, sort: 100, category: '' };
    const elecs = S.accounts.filter(a => a.kind === 'electronique' || a.kind === 'banque');
    sheet({
      title: p ? 'Modifier le produit' : 'Nouveau produit',
      body: `${field('Nom du bouton', `<input class="input" name="name" value="${esc(x.name || '')}">`)}
        ${field('Sens pour la caisse (espèces)', `<select class="input" name="direction">${options([['entree', 'Entrée : le client donne de l’argent'], ['sortie', 'Sortie : on remet de l’argent au client']], x.direction)}</select>`)}
        ${field('Compte électronique lié', `<select class="input" name="counter">${options(elecs.map(a => [a.id, a.name]), x.counter_account_id, '— Aucun —')}</select>`, 'Ex. Dépôt Mixx : les espèces entrent, le solde Mixx diminue. Retrait : l’inverse.')}
        ${field('Catégorie', `<input class="input" name="category" value="${esc(x.category || '')}" placeholder="Transfert, Crédit, Services…">`)}
        ${field('Ordre d’affichage', `<input class="input" name="sort" inputmode="numeric" value="${esc(x.sort)}">`)}
        <label class="check"><input type="checkbox" name="active" ${x.active ? 'checked' : ''}> Afficher ce bouton aux agents</label>`,
      foot: '<button class="btn primary" data-x="ok">Enregistrer</button>',
      onMount(c) {
        c.q('[data-x=ok]').onclick = e => busy(e.currentTarget, async () => {
          await rpc('save_product', { id: p && p.id, name: c.q('[name=name]').value, direction: c.q('[name=direction]').value,
            counter_account_id: c.q('[name=counter]').value || null, category: c.q('[name=category]').value,
            sort: Number(c.q('[name=sort]').value) || 100, active: c.q('[name=active]').checked });
          c.close(); toast('Produit enregistré.', 'ok'); await loadBoot(); route();
        });
      },
    });
  }

  function sheetMe() {
    const demo = S.api.mode === 'demo';
    sheet({
      title: 'Mon compte',
      body: `<dl class="kv"><dt>Nom</dt><dd>${esc(S.me.full_name)}</dd><dt>Rôle</dt><dd>${esc(roleLabel(S.me.role))}</dd>${S.me.agence ? `<dt>Agence</dt><dd>${esc(S.me.agence)}</dd>` : ''}<dt>Identifiant</dt><dd>${esc(S.me.login || '')}</dd>
          ${!isMgr() ? `<dt>Droits</dt><dd>${[S.me.can_caisse && 'caisse', S.me.can_recharge && 'recharges'].filter(Boolean).join(', ') || 'aucun'}</dd>` : ''}</dl>
        ${!demo ? `<div style="margin-top:18px">${field('Nouveau mot de passe', '<input class="input" type="password" name="pwd" minlength="6" autocomplete="new-password">')}
          <button class="btn block" data-x="pwd">Changer mon mot de passe</button></div>` : ''}
        ${demo ? '<button class="btn block" data-x="reset" style="margin-top:18px">Réinitialiser la démo</button>' : ''}`,
      foot: `<button class="btn danger" data-x="out">${demo ? 'Changer d’utilisateur' : 'Se déconnecter'}</button>`,
      onMount(c) {
        c.q('[data-x=out]').onclick = async () => { await S.api.signOut(); location.hash = ''; location.reload(); };
        const p = c.q('[data-x=pwd]');
        if (p) p.onclick = () => busy(p, async () => {
          const v = c.q('[name=pwd]').value;
          if (v.length < 6) throw new Error('Au moins 6 caractères.');
          await S.api.changePassword(v); c.close(); toast('Mot de passe changé.', 'ok');
        });
        const r = c.q('[data-x=reset]');
        if (r) r.onclick = async () => { if (await ask('Réinitialiser la démo ?', 'Toutes les données d’essai seront remplacées par les données de départ.', 'Réinitialiser', true)) { await S.api.reset(); location.hash = ''; location.reload(); } };
      },
    });
  }

  // ---------------------------------------------------------------- Excel : chargement, export
  function loadXLSX() {
    if (window.XLSX) return Promise.resolve();
    return new Promise((res, rej) => {
      const s = document.createElement('script');
      s.src = XLSX_URL; s.onload = res; s.onerror = () => rej(new Error('Impossible de charger le module Excel (connexion Internet ?).'));
      document.head.append(s);
    });
  }

  function sheetExport() {
    const [f, t] = period('month');
    sheet({
      title: 'Exporter en Excel',
      body: `<p class="small muted" style="margin-top:0">Le fichier contient toutes les opérations (y compris annulées), les soldes actuels et les rapports de la période. Faites-le régulièrement : c’est votre sauvegarde.</p>
        <div class="grid2">${field('Du', `<input class="input" type="date" name="from" value="${f}">`)}${field('Au', `<input class="input" type="date" name="to" value="${t}">`)}</div>`,
      foot: `<button class="btn primary" data-x="ok">${icon('download')} Télécharger</button>`,
      onMount(c) { c.q('[data-x=ok]').onclick = e => busy(e.currentTarget, async () => { await exportExcel(c.q('[name=from]').value, c.q('[name=to]').value); c.close(); }); },
    });
  }

  async function exportExcel(from, to, filters = {}) {
    await loadXLSX();
    const XL = window.XLSX;
    const [ops, rep, dash] = await Promise.all([rpc('list_operations', Object.assign({ from, to, limit: 20000 }, filters)), rpc('report', { from, to }), rpc('dashboard')]);
    const wb = XL.utils.book_new();
    const add = (name, rows, widths) => {
      const ws = XL.utils.json_to_sheet(rows.length ? rows : [{ '': 'Aucune donnée' }]);
      if (widths) ws['!cols'] = widths.map(w => ({ wch: w }));
      XL.utils.book_append_sheet(wb, ws, name);
    };
    const opRows = ops.slice().reverse().map(o => ({
      'N°': o.id, Date: fmtDate(o.op_day), Heure: fmtTime(o.created_at), Type: KIND[o.kind] || o.kind, Produit: o.product || '',
      'N° carte': o.card_number || '', Client: o.holder_name || '', Compte: o.account || '', 'Compte 2': o.account2 || '',
      Sens: o.kind === 'caisse' ? (o.direction === 'entree' ? 'Entrée' : 'Sortie') : '',
      Montant: Number(o.amount), Commission: Number(o.commission) || 0, 'Reçu': o.received == null ? '' : Number(o.received),
      Rendu: o.change_given == null ? '' : Number(o.change_given), 'Détails': o.details || '', 'Catégorie': o.category || '',
      Agent: o.agent || '', Statut: o.status === 'ok' ? 'OK' : 'ANNULÉE', 'Motif annulation': o.cancel_reason || '', Signalement: o.flags || '' }));
    add('Opérations', opRows, [7, 11, 7, 16, 22, 11, 28, 18, 18, 8, 12, 11, 10, 10, 28, 14, 16, 9, 24, 22]);
    add('Recharges', opRows.filter(r => r.Type === 'Recharge').map(r => ({ Date: r.Date, Heure: r.Heure, 'N° carte': r['N° carte'], Client: r.Client, 'Compte GTP': r.Compte,
      'Montant TTC': r.Montant, Commission: r.Commission, 'Moyen de paiement': /^Caisse/.test(r['Compte 2']) ? 'ESPÈCES' : r['Compte 2'], Agent: r.Agent, Statut: r.Statut })),
      [11, 7, 11, 30, 18, 12, 11, 18, 16, 9]);
    add('Caisse', opRows.filter(r => ['Caisse', 'Dépense'].includes(r.Type)).map(r => ({ Date: r.Date, Heure: r.Heure, Produit: r.Type === 'Dépense' ? 'DÉPENSE' : r.Produit,
      'Détails': r['Détails'], Entrée: r.Sens === 'Entrée' ? r.Montant : '', Sortie: r.Sens === 'Sortie' || r.Type === 'Dépense' ? r.Montant : '', Compte: r.Compte, Agent: r.Agent, Statut: r.Statut })),
      [11, 7, 24, 28, 11, 11, 18, 16, 9]);
    add('Soldes actuels', dash.balances.map(b => ({ Compte: b.name, Type: ACCOUNT_KIND[b.kind] || b.kind, Banque: b.bank || '', Solde: Number(b.balance) })), [24, 20, 10, 14]);
    add('Commissions par mois', rep.by_month.map(x => ({ Mois: x.month, Compte: x.name, 'Nb recharges': Number(x.n), 'Montant rechargé': Number(x.amount), Commission: Number(x.commission) })), [10, 20, 13, 16, 12]);
    add('Par GTP', rep.by_gtp.map(x => ({ Compte: x.name, Banque: x.bank, 'Nb recharges': Number(x.n), Rechargé: Number(x.amount), Commission: Number(x.commission), Approvisionné: Number(x.appro) })), [20, 10, 13, 14, 12, 14]);
    add('Par produit', rep.by_product.map(x => ({ Produit: x.name, Sens: x.direction === 'entree' ? 'Entrée' : 'Sortie', Nombre: Number(x.n), Montant: Number(x.total) })), [26, 8, 9, 14]);
    add('Par agent', rep.by_agent.map(x => ({ Agent: x.agent, 'Opérations': Number(x.n), Recharges: Number(x.recharges_n), 'Rechargé': Number(x.recharges_amount), Commission: Number(x.commission),
      'Opérations caisse': Number(x.caisse_n), 'Dépenses': Number(x.depenses), 'Annulées': Number(x.annulees) })), [18, 11, 10, 13, 12, 16, 11, 9]);
    XL.writeFile(wb, `AYAEL_export_${from}_au_${to}.xlsx`);
    toast('Fichier Excel téléchargé.', 'ok');
  }

  // ---------------------------------------------------------------- import Excel
  const MONTHS = [['JANV', 1], ['JAN', 1], ['FEVR', 2], ['FERV', 2], ['FEV', 2], ['MARS', 3], ['AVR', 4], ['MAI', 5], ['JUIN', 6], ['JUIL', 7], ['AOUT', 8], ['SEPT', 9], ['OCT', 10], ['NOV', 11], ['DEC', 12]];
  const norm = s => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase();
  function sheetMonth(name) {
    const n = norm(name), y = (n.match(/(20\d\d)/) || [])[1];
    if (!y) return null;
    const m = MONTHS.find(([k]) => n.includes(k));
    return m ? Number(y) * 100 + m[1] : null;
  }
  const CARD_RE = /^\s*(?:RECHARGEMENT\s+[A-Z]+\s+)?(\d{6,9})\s*(?:\(\s*(\d{3,5})\s*\))?\s*-?\s*([A-Za-zÀ-ÿ'’][A-Za-zÀ-ÿ'’ .\-]{2,}?)\s*$/;
  function guessBank(sheetName, fileBank) {
    const n = norm(sheetName);
    if (n.includes('ECOBANK') || n.includes('CASHXPRESS') || n.includes('CASH XPRESS')) return 'ECOBANK';
    if (n.includes('GTCO') || n.includes('GTBANK') || n.includes('AYAEL CI')) return 'GTCO';
    if (n.includes('UBA BURKINA') || n.includes('UBA BF')) return 'UBA BF';
    if (n.includes('UBA') || n.includes('AFRICSIKA')) return 'UBA';
    if (n.includes('BNIFIC')) return 'BNIFIC';
    if (n.includes('BSIC')) return 'BSIC';
    if (/\bBNI\b/.test(n)) return 'BNI';
    return fileBank;
  }

  // Fiche clients en colonnes (N° Identifiant, Noms & prénoms, Téléphone, Statut, Date d'expiration…)
  function parseClientSheet(rows, bank) {
    const h = rows.slice(0, 3).findIndex(r => (r || []).some(v => /IDENTIFIANT/.test(norm(v))) && (r || []).some(v => /^NOMS?\b/.test(norm(v).trim())));
    if (h < 0) return null;
    const head = rows[h].map(v => norm(v).trim());
    const col = re => head.findIndex(v => re.test(v));
    const cId = col(/IDENTIFIANT/), cName = col(/^NOMS?\b/), cPhone = col(/TELEPHONE/), cStatus = col(/STATUT/), cBuy = col(/ACHAT/);
    const cExp = head.map((v, i) => /EXPIRATION/.test(v) ? i : -1).filter(i => i >= 0);
    const today = todayIso();
    const toDate = v => {
      if (typeof v === 'number') return excelDate(v);
      const s = String(v || '').trim();
      let m = s.match(/^(\d{1,2})\s*\/\s*(\d{2})$/);          // 11/26 = fin novembre 2026
      if (m) return isoDay(new Date(2000 + Number(m[2]), Number(m[1]), 0));
      return excelDate(s);
    };
    const out = [];
    for (let i = h + 1; i < rows.length; i++) {
      const r = rows[i] || [];
      const rawId = r[cId], rawName = r[cName];
      if (rawId == null || rawName == null) continue;
      let num = null, last4 = null;
      if (typeof rawId === 'number') num = String(Math.round(rawId));
      else { const m = String(rawId).match(/(\d{5,12})\s*(?:\(\s*(\d{3,5})\s*\))?/); if (m) { num = m[1]; last4 = m[2] || null; } }
      const name = String(rawName).replace(/\s+/g, ' ').trim().toUpperCase();
      if (!num || name.length < 2 || /IDENTIFIANT|NOMS/.test(norm(name))) continue;
      let phone = null;
      if (cPhone >= 0 && r[cPhone] != null) {
        let d = typeof r[cPhone] === 'number' ? String(Math.round(r[cPhone])) : String(r[cPhone]).replace(/\D/g, '');
        if (d.length === 11 && d.startsWith('228')) d = d.slice(3);
        phone = d.length === 8 ? d.replace(/(\d{2})(?=\d)/g, '$1 ') : (d || null);
      }
      const rawStatus = cStatus >= 0 ? r[cStatus] : null;
      const status = typeof rawStatus === 'string' ? norm(rawStatus).trim() : '';
      let exp = typeof rawStatus === 'number' ? toDate(rawStatus) : null;   // certaines lignes ont la date d'expiration dans « Statut »
      cExp.forEach(c => { if (r[c] != null && String(r[c]).trim() !== '') { const d = toDate(r[c]); if (d) exp = d; } });
      const buy = cBuy >= 0 && r[cBuy] != null ? toDate(r[cBuy]) : null;
      const dead = /EXPIR|REMPLACEE|BLOQU|PERDU|RESILI|ANNUL|DESACTIV/.test(status) || (exp && exp < today);
      const note = [buy ? 'Achat ' + fmtDate(buy) : '', exp ? 'Expire ' + fmtDate(exp) : '', last4 ? '4 derniers chiffres : ' + last4 : '',
        status ? 'Statut : ' + status.toLowerCase() : ''].filter(Boolean).join(' · ');
      out.push({ card_number: num, bank, holder_name: name, phone, active: !dead, note });
    }
    return out;
  }
  const excelDate = v => {
    if (typeof v === 'number' && v > 30000 && v < 80000) { const d = new Date(Math.round((v - 25569) * 86400000)); return d.toISOString().slice(0, 10); }
    if (typeof v === 'string') { const m = v.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/); if (m) return `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`; }
    return null;
  };

  function analyseWorkbook(wb, fileName, fileBank) {
    const XL = window.XLSX;
    const cards = new Map();
    const soldes = [];
    wb.SheetNames.forEach(sn => {
      const rows = XL.utils.sheet_to_json(wb.Sheets[sn], { header: 1, raw: true, defval: null, blankrows: false });
      const bank = guessBank(sn, fileBank);
      // Fiche clients en colonnes : reprise ligne par ligne (la dernière ligne d'une même carte l'emporte).
      const fiche = parseClientSheet(rows, bank);
      if (fiche) { fiche.forEach(c => cards.set(c.bank + '|' + c.card_number, Object.assign(c, { names: { [c.holder_name]: 1 } }))); return; }
      // Sinon (fiches de rechargement) : toute cellule texte « numéro + nom »
      rows.forEach(r => r.forEach(v => {
        if (typeof v !== 'string') return;
        const m = v.match(CARD_RE);
        if (!m) return;
        const key = bank + '|' + m[1];
        const name = m[3].replace(/\s+/g, ' ').trim().toUpperCase();
        if (/^(REPORT|SOLDE|VERSEMENT|APPRO)/.test(name)) return;
        const c = cards.get(key) || { card_number: m[1], bank, names: {}, note: m[2] ? `4 derniers chiffres : ${m[2]}` : '' };
        c.names[name] = (c.names[name] || 0) + 1;
        if (m[2] && !c.note) c.note = `4 derniers chiffres : ${m[2]}`;
        cards.set(key, c);
      }));
      // Solde final de la feuille
      let h = -1, dc = -1, sc = -1;
      for (let i = 0; i < Math.min(rows.length, 6) && h < 0; i++) {
        const r = rows[i] || [];
        const di = r.findIndex(v => /^DESIGNATION/.test(norm(v).trim()));
        let si = r.findIndex(v => /SOLDE GTP/.test(norm(v)));
        if (si < 0) si = r.findIndex(v => /^SOLDES?$/.test(norm(v).trim()));
        if (di >= 0 && si >= 0) { h = i; dc = di; sc = si; }
      }
      if (h < 0) return;
      let last = null, lastDate = null;
      for (let i = h + 1; i < rows.length; i++) {
        const r = rows[i] || [];
        const d = excelDate(r[0]); if (d) lastDate = d;
        if (r[dc] != null && String(r[dc]).trim() !== '' && typeof r[sc] === 'number') last = { solde: Math.round(r[sc]), label: String(r[dc]).trim(), date: lastDate };
      }
      if (last) soldes.push({ file: fileName, sheet: sn, month: sheetMonth(sn), bank, ...last });
    });
    const cardList = [...cards.values()].map(c => ({ card_number: c.card_number, bank: c.bank, note: c.note,
      phone: c.phone || null, active: c.active === undefined ? null : c.active,
      holder_name: Object.entries(c.names).sort((a, b) => b[1] - a[1])[0][0] }));
    // Feuilles mensuelles : on garde la plus récente. Autres feuilles : toutes.
    const monthly = soldes.filter(s => s.month).sort((a, b) => b.month - a.month);
    const keep = [...(monthly.length ? [monthly[0]] : []), ...soldes.filter(s => !s.month)];
    return { cards: cardList, soldes: keep };
  }

  async function vImport(m) {
    const gtps = S.accounts.filter(a => a.kind === 'gtp');
    const banks = [...new Set(gtps.map(a => a.bank).filter(Boolean))];
    m.innerHTML = `
      <div class="alert info"><span>Choisissez un ou plusieurs fichiers Excel (fiches de rechargement, suivi AfricSika…). L’application y trouve les cartes clients (« numéro + nom ») et le dernier solde de chaque feuille. Rien n’est enregistré avant votre confirmation.</span></div>
      <div class="card pad section">
        ${field('Banque par défaut des cartes', `<select class="input" id="ibank">${options([['auto', 'Deviner d’après le nom du fichier'], ...banks.map(b => [b, b])], 'auto')}</select>`, 'Les feuilles nommées ECOBANK, GTCO, AYAEL CI… sont reconnues automatiquement.')}
        <label class="btn primary block" style="cursor:pointer">${icon('upload')} Choisir les fichiers Excel<input type="file" id="ifile" accept=".xlsx,.xls,.xlsm" multiple hidden></label>
      </div>
      <div id="ires"></div>`;
    m.querySelector('#ifile').onchange = async e => {
      const files = [...e.target.files];
      if (!files.length) return;
      const box = m.querySelector('#ires');
      box.innerHTML = '<div class="spinner"></div>';
      try {
        await loadXLSX();
        const all = { cards: new Map(), soldes: [] };
        for (const f of files) {
          const wb = window.XLSX.read(await f.arrayBuffer(), { type: 'array', cellDates: false });
          const sel = m.querySelector('#ibank').value;
          const n = norm(f.name);
          const fileBank = sel !== 'auto' ? sel : n.includes('GTCO') ? 'GTCO' : n.includes('AFRICSIKA') || n.includes('UBA') ? 'UBA' : n.includes('ECOBANK') ? 'ECOBANK' : n.includes('BNIFIC') ? 'BNIFIC' : (banks[0] || 'GTCO');
          const r = analyseWorkbook(wb, f.name, fileBank);
          r.cards.forEach(c => { const k = c.bank + '|' + c.card_number; if (!all.cards.has(k)) all.cards.set(k, c); });
          all.soldes.push(...r.soldes);
        }
        renderImport(box, [...all.cards.values()], all.soldes);
      } catch (err) { box.innerHTML = `<div class="alert bad">${esc(err.message || err)}</div>`; }
    };
  }

  function renderImport(box, cards, soldes) {
    const byBank = {};
    cards.forEach(c => { byBank[c.bank] = (byBank[c.bank] || 0) + 1; });
    const nInactive = cards.filter(c => c.active === false).length;
    const nPhone = cards.filter(c => c.phone).length;
    const nClients = new Set(cards.map(c => c.holder_name.replace(/[^A-Z0-9À-Ý]+/g, ' ').trim())).size;
    const targets = S.accounts.filter(a => a.active && a.kind !== 'caisse');
    const guessTarget = s => {
      const n = norm(s.sheet);
      if (n.includes('AYAEL CI')) return (targets.find(a => norm(a.name).includes('AYAEL CI')) || {}).id;
      if (n.includes('ESPECE')) return '';
      if (n.includes('ECOBANK')) return (targets.find(a => a.kind === 'gtp' && a.bank === 'ECOBANK') || {}).id;
      if (s.month) {
        if (norm(s.file).includes('AFRICSIKA')) return (targets.find(a => norm(a.name).includes('AFRICSIKA')) || {}).id;
        return (targets.find(a => a.kind === 'gtp' && a.bank === s.bank) || {}).id;
      }
      return '';
    };
    box.innerHTML = `
      ${section(`Cartes clients trouvées : ${cards.length}`, `<div class="card pad">
        <p style="margin-top:0">${Object.entries(byBank).map(([b, n]) => `<span class="tag brand">${esc(b)} : ${n}</span>`).join(' ')}</p>
        <p class="small">Environ <b>${nClients}</b> clients différents · ${nPhone} avec téléphone${nInactive ? ` · <b>${nInactive}</b> carte(s) expirée(s) ou remplacée(s), importée(s) comme inactives` : ''}.</p>
        <div class="scrollx" style="max-height:300px;overflow-y:auto;margin-bottom:12px"><table class="tbl"><thead><tr><th>N°</th><th>Nom</th><th>Banque</th><th>Téléphone</th><th>État</th></tr></thead><tbody>
          ${cards.slice(0, 300).map(c => `<tr><td class="num">${esc(c.card_number)}</td><td>${esc(c.holder_name)}</td><td>${esc(c.bank)}</td><td class="num">${esc(c.phone || '')}</td>
            <td>${c.active === false ? '<span class="tag">inactive</span>' : c.active ? '<span class="tag green">active</span>' : ''}</td></tr>`).join('')}</tbody></table></div>
        ${cards.length > 300 ? `<p class="small muted">… et ${cards.length - 300} autres.</p>` : ''}
        <button class="btn primary block" id="icards" ${cards.length ? '' : 'disabled'}>Importer ces ${cards.length} cartes et leurs clients</button>
        <p class="small muted">Rien n’est dupliqué : une carte déjà présente est complétée, un client déjà connu (même nom) reçoit ses nouvelles cartes. Les pièces d’identité, dates de naissance, e-mails et adresses ne sont pas repris.</p></div>`)}
      ${section('Soldes de départ', soldes.length ? `<div class="card pad">
        <p class="small muted" style="margin-top:0">Dernier solde lu dans chaque feuille. Choisissez le compte auquel l’appliquer : un ajustement « solde d’ouverture » mettra ce compte exactement à ce montant.</p>
        ${soldes.map((s, i) => `<div style="border-top:1px solid var(--line);padding:10px 0">
          <div style="display:flex;justify-content:space-between;gap:8px"><b>${esc(s.sheet)}</b><b class="num ${s.solde < 0 ? 'neg' : ''}">${fmt(s.solde)}</b></div>
          <div class="small muted">${esc(s.file)}${s.date ? ' · dernière date ' + fmtDate(s.date) : ''} · dernière ligne : ${esc(s.label.slice(0, 40))}</div>
          <select class="input" data-si="${i}" style="margin-top:8px">${options(targets.map(a => [a.id, `${a.name} (actuel ${fmt(S.bal.get(a.id) ?? 0)})`]), guessTarget(s), '— Ne pas utiliser —')}</select></div>`).join('')}
        <button class="btn primary block" id="isoldes" style="margin-top:12px">Enregistrer les soldes choisis</button></div>` : '<div class="list"><div class="empty">Aucune colonne « SOLDE GTP » trouvée.</div></div>')}`;
    const ic = box.querySelector('#icards');
    if (ic) ic.onclick = () => busy(ic, async () => {
      if (!await ask('Importer les clients ?', `<p>${cards.length} cartes (environ ${nClients} clients) vont être enregistrées dans la base de l’équipe.</p>`, 'Importer')) return;
      const tot = { cards_new: 0, cards_updated: 0, clients_new: 0, inactive: 0 };
      for (let i = 0; i < cards.length; i += 300) {
        ic.textContent = `Import en cours… ${Math.min(i + 300, cards.length)} / ${cards.length}`;
        const r = await rpc('import_clients', { rows: cards.slice(i, i + 300) });
        Object.keys(tot).forEach(k => { tot[k] += r[k] || 0; });
      }
      const msg = `${tot.clients_new} client(s) créé(s), ${tot.cards_new} carte(s) ajoutée(s), ${tot.cards_updated} déjà présente(s) mise(s) à jour, dont ${tot.inactive} inactive(s).`;
      toast(msg, 'ok');
      box.querySelector('#icards').insertAdjacentHTML('afterend', `<div class="alert ok" style="margin-top:10px">${esc(msg)}</div>`);
      return 'done';
    }).then(r => { if (r === 'done' && ic.isConnected) { ic.textContent = 'Import terminé ✓'; ic.disabled = true; } });
    const is = box.querySelector('#isoldes');
    if (is) is.onclick = async () => {
      const picks = [...box.querySelectorAll('[data-si]')].filter(s => s.value).map(s => ({ acc: acc(s.value), s: soldes[Number(s.dataset.si)] }));
      if (!picks.length) return toast('Choisissez au moins un compte.', 'bad');
      const dupl = picks.map(p => p.acc.id).filter((v, i, a) => a.indexOf(v) !== i);
      if (dupl.length) return toast('Un même compte est choisi deux fois.', 'bad');
      const ok = await ask('Enregistrer les soldes de départ ?', `<dl class="kv">${picks.map(p => `<dt>${esc(p.acc.name)}</dt><dd>${fmt(p.s.solde)}</dd>`).join('')}</dl>`, 'Enregistrer');
      if (!ok) return;
      busy(is, async () => {
        for (const p of picks) {
          await rpc('op_ajustement', { account_id: p.acc.id, target: p.s.solde, details: `Solde d’ouverture importé (${p.s.file} / ${p.s.sheet})` });
        }
        toast('Soldes de départ enregistrés.', 'ok');
        const d = await rpc('dashboard'); rememberBalances(d);
        is.textContent = 'Soldes enregistrés ✓'; is.disabled = true;
      });
    };
  }

  // ---------------------------------------------------------------- routes et actions globales
  const VIEWS = {
    accueil: { title: 'Accueil', render: vAccueil, live: true },
    journal: { title: 'Journal', render: vJournal, live: true, refresh: () => loadJournal() },
    clients: { title: 'Clients', render: vClients },
    cloture: { title: 'Clôture du soir', render: vCloture },
    tableau: { title: 'Tableau de bord', render: vTableau, live: true, mgr: true },
    clotures: { title: 'Clôtures', render: vClotures, live: true, mgr: true },
    plus: { title: 'Plus', render: vPlus, mgr: true },
    rapports: { title: 'Rapports', render: vRapports, mgr: true },
    cartes: { title: 'Cartes clients', render: vCartes, mgr: true },
    doublons: { title: 'Clients en double', render: vDoublons, dir: true },
    commissions: { title: 'Commissions à vérifier', render: vCommissions, need: 'can_recharge' },
    equipe: { title: 'Équipe', render: vEquipe, live: true, dir: true },
    comptes: { title: 'Comptes et commissions', render: vComptes, dir: true },
    produits: { title: 'Produits de caisse', render: vProduits, dir: true },
    import: { title: 'Importer depuis Excel', render: vImport, dir: true },
    agences: { title: 'Agences', render: vAgences, live: true, dir: true },
  };

  // Après une saisie : mise à jour immédiate de l'écran (sans attendre le signal temps réel).
  function refreshNow() {
    if (S.view === 'cartes' || S.view === 'clients') { route(); return; }
    refresh();
  }

  // Ouvre une nouvelle opération pour la carte choisie (client retrouvé par son n° de carte).
  async function venteForCard(card) {
    const res = await rpc('search_clients', { q: card.card_number, limit: 5 });
    const cl = res.find(x => x.id === card.client_id) || res[0] || null;
    sheetVente({ client: cl, card: cl ? (cl.cards || []).find(k => k.id === card.id) || card : null });
  }

  const DO = {
    vente: () => sheetVente(),
    depense: () => sheetDepense(),
    appro: () => sheetAppro(),
    transfert: () => sheetTransfert(),
    ajust: () => sheetAjust(),
    export: () => sheetExport(),
    reversement: () => sheetReversement(),
    company: () => sheetCompany(),
    me: () => sheetMe(),
    newcard: () => sheetCard(null, {}),
  };

  document.addEventListener('click', e => {
    const t = e.target.closest('[data-do],[data-op],[data-acc],[data-client],[data-cardedit],[data-go],[data-liaison]');
    if (!t) return;
    const d = t.dataset;
    if (t.tagName === 'A' && (d.do || d.go)) e.preventDefault();
    if (d.go) { if (location.hash === '#/' + d.go) route(); else location.hash = '#/' + d.go; return; }
    if (d.op) return opDetail(Number(d.op));
    if (d.liaison) return sheetLiaison(Number(d.liaison));
    if (d.acc) return isMgr() && sheetAccountView(Number(d.acc));
    if (d.client) { const c = S.clients.get(Number(d.client)); return c && sheetClient(c); }
    if (d.cardedit) { const c = S.cards.get(Number(d.cardedit)); return c && sheetCard(c, { onRecharge: (S.me.can_recharge || isMgr()) ? venteForCard : null }); }
    if (d.do && DO[d.do]) DO[d.do](t);
  });
  // Toucher l'onglet déjà ouvert le recharge.
  document.addEventListener('click', e => {
    const a = e.target.closest('.tabbar a');
    if (a && a.getAttribute('href') === location.hash) { e.preventDefault(); route(); }
  });

  boot();
})();
