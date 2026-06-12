(function () {
  // ── Directory: company list ───────────────────
  const DIR_URL = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vT9_K7L4VsiykK_3wQT4I5vAyzLIdqjn9meayzoaQmLfa_IWmrNc9_C511zSVxqgAhMoCR8a1Xv_YWI/pub?output=csv';

  // ── Broker Config tab: one row per broker ─────
  // Publish separately: Sheets → File → Share → Publish to web → "Broker Config" tab → CSV
  const BROKER_CONFIG_URL = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vT9_K7L4VsiykK_3wQT4I5vAyzLIdqjn9meayzoaQmLfa_IWmrNc9_C511zSVxqgAhMoCR8a1Xv_YWI/pub?gid=606899757&single=true&output=csv';

  // ── Color helpers ─────────────────────────────
  function hexToRgb(h) {
    return [parseInt(h.slice(1,3),16), parseInt(h.slice(3,5),16), parseInt(h.slice(5,7),16)];
  }
  function rgbToHex(r,g,b) {
    return '#'+[r,g,b].map(v => Math.max(0,Math.min(255,Math.round(v))).toString(16).padStart(2,'0')).join('');
  }
  function darken(hex, amt)  { const [r,g,b]=hexToRgb(hex); return rgbToHex(r*(1-amt),g*(1-amt),b*(1-amt)); }
  function lighten(hex, amt) { const [r,g,b]=hexToRgb(hex); return rgbToHex(r+(255-r)*amt,g+(255-g)*amt,b+(255-b)*amt); }
  function hexToRgba(hex, a) { const [r,g,b]=hexToRgb(hex); return `rgba(${r},${g},${b},${a})`; }
  function isHex(v) { return /^#[0-9a-f]{6}$/i.test((v||'').trim()); }

  // ── Minimal CSV parser ────────────────────────
  function parseCSV(text) {
    return text.trim().replace(/\r/g,'').split('\n').map(line => {
      const row=[]; let cur='', inQ=false;
      for (let i=0; i<line.length; i++) {
        const c = line[i];
        if (c==='"') { inQ=!inQ; }
        else if (c===',' && !inQ) { row.push(cur.trim()); cur=''; }
        else { cur+=c; }
      }
      row.push(cur.trim());
      return row;
    });
  }

  // ── Fetch broker config row ───────────────────
  async function fetchBrokerConfig(broker) {
    if (!BROKER_CONFIG_URL || BROKER_CONFIG_URL.startsWith('YOUR_')) {
      console.warn('[Theme] Broker Config URL not set — skipping theme');
      return null;
    }
    const resp = await fetch(BROKER_CONFIG_URL);
    if (!resp.ok) throw new Error(`Broker Config fetch failed: HTTP ${resp.status}`);
    const rows = parseCSV(await resp.text());
    if (rows.length < 2) return null;

    const headers = rows[0].map(v => v.trim().toLowerCase());
    const col = name => headers.findIndex(v => v === name.toLowerCase());

    const iBroker    = col('broker id');
    const iPrimary   = col('primary');
    const iSecond    = col('secondary');
    const iLogo      = col('logo url');
    const iBrand     = col('brand name');
    const iEnrollUrl = col('enroll url');
    const iCustomUrl = col('customize url');
    const iAdvisor   = col('advisor booking url');
    const iEmail     = col('support email');
    const iPhone     = col('support phone');
    const iWebsite   = col('website url');

    const row = rows.slice(1).find(r =>
      (r[iBroker] || '').trim().toLowerCase() === broker.toLowerCase()
    );
    if (!row) { console.warn('[Theme] No broker config row found for:', broker); return null; }

    const enrollUrl = iEnrollUrl >= 0 ? (row[iEnrollUrl] || '').trim() : '';
    return {
      primary   : iPrimary   >= 0 ? (row[iPrimary]   || '').trim() : '',
      secondary : iSecond    >= 0 ? (row[iSecond]    || '').trim() : '',
      logoUrl   : iLogo      >= 0 ? (row[iLogo]      || '').trim() : '',
      brand     : iBrand     >= 0 ? (row[iBrand]     || '').trim() : '',
      enrollUrl,
      customUrl         : iCustomUrl >= 0 ? (row[iCustomUrl] || '').trim() : enrollUrl,
      advisorBookingUrl : iAdvisor   >= 0 ? (row[iAdvisor]   || '').trim() : '',
      supportEmail      : iEmail     >= 0 ? (row[iEmail]     || '').trim() : '',
      supportPhone      : iPhone     >= 0 ? (row[iPhone]     || '').trim() : '',
      websiteUrl        : iWebsite   >= 0 ? (row[iWebsite]   || '').trim() : '',
    };
  }

  // ── Apply theme to page ───────────────────────
  function applyTheme(cfg) {
    const root = document.documentElement;
    const pageParams = new URLSearchParams(window.location.search);

    if (isHex(cfg.primary) && !pageParams.get('primary')) {
      root.style.setProperty('--forest',    cfg.primary);
      root.style.setProperty('--forest-d',  darken(cfg.primary, 0.15));
      root.style.setProperty('--forest-lt', lighten(cfg.primary, 0.92));
      const tm = document.querySelector('meta[name="theme-color"]');
      if (tm) tm.content = cfg.primary;
    }
    if (isHex(cfg.secondary) && !pageParams.get('secondary')) {
      root.style.setProperty('--teal',     cfg.secondary);
      root.style.setProperty('--teal-lt',  lighten(cfg.secondary, 0.92));
      root.style.setProperty('--teal-mid', hexToRgba(cfg.secondary, 0.13));
    }
    if (cfg.logoUrl) {
      document.querySelectorAll(
        'img[src*="filesafe.space"]:not([data-company-logo]), img[src*="thrivebg"]:not([data-company-logo]), img[data-broker-logo]'
      ).forEach(img => {
        img.src = cfg.logoUrl;
        if (cfg.brand) img.alt = cfg.brand;
      });
      document.querySelectorAll('.logo-mark').forEach(el => {
        el.innerHTML = `<img src="${cfg.logoUrl}" alt="${cfg.brand}" style="height:100%;width:100%;object-fit:contain;border-radius:inherit">`;
      });
      const desktop = document.getElementById('brokerLogoDesktop');
      if (desktop) desktop.style.display = 'flex';
    }
    if (cfg.brand) {
      document.querySelectorAll('[data-broker-brand]').forEach(el => {
        el.textContent = cfg.brand;
      });
    }
    if (cfg.enrollUrl) {
      document.querySelectorAll('[data-broker-cta="enroll"]').forEach(a => {
        const url = new URL(cfg.enrollUrl);
        pageParams.forEach((v, k) => url.searchParams.set(k, v));
        a.href = url.toString();
      });
    }
    if (cfg.customUrl) {
      document.querySelectorAll('[data-broker-cta="customize"]').forEach(a => {
        const url = new URL(cfg.customUrl);
        pageParams.forEach((v, k) => url.searchParams.set(k, v));
        a.href = url.toString();
      });
    }
    if (cfg.advisorBookingUrl) {
      document.querySelectorAll('[data-broker-cta="advisor"]').forEach(a => {
        a.href = cfg.advisorBookingUrl;
      });
    }
    if (cfg.supportEmail) {
      document.querySelectorAll('[data-broker-email]').forEach(a => {
        a.href = 'mailto:' + cfg.supportEmail;
        if (a.textContent.includes('@')) a.textContent = cfg.supportEmail;
      });
    }
    if (cfg.supportPhone) {
      document.querySelectorAll('[data-broker-phone]').forEach(a => {
        a.href = 'tel:' + cfg.supportPhone.replace(/[^0-9]/g, '');
        if (a.textContent.trim().startsWith('(') || a.textContent.trim().match(/^\d/)) {
          a.textContent = cfg.supportPhone;
        }
      });
    }
    if (cfg.websiteUrl) {
      document.querySelectorAll('[data-broker-website]').forEach(a => {
        a.href = cfg.websiteUrl;
      });
    }
  }

  // ── Main ──────────────────────────────────────
  async function applyBrokerTheme() {
    const broker = new URLSearchParams(window.location.search).get('broker');
    if (!broker) return;
    try {
      const cfg = await fetchBrokerConfig(broker);
      if (cfg) {
        applyTheme(cfg);
        console.log('[Theme] Applied:', broker, cfg);
      }
    } catch (e) {
      console.warn('[Theme] Failed:', e.message);
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', applyBrokerTheme);
  } else {
    applyBrokerTheme();
  }
})();
