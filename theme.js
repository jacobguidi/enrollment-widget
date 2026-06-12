(function () {
  // ── Broker config — fetched from repo JSON ────
  const BROKERS_URL = 'https://raw.githubusercontent.com/jacobguidi/enrollment-widget/main/config/brokers.json';

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

  // ── Fetch broker config from JSON ────────────
  async function fetchBrokerConfig(broker) {
    const resp = await fetch(BROKERS_URL);
    if (!resp.ok) throw new Error(`Brokers fetch failed: HTTP ${resp.status}`);
    const data = await resp.json();
    const cfg = data[broker.toLowerCase()];
    if (!cfg) { console.warn('[Theme] No broker config found for:', broker); return null; }
    return {
      primary           : cfg.primary           || '',
      secondary         : cfg.secondary         || '',
      logoUrl           : cfg.logoUrl           || '',
      brand             : cfg.brand             || '',
      enrollUrl         : cfg.enrollUrl         || '',
      customUrl         : cfg.customizeUrl      || cfg.enrollUrl || '',
      advisorBookingUrl : cfg.advisorBookingUrl || '',
      supportEmail      : cfg.supportEmail      || '',
      supportPhone      : cfg.supportPhone      || '',
      websiteUrl        : cfg.websiteUrl        || '',
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
