(function () {
  const DIR_URL = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vT9_K7L4VsiykK_3wQT4I5vAyzLIdqjn9meayzoaQmLfa_IWmrNc9_C511zSVxqgAhMoCR8a1Xv_YWI/pub?output=csv';

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

  // ── Main ──────────────────────────────────────
  async function applyBrokerTheme() {
    const broker = new URLSearchParams(window.location.search).get('broker');
    if (!broker) return;

    try {
      const resp = await fetch(DIR_URL);
      if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
      const rows = parseCSV(await resp.text());
      if (rows.length < 2) return;

      const headers = rows[0].map(v => v.trim().toLowerCase());
      const col = name => headers.findIndex(v => v === name.toLowerCase());

      const iBroker   = col('broker id');
      const iPrimary  = col('primary color');
      const iSecond   = col('secondary color');
      const iLogo     = col('logo url');
      const iBrand    = col('brand name');

      // Find first row belonging to this broker (any active company row works for theme data)
      const row = rows.slice(1).find(r =>
        (r[iBroker] || '').trim().toLowerCase() === broker.toLowerCase()
      );
      if (!row) { console.warn('[Theme] No row found for broker:', broker); return; }

      const primary = iPrimary >= 0 ? (row[iPrimary] || '').trim() : '';
      const second  = iSecond  >= 0 ? (row[iSecond]  || '').trim() : '';
      const logoUrl = iLogo    >= 0 ? (row[iLogo]    || '').trim() : '';
      const brand   = iBrand   >= 0 ? (row[iBrand]   || '').trim() : '';

      const root = document.documentElement;

      // Apply primary color + derived variants
      if (isHex(primary)) {
        root.style.setProperty('--forest',    primary);
        root.style.setProperty('--forest-d',  darken(primary, 0.15));
        root.style.setProperty('--forest-lt', lighten(primary, 0.92));
        // Update theme-color meta if present
        const tm = document.querySelector('meta[name="theme-color"]');
        if (tm) tm.content = primary;
      }

      // Apply secondary color + derived variants
      if (isHex(second)) {
        root.style.setProperty('--teal',     second);
        root.style.setProperty('--teal-lt',  lighten(second, 0.92));
        root.style.setProperty('--teal-mid', hexToRgba(second, 0.13));
      }

      // Swap logos
      if (logoUrl) {
        // img tags: TBG CDN images and any tagged with data-broker-logo
        document.querySelectorAll(
          'img[src*="filesafe.space"], img[src*="thrivebg"], img[data-broker-logo]'
        ).forEach(img => {
          img.src = logoUrl;
          if (brand) img.alt = brand;
        });
        // Logo-mark divs (letter mark used in landing nav/footer)
        document.querySelectorAll('.logo-mark').forEach(el => {
          el.innerHTML = `<img src="${logoUrl}" alt="${brand || broker}" style="height:100%;width:100%;object-fit:contain;border-radius:inherit">`;
        });
      }

      // Swap brand name in explicitly tagged elements only
      if (brand) {
        document.querySelectorAll('[data-broker-brand]').forEach(el => {
          el.textContent = brand;
        });
      }

      console.log('[Theme] Applied broker theme:', broker,
        { primary: primary || '(default)', second: second || '(default)', logoUrl: logoUrl || '(none)' });

    } catch (e) {
      console.warn('[Theme] Failed to apply broker theme:', e.message);
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', applyBrokerTheme);
  } else {
    applyBrokerTheme();
  }
})();
