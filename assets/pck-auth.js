/* ═══════════════════════════════════════════════════════════════════════════
   pck-auth.js — หน้ากั้นก่อนเข้าใช้งาน (Access Gate)
   Molecular Lab · Phrachomklao Hospital
   ───────────────────────────────────────────────────────────────────────────
   ⚠️ นี่คือ "ป้ายกั้น" ไม่ใช่ "ระบบความปลอดภัย"
     • กันคนที่เปิดมาเจอโดยบังเอิญ — ได้
     • กันคนที่ตั้งใจเจาะ — ไม่ได้ (เปิด DevTools ลบ overlay ก็ผ่าน)
     • ข้อมูลใน Google Sheet ยังเปิดอยู่ ใครรู้ URL ของ Apps Script ก็ดึงได้

   วิธีใช้ : ใส่บรรทัดนี้ใน <head> ของทุกหน้า
     <script src="/PCK/assets/pck-auth.js?v=1"></script>
   ═══════════════════════════════════════════════════════════════════════════ */
(function () {
  'use strict';

  /* ── ตั้งค่า ────────────────────────────────────────────────────────────
     เปลี่ยนรหัสผ่าน : เปิดเว็บ กด F12 → Console → พิมพ์
         PCK_AUTH.hash('รหัสใหม่ของคุณ').then(console.log)
     แล้วเอาค่าที่ได้มาใส่ใน HASH ด้านล่าง                                */
  var CFG = {
    // SHA-256 ของรหัสผ่าน — ค่าเริ่มต้นคือ  PCK-Lab-2569
    HASH: '52987c732036566cdede8d497a781a47ad5bd49c49a5fe1e39af3bf81ee15ad8',

    // จำกัดโดเมนอีเมล — [] = รับทุกโดเมน | ตัวอย่าง ['gmail.com','moph.go.th']
    DOMAINS: [],

    // จำกัดรายชื่ออีเมล — [] = ไม่จำกัด | ตัวอย่าง ['somchai@gmail.com']
    EMAILS: [],

    HOURS: 12,                 // จำการเข้าสู่ระบบไว้กี่ชั่วโมง
    KEY: 'pck-auth-session'
  };

  var LS = {
    get: function () { try { return JSON.parse(localStorage.getItem(CFG.KEY) || 'null'); } catch (e) { return null; } },
    set: function (v) { try { localStorage.setItem(CFG.KEY, JSON.stringify(v)); } catch (e) {} },
    del: function () { try { localStorage.removeItem(CFG.KEY); } catch (e) {} }
  };

  function sha256(text) {
    if (!window.crypto || !crypto.subtle) return Promise.reject(new Error('เบราว์เซอร์ไม่รองรับ'));
    return crypto.subtle.digest('SHA-256', new TextEncoder().encode(text))
      .then(function (buf) {
        return Array.prototype.map.call(new Uint8Array(buf), function (b) {
          return b.toString(16).padStart(2, '0');
        }).join('');
      });
  }

  function session() {
    var s = LS.get();
    if (s && s.email && s.exp && Date.now() < s.exp) return s;
    if (s) LS.del();
    return null;
  }

  function emailOk(e) {
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e)) return 'รูปแบบอีเมลไม่ถูกต้อง';
    if (CFG.EMAILS.length && CFG.EMAILS.indexOf(e) < 0) return 'อีเมลนี้ไม่มีสิทธิ์เข้าใช้งาน';
    if (CFG.DOMAINS.length && CFG.DOMAINS.indexOf(e.split('@')[1]) < 0)
      return 'อนุญาตเฉพาะอีเมล @' + CFG.DOMAINS.join(' หรือ @');
    return '';
  }

  /* ── ใส่ favicon อัตโนมัติ ไม่ต้องแก้ HTML ทีละหน้า ─────────────────── */
  (function () {
    if (document.querySelector('link[rel="icon"]')) return;
    var l = document.createElement('link');
    l.rel = 'icon'; l.type = 'image/png'; l.href = '/PCK/assets/favicon-32.png';
    (document.head || document.documentElement).appendChild(l);
  })();

  /* ── ซ่อนเนื้อหาทันทีที่สคริปต์ทำงาน ────────────────────────────────── */
  var hideCSS = document.createElement('style');
  hideCSS.id = 'pck-gate-hide';
  hideCSS.textContent = 'body{visibility:hidden}#pck-gate{visibility:visible}';
  (document.head || document.documentElement).appendChild(hideCSS);

  function reveal() {
    var s = document.getElementById('pck-gate-hide');
    if (s) s.remove();
    var g = document.getElementById('pck-gate');
    if (g) g.remove();
  }

  /* ── ปุ่มออกจากระบบ — แทรกใน topbar ถ้ามี ───────────────────────────── */
  function addLogout(email) {
    var tries = 0;
    (function place() {
      var host = document.querySelector('.pck-topbar-right');
      if (!host) { if (++tries < 40) return setTimeout(place, 150); return; }
      if (document.getElementById('pck-logout')) return;
      var wrap = document.createElement('span');
      wrap.id = 'pck-logout';
      wrap.style.cssText = 'display:flex;align-items:center;gap:8px';
      wrap.innerHTML =
        '<span title="' + email + '" style="font-size:11.5px;color:var(--text3);max-width:170px;' +
        'overflow:hidden;text-overflow:ellipsis;white-space:nowrap">' + email + '</span>' +
        '<button class="pck-btn icon" title="ออกจากระบบ" aria-label="ออกจากระบบ">⏻</button>';
      wrap.querySelector('button').addEventListener('click', function () {
        if (confirm('ออกจากระบบ?')) { LS.del(); location.reload(); }
      });
      host.insertBefore(wrap, host.firstChild);
    })();
  }

  /* ── หน้าจอเข้าสู่ระบบ ───────────────────────────────────────────────── */
  function gate() {
    var g = document.createElement('div');
    g.id = 'pck-gate';
    g.innerHTML =
      '<style>' +
      '#pck-gate{position:fixed;inset:0;z-index:99999;display:grid;place-items:center;padding:20px;' +
        'background:linear-gradient(140deg,#fdf0f4 0%,#fdf8f9 55%,#f4ecf5 100%);' +
        "font-family:'Sarabun',Tahoma,sans-serif;color:#2d2028;overflow:auto}" +
      '@media (prefers-color-scheme:dark){#pck-gate{background:linear-gradient(140deg,#241a20,#191317 60%);color:#f4e9ee}' +
        '#pck-gate .box{background:#231b20;border-color:#3a2d34}' +
        '#pck-gate input{background:#2b2127;border-color:#3a2d34;color:#f4e9ee}' +
        '#pck-gate .note{background:#3a2f14;color:#f0d060;border-color:#6b5620}}' +
      '#pck-gate .box{width:100%;max-width:380px;background:#fff;border:1px solid #eedde4;border-radius:16px;' +
        'padding:30px 28px;box-shadow:0 8px 40px rgba(180,80,100,.14)}' +
      '#pck-gate .mark{display:block;width:100%;max-width:180px;height:auto;margin:0 auto 18px}' +
      '#pck-gate .sub{font-size:12px;text-align:center;color:#896b74;margin-bottom:22px}' +
      '#pck-gate label{display:block;font-size:11px;font-weight:700;letter-spacing:.4px;' +
        'text-transform:uppercase;color:#896b74;margin-bottom:5px}' +
      '#pck-gate input{width:100%;padding:10px 13px;border:1px solid #eedde4;border-radius:9px;' +
        'font-size:14px;font-family:inherit;margin-bottom:14px;background:#fff}' +
      '#pck-gate input:focus{outline:2px solid #c9687e;outline-offset:1px;border-color:#c9687e}' +
      '#pck-gate button.go{width:100%;padding:11px;border:0;border-radius:99px;background:#8b3a52;color:#fff;' +
        'font-size:14px;font-weight:700;font-family:inherit;cursor:pointer;transition:background .15s}' +
      '#pck-gate button.go:hover{background:#c9687e}' +
      '#pck-gate button.go:disabled{opacity:.6;cursor:wait}' +
      '#pck-gate .err{font-size:12.5px;color:#721c24;background:#f8d7da;border:1px solid #e8a8b0;' +
        'border-radius:9px;padding:9px 12px;margin-bottom:14px;display:none}' +
      '#pck-gate .note{font-size:11px;line-height:1.55;color:#8a6000;background:#fff8e0;' +
        'border:1px solid #f0d060;border-radius:9px;padding:9px 12px;margin-top:18px}' +
      '</style>' +
      '<form class="box" autocomplete="on">' +
        '<img class="mark" src="/PCK/assets/logo-full.png" alt="Molecular Lab PCK">' +
        '<div class="sub">Phrachomklao Hospital · ระบบประกันคุณภาพห้องปฏิบัติการ</div>' +
        '<div class="err" id="pck-err" role="alert"></div>' +
        '<label for="pck-em">อีเมล</label>' +
        '<input id="pck-em" type="email" inputmode="email" autocomplete="username" required placeholder="name@example.com">' +
        '<label for="pck-pw">รหัสผ่าน</label>' +
        '<input id="pck-pw" type="password" autocomplete="current-password" required placeholder="••••••••">' +
        '<button class="go" type="submit">เข้าสู่ระบบ</button>' +
      '</form>';
    document.body.appendChild(g);

    var form = g.querySelector('form');
    var em   = g.querySelector('#pck-em');
    var pw   = g.querySelector('#pck-pw');
    var err  = g.querySelector('#pck-err');
    var btn  = g.querySelector('button.go');

    try { em.value = (LS.get() || {}).email || localStorage.getItem('pck-last-email') || ''; } catch (e) {}
    setTimeout(function () { (em.value ? pw : em).focus(); }, 60);

    function fail(msg) {
      err.textContent = msg; err.style.display = 'block';
      btn.disabled = false; btn.textContent = 'เข้าสู่ระบบ';
      pw.value = ''; pw.focus();
    }

    form.addEventListener('submit', function (ev) {
      ev.preventDefault();
      var email = em.value.trim().toLowerCase();
      var bad = emailOk(email);
      if (bad) return fail(bad);

      btn.disabled = true; btn.textContent = 'กำลังตรวจสอบ…';
      sha256(pw.value)
        .then(function (h) {
          if (h !== CFG.HASH) return fail('รหัสผ่านไม่ถูกต้อง');
          LS.set({ email: email, exp: Date.now() + CFG.HOURS * 3600000, at: new Date().toISOString() });
          try { localStorage.setItem('pck-last-email', email); } catch (e) {}
          reveal();
          addLogout(email);
          window.PCK_USER = email;
        })
        .catch(function (e) { fail('ตรวจสอบไม่สำเร็จ: ' + e.message); });
    });
  }

  function start() {
    var s = session();
    if (s) { reveal(); window.PCK_USER = s.email; addLogout(s.email); }
    else gate();
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start);
  else start();

  window.PCK_AUTH = {
    hash: sha256,
    logout: function () { LS.del(); location.reload(); },
    user: function () { var s = session(); return s ? s.email : null; }
  };
})();
