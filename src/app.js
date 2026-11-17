(() => {
'use strict';
const $ = id => document.getElementById(id);
const VER = '1.0';
const PT = UD.PT;
const save = (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) {} };
const load = (k, d) => { try { const v = localStorage.getItem(k); return v == null ? d : JSON.parse(v); } catch (e) { return d; } };
const DEF = { bg: '#101217', text: '#ffffff', accent: '#ffff00', images: true };
const st = Object.assign({}, DEF, load('ud-settings', {}));
let scope = load('ud-scope2', 'all');
let busy = false;

// ---------- status / results ----------
const status = $('status'), res = $('res');
const say = (m, cls) => { status.textContent = m || ''; status.className = 'status' + (cls ? ' ' + cls : ''); };
const prog = f => { $('prog').hidden = f == null; if (f != null) $('prog').firstElementChild.style.width = Math.round(f * 100) + '%'; };
function line(cls, html) { const d = document.createElement('div'); d.className = 'ln ' + cls; d.innerHTML = html; res.appendChild(d); return d; }
function slideLinks(ids, nums) {
  return ids.map((id, k) => '<a data-id="' + id + '">' + nums[k] + '</a>').join('، ');
}
res.addEventListener('click', e => {
  const a = e.target.closest('a[data-id]'); if (!a) return;
  PowerPoint.run(async ctx => { ctx.presentation.setSelectedSlides([a.dataset.id]); await ctx.sync(); }).catch(() => {});
});

// ---------- Office helpers ----------
const run = f => PowerPoint.run(f);
const timeout = (p, ms) => Promise.race([p, new Promise((_, rej) => setTimeout(() => rej(new Error('timeout')), ms))]);
async function targetIds() {
  return run(async ctx => {
    const col = scope === 'sel' ? ctx.presentation.getSelectedSlides() : ctx.presentation.slides;
    col.load('items/id'); await ctx.sync();
    return col.items.map(s => s.id);
  });
}
async function allIds() { return run(async ctx => { const s = ctx.presentation.slides; s.load('items/id'); await ctx.sync(); return s.items.map(x => x.id); }); }
async function exportSlide(id) { return run(async ctx => { const r = ctx.presentation.slides.getItem(id).exportAsBase64(); await ctx.sync(); return r.value; }); }
async function deleteSlide(id) { return run(async ctx => { ctx.presentation.slides.getItem(id).delete(); await ctx.sync(); }); }
async function insertAfter(b64, afterId) {
  const tryIns = async target => run(async ctx => {
    ctx.presentation.insertSlidesFromBase64(b64, { formatting: 'UseDestinationTheme', targetSlideId: target });
    await ctx.sync();
  });
  try { await tryIns(afterId); }
  catch (e) { await tryIns(String(afterId).split('#')[0] + '#'); }
  const ids = await allIds();
  return ids[ids.indexOf(afterId) + 1];
}
async function isDarkDone(id) {
  return run(async ctx => {
    const t = ctx.presentation.slides.getItem(id).tags; t.load('items/key,items/value'); await ctx.sync();
    return t.items.some(x => x.key.toUpperCase() === 'UULADARK');
  }).catch(() => false);
}
async function markDark(id) { return run(async ctx => { ctx.presentation.slides.getItem(id).tags.add('UULADARK', VER); await ctx.sync(); }).catch(() => {}); }

// refine the layout with real text heights; returns the lowest bottom (pt)
async function refine(id, plan) {
  return run(async ctx => {
    const shapes = ctx.presentation.slides.getItem(id).shapes;
    shapes.load('items/id'); await ctx.sync();
    const L = plan.items.map(s => ({ s, sh: shapes.items[s.idx] }));
    if (L.some(x => !x.sh)) return plan.bottom / PT;
    const texts = L.filter(x => x.s.kind === 'text');
    for (const x of L) if (x.s.kind !== 'text') {
      x.sh.left = x.s.x / PT;
      if (x.s.kind !== 'table') { x.sh.width = Math.max(1, x.s.w / PT); x.sh.height = Math.max(1, x.s.h / PT); }
    }
    for (const x of texts) { x.sh.left = x.s.x / PT; x.sh.width = x.s.w / PT; x.sh.textFrame.autoSizeSetting = 'AutoSizeNone'; }
    await ctx.sync();
    for (const x of texts) { x.sh.textFrame.autoSizeSetting = 'AutoSizeShapeToFitText'; x.sh.textFrame.verticalAlignment = 'Top'; }
    await ctx.sync();
    L.forEach(x => x.sh.load('height'));
    await ctx.sync();
    const H = x => x.s.kind === 'text' ? x.sh.height : Math.max(x.sh.height || 0, x.s.h / PT);
    const M = UD.MARGIN / PT, G = UD.GAP / PT;
    const title = L.find(x => x.s.col === 'title');
    let start = M, bottom = M;
    if (title) { title.sh.top = M; start = M + H(title) + G; bottom = M + H(title); }
    const y = { full: M, left: start, right: start };
    for (const x of L) {
      if (x.s.col === 'title') continue;
      x.sh.top = y[x.s.col];
      y[x.s.col] += H(x) + G;
      bottom = Math.max(bottom, y[x.s.col] - G);
    }
    await ctx.sync();
    return bottom;
  });
}

// ---------- dark (splits at the separators first, then converts each part) ----------
const LEVELS = ['3600', '3200', '3000'];
// if PowerPoint refuses the converted slide, retry without parts of the conversion (also tells us which part)
const SKIPS = [[], ['layout'], ['images'], ['tables'], ['shapes'], ['text'], ['layout', 'images', 'tables', 'shapes']];
const SKIP_AR = { layout: 'الترتيب', images: 'تلوين الصور', tables: 'الجداول', shapes: 'الأشكال', text: 'النصوص' };
async function doDark(ids) {
  const done = [], small = [], over = [], gray = [], skipped = [], errs = [];
  let srcCount = 0; const fallbacks = [];
  const ctxBase = { bg: st.bg.slice(1).toUpperCase(), text: st.text.slice(1).toUpperCase(), accent: st.accent.slice(1).toUpperCase(), title: '4800', font: 'UULA Sans Medium', images: st.images };
  for (let k = 0; k < ids.length; k++) {
    prog(k / ids.length); say('تحويل الشريحة ' + (k + 1) + ' من ' + ids.length + '…');
    const id = ids[k];
    let stage = 'قراءة الشريحة', origB64 = null;
    try {
      if (await isDarkDone(id)) { skipped.push(id); continue; }
      origB64 = await exportSlide(id);
      const pkg = await UD.openPackage(origB64);
      const slideHpt = pkg.slideH / PT;
      const ctx0 = { scheme: UD.readTheme(pkg.files) };
      const sec = UD.sections(UD.analyze(UD.parseXml(pkg.files.slide), pkg, ctx0));
      const parts = sec.count && sec.parts.length ? sec.parts.map(p => UD.sectionXml(pkg.files.slide, pkg, ctx0, p)) : [pkg.files.slide];
      let prev = id;
      for (const base of parts) {
        let newId = null, lvl = 0, bottom = 0, rep = null;
        for (lvl = 0; lvl < LEVELS.length; lvl++) {
          stage = 'إدراج الشريحة';
          let d = null, lastErr = null;
          for (const sk of SKIPS) {
            d = UD.darkXml(base, pkg, Object.assign({}, ctxBase, { body: LEVELS[lvl], skip: sk }));
            try { newId = await insertAfter(await UD.buildPackage(pkg, d.xml), prev); lastErr = null; if (sk.length) fallbacks.push([sk, null]); d.skipped = sk; break; }
            catch (e) { lastErr = e; }
          }
          if (lastErr) throw lastErr;
          rep = d.report;
          stage = 'ترتيب المحتوى';
          bottom = await refine(newId, d.plan);
          if (bottom <= slideHpt - UD.MARGIN / PT + 0.5 || lvl === LEVELS.length - 1) break;
          await deleteSlide(newId);
        }
        if (fallbacks.length && fallbacks[fallbacks.length - 1][1] === null) fallbacks[fallbacks.length - 1][1] = newId;
        await markDark(newId);
        prev = newId;
        done.push(newId);
        if (lvl > 0) small.push([newId, LEVELS[lvl] / 100]);
        if (bottom > slideHpt - UD.MARGIN / PT + 0.5) over.push(newId);
        if (rep && rep.duotone) gray.push(newId);
      }
      stage = 'حذف الأصل';
      await deleteSlide(id);
      srcCount++;
    } catch (e) {
      let extra = '';
      if (stage === 'إدراج الشريحة' && origB64) {
        // diagnosis: can the untouched slide be inserted? then the problem is in the converted XML
        try { const t = await insertAfter(origB64, id); await deleteSlide(t); extra = ' — الأصل يُدرج بنجاح، المشكلة في الملف المحوَّل'; }
        catch (e2) { extra = ' — حتى الأصل لا يُدرج: ' + errText(e2); }
      }
      errs.push(stage + ': ' + errText(e) + extra);
    }
  }
  prog(null);
  const all = await allIds(), num = i => all.indexOf(i) + 1;
  res.innerHTML = '';
  if (done.length) line('ok', '✓ ' + cnt(srcCount) + ' ← ' + cnt(done.length) + ' داكنة: ' + slideLinks(done, done.map(num)));
  if (small.length) line('warn', '• صُغّر الخط: ' + small.map(([i, s]) => '<a data-id="' + i + '">' + num(i) + '</a> (' + s + ')').join('، '));
  if (over.length) line('err', '• ممتلئة، قسّمها يدويًا: ' + slideLinks(over, over.map(num)));
  if (gray.length) line('warn', '• صور فقدت ألوانها: ' + slideLinks(gray, gray.map(num)));
  if (skipped.length) line('warn', '• محوّلة سابقًا: ' + cnt(skipped.length));
  fallbacks.forEach(([sk, i]) => { if (i) line('warn', '• شريحة <a data-id="' + i + '">' + num(i) + '</a> حُوّلت بدون: ' + sk.map(k => SKIP_AR[k]).join('، ')); });
  errs.forEach(m => line('err', '✗ ' + esc(m)));
  if (done.length) line('', '<span class="note">يلزم تثبيت خط UULA Sans Medium على الجهاز.</span>');
  say(done.length ? 'تم التحويل.' : 'لم يتم تحويل أي شريحة.');
}
const errText = e => { const loc = e && e.debugInfo && e.debugInfo.errorLocation; return (e && (e.code || e.message) || String(e)) + (loc ? ' (' + loc + ')' : ''); };
const cnt = n => n === 1 ? 'شريحة واحدة' : n === 2 ? 'شريحتان' : n <= 10 ? n + ' شرائح' : n + ' شريحة';
const esc = s => String(s).replace(/[&<>]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));

async function go(fn) {
  if (busy || !canConvert()) return;
  busy = true; refreshLock(); res.innerHTML = '';
  try {
    const ids = await targetIds();
    if (!ids.length) { say('حدّد شريحة أولًا.', ''); return; }
    await fn(ids);
  } catch (e) { prog(null); say('حدث خطأ: ' + (e.message || e)); }
  finally { busy = false; refreshLock(); }
}
$('darkBtn').addEventListener('click', () => go(doDark));

// ---------- the two steps the user does himself ----------
let supported = false, step1 = false, step2 = false;
const step2ok = () => step2;
const canConvert = () => supported && step1 && step2ok();
function refreshLock() {
  $('darkBtn').disabled = busy || !canConvert();
  $('step1').classList.toggle('done', step1);
  $('step2').classList.toggle('done', step2ok());
  const miss = [!step1 && '١', !step2ok() && '٢'].filter(Boolean);
  $('lockNote').hidden = !supported || !miss.length;
  $('lockNote').textContent = miss.length === 2 ? 'أكمل الخطوتين ١ و٢ أولًا.' : 'أكمل الخطوة ' + (miss[0] || '') + ' أولًا.';
}
// step 1: remembered inside the presentation file
function docSettings() { try { return Office.context.document.settings; } catch (e) { return null; } }
$('s1Done').addEventListener('change', () => {
  step1 = $('s1Done').checked; refreshLock();
  const s = docSettings(); if (s) { s.set('udStep1', step1); s.saveAsync(); }
});
// "أضف فواصل": a red full-width line in the middle of the current slide, selected so it can be dragged into place
$('addSep').addEventListener('click', async () => {
  if (!supported) return;
  try {
    await run(async ctx => {
      const sel = ctx.presentation.getSelectedSlides(); sel.load('items/id'); await ctx.sync();
      if (!sel.items.length) { say('حدّد شريحة أولًا.'); return; }
      const slide = sel.items[0];
      let W = 960, H = 540;
      try { const pkg = await UD.openPackage(await exportSlide(slide.id)); W = pkg.slideW / PT; H = pkg.slideH / PT; } catch (e) {}
      // a thin red bar (not a connector line): it always stays horizontal, even when dragged
      const bw = (W - 80) * 1.2; // 20% longer than the slide width minus the side margins, centered
      const ln = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, { left: (W - bw) / 2, top: H / 2, width: bw, height: 3 });
      ln.name = 'UULA فاصل';
      ln.fill.setSolidColor('#FF0000');
      ln.lineFormat.visible = false;
      ln.load('id'); await ctx.sync();
      try { slide.setSelectedShapes([ln.id]); await ctx.sync(); } catch (e) {}
      say('اسحب الفاصل إلى مكانه.');
    });
  } catch (e) { say('تعذّرت إضافة الفاصل: ' + (e.message || e)); }
});
// step 2: the user changes the slide size himself and ticks the box
function setStep2(v) { step2 = v; $('s2Done').checked = v; const s = docSettings(); if (s) { s.set('udStep2', v); s.saveAsync(); } refreshLock(); }
$('s2Done').addEventListener('change', () => setStep2($('s2Done').checked));

// ---------- scope ----------
function setScope(v) { scope = v; save('ud-scope2', v); $('scope').querySelectorAll('button').forEach(b => b.setAttribute('aria-pressed', String(b.dataset.v === v))); }
$('scope').addEventListener('click', e => { const b = e.target.closest('button'); if (b) setScope(b.dataset.v); });
setScope(scope);

// ---------- settings ----------
$('setTog').addEventListener('click', () => { const o = $('setTog').getAttribute('aria-expanded') !== 'true'; $('setTog').setAttribute('aria-expanded', String(o)); $('setBody').hidden = !o; });
$('imgChk').checked = st.images;
$('imgChk').addEventListener('change', () => { st.images = $('imgChk').checked; save('ud-settings', st); });
$('resetBtn').addEventListener('click', () => { Object.assign(st, DEF); save('ud-settings', st); $('imgChk').checked = st.images; paintSwatches(); });
function paintSwatches() { document.querySelectorAll('.colorbtn').forEach(b => { b.querySelector('.sw').style.background = st[b.dataset.k]; }); }
paintSwatches();

// ---------- color picker (standard for all tools) ----------
const COLORS = [['أبيض', '#ffffff'], ['أسود', '#000000'], ['أزرق', '#1f71de'], ['أرجواني', '#fe00fe'], ['أصفر', '#ffe600']];
const THEME = ['#ffffff', '#000000', '#e7e6e6', '#44546a', '#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47'];
const STANDARD = ['#c00000', '#ff0000', '#ffc000', '#ffff00', '#92d050', '#00b050', '#00b0f0', '#0070c0', '#002060', '#7030a0'];
function hex2hsl(h) {
  let [r, g, b] = [1, 3, 5].map(i => parseInt(h.substr(i, 2), 16) / 255);
  const mx = Math.max(r, g, b), mn = Math.min(r, g, b), L = (mx + mn) / 2; let H = 0, S = 0;
  if (mx !== mn) { const d = mx - mn; S = L > .5 ? d / (2 - mx - mn) : d / (mx + mn);
    H = mx === r ? (g - b) / d + (g < b ? 6 : 0) : mx === g ? (b - r) / d + 2 : (r - g) / d + 4; H /= 6; }
  return [H, S, L];
}
function hsl2hex(H, S, L) {
  let r, g, b;
  if (!S) r = g = b = L; else {
    const f = (p, q, t) => { if (t < 0) t += 1; if (t > 1) t -= 1; return t < 1 / 6 ? p + (q - p) * 6 * t : t < 1 / 2 ? q : t < 2 / 3 ? p + (q - p) * (2 / 3 - t) * 6 : p; };
    const q = L < .5 ? L * (1 + S) : L + S - L * S, p = 2 * L - q; r = f(p, q, H + 1 / 3); g = f(p, q, H); b = f(p, q, H - 1 / 3);
  }
  return '#' + [r, g, b].map(v => Math.round(v * 255).toString(16).padStart(2, '0')).join('');
}
function shades(h) {
  const [H, S, L] = hex2hsl(h);
  const mods = L > .95 ? [.95, .85, .75, .65, .5].map(m => [0, m]) : L < .05 ? [.5, .35, .25, .15, .05].map(m => [1, m]) : [[1, .8], [1, .6], [1, .4], [0, .75], [0, .5]];
  return mods.map(([up, m]) => hsl2hex(H, S, up ? L + (1 - L) * m : L * m));
}
const cpop = $('cpop'), custom = $('custom');
let cKey = null, cBtn = null;
function setColorKey(k, c) { st[k] = c.toLowerCase(); save('ud-settings', st); paintSwatches(); }
function swatch(c, title) {
  const b = document.createElement('button'); b.type = 'button'; b.className = 'c'; b.style.background = c; b.title = title || c; b.dataset.c = c;
  b.addEventListener('click', () => { const k = cKey; closeColors(); if (k) setColorKey(k, c); });
  return b;
}
function buildColors() {
  cpop.innerHTML = '';
  const lab = t => { const d = document.createElement('div'); d.className = 'clab'; d.textContent = t; cpop.appendChild(d); };
  lab('ألوان المذكرة');
  const mine = document.createElement('div'); mine.className = 'mine'; for (const [n, c] of COLORS) mine.appendChild(swatch(c, n)); cpop.appendChild(mine);
  lab('ألوان أخرى');
  const top = document.createElement('div'); top.className = 'crow'; THEME.forEach(c => top.appendChild(swatch(c))); cpop.appendChild(top);
  const grid = document.createElement('div'); grid.className = 'ccol'; const sh = THEME.map(shades);
  for (let k = 0; k < 5; k++) THEME.forEach((_, i) => grid.appendChild(swatch(sh[i][k]))); cpop.appendChild(grid);
  const std = document.createElement('div'); std.className = 'crow'; STANDARD.forEach(c => std.appendChild(swatch(c))); cpop.appendChild(std);
  const more = document.createElement('button'); more.type = 'button'; more.className = 'more'; more.textContent = 'لون آخر…';
  more.addEventListener('click', () => { const k = cKey; closeColors(); custom.value = st[k]; custom.dataset.k = k; custom.click(); });
  cpop.appendChild(more);
  cpop.querySelectorAll('button.c').forEach(b => b.classList.toggle('on', b.dataset.c.toLowerCase() === String(st[cKey]).toLowerCase()));
}
function openColors(btn) {
  cKey = btn.dataset.k; cBtn = btn; buildColors(); cpop.hidden = false; btn.setAttribute('aria-expanded', 'true');
  const r = btn.getBoundingClientRect(), pw = cpop.offsetWidth, ph = cpop.offsetHeight;
  let left = Math.max(6, Math.min(r.left, innerWidth - pw - 6));
  let top = r.bottom + 4; if (top + ph > innerHeight - 6) top = Math.max(6, r.top - ph - 4);
  cpop.style.position = 'fixed'; cpop.style.left = left + 'px'; cpop.style.top = (top + scrollY * 0) + 'px';
}
function closeColors() { cpop.hidden = true; if (cBtn) cBtn.setAttribute('aria-expanded', 'false'); cKey = null; cBtn = null; }
document.querySelectorAll('.colorbtn').forEach(b => b.addEventListener('click', e => { e.stopPropagation(); if (cBtn === b) closeColors(); else { closeColors(); openColors(b); } }));
custom.addEventListener('input', () => { if (custom.dataset.k) setColorKey(custom.dataset.k, custom.value); });
document.addEventListener('click', e => { if (!cpop.hidden && !cpop.contains(e.target)) closeColors(); if (!$('appMenu').hidden && !e.target.closest('.appmenu')) menu(false); });
document.addEventListener('keydown', e => { if (e.key === 'Escape') { closeColors(); menu(false); $('modal').hidden = true; } });

// ---------- menu / about ----------
function menu(o) { $('appMenu').hidden = !o; $('menuBtn').setAttribute('aria-expanded', String(o)); }
$('menuBtn').addEventListener('click', e => { e.stopPropagation(); menu($('appMenu').hidden); });
function modal(t, html) { $('modalT').textContent = t; $('modalB').innerHTML = html; $('modal').hidden = false; }
$('modalX').addEventListener('click', () => { $('modal').hidden = true; });
$('modal').addEventListener('click', e => { if (e.target === $('modal')) $('modal').hidden = true; });
$('aboutBtn').addEventListener('click', () => {
  menu(false);
  modal('حول الأداة', '<p>تحوّل أوراق العمل إلى شرائح داكنة للعرض على الطلاب: شريحة لكل جزء بين فاصلين.</p><p>قبل التحويل: أضف الفواصل بنفسك، وحوّل مقاس الشرائح إلى 16:9.</p>'
    + '<ul><li>خلفية داكنة ونص أبيض، والأزرق يصير أصفر</li><li>خط UULA Sans Medium بحجم 36 والعناوين 48</li><li>حذف رقم الصفحة ورمز QR والبطاقات الفاتحة</li><li>تصغير الخط تلقائيًا عند الامتلاء</li></ul>'
    + '<p class="foot" dir="ltr">Developed by <b>Zaher Alkablawi</b> · Version ' + VER + '<br>© 2026 UULA. All rights reserved.</p>');
});

// ---------- updates (same method as UULA Arabic Math) ----------
(function () {
  const head = () => fetch('app.html?chk=' + Date.now(), { method: 'HEAD', cache: 'no-store' })
    .then(r => r.ok ? (r.headers.get('etag') || r.headers.get('last-modified')) : null).catch(() => null);
  let base = null, last = 0;
  const note = m => { if (busy) return; say(m); setTimeout(() => { if (status.textContent === m) say(''); }, 5000); };
  async function check(manual) {
    last = Date.now(); const t = await head();
    if (!base) { base = t; if (manual) note(t ? 'لديك آخر إصدار.' : 'تعذّر الاتصال.'); return; }
    if (t && t !== base) { $('updBtn').hidden = false; if (manual) note('يوجد تحديث.'); }
    else if (manual) note(t ? 'لديك آخر إصدار.' : 'تعذّر الاتصال.');
  }
  head().then(t => { base = t; });
  setInterval(() => check(false), 60 * 60 * 1000);
  const soon = () => { if (Date.now() - last > 60 * 60 * 1000) check(false); };
  window.addEventListener('focus', soon); document.addEventListener('visibilitychange', () => { if (!document.hidden) soon(); });
  $('updBtn').addEventListener('click', () => { if (!busy) location.replace('app.html?v=' + Date.now()); });
  $('chkUpd').addEventListener('click', () => { menu(false); check(true); });
})();

// ---------- start ----------
$('appVer').textContent = VER;
refreshLock();
const ready = (window.Office && Office.onReady) ? Office.onReady() : Promise.resolve({});
ready.then(info => {
  if (!window.PowerPoint || (info && info.host && info.host !== Office.HostType.PowerPoint)) { say('هذه الأداة تعمل داخل PowerPoint فقط.'); return; }
  if (!Office.context.requirements.isSetSupported('PowerPointApi', '1.8')) { say('حدّث برنامج PowerPoint لاستخدام الأداة.'); return; }
  supported = true;
  const s = docSettings(); step1 = !!(s && s.get('udStep1')); $('s1Done').checked = step1;
  step2 = !!(s && s.get('udStep2')); $('s2Done').checked = step2;
  refreshLock();
});
})();
