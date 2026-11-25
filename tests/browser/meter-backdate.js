// Uji UI: peringatan tanggal kosong di panel Meteran, peringatan input susulan + pencatatan berikutnya di formulir, penanda anomali di detail.
const { chromium } = require('playwright');
const { install, handlers } = require('./mockapi');
const BASE = process.env.BASE || 'http://localhost:8765';
let fail = 0; const chk = (n, c, d) => { if (!c) fail++; console.log((c ? 'ok   ' : 'FAIL ') + n + (c ? '' : ' -> ' + d)); };
const days = ['2026-10-04', '2026-10-06'];
handlers.meter_stats = () => ({ locations: ['Gedung A'], latest: [], group: 'day', from: '2026-10-01', to: '2026-10-31', since: days[0],
  series: [{ k: '2026-10-06', lwbp: 30, wbp: 40, genset: 0, air: 8 }], count: { total: 2, period: 2 }, cum: { lwbp: 30, wbp: 40, genset: 0, air: 8 }, per: { lwbp: 30, wbp: 40, genset: 0, air: 8 },
  gaps: [{ loc: 'Gedung A', m: 'listrik', n: 1, dates: ['2026-10-05'] }, { loc: 'Gedung A', m: 'air', n: 1, dates: ['2026-10-05'] }] });
// formulir: pencatatan terakhir 4 Okt, pencatatan berikutnya 6 Okt (jadi pengisian 5 Okt = susulan)
handlers.meter_prev = () => ({ locations: ['Gedung A'], count: 1, since: '2026-10-04', last: { date: '2026-10-04', time: '08:00' },
  elec: { lwbp: 1000, wbp: 500, genset: 50, date: '2026-10-04', time: '08:00' }, water: { reading: 100, date: '2026-10-04', time: '08:00' },
  cum: { lwbp: 0, wbp: 0, genset: 0, air: 0 }, has: { e: true, w: true },
  next: { date: '2026-10-06', time: '08:00', elec: { lwbp: 1030, wbp: 540, genset: 50 }, water: { reading: 108 } } });
(async () => {
  const b = await chromium.launch(); const ctx = await b.newContext({ viewport: { width: 1280, height: 900 } });
  await ctx.addInitScript(() => { localStorage.setItem('stc_t', 't'); localStorage.setItem('stc_u', JSON.stringify({ userId: 'U1', name: 'M', username: 'a', role: 'ADMIN' })); });
  const page = await ctx.newPage(); await page.route(/^https?:\/\/(?!localhost)/, r => r.abort()); await install(page);
  const errs = []; page.on('pageerror', e => errs.push(e.message));
  await page.goto(BASE + '/index.html', { waitUntil: 'load' }); await page.waitForTimeout(900);
  await page.evaluate(() => document.querySelector('#nav a[data-p="meter"]').click()); await page.waitForTimeout(1500);
  const gap = await page.evaluate(() => [...document.querySelectorAll('#mres .mwarn')].map(e => e.textContent).join(' | '));
  chk('panel menampilkan "Pencatatan belum tersedia" untuk tanggal kosong', /Pencatatan belum tersedia/.test(gap) && /Gedung A/.test(gap) && /Listrik/.test(gap) && /Air/.test(gap), gap);
  const tbl = await page.evaluate(() => document.querySelector('.meter-data-table')?.textContent || '');
  chk('tabel harian tidak memuat tanggal fiktif (hanya 6 Okt)', /6 Okt/.test(tbl) && !/5 Okt/.test(tbl), tbl.slice(0, 200));
  await page.evaluate(() => document.getElementById('ra').click()); await page.waitForTimeout(800);
  await page.fill('#rfm [name=location]', 'Gedung A'); await page.fill('#rfm [name=date]', '2026-10-05'); await page.fill('#rfm [name=time]', '08:00');
  await page.dispatchEvent('#rfm [name=date]', 'change'); await page.waitForTimeout(900);
  await page.fill('#rfm [name=lwbp]', '1010'); await page.waitForTimeout(300);
  let live = await page.evaluate(() => document.getElementById('mlive').textContent);
  chk('formulir 5 Okt: info pencatatan berikutnya (6 Okt) akan dihitung ulang', /pencatatan berikutnya/.test(live) && /dihitung ulang/.test(live), live);
  chk('formulir 5 Okt: tidak ada peringatan tanggal terlewat (4 -> 5 berurutan)', !/belum ada pencatatan\. Pemakaian/.test(live), live);
  await page.fill('#rfm [name=lwbp]', '1040'); await page.waitForTimeout(300);
  live = await page.evaluate(() => document.getElementById('mlive').textContent);
  chk('angka lebih besar dari pencatatan berikutnya diperingatkan', /lebih besar dari pencatatan berikutnya/.test(live), live);
  await page.fill('#rfm [name=date]', '2026-10-03'); await page.dispatchEvent('#rfm [name=date]', 'change'); await page.waitForTimeout(500);
  handlers.meter_prev = () => ({ locations: ['Gedung A'], count: 1, since: '2026-10-04', last: { date: '2026-10-04', time: '08:00' },
    elec: { lwbp: 1000, wbp: 500, genset: 50 }, water: { reading: 100 }, cum: { lwbp: 0, wbp: 0, genset: 0, air: 0 }, has: { e: true, w: true }, next: null });
  await page.fill('#rfm [name=date]', '2026-10-08'); await page.dispatchEvent('#rfm [name=date]', 'change'); await page.waitForTimeout(900);
  live = await page.evaluate(() => document.getElementById('mlive').textContent);
  chk('formulir 8 Okt (4 Okt terakhir): peringatan 3 tanggal belum ada pencatatan', /3 tanggal antara/.test(live) && /data terlambat/.test(live), live);
  await page.setViewportSize({ width: 390, height: 844 }); await page.waitForTimeout(400);
  const ov = await page.evaluate(() => document.documentElement.scrollWidth - document.documentElement.clientWidth);
  chk('ponsel 390px: tidak ada overflow horizontal', ov <= 0, ov);
  chk('tidak ada error JS', errs.length === 0, errs.join(';'));
  await b.close(); console.log(fail ? 'HASIL: ' + fail + ' GAGAL' : 'HASIL: SEMUA LULUS'); process.exit(fail ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
