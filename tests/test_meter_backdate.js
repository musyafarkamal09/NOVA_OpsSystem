/* Tes Pencatatan Meteran: input terlambat (backdated), rekalkulasi berantai, edit/hapus, anomali, tanggal kosong.
   Jalankan: node tests/test_meter_backdate.js  (butuh backend/Code.gs = berkas backend/Code.gs) */
'use strict';
const assert = require('assert');
const { makeEnv } = require('./gas-harness');
let pass = 0, fail = 0;
function t(name, fn) { try { fn(); pass++; console.log('  ok  ' + name); } catch (e) { fail++; console.log('  GAGAL ' + name + '\n      ' + (e && e.message)); } }
const ok = (r) => { assert.ok(r.ok, 'harus ok, dapat: ' + r.message); return r.data; };

const env = makeEnv();
env.addUser('adm', 'ADMIN', 'Admin'); env.addUser('tek', 'USER', 'Teknisi');
const tk = { adm: env.login('adm'), tek: env.login('tek') };
const call = (u, a, p) => env.call(a, Object.assign({ token: tk[u] }, p || {}));
const G = (loc, date, l, w, g, air, time, kind) => ({ kind: kind || 'Gabungan', location: loc, date, time: time || '08:00', lwbp: l, wbp: w, genset: g, reading: air });
const save = (d, id) => ok(call('tek', 'rep_save', { type: 'meter', id, data: d })).id;
const get = (id) => ok(call('adm', 'rep_get', { id })).data;
const stats = (loc) => ok(call('adm', 'meter_stats', { location: loc, from: '', to: '', group: 'day' }));
const A = 'Gedung A', B = 'Gedung B';

console.log('== 1. tanggal 4, lalu 6 (5 terlewat) ==');
const id4 = save(G(A, '2026-10-04', 1000, 500, 50, 100));
const id6 = save(G(A, '2026-10-06', 1030, 540, 50, 108));
t('tgl 6 dihitung dari tgl 4 sementara 5 kosong', () => { const d = get(id6); assert.deepEqual([d.diffs.lwbp, d.diffs.wbp, d.diffs.genset, d.diffs.listrik, d.diff], [30, 40, 0, 70, 8]); });
t('tidak ada angka fiktif: tgl 5 tidak muncul di seri harian', () => { const s = stats(A); assert.ok(!s.series.some((x) => x.k === '2026-10-05')); assert.deepEqual(s.series.map((x) => x.k), ['2026-10-06']); });
t('tanggal 5 dilaporkan kosong (listrik dan air)', () => { const g = stats(A).gaps; ['listrik', 'air'].forEach((m) => assert.ok(g.some((x) => x.m === m && x.dates.indexOf('2026-10-05') >= 0), m)); });

console.log('== 2. tgl 5 diisi terlambat ==');
const id5 = save(G(A, '2026-10-05', 1010, 520, 50, 103));
t('tgl 5 dihitung dari tgl 4', () => { const d = get(id5); assert.deepEqual([d.prevs.lwbp, d.diffs.lwbp, d.diffs.wbp, d.diffs.listrik, d.prev, d.diff], [1000, 10, 20, 30, 100, 3]); });
t('tgl 6 otomatis dihitung ulang dari tgl 5', () => { const d = get(id6); assert.deepEqual([d.prevs.lwbp, d.diffs.lwbp, d.diffs.wbp, d.diffs.listrik, d.prev, d.diff], [1010, 20, 20, 40, 103, 5]); });
t('tgl 4 (awal) tetap tanpa selisih', () => { assert.equal(get(id4).diffs.lwbp, null); });
t('seri harian: 5 = 30, 6 = 40; total = selisih meter awal-akhir', () => {
  const s = stats(A); const m = {}; s.series.forEach((x) => (m[x.k] = x));
  assert.equal(m['2026-10-05'].lwbp + m['2026-10-05'].wbp, 30); assert.equal(m['2026-10-06'].lwbp + m['2026-10-06'].wbp, 40);
  assert.equal(s.cum.lwbp + s.cum.wbp, (1030 + 540) - (1000 + 500)); assert.equal(s.cum.air, 8); assert.equal(s.cum.genset, 0);
});
t('gap tgl 5 hilang', () => { assert.ok(!stats(A).gaps.some((x) => x.dates.indexOf('2026-10-05') >= 0)); });
t('selisih tersimpan di REPORTS dan sheet METER_LISTRIK_AIR ikut diperbarui (bukan hanya dihitung saat dibaca)', () => {
  const r = env.rows('REPORTS').find((x) => x.id === id6); assert.equal(JSON.parse(r.data).diffs.lwbp, 20);
  const m = env.rows('METER_LISTRIK_AIR').find((x) => x['ID Laporan'] === id6); assert.equal(Number(m['Selisih LWBP']), 20); assert.equal(Number(m['Selisih Air (m³)']), 5);
});

console.log('== 3. ulangi: tgl 8 dulu, 7 terlambat ==');
const id8 = save(G(A, '2026-10-08', 1100, 600, 60, 120));
t('tgl 8 sementara dari tgl 6', () => { const d = get(id8); assert.deepEqual([d.diffs.lwbp, d.diffs.wbp, d.diffs.genset, d.diff], [70, 60, 10, 12]); });
const id7 = save(G(A, '2026-10-07', 1050, 560, 55, 112));
t('tgl 7 dari tgl 6', () => { const d = get(id7); assert.deepEqual([d.diffs.lwbp, d.diffs.wbp, d.diffs.genset, d.diff], [20, 20, 5, 4]); });
t('tgl 8 dihitung ulang dari tgl 7', () => { const d = get(id8); assert.deepEqual([d.prevs.lwbp, d.diffs.lwbp, d.diffs.wbp, d.diffs.genset, d.diff], [1050, 50, 40, 5, 8]); });
t('genset terpisah dari total listrik', () => { const d = get(id8); assert.equal(d.diffs.listrik, d.diffs.lwbp + d.diffs.wbp); });

console.log('== 4. listrik, air, dan laporan lama terpisah ==');
const L1 = save(G(B, '2026-10-04', 500, 100, 0, null, '08:00', 'Listrik'));
const W1 = save({ kind: 'Air', location: B, date: '2026-10-04', time: '09:00', reading: 50 });
const W3 = save({ kind: 'Air', location: B, date: '2026-10-06', time: '09:00', reading: 60 });
const L3 = save(G(B, '2026-10-06', 540, 120, 0, null, '08:00', 'Listrik'));
const W2 = save({ kind: 'Air', location: B, date: '2026-10-05', time: '09:00', reading: 54 });
t('air terlambat hanya memengaruhi air', () => { assert.equal(get(W2).diff, 4); assert.equal(get(W3).diff, 6); assert.equal(get(W3).prev, 54); assert.equal(get(L3).diffs.lwbp, 40); });
const L2 = save(G(B, '2026-10-05', 520, 110, 0, null, '08:00', 'Listrik'));
t('listrik terlambat hanya memengaruhi listrik', () => { assert.equal(get(L2).diffs.lwbp, 20); assert.equal(get(L3).diffs.lwbp, 20); assert.equal(get(L3).diffs.wbp, 10); assert.equal(get(W3).diff, 6); });
t('lokasi tidak tercampur (Gedung A tidak berubah)', () => { assert.equal(get(id6).diffs.lwbp, 20); });

console.log('== 5. edit & hapus ==');
t('edit tgl 5 memperbarui tgl 5 dan tgl 6', () => {
  save(G(A, '2026-10-05', 1015, 520, 50, 103), id5);
  assert.equal(get(id5).diffs.lwbp, 15); assert.equal(get(id6).diffs.lwbp, 15); assert.equal(get(id6).prevs.lwbp, 1015);
});
t('edit memindahkan tanggal: tgl 5 -> tgl 9 mengubah urutan', () => {
  save(G(A, '2026-10-09', 1015, 520, 50, 103), id5);
  assert.equal(get(id6).prevs.lwbp, 1000); assert.equal(get(id6).diffs.lwbp, 30);   // 6 kembali dari 4
  assert.equal(get(id5).prevs.lwbp, 1100); assert.equal(get(id5).diffs.lwbp, null); // angka lebih kecil dari tgl 8 = anomali
  save(G(A, '2026-10-05', 1010, 520, 50, 103), id5);                                 // kembalikan
  assert.equal(get(id6).prevs.lwbp, 1010); assert.equal(get(id5).diffs.lwbp, 10);
});
t('hapus (arsip) tgl 5 menghitung ulang tgl 6 dari tgl 4; data tidak hilang', () => {
  ok(call('adm', 'rep_del', { id: id5, reason: 'uji' }));
  assert.equal(get(id6).prevs.lwbp, 1000); assert.equal(get(id6).diffs.lwbp, 30); assert.equal(get(id6).diff, 8);
  assert.ok(env.rows('REPORTS').some((x) => x.id === id5 && x.status === 'Diarsipkan'));
  assert.ok(stats(A).gaps.some((x) => x.dates.indexOf('2026-10-05') >= 0));
  const m = env.rows('METER_LISTRIK_AIR').find((x) => x['ID Laporan'] === id6); assert.equal(Number(m['Selisih LWBP']), 30);
});

console.log('== 6. anomali: reset / ganti meter ==');
const idR = save(G(A, '2026-10-10', 10, 5, 0, 130));
t('tidak ada pemakaian negatif; reset tidak diakumulasi', () => {
  const d = get(idR); assert.equal(d.diffs.lwbp, null); assert.equal(d.diffs.wbp, null); assert.equal(d.diffs.listrik, null); assert.equal(d.diff, 10);
  const s = stats(A); assert.equal(s.cum.lwbp, 100);   // 1100 - 1000: angka reset tidak ikut diakumulasi
  s.series.forEach((x) => ['lwbp', 'wbp', 'genset', 'air'].forEach((m) => assert.ok(x[m] >= 0)));
});

console.log('== 7. waktu pencatatan, bukan waktu input ==');
t('dua pencatatan di tanggal sama: urut menurut jam meter', () => {
  const late = save(G(B, '2026-10-07', 560, 130, 0, null, '20:00', 'Listrik'));
  const early = save(G(B, '2026-10-07', 550, 125, 0, null, '08:00', 'Listrik'));
  assert.equal(get(early).prevs.lwbp, 540); assert.equal(get(early).diffs.lwbp, 10); assert.equal(get(late).prevs.lwbp, 550); assert.equal(get(late).diffs.lwbp, 10);
});
t('jam identik: urutan stabil (createdAt), tanpa saling mengabaikan', () => {
  const a = save(G(B, '2026-10-08', 570, 135, 0, null, '08:00', 'Listrik')), b = save(G(B, '2026-10-08', 575, 136, 0, null, '08:00', 'Listrik'));
  assert.equal(get(b).prevs.lwbp, 570); assert.equal(get(a).diffs.lwbp, 10);
});

console.log('== 8. ekspor/daftar memakai hasil rekalkulasi ==');
t('rep_export meter: selisih tgl 6 sama dengan rep_get', () => {
  const ex = ok(call('adm', 'rep_export', { type: 'meter', q: '', from: '', to: '', status: '' })); const it = ex.items.find((x) => x.id === id6);
  assert.equal(it.data.diffs.lwbp, 30);
});
t('meter_prev menyertakan pencatatan berikutnya untuk peringatan input susulan', () => {
  const r = ok(call('adm', 'meter_prev', { location: A, until: '2026-10-05 08:00' })); assert.equal(r.last.date, '2026-10-04'); assert.equal(r.next.date, '2026-10-06');
});

console.log('\nHASIL: ' + pass + ' lulus, ' + fail + ' gagal');
process.exit(fail ? 1 : 0);
