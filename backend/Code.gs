/** STC XXI TECHNICAL REPORT - Backend (Google Apps Script). Tempel di proyek Apps Script yang dibuat dari Spreadsheet.
 *  Versi 10 (revisi): perhitungan meteran kronologis (selisih dicatat pada tanggal pencatatan aktual, tidak dibagi rata; tanggal kosong = gaps), masa simpan laporan 45 hari sejak diterbitkan,
 *  laporan meteran dikecualikan dari penghapusan otomatis, rep_export dibatasi untuk Guide/Security, daftar & ekspor meteran membaca data sekali.
 *  Versi 10 (patch): flag AC Eror (err) per area disimpan di server; notifikasi push saat ada AC error.
 *  Versi 9: SCHEMA dinaikkan ke 11 agar kolom USERS.photoFileId pasti ada (penyebab foto profil gagal tersimpan), foto profil menerima data base64 dengan/ tanpa awalan data:.
 *  Versi 8: STANDAR SUHU AC (default 23-23,5 C, hanya Admin yang dapat mengubah; status APPROVED / TOO-WARM / TOO-COLD dihitung di server,
 *  disimpan terpisah dari status lain), FOTO PROFIL dengan penggantian aman (foto lama dihapus hanya setelah foto baru + database berhasil).
 *  Versi 7: METERAN LISTRIK + AIR DALAM 1 LAPORAN (kind 'Gabungan'), akumulasi pemakaian, statistik untuk grafik periode
 *  (handler meter_prev & meter_stats). Laporan meteran lama (Listrik / Air terpisah) tetap terbaca dan ikut dihitung.
 *  Versi 6: foto sampai 50 per laporan (unggah bertahap), folder Drive per laporan, ZIP otomatis di atas 10 foto,
 *  Daily Monitoring AC multi-area, ekspor Excel, Report ID TR-YYYYMMDD-NNNN, validasi sisi server, migrasi skema otomatis,
 *  ARSIP BULANAN semua kategori, SHEET FOTO per kategori (setiap foto satu baris + pemeriksaan kelengkapan),
 *  role PELAPOR dan alur PERLU DIPERBAIKI (dilaporkan siapa saja, ditanggapi & diselesaikan Teknisi). */
var TZ='Asia/Jakarta', PAGE=15, MAXPH=50, PHBATCH=6, ZIP_AT=10, MAXB=4000000, EDIT_MIN=30, PUBLIC_PHOTOS=true, TOKEN_TTL=14*24*3600*1000, IDLE_MS=30*60*1000, MAX_FAIL=10, LOCK_MS=15*60*1000, SCHEMA='18', MAXEXP=3000;
var SH={
 USERS:['userId','fullName','username','role','status','salt','hash','mustChange','createdAt','updatedAt','photoFileId','failCount','lockedAt','outlet'],
 ANNOUNCEMENTS:['id','title','body','priority','publishedAt','status','createdBy','updatedAt'],
 ATS_FILMS:['id','title','genre','duration','rating','atsAt','showAt','status','posterFileId','note','updatedAt','deleteAt'],
 REPORTS:['id','type','title','data','userId','reporter','createdAt','updatedAt','editDeadline','status','reqId','foto','folderId','zipId'],
 PHOTOS:['photoId','reportId','fileId','kind','uploadedAt','capturedAt','url','preview'],
 AUDIT_LOG:['logId','userId','userName','role','action','category','recordId','at','summary'],
 FILM_REPORTS:['id','title','status','format','kdm','reportDate','outlet','capacity','createdBy','createdAt','updatedAt','signature'],
 GLOSSARY:['id','term','definition','category','source','updatedBy','updatedAt'],
 STORAGE:['id','title','category','description','fileName','ext','mime','size','fileId','version','createdBy','createdAt','updatedBy','updatedAt'],
 /* Data Penonton per-show (v20.1.0): 1 baris = 1 show; data aktif unik per showKey = Tanggal|Studio|Show */
 AUDIENCE_SHOWS:['showKey','date','studio','show','film','jam','tiket','penonton','petugas','ket','catatan','userId','reporter','createdAt','updatedAt','editDeadline','status','source','sourceId','sourceKey','sourceRef','review']
};
var PFX={work:'WORK',meter:'METER',purchase:'PURCHASE',ac:'AC',issue:'ISSUE',patrol:'IPS',audience:'AUDIENCE'};
var REPORT_PREFIX={work:'LP',meter:'LM',purchase:'LB',ac:'AC',issue:'LPR',patrol:'IPS',audience:'DP'};
var FILM_REPORT_PREFIX='LF';
var FOLDERS={work:'WORK REPORTS',meter:'METER REPORTS',purchase:'PURCHASE REPORTS',ac:'AC MONITORING',issue:'ISSUE REPORTS',patrol:'INSPEKSI PATROLI STUDIO',audience:'DATA PENONTON',poster:'ATS POSTERS',avatar:'PROFILE PHOTOS',store:'DATA STORAGE'};
var WCAT=['Electrical','Mechanical','HVAC / AC','Plumbing','Audio','Projector','Building Maintenance','Preventive Maintenance','Corrective Maintenance','Lainnya'];
var WST=['Open','In Progress','Completed','Pending','Need Follow-up'];
var FST=['Upcoming ATS','ATS Open','Coming Soon','Coming to Theater','Now Showing','Completed'];
var ACUNIT=['Daikin','York'], ACMAX=20;
/* kondisi area AC selain suhu terisi: EMPTY = studio kosong, NOSHOW = tidak ada show. Suhu tidak perlu diisi dan area tetap berstatus APPROVED. */
var ACCOND={EMPTY:'Studio kosong',NOSHOW:'Tidak ada show'};
function acCond_(a){var c=a&&a.cond;return (c&&ACCOND[c])?c:'';}
/* standar suhu AC: default 23 - 23,5 C; nilai aktif disimpan di Script Properties (AC_TMIN / AC_TMAX), hanya Admin yang dapat mengubah */
var ACMIN0=23, ACMAX0=23.5, AC_C_=null;
function acCfg_(){
 if(AC_C_)return AC_C_;
 var p=P_(),mn=parseFloat(p.getProperty('AC_TMIN')),mx=parseFloat(p.getProperty('AC_TMAX'));
 if(!isFinite(mn)||!isFinite(mx)||mn>mx){mn=ACMIN0;mx=ACMAX0;}
 AC_C_={min:mn,max:mx};return AC_C_;
}
/* Kode status tersimpan tetap APPROVED / TOO-WARM / TOO-COLD (data lama tetap valid).
   Label yang ditampilkan: APPROVED / PANAS (di atas batas atas) / DINGIN (di bawah batas bawah). */
function acLabel_(s){return s==='TOO-WARM'?'PANAS':(s==='TOO-COLD'?'DINGIN':(s||''));}
function acEval_(v){var c=acCfg_(),t=Math.round(Number(v)*100)/100;return t>c.max?'TOO-WARM':(t<c.min?'TOO-COLD':'APPROVED');}
/* tampilan: hitung ulang status sesuai standar terbaru (hanya untuk output; data tersimpan tidak diubah) */
function acApply_(d){
 var c=acCfg_();
 d.rows=acRows_(d).map(function(a){var t=a.roomTemp,ok=t!==''&&t!=null&&isFinite(Number(t));a.tempStatus=acCond_(a)?'APPROVED':(ok?acEval_(t):'');return a;});
 d.tempStd={min:c.min,max:c.max};return d;
}
var ISP=['Rendah','Sedang','Tinggi','Darurat'];
var IST=['Menunggu','Ditanggapi','Sedang Dikerjakan','Selesai'];
var ROLES=['ADMIN','USER','REPORTER','GUIDE_SECURITY','MANAGER']; /* ADMIN=Admin, USER=Teknisi, REPORTER=Pelapor (hanya melaporkan "Perlu Diperbaiki") */
var PHS={work:'FOTO_PEKERJAAN',meter:'FOTO_METERAN',purchase:'FOTO_PEMBELIAN',ac:'FOTO_AC',issue:'FOTO_PERLU_DIPERBAIKI',patrol:'FOTO_INSPEKSI_PATROLI',audience:'FOTO_DATA_PENONTON'};
var PHEAD=['ID Laporan','No','Kelompok / Area','Foto','Link Foto','Waktu Unggah','ID Foto'];
var ARCHL={work:'Laporan Pekerjaan',meter_listrik:'Meteran Listrik',meter_air:'Meteran Air',meter_gabungan:'Meteran Listrik dan Air',purchase:'Laporan Pembelian',ac:'Monitoring AC',issue:'Perlu Diperbaiki',audience:'Data Penonton'};
/* laporan AC baru berisi banyak area (d.rows); laporan lama hanya satu area (location/unit/roomTemp) */
function acRows_(d){if(d&&Array.isArray(d.rows)&&d.rows.length)return d.rows;return (d&&d.location)?[{location:d.location,unit:d.unit,roomTemp:d.roomTemp,note:''}]:[];}
/* flag AC Eror per area: true / 'true' / 1 / 'Ya' dianggap error */
function acErrIn_(r){var v=r&&r.err;return v===true||v==='true'||v===1||v==='1'||v==='Ya';}

/* ===== SHEET PER KATEGORI =====
   REPORTS tetap menjadi database utama yang dibaca website. Setiap laporan juga ditulis otomatis
   ke sheet sesuai kategorinya (kolom terpisah + kolom foto/folder/ZIP). */
/* Menulis blok baris ke sheet. Sel berisi =HYPERLINK("url","label") ditulis sebagai TEKS BERTAUTAN (rich text),
   bukan rumus, sehingga tidak bergantung pada lokal spreadsheet (lokal Indonesia memakai ';' sebagai pemisah
   argumen rumus, sehingga rumus berkoma menjadi #ERROR!). Rumus lain (mis. =IMAGE) tetap ditulis seperti biasa. */
function putRows_(s,row,col,m){
 if(!m||!m.length)return;
 var re=/^=HYPERLINK\("(https:\/\/drive\.google\.com\/[^"]+)","([^"]+)"\)$/,nr=m.length,nc=m[0].length,links={};
 var vals=m.map(function(r,i){return r.map(function(v,j){var x=typeof v==='string'&&re.exec(v);if(x){(links[j]=links[j]||{})[i]=x;return '';}return v;});});
 s.getRange(row,col,nr,nc).setValues(vals);
 Object.keys(links).forEach(function(j){
  var rv=m.map(function(r,i){var x=links[j][i],b=SpreadsheetApp.newRichTextValue();
   return [x?b.setText(x[2]).setLinkUrl(x[1]).build():b.setText(r[j]==null?'':String(r[j])).build()];});
  s.getRange(row,col+Number(j),nr,1).setRichTextValues(rv);
 });
}
/* JALANKAN sekali dari editor setelah kode ini dipasang: memperbaiki sel "#ERROR!" pada kolom Link Foto
   di semua sheet FOTO_*, serta kolom Folder/ZIP di sheet kategori. Tidak menghapus atau membangun ulang sheet. */
function perbaikiLinkFoto(){
 var t0=Date.now(),ph={},cnt=0,sh=0;
 all_('PHOTOS').forEach(function(p){if(p.fileId)ph[String(p.photoId)]=p.fileId;});
 Object.keys(PHS).forEach(function(t){
  var s=ss_().getSheetByName(PHS[t]);if(!s)return;var n=s.getLastRow();if(n<2)return;
  var ids=s.getRange(2,7,n-1,1).getValues(),
   rv=ids.map(function(r){var f=ph[String(r[0])],b=SpreadsheetApp.newRichTextValue();
    return [f?b.setText('Buka foto').setLinkUrl('https://drive.google.com/file/d/'+f+'/view').build():b.setText('').build()];});
  s.getRange(2,5,n-1,1).setRichTextValues(rv);cnt+=n-1;
 });
 all_('REPORTS').forEach(function(r){
  if(Date.now()-t0>300000||r.status==='Diarsipkan'||!PFX[r.type]||!r.folderId)return;
  try{syncFoto_(r.id);sh++;}catch(e){Logger.log('Gagal '+r.id+': '+e);}
 });
 Logger.log('Link foto diperbaiki: '+cnt+' baris di sheet FOTO_*; kolom folder/ZIP disinkronkan untuk '+sh+' laporan.'+((Date.now()-t0>300000)?' BATAS WAKTU TERCAPAI: jalankan sekali lagi.':''));
}

function fmtTs_(ms){var n=Number(ms);return n?Utilities.formatDate(new Date(n),TZ,'dd/MM/yyyy HH:mm'):'';}
function nz_(v){return v==null?'':v;}
function legacy_(d){var a=[];if(d.meterNo)a.push('No. meteran lama: '+d.meterNo);if(d.kind!=='Air'&&d.lwbp==null&&d.reading!=null)a.push('Pembacaan lama: '+d.reading+' kWh');if(d.note)a.push(d.note);return a.join(' | ');}
function mirKey_(t,d){return t==='meter'?(d&&d.kind==='Air'?'meter_air':(d&&d.kind==='Gabungan'?'meter_gabungan':'meter_listrik')):t;}
var MIR={
 work:{name:'WORK_REPORTS',num:[],wrap:[10],
  head:['ID Laporan','Dibuat','Tanggal','Jam','Judul Pekerjaan','Kategori','Lokasi','Status Pekerjaan','Tim','Deskripsi','Catatan','Pelapor','Terakhir Diubah','Status Data'],
  row:function(r,d){return [r.id,fmtTs_(r.createdAt),d.date,d.time,d.title,d.category,d.location,d.status,(d.team||[]).join(', '),d.desc,d.note,r.reporter,fmtTs_(r.updatedAt),r.status];}},
 meter_listrik:{name:'METER_LISTRIK',num:[6,7,8,9,10,11],wrap:[],
  head:['ID Laporan','Dibuat','Tanggal','Jam','Lokasi','LWBP (kWh)','WBP (kWh)','Genset (kWh)','Selisih LWBP','Selisih WBP','Selisih Genset','Catatan','Pelapor','Terakhir Diubah','Status Data'],
  row:function(r,d){var f=d.diffs||{};return [r.id,fmtTs_(r.createdAt),d.date,d.time,d.location,nz_(d.lwbp),nz_(d.wbp),nz_(d.genset),nz_(f.lwbp),nz_(f.wbp),nz_(f.genset),legacy_(d),r.reporter,fmtTs_(r.updatedAt),r.status];}},
 meter_air:{name:'METER_AIR',num:[6,7,8],wrap:[],
  head:['ID Laporan','Dibuat','Tanggal','Jam','Lokasi','Angka Meteran (m³)','Angka Sebelumnya','Selisih (m³)','Catatan','Pelapor','Terakhir Diubah','Status Data'],
  row:function(r,d){return [r.id,fmtTs_(r.createdAt),d.date,d.time,d.location,nz_(d.reading),nz_(d.prev),nz_(d.diff),legacy_(d),r.reporter,fmtTs_(r.updatedAt),r.status];}},
 meter_gabungan:{name:'METER_LISTRIK_AIR',num:[6,7,8,9,10,11,12,13,14],wrap:[],
  head:['ID Laporan','Dibuat','Tanggal','Jam','Lokasi','LWBP (kWh)','WBP (kWh)','Genset (kWh)','Selisih LWBP','Selisih WBP','Selisih Genset','Meteran Air (m³)','Air Sebelumnya','Selisih Air (m³)','Catatan','Pelapor','Terakhir Diubah','Status Data'],
  row:function(r,d){var f=d.diffs||{};return [r.id,fmtTs_(r.createdAt),d.date,d.time,d.location,nz_(d.lwbp),nz_(d.wbp),nz_(d.genset),nz_(f.lwbp),nz_(f.wbp),nz_(f.genset),nz_(d.reading),nz_(d.prev),nz_(d.diff),legacy_(d),r.reporter,fmtTs_(r.updatedAt),r.status];}},
 purchase:{name:'PURCHASE_REPORTS',num:[10,11],wrap:[9],
  head:['ID Laporan','Dibuat','Tanggal','Jam','Toko','Pembeli','Penerima','Keperluan','Rincian Barang','Jumlah Item','Total (Rp)','Catatan','Pelapor','Terakhir Diubah','Status Data'],
  row:function(r,d){var it=d.items||[];return [r.id,fmtTs_(r.createdAt),d.date,d.time,d.store,d.buyer,d.receiver,d.purpose,
   it.map(function(i,n){return (n+1)+'. '+i.name+(i.category?' ['+i.category+']':'')+' - '+i.qty+' '+(i.unit||'')+' x '+i.price+' = '+i.total;}).join('\n'),
   it.length,d.total,d.note,r.reporter,fmtTs_(r.updatedAt),r.status];}},
 ac:{name:'AC_MONITORING',num:[5],wrap:[6],
  head:['ID Laporan','Dibuat','Tanggal','Jam','Jumlah Area','Rincian (Area | Merk | Suhu | Catatan)','Catatan Umum','Pelapor','Terakhir Diubah','Status Data'],
  row:function(r,d){var rs=acRows_(d);return [r.id,fmtTs_(r.createdAt),d.date,d.time,rs.length,
   rs.map(function(a,i){return (i+1)+'. '+a.location+' | '+a.unit+' | '+(acCond_(a)?ACCOND[a.cond]:nz_(a.roomTemp)+' °C')+(a.tempStatus?' ['+acLabel_(a.tempStatus)+']':'')+(acErrIn_(a)?' | AC ERROR':'')+(a.note?' | '+a.note:'');}).join('\n'),d.note,r.reporter,fmtTs_(r.updatedAt),r.status];}},
 audience:{name:'DATA_PENONTON',num:[4,5,6],wrap:[7],
  head:['ID Laporan','Dibuat','Tanggal','Jumlah Show','Total Tiket Terjual','Total Penonton Masuk','Rincian (Studio | Show | Film | Jam | Tiket | Penonton | Petugas | Keterangan)','Catatan','Pelapor','Terakhir Diubah','Status Data'],
  row:function(r,d){var det=[];(d.studios||[]).forEach(function(a){(a.shows||[]).forEach(function(w){det.push(a.studio+' | '+w.show+' | '+w.film+' | '+nz_(w.jam)+' | '+nz_(w.tiket)+' | '+nz_(w.penonton)+' | '+nz_(w.petugas)+' | '+nz_(w.ket));});});
   return [r.id,fmtTs_(r.createdAt),d.date,nz_(d.totalShow),nz_(d.totalTiket),nz_(d.totalPenonton),det.join('\n'),d.note,r.reporter,fmtTs_(r.updatedAt),d.draft===true?'Draf':(d.draft===false?'Terkirim':r.status)];}},
 issue:{name:'PERLU_DIPERBAIKI',num:[8],wrap:[6,9],
  head:['ID Laporan','Dibuat','Judul Masalah','Lokasi','Prioritas','Uraian Masalah','Status Perbaikan','Jumlah Tanggapan','Tanggapan Terakhir','Pelapor','Terakhir Diubah','Status Data'],
  row:function(r,d){var rp=d.responses||[],l=rp[rp.length-1];return [r.id,fmtTs_(r.createdAt),d.title,d.location,d.priority,d.desc,d.status,rp.length,
   l?(fmtTs_(l.at)+' | '+l.by+' | '+l.status+' | '+l.text):'',r.reporter,fmtTs_(r.updatedAt),r.status];}}
};
function mo_(ms){return Utilities.formatDate(new Date(Number(ms)),TZ,'yyyy-MM');}
function keyType_(k){return k.indexOf('meter')===0?'meter':k;}
/* bulan yang sudah diarsipkan per kategori: "<kunci>|YYYY-MM" disimpan di Script Properties */
function archDone_(){try{return JSON.parse(P_().getProperty('ARCH_DONE')||'[]');}catch(e){return [];}}
function isArch_(k,m){return archDone_().indexOf(k+'|'+m)>=0;}
function archMark_(k,m){var a=archDone_();if(a.indexOf(k+'|'+m)<0){a.push(k+'|'+m);P_().setProperty('ARCH_DONE',JSON.stringify(a));}}
function phHead_(k){return k==='work'?['Foto Before','Foto After','Jumlah Foto','Folder Foto','ZIP Foto']:['Foto','Jumlah Foto','Folder Foto','ZIP Foto'];}
function fullHead_(k){return MIR[k].head.concat(phHead_(k));}
function zipUrl_(id){return 'https://drive.google.com/uc?export=download&id='+id;}
function folderUrl_(id){return 'https://drive.google.com/drive/folders/'+id;}
/* info foto satu laporan: jumlah, foto pertama/Before/After, folder, ZIP */
function phInfo_(list,rep){
 list=list.slice().sort(function(a,b){return Number(a.uploadedAt)-Number(b.uploadedAt);});
 var f=function(k){return list.filter(function(p){return p.kind===k;})[0];},b=f('Sebelum'),a=f('Sesudah'),fi=list[0];
 return {n:list.length,first:fi&&fi.fileId,before:b&&b.fileId,after:a&&a.fileId,folder:(rep&&rep.folderId)||'',zip:(rep&&rep.zipId)||''};
}
function phCells_(k,i){
 var img=function(id){return '';};
 var fo=i.folder?'=HYPERLINK("'+folderUrl_(i.folder)+'","Buka folder")':'',zp=i.zip?'=HYPERLINK("'+zipUrl_(i.zip)+'","Unduh ZIP")':'';
 return k==='work'?[img(i.before),img(i.after),i.n||'',fo,zp]:[img(i.first),i.n||'',fo,zp];
}
function mirrorSheet_(k){
 var m=MIR[k],ss=ss_(),s=ss.getSheetByName(m.name);if(s)return s;
 s=ss.insertSheet(m.name);var rows=Math.max(1,s.getMaxRows()-1),h=fullHead_(k),n=m.head.length;
 s.getRange(1,1,1,h.length).setValues([h]).setFontWeight('bold').setBackground('#0a3a2a').setFontColor('#ffffff');
 for(var c=1;c<=n;c++){if(m.num.indexOf(c)<0)s.getRange(2,c,rows,1).setNumberFormat('@');}
 m.wrap.forEach(function(c){s.getRange(2,c,rows,1).setWrap(true);s.setColumnWidth(c,320);});
 for(var j=n+1;j<=h.length;j++)s.setColumnWidth(j,j<=n+(k==='work'?2:1)?160:110);
 s.setFrozenRows(1);return s;
}
function mirrorLoc_(k,id){
 var s=ss_().getSheetByName(MIR[k].name);if(!s)return null;
 var n=s.getLastRow();if(n<2)return {s:s,row:0};
 var v=s.getRange(2,1,n-1,1).getValues();
 for(var i=0;i<v.length;i++)if(String(v[i][0])===String(id))return {s:s,row:i+2};
 return {s:s,row:0};
}
function mirror_(r,d){
 var k=mirKey_(r.type,d),m=MIR[k];if(!m)return;
 if(isArch_(k,mo_(r.createdAt)))return;
 var s=mirrorSheet_(k),loc=mirrorLoc_(k,r.id),v=m.row(r,d),row=(loc&&loc.row)?loc.row:s.getLastRow()+1;
 if(row>s.getMaxRows())s.insertRowsAfter(s.getMaxRows(),200);
 s.getRange(row,1,1,v.length).setValues([v]);
 /* bila jenis meteran berubah saat diedit, hapus baris lama di sheet meteran lainnya */
 if(r.type==='meter')['meter_listrik','meter_air','meter_gabungan'].forEach(function(o){if(o===k)return;var l=mirrorLoc_(o,r.id);if(l&&l.row)l.s.deleteRow(l.row);});
}
function mirrorSafe_(r,d){try{mirror_(r,d);}catch(e){Logger.log('Gagal menulis sheet kategori: '+e);}}
/* membangun ulang semua sheet kategori dari REPORTS + PHOTOS (aman dijalankan ulang) */
function rebuildMirrors_(){
 var ss=ss_(),ks=Object.keys(MIR),rows={},by={};
 ks.forEach(function(k){rows[k]=[];});
 var mrecs=meterRecs_();
 all_('PHOTOS').forEach(function(p){if(p.fileId)(by[p.reportId]=by[p.reportId]||[]).push(p);});
 ks.forEach(function(k){var s=ss.getSheetByName(MIR[k].name);if(s)ss.deleteSheet(s);});
 all_('REPORTS').sort(function(a,b){return Number(a.createdAt)-Number(b.createdAt);}).forEach(function(r){
  var d,k;try{d=JSON.parse(r.data);k=mirKey_(r.type,d);}catch(e){Logger.log('Data rusak dilewati: '+r.id);return;}
  if(r.type==='meter')meterDiff_(d,r.id,mrecs);
  if(!MIR[k]||!PFX[r.type]||isArch_(k,mo_(r.createdAt)))return;
  var i=phInfo_(by[r.id]||[],r);rows[k].push({v:rowSafe_(MIR[k].row(r,d)).concat(phCells_(k,i)),has:i.n>0});
 });
 ks.forEach(function(k){
  var s=mirrorSheet_(k),list=rows[k],n=list.length;if(!n)return;
  if(n+1>s.getMaxRows())s.insertRowsAfter(s.getMaxRows(),n+1-s.getMaxRows()+50);
  putRows_(s,2,1,list.map(function(x){return x.v;}));
  list.forEach(function(x,j){if(x.has)s.setRowHeight(j+2,100);});
  Logger.log(MIR[k].name+': '+n+' laporan.');
 });
 rebuildPhotoSheets_();
}
/* JALANKAN dari editor bila ingin memaksa pembuatan ulang sheet kategori */
function pisahkanData(){rebuildMirrors_();Logger.log('Selesai. Sheet kategori dibangun ulang.');}

/* ===== SHEET FOTO PER KATEGORI =====
   Setiap foto dokumentasi = satu baris (gambar + tautan + kelompok/area). Dibuat saat foto diunggah, dan
   diperiksa kelengkapannya otomatis di akhir unggahan (photo_done) serta lewat periksaKelengkapanFoto(). */
function photoSheet_(t){
 var ss=ss_(),n=PHS[t],s=ss.getSheetByName(n);if(s)return s;
 s=ss.insertSheet(n);
 s.getRange(1,1,1,PHEAD.length).setValues([PHEAD]).setFontWeight('bold').setBackground('#0a3a2a').setFontColor('#ffffff');
 var rows=Math.max(1,s.getMaxRows()-1);
 s.getRange(2,1,rows,1).setNumberFormat('@');s.getRange(2,7,rows,1).setNumberFormat('@');
 s.setColumnWidth(1,170);s.setColumnWidth(3,170);s.setColumnWidth(4,160);s.setColumnWidth(5,110);s.setColumnWidth(6,130);s.setColumnWidth(7,160);
 s.setFrozenRows(1);return s;
}
function photoRow_(p,n){
 return [p.reportId,n,safeTxt_(String(p.kind||'Foto')),
  '',
  p.fileId?'=HYPERLINK("https://drive.google.com/file/d/'+p.fileId+'/view","Buka foto")':'',
  fmtTs_(p.uploadedAt),p.photoId];
}
/* menambahkan baris foto yang belum ada untuk satu laporan; mengembalikan jumlah baris yang ditambahkan */
function syncPhotoRows_(rep,photosAll){
 if(!rep||!PHS[rep.type])return 0;
 var k;try{k=mirKey_(rep.type,JSON.parse(rep.data));}catch(e){k=rep.type;}
 if(isArch_(k,mo_(rep.createdAt)))return 0;
 var s=photoSheet_(rep.type),n=s.getLastRow(),have={},
  list=(photosAll||all_('PHOTOS')).filter(function(p){return p.reportId===rep.id&&p.fileId;}).sort(function(a,b){return Number(a.uploadedAt)-Number(b.uploadedAt);});
 if(n>1)s.getRange(2,7,n-1,1).getValues().forEach(function(r){have[String(r[0])]=1;});
 var miss=[];list.forEach(function(p,i){if(!have[String(p.photoId)])miss.push(photoRow_(p,i+1));});
 if(!miss.length)return 0;
 var start=s.getLastRow()+1;
 if(start+miss.length-1>s.getMaxRows())s.insertRowsAfter(s.getMaxRows(),miss.length+50);
 putRows_(s,start,1,miss);
 if(PUBLIC_PHOTOS)s.setRowHeights(start,miss.length,100);
 return miss.length;
}
/* JALANKAN dari editor kapan saja: memastikan SEMUA foto di tabel PHOTOS tampil di sheet foto kategorinya */
function periksaKelengkapanFoto(){
 var ph=all_('PHOTOS'),add=0,rp=0;
 all_('REPORTS').forEach(function(r){if(r.status==='Diarsipkan'||!PHS[r.type])return;rp++;
  try{add+=syncPhotoRows_(r,ph);}catch(e){Logger.log('Gagal memeriksa '+r.id+': '+e);}});
 Logger.log(rp+' laporan diperiksa, '+add+' baris foto yang kurang sudah ditambahkan.');return add;
}
function rebuildPhotoSheets_(){
 var ss=ss_(),by={},rows={};
 all_('PHOTOS').forEach(function(p){if(p.fileId)(by[p.reportId]=by[p.reportId]||[]).push(p);});
 Object.keys(PHS).forEach(function(t){var s=ss.getSheetByName(PHS[t]);if(s)ss.deleteSheet(s);rows[t]=[];});
 all_('REPORTS').sort(function(a,b){return Number(a.createdAt)-Number(b.createdAt);}).forEach(function(r){
  if(!rows[r.type]||r.status==='Diarsipkan')return;
  var k;try{k=mirKey_(r.type,JSON.parse(r.data));}catch(e){return;}
  if(isArch_(k,mo_(r.createdAt)))return;
  (by[r.id]||[]).sort(function(a,b){return Number(a.uploadedAt)-Number(b.uploadedAt);}).forEach(function(p,i){rows[r.type].push(photoRow_(p,i+1));});
 });
 Object.keys(PHS).forEach(function(t){
  var s=photoSheet_(t),n=rows[t].length;if(!n)return;
  if(n+1>s.getMaxRows())s.insertRowsAfter(s.getMaxRows(),n+1-s.getMaxRows()+50);
  putRows_(s,2,1,rows[t]);
  if(PUBLIC_PHOTOS)s.setRowHeights(2,n,100);
  Logger.log(PHS[t]+': '+n+' foto.');
 });
}

/* ===== ARSIP BULANAN (semua kategori) =====
   Awal bulan baru: laporan bulan lalu disalin ke file Google Sheets sendiri (Laporan + Foto + Rekap) di Drive,
   diverifikasi jumlah barisnya, lalu dihapus dari sheet kategori & sheet foto sehingga sheet mulai kosong lagi.
   Database utama REPORTS/PHOTOS TIDAK disentuh, jadi website tetap menyimpan semua riwayat. */
function archFolder_(k){
 var p=P_(),id=p.getProperty('F_arch_'+k);
 if(id)try{return DriveApp.getFolderById(id);}catch(e){}
 var f=folder_(folder_(folder_(null,'STC XXI TECHNICAL REPORT'),'ARSIP BULANAN'),ARCHL[k]);p.setProperty('F_arch_'+k,f.getId());return f;
}
/* hapus baris (kolom A cocok dengan ids) dari sheet, dari bawah ke atas, per blok berurutan */
function delRows_(name,ids){
 var s=ss_().getSheetByName(name);if(!s)return 0;
 var n=s.getLastRow();if(n<2)return 0;
 var v=s.getRange(2,1,n-1,1).getValues(),del=0,i=v.length-1;
 while(i>=0){
  if(ids[String(v[i][0])]){var j=i;while(j-1>=0&&ids[String(v[j-1][0])])j--;
   var c=i-j+1;if(s.getMaxRows()-c<2)s.insertRowsAfter(s.getMaxRows(),5);
   s.deleteRows(j+2,c);del+=c;i=j-1;}
  else i--;
 }
 return del;
}
function recap_(k,list){
 var R=[],ln=function(a){R.push(a);},num=function(v){v=Number(v);return isFinite(v)?v:0;},rd=function(v){return Math.round(v*1000)/1000;};
 var cnt=function(f){var o={};list.forEach(function(x){var v=f(x)||'-';o[v]=(o[v]||0)+1;});return o;};
 var sec=function(title,head,rows){ln([]);ln([title]);ln(head);rows.forEach(ln);};
 var cr=function(o){return Object.keys(o).sort().map(function(n){return [n,o[n]];});};
 ln(['Rekap '+ARCHL[k]]);ln(['Jumlah laporan',list.length]);
 if(k==='ac'){
  var a={};list.forEach(function(x){acRows_(x.d).forEach(function(r){if(acCond_(r)||r.roomTemp===''||r.roomTemp==null||!r.location)return;var v=Number(r.roomTemp);if(!isFinite(v))return;
   var o=a[r.location]=a[r.location]||{n:0,s:0,mn:1e9,mx:-1e9};o.n++;o.s+=v;o.mn=Math.min(o.mn,v);o.mx=Math.max(o.mx,v);});});
  sec('Suhu per area',['Area','Jumlah data','Rata-rata (°C)','Terendah (°C)','Tertinggi (°C)'],Object.keys(a).sort().map(function(n){var o=a[n];return [n,o.n,Math.round(o.s/o.n*100)/100,o.mn,o.mx];}));
 }else if(k==='work'){
  sec('Berdasarkan status',['Status','Jumlah'],cr(cnt(function(x){return x.d.status;})));
  sec('Berdasarkan kategori',['Kategori','Jumlah'],cr(cnt(function(x){return x.d.category;})));
 }else if(k==='meter_listrik'){
  var l={};list.forEach(function(x){var f=x.d.diffs||{},o=l[x.d.location]=l[x.d.location]||{n:0,a:0,b:0,c:0};o.n++;o.a+=num(f.lwbp);o.b+=num(f.wbp);o.c+=num(f.genset);});
  sec('Pemakaian per lokasi (jumlah selisih)',['Lokasi','Jumlah pencatatan','LWBP (kWh)','WBP (kWh)','Genset (kWh)'],Object.keys(l).sort().map(function(n){var o=l[n];return [n,o.n,rd(o.a),rd(o.b),rd(o.c)];}));
 }else if(k==='meter_air'){
  var w={};list.forEach(function(x){var o=w[x.d.location]=w[x.d.location]||{n:0,a:0};o.n++;o.a+=num(x.d.diff);});
  sec('Pemakaian air per lokasi',['Lokasi','Jumlah pencatatan','Selisih total (m³)'],Object.keys(w).sort().map(function(n){return [n,w[n].n,rd(w[n].a)];}));
 }else if(k==='meter_gabungan'){
  var gb={};list.forEach(function(x){var f=x.d.diffs||{},o=gb[x.d.location]=gb[x.d.location]||{n:0,a:0,b:0,c:0,w:0};o.n++;o.a+=num(f.lwbp);o.b+=num(f.wbp);o.c+=num(f.genset);o.w+=num(x.d.diff);});
  sec('Pemakaian per lokasi (jumlah selisih)',['Lokasi','Jumlah pencatatan','LWBP (kWh)','WBP (kWh)','Genset (kWh)','Air (m³)'],Object.keys(gb).sort().map(function(n){var o=gb[n];return [n,o.n,rd(o.a),rd(o.b),rd(o.c),rd(o.w)];}));
 }else if(k==='purchase'){
  var g=0,st={};list.forEach(function(x){g+=num(x.d.total);var o=st[x.d.store]=st[x.d.store]||{n:0,t:0};o.n++;o.t+=num(x.d.total);});
  ln(['Total belanja (Rp)',g]);
  sec('Per toko / pemasok',['Toko','Jumlah transaksi','Total (Rp)'],Object.keys(st).sort().map(function(n){return [n,st[n].n,st[n].t];}));
 }else if(k==='audience'){
  var au={},tt=0,tp=0;list.forEach(function(x){(x.d.studios||[]).forEach(function(a){var o=au[a.studio]=au[a.studio]||{n:0,t:0,p:0};(a.shows||[]).forEach(function(w){o.n++;o.t+=num(w.tiket);o.p+=num(w.penonton);});});tt+=num(x.d.totalTiket);tp+=num(x.d.totalPenonton);});
  ln(['Total tiket terjual',tt]);ln(['Total penonton masuk',tp]);
  sec('Per studio',['Studio','Jumlah show','Tiket terjual','Penonton masuk'],Object.keys(au).sort().map(function(n){return [n,au[n].n,au[n].t,au[n].p];}));
 }else if(k==='issue'){
  sec('Berdasarkan status',['Status','Jumlah'],cr(cnt(function(x){return x.d.status;})));
  sec('Berdasarkan prioritas',['Prioritas','Jumlah'],cr(cnt(function(x){return x.d.priority;})));
 }
 var w2=0;R.forEach(function(r){if(r.length>w2)w2=r.length;});
 return R.map(function(r){var o=r.slice();while(o.length<w2)o.push('');return o;});
}
function arsipSatu_(k,m,list,by){
 var t=keyType_(k),mm=MIR[k],all={},res='',
  live=list.filter(function(x){return x.r.status!=='Diarsipkan';}).sort(function(a,b){return Number(a.r.createdAt)-Number(b.r.createdAt);});
 list.forEach(function(x){all[String(x.r.id)]=1;});
 if(live.length){
  var head=fullHead_(k),n=mm.head.length,prow=[],
   rows=live.map(function(x){return rowSafe_(mm.row(x.r,x.d)).concat(phCells_(k,phInfo_(by[x.r.id]||[],x.r)));});
  live.forEach(function(x){(by[x.r.id]||[]).slice().sort(function(a,b){return Number(a.uploadedAt)-Number(b.uploadedAt);}).forEach(function(p,i){prow.push(photoRow_(p,i+1));});});
  var nm='Arsip '+ARCHL[k]+' '+m,fol=archFolder_(k),it=fol.getFilesByName(nm);
  while(it.hasNext())it.next().setTrashed(true); /* sisa percobaan sebelumnya yang belum selesai */
  var ss=SpreadsheetApp.create(nm);DriveApp.getFileById(ss.getId()).moveTo(fol);
  var s1=ss.getSheets()[0].setName('Laporan');
  s1.getRange(1,1,1,head.length).setValues([head]).setFontWeight('bold').setBackground('#0a3a2a').setFontColor('#ffffff');
  for(var c=1;c<=n;c++){if(mm.num.indexOf(c)<0)s1.getRange(2,c,rows.length,1).setNumberFormat('@');}
  putRows_(s1,2,1,rows);
  mm.wrap.forEach(function(c){s1.getRange(2,c,rows.length,1).setWrap(true);s1.setColumnWidth(c,320);});
  s1.setFrozenRows(1);
  var s2=ss.insertSheet('Foto');
  s2.getRange(1,1,1,PHEAD.length).setValues([PHEAD]).setFontWeight('bold').setBackground('#0a3a2a').setFontColor('#ffffff');
  if(prow.length){
   s2.getRange(2,1,prow.length,1).setNumberFormat('@');s2.getRange(2,7,prow.length,1).setNumberFormat('@');
   putRows_(s2,2,1,prow);if(PUBLIC_PHOTOS)s2.setRowHeights(2,prow.length,100);
  }
  s2.setColumnWidth(4,160);s2.setFrozenRows(1);
  var rc=recap_(k,live),s3=ss.insertSheet('Rekap');
  s3.getRange(1,1,rc.length,rc[0].length).setValues(rc);s3.getRange(1,1).setFontWeight('bold').setFontSize(13);s3.setColumnWidth(1,220);
  SpreadsheetApp.flush();
  if(s1.getLastRow()-1!==live.length||s2.getLastRow()-1!==prow.length)throw new Error('Verifikasi arsip gagal: jumlah baris tidak sama.');
  res=nm+' ('+live.length+' laporan, '+prow.length+' foto)';
 }
 /* baru setelah terverifikasi: kosongkan dari sheet bulan berjalan */
 delRows_(mm.name,all);delRows_(PHS[t],all);archMark_(k,m);
 return res||('(tanpa laporan aktif) '+ARCHL[k]+' '+m);
}
function arsipBulan_(){
 var cur=mo_(Date.now()),reps=all_('REPORTS'),by={},made=[];
 var mrecs=meterRecs_();
 all_('PHOTOS').forEach(function(p){if(p.fileId)(by[p.reportId]=by[p.reportId]||[]).push(p);});
 Object.keys(MIR).forEach(function(k){
  var t=keyType_(k),g={};
  reps.forEach(function(r){
   if(r.type!==t)return;var d;try{d=JSON.parse(r.data);}catch(e){return;}
   if(r.type==='meter')meterDiff_(d,r.id,mrecs);
   if(mirKey_(r.type,d)!==k)return;
   var m=mo_(r.createdAt);if(m>=cur||isArch_(k,m))return;
   (g[m]=g[m]||[]).push({r:r,d:d});
  });
  Object.keys(g).sort().forEach(function(m){
   try{made.push(arsipSatu_(k,m,g[m],by));}catch(e){Logger.log('Arsip GAGAL '+k+' '+m+': '+e);}
  });
 });
 return made;
}
function arsipBulananJob(){
 var lock=LockService.getScriptLock();lock.waitLock(60000);
 try{var r=arsipBulan_();Logger.log(r.length?('Arsip dibuat:\n'+r.join('\n')):'Tidak ada bulan yang perlu diarsipkan.');return r;}
 finally{lock.releaseLock();}
}
/* otomatis: dicek sekali per hari saat ada yang membuka website */
function autoArsip_(){
 var p=P_(),day=Utilities.formatDate(new Date(),TZ,'yyyyMMdd');
 if(p.getProperty('ARCH_DAY')===day)return;
 /* tandai "selesai" hanya bila berhasil; bila gagal, coba lagi paling cepat 1 jam kemudian (bukan tiap permintaan, bukan besok) */
 var tr=Number(p.getProperty('ARCH_TRY'))||0;if(Date.now()-tr<3600000)return;
 p.setProperty('ARCH_TRY',String(Date.now()));
 try{arsipBulananJob();p.setProperty('ARCH_DAY',day);}catch(e){Logger.log('Arsip otomatis gagal: '+e);}
}
/* JALANKAN SEKALI dari editor: pemicu tanggal 1 tiap bulan (sekitar pukul 01.00) + langsung mengarsipkan bulan lama */
function aktifkanArsipBulanan(){
 ScriptApp.getProjectTriggers().forEach(function(t){if(t.getHandlerFunction()==='arsipBulananJob')ScriptApp.deleteTrigger(t);});
 ScriptApp.newTrigger('arsipBulananJob').timeBased().onMonthDay(1).atHour(1).create();
 arsipBulananJob();Logger.log('Pemicu arsip bulanan aktif (tanggal 1, sekitar pukul 01.00).');
}

/* ===== v10: AUTO-HAPUS ATS, GLOSARIUM, PENYIMPANAN DATA ===== */
var SYS_={userId:'SYSTEM',fullName:'Sistem (Auto-Hapus)',role:'SYSTEM'}, ATS_HARD_DELETE=false, MAXSTORE=6000000;
var STORE_EXT={pdf:'application/pdf',doc:'application/msword',docx:'application/vnd.openxmlformats-officedocument.wordprocessingml.document',xls:'application/vnd.ms-excel',xlsx:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',ppt:'application/vnd.ms-powerpoint',pptx:'application/vnd.openxmlformats-officedocument.presentationml.presentation',csv:'text/csv',txt:'text/plain',json:'application/json',png:'image/png',jpg:'image/jpeg',jpeg:'image/jpeg',webp:'image/webp',zip:'application/zip'};
/* --- Auto-hapus ATS: film yang countdown hapus-nya habis dibuang permanen (baris sheet + poster) --- */
function atsExp_(f){var t=Number(f.deleteAt);return t>0&&t<=Date.now();}
function atsLive_(){
 var r=all_('ATS_FILMS');
 if(r.some(atsExp_)){try{atsPurge_();r=all_('ATS_FILMS');}catch(e){Logger.log('Auto-hapus ATS gagal: '+e);}}
 return r.filter(function(f){return !atsExp_(f);});
}
function killFile_(id){
 try{
  if(ATS_HARD_DELETE&&typeof Drive!=='undefined'&&Drive.Files&&Drive.Files.remove){Drive.Files.remove(id);return;}
  DriveApp.getFileById(id).setTrashed(true);
 }catch(e){Logger.log('Berkas Drive tidak dapat dihapus ('+id+'): '+e);}
}
function atsPurge_(){
 var lock=LockService.getScriptLock();
 if(!lock.tryLock(10000))return 0;
 try{
  var rows=all_('ATS_FILMS').filter(atsExp_);
  if(!rows.length)return 0;
  var sh=ss_().getSheetByName('ATS_FILMS');
  rows.sort(function(a,b){return b._row-a._row;}).forEach(function(f){
   sh.deleteRow(f._row);                    /* baris dihapus dulu; poster baru dibuang setelahnya */
   if(f.posterFileId)killFile_(f.posterFileId);
   log_(SYS_,'HAPUS_OTOMATIS_ATS','ATS_FILMS',f.id,f.title+' (countdown hapus habis)');
  });
  return rows.length;
 }finally{lock.releaseLock();}
}
function atsPurgeJob(){Logger.log('Auto-hapus ATS: '+atsPurge_()+' film dihapus.');}
/* JALANKAN SEKALI dari editor: pemicu tiap 1 jam sebagai cadangan (saat website dibuka, pemeriksaan juga berjalan otomatis) */
function aktifkanAutoHapusATS(){
 ScriptApp.getProjectTriggers().forEach(function(t){if(t.getHandlerFunction()==='atsPurgeJob')ScriptApp.deleteTrigger(t);});
 ScriptApp.newTrigger('atsPurgeJob').timeBased().everyHours(1).create();
 atsPurgeJob();Logger.log('Pemicu auto-hapus ATS aktif (tiap 1 jam).');
}
/* --- Penyimpanan data multi-format: validasi isi berkas terhadap ekstensi --- */
function storeCheck_(b,ext){
 var a=b[0]&255,c=b[1]&255,d=b[2]&255,e=b[3]&255,ok=true;
 if(ext==='pdf')ok=a===0x25&&c===0x50&&d===0x44&&e===0x46;
 else if(['docx','xlsx','pptx','zip'].indexOf(ext)>=0)ok=a===0x50&&c===0x4B;
 else if(['doc','xls','ppt'].indexOf(ext)>=0)ok=a===0xD0&&c===0xCF&&d===0x11&&e===0xE0;
 else if(['png','jpg','jpeg','webp'].indexOf(ext)>=0){var k=imgKind_(b);ok=!!k&&k===STORE_EXT[ext];}
 else{
  for(var i=0;i<Math.min(b.length,4000);i++)if(b[i]===0)ok=false;
  if(ok&&ext==='json'){try{JSON.parse(Utilities.newBlob(b).getDataAsString('UTF-8'));}catch(x){bad_('Isi berkas JSON tidak valid.');}}
 }
 if(!ok)bad_('Isi berkas tidak sesuai dengan ekstensinya (.'+ext+').');
}
function storeFile_(q){
 var f=q.file;if(!f||!f.b64)return null;
 var name=s_(f.name,120).replace(/[\\\/:*?"<>|]+/g,'_'),m=/\.([A-Za-z0-9]+)$/.exec(name),ext=m?m[1].toLowerCase():'';
 if(!STORE_EXT[ext])bad_('Format tidak didukung. Diizinkan: '+Object.keys(STORE_EXT).join(', ')+'.');
 var b=Utilities.base64Decode(String(f.b64).replace(/^data:[^,]*,/,''));
 if(!b.length)bad_('Berkas kosong.');
 if(b.length>MAXSTORE)bad_('Ukuran berkas maksimal 6 MB.');
 storeCheck_(b,ext);
 return {name:name,ext:ext,mime:STORE_EXT[ext],bytes:b};
}

/* ===== migrasi skema otomatis (sekali setelah kode baru dipasang) ===== */
function ensureSchema_(){
 var p=P_();if(p.getProperty('SCHEMA')===SCHEMA)return;
 var lock=LockService.getScriptLock();lock.waitLock(30000);
 try{
  if(p.getProperty('SCHEMA')===SCHEMA)return;
  var ss=ss_();
  Object.keys(SH).forEach(function(n){var s=ss.getSheetByName(n);
   if(!s){s=ss.insertSheet(n);s.appendRow(SH[n]);s.setFrozenRows(1);if(n==='AUDIENCE_SHOWS')s.getRange(2,1,Math.max(1,s.getMaxRows()-1),SH[n].length).setNumberFormat('@');return;}
   var h=s.getRange(1,1,1,Math.max(1,s.getLastColumn())).getValues()[0].map(String);
   SH[n].forEach(function(k){if(h.indexOf(k)<0){h.push(k);s.getRange(1,h.length).setValue(k);}});});
  var stale=Object.keys(MIR).some(function(k){var s=ss.getSheetByName(MIR[k].name);
   return s&&String(s.getRange(1,1,1,Math.max(1,s.getLastColumn())).getValues()[0])!==String(fullHead_(k));});
  var noPh=Object.keys(PHS).some(function(t){return !ss.getSheetByName(PHS[t]);});
  if(stale||noPh)try{rebuildMirrors_();}catch(e){Logger.log('Rebuild sheet kategori gagal: '+e);}
  p.setProperty('SCHEMA',SCHEMA);
 }finally{lock.releaseLock();}
}

/* ===== util ===== */
function bad_(m){throw new Error('!'+m);}
function P_(){return PropertiesService.getScriptProperties();}
/* spreadsheet dibuka sekali per permintaan (openById lambat dan dulu dipanggil puluhan kali) */
var SS_=null,SS_ID='';
function ss_(){var id=P_().getProperty('SHEET_ID');if(!id)bad_('Setup belum dijalankan.');if(!SS_||SS_ID!==id){SS_=SpreadsheetApp.openById(id);SS_ID=id;}return SS_;}
function all_(n){var v=ss_().getSheetByName(n).getDataRange().getValues(),h=v.shift();return v.map(function(r,i){var o={_row:i+2};h.forEach(function(k,j){o[k]=r[j];});return o;});}
/* teks dari pengguna yang diawali = + - @ atau tab/CR dibaca Sheets sebagai rumus; diberi apostrof agar tersimpan sebagai teks. */
function cellSafe_(v){return (typeof v==='string'&&/^[=+\-@\t\r]/.test(v))?"'"+v:v;}
/* teks buatan pengguna yang ditulis lewat setValues: awalan = + - @ ' tab/CR diberi spasi agar tidak dibaca sebagai rumus atau tautan HYPERLINK */
function safeTxt_(v){return (typeof v==='string'&&/^['=+\-@\t\r]/.test(v))?' '+v:v;}
function rowSafe_(a){return a.map(safeTxt_);}
function add_(n,o){ss_().getSheetByName(n).appendRow(SH[n].map(function(k){return o[k]===undefined?'':cellSafe_(o[k]);}));}
function upd_(n,row,o){var s=ss_().getSheetByName(n);SH[n].forEach(function(k,j){if(o[k]!==undefined)s.getRange(row,j+1).setValue(cellSafe_(o[k]));});}
function uid_(p){return p+'-'+Utilities.getUuid().replace(/-/g,'').slice(0,10);}
function hx_(b){return b.map(function(x){return ('0'+(x&255).toString(16)).slice(-2);}).join('');}
function hash_(pw,salt){var h=pw;for(var i=0;i<1000;i++)h=hx_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,salt+h+pw,Utilities.Charset.UTF_8));return h;}
function s_(v,m){return String(v==null?'':v).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g,'').trim().slice(0,m||500);}
/* Username case-sensitive: 'Admin' dan 'admin' adalah dua nama berbeda saat login. */
function find_(un){un=String(un);return all_('USERS').filter(function(u){return String(u.username)===un;})[0];}
/* Pembuatan akun tetap mencegah nama yang hanya berbeda huruf besar/kecil agar tidak membingungkan. */
function findCI_(un){un=String(un).toLowerCase();return all_('USERS').filter(function(u){return String(u.username).toLowerCase()===un;})[0];}
function pub_(u){return {userId:u.userId,name:u.fullName,username:u.username,role:u.role,mustChange:u.mustChange==='Y',photo:!!u.photoFileId};}
/* versi cache notifikasi: setiap penulisan (semua aksi tulis memanggil log_) mengganti versi, sehingga cache lama tidak terpakai lagi */
function npBump_(){try{CacheService.getScriptCache().put('npv',Utilities.getUuid(),21600);}catch(e){Logger.log('npBump_: '+e);}}
function log_(u,act,cat,rid,sum){add_('AUDIT_LOG',{logId:uid_('L'),userId:u.userId,userName:u.fullName,role:u.role,action:act,category:cat,recordId:rid||'',at:Date.now(),summary:s_(sum,300)});npBump_();}
function folder_(parent,name){var it=parent?parent.getFoldersByName(name):DriveApp.getFoldersByName(name);return it.hasNext()?it.next():(parent?parent.createFolder(name):DriveApp.createFolder(name));}
function chkPw_(p){if(String(p||'').length<8)bad_('Password minimal 8 karakter.');}

/* ===== token & foto ===== */
function secret_(){var p=P_(),v=p.getProperty('TOKEN_SECRET');if(!v){v=Utilities.getUuid()+Utilities.getUuid();p.setProperty('TOKEN_SECRET',v);}return v;}
function purl_(fid,w){return 'https://lh3.googleusercontent.com/d/'+fid+'=w'+(w||800);}
function shareFile_(f){if(!PUBLIC_PHOTOS)return true;try{f.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);if(f.getSharingAccess()!==DriveApp.Access.ANYONE_WITH_LINK)throw new Error('Izin ANYONE_WITH_LINK tidak terverifikasi');return true;}catch(e){Logger.log('Gagal memvalidasi izin Drive: '+e);throw new Error('Izin akses Google Drive gagal diverifikasi. Periksa kebijakan berbagi Drive.');}}
/* folder kategori (dibuat otomatis bila belum ada) */
function catFolder_(t){
 var p=P_(),id=p.getProperty('F_'+t);
 if(id)try{return DriveApp.getFolderById(id);}catch(e){}
 var f=folder_(folder_(null,'STC XXI TECHNICAL REPORT'),FOLDERS[t]);p.setProperty('F_'+t,f.getId());return f;
}
/* satu folder Drive per laporan: <kategori>/<ID laporan> */
function repFolder_(rep){
 if(rep.folderId)try{return DriveApp.getFolderById(rep.folderId);}catch(e){}
 var f=folder_(catFolder_(rep.type),rep.id);shareFile_(f);upd_('REPORTS',rep._row,{folderId:f.getId()});rep.folderId=f.getId();return f;
}
function imgKind_(b){
 if(!b||b.length<12)return null;var a=b[0]&255,c=b[1]&255,d=b[2]&255,e=b[3]&255;
 if(a===0xFF&&c===0xD8)return 'image/jpeg';
 if(a===0x89&&c===0x50&&d===0x4E&&e===0x47)return 'image/png';
 if(a===0x52&&c===0x49&&d===0x46&&e===0x46&&(b[8]&255)===0x57)return 'image/webp';
 return null;
}
var EXT={'image/jpeg':'jpg','image/png':'png','image/webp':'webp'};
function addPhotos_(rid,photos){
 var rep=all_('REPORTS').filter(function(r){return r.id===rid;})[0];if(!rep)bad_('Laporan tidak ditemukan.');
 var fol=repFolder_(rep),n=all_('PHOTOS').filter(function(r){return r.reportId===rid;}).length,add=0,fail=0,ps=ss_().getSheetByName('PHOTOS');
 (photos||[]).slice(0,PHBATCH).forEach(function(ph){try{
  if(n>=MAXPH)throw new Error('melebihi batas '+MAXPH+' foto');
  var b=Utilities.base64Decode(String(ph&&ph.b64||'')),mime=imgKind_(b);if(!mime||b.length>MAXB)throw new Error('format/ukuran tidak valid');
  var kd=rep.type==='work'?(['Sebelum','Sesudah'].indexOf(ph&&ph.kind)>=0?ph.kind:'Foto'):rep.type==='ac'?(s_(ph&&ph.kind,60)||'Foto'):rep.type==='issue'?(['Masalah','Penyelesaian'].indexOf(ph&&ph.kind)>=0?ph.kind:'Masalah'):'Foto';
  var f=fol.createFile(Utilities.newBlob(b,mime,rid+'-'+('0'+(n+1)).slice(-2)+'.'+EXT[mime]));
  shareFile_(f);var fid=f.getId();n++;add++;
  var uploadedAt=Date.now(),capturedAt=Number(ph&&ph.capturedAt)||uploadedAt;
  add_('PHOTOS',{photoId:uid_('PH'),reportId:rid,fileId:fid,kind:kd,uploadedAt:uploadedAt,capturedAt:capturedAt,url:PUBLIC_PHOTOS?purl_(fid,1200):'',preview:''});
  try{ps.setRowHeight(ps.getLastRow(),90);}catch(e2){}
 }catch(e){fail++;Logger.log('foto gagal: '+e);}});
 return {added:add,failed:fail,total:n};
}
/* perbarui kolom foto di REPORTS dan di sheet kategori */
function syncFoto_(rid){
 var rep=all_('REPORTS').filter(function(r){return r.id===rid;})[0];if(!rep)return;
 var info=phInfo_(all_('PHOTOS').filter(function(p){return p.reportId===rid;}),rep);if(!info.n)return;
 var s=ss_().getSheetByName('REPORTS');
 /* Foto tidak ditanam ke Spreadsheet; tautan Drive dan ZIP menjadi rujukan dokumentasi. */
 try{s.getRange(rep._row,SH.REPORTS.indexOf('foto')+1).clearContent();}catch(e0){}
 try{var mk=mirKey_(rep.type,JSON.parse(rep.data)),loc=mirrorLoc_(mk,rid);
  if(loc&&loc.row){var cells=phCells_(mk,info);putRows_(loc.s,loc.row,MIR[mk].head.length+1,[cells]);if(PUBLIC_PHOTOS)loc.s.setRowHeight(loc.row,100);}}
 catch(e){Logger.log(e);}
}
/* ZIP seluruh foto laporan, disimpan di folder laporan (dibuat ulang bila jumlah foto berubah) */
function makeZip_(rid){
 var lock=LockService.getScriptLock();lock.waitLock(30000);
 try{
  var rep=all_('REPORTS').filter(function(r){return r.id===rid;})[0];if(!rep)bad_('Laporan tidak ditemukan.');
  var ph=all_('PHOTOS').filter(function(p){return p.reportId===rid&&p.fileId;}).sort(function(a,b){return Number(a.uploadedAt)-Number(b.uploadedAt);});
  if(!ph.length)bad_('Laporan ini belum memiliki foto.');
  var fol=repFolder_(rep),it=fol.getFilesByName(rid+'.zip'),cur=null;
  while(it.hasNext()){var f=it.next();if(!cur&&f.getDescription()==='n='+ph.length)cur=f;else f.setTrashed(true);}
  if(!cur){var cnt={},blobs=ph.map(function(p){var b=DriveApp.getFileById(p.fileId).getBlob(),k=rep.type==='ac'?String(p.kind||'Foto').replace(/[^A-Za-z0-9_-]+/g,'_'):(p.kind==='Sebelum'||p.kind==='Sesudah')?p.kind:'Foto';
    cnt[k]=(cnt[k]||0)+1;b.setName(k+'_'+('0'+cnt[k]).slice(-2)+'.'+(EXT[b.getContentType()]||'jpg'));return b;});
   cur=fol.createFile(Utilities.zip(blobs,rid+'.zip'));cur.setDescription('n='+ph.length);shareFile_(cur);}
  upd_('REPORTS',rep._row,{zipId:cur.getId()});
  return {url:zipUrl_(cur.getId()),count:ph.length};
 }finally{lock.releaseLock();}
}
/* JALANKAN SEKALI bila perlu: membagikan ulang file foto lama & membangun ulang sheet kategori */
function perbaikiFoto(){
 var ps=ss_().getSheetByName('PHOTOS'),cu=SH.PHOTOS.indexOf('url')+1,cp=SH.PHOTOS.indexOf('preview')+1;
 all_('PHOTOS').forEach(function(p){if(!p.fileId)return;
  try{shareFile_(DriveApp.getFileById(p.fileId));}catch(e){Logger.log('File tidak ditemukan: '+p.fileId);return;}
  ps.getRange(p._row,cu).setValue(purl_(p.fileId,1200));ps.getRange(p._row,cp).clearContent();});
 all_('ATS_FILMS').forEach(function(f){if(f.posterFileId)try{shareFile_(DriveApp.getFileById(f.posterFileId));}catch(e){}});
 ps.setColumnWidth(cp,160);rebuildMirrors_();Logger.log('Selesai.');
}
/* hapus foto profil LAMA milik pengguna ini (ke tempat sampah Drive, dapat dipulihkan 30 hari).
   Hanya dihapus bila: file ada di folder PROFILE PHOTOS, namanya diawali AVATAR-<userId>-, dan tidak dipakai pengguna lain.
   File default / aset bersama / file lain tidak pernah disentuh. Kegagalan hanya dicatat, tidak membatalkan penggantian. */
/* pastikan semua kolom USERS (termasuk photoFileId) ada di header sheet; aman dijalankan berulang, tidak mengubah data */
function ensureUserCols_(){var s=ss_().getSheetByName('USERS'),h=s.getRange(1,1,1,Math.max(1,s.getLastColumn())).getValues()[0].map(String);
 SH.USERS.forEach(function(k){if(h.indexOf(k)<0){h.push(k);s.getRange(1,h.length).setValue(k);}});}
function delOldAvatar_(u,fid,fol){
 if(!fid)return false;
 try{
  if(all_('USERS').some(function(x){return x.userId!==u.userId&&String(x.photoFileId)===String(fid);}))return false;
  var f=DriveApp.getFileById(fid);
  if(f.getName().indexOf('AVATAR-'+u.userId+'-')!==0)return false;
  var ps=f.getParents(),ok=false;while(ps.hasNext()){if(ps.next().getId()===fol.getId())ok=true;}
  if(!ok)return false;
  f.setTrashed(true);return true;
 }catch(e){Logger.log('Foto profil lama tidak ditemukan / gagal dihapus: '+e);return false;}
}

/* ===== setup (jalankan sekali dari editor) ===== */
function setupAwal(){
 var p=P_(),ss=SpreadsheetApp.getActiveSpreadsheet();
 if(!ss)bad_('Proyek harus dibuat dari dalam Google Spreadsheet (Ekstensi > Apps Script).');
 var acc=[['Technical','Technical','USER','2121'],['Syafar','Syafar','ADMIN','master2121']];
 p.setProperty('SHEET_ID',ss.getId());
 Object.keys(SH).forEach(function(n){var s=ss.getSheetByName(n)||ss.insertSheet(n);
  if(s.getLastRow()===0){s.getRange(1,1,s.getMaxRows(),SH[n].length).setNumberFormat('@');s.appendRow(SH[n]);s.setFrozenRows(1);}});
 Object.keys(FOLDERS).forEach(function(k){catFolder_(k);});
 Object.keys(MIR).forEach(function(t){mirrorSheet_(t);});
 Object.keys(PHS).forEach(function(t){photoSheet_(t);});
 var made=[];
 acc.forEach(function(a){if(!find_(a[1])){var salt=Utilities.getUuid(),now=Date.now();
  add_('USERS',{userId:uid_('U'),fullName:a[0],username:a[1],role:a[2],status:'Aktif',salt:salt,hash:hash_(a[3],salt),mustChange:'Y',createdAt:now,updatedAt:now});made.push(a[1]);}});
 p.setProperty('SCHEMA',SCHEMA);
 Logger.log('Setup selesai. Akun baru: '+(made.join(', ')||'tidak ada (sudah ada)')+'. Password awal wajib diganti saat login pertama.');
}



/* ===== AUTO-HAPUS LAPORAN > 45 HARI SEJAK DITERBITKAN =====
 * Jalankan aktifkanAutoHapusLaporan45Hari() satu kali dari editor Apps Script (menggantikan pemicu lama autoHapusLaporan30HariJob).
 * Dasar umur: createdAt (waktu laporan diterbitkan); semua data report utama, foto terdaftar, ZIP, folder per laporan,
 * serta laporan film dan tanda tangan terkait. File Drive yang direferensikan entitas lain dilindungi.
 */
var REPORT_RETENTION_DAYS = 45;
function fileReferencedOutsideReport_(fid, reportId, filmRowId) {
 if (!fid) return false;
 var id=String(fid);
 // Foto atau ZIP yang masih direferensikan laporan lain tetap dilindungi.
 if (all_('PHOTOS').some(function(p){return String(p.fileId||'')===id && String(p.reportId||'')!==String(reportId||'');})) return true;
 if (all_('REPORTS').some(function(r){return String(r.id||'')!==String(reportId||'') && (String(r.zipId||'')===id || String(r.folderId||'')===id);})) return true;
 // Lindungi file yang dipakai oleh data film, ATS, penyimpanan, profil pengguna, atau foto laporan lain.
 if (all_('FILM_REPORTS').some(function(r){return String(r.id||'')!==String(filmRowId||'') && String(r.signature||'')===id;})) return true;
 if (all_('ATS_FILMS').some(function(r){return String(r.posterFileId||'')===id;})) return true;
 if (all_('STORAGE').some(function(r){return String(r.fileId||'')===id;})) return true;
 if (all_('USERS').some(function(r){return String(r.photoFileId||'')===id;})) return true;
 return false;
}
function trashReportFileIfSafe_(fid, reportId, filmRowId) {
 if (!fid || fileReferencedOutsideReport_(fid,reportId,filmRowId)) return false;
 try { DriveApp.getFileById(String(fid)).setTrashed(true); return true; }
 catch(e) { Logger.log('File laporan tidak dapat dipindahkan ke Sampah ('+fid+'): '+e); return false; }
}
function trashReportFolderIfSafe_(folderId, reportId) {
 if (!folderId) return false;
 // Folder hanya milik satu laporan jika ID-nya tidak direferensikan laporan lain.
 if (all_('REPORTS').some(function(r){return String(r.id||'')!==String(reportId||'') && String(r.folderId||'')===String(folderId);})) return false;
 try {
  var f=DriveApp.getFolderById(String(folderId));
  // Hanya buang file di folder ini jika file tidak dipakai entitas lain.
  var files=f.getFiles(), retained=false;
  while(files.hasNext()) { var child=files.next(); if(fileReferencedOutsideReport_(child.getId(),reportId,'')) retained=true; else trashReportFileIfSafe_(child.getId(),reportId,''); }
  // Jangan buang folder jika ada file yang masih direferensikan entitas lain.
  if(retained) return false;
  f.setTrashed(true); return true;
 } catch(e) { Logger.log('Folder laporan tidak dapat dipindahkan ke Sampah ('+folderId+'): '+e); return false; }
}
function autoHapusLaporan45HariJob() {
 var lock=LockService.getScriptLock(); if(!lock.tryLock(10000)) return {reports:0,filmReports:0,photos:0};
 var result={reports:0,filmReports:0,photos:0,driveFiles:0};
 try {
  var cutoff=Date.now()-REPORT_RETENTION_DAYS*24*60*60*1000, ss=ss_();
  var reports=all_('REPORTS').filter(function(r){return r.type!=='meter' && Number(r.createdAt)>0 && Number(r.createdAt)<cutoff;});
  // Simpan referensi file sebelum menghapus baris PHOTOS/REPORTS.
  var ph=all_('PHOTOS'), shPh=ss.getSheetByName('PHOTOS'), shRep=ss.getSheetByName('REPORTS'), expiredIds={};
  reports.forEach(function(r){expiredIds[String(r.id)]=true;});
  reports.sort(function(a,b){return b._row-a._row;}).forEach(function(r){
   var photos=ph.filter(function(x){return String(x.reportId)===String(r.id);});
   photos.forEach(function(x){ if(trashReportFileIfSafe_(x.fileId,r.id,'')) result.driveFiles++; });
   if(r.zipId && trashReportFileIfSafe_(r.zipId,r.id,'')) result.driveFiles++;
   if(r.folderId && trashReportFolderIfSafe_(r.folderId,r.id)) result.driveFiles++;
   // Metadata PHOTOS dihapus sekali di bawah dengan indeks baris menurun agar tidak menggeser indeks berikutnya.
   try { shRep.deleteRow(r._row); result.reports++; }
   catch(e) { Logger.log('Gagal menghapus laporan '+r.id+': '+e); }
   log_(SYS_,'AUTO_HAPUS_LAPORAN','REPORTS',r.id,'Laporan dan lampiran melewati retensi 45 hari.');
  });
  // Hapus metadata foto kedaluwarsa secara menurun agar nomor baris tetap akurat.
  ph.filter(function(x){return !!expiredIds[String(x.reportId)];}).sort(function(a,b){return b._row-a._row;}).forEach(function(x){try{shPh.deleteRow(x._row);result.photos++;}catch(e){Logger.log('Gagal menghapus metadata foto '+x.photoId+': '+e);}});
  // Hapus juga laporan film yang dibuat lebih dari 45 hari lalu dan tanda tangan eksklusifnya.
  var film=all_('FILM_REPORTS').filter(function(r){return Number(r.createdAt)>0 && Number(r.createdAt)<cutoff;});
  var shFilm=ss.getSheetByName('FILM_REPORTS');
  film.sort(function(a,b){return b._row-a._row;}).forEach(function(r){
   if(r.signature && trashReportFileIfSafe_(r.signature,'',r.id)) result.driveFiles++;
   try { shFilm.deleteRow(r._row); result.filmReports++; log_(SYS_,'AUTO_HAPUS_LAPORAN_FILM','FILM_REPORTS',r.id,'Laporan film melewati retensi 45 hari.'); }
   catch(e) { Logger.log('Gagal menghapus laporan film '+r.id+': '+e); }
  });
  if(result.reports || result.photos) { try { rebuildMirrors_(); } catch(e) { Logger.log('Rebuild sheet kategori setelah penghapusan gagal: '+e); } }
  Logger.log('Auto-hapus 45 hari selesai: '+JSON.stringify(result));
  return result;
 } finally { lock.releaseLock(); }
}
function aktifkanAutoHapusLaporan45Hari() {
 /* hapus pemicu lama (nama 30 hari) dan pemicu baru yang mungkin sudah ada, lalu pasang satu pemicu baru */
 ScriptApp.getProjectTriggers().forEach(function(t){var h=t.getHandlerFunction();if(h==='autoHapusLaporan30HariJob'||h==='autoHapusLaporan45HariJob')ScriptApp.deleteTrigger(t);});
 ScriptApp.newTrigger('autoHapusLaporan45HariJob').timeBased().everyDays(1).atHour(3).create();
 Logger.log('Pemicu auto-hapus laporan (masa simpan 45 hari) aktif. Penghapusan pertama dijalankan pada jadwal berikutnya, sekitar pukul 03.00.');
}

/* ===== entry ===== */
function doGet(){return ContentService.createTextOutput('STC XXI API aktif | versi: v20.1.0-penonton-per-show');}
function doPost(e){
 var out;
 try{out={ok:true,data:route_(JSON.parse(e.postData.contents))};}
 catch(err){var m=(err&&err.message)||'';
  if(m.charAt(0)==='!'){out=m.charAt(1)==='@'?{ok:false,code:'AUTH',message:m.slice(2)}:{ok:false,message:m.slice(1)};}
  else{Logger.log(err&&err.stack||err);out={ok:false,message:'Terjadi kesalahan server: '+m};}}
 out.now=Date.now();
 return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
/* Blokir login bersifat sementara (LOCK_MS) agar orang luar tidak bisa mengunci akun selamanya; Admin juga bisa membuka lebih awal lewat Pengguna. */
function isLocked_(u){var l=Number(u.lockedAt);return l>0&&Date.now()-l<LOCK_MS;}
function sig_(u,exp){return hx_(Utilities.computeHmacSha256Signature(u.userId+'.'+exp,secret_()+u.hash));}
function mk_(u){var exp=Date.now()+TOKEN_TTL;return u.userId+'.'+exp+'.'+sig_(u,exp);}
function session_(t){
 var p=String(t||'').split('.'),u=p.length===3&&all_('USERS').filter(function(x){return x.userId===p[0];})[0];
 if(!u||u.status!=='Aktif'||(u.role!=='ADMIN'&&isLocked_(u))||!(Number(p[1])>Date.now())||sig_(u,p[1])!==p[2])bad_('@Sesi berakhir, silakan login ulang.');
 /* Idle timeout 30 menit, ditegakkan di server. Aktivitas terakhir disimpan di cache; bila cache hilang,
    token hanya diterima selama 30 menit sejak diterbitkan (iat = exp - TOKEN_TTL). */
 var c=CacheService.getScriptCache(),k='ls_'+u.userId,now=Date.now(),last=Number(c.get(k)||0);
 if(!last)last=Number(p[1])-TOKEN_TTL;
 if(now-last>IDLE_MS){try{c.remove(k);}catch(e){}bad_('@Sesi berakhir karena tidak aktif 30 menit, silakan login ulang.');}
 try{c.put(k,String(now),21600);}catch(e){}
 return u;
}
function route_(q){
 var own=Object.prototype.hasOwnProperty,act=String(q.action);
 /* hanya aksi yang benar-benar didefinisikan; nama bawaan JavaScript ("constructor", "toString") bukan aksi */
 var h=own.call(H,act)?H[act]:(typeof DCP_H!=='undefined'&&own.call(DCP_H,act)?DCP_H[act]:null);if(typeof h!=='function')bad_('Aksi tidak dikenal.');
 try{ensureSchema_();}catch(e){Logger.log('Migrasi skema gagal: '+e);}
 if(q.action==='login')return h(q);
 var u=session_(q.token);
 if(u.mustChange==='Y'&&['changepw','logout','me'].indexOf(q.action)<0)bad_('Anda harus mengganti password terlebih dahulu.');
 if(ROLES.indexOf(u.role)<0)bad_('Role akun tidak dikenal.');
 if(u.role==='REPORTER'&&RP_OK.indexOf(q.action)<0)bad_('Akses ditolak. Akun Pelapor hanya dapat melaporkan masalah yang perlu diperbaiki dan mengisi Data Penonton.');
 if(u.role==='GUIDE_SECURITY'&&['me','logout','changepw','dashboard','ann_list','ats_list','rep_list','rep_get','rep_save','rep_export','photo_save','photo_done','rep_zip','img','profile_photo_save','notif_poll','aud_day','aud_list','aud_range','aud_save'].indexOf(q.action)<0)bad_('Akses ditolak. Guide/Security hanya dapat mengakses Inspeksi & Patroli Studio, Data Penonton, Dashboard, Pengumuman, dan Advance Ticket Sales.');

 if(u.role!=='REPORTER')autoArsip_();
 var lock=null;
 if(/_(save|del)$|^(changepw|usr_reset|usr_unlock|pel_validasi|photo_done|aud_migrate)$/.test(q.action)){lock=LockService.getScriptLock();lock.waitLock(20000);}
 try{return h(u,q);}finally{if(lock)lock.releaseLock();}
}
var RP_OK=['me','logout','changepw','dashboard','ann_list','ats_list','rep_list','rep_get','rep_save','photo_save','photo_done','rep_zip','img','profile_photo_save','gloss_list','notif_poll','aud_day','aud_list','aud_range','aud_save'];
function adm_(u){if(u.role!=='ADMIN')bad_('Akses ditolak. Hanya Admin.');}

/* ===== laporan: validasi ===== */
/* Data Penonton: batas studio dan show per studio (bawaan formulir: 4 studio x 5 show) */
var AUDMAXS=12,AUDMAXR=15;
function audInt_(v){var x=String(v==null?'':v).trim();if(x==='')return '';return /^\d{1,7}$/.test(x)?Number(x):NaN;}
/* Data Penonton: draf (data.draft===true) boleh diubah pemilik kapan saja; setelah Kirim, edit hanya EDIT_MIN menit. Data lama (tanpa flag) = sudah terkirim. */
var AUDREQ=5,AUDSTD=['Studio 1','Studio 2','Studio 3','Studio 4'];
function isDraft_(r){if(!r||r.type!=='audience')return false;try{return JSON.parse(r.data).draft===true;}catch(e){return false;}}
/* lengkap = SHOW1..SHOW5 setiap studio baku (dan studio tambahan yang ada) berisi nama film, tiket terjual, dan penonton masuk. Mengembalikan pesan galat atau ''. */
function audComplete_(d){
 var by={};(d.studios||[]).forEach(function(a){by[a.studio]=a.shows||[];});
 var names=AUDSTD.slice();Object.keys(by).forEach(function(n){if(names.indexOf(n)<0)names.push(n);});
 for(var i=0;i<names.length;i++){var m={};(by[names[i]]||[]).forEach(function(w){m[w.show]=w;});
  for(var j=1;j<=AUDREQ;j++){var w=m['SHOW'+j];
   if(!w||!w.film||w.tiket===''||w.penonton==='')return 'Belum lengkap: '+names[i]+', SHOW'+j+' (nama film, tiket terjual, dan penonton masuk wajib diisi).';}}
 return '';
}
/* ===== DATA PENONTON PER-SHOW (v20.1.0) =====
 * 1 baris sheet AUDIENCE_SHOWS = 1 show. Kunci data aktif = Tanggal|Studio|Show (kolom showKey, hanya terisi untuk status Aktif), unik.
 * Status: Aktif (dipakai laporan harian) | Review (hasil migrasi yang tidak aman dipetakan, menunggu Admin) | Dihapus | Diabaikan.
 * Data lama (REPORTS type 'audience') tidak diubah atau dihapus; migrasi hanya menyalin dan aman dijalankan ulang (sourceKey). */
var AUDSH='AUDIENCE_SHOWS',AUDMXD=93;
function audD_(v){if(Object.prototype.toString.call(v)==='[object Date]')return isNaN(v.getTime())?'':Utilities.formatDate(v,TZ,'yyyy-MM-dd');return String(v==null?'':v).trim();}
function audJ_(v){if(Object.prototype.toString.call(v)==='[object Date]')return isNaN(v.getTime())?'':Utilities.formatDate(v,TZ,'HH:mm');return String(v==null?'':v).trim();}
function audSt_(v){var m=/^studio\s*0*(\d{1,2})$/i.exec(String(v==null?'':v).trim()),n=m?parseInt(m[1],10):0;return (n>=1&&n<=AUDMAXS)?'Studio '+n:'';}
function audSw_(v){var m=/^show\s*0*(\d{1,2})$/i.exec(String(v==null?'':v).trim()),n=m?parseInt(m[1],10):0;return (n>=1&&n<=AUDMAXR)?'SHOW'+n:'';}
function audKey_(d,s,w){return d+'|'+s+'|'+w;}
function audNum_(v){if(v===''||v==null)return '';var n=Number(v);return isFinite(n)?n:'';}
function audDateOk_(v){return /^\d{4}-\d{2}-\d{2}$/.test(v)&&!isNaN(Date.parse(v+'T00:00:00+07:00'));}
function audRow_(r){return {row:r._row,showKey:String(r.showKey||''),date:audD_(r.date),studio:String(r.studio||''),show:String(r.show||''),film:String(r.film||''),jam:audJ_(r.jam),tiket:audNum_(r.tiket),penonton:audNum_(r.penonton),petugas:String(r.petugas||''),ket:String(r.ket||''),catatan:String(r.catatan||''),userId:String(r.userId||''),reporter:String(r.reporter||''),createdAt:Number(r.createdAt)||0,updatedAt:Number(r.updatedAt)||0,editDeadline:Number(r.editDeadline)||0,status:String(r.status||''),source:String(r.source||''),sourceId:String(r.sourceId||''),sourceKey:String(r.sourceKey||''),sourceRef:String(r.sourceRef||''),review:String(r.review||'')};}
function audRows_(){return all_(AUDSH).map(audRow_);}
function audVals_(o){return SH[AUDSH].map(function(k){return o[k]===undefined?'':cellSafe_(o[k]);});}
function audAppend_(list){
 if(!list.length)return;
 var s=ss_().getSheetByName(AUDSH),n=SH[AUDSH].length,start=s.getLastRow()+1,last=start+list.length-1;
 if(last>s.getMaxRows())s.insertRowsAfter(s.getMaxRows(),last-s.getMaxRows()+50);
 s.getRange(start,1,list.length,n).setValues(list.map(audVals_));
}
function audPut_(o){ss_().getSheetByName(AUDSH).getRange(o.row,1,1,SH[AUDSH].length).setValues([audVals_(o)]);}
function audActive_(rows){var m={};rows.forEach(function(r){if(r.status==='Aktif'&&r.showKey)m[r.showKey]=r;});return m;}
function audCan_(u,r,now){return u.role==='ADMIN'||(r.userId===u.userId&&now<=r.editDeadline);}
function audDayCount_(){try{var d={};audRows_().forEach(function(r){if(r.status==='Aktif'&&r.showKey)d[r.date]=1;});return Object.keys(d).length;}catch(e){return 0;}}
function audNo_(s,re){var m=re.exec(s);return m?parseInt(m[1],10):0;}
/* laporan harian satu tanggal: 4 studio x SHOW1..5 baku + studio/show tambahan yang terisi. act = baris Aktif tanggal tsb */
function audDay_(date,act,u,now){
 var by={},sn={};
 AUDSTD.forEach(function(n){sn[n]=AUDREQ;by[n]={};});
 act.forEach(function(r){if(!sn[r.studio])sn[r.studio]=AUDREQ;if(!by[r.studio])by[r.studio]={};by[r.studio][r.show]=r;var k=audNo_(r.show,/(\d+)/);if(k>sn[r.studio])sn[r.studio]=k;});
 var names=Object.keys(sn).sort(function(a,b){return audNo_(a,/(\d+)/)-audNo_(b,/(\d+)/);});
 var T={show:0,filled:0,tiket:0,penonton:0},pt={},cat={};
 var studios=names.map(function(nm){
  var shows=[],f=0,tk=0,pn=0;
  for(var j=1;j<=sn[nm];j++){
   var w='SHOW'+j,r=by[nm][w],o={show:w,status:r?'Terisi':'Belum Terisi'};
   T.show++;
   if(r){f++;T.filled++;tk+=r.tiket||0;pn+=r.penonton||0;
    if(r.petugas||r.reporter)pt[r.petugas||r.reporter]=1;
    if(r.catatan)cat[r.catatan]=1;
    o.film=r.film;o.jam=r.jam;o.tiket=r.tiket;o.penonton=r.penonton;o.petugas=r.petugas||r.reporter;o.ket=r.ket;o.reporter=r.reporter;
    o.updatedAt=r.updatedAt;o.baseAt=r.updatedAt;o.sourceId=r.sourceId;o.over=(r.tiket!==''&&r.penonton!==''&&r.penonton>r.tiket);
    o.editable=u?audCan_(u,r,now):false;
   }else o.editable=true;
   shows.push(o);}
  T.tiket+=tk;T.penonton+=pn;
  return {studio:nm,total:shows.length,filled:f,tiket:tk,penonton:pn,status:f===shows.length?'Terisi':'Belum Terisi',shows:shows};});
 return {date:date,status:T.filled===T.show?'Terisi':'Belum Terisi',progress:{filled:T.filled,total:T.show,pct:T.show?Math.round(T.filled*100/T.show):0},
  total:{tiket:T.tiket,penonton:T.penonton,show:T.filled},petugas:Object.keys(pt).sort(),catatan:Object.keys(cat),studios:studios};
}
function audByDate_(rows){var m={};rows.forEach(function(r){if(r.status==='Aktif'&&r.showKey)(m[r.date]=m[r.date]||[]).push(r);});return m;}
function audSig_(x){return [String(x.film).toLowerCase().replace(/\s+/g,' '),x.jam,x.tiket,x.penonton].join('|');}
/* migrasi REPORTS(type audience) -> AUDIENCE_SHOWS. Sumber tidak diubah. Idempoten lewat sourceKey (dan sourceRef untuk duplikat identik). */
function audMigrate_(u,dry){
 var lock=LockService.getScriptLock();lock.waitLock(30000);
 try{
  var legacy=all_('REPORTS').filter(function(r){return r.type==='audience';}).sort(function(a,b){return (Number(a.createdAt)||0)-(Number(b.createdAt)||0);});
  var rows=audRows_(),seen={},active=audActive_(rows),used={},now=Date.now();
  rows.forEach(function(r){if(r.sourceKey)seen[r.sourceKey]=1;if(r.sourceRef)r.sourceRef.split(',').forEach(function(k){if(k)seen[k]=1;});});
  var st={dryRun:!!dry,laporanSumber:legacy.length,itemSumber:0,dimigrasi:0,digabung:0,review:0,dilewati:0,alasanReview:{},contohReview:[]};
  var fresh=[],dirty={},groups={},order=[];
  function skOf(base){var k=base,i=1;while(used[k]){i++;k=base+'#'+i;}used[k]=1;return k;}
  function toReview(c,reason){
   var o={showKey:'',date:c.dateOk?c.date:c.rawDate,studio:c.stN||c.rawStudio,show:c.swN||c.rawShow,film:c.film,jam:c.jam,tiket:c.tiket,penonton:c.penonton,petugas:c.petugas,ket:c.ket,catatan:c.catatan,userId:c.userId,reporter:c.reporter,createdAt:c.createdAt,updatedAt:c.updatedAt,editDeadline:c.editDeadline,status:'Review',source:'migrasi',sourceId:c.sourceId,sourceKey:c.sk,sourceRef:'',review:reason};
   fresh.push(o);st.review++;var g=reason.split(':')[0].split(' (')[0];st.alasanReview[g]=(st.alasanReview[g]||0)+1;
   if(st.contohReview.length<15)st.contohReview.push({sumber:c.sourceId,tanggal:o.date,studio:o.studio,show:o.show,alasan:reason});
  }
  function addRef(ex,sk){ex.sourceRef=ex.sourceRef?ex.sourceRef+','+sk:sk;if(ex.row)dirty[ex.row]=ex;st.digabung++;}
  legacy.forEach(function(r){
   var d=null;try{d=JSON.parse(r.data);}catch(e){}
   var base={sourceId:String(r.id),userId:String(r.userId||''),reporter:String(r.reporter||''),createdAt:Number(r.createdAt)||0,updatedAt:Number(r.updatedAt)||Number(r.createdAt)||0,editDeadline:Number(r.editDeadline)||0};
   if(!d||typeof d!=='object'){var k0=r.id+'#?';if(seen[k0]){st.dilewati++;return;}st.itemSumber++;used[k0]=1;
    toReview(Object.assign({sk:k0,rawDate:'',rawStudio:'',rawShow:'',film:'',jam:'',tiket:'',penonton:'',petugas:'',ket:'',catatan:''},base),'Data sumber tidak dapat dibaca (JSON rusak)');return;}
   (Array.isArray(d.studios)?d.studios:[]).forEach(function(a){a=a||{};(Array.isArray(a.shows)?a.shows:[]).forEach(function(w){w=w||{};
    var rawS=String(a.studio==null?'':a.studio),rawW=String(w.show==null?'':w.show),sk0=r.id+'#'+rawS+'#'+rawW;
    if(seen[sk0]&&!used[sk0]){st.dilewati++;return;}
    st.itemSumber++;
    var c=Object.assign({sk:skOf(sk0),rawDate:audD_(d.date),rawStudio:rawS,rawShow:rawW,film:s_(w.film,200),jam:audJ_(w.jam),tiket:audNum_(w.tiket),penonton:audNum_(w.penonton),petugas:s_(w.petugas,100),ket:s_(w.ket,500),catatan:s_(d.note,2000)},base);
    c.date=c.rawDate;c.dateOk=audDateOk_(c.date);c.stN=audSt_(rawS);c.swN=audSw_(rawW);
    var why=[];
    if(!c.dateOk)why.push('Tanggal tidak valid ("'+c.rawDate+'")');
    if(!c.stN)why.push('Studio tidak dapat dipetakan ("'+rawS+'")');
    if(!c.swN)why.push('Show tidak dapat dipetakan ("'+rawW+'")');
    var miss=[];if(!c.film)miss.push('nama film');if(c.tiket==='')miss.push('tiket');if(c.penonton==='')miss.push('penonton');
    if(miss.length)why.push('Belum lengkap ('+miss.join(', ')+')');
    if(why.length){toReview(c,why.join('; '));return;}
    c.key=audKey_(c.date,c.stN,c.swN);if(!groups[c.key]){groups[c.key]=[];order.push(c.key);}groups[c.key].push(c);
   });});
  });
  order.forEach(function(key){
   var cs=groups[key],ex=active[key];
   if(ex){cs.forEach(function(c){if(audSig_(c)===audSig_(ex))addRef(ex,c.sk);else toReview(c,'Konflik: tanggal+studio+show sudah terisi dengan data berbeda (oleh '+(ex.petugas||ex.reporter||'-')+')');});return;}
   var sigs={};cs.forEach(function(c){sigs[audSig_(c)]=1;});
   if(Object.keys(sigs).length===1){
    var c0=cs[0],o={showKey:key,date:c0.date,studio:c0.stN,show:c0.swN,film:c0.film,jam:c0.jam,tiket:c0.tiket,penonton:c0.penonton,petugas:c0.petugas,ket:c0.ket,catatan:c0.catatan,userId:c0.userId,reporter:c0.reporter,createdAt:c0.createdAt,updatedAt:c0.updatedAt,editDeadline:c0.editDeadline,status:'Aktif',source:'migrasi',sourceId:c0.sourceId,sourceKey:c0.sk,sourceRef:'',review:''};
    fresh.push(o);active[key]=o;st.dimigrasi++;
    cs.slice(1).forEach(function(c){addRef(o,c.sk);});
   }else{var ids=cs.map(function(c){return c.sourceId;}).join(', ');cs.forEach(function(c){toReview(c,'Konflik: '+cs.length+' laporan lama memuat show yang sama dengan data berbeda ('+ids+')');});}
  });
  if(!dry){audAppend_(fresh);Object.keys(dirty).forEach(function(k){audPut_(dirty[k]);});
   log_(u,'MIGRASI_DATA_PENONTON','audience','','Migrasi: '+st.dimigrasi+' dipindah, '+st.digabung+' digabung, '+st.review+' review, '+st.dilewati+' dilewati');}
  return st;
 }finally{lock.releaseLock();}
}
/* fungsi manual untuk editor Apps Script (Jalankan > migrasiDataPenonton) */
function migrasiDataPenonton(){var r=audMigrate_(SYS_,false);Logger.log(JSON.stringify(r));return r;}
function chkT_(t){if(!PFX[t])bad_('Jenis laporan tidak valid.');}
function dOk_(v){return /^\d{4}-\d{2}-\d{2}$/.test(v)&&!isNaN(Date.parse(v+'T00:00:00+07:00'))&&v>='2020-01-01'&&Date.parse(v+'T00:00:00+07:00')<=Date.now()+36*3600*1000;}
function tOk_(v){return !v||/^([01]\d|2[0-3]):[0-5]\d$/.test(v);}
function num_(v,min,max){if(v===''||v==null)return NaN;var n=Number(v);return (isFinite(n)&&n>=min&&n<=max)?n:NaN;}
function pick_(v,list,msg){v=s_(v,60);if(list.indexOf(v)<0)bad_(msg);return v;}
function clean_(t,x){
 var d={note:s_(x.note,2000),date:s_(x.date,10),time:s_(x.time,5),location:s_(x.location,200)};
 if(!dOk_(d.date)||!tOk_(d.time))bad_('Tanggal atau jam tidak valid.');
 if(t==='work'){
  d.title=s_(x.title,200);d.category=s_(x.category,50);d.status=s_(x.status,30);d.desc=s_(x.desc,4000);
  d.team=(Array.isArray(x.team)?x.team:[]).slice(0,20).map(function(n){return s_(n,100);}).filter(Boolean);
  if(!d.title||!d.location||WCAT.indexOf(d.category)<0||WST.indexOf(d.status)<0)bad_('Lengkapi judul, kategori, lokasi, dan status pekerjaan.');
 }else if(t==='meter'){
  /* laporan baru = 'Gabungan' (listrik + air sekaligus); 'Listrik' / 'Air' hanya untuk mengedit laporan lama */
  d.kind=(x.kind==='Air'||x.kind==='Listrik')?x.kind:'Gabungan';
  d.unit=d.kind==='Air'?'m³':(d.kind==='Listrik'?'kWh':'');
  if(!d.location||!d.time)bad_('Lengkapi lokasi dan jam pencatatan.');
  if(d.kind!=='Listrik'){d.reading=num_(x.reading,0,1e9);if(isNaN(d.reading))bad_('Isi angka meteran air dengan benar.');}
  if(d.kind!=='Air'){d.lwbp=num_(x.lwbp,0,1e9);d.wbp=num_(x.wbp,0,1e9);d.genset=num_(x.genset,0,1e9);
   if([d.lwbp,d.wbp,d.genset].some(isNaN))bad_('Isi angka LWBP, WBP, dan Genset dengan benar.');}
 }else if(t==='issue'){
  d.title=s_(x.title,200);d.desc=s_(x.desc,4000);d.priority=pick_(x.priority,ISP,'Pilih prioritas masalah.');d.status='Menunggu';d.responses=[];
  if(!d.title||!d.location||!d.desc)bad_('Lengkapi judul, lokasi, dan uraian masalah.');
 }else if(t==='ac'){
  var src=Array.isArray(x.rows)?x.rows:[];delete d.location;
  if(src.length>ACMAX)bad_('Maksimal '+ACMAX+' area.');
  /* baris yang benar-benar kosong (tanpa suhu dan tanpa kondisi Studio kosong / Tidak ada show) dilewati, bukan ditolak */
  src=src.filter(function(r){r=r||{};return !!acCond_(r)||String(r.roomTemp==null?'':r.roomTemp).trim()!=='';});
  if(!src.length)bad_('Isi suhu atau pilih Studio kosong / Tidak ada show minimal untuk satu area.');
  /* suhu: hanya angka (titik atau koma desimal); status suhu SELALU dihitung di server dari standar aktif, nilai kiriman klien diabaikan */
  d.rows=src.map(function(r){r=r||{};var cd=acCond_(r);
   if(cd){ /* studio kosong / tidak ada show: suhu dikosongkan, status otomatis APPROVED */
    var oc={location:s_(r.location,100),unit:pick_(r.unit,ACUNIT,'Pilih merk AC (Daikin atau York).'),roomTemp:'',cond:cd,note:s_(r.note,300),err:acErrIn_(r)};
    if(!oc.location)bad_('Setiap area wajib punya nama.');
    oc.tempStatus='APPROVED';return oc;}
   var tv=String(r.roomTemp==null?'':r.roomTemp).trim().replace(',','.');tv=/^-?\d+(\.\d+)?$/.test(tv)?tv:'';
   var o={location:s_(r.location,100),unit:pick_(r.unit,ACUNIT,'Pilih merk AC (Daikin atau York).'),roomTemp:num_(tv,-10,80),note:s_(r.note,300),err:acErrIn_(r)};
   if(!o.location||isNaN(o.roomTemp))bad_('Setiap area wajib punya nama dan suhu AC (angka -10 sampai 80), atau pilih Studio kosong / Tidak ada show.');
   o.tempStatus=acEval_(o.roomTemp);return o;});
  d.tempStd={min:acCfg_().min,max:acCfg_().max};
 }else if(t==='patrol'){
  d.studios=(Array.isArray(x.studios)?x.studios:[]).slice(0,4).map(function(r,i){r=r||{};return {studio:s_(r.studio,40)||('Studio '+(i+1)),status:s_(r.status,80),note:s_(r.note,1000)};});
  while(d.studios.length<4)d.studios.push({studio:'Studio '+(d.studios.length+1),status:'',note:''});
  d.otherInfo=s_(x.otherInfo,3000);
  if(!d.date||!d.time)bad_('Tanggal dan jam inspeksi wajib diisi.');
  if(!d.studios.some(function(r){return r.status||r.note;})&&!d.otherInfo&&!d.note)bad_('Isi hasil inspeksi minimal satu studio atau informasi lainnya.');
 }else if(t==='audience'){
  /* Data Penonton: studios[] -> shows[]; baris show yang benar-benar kosong dilewati; total dihitung di server */
  delete d.location;
  if(!d.date)bad_('Tanggal wajib diisi.');
  var srcS=Array.isArray(x.studios)?x.studios:[],seenS={};
  if(srcS.length>AUDMAXS)bad_('Maksimal '+AUDMAXS+' studio.');
  d.studios=[];d.totalShow=0;d.totalTiket=0;d.totalPenonton=0;
  srcS.forEach(function(st,i){st=st||{};
   var nm=s_(st.studio,40)||('Studio '+(i+1)),srcR=Array.isArray(st.shows)?st.shows:[],rows=[];
   if(seenS[nm.toLowerCase()])bad_('Studio '+nm+' terisi lebih dari satu kali.');seenS[nm.toLowerCase()]=1;
   if(srcR.length>AUDMAXR)bad_('Maksimal '+AUDMAXR+' show per studio.');
   srcR.forEach(function(r,j){r=r||{};
    var o={show:s_(r.show,20)||('SHOW'+(j+1)),film:s_(r.film,200),jam:s_(r.jam,5),tiket:audInt_(r.tiket),penonton:audInt_(r.penonton),petugas:s_(r.petugas,100),ket:s_(r.ket,500)};
    if(o.tiket!==o.tiket||o.penonton!==o.penonton)bad_('Tiket terjual dan penonton masuk harus bilangan bulat 0 atau lebih ('+nm+', '+o.show+').');
    if(!tOk_(o.jam))bad_('Jam show tidak valid ('+nm+', '+o.show+').');
    if(!o.film&&!o.jam&&o.tiket===''&&o.penonton===''&&!o.petugas&&!o.ket)return;
    if(!o.film)bad_('Nama film wajib diisi ('+nm+', '+o.show+').');
    d.totalShow++;if(o.tiket!=='')d.totalTiket+=o.tiket;if(o.penonton!=='')d.totalPenonton+=o.penonton;rows.push(o);});
   if(rows.length)d.studios.push({studio:nm,shows:rows});});
  if(!d.studios.length)bad_('Isi minimal satu show (nama film) pada salah satu studio.');
 }else{
  d.store=s_(x.store,200);d.buyer=s_(x.buyer,100);d.receiver=s_(x.receiver,100);d.purpose=s_(x.purpose,300);d.total=0;
  d.items=(Array.isArray(x.items)?x.items:[]).slice(0,50).map(function(i){var q=num_(i.qty,0,1e7),p=num_(i.price,0,1e10);
   if(!s_(i.name)||!(q>0)||isNaN(p))bad_('Setiap barang wajib punya nama, jumlah > 0, dan harga valid.');
   if(!s_(i.unit,20))bad_('Satuan barang wajib diisi.');
   var tot=Math.round(q*p);d.total+=tot;return {name:s_(i.name,200),category:s_(i.category,60),qty:q,unit:s_(i.unit,20),price:p,total:tot};});
  if(!d.items.length||!d.store||!d.buyer)bad_('Lengkapi toko, pembeli, dan minimal satu barang.');
 }
 return d;
}
function title_(t,d){return (t==='work'||t==='issue')?d.title:t==='meter'?'Meteran '+(d.kind==='Gabungan'?'Listrik & Air':d.kind)+' - '+d.location:t==='ac'?'Monitoring AC - '+d.rows.length+' area: '+d.rows.map(function(a){return a.location;}).join(', ').slice(0,120):t==='patrol'?'Inspeksi & Patroli Studio - '+d.date:t==='audience'?'Data Penonton - '+d.date+' - '+d.totalShow+' show, '+d.totalPenonton+' penonton':d.store+' - '+d.items.map(function(i){return i.name;}).join(', ').slice(0,120);}

/* ===== METERAN: data, selisih, akumulasi =====
   Pemakaian = selisih angka antara dua pencatatan berurutan di LOKASI yang sama, dihitung terpisah untuk
   LWBP, WBP, Genset, dan Air. Laporan lama (Listrik / Air terpisah) dan laporan baru (Gabungan) dihitung bersama.
   Bila angka baru lebih kecil dari sebelumnya (meteran diganti/reset), selisih dikosongkan, dan tidak diakumulasi (ditandai ⚠ di detail laporan dan panel pemeriksaan data).
   Urutan SELALU menurut tanggal + jam meter yang dicatat (bukan waktu input), sehingga laporan susulan (backdated) memperbaiki seluruh riwayat.
   Selisih TIDAK dibagi rata ke tanggal tanpa pencatatan: pemakaian dicatat pada tanggal pencatatan aktual yang lebih baru; tanggal kosong dilaporkan sebagai gaps. */
function rd3_(v){return Math.round(v*1000)/1000;}
/* urutan kronologis pencatatan: tanggal + jam meter; createdAt lalu id hanya pemecah seri bila tanggal dan jam identik */
function mCmp_(a,b){var x=a.date+' '+a.time,y=b.date+' '+b.time;if(x!==y)return x<y?-1:1;
 var c=(a.createdAt||0)-(b.createdAt||0);if(c)return c<0?-1:1;var i=String(a.id),j=String(b.id);return i<j?-1:(i>j?1:0);}
/* seluruh pencatatan meteran yang aktif, urut waktu: e = angka listrik (atau null), w = angka air (atau null) */
function meterRecs_(){
 var out=[];
 all_('REPORTS').forEach(function(r){
  if(r.type!=='meter'||r.status==='Diarsipkan')return;
  var o;try{o=JSON.parse(r.data);}catch(e){return;}
  var loc=s_(o.location,200);if(!loc||!o.date)return;
  var k=o.kind==='Air'?'Air':(o.kind==='Gabungan'?'Gabungan':'Listrik');
  var e=(k!=='Air'&&o.lwbp!=null&&o.wbp!=null&&o.genset!=null&&isFinite(o.lwbp)&&isFinite(o.wbp)&&isFinite(o.genset))?{lwbp:Number(o.lwbp),wbp:Number(o.wbp),genset:Number(o.genset)}:null,
   w=(k!=='Listrik'&&o.reading!=null&&isFinite(o.reading))?Number(o.reading):null;
  if(!e&&w===null)return; /* data lama tanpa angka LWBP/WBP/Genset tidak bisa dihitung */
  out.push({id:r.id,date:String(o.date),time:o.time||'00:00',createdAt:Number(r.createdAt)||0,loc:loc,lk:loc.toLowerCase(),e:e,w:w});
 });
 out.sort(mCmp_);
 var g={};out.forEach(function(x){(g[x.lk]=g[x.lk]||[]).push(x);});out.byLoc=g; /* indeks per lokasi: meterDiff_ hanya memindai lokasi yang sama */
 return out;
}
/* ---- tanggal sebagai bilangan hari (UTC) ---- */
function dayNum_(s){return Math.round(Date.parse(s+'T00:00:00Z')/864e5);}
function dayStr_(n){return new Date(n*864e5).toISOString().slice(0,10);}
/* peristiwa pemakaian per pencatatan aktual: selisih dua pencatatan berurutan (lokasi + jenis meter sama) dicatat pada tanggal pencatatan yang lebih baru.
   Tidak ada angka fiktif untuk tanggal tanpa pencatatan. Angka turun (reset/ganti meter) tidak dihitung dan tidak diakumulasi; pembanding berikutnya = angka baru.
   Hasil: [{d:tanggal 'YYYY-MM-DD', lk:lokasi, m:'lwbp|wbp|genset|air', v:pemakaian}] */
function meterUse_(recs){
 var ev=[],last={};
 function put_(x,m,cur){var k=x.lk+'|'+m;if(last[k]!==undefined&&cur>=last[k])ev.push({d:x.date,lk:x.lk,m:m,v:rd3_(cur-last[k])});last[k]=cur;}
 recs.forEach(function(x){
  if(x.e)['lwbp','wbp','genset'].forEach(function(m){put_(x,m,x.e[m]);});
  if(x.w!==null)put_(x,'air',x.w);
 });
 return ev;
}
/* tanggal tanpa pencatatan per lokasi dan jenis (listrik / air): dari pencatatan pertama sampai kemarin (WIB) atau pencatatan terakhir */
function meterGaps_(recs){
 var yd=dayNum_(Utilities.formatDate(new Date(),TZ,'yyyy-MM-dd'))-1,by={},out=[];
 recs.forEach(function(x){var o=by[x.lk]=by[x.lk]||{loc:x.loc,e:{},w:{}};o.loc=x.loc;if(x.e)o.e[x.date]=1;if(x.w!==null)o.w[x.date]=1;});
 Object.keys(by).sort().forEach(function(k){var o=by[k];
  [['listrik','e'],['air','w']].forEach(function(t){var ds=Object.keys(o[t[1]]).sort();if(!ds.length)return;
   var end=Math.max(dayNum_(ds[ds.length-1]),yd),miss=[];
   for(var n=dayNum_(ds[0])+1;n<=end;n++){var d=dayStr_(n);if(!o[t[1]][d])miss.push(d);}
   if(miss.length)out.push({loc:o.loc,m:t[0],n:miss.length,dates:miss.slice(0,31)});});});
 return out;
}
/* mengisi d.prevs/d.diffs (listrik) dan d.prev/d.diff (air) terhadap pencatatan sebelumnya di lokasi yang sama */
function meterDiff_(d,selfId,recsAll){
 /* Pemakaian = angka meteran saat ini - angka pencatatan aktual sebelumnya pada lokasi dan jenis meter yang sama.
    Genset tidak dijumlahkan ke LWBP/WBP. "Sebelumnya" menurut mCmp_ (tanggal + jam meter), bukan waktu input,
    sehingga laporan susulan dibandingkan dengan pembacaan yang benar-benar terjadi sebelumnya. */
 var lk=String(d.location||'').trim().toLowerCase(),pe=null,pw=null,all=recsAll||meterRecs_(),
     me={date:String(d.date||''),time:String(d.time||'00:00'),createdAt:Number.MAX_SAFE_INTEGER,id:String(selfId||'')};
 if(all.byLoc)all=all.byLoc[lk]||[];
 all.forEach(function(x){if(selfId&&x.id===selfId)me.createdAt=x.createdAt;});
 all.forEach(function(x){
  if(x.lk!==lk||(selfId&&x.id===selfId)||mCmp_(x,me)>=0)return;
  if(x.e&&(!pe||mCmp_(x,pe)>0))pe=x;
  if(x.w!==null&&(!pw||mCmp_(x,pw)>0))pw=x;
 });
 function delta(previous,current){
  if(previous===null||previous===undefined||current===null||current===undefined||
     current===''||previous===''||!isFinite(Number(previous))||!isFinite(Number(current)))return null;
  previous=Number(previous);current=Number(current);
  /* Jika meter turun (reset/penggantian meter), jangan tampilkan pemakaian negatif. */
  if(current<previous)return null;
  return rd3_(current-previous);
 }
 if(d.kind!=='Air'){
  d.prevs=pe?{lwbp:pe.e.lwbp,wbp:pe.e.wbp,genset:pe.e.genset}:null;
  d.diffs={
   lwbp:delta(pe&&pe.e?pe.e.lwbp:null,d.lwbp),
   wbp:delta(pe&&pe.e?pe.e.wbp:null,d.wbp),
   genset:delta(pe&&pe.e?pe.e.genset:null,d.genset)
  };
  /* Total listrik = selisih LWBP + selisih WBP. Genset tetap terpisah. */
  d.diffs.listrik=(d.diffs.lwbp!==null&&d.diffs.wbp!==null)
      ?rd3_(d.diffs.lwbp+d.diffs.wbp):null;
 }
 if(d.kind!=='Listrik'){
  d.prev=pw?pw.w:null;
  d.diff=delta(pw?pw.w:null,d.reading);
 }
}
function meterSig_(d){return JSON.stringify([d.prevs,d.diffs,d.prev,d.diff]);}
/* segarkan Selisih laporan di lokasi yang sama yang tanggalnya >= fromKey ('YYYY-MM-DD HH:mm'), karena selisihnya bergantung pada laporan
   sebelumnya (laporan susulan / diedit / diarsipkan). Hanya laporan yang selisihnya benar-benar berubah yang ditulis ulang (data + sheet METER_*),
   biasanya laporan itu sendiri dan satu laporan sesudahnya, jadi tetap cepat walau riwayat panjang. */
function refreshMeterMirrors_(loc,fromKey){
 try{
  var lk=String(loc||'').trim().toLowerCase(),recs=meterRecs_();
  all_('REPORTS').forEach(function(r){
   if(r.type!=='meter'||r.status==='Diarsipkan')return;
   var d;try{d=JSON.parse(r.data);}catch(e){return;}
   if(String(d.location||'').trim().toLowerCase()!==lk)return;
   if(fromKey&&(String(d.date)+' '+String(d.time||'00:00'))<fromKey)return;
   var was=meterSig_(d);meterDiff_(d,r.id,recs);
   if(was===meterSig_(d))return;
   upd_('REPORTS',r._row,{data:JSON.stringify(d)});mirrorSafe_(r,d);
  });
 }catch(e){Logger.log('Penyegaran selisih meteran gagal: '+e);}
}
/* Draf Data Penonton belum dikirim: semua akun kecuali Pelapor boleh mengisi/mengubah milik siapa pun; Pelapor hanya miliknya sendiri. Admin selalu boleh. */
function canDraftEdit_(u,r){return u.role==='ADMIN'||r.userId===u.userId||(u.role!=='REPORTER'&&isDraft_(r));}
function editInfo_(u,r,now){
 if(u.role==='ADMIN')return {editable:true,own:r.userId===u.userId};
 if(isDraft_(r))return canDraftEdit_(u,r)?{editable:true,draft:true,own:r.userId===u.userId}:{editable:false,own:false};
 var dl=Number(r.editDeadline);
 return (r.userId===u.userId&&now<=dl)?{editable:true,remainMs:dl-now}:{editable:false,own:r.userId===u.userId};
}
function out_(u,r,now,mrecs){var o=editInfo_(u,r,now),d=JSON.parse(r.data);if(r.type==='ac')acApply_(d);if(r.type==='meter')meterDiff_(d,r.id,mrecs);return {id:r.id,type:r.type,title:r.title,reporter:r.reporter,createdAt:Number(r.createdAt),updatedAt:Number(r.updatedAt),data:d,editable:o.editable&&r.type!=='audience',remainMs:o.remainMs,own:o.own,draft:isDraft_(r)};}
/* Nomor laporan unik per kategori dan tahun: PREFIX-YYYY-0001.
   Counter diturunkan dari ID tersimpan di database, dipanggil dalam ScriptLock pada rep_save.
   ID lama tidak dimigrasikan atau diubah; nomor baru hanya berlaku untuk laporan baru. */
function nextId_(now,type){
 var year=Utilities.formatDate(now,TZ,'yyyy'),prefix=REPORT_PREFIX[type]||'TR',pre=prefix+'-'+year+'-',mx=0;
 all_('REPORTS').forEach(function(r){
  if(r.type!==type)return;
  var id=String(r.id||'');
  if(id.indexOf(pre)!==0)return;
  var tail=id.slice(pre.length),n=/^\d+$/.test(tail)?parseInt(tail,10):0;
  if(n>mx)mx=n;
 });
 return pre+('0000'+(mx+1)).slice(-4);
}
/* Nomor laporan film terpisah dari kategori laporan operasional. ID FR lama tetap valid untuk edit/hapus. */
function nextFilmReportId_(now){
 var year=Utilities.formatDate(now,TZ,'yyyy'),pre=FILM_REPORT_PREFIX+'-'+year+'-',mx=0;
 all_('FILM_REPORTS').forEach(function(r){var id=String(r.id||'');if(id.indexOf(pre)!==0)return;var tail=id.slice(pre.length),n=/^\d+$/.test(tail)?parseInt(tail,10):0;if(n>mx)mx=n;});
 return pre+('0000'+(mx+1)).slice(-4);
}
function repFilter_(q,u,all){
 var s=s_(q.q,100).toLowerCase(),f=q.from?Date.parse(q.from+'T00:00:00+07:00'):0,to=q.to?Date.parse(q.to+'T23:59:59+07:00'):9e15;
 if(isNaN(f))f=0;if(isNaN(to))to=9e15;
 return (all||all_('REPORTS')).filter(function(r){
  if(r.type!==q.type||r.status==='Diarsipkan'||Number(r.createdAt)<f||Number(r.createdAt)>to)return false;
  if(u&&u.role==='REPORTER'&&r.type!=='audience'&&r.userId!==u.userId)return false; /* Data Penonton terlihat oleh semua pengguna */
  if(s&&(r.id+' '+r.title+' '+r.reporter+' '+r.data).toLowerCase().indexOf(s)<0)return false;
  if(q.status&&JSON.parse(r.data).status!==q.status)return false;return true;}).sort(function(a,b){return b.createdAt-a.createdAt;});
}
function canPhoto_(u,rep){
 if(!rep||rep.status==='Diarsipkan')bad_('Laporan tidak ditemukan.');
 if(u.role==='ADMIN')return;
 if(rep.type==='issue'&&u.role==='USER')return; /* Teknisi boleh melampirkan foto penyelesaian */
 if(rep.type==='audience'&&canDraftEdit_(u,rep)&&isDraft_(rep))return;
 if(rep.userId!==u.userId)bad_('Anda hanya dapat menambah foto pada laporan milik sendiri.');
 if(!isDraft_(rep)&&Date.now()>Number(rep.editDeadline)+EDIT_MIN*60000)bad_('Batas waktu unggah foto telah berakhir. Hubungi Admin.');
}

function notifBuild_(u,q){
  var all=null;function reps(){if(all===null)all=all_('REPORTS');return all;}
  var one=function(fn){try{return fn();}catch(e){Logger.log('notif_poll: '+e);return null;}};
  var pq=function(t){return {type:t,q:'',from:'',to:'',status:'',page:1};};
  return {
   ann:one(function(){return H.ann_list(u,q);}),
   issue:u.role==='GUIDE_SECURITY'?null:one(function(){return repPage_(u,pq('issue'),reps());}),
   work:(u.role==='REPORTER'||u.role==='GUIDE_SECURITY')?null:one(function(){return repPage_(u,pq('work'),reps());}),
   ats:u.role==='ADMIN'?null:one(function(){return H.ats_list(u,q);})
  };
}
/* isi daftar laporan (dipakai rep_list dan notif_poll). all = baris REPORTS yang sudah dibaca (opsional). */
function repPage_(u,q,all){
  var pg=Math.max(1,Number(q.page)||1),now=Date.now(),rows=repFilter_(q,u,all);
  var mr=q.type==='meter'?meterRecs_():null;
  return {total:rows.length,page:pg,size:PAGE,items:rows.slice((pg-1)*PAGE,pg*PAGE).map(function(r){return out_(u,r,now,mr);})};
}
/* ===== handler ===== */
var H={
 login:function(q){
  /* Username dicocokkan persis (case-sensitive). Salah password dihitung per akun di database; setelah MAX_FAIL kali,
     akun dikunci LOCK_MS (semua role, termasuk Admin). Kunci habis sendiri, atau Admin dapat membukanya lebih awal. */
  var un=s_(q.username,60),pw=String(q.password||''),u=un&&find_(un);
  if(!u||u.status!=='Aktif')bad_('Username atau password salah.');
  var LOCKMSG='Akun dikunci sementara karena terlalu banyak kesalahan password. Coba lagi dalam 15 menit atau hubungi Admin.';
  if(isLocked_(u))bad_(LOCKMSG);
  if(hash_(pw,u.salt)!==u.hash){
   /* baca ulang di dalam kunci: dua tebakan bersamaan tidak boleh dihitung sekali */
   var lk=LockService.getScriptLock(),n;lk.waitLock(10000);
   try{var fresh=find_(un)||u,base=(Number(fresh.lockedAt)>0&&!isLocked_(fresh))?0:(Number(fresh.failCount)||0);
    n=base+1;
    if(n>=MAX_FAIL){upd_('USERS',u._row,{failCount:n,lockedAt:Date.now()});log_(u,'AKUN_TERBLOKIR','USERS',u.userId,'Dikunci '+(LOCK_MS/60000)+' menit setelah '+n+' kali salah password');}
    else upd_('USERS',u._row,{failCount:n,lockedAt:''});
   }finally{lk.releaseLock();}
   if(n>=MAX_FAIL)bad_(LOCKMSG);
   bad_('Username atau password salah. Sisa percobaan: '+(MAX_FAIL-n)+'.');
  }
  if(Number(u.failCount)>0||Number(u.lockedAt)>0)upd_('USERS',u._row,{failCount:0,lockedAt:''});
  var t=mk_(u);try{CacheService.getScriptCache().put('ls_'+u.userId,String(Date.now()),21600);}catch(e){}
  log_(u,'LOGIN','USERS',u.userId,'Login berhasil');return {token:t,user:pub_(u)};
 },
 logout:function(){return true;},
 me:function(u){var o=pub_(u);o.token=mk_(u);return o;},
 changepw:function(u,q){
  if(hash_(String(q.old||''),u.salt)!==u.hash)bad_('Password lama salah.');chkPw_(q.pw);
  var salt=Utilities.getUuid(),nh=hash_(q.pw,salt);upd_('USERS',u._row,{salt:salt,hash:nh,mustChange:'N',updatedAt:Date.now()});
  log_(u,'GANTI_PASSWORD','USERS',u.userId,'Pengguna mengganti password');return {token:mk_({userId:u.userId,hash:nh})};
 },
 /* FOTO PROFIL (milik sendiri). Urutan aman: validasi -> simpan file baru -> perbarui database -> verifikasi -> baru hapus foto lama.
    Bila langkah mana pun sebelum penghapusan gagal, file baru dibuang dan foto lama tetap dipakai. */
 profile_photo_save:function(u,q){
  ensureUserCols_();u=session_(q.token);
  var b;try{b=Utilities.base64Decode(String(q.b64||'').replace(/^data:[^,]*,/,''));}catch(e0){bad_('Data foto tidak valid.');}
  var mime=imgKind_(b);if(!mime||b.length>MAXB)bad_('Foto harus JPG/PNG/WebP dan maksimal 4 MB.');
  var old=String(u.photoFileId||''),f=null,fol=catFolder_('avatar'),nid='';
  try{
   f=fol.createFile(Utilities.newBlob(b,mime,'AVATAR-'+u.userId+'-'+Date.now()+'.'+EXT[mime]));nid=f.getId();
   upd_('USERS',u._row,{photoFileId:nid,updatedAt:Date.now()});
   var chk=all_('USERS').filter(function(x){return x.userId===u.userId;})[0];
   if(!chk||String(chk.photoFileId)!==nid)throw new Error('verifikasi database gagal');
  }catch(e){
   Logger.log('Ganti foto profil gagal: '+e);
   if(old)try{upd_('USERS',u._row,{photoFileId:old});}catch(e1){} /* pastikan referensi lama utuh */
   if(f)try{f.setTrashed(true);}catch(e2){}
   bad_('Gagal menyimpan foto profil. Foto lama dipertahankan.');
  }
  var removed=old&&old!==nid?delOldAvatar_(u,old,fol):false;
  log_(u,'GANTI_FOTO_PROFIL','USERS',u.userId,removed?'Foto lama dihapus':'Foto profil diperbarui');
  return {photo:true,removedOld:removed};
 },
 /* STANDAR SUHU AC: dibaca semua non-Pelapor, diubah HANYA oleh Admin (dicek di sini, bukan di frontend) */
 ac_cfg_get:function(u){var c=acCfg_();return {min:c.min,max:c.max,def:{min:ACMIN0,max:ACMAX0}};},
 ac_cfg_save:function(u,q){adm_(u);
  var cv=function(v){v=String(v==null?'':v).trim().replace(',','.');return /^-?\d+(\.\d+)?$/.test(v)?Math.round(Number(v)*100)/100:NaN;},mn=cv(q.min),mx=cv(q.max);
  if(isNaN(mn)||isNaN(mx)||mn<-10||mx>80||mx<-10||mn>80)bad_('Isi batas suhu bawah dan atas dengan angka -10 sampai 80.');
  if(mn>mx)bad_('Batas bawah tidak boleh lebih besar dari batas atas.');
  P_().setProperties({AC_TMIN:String(mn),AC_TMAX:String(mx)});AC_C_=null;
  log_(u,'UBAH_STANDAR_SUHU_AC','SETTINGS','','Standar suhu AC: '+mn+' - '+mx+' C');
  return {min:mn,max:mx};},
 team:function(){return all_('USERS').filter(function(x){return x.status==='Aktif'&&x.role!=='REPORTER';}).map(function(x){return x.fullName;});},
 dashboard:function(u){
  var now=new Date(),day=Utilities.formatDate(now,TZ,'yyyyMMdd'),rep=u.role==='REPORTER',
   rp=all_('REPORTS').filter(function(r){return r.status!=='Diarsipkan'&&(!rep||r.type==='audience'||(r.type==='issue'&&r.userId===u.userId));});
  var c=function(t){return rp.filter(function(r){return r.type===t&&!isDraft_(r);}).length;}; /* draf Data Penonton belum dihitung */
  var isOpen=rp.filter(function(r){return r.type==='issue'&&JSON.parse(r.data).status!=='Selesai';}).length;
  var films=atsLive_().filter(function(f){return ['Upcoming ATS','ATS Open','Coming Soon','Coming to Theater'].indexOf(f.status)>=0;});
  return {work:c('work'),meter:c('meter'),purchase:c('purchase'),ac:c('ac'),audience:audDayCount_(),issue:c('issue'),issueOpen:isOpen,
   today:rp.filter(function(r){return Utilities.formatDate(new Date(Number(r.createdAt)),TZ,'yyyyMMdd')===day;}).length,
   films:films.length,
   announcements:all_('ANNOUNCEMENTS').filter(function(a){return a.status==='Aktif';}).sort(function(a,b){return b.publishedAt-a.publishedAt;}).slice(0,5).map(annOut_),
   recent:u.role==='ADMIN'?rp.sort(function(a,b){return b.createdAt-a.createdAt;}).slice(0,8).map(function(r){return out_(u,r,now.getTime());}):[]};
 },
 ann_list:function(u){return all_('ANNOUNCEMENTS').filter(function(a){return u.role==='ADMIN'||a.status==='Aktif';}).sort(function(a,b){return b.publishedAt-a.publishedAt;}).map(annOut_);},
 /* NOTIFIKASI GABUNGAN: satu panggilan menggantikan ann_list + rep_list(issue) + rep_list(work) + ats_list.
   Tiap bagian memakai logika dan hak akses aksi aslinya; bagian yang tidak diizinkan / gagal = null (sama seperti klien lama).
   REPORTS dibaca sekali untuk kedua daftar laporan. */
 notif_poll:function(u,q){
  /* cache singkat per pengguna+role; versi cache berubah pada setiap penulisan (npBump_). Dibatasi 60 detik. */
  var cache=null,ck='';
  try{cache=CacheService.getScriptCache();ck='np|'+(cache.get('npv')||'0')+'|'+u.userId+'|'+u.role;var hit=cache.get(ck);if(hit)return JSON.parse(hit);}catch(e0){cache=null;}
  var out=notifBuild_(u,q);
  try{
   var soon=false;all_('ATS_FILMS').forEach(function(f){var t=Number(f.deleteAt);if(t>0&&t<Date.now()+120000)soon=true;}); /* film yang akan terhapus otomatis: jangan di-cache */
   var str=JSON.stringify(out);if(cache&&!soon&&str.length<90000)cache.put(ck,str,60);
  }catch(e1){Logger.log('cache notif: '+e1);}
  return out;
 },
 ann_save:function(u,q){adm_(u);
  var o={title:s_(q.title,200),body:s_(q.body,4000),priority:['Informasi','Penting','Mendesak'].indexOf(q.priority)>=0?q.priority:'Informasi',updatedAt:Date.now()};
  if(!o.title||!o.body)bad_('Judul dan isi wajib diisi.');
  if(q.id){var a=all_('ANNOUNCEMENTS').filter(function(x){return x.id===q.id;})[0];if(!a)bad_('Pengumuman tidak ditemukan.');upd_('ANNOUNCEMENTS',a._row,o);log_(u,'EDIT_PENGUMUMAN','ANNOUNCEMENTS',q.id,o.title);return q.id;}
  o.id=uid_('A');o.publishedAt=Date.now();o.status='Aktif';o.createdBy=u.fullName;add_('ANNOUNCEMENTS',o);log_(u,'BUAT_PENGUMUMAN','ANNOUNCEMENTS',o.id,o.title);return o.id;
 },
 ann_del:function(u,q){adm_(u);var a=all_('ANNOUNCEMENTS').filter(function(x){return x.id===q.id;})[0];if(!a)bad_('Pengumuman tidak ditemukan.');
  var st=a.status==='Aktif'?'Diarsipkan':'Aktif';upd_('ANNOUNCEMENTS',a._row,{status:st,updatedAt:Date.now()});log_(u,'ARSIP_PENGUMUMAN','ANNOUNCEMENTS',q.id,'Status: '+st);return st;},
 ats_list:function(){return atsLive_().sort(function(a,b){return (Number(a.atsAt)||9e15)-(Number(b.atsAt)||9e15);}).map(function(f){
  return {id:f.id,title:f.title,genre:f.genre,duration:f.duration,rating:f.rating,atsAt:Number(f.atsAt)||0,showAt:Number(f.showAt)||0,deleteAt:Number(f.deleteAt)||0,status:f.status,poster:f.posterFileId,posterUrl:(PUBLIC_PHOTOS&&f.posterFileId)?purl_(f.posterFileId,700):'',note:f.note};});},
 ats_save:function(u,q){adm_(u);
  var o={title:s_(q.title,200),genre:s_(q.genre,100),duration:s_(q.duration,30),rating:s_(q.rating,20),atsAt:Number(q.atsAt)||'',showAt:Number(q.showAt)||'',status:FST.indexOf(q.status)>=0?q.status:'Coming Soon',note:s_(q.note,1000),updatedAt:Date.now()};
  if(!o.title)bad_('Judul film wajib diisi.');
  var da=Number(q.deleteAt)||'';if(da!==''&&da<Date.now()+60000)bad_('Waktu hapus otomatis harus di masa depan (minimal 1 menit dari sekarang).');o.deleteAt=da;
  if(q.poster&&q.poster.b64){var pb=Utilities.base64Decode(String(q.poster.b64)),pm=imgKind_(pb);
   if(!pm||pb.length>MAXB)bad_('Poster harus JPG/PNG/WebP dan maks 4 MB.');
   var pf=catFolder_('poster').createFile(Utilities.newBlob(pb,pm,'POSTER-'+Date.now()+'.'+EXT[pm]));shareFile_(pf);o.posterFileId=pf.getId();}
  if(q.id){var f=all_('ATS_FILMS').filter(function(x){return x.id===q.id;})[0];if(!f)bad_('Film tidak ditemukan.');upd_('ATS_FILMS',f._row,o);log_(u,'EDIT_ATS','ATS_FILMS',q.id,o.title);return q.id;}
  o.id=uid_('F');add_('ATS_FILMS',o);log_(u,'BUAT_ATS','ATS_FILMS',o.id,o.title);return o.id;
 },
 ats_del:function(u,q){adm_(u);var f=all_('ATS_FILMS').filter(function(x){return x.id===q.id;})[0];if(!f)bad_('Film tidak ditemukan.');
  ss_().getSheetByName('ATS_FILMS').deleteRow(f._row);log_(u,'HAPUS_ATS','ATS_FILMS',q.id,f.title);return true;},
 /* ---- Glossarium (baca: semua pengguna; tambah/edit/hapus: Admin) ---- */
 gloss_list:function(u){return all_('GLOSSARY').map(function(g){return {id:g.id,term:g.term,definition:g.definition,category:g.category,source:g.source,updatedAt:Number(g.updatedAt)||0};}).sort(function(a,b){return String(a.term).localeCompare(String(b.term),'id',{sensitivity:'base'});});},
 gloss_save:function(u,q){adm_(u);
  var o={term:s_(q.term,120),definition:s_(q.definition,2000),category:s_(q.category,60),source:s_(q.source,120),updatedBy:u.fullName,updatedAt:Date.now()};
  if(!o.term||!o.definition)bad_('Istilah dan definisi wajib diisi.');
  var key=o.term.toLowerCase();
  if(all_('GLOSSARY').some(function(g){return g.id!==q.id&&String(g.term).toLowerCase()===key;}))bad_('Istilah sudah ada di glossarium.');
  if(q.id){var g=all_('GLOSSARY').filter(function(x){return x.id===q.id;})[0];if(!g)bad_('Istilah tidak ditemukan.');upd_('GLOSSARY',g._row,o);log_(u,'EDIT_GLOSSARIUM','GLOSSARY',q.id,o.term);return q.id;}
  o.id=uid_('G');add_('GLOSSARY',o);log_(u,'BUAT_GLOSSARIUM','GLOSSARY',o.id,o.term);return o.id;
 },
 gloss_del:function(u,q){adm_(u);var g=all_('GLOSSARY').filter(function(x){return x.id===q.id;})[0];if(!g)bad_('Istilah tidak ditemukan.');
  ss_().getSheetByName('GLOSSARY').deleteRow(g._row);log_(u,'HAPUS_GLOSSARIUM','GLOSSARY',q.id,g.term);return true;},
 /* ---- Penyimpanan data multi-format (lihat/unduh: Admin & Teknisi; tambah/edit/hapus: Admin) ---- */
 store_list:function(u){return all_('STORAGE').map(function(r){return {id:r.id,title:r.title,category:r.category,description:r.description,fileName:r.fileName,ext:r.ext,size:Number(r.size)||0,version:Number(r.version)||1,updatedBy:r.updatedBy,updatedAt:Number(r.updatedAt)||0};}).sort(function(a,b){return b.updatedAt-a.updatedAt;});},
 store_get:function(u,q){var r=all_('STORAGE').filter(function(x){return x.id===q.id;})[0];if(!r)bad_('Berkas tidak ditemukan.');
  var b=DriveApp.getFileById(r.fileId).getBlob();return {name:r.fileName,mime:r.mime,b64:Utilities.base64Encode(b.getBytes())};},
 store_save:function(u,q){adm_(u);
  var o={title:s_(q.title,150),category:s_(q.category,60)||'Lainnya',description:s_(q.description,1000),updatedBy:u.fullName,updatedAt:Date.now()};
  if(!o.title)bad_('Judul wajib diisi.');
  var x=storeFile_(q);
  if(q.id){var r=all_('STORAGE').filter(function(y){return y.id===q.id;})[0];if(!r)bad_('Data tidak ditemukan.');
   var old=null;
   if(x){var nf=catFolder_('store').createFile(Utilities.newBlob(x.bytes,x.mime,r.id+'-'+x.name));
    o.fileName=x.name;o.ext=x.ext;o.mime=x.mime;o.size=x.bytes.length;o.fileId=nf.getId();o.version=(Number(r.version)||1)+1;old=r.fileId;}
   upd_('STORAGE',r._row,o);
   if(old)killFile_(old);                   /* berkas lama dibuang hanya setelah berkas baru + database berhasil */
   log_(u,'EDIT_PENYIMPANAN','STORAGE',r.id,o.title+(x?' (berkas diganti)':''));return r.id;}
  if(!x)bad_('Pilih berkas yang akan disimpan.');
  o.id=uid_('S');
  var nf2=catFolder_('store').createFile(Utilities.newBlob(x.bytes,x.mime,o.id+'-'+x.name));   /* folder TIDAK dibagikan publik */
  o.fileName=x.name;o.ext=x.ext;o.mime=x.mime;o.size=x.bytes.length;o.fileId=nf2.getId();o.version=1;o.createdBy=u.fullName;o.createdAt=Date.now();
  try{add_('STORAGE',o);}catch(e){killFile_(nf2.getId());throw e;}
  log_(u,'BUAT_PENYIMPANAN','STORAGE',o.id,o.title+' ('+x.name+')');return o.id;
 },
 store_del:function(u,q){adm_(u);var r=all_('STORAGE').filter(function(x){return x.id===q.id;})[0];if(!r)bad_('Data tidak ditemukan.');
  ss_().getSheetByName('STORAGE').deleteRow(r._row);if(r.fileId)killFile_(r.fileId);log_(u,'HAPUS_PENYIMPANAN','STORAGE',q.id,r.title);return true;},
 filmrep_list:function(u,q){
  if(u.role==='REPORTER')bad_('Akses ditolak.');
  var rows=all_('FILM_REPORTS'),term=s_(q.q,200).toLowerCase();
  rows=rows.filter(function(r){return !term||[r.title,r.status,r.format,r.kdm,r.outlet,r.reportDate].join(' ').toLowerCase().indexOf(term)>=0;});
  var order={'BELUM TAYANG':0,'SEDANG TAYANG':1,'SUDAH TAYANG':2};
  rows.sort(function(a,b){var x=order[String(a.status).toUpperCase()],y=order[String(b.status).toUpperCase()];return (x==null?9:x)-(y==null?9:y)||String(b.reportDate||'').localeCompare(String(a.reportDate||''))||(Number(b.createdAt)||0)-(Number(a.createdAt)||0)||String(a.title).localeCompare(String(b.title),'id',{sensitivity:'base'});});
  return rows.map(function(r){return {id:r.id,title:r.title,status:r.status,format:r.format,kdm:r.kdm,reportDate:r.reportDate,outlet:r.outlet,capacity:r.capacity,createdBy:r.createdBy,signature:r.signature?(String(r.signature).indexOf('http')===0?r.signature:purl_(r.signature,500)):'',createdAt:Number(r.createdAt)||0,updatedAt:Number(r.updatedAt)||0};});
 },
 filmrep_save:function(u,q){
  if(u.role==='REPORTER')bad_('Akses ditolak.');
  var title=s_(q.title,200),status=pick_(q.status,['BELUM TAYANG','SEDANG TAYANG','SUDAH TAYANG'],'Pilih status film.'),format=pick_(q.format,['2D','3D','ATMOS','3D-ATMOS','2D-ATMOS'],'Pilih format film.'),kdm=s_(q.kdm,300),date=s_(q.reportDate,10),outlet=s_(q.outlet,160),capacity=s_(q.capacity,2000),createdBy=s_(q.createdBy,120)||u.fullName,signature='',now=Date.now();
  if(!title||!outlet||!/^\d{4}-\d{2}-\d{2}$/.test(date)||isNaN(Date.parse(date+'T00:00:00+07:00')))bad_('Judul, outlet, dan tanggal laporan wajib diisi dengan benar.');
  var old=q.id?all_('FILM_REPORTS').filter(function(x){return x.id===q.id;})[0]:null;if(q.id&&!old)bad_('Data film tidak ditemukan.');
  var sigIn=String(q.signature||'');if(/^data:image\/(png|jpeg|webp);base64,/i.test(sigIn)){var sb;try{sb=Utilities.base64Decode(sigIn.replace(/^data:[^,]*,/,'') );}catch(se){bad_('Data tanda tangan tidak valid.');}var sm=imgKind_(sb);if(!sm||sb.length>MAXB)bad_('Tanda tangan harus JPG/PNG/WebP maksimal 4 MB.');var sf=catFolder_('store').createFile(Utilities.newBlob(sb,sm,'SIGNATURE-FILM-'+Date.now()+'.'+EXT[sm]));shareFile_(sf);signature=sf.getId();}else if(old)signature=String(old.signature||'');
  var o={title:title,status:status,format:format,kdm:kdm,reportDate:date,outlet:outlet,capacity:capacity,createdBy:createdBy,signature:signature,updatedAt:now};
  if(q.id){upd_('FILM_REPORTS',old._row,o);log_(u,'EDIT_LAPORAN_FILM','FILM_REPORTS',q.id,title);return q.id;}
  o.id=nextFilmReportId_(new Date(now));o.createdAt=now;add_('FILM_REPORTS',o);log_(u,'BUAT_LAPORAN_FILM','FILM_REPORTS',o.id,title);return o.id;
 },
 filmrep_del:function(u,q){adm_(u);var r=all_('FILM_REPORTS').filter(function(x){return x.id===q.id;})[0];if(!r)bad_('Data film tidak ditemukan.');ss_().getSheetByName('FILM_REPORTS').deleteRow(r._row);log_(u,'HAPUS_LAPORAN_FILM','FILM_REPORTS',q.id,r.title);return true;},
 rep_list:function(u,q){
  chkT_(q.type);if(u.role==='REPORTER'&&q.type!=='issue'&&q.type!=='audience')bad_('Akses ditolak.');if(u.role==='GUIDE_SECURITY'&&q.type!=='patrol'&&q.type!=='audience')bad_('Akses ditolak.');
  return repPage_(u,q,null);
 },
 /* ekspor Excel/PDF: semua laporan sesuai filter + info foto/folder/ZIP (Pelapor tidak boleh mengekspor) */
 rep_export:function(u,q){
  chkT_(q.type);if(u.role==='REPORTER')bad_('Akses ditolak.');if(u.role==='GUIDE_SECURITY'&&q.type!=='patrol'&&q.type!=='audience')bad_('Akses ditolak.');var now=Date.now(),rows=repFilter_(q,u).slice(0,MAXEXP),by={};
  var mr=q.type==='meter'?meterRecs_():null;
  all_('PHOTOS').forEach(function(p){if(p.fileId)(by[p.reportId]=by[p.reportId]||[]).push(p);});
  return {total:rows.length,items:rows.map(function(r){var o=out_(u,r,now,mr),i=phInfo_(by[r.id]||[],r),w=function(id){return id?purl_(id,1200):'';};
   o.ph={n:i.n,first:w(i.first),before:w(i.before),after:w(i.after),folder:i.folder?folderUrl_(i.folder):'',zip:i.zip?zipUrl_(i.zip):''};return o;})};
 },
 /* METERAN: angka sebelumnya + akumulasi pemakaian untuk satu lokasi.
    q.location, q.until ('YYYY-MM-DD HH:mm'; tanpa until = semua pencatatan), q.inclusive (sertakan pencatatan tepat di until),
    q.exclude (ID laporan yang sedang diedit). Selalu mengembalikan daftar lokasi yang pernah dipakai. */
 meter_prev:function(u,q){
  if(u.role==='REPORTER')bad_('Akses ditolak.');
  var lk=s_(q.location,200).toLowerCase(),until=s_(q.until,20),inc=!!q.inclusive,ex=s_(q.exclude,80),all=meterRecs_(),lm={};
  all.forEach(function(x){lm[x.lk]=x.loc;});
  var locations=Object.keys(lm).map(function(k){return lm[k];}).sort(),
   mine=all.filter(function(x){
    if(!lk||x.lk!==lk||(ex&&x.id===ex))return false;
    if(!until)return true;
    var k=x.date+' '+x.time;return inc?k<=until:k<until;}),
   le=null,lw=null,last=null,cum={lwbp:0,wbp:0,genset:0,air:0};
  mine.forEach(function(x){if(x.e)le=x;if(x.w!==null)lw=x;last=x;});
  meterUse_(mine).forEach(function(e){cum[e.m]+=e.v;});
  Object.keys(cum).forEach(function(k){cum[k]=rd3_(cum[k]);});
  /* pencatatan berikutnya (jika ada) yang selisihnya ikut dihitung ulang bila laporan ini disimpan susulan */
  var nx=(until&&lk)?all.filter(function(x){return x.lk===lk&&!(ex&&x.id===ex)&&(x.date+' '+x.time)>=until;})[0]:null;
  return {locations:locations,count:mine.length,since:mine.length?mine[0].date:'',
   next:nx?{date:nx.date,time:nx.time,elec:nx.e?{lwbp:nx.e.lwbp,wbp:nx.e.wbp,genset:nx.e.genset}:null,water:nx.w!==null?{reading:nx.w}:null}:null,
   last:last?{date:last.date,time:last.time}:null,
   elec:le?{lwbp:le.e.lwbp,wbp:le.e.wbp,genset:le.e.genset,date:le.date,time:le.time}:null,
   water:lw?{reading:lw.w,date:lw.date,time:lw.time}:null,
   cum:cum,has:{e:mine.some(function(x){return !!x.e;}),w:mine.some(function(x){return x.w!==null;})}};
 },
 /* METERAN: statistik untuk grafik. q.location (kosong = semua lokasi), q.from/q.to ('YYYY-MM-DD'), q.group ('day'|'month').
    cum = akumulasi seluruh riwayat; per = pemakaian dalam periode; series = pemakaian per hari/bulan dalam periode. */
 meter_stats:function(u,q){
  if(u.role==='REPORTER')bad_('Akses ditolak.');
  var dr=function(v){v=s_(v,10);return /^\d{4}-\d{2}-\d{2}$/.test(v)?v:'';},
   lk=s_(q.location,200).toLowerCase(),fromQ=dr(q.from),toQ=dr(q.to),from=fromQ||'0000-01-01',to=toQ||'9999-12-31',
   gr=q.group==='month'?'month':'day',all=meterRecs_(),recs=lk?all.filter(function(x){return x.lk===lk;}):all,lm={},
   M=['lwbp','wbp','genset','air'],cum={},per={},bk={},lt={};
  all.forEach(function(x){lm[x.lk]=x.loc;});
  M.forEach(function(m){cum[m]=0;per[m]=0;});
  var ev=meterUse_(recs);
  ev.forEach(function(e){
   cum[e.m]+=e.v;if(e.d<from||e.d>to)return;
   per[e.m]+=e.v;var b=gr==='month'?e.d.slice(0,7):e.d;
   if(!bk[b]){bk[b]={k:b};M.forEach(function(m){bk[b][m]=0;});}
   bk[b][e.m]+=e.v;});
  M.forEach(function(m){cum[m]=rd3_(cum[m]);per[m]=rd3_(per[m]);});
  recs.forEach(function(x){var o=lt[x.lk]=lt[x.lk]||{loc:x.loc,elec:null,water:null};o.loc=x.loc;
   if(x.e)o.elec={lwbp:x.e.lwbp,wbp:x.e.wbp,genset:x.e.genset,date:x.date,time:x.time};
   if(x.w!==null)o.water={reading:x.w,date:x.date,time:x.time};});
  return {locations:Object.keys(lm).map(function(k){return lm[k];}).sort(),
   latest:Object.keys(lt).sort().map(function(k){return lt[k];}),
   cum:cum,per:per,
   series:Object.keys(bk).sort().map(function(k){var o=bk[k];M.forEach(function(m){o[m]=rd3_(o[m]);});return o;}),
   since:recs.length?recs[0].date:'',
   count:{total:recs.length,period:recs.filter(function(x){return x.date>=from&&x.date<=to;}).length},
   gaps:meterGaps_(recs),
   from:fromQ,to:toQ,group:gr};
 },
 /* ---- Data Penonton per-show (v20.1.0). Semua peran yang berhak mengakses modul boleh membaca dan mengisi. ---- */
 aud_day:function(u,q){
  var date=audD_(q.date);if(!audDateOk_(date))bad_('Tanggal tidak valid.');
  var rows=audRows_(),act=audByDate_(rows)[date]||[],o=audDay_(date,act,u,Date.now());
  if(u.role==='ADMIN')o.reviewCount=rows.filter(function(r){return r.status==='Review';}).length;
  return o;
 },
 aud_list:function(u,q){
  var pg=Math.max(1,Number(q.page)||1),now=Date.now(),rows=audRows_(),by=audByDate_(rows),today=Utilities.formatDate(new Date(now),TZ,'yyyy-MM-dd'),
   from=audD_(q.from),to=audD_(q.to),s=s_(q.q,100).toLowerCase();
  if(!audDateOk_(from))from='';if(!audDateOk_(to))to='';
  if(!by[today]&&(!from||today>=from)&&(!to||today<=to)&&!s)by[today]=[];
  var dates=Object.keys(by).filter(function(d){
   if(from&&d<from)return false;if(to&&d>to)return false;
   if(s&&(d+' '+by[d].map(function(r){return r.film+' '+r.petugas+' '+r.reporter+' '+r.studio;}).join(' ')).toLowerCase().indexOf(s)<0)return false;return true;}).sort().reverse();
  var items=dates.slice((pg-1)*PAGE,pg*PAGE).map(function(d){var o=audDay_(d,by[d],u,now);delete o.studios;return o;});
  return {total:dates.length,page:pg,size:PAGE,items:items,reviewCount:u.role==='ADMIN'?rows.filter(function(r){return r.status==='Review';}).length:0};
 },
 /* rentang tanggal lengkap (untuk unduhan Excel/PDF), maksimal AUDMXD hari */
 aud_range:function(u,q){
  var from=audD_(q.from),to=audD_(q.to);if(!audDateOk_(from)||!audDateOk_(to)||from>to)bad_('Rentang tanggal tidak valid.');
  var now=Date.now(),by=audByDate_(audRows_());
  var dates=Object.keys(by).filter(function(d){return d>=from&&d<=to;}).sort();
  if(dates.length>AUDMXD)bad_('Maksimal '+AUDMXD+' hari per unduhan. Persempit rentang tanggal.');
  return {from:from,to:to,days:dates.map(function(d){return audDay_(d,by[d],u,now);})};
 },
 /* simpan satu atau banyak show pada satu tanggal. Semua divalidasi dulu; bila ada yang gagal, tidak ada yang tersimpan. */
 aud_save:function(u,q){
  var date=audD_(q.date);if(!dOk_(date))bad_('Tanggal tidak valid.');
  var src=Array.isArray(q.shows)?q.shows:[];if(!src.length)bad_('Tidak ada show yang diisi.');if(src.length>AUDMAXS*AUDMAXR)bad_('Terlalu banyak show dalam satu kiriman.');
  var now=Date.now(),seen={},list=[];
  src.forEach(function(r){r=r||{};
   var st=audSt_(r.studio),sw=audSw_(r.show);if(!st||!sw)bad_('Studio atau show tidak valid.');
   var key=audKey_(date,st,sw),lbl=st+' '+sw;if(seen[key])bad_(lbl+' terisi lebih dari satu kali.');seen[key]=1;
   var o={key:key,studio:st,show:sw,film:s_(r.film,200),jam:s_(r.jam,5),tiket:audInt_(r.tiket),penonton:audInt_(r.penonton),petugas:s_(r.petugas,100)||u.fullName,ket:s_(r.ket,500),baseAt:Number(r.baseAt)||0,lbl:lbl};
   if(o.tiket!==o.tiket||o.penonton!==o.penonton)bad_('Tiket terjual dan penonton masuk harus bilangan bulat 0 atau lebih ('+lbl+').');
   if(!tOk_(o.jam))bad_('Jam show tidak valid ('+lbl+').');
   if(!o.film||o.tiket===''||o.penonton==='')bad_('Nama film, tiket terjual, dan penonton masuk wajib diisi ('+lbl+').');
   list.push(o);});
  var act=audActive_(audRows_()),add=[],upd=[],same=0;
  list.forEach(function(o){var ex=act[o.key];
   if(!ex){add.push({showKey:o.key,date:date,studio:o.studio,show:o.show,film:o.film,jam:o.jam,tiket:o.tiket,penonton:o.penonton,petugas:o.petugas,ket:o.ket,catatan:'',userId:u.userId,reporter:u.fullName,createdAt:now,updatedAt:now,editDeadline:now+EDIT_MIN*60000,status:'Aktif',source:'',sourceId:'',sourceKey:'',sourceRef:'',review:''});return;}
   var who=ex.petugas||ex.reporter||'pengguna lain';
   if(!o.baseAt)bad_(o.lbl+' baru saja diisi oleh '+who+'. Muat ulang tampilan lalu periksa datanya.');
   if(ex.updatedAt>o.baseAt)bad_(o.lbl+' baru saja diubah oleh '+who+'. Muat ulang tampilan lalu periksa datanya.');
   if(!audCan_(u,ex,now))bad_(o.lbl+' sudah diisi oleh '+who+'. Hanya Admin, atau pengisi pertama dalam '+EDIT_MIN+' menit, yang dapat mengubahnya.');
   if(o.film===ex.film&&o.jam===ex.jam&&o.tiket===ex.tiket&&o.penonton===ex.penonton&&o.petugas===ex.petugas&&o.ket===ex.ket){same++;return;}
   ex.film=o.film;ex.jam=o.jam;ex.tiket=o.tiket;ex.penonton=o.penonton;ex.petugas=o.petugas;ex.ket=o.ket;ex.updatedAt=now;upd.push(ex);});
  audAppend_(add);upd.forEach(audPut_);
  if(add.length||upd.length)log_(u,'ISI_DATA_PENONTON','audience',date,'Data Penonton '+date+': '+add.length+' show baru, '+upd.length+' diubah');
  var out=H.aud_day(u,{date:date});out.saved={created:add.length,updated:upd.length,unchanged:same};return out;
 },
 aud_del:function(u,q){
  adm_(u);var date=audD_(q.date),key=audKey_(date,audSt_(q.studio),audSw_(q.show)),ex=audActive_(audRows_())[key];
  if(!ex)bad_('Data show tidak ditemukan.');
  ex.status='Dihapus';ex.showKey='';ex.review='Dihapus oleh '+u.fullName+' pada '+fmtTs_(Date.now());ex.updatedAt=Date.now();audPut_(ex);
  log_(u,'HAPUS_DATA_PENONTON','audience',date,'Hapus '+ex.studio+' '+ex.show+' ('+date+')');
  return H.aud_day(u,{date:date});
 },
 aud_migrate:function(u,q){adm_(u);return audMigrate_(u,q.dryRun===true);},
 aud_review_list:function(u){
  adm_(u);return audRows_().filter(function(r){return r.status==='Review';}).sort(function(a,b){return a.date<b.date?1:(a.date>b.date?-1:0);}).slice(0,500).map(function(r){delete r.row;return r;});
 },
 /* Admin menyelesaikan satu item review: koreksi isian bila perlu, lalu jadikan data aktif (menggantikan data aktif lama pada kunci yang sama) */
 aud_review_save:function(u,q){
  adm_(u);var rows=audRows_(),r=rows.filter(function(x){return x.sourceKey===String(q.sourceKey)&&x.status==='Review';})[0];
  if(!r)bad_('Item review tidak ditemukan atau sudah diselesaikan.');
  var f=q.fix||{},date=audD_(f.date||r.date),st=audSt_(f.studio||r.studio),sw=audSw_(f.show||r.show);
  if(!audDateOk_(date))bad_('Tanggal tidak valid.');if(!st||!sw)bad_('Studio atau show tidak valid.');
  var film=s_(f.film!=null?f.film:r.film,200),jam=s_(f.jam!=null?f.jam:r.jam,5),tk=audInt_(f.tiket!=null?f.tiket:r.tiket),pn=audInt_(f.penonton!=null?f.penonton:r.penonton);
  if(tk!==tk||pn!==pn)bad_('Tiket terjual dan penonton masuk harus bilangan bulat 0 atau lebih.');
  if(!tOk_(jam))bad_('Jam show tidak valid.');if(!film||tk===''||pn==='')bad_('Nama film, tiket terjual, dan penonton masuk wajib diisi.');
  var key=audKey_(date,st,sw),ex=audActive_(rows)[key],now=Date.now();
  if(ex){ex.status='Dihapus';ex.showKey='';ex.review='Digantikan oleh '+r.sourceKey+' ('+u.fullName+')';ex.updatedAt=now;audPut_(ex);}
  r.date=date;r.studio=st;r.show=sw;r.film=film;r.jam=jam;r.tiket=tk;r.penonton=pn;
  if(f.petugas!=null)r.petugas=s_(f.petugas,100);if(f.ket!=null)r.ket=s_(f.ket,500);
  r.showKey=key;r.status='Aktif';r.review='Diselesaikan Admin '+u.fullName+': '+r.review;r.updatedAt=now;audPut_(r);
  log_(u,'REVIEW_DATA_PENONTON','audience',key,'Review diterapkan: '+r.sourceKey);
  return {ok:true,key:key,replaced:!!ex};
 },
 aud_review_del:function(u,q){
  adm_(u);var r=audRows_().filter(function(x){return x.sourceKey===String(q.sourceKey)&&x.status==='Review';})[0];
  if(!r)bad_('Item review tidak ditemukan atau sudah diselesaikan.');
  r.status='Diabaikan';r.review='Diabaikan Admin '+u.fullName+': '+r.review;r.updatedAt=Date.now();audPut_(r);
  log_(u,'ABAIKAN_REVIEW_PENONTON','audience',r.sourceKey,'Review diabaikan');return {ok:true};
 },
 rep_get:function(u,q){
  var r=all_('REPORTS').filter(function(x){return x.id===q.id&&x.status!=='Diarsipkan';})[0];if(!r)bad_('Laporan tidak ditemukan.');
  if(u.role==='REPORTER'&&r.type!=='audience'&&(r.type!=='issue'||r.userId!==u.userId))bad_('Akses ditolak.');if(u.role==='GUIDE_SECURITY'&&r.type!=='patrol'&&r.type!=='audience')bad_('Akses ditolak.');
  var o=out_(u,r,Date.now());o.photos=all_('PHOTOS').filter(function(p){return p.reportId===r.id;}).sort(function(a,b){return Number(a.uploadedAt)-Number(b.uploadedAt);}).map(function(p){return {id:p.photoId,kind:p.kind,uploadedAt:Number(p.uploadedAt)||0,capturedAt:Number(p.capturedAt)||Number(p.uploadedAt)||0,url:PUBLIC_PHOTOS?purl_(p.fileId,1200):''};});
  o.folder=r.folderId?folderUrl_(r.folderId):'';return o;
 },
 rep_save:function(u,q){
  var t=q.type;chkT_(t);if(t==='audience')bad_('Data Penonton kini diisi per show (Tanggal, Studio, Show). Muat ulang aplikasi (Ctrl+F5) untuk memakai formulir terbaru.');if(u.role==='REPORTER'&&t!=='issue'&&t!=='audience')bad_('Akses ditolak. Pelapor hanya dapat membuat laporan Perlu Diperbaiki dan Data Penonton.');if(u.role==='GUIDE_SECURITY'&&t!=='patrol'&&t!=='audience')bad_('Akses ditolak. Guide/Security hanya dapat membuat Inspeksi & Patroli Studio dan Data Penonton.');
  var ts=Date.now(),x0=q.data||{};
  if(t==='issue'){x0.date=Utilities.formatDate(new Date(ts),TZ,'yyyy-MM-dd');x0.time=Utilities.formatDate(new Date(ts),TZ,'HH:mm');}
  var d=clean_(t,x0),id,rep,sub=false,oldD=null;
  if(t==='audience'){sub=q.submit===true;}
  if(q.id){
   rep=all_('REPORTS').filter(function(r){return r.id===q.id;})[0];
   if(!rep||rep.status==='Diarsipkan'||rep.type!==t)bad_('Laporan tidak ditemukan.');
   if(t==='issue'){var od=JSON.parse(rep.data);d.status=od.status;d.responses=od.responses||[];d.date=od.date;d.time=od.time;}
   if(t==='audience'){try{oldD=JSON.parse(rep.data);}catch(e9){oldD={};}}
   /* dua orang membuka draf yang sama: yang menyimpan belakangan tidak boleh menimpa diam-diam */
   if(t==='audience'&&q.baseAt&&Number(rep.updatedAt)>Number(q.baseAt))bad_('Data ini baru saja diubah oleh pengguna lain. Tutup lalu buka lagi laporan ini agar data terbaru tampil.');
   if(u.role!=='ADMIN'){
    if(!(rep.userId===u.userId||(t==='audience'&&canDraftEdit_(u,rep))))bad_('Anda hanya dapat mengedit laporan milik sendiri.');

    if(!isDraft_(rep)&&ts>Number(rep.editDeadline))bad_('Batas waktu edit telah berakhir. Hubungi Admin jika diperlukan koreksi.');}
   if(t==='audience'){
    if(oldD.draft===true){d.draft=!sub;if(sub){d.submittedAt=ts;}}
    else if(oldD.draft===false){d.draft=false;d.submittedAt=oldD.submittedAt;} /* sudah terkirim: tetap terkirim, batas waktu tidak diperpanjang */
    else{delete d.draft;} /* data lama sebelum fitur draf */
    if(d.draft===false){var ce=audComplete_(d);if(ce)bad_(ce);}}
   if(t==='meter')meterDiff_(d,rep.id);
   var upf={title:title_(t,d),data:JSON.stringify(d),updatedAt:ts};
   if(t==='audience'&&oldD&&oldD.draft===true&&sub)upf.editDeadline=ts+EDIT_MIN*60000; /* jam edit 30 menit mulai saat Kirim */
   upd_('REPORTS',rep._row,upf);id=rep.id;
   mirrorSafe_({id:id,type:t,reporter:rep.reporter,createdAt:rep.createdAt,updatedAt:ts,status:'Aktif'},d);
   log_(u,t==='audience'&&oldD&&oldD.draft===true&&sub?'KIRIM_LAPORAN':'EDIT_LAPORAN',t,id,t==='audience'&&oldD&&oldD.draft===true&&sub?'Data Penonton dikirim (submit)':'Laporan diedit oleh '+u.role);
   if(t==='meter'){try{var od2=JSON.parse(rep.data);refreshMeterMirrors_(od2.location,od2.date+' '+(od2.time||'00:00'));}catch(e5){}refreshMeterMirrors_(d.location,d.date+' '+(d.time||'00:00'));}
  }else{
   if(t==='audience'){d.draft=!sub;if(sub){d.submittedAt=ts;var ce2=audComplete_(d);if(ce2)bad_(ce2);}}
   var rid=s_(q.reqId,80),dup=rid&&all_('REPORTS').filter(function(r){return r.reqId===rid;})[0];
   if(dup)return {id:dup.id,photoFailed:0,duplicate:true};
   id=nextId_(new Date(ts),t);if(t==='meter')meterDiff_(d,null);
   add_('REPORTS',{id:id,type:t,title:title_(t,d),data:JSON.stringify(d),userId:u.userId,reporter:u.fullName,createdAt:ts,updatedAt:ts,editDeadline:ts+EDIT_MIN*60000,status:'Aktif',reqId:rid});
   log_(u,'BUAT_LAPORAN',t,id,title_(t,d));
   mirrorSafe_({id:id,type:t,reporter:u.fullName,createdAt:ts,updatedAt:ts,status:'Aktif'},d);
   if(t==='meter')refreshMeterMirrors_(d.location,d.date+' '+(d.time||'00:00'));
  }
  return {id:id,photoFailed:0,submitted:t==='audience'?d.draft===false:undefined,draft:t==='audience'?d.draft===true:undefined};
 },
 /* unggah foto bertahap: maks PHBATCH foto per panggilan, total MAXPH per laporan */
 photo_save:function(u,q){
  var rep=all_('REPORTS').filter(function(r){return r.id===q.id;})[0];canPhoto_(u,rep);
  var r=addPhotos_(rep.id,q.photos);
  if(r.added){try{syncFoto_(rep.id);}catch(e){Logger.log(e);}try{syncPhotoRows_(rep);}catch(e3){Logger.log('Sheet foto: '+e3);}}
  return r;
 },
 /* panggilan penutup setelah semua foto terunggah: sinkron sheet + ZIP otomatis bila foto > ZIP_AT */
 photo_done:function(u,q){
  var rep=all_('REPORTS').filter(function(r){return r.id===q.id;})[0];canPhoto_(u,rep);
  var n=all_('PHOTOS').filter(function(p){return p.reportId===rep.id;}).length,zip=false;
  if(n>ZIP_AT)try{makeZip_(rep.id);zip=true;}catch(e){Logger.log('ZIP gagal: '+e);}
  try{syncFoto_(rep.id);}catch(e2){Logger.log(e2);}
  var fixed=0;try{fixed=syncPhotoRows_(rep);}catch(e4){Logger.log('Pemeriksaan foto: '+e4);}
  log_(u,'FOTO_LAPORAN',rep.type,rep.id,n+' foto'+(zip?', ZIP dibuat':''));
  return {total:n,zip:zip,sheetFixed:fixed};
 },
 rep_zip:function(u,q){
  var r=all_('REPORTS').filter(function(x){return x.id===q.id&&x.status!=='Diarsipkan';})[0];if(!r)bad_('Laporan tidak ditemukan.');
  if(u.role==='REPORTER'&&r.type!=='audience'&&(r.type!=='issue'||r.userId!==u.userId))bad_('Akses ditolak.');if(u.role==='GUIDE_SECURITY'&&r.type!=='patrol'&&r.type!=='audience')bad_('Akses ditolak.');
  var z=makeZip_(r.id);try{syncFoto_(r.id);}catch(e){}return z;
 },
 /* tanggapan Teknisi/Admin pada laporan Perlu Diperbaiki; dapat mengubah status sampai Selesai */
 resp_save:function(u,q){
  if(u.role==='REPORTER')bad_('Akses ditolak. Hanya Teknisi atau Admin yang dapat menanggapi.');
  var r=all_('REPORTS').filter(function(x){return x.id===q.id&&x.type==='issue'&&x.status!=='Diarsipkan';})[0];if(!r)bad_('Laporan tidak ditemukan.');
  var text=s_(q.text,2000),st=s_(q.status,30);
  if(IST.indexOf(st)<1)bad_('Pilih status tanggapan (Ditanggapi, Sedang Dikerjakan, atau Selesai).');
  if(!text)bad_('Tulis tanggapan atau tindakan yang dilakukan.');
  var d=JSON.parse(r.data),ts=Date.now();d.responses=d.responses||[];
  d.responses.push({by:u.fullName,byId:u.userId,at:ts,text:text,status:st});d.status=st;
  if(st==='Selesai')d.doneAt=ts;else delete d.doneAt;
  upd_('REPORTS',r._row,{data:JSON.stringify(d),updatedAt:ts});
  mirrorSafe_({id:r.id,type:'issue',reporter:r.reporter,createdAt:r.createdAt,updatedAt:ts,status:r.status},d);
  log_(u,'TANGGAPAN_ISSUE','issue',r.id,st+': '+text.slice(0,120));
  return {id:r.id,status:st};
 },
 rep_del:function(u,q){adm_(u);
  var r=all_('REPORTS').filter(function(x){return x.id===q.id;})[0];if(!r)bad_('Laporan tidak ditemukan.');
  upd_('REPORTS',r._row,{status:'Diarsipkan',updatedAt:Date.now()});log_(u,'ARSIP_LAPORAN',r.type,r.id,'Alasan: '+s_(q.reason,200));
  try{mirrorSafe_({id:r.id,type:r.type,reporter:r.reporter,createdAt:r.createdAt,updatedAt:Date.now(),status:'Diarsipkan'},JSON.parse(r.data));}catch(e){}
  if(r.type==='meter'){try{var dd2=JSON.parse(r.data);refreshMeterMirrors_(dd2.location,dd2.date+' '+(dd2.time||'00:00'));}catch(e6){}}
  return true;},
 img:function(u,q){
  var fid;
  if(q.kind==='avatar'){var au=all_('USERS').filter(function(x){return x.userId===q.id;})[0];
   if(!au||(au.userId!==u.userId&&u.role!=='ADMIN'))bad_('Akses ditolak.');fid=au.photoFileId;}
  else if(q.kind==='poster'){var f=all_('ATS_FILMS').filter(function(x){return x.posterFileId===q.id;})[0];fid=f&&f.posterFileId;}
  else{var p=all_('PHOTOS').filter(function(x){return x.photoId===q.id;})[0];fid=p&&p.fileId;}
  if(!fid)bad_('Foto tidak ditemukan.');
  /* foto laporan mengikuti batas peran yang sama dengan rep_get: Pelapor hanya Data Penonton dan miliknya, Guide/Security hanya Patroli dan Data Penonton */
  if((u.role==='REPORTER'||u.role==='GUIDE_SECURITY')&&q.kind!=='poster'&&q.kind!=='avatar'){var pr=all_('REPORTS').filter(function(x){return x.id===p.reportId;})[0];
   if(!pr||(u.role==='REPORTER'&&pr.type!=='audience'&&pr.userId!==u.userId)||(u.role==='GUIDE_SECURITY'&&pr.type!=='patrol'&&pr.type!=='audience'))bad_('Akses ditolak.');}
  var b=DriveApp.getFileById(fid).getBlob();return 'data:'+b.getContentType()+';base64,'+Utilities.base64Encode(b.getBytes());
 },
 usr_list:function(u){adm_(u);return all_('USERS').map(function(x){return {userId:x.userId,fullName:x.fullName,username:x.username,role:x.role,outlet:String(x.outlet||''),status:x.status,locked:isLocked_(x),failCount:Number(x.failCount)||0,createdAt:Number(x.createdAt)};});},
 usr_unlock:function(u,q){adm_(u);var x=all_('USERS').filter(function(y){return y.userId===q.userId;})[0];if(!x)bad_('Pengguna tidak ditemukan.');
  upd_('USERS',x._row,{failCount:0,lockedAt:'',updatedAt:Date.now()});log_(u,'BUKA_BLOKIR','USERS',x.userId,'Blokir dibuka untuk '+x.username);return true;},
 usr_save:function(u,q){adm_(u);
  var role=ROLES.indexOf(q.role)>=0?q.role:'USER',name=s_(q.fullName,100),now=Date.now(),ol=PEL_OUTLET.indexOf(q.outlet)>=0?q.outlet:'';
  if(!name)bad_('Nama lengkap wajib diisi.');
  if(q.userId){var x=all_('USERS').filter(function(y){return y.userId===q.userId;})[0];if(!x)bad_('Pengguna tidak ditemukan.');
   var st=q.status==='Nonaktif'?'Nonaktif':'Aktif';
   if(x.userId===u.userId&&(role!=='ADMIN'||st!=='Aktif'))bad_('Anda tidak dapat menurunkan atau menonaktifkan akun sendiri.');
   upd_('USERS',x._row,{fullName:name,role:role,status:st,outlet:ol,updatedAt:now});
   log_(u,x.role!==role?'UBAH_ROLE':st!==x.status?'STATUS_AKUN':'EDIT_AKUN','USERS',x.userId,x.username+' -> '+role+'/'+st);return x.userId;}
  var un=s_(q.username,30);if(!/^[A-Za-z0-9_.-]{3,30}$/.test(un))bad_('Username 3-30 karakter (huruf, angka, . _ -).');
  if(findCI_(un))bad_('Username sudah dipakai.');chkPw_(q.password);
  var salt=Utilities.getUuid(),id=uid_('U');
  add_('USERS',{userId:id,fullName:name,username:un,role:role,outlet:ol,status:'Aktif',salt:salt,hash:hash_(q.password,salt),mustChange:'Y',createdAt:now,updatedAt:now});
  log_(u,'BUAT_AKUN','USERS',id,un+' ('+role+')');return id;
 },
 usr_reset:function(u,q){adm_(u);chkPw_(q.password);var x=all_('USERS').filter(function(y){return y.userId===q.userId;})[0];if(!x)bad_('Pengguna tidak ditemukan.');
  var salt=Utilities.getUuid();upd_('USERS',x._row,{salt:salt,hash:hash_(q.password,salt),mustChange:'Y',failCount:0,lockedAt:'',updatedAt:Date.now()});
  log_(u,'RESET_PASSWORD','USERS',x.userId,'Reset untuk '+x.username);return true;},
 audit:function(u,q){adm_(u);var pg=Math.max(1,Number(q.page)||1),rows=all_('AUDIT_LOG').sort(function(a,b){return b.at-a.at;});
  return {total:rows.length,page:pg,size:25,items:rows.slice((pg-1)*25,pg*25).map(function(r){return {at:Number(r.at),userName:r.userName,role:r.role,action:r.action,category:r.category,recordId:r.recordId,summary:r.summary};})};}
};
function annOut_(a){return {id:a.id,title:a.title,body:a.body,priority:a.priority,publishedAt:Number(a.publishedAt),status:a.status,createdBy:a.createdBy};}

/* ===== PEMBERSIHAN SHEET YANG TIDAK TERPAKAI =====
   Sheet "terpakai" = semua sheet yang namanya dirujuk kode: database website (SH), sheet kategori (MIR),
   sheet foto (PHS), dan sheet modul DCP & Film Testing (DCP_SH di dcp.gs). Selain itu (mis. Sheet1 bawaan atau sheet lama) dianggap tidak terpakai.
   Urutan aman: 1) periksaSheetTidakTerpakai  2) sembunyikanSheetTidakTerpakai (bisa dibatalkan)  3) hapusSheetTidakTerpakai. */
function sheetPakai_(){
 var o={};Object.keys(SH).forEach(function(n){o[n]=1;});
 Object.keys(MIR).forEach(function(k){o[MIR[k].name]=1;});
 Object.keys(PHS).forEach(function(t){o[PHS[t]]=1;});
 /* sheet modul DCP & Film Testing (dcp.gs) juga dipakai kode; jangan pernah dianggap sampah */
 if(typeof DCP_SH!=='undefined')Object.keys(DCP_SH).forEach(function(n){o[n]=1;});
 o.CAFE_SALES=1;o.CAFE_TARGET=1; /* data lama modul Cafe dibiarkan utuh */
 return o;
}
function sheetTakTerpakai_(){var use=sheetPakai_();return ss_().getSheets().filter(function(s){return !use[s.getName()];});}
function periksaSheetTidakTerpakai(){
 var use=sheetPakai_(),c=sheetTakTerpakai_();
 Logger.log('DIPAKAI (tidak akan disentuh): '+Object.keys(use).join(', '));
 Logger.log(c.length?('TIDAK TERPAKAI ('+c.length+'):\n'+c.map(function(s){return ' - '+s.getName()+' ('+Math.max(0,s.getLastRow())+' baris, '+s.getLastColumn()+' kolom)';}).join('\n')):'Tidak ada sheet yang tidak terpakai.');
 return c.map(function(s){return s.getName();});
}
function sembunyikanSheetTidakTerpakai(){
 var ss=ss_(),c=sheetTakTerpakai_();if(!c.length){Logger.log('Tidak ada sheet yang tidak terpakai.');return;}
 try{ss.setActiveSheet(ss.getSheetByName('REPORTS'));}catch(e){}
 c.forEach(function(s){try{s.hideSheet();}catch(e){Logger.log('Gagal menyembunyikan '+s.getName()+': '+e);}});
 Logger.log('Disembunyikan: '+c.map(function(s){return s.getName();}).join(', ')+'. Untuk menampilkan lagi: menu Tampilkan > Sheet tersembunyi.');
}
function hapusSheetTidakTerpakai(){
 var ss=ss_(),c=sheetTakTerpakai_(),del=[];
 if(!c.length){Logger.log('Tidak ada sheet yang dihapus.');return;}
 if(c.length>=ss.getSheets().length)throw new Error('Semua sheet akan terhapus. Spreadsheet harus menyisakan minimal satu sheet. Dibatalkan.');
 /* salinan penuh dulu; bila gagal, tidak ada yang dihapus */
 var bak=DriveApp.getFileById(ss.getId()).makeCopy('BACKUP sebelum hapus sheet '+Utilities.formatDate(new Date(),TZ,'yyyy-MM-dd HHmm')+' - '+ss.getName());
 Logger.log('BACKUP: '+bak.getUrl());
 try{ss.setActiveSheet(ss.getSheetByName('REPORTS'));}catch(e){}
 c.forEach(function(s){var n=s.getName()+' ('+Math.max(0,s.getLastRow()-1)+' baris)';try{ss.deleteSheet(s);del.push(n);}catch(e){Logger.log('Gagal menghapus '+n+': '+e);}});
 Logger.log(del.length?('Dihapus permanen: '+del.join(', ')):'Tidak ada sheet yang dihapus.');
}

/** Glossarium awal dari 2 dokumen Word (93 istilah). Jalankan seedGlossarium() SEKALI dari editor Apps Script.
 *  Aman dijalankan ulang: istilah yang sudah ada dilewati. Kolom: [istilah, definisi, kategori, sumber]. */
var GLOSS_SEED=[
 ["2D / 3D", "Penanda format film yang dicantumkan pada bagian content type modifiers.", "Operasional & Projectionist", "Roadshow Projectionist"],
 ["2K / 4K", "Resolusi konten film yang dicantumkan dalam panduan membaca file konten.", "Operasional & Projectionist", "Roadshow Projectionist"],
 ["5.1 / 7.1", "Konfigurasi audio surround yang dibahas dalam materi pelatihan.", "Operasional & Projectionist", "Roadshow Projectionist"],
 ["Aspect ratio", "Perbandingan dimensi gambar. Dokumen mencantumkan Scope 1:2,39 dan Flat 1:1,85.", "Operasional & Projectionist", "Roadshow Projectionist"],
 ["Centralized", "Sistem/ruang terpusat untuk perangkat dan pengelolaan materi seperti LMS, CPU Cinead, CPU musik, dan monitor.", "Operasional & Projectionist", "Roadshow Projectionist"],
 ["Cinead Player", "Player yang ditampilkan dalam materi sebagai bagian sinkronisasi waktu dengan TMS Player.", "Operasional & Projectionist", "Roadshow Projectionist"],
 ["Cinema audio processor", "Perangkat untuk memfilter output audio sebelum amplifier agar sesuai standar suara yang diinginkan.", "Operasional & Projectionist", "Roadshow Projectionist"],
 ["Closing", "Tahap penutupan setelah operasional, termasuk mematikan perangkat sesuai urutan.", "Operasional & Projectionist", "Roadshow Projectionist"],
 ["Content type (FTR)", "Kolom informasi pada file konten film; dokumen mencantumkan FTR.", "Operasional & Projectionist", "Roadshow Projectionist"],
 ["Content type modifiers", "Penanda format konten, misalnya 2D atau 3D.", "Operasional & Projectionist", "Roadshow Projectionist"],
 ["Cooling down", "Proses pendinginan proyektor setelah digunakan; dokumen mengarahkan proyektor dimatikan setelah proses ini selesai, sekitar 10 menit.", "Operasional & Projectionist", "Roadshow Projectionist"],
 ["DCI (Digital Cinema Initiatives)", "Standar/inisiatif yang disebut pada uraian server playback cinema digital.", "Operasional & Projectionist", "Roadshow Projectionist"],
 ["DCP (Digital Cinema Package)", "Materi/paket film digital yang ditransfer ke TMS dan server teater. | Paket konten film digital yang disebut dalam proses transfer ke TMS/server.", "Umum (kedua sumber)", "Roadshow Projectionist + Projector & Sound"],
 ["DMD (Digital Micromirror Device)", "Chip berisi jutaan cermin mikro untuk memantulkan gambar; satu cermin mewakili satu piksel menurut dokumen. | Chip elektronik berisi jutaan cermin kecil; dokumen menjelaskan satu cermin mewakili satu piksel.", "Umum (kedua sumber)", "Roadshow Projectionist + Projector & Sound"],
 ["Dolby Atmos", "Teknologi audio cinema yang pada dokumen dijelaskan memberikan channel individual pada speaker, dengan jumlah maksimal 64 channel.", "Operasional & Projectionist", "Roadshow Projectionist"],
 ["Doremi Playback Server", "Perangkat pemutaran cinema digital yang dapat memutar materi DCP; dokumen menyebut kontennya terenkripsi.", "Operasional & Projectionist", "Roadshow Projectionist"],
 ["Exhaust blower", "Perangkat ventilasi yang dinyalakan pada tahap persiapan sistem proyektor.", "Operasional & Projectionist", "Roadshow Projectionist"],
 ["Filter air intake", "Salah satu filter yang tercantum untuk proyektor Christie.", "Operasional & Projectionist", "Roadshow Projectionist"],
 ["Filter Cardcage", "Salah satu filter yang tercantum untuk proyektor Barco.", "Operasional & Projectionist", "Roadshow Projectionist"],
 ["Filter Heat Exchanger", "Salah satu filter yang tercantum untuk proyektor Barco.", "Operasional & Projectionist", "Roadshow Projectionist"],
 ["Filter LAD", "Salah satu filter yang tercantum untuk proyektor Christie.", "Operasional & Projectionist", "Roadshow Projectionist"],
 ["Filter Radiator", "Salah satu filter yang tercantum untuk proyektor Christie.", "Operasional & Projectionist", "Roadshow Projectionist"],
 ["Filter SPG", "Salah satu filter yang tercantum untuk proyektor Barco.", "Operasional & Projectionist", "Roadshow Projectionist"],
 ["Flat", "Jenis aspect ratio yang dicantumkan sebagai 1:1,85.", "Operasional & Projectionist", "Roadshow Projectionist"],
 ["HDD (Hard Disk Drive)", "Penyimpanan server yang perlu dipantau kapasitas dan status slot A/B/C-nya. | Media penyimpanan server; dokumen menyarankan pengecekan status slot A/B/C dan menyebut kapasitas penyimpanan film maksimal 70%.", "Umum (kedua sumber)", "Roadshow Projectionist + Projector & Sound"],
 ["HFR (High Frame Rate)", "Dukungan frame rate tinggi; istilah ini muncul pada daftar kemampuan beberapa perangkat server/IMB.", "Operasional & Projectionist", "Roadshow Projectionist"],
 ["Humidity", "Kelembapan udara; termasuk parameter monitoring ruang proyektor.", "Operasional & Projectionist", "Roadshow Projectionist"],
 ["KDM (Key Delivery Message)", "Pesan lisensi yang mengatur kapan konten film dapat aktif/diputar. | Pesan/lisensi digital untuk membuka konten film pada waktu yang diizinkan.", "Umum (kedua sumber)", "Roadshow Projectionist + Projector & Sound"],
 ["LMS (Library Management System)", "Perangkat/sistem yang dinyalakan dalam prosedur centralized.", "Operasional & Projectionist", "Roadshow Projectionist"],
 ["LSS-100 (Light Sound Sensor)", "Perangkat yang diperiksa melalui log; dokumen menyebut nilai dBc dan luminance/FL. | Perangkat sensor cahaya dan suara; dokumen menyebut pemeriksaan log, nilai dBc, dan luminance.", "Umum (kedua sumber)", "Roadshow Projectionist + Projector & Sound"],
 ["Luminance / FL", "Parameter tingkat cahaya pada Log LSS-100; dokumen menyebut nilai 12–16 fl. | Nilai tingkat cahaya yang dicantumkan dalam pemeriksaan Log LSS; dokumen menyebut rentang 12–16 fl.", "Umum (kedua sumber)", "Roadshow Projectionist + Projector & Sound"],
 ["Management IP", "Alamat untuk mengoperasikan proyektor, menurut dokumen glossary pendamping. | Alamat IP untuk mengoperasikan proyektor.", "Umum (kedua sumber)", "Roadshow Projectionist + Projector & Sound"],
 ["Monitoring preshow", "Pemeriksaan checklist preshow, playlist iklan, serta materi/order yang bermasalah.", "Operasional & Projectionist", "Roadshow Projectionist"],
 ["Movie title", "Judul film yang perlu diperiksa pada informasi file konten.", "Operasional & Projectionist", "Roadshow Projectionist"],
 ["Opening", "Tahap persiapan sebelum operasional bioskop dimulai.", "Operasional & Projectionist", "Roadshow Projectionist"],
 ["Operasional", "Tahap ketika pemutaran film berjalan dan perangkat serta kualitas gambar/suara dimonitor.", "Operasional & Projectionist", "Roadshow Projectionist"],
 ["Playback server", "Server yang memutar konten film; dokumen menyebut kapasitas penyimpanan maksimal 70%.", "Operasional & Projectionist", "Roadshow Projectionist"],
 ["Playlist", "Susunan materi preshow atau film pokok yang dijadwalkan untuk diputar.", "Operasional & Projectionist", "Roadshow Projectionist"],
 ["Porthole", "Kaca/jendela antara ruang proyektor dan studio; dokumen menyebut pembersihan setiap tiga hari.", "Operasional & Projectionist", "Roadshow Projectionist"],
 ["Power amplifier", "Perangkat penguat sinyal audio sebelum diteruskan ke speaker.", "Operasional & Projectionist", "Roadshow Projectionist"],
 ["Preshow", "Materi sebelum film utama, seperti welcome, pengumuman, trailer, dan iklan. | Materi sebelum film utama, meliputi welcome, pengumuman awal, trailer, iklan, dan pengumuman akhir.", "Umum (kedua sumber)", "Roadshow Projectionist + Projector & Sound"],
 ["Preventive maintenance", "Tindakan pencegahan/perawatan terjadwal untuk perangkat, termasuk filter dan perlengkapan theater.", "Operasional & Projectionist", "Roadshow Projectionist"],
 ["Resolution", "Resolusi konten film; dokumen mencantumkan 2K dan 4K.", "Operasional & Projectionist", "Roadshow Projectionist"],
 ["Scope", "Jenis aspect ratio yang dicantumkan sebagai 1:2,39.", "Operasional & Projectionist", "Roadshow Projectionist"],
 ["Show schedule", "Jadwal preshow dan film utama per studio, termasuk waktu mulai/akhir dan judul film.", "Operasional & Projectionist", "Roadshow Projectionist"],
 ["Storage capacity", "Kapasitas penyimpanan server; dokumen menyebut batas maksimal penggunaan 70%.", "Operasional & Projectionist", "Roadshow Projectionist"],
 ["Surround speaker", "Speaker surround di studio yang perlu diperiksa secara berkala.", "Operasional & Projectionist", "Roadshow Projectionist"],
 ["Temperature", "Temperatur proyektor dan ruang proyektor yang dipantau secara berkala.", "Operasional & Projectionist", "Roadshow Projectionist"],
 ["Thermogun", "Alat yang disebut untuk mengukur temperatur ruang proyektor; materi mencantumkan 24°C sebagai acuan.", "Operasional & Projectionist", "Roadshow Projectionist"],
 ["TMS", "Sistem yang digunakan untuk melihat konten film, lisensi/KDM, status transfer, dan jadwal film/preshow.", "Operasional & Projectionist", "Roadshow Projectionist"],
 ["TMS Player", "Player yang ditampilkan dalam materi terkait sinkronisasi jam dengan Cinead Player.", "Operasional & Projectionist", "Roadshow Projectionist"],
 ["Transfer status", "Tampilan untuk memonitor proses transfer konten dan KDM dari DCP ke TMS atau dari TMS ke server teater.", "Operasional & Projectionist", "Roadshow Projectionist"],
 ["Xenon lamp hours", "Jam pemakaian lampu Xenon yang dipantau berkala; dokumen mencantumkan ambang pemesanan berdasarkan lokasi.", "Operasional & Projectionist", "Roadshow Projectionist"],
 ["5.1 / 7.1 / Atmos", "Format/konfigurasi audio yang disebut untuk pemutaran film di bioskop.", "Proyektor, Suara & Pendukung", "Projector & Sound"],
 ["AC (Air Conditioner)", "Perangkat untuk mengatur/menjaga kelembapan udara dan menurunkan temperatur ruangan.", "Proyektor, Suara & Pendukung", "Projector & Sound"],
 ["Amplifier / Power Amplifier", "Perangkat yang menguatkan sinyal suara input lalu meneruskan output ke speaker.", "Proyektor, Suara & Pendukung", "Projector & Sound"],
 ["Audio Crossover", "Perangkat/rangkaian yang membagi rentang frekuensi suara sebelum sinyal diperkuat power amplifier. Dokumen menyebut crossover aktif dan pasif.", "Proyektor, Suara & Pendukung", "Projector & Sound"],
 ["Audio Processor / Cinema Audio Processor", "Perangkat yang memproses atau memfilter output audio agar suara sesuai standar bioskop; mendukung konfigurasi seperti 5.1, 7.1, dan Atmos.", "Proyektor, Suara & Pendukung", "Projector & Sound"],
 ["CC (Cinema Controller)", "Komponen yang disebut berada pada proyektor Barco.", "Proyektor, Suara & Pendukung", "Projector & Sound"],
 ["COGS (Cost of Goods Sales, sesuai dokumen)", "Di dokumen dijelaskan sebagai HPP (Harga Pokok Penjualan).", "Proyektor, Suara & Pendukung", "Projector & Sound"],
 ["Compressor / Kompresor", "Komponen AC yang memompa/mendorong refrigeran ke seluruh sistem pendingin.", "Proyektor, Suara & Pendukung", "Projector & Sound"],
 ["Content IP", "Alamat IP yang digunakan untuk transfer film dan Cinead.", "Proyektor, Suara & Pendukung", "Projector & Sound"],
 ["Crossover 2-way", "Susunan speaker depan layar dengan rentang high dan low.", "Proyektor, Suara & Pendukung", "Projector & Sound"],
 ["Crossover 3-way", "Susunan speaker depan layar dengan rentang high, mid, dan low.", "Proyektor, Suara & Pendukung", "Projector & Sound"],
 ["Crossover 4-way", "Susunan speaker depan layar dengan rentang very high, high, mid, dan low.", "Proyektor, Suara & Pendukung", "Projector & Sound"],
 ["Digital Signage", "Distribusi/penayangan konten dinamis seperti promo, menu F&B, trailer, dan iklan pada layar LED/TV.", "Proyektor, Suara & Pendukung", "Projector & Sound"],
 ["DLP Cinema (Digital Light Processing)", "Teknologi proyektor digital yang menggunakan chip berisi banyak cermin mikro untuk memantulkan gambar dan warna.", "Proyektor, Suara & Pendukung", "Projector & Sound"],
 ["EBITDA (Earnings Before Interest, Taxes, Depreciation, and Amortization)", "Ukuran keuangan yang kepanjangannya dicantumkan dalam dokumen.", "Proyektor, Suara & Pendukung", "Projector & Sound"],
 ["Ekspansi / pipa kapiler", "Komponen AC untuk menurunkan tekanan refrigeran, mengalirkannya ke evaporator, dan mengatur jumlah refrigeran.", "Proyektor, Suara & Pendukung", "Projector & Sound"],
 ["Equalizer", "Salah satu dari empat peralatan utama sound system yang disebut dalam dokumen.", "Proyektor, Suara & Pendukung", "Projector & Sound"],
 ["Evaporator", "Komponen AC yang menyerap panas dari udara sehingga menghasilkan udara dingin.", "Proyektor, Suara & Pendukung", "Projector & Sound"],
 ["EVB", "Pada dokumen, disebut sebagai pengendali suhu, fan, dan sensor pada proyektor Christie.", "Proyektor, Suara & Pendukung", "Projector & Sound"],
 ["Fan Controller", "Komponen pengendali kipas pada proyektor Barco.", "Proyektor, Suara & Pendukung", "Projector & Sound"],
 ["FNB", "Food and Beverage; istilah ini muncul dalam konteks omset FNB pada dokumen.", "Proyektor, Suara & Pendukung", "Projector & Sound"],
 ["Freon / refrigeran", "Media pendinginan AC menurut dokumen; AC chiller disebut menggunakan refrigeran dan air.", "Proyektor, Suara & Pendukung", "Projector & Sound"],
 ["ICP (Integrated Cinema Processor)", "Komponen pada proyektor Barco dan Christie; dokumen menyebut gangguan/block pada ICP dapat membuat sistem tidak masuk ke Cinema Controller.", "Proyektor, Suara & Pendukung", "Projector & Sound"],
 ["IMCB (Integrated Main Control Board)", "Komponen pada proyektor Christie yang mengontrol Xenon dan dowser/penutup sinar Xenon.", "Proyektor, Suara & Pendukung", "Projector & Sound"],
 ["IP (Internet Protocol) address", "Alamat jaringan; pada dokumen dibedakan menjadi Management IP dan Content IP.", "Proyektor, Suara & Pendukung", "Projector & Sound"],
 ["Kondensor", "Komponen AC yang mengatur temperatur refrigeran dan mengubah gas menjadi cairan.", "Proyektor, Suara & Pendukung", "Projector & Sound"],
 ["LEQD board", "Board yang disebut dalam langkah troubleshooting masalah gambar proyektor; dokumen tidak memberikan kepanjangan istilah ini.", "Proyektor, Suara & Pendukung", "Projector & Sound"],
 ["License / lisensi film", "Izin pada server digital yang menentukan film dapat dibuka/diputar dan masa berlakunya.", "Proyektor, Suara & Pendukung", "Projector & Sound"],
 ["Light Engine", "Bagian optik/pemrosesan cahaya proyektor; dokumen mencantumkan pemeriksaan formatter board, DMD, dan satelite saat masalah gambar.", "Proyektor, Suara & Pendukung", "Projector & Sound"],
 ["LPS (Lamp Power Supply)", "Catu daya untuk lampu.", "Proyektor, Suara & Pendukung", "Projector & Sound"],
 ["LVPS (Low Voltage Power Supply)", "Catu daya tegangan rendah; dokumen menyebut komponen ini pada proyektor Christie dan untuk konversi ke 12–24 VDC.", "Proyektor, Suara & Pendukung", "Projector & Sound"],
 ["Marriage (Projector/IMB)", "Proses pairing/pernikahan perangkat yang dijelaskan pada dokumen untuk proyektor Barco dan Christie.", "Proyektor, Suara & Pendukung", "Projector & Sound"],
 ["PIB (Projector Intelligence Board)", "Komponen yang disebut berada pada proyektor Christie.", "Proyektor, Suara & Pendukung", "Projector & Sound"],
 ["QNAP TMS", "Disebut sebagai perangkat penerima/penampung materi preshow; dokumen menyebutnya menggantikan Cinead pada Gen 21.", "Proyektor, Suara & Pendukung", "Projector & Sound"],
 ["Server playback", "Perangkat untuk menyimpan dan memutar konten film digital; dokumen menyebut merek seperti Doremi dan Dolby.", "Proyektor, Suara & Pendukung", "Projector & Sound"],
 ["SMPS (Switch Mode Power Supply)", "Mengubah arus AC menjadi DC; dokumen menyebut lokasinya pada proyektor Barco.", "Proyektor, Suara & Pendukung", "Projector & Sound"],
 ["Subwoofer", "Speaker untuk reproduksi frekuensi rendah; dokumen menyebut tiga titik subwoofer pada konfigurasi Dolby Atmos yang dibahas.", "Proyektor, Suara & Pendukung", "Projector & Sound"],
 ["TPC (Touch Panel Control)", "Panel sentuh untuk mengontrol proyektor Christie.", "Proyektor, Suara & Pendukung", "Projector & Sound"],
 ["UPS (Uninterruptible Power Supply)", "Perangkat catu daya cadangan yang dinyalakan pada urutan persiapan sistem.", "Proyektor, Suara & Pendukung", "Projector & Sound"],
 ["Xenon lamp", "Lampu proyektor Xenon; dokumen mencantumkan contoh tipe/seri Ushio, Christie, dan Osram.", "Proyektor, Suara & Pendukung", "Projector & Sound"]
];
function seedGlossarium(){
 ensureSchema_();
 var have={};all_('GLOSSARY').forEach(function(g){have[String(g.term).toLowerCase()]=1;});
 var n=0,now=Date.now();
 GLOSS_SEED.forEach(function(r){
  if(have[String(r[0]).toLowerCase()])return;
  add_('GLOSSARY',{id:uid_('G'),term:r[0],definition:r[1],category:r[2],source:r[3],updatedBy:'Impor awal (dokumen Word)',updatedAt:now});n++;
 });
 Logger.log('Glossarium: '+n+' istilah ditambahkan, '+(GLOSS_SEED.length-n)+' dilewati (sudah ada).');
}
/* ===== NOTIFIKASI OTOMATIS (OneSignal) =====
   Penerima: Admin & Teknisi (laporan baru / diedit); Pelapor (laporannya ditanggapi).
   Kegagalan notifikasi tidak membatalkan penyimpanan.
   Script Properties yang dibutuhkan: ONESIGNAL_APP_ID, ONESIGNAL_REST_KEY. */

var OS_URL_ = 'https://api.onesignal.com/notifications';
var OS_APP_URL_ = 'https://technicalreport-stcxxi.pages.dev/';
var OS_LABEL_ = {work:'Pekerjaan', meter:'Meteran', purchase:'Pembelian', ac:'Monitoring AC', issue:'Perlu Diperbaiki', patrol:'Inspeksi Studio', audience:'Data Penonton'};

function osSend_(ids, title, body, data) {
  ids = (ids || []).filter(function (x, i, a) { return x && a.indexOf(x) === i; });
  if (!ids.length) return;
  var p = P_(), appId = p.getProperty('ONESIGNAL_APP_ID'), key = p.getProperty('ONESIGNAL_REST_KEY');
  if (!appId || !key) { Logger.log('OneSignal belum dikonfigurasi di Script Properties.'); return; }
  var payload = {
    app_id: appId,
    target_channel: 'push',
    include_aliases: {external_id: ids},
    headings: {en: title},
    contents: {en: body},
    url: OS_APP_URL_,
    data: data || {}
  };
  try {
    var r = UrlFetchApp.fetch(OS_URL_, {
      method: 'post',
      contentType: 'application/json',
      headers: {Authorization: 'Key ' + key},
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    if (r.getResponseCode() >= 300) Logger.log('OneSignal gagal ' + r.getResponseCode() + ': ' + r.getContentText());
  } catch (e) { Logger.log('OneSignal error: ' + e); }
}

/* pengguna aktif dengan role tertentu, kecuali pelaku aksi */
function notifAudience_(roles, exceptId) {
  return all_('USERS').filter(function (x) {
    return x.status === 'Aktif' && roles.indexOf(x.role) >= 0 && String(x.userId) !== String(exceptId);
  }).map(function (x) { return String(x.userId); });
}

function notifRep_(u, act, t, id, title) {
  try {
    var lbl = OS_LABEL_[t] || 'Laporan';
    var head = act === 'create'
      ? (t === 'issue' ? 'Perlu diperbaiki: laporan baru' : 'Laporan ' + lbl + ' baru')
      : 'Laporan ' + lbl + ' diperbarui';
    osSend_(notifAudience_(['ADMIN', 'USER'], u.userId), head, (title || id) + ' · ' + u.fullName + ' · ' + id, {reportId: id, type: t});
  } catch (e) { Logger.log('notif laporan: ' + e); }
}

function notifResp_(u, r, st, text) {
  try {
    if (String(r.userId) === String(u.userId)) return; /* tidak memberi tahu diri sendiri */
    osSend_([String(r.userId)], 'Tanggapan pada laporan Anda', r.id + ' · status: ' + st + ' · ' + String(text || '').slice(0, 120), {reportId: r.id, type: 'issue'});
  } catch (e) { Logger.log('notif tanggapan: ' + e); }
}

/* Pemasangan pembungkus. H didefinisikan di atas; file ini harus dimuat setelah bagian H. */
(function installNotifHooks_() {
  if (typeof H === 'undefined') throw new Error('notifikasi: H belum dimuat.');
  var origRepSave = H.rep_save, origRespSave = H.resp_save;
  H.rep_save = function (u, q) {
    var wasDraft = false;
    try { if (q && q.type === 'audience' && q.id) wasDraft = isDraft_(all_('REPORTS').filter(function (x) { return x.id === q.id; })[0]); } catch (e0) {}
    var out = origRepSave(u, q);
    try {
      if (out && out.id && !out.duplicate && !(q.type === 'audience' && out.draft === true)) {
        var rep = all_('REPORTS').filter(function (x) { return x.id === out.id; })[0];
        notifRep_(u, (q.id && !wasDraft) ? 'edit' : 'create', q.type, out.id, rep ? rep.title : '');
        /* AC error: kirim peringatan terpisah ke Admin & Teknisi, berisi nama area yang error */
        if (q.type === 'ac' && rep) {
          var dd = JSON.parse(rep.data);
          var errs = (dd.rows || []).filter(acErrIn_).map(function (a) { return a.location; });
          if (errs.length) osSend_(notifAudience_(['ADMIN', 'USER'], u.userId),
            'AC error terdeteksi', errs.join(', ') + ' · ' + u.fullName, {reportId: out.id, type: 'ac'});
        }
      }
    } catch (e) { Logger.log('hook rep_save: ' + e); }
    return out;
  };
  H.resp_save = function (u, q) {
    var out = origRespSave(u, q);
    try {
      var rep = all_('REPORTS').filter(function (x) { return x.id === out.id; })[0];
      if (rep) notifResp_(u, rep, out.status, q.text);
    } catch (e) { Logger.log('hook resp_save: ' + e); }
    return out;
  };
})();

/* JALANKAN SEKALI dari editor: tes pengiriman ke semua Admin yang sudah mengaktifkan notifikasi */
function testNotifikasi() {
  osSend_(notifAudience_(['ADMIN'], ''), 'Tes notifikasi', 'Jika Anda menerima ini, pengaturan OneSignal berhasil.', {test: 1});
  Logger.log('Tes dikirim.');
}
function izinUrlFetch() {
     UrlFetchApp.fetch('https://www.googleapis.com/', {muteHttpExceptions: true});
     Logger.log('Izin OK');
   }
function cekUser() {
     var rows = all_('USERS');
     Logger.log('Jumlah user: ' + rows.length);
     rows.forEach(function (x) {
       Logger.log('userId=[' + x.userId + '] role=[' + x.role + '] status=[' + x.status + ']');
     });
   }
