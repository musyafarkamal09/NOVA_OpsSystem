/* ===== MODUL DCP MANAGEMENT & DAILY FILM TESTING (additive) =====
 * Berkas ini berdiri sendiri. Code.gs hanya mendelegasikan aksi yang tidak dikenal ke DCP_H (lihat route_).
 * Data disimpan di sheet sendiri (DCP_TRANSAKSI, DCP_FOTO, FT_LAPORAN, FT_FOTO), BUKAN di REPORTS / FILM_REPORTS,
 * sehingga tidak ikut penghapusan otomatis 45 hari dan tidak ikut arsip bulanan.
 * Sheet dibuat otomatis saat pertama dipakai (dcpSheets_). Tidak ada sheet lama yang diubah.
 * Aturan akses: REPORTER dan GUIDE_SECURITY ditolak oleh route_. Di sini: lihat/buat/ubah = ADMIN & USER (Teknisi);
 * hapus dan verifikasi hanya ADMIN. Izin granular per pengguna belum dibuat (butuh kolom baru di USERS).
 */
var DCP_SH = {
 DCP_TRANSAKSI: ['id','jenis','tanggal','jam','film','kodeDisk','ph','format','penerima','penyerah','tujuan','keterangan','status','kelengkapan','relasiId','createdBy','createdByUserId','createdAt','updatedBy','updatedAt','verifiedBy','verifiedAt','overdueNotifiedAt','modul','outlet','outletTujuan','alasanBatal','batalBy','batalAt','asalDari','namaPengirim','jenisDcp','batchId','tipeKeluar','noResi','tglKirim'],
 USER_PERMS: ['userId','perms','updatedAt','updatedBy'],
 DCP_FOTO:      ['fotoId','transaksiId','kind','fileId','uploadedAt','capturedAt','url'],
 FT_LAPORAN:    ['id','tanggal','jam','film','studio','format','server','projector','tester','status','checklist','masalah','tindakan','hasilUlang','catatanTeknisi','parentId','createdBy','createdByUserId','createdAt','updatedBy','updatedAt','batchId','ketDok','mode'],
 FT_FOTO:       ['fotoId','ftId','kind','fileId','uploadedAt','capturedAt','url']
};
var DCP_JENIS = ['KELUAR','MASUK'];
var DCP_STATUS = ['Diproses','Selesai','Bermasalah'];
var DCP_KOMP = ['HDD','USB','Kabel Power','Adaptor'];
var DCP_KOMP_ST = ['Lengkap','Tidak Ada','Rusak','Tidak Diperlukan'];
/* DCP Besar: HDD, USB, Kabel Power, Adaptor. DCP Kecil: HDD, USB. */
var DCP_KOMP_BESAR = ['HDD','USB','Kabel Power','Adaptor'];
var DCP_KOMP_KECIL = ['HDD','USB'];
var DCP_UKURAN = ['Besar','Kecil'];
var DCP_FMT = ['2D','3D','ATMOS','3D-ATMOS','2D-ATMOS','Lainnya'];
var DCP_MAXFOTO = 30, DCP_BATCH = 6, DCP_MAXB = 12000000; /* foto asli ditayangkan apa adanya; batas per foto 12 MB */
var FT_FMT = ['2D','3D','ATMOS','3D-ATMOS','Lainnya'];
var FT_STATUS = ['PASSED','FAILED','RETEST'];
var FT_VAL = ['PASS','FAIL','N/A'];
var FT_MAXFOTO = 30;
/* checklist pemeriksaan film: grup -> daftar item. id item = prefiks grup + nomor urut (stabil untuk data tersimpan). */
var FT_GROUPS = [
 {key:'pq', name:'Picture Quality', items:['Kualitas gambar','Brightness','Focus','Framing','Aspect ratio','Warna gambar','Pixel atau artifact','Kestabilan gambar']},
 {key:'aq', name:'Audio Quality', items:['Audio kiri dan kanan','Center channel','Surround channel','Subwoofer','Sinkronisasi audio dan video','Volume audio','Gangguan suara atau noise']},
 {key:'sc', name:'Subtitle dan Konten', items:['Subtitle tampil dengan benar','Sinkronisasi subtitle','Tidak ada subtitle terpotong','Tidak ada gambar atau konten yang hilang','Kesesuaian versi film']},
 {key:'tp', name:'Technical Playback', items:['KDM valid','Playback berjalan normal','Tidak ada error server','Tidak ada gangguan selama pemutaran','Durasi film sesuai','Trailer dan konten pendukung jika diperlukan']}
];
function ftItemIds_(){var out=[];FT_GROUPS.forEach(function(g){g.items.forEach(function(lbl,i){out.push({id:g.key+(i+1),grp:g.name,label:lbl});});});return out;}

/* ---------- sheet & utilitas ---------- */
var DCP_READY_=false;
function dcpSheets_(){
 if(DCP_READY_)return; /* sekali per eksekusi: hindari baca header berulang tiap dcpRows_ */
 var ss=ss_(),fresh=[];
 Object.keys(DCP_SH).forEach(function(n){
  var s=ss.getSheetByName(n);
  if(!s){s=ss.insertSheet(n);s.getRange(1,1,1,DCP_SH[n].length).setValues([DCP_SH[n]]).setFontWeight('bold');s.setFrozenRows(1);fresh.push(n);return;}
  /* kolom baru (aditif): ditambahkan di ujung header; data lama tidak diubah */
  var lastCol=s.getLastColumn(),head=lastCol?s.getRange(1,1,1,lastCol).getValues()[0]:[];
  var miss=DCP_SH[n].filter(function(k){return head.indexOf(k)<0;});
  if(miss.length)s.getRange(1,lastCol+1,1,miss.length).setValues([miss]).setFontWeight('bold');
 });
 /* tanggal (yyyy-MM-dd) dan jam (HH:mm) wajib teks: tanpa ini Spreadsheet mengubahnya jadi Date dan filter/rekap tanggal tidak cocok */
 var pr=PropertiesService.getScriptProperties();
 if(fresh.length||pr.getProperty('DCP_TXTFMT')!=='1'){
  Object.keys(DCP_SH).forEach(function(n){
   var sh=ss.getSheetByName(n);if(!sh)return;
   ['tanggal','jam'].forEach(function(k){var c=DCP_SH[n].indexOf(k);if(c>=0)sh.getRange(1,c+1,sh.getMaxRows(),1).setNumberFormat('@');});
  });
  pr.setProperty('DCP_TXTFMT','1');
 }
 DCP_READY_=true;
}
/* nilai Date dari sel (data lama yang sempat diubah Spreadsheet) dikembalikan ke teks WIB yang konsisten */
function dcpNormDT_(k,v){
 if(Object.prototype.toString.call(v)!=='[object Date]'||isNaN(v.getTime()))return v;
 return Utilities.formatDate(v,TZ,k==='jam'?'HH:mm':'yyyy-MM-dd');
}
/* penulis baris khusus modul ini (add_/upd_ bawaan hanya mengenal sheet di SH Code.gs) */
/* teks yang diawali = + - @ dibaca sebagai rumus oleh Spreadsheet (CSV/formula injection). Diberi apostrof saat ditulis. */
function dcpSafe_(v){return (typeof v==='string'&&/^[=+\-@\t\r]/.test(v))?"'"+v:v;}
function dcpUnsafe_(v){return (typeof v==='string'&&/^'[=+\-@\t\r]/.test(v))?v.slice(1):v;}
function dcpAdd_(n,o){dcpSheets_();ss_().getSheetByName(n).appendRow(DCP_SH[n].map(function(k){return o[k]===undefined?'':dcpSafe_(o[k]);}));}
function dcpUpd_(n,row,o){dcpSheets_();var s=ss_().getSheetByName(n);DCP_SH[n].forEach(function(k,j){if(o[k]!==undefined)s.getRange(row,j+1).setValue(dcpSafe_(o[k]));});}
function dcpRows_(name){dcpSheets_();return all_(name).map(function(r){Object.keys(r).forEach(function(k){if(k!=='_row')r[k]=dcpNormDT_(k,dcpUnsafe_(r[k]));});return r;});}
function dcpToday_(){return Utilities.formatDate(new Date(),TZ,'yyyy-MM-dd');}
/* izin granular. ADMIN selalu penuh. USER: default view/create/edit/photo/pdf; selebihnya lewat USER_PERMS (diatur Admin).
   REPORTER & GUIDE_SECURITY tidak punya izin modul ini (juga diblokir route_). */
var DCP_PERM_ALL=['view','create','edit','delete','verify','photo','pdf'];
var DCP_PERM_DEF={USER:['view','create','edit','photo','pdf']};
function dcpPermOf_(userId){
 var row=dcpRows_('USER_PERMS').filter(function(r){return String(r.userId)===String(userId);})[0];
 if(!row)return null;
 var arr=dcpParseJson_(row.perms,[]);
 return Array.isArray(arr)?arr.filter(function(k){return DCP_PERM_ALL.indexOf(k)>=0;}):[];
}
function dcpHas_(u,key){
 if(!u)return false;
 if(u.role==='ADMIN')return true;
 if(u.role!=='USER')return false;
 var list=dcpPermOf_(u.userId);
 if(list===null)list=DCP_PERM_DEF.USER;
 return list.indexOf(key)>=0;
}
function dcpCan_(u,need){
 if(!dcpHas_(u,need))bad_('Akses ditolak untuk tindakan ini ('+need+').');
}
function dcpNext_(prefix,rows,field){
 var pre=prefix,mx=0;
 rows.forEach(function(r){var id=String(r[field]||'');if(id.indexOf(pre)!==0)return;var t=id.slice(pre.length);var n=/^\d+$/.test(t)?parseInt(t,10):0;if(n>mx)mx=n;});
 return pre+('0000'+(mx+1)).slice(-4);
}
function dcpParseJson_(s,fb){try{return s?JSON.parse(s):fb;}catch(e){return fb;}}
function dcpRoot_(){return folder_(folder_(null,'STC XXI TECHNICAL REPORT'),'DCP & TESTING');}
function dcpSub_(parent,name){return folder_(parent,name);}

/* hapus baris foto dan file Drive milik satu induk (transaksi DCP atau laporan Testing) */
function purgeFotos_(sheetName,keyField,parentId){
 var sh=ss_().getSheetByName(sheetName),rows=dcpRows_(sheetName);
 rows.filter(function(p){return String(p[keyField])===String(parentId);}).reverse().forEach(function(p){
  try{DriveApp.getFileById(String(p.fileId)).setTrashed(true);}catch(e){Logger.log('file foto tidak dipindahkan: '+e);}
  sh.deleteRow(p._row);
 });
}

/* baca satu foto sebagai data URL. Foto Drive bersifat privat; hanya lewat aksi ini dan hanya bagi yang punya izin view. */
function fotoGet_(u,sheetName,fotoId,key){
 dcpCan_(u,'view');
 var p=dcpRows_(sheetName).filter(function(x){return String(x.fotoId)===String(fotoId);})[0];
 if(!p||!p.fileId)bad_('Foto tidak ditemukan.');
 var blob=DriveApp.getFileById(String(p.fileId)).getBlob();
 var bytes=blob.getBytes();
 if(bytes.length>DCP_MAXB)bad_('Foto terlalu besar untuk ditampilkan.');
 var mime=imgKind_(bytes)||'image/jpeg';
 return {fotoId:p.fotoId,url:'data:'+mime+';base64,'+Utilities.base64Encode(bytes)};
}
/* DCP Management (data legacy) hanya untuk Admin. Transaksi Pelaporan (modul PEL) tetap dapat diakses sesuai izin. */
function dcpAdmin_(u){if(!u||u.role!=='ADMIN')bad_('DCP Management hanya dapat diakses oleh Admin.');}
function dcpGuardTr_(u,tr){if(tr&&tr.modul!=='PEL')dcpAdmin_(u);}
function dcpGuardFoto_(u,p){if(!p)return;var tr=dcpRows_('DCP_TRANSAKSI').filter(function(x){return x.id===p.transaksiId;})[0];dcpGuardTr_(u,tr);}
/* Foto Pelaporan/DCP: ambil beberapa foto dalam SATU panggilan (sheet dibaca sekali). thumb=true -> gambar kecil dari Drive
   (puluhan KB, cukup untuk kotak 84px); bila thumbnail belum tersedia, jatuh ke ukuran penuh. Aturan akses sama dengan sebelumnya. */
function dcpPhotoBatch_(u,q){
 dcpCan_(u,'view');
 var ids=(Array.isArray(q.ids)?q.ids:[]).map(function(x){return String(x||'');}).filter(function(x,i,a){return x&&a.indexOf(x)===i;}).slice(0,8);
 if(!ids.length)return {items:[]};
 var want={},fotos={},trs=null,thumb=!!q.thumb;
 ids.forEach(function(i){want[i]=1;});
 dcpRows_('DCP_FOTO').forEach(function(p){if(want[String(p.fotoId)])fotos[String(p.fotoId)]=p;});
 return {items:ids.map(function(id){
  try{
   var p=fotos[id];if(!p||!p.fileId)return {fotoId:id,error:'Foto tidak ditemukan.'};
   if(u.role!=='ADMIN'){
    if(!trs){trs={};dcpRows_('DCP_TRANSAKSI').forEach(function(t){trs[String(t.id)]=t;});}
    var tr=trs[String(p.transaksiId)];if(tr&&tr.modul!=='PEL')return {fotoId:id,error:'DCP Management hanya dapat diakses oleh Admin.'};
   }
   var cc=CacheService.getScriptCache(),ck='dth_'+p.fileId;
   if(thumb){var hit=cc.get(ck);if(hit)return {fotoId:id,url:hit,thumb:true};}
   var file=DriveApp.getFileById(String(p.fileId)),blob=null,bytes=null,isThumb=false;
   if(thumb){try{blob=file.getThumbnail();bytes=blob?blob.getBytes():null;isThumb=!!(bytes&&bytes.length);}catch(e){bytes=null;}}
   if(!bytes||!bytes.length){bytes=file.getBlob().getBytes();isThumb=false;}
   if(bytes.length>DCP_MAXB)return {fotoId:id,error:'Foto terlalu besar untuk ditampilkan.'};
   var mime=imgKind_(bytes)||(isThumb&&blob&&blob.getContentType())||'image/jpeg';
   var du='data:'+mime+';base64,'+Utilities.base64Encode(bytes);
   if(isThumb&&du.length<90000){try{cc.put(ck,du,21600);}catch(e){}}
   return {fotoId:id,url:du,thumb:isThumb};
  }catch(e){return {fotoId:id,error:'Foto tidak dapat dimuat.'};}
 })};
}
function dcpPhotoGet_(u,q){var r=dcpPhotoBatch_(u,{ids:[q.fotoId]}).items[0];if(!r||r.error)bad_((r&&r.error)||'Foto tidak ditemukan.');return {fotoId:r.fotoId,url:r.url};}
function ftPhotoGet_(u,q){return fotoGet_(u,'FT_FOTO',q.fotoId);}

/* ---------- DCP: validasi & kelengkapan ---------- */
function dcpCheckList_(raw,ukuran){
 var src=Array.isArray(raw)?raw:[],out=[];
 (ukuran==='Kecil'?DCP_KOMP_KECIL:DCP_KOMP_BESAR).forEach(function(nm){
  var r=null;src.forEach(function(x){if(x&&String(x.name)===nm)r=x;});r=r||{};
  var st=String(r.status||'');if(DCP_KOMP_ST.indexOf(st)<0)bad_('Pilih status untuk '+nm+'.');
  var ket=s_(r.note,300);
  if((st==='Tidak Ada'||st==='Rusak')&&!ket)bad_('Keterangan wajib diisi untuk '+nm+' ('+st+').');
  out.push({name:nm,status:st,note:ket});
 });
 return out;
}
function dcpIncomplete_(list){return list.filter(function(c){return c.status==='Tidak Ada'||c.status==='Rusak';});}

/* tautan DCP keluar -> masuk: MASUK menyimpan relasiId = id KELUAR. Satu KELUAR hanya boleh punya satu MASUK. */
function dcpLinkedMap_(rows){var m={};rows.forEach(function(r){if(r.jenis==='MASUK'&&r.relasiId)m[String(r.relasiId)]=r.id;});return m;}

/* ---------- DCP: handler ---------- */
function dcpSave_(u,q){
 dcpAdmin_(u);
 if(!q.id)bad_('DCP Keluar/Masuk dilaporkan lewat Pelaporan DCP. DCP Management hanya untuk melihat, mengubah, dan menghapus data.');
 dcpCan_(u,'edit');
 var rows=dcpRows_('DCP_TRANSAKSI'),now=Date.now();
 /* satu sumber data: baris Pelaporan (modul PEL) selalu disimpan lewat aturan Pelaporan, dari modul mana pun diubah */
 if(q.id){var o0=rows.filter(function(r){return r.id===q.id;})[0];if(o0&&o0.modul==='PEL')return pelSave_(u,q);}
 var jenis=s_(q.jenis,10);if(DCP_JENIS.indexOf(jenis)<0)bad_('Jenis transaksi harus KELUAR atau MASUK.');
 var tanggal=s_(q.tanggal,10),jam=s_(q.jam,5);
 if(!/^\d{4}-\d{2}-\d{2}$/.test(tanggal)||isNaN(Date.parse(tanggal+'T00:00:00+07:00')))bad_('Tanggal transaksi tidak valid.');
 if(!/^([01]\d|2[0-3]):[0-5]\d$/.test(jam))bad_('Jam transaksi tidak valid.');
 var film=s_(q.film,200),kodeDisk=s_(q.kodeDisk,60),ph=s_(q.ph,160),format=s_(q.format,20),penerima=s_(q.penerima,120),penyerah=s_(q.penyerah,120),tujuan=s_(q.tujuan,200),ket=s_(q.keterangan,2000);
 if(!film||!kodeDisk||!ph||!penerima||!penyerah||!tujuan)bad_('Lengkapi film, kode disk, PH, penerima, penyerah, dan tujuan/asal.');
 if(format&&DCP_FMT.indexOf(format)<0)bad_('Format film tidak dikenal.');
 var status=s_(q.status,20);if(DCP_STATUS.indexOf(status)<0)bad_('Pilih status: Diproses, Selesai, atau Bermasalah.');
 var kel=dcpCheckList_(q.kelengkapan);
 if(dcpIncomplete_(kel).length&&!ket)bad_('Ada komponen tidak lengkap atau rusak. Isi keterangan.');
 var old=q.id?rows.filter(function(r){return r.id===q.id;})[0]:null;
 if(q.id&&!old)bad_('Transaksi tidak ditemukan.');
 if(old&&u.role!=='ADMIN'&&String(old.createdByUserId)!==String(u.userId))bad_('Hanya Admin atau pembuat transaksi yang dapat mengubah.');
 if(old&&old.jenis!==jenis)bad_('Jenis transaksi tidak dapat diubah.');
 if(old&&old.verifiedAt&&old.modul==='PEL'&&u.role!=='ADMIN')bad_('Laporan sudah tervalidasi dan tidak dapat diubah. Batalkan validasi terlebih dahulu.');
 /* tautan keluar-masuk */
 var relasiId='';
 if(jenis==='MASUK'){
  if(old&&old.relasiId)relasiId=String(old.relasiId);
  var want=s_(q.relasiId,40)||relasiId;
  if(want){
   var k=rows.filter(function(r){return r.id===want&&r.jenis==='KELUAR';})[0];
   if(!k)bad_('Transaksi keluar yang dipilih tidak ditemukan.');
   if(String(k.kodeDisk)!==kodeDisk)bad_('Kode disk tidak sama dengan transaksi keluar ('+k.kodeDisk+').');
   relasiId=k.id;
  }else{
   /* otomatis: KELUAR terbaru dengan kode disk yang sama dan belum punya MASUK */
   var linked=dcpLinkedMap_(rows);
   var cand=rows.filter(function(r){return r.jenis==='KELUAR'&&String(r.kodeDisk)===kodeDisk&&!linked[String(r.id)];})
    .sort(function(a,b){return Number(b.createdAt)-Number(a.createdAt);});
   if(cand[0]){relasiId=cand[0].id;}
  }
  if(relasiId){
   var lk=dcpLinkedMap_(rows);
   if(lk[relasiId]&&(!old||lk[relasiId]!==old.id))bad_('Transaksi keluar ini sudah memiliki transaksi masuk.');
  }
 }
 var o={jenis:jenis,tanggal:tanggal,jam:jam,film:film,kodeDisk:kodeDisk,ph:ph,format:format,penerima:penerima,penyerah:penyerah,tujuan:tujuan,keterangan:ket,status:status,kelengkapan:JSON.stringify(kel),relasiId:relasiId,updatedBy:u.fullName,updatedAt:now};
 if(old){if(!old.batchId&&q.batchId)o.batchId=s_(q.batchId,40);dcpUpd_('DCP_TRANSAKSI',old._row,o);log_(u,'EDIT_DCP','DCP_TRANSAKSI',old.id,jenis+' '+film+' ('+kodeDisk+')');return old.id;}
 o.id=dcpNext_('DCP-'+(jenis==='KELUAR'?'OUT':'IN')+'-'+Utilities.formatDate(new Date(now),TZ,'yyyy')+'-',rows.filter(function(r){return r.jenis===jenis;}),'id');
 o.createdBy=u.fullName;o.createdByUserId=u.userId;o.createdAt=now;o.batchId=s_(q.batchId,40);
 dcpAdd_('DCP_TRANSAKSI',o);log_(u,'BUAT_DCP','DCP_TRANSAKSI',o.id,jenis+' '+film+' ('+kodeDisk+')');
 return o.id;
}

/* Simpan banyak item DCP dalam satu laporan (batchId sama). Atomik: bila satu item gagal, item yang sudah
   tertulis dalam permintaan ini dihapus kembali (rollback) sehingga tidak ada laporan setengah jadi.
   kind: 'dcp' (DCP Management, Admin) | 'pel' (Pelaporan DCP). Berjalan di bawah LockService (akhiran _save). */
function dcpBatchSave_(u,q){
 var items=q.items;
 if(Object.prototype.toString.call(items)!=='[object Array]'||!items.length)bad_('Tidak ada item DCP untuk disimpan.');
 if(items.length>30)bad_('Maksimal 30 item DCP dalam satu laporan.');
 var kind=q.kind==='pel'?'pel':'dcp';
 if(kind!=='pel')bad_('DCP Keluar/Masuk dilaporkan lewat Pelaporan DCP. DCP Management hanya untuk melihat, mengubah, dan menghapus data.');
 var batchId=s_(q.batchId,40)||('B'+Date.now().toString(36).toUpperCase());
 var seen={};
 items.forEach(function(it,i){var k=String(it&&it.kodeDisk||'').trim().toLowerCase();if(!k)return;if(seen[k]!==undefined)bad_('Item '+(i+1)+': kode disk '+it.kodeDisk+' sama dengan item '+(seen[k]+1)+'. Satu kode disk hanya boleh sekali per laporan.');seen[k]=i;});
 var ids=[];
 try{
  items.forEach(function(it,i){
   var x={};Object.keys(it||{}).forEach(function(k){x[k]=it[k];});
   x.batchId=batchId;x.id='';
   try{ids.push(kind==='pel'?pelSave_(u,x):dcpSave_(u,x));}
   catch(e){throw new Error('!Item '+(i+1)+(x.film?' ('+s_(x.film,60)+')':'')+': '+String(e&&e.message||e).replace(/^!/,''));}
  });
 }catch(err){
  var sh=ss_().getSheetByName('DCP_TRANSAKSI');
  ids.slice().reverse().forEach(function(id){
   var r=dcpRows_('DCP_TRANSAKSI').filter(function(x){return x.id===id;})[0];
   if(r)sh.deleteRow(r._row);
  });
  if(ids.length)log_(u,'ROLLBACK_BATCH','DCP_TRANSAKSI',batchId,'Dibatalkan: '+ids.join(', '));
  throw err;
 }
 return {batchId:batchId,ids:ids};
}
function dcpShape_(r,rows){var o=dcpShape0_(r,rows);o.statusAsli=o.status;if(o.jenis==='KELUAR')o.status='Sudah Keluar'; /* keluar dari outlet STC = hijau, bukan DCP aktif */
 o.tipeKeluar=r.tipeKeluar||(r.jenis==='KELUAR'&&r.modul==='PEL'?'PINDAH':'');o.noResi=r.noResi||'';o.tglKirim=dcpDay_(r.tglKirim)||'';return o;}
function dcpShape0_(r,rows){
 var kel=dcpParseJson_(r.kelengkapan,[]);
 var linked=dcpLinkedMap_(rows);
 return {id:r.id,jenis:r.jenis,tanggal:r.tanggal,jam:r.jam,film:r.film,kodeDisk:r.kodeDisk,ph:r.ph,format:r.format,penerima:r.penerima,penyerah:r.penyerah,tujuan:r.tujuan,keterangan:r.keterangan,status:r.status,kelengkapan:kel,relasiId:r.relasiId||'',masukId:linked[String(r.id)]||'',createdBy:r.createdBy,createdByUserId:r.createdByUserId,createdAt:Number(r.createdAt)||0,updatedBy:r.updatedBy,updatedAt:Number(r.updatedAt)||0,incomplete:dcpIncomplete_(kel).length>0,verifiedBy:r.verifiedBy||'',verifiedAt:Number(r.verifiedAt)||0,batchId:r.batchId||'',modul:r.modul||'',outlet:r.outlet||'',outletTujuan:r.outletTujuan||'',jenisDcp:r.jenisDcp||'',asalDari:r.asalDari||'',namaPengirim:r.namaPengirim||''};
}
function dcpList_(u,q){
 dcpAdmin_(u);
 dcpCan_(u,'view');
 var rows=dcpRows_('DCP_TRANSAKSI'),term=s_(q.q,120).toLowerCase(),from=q.from||'',to=q.to||'';
 var list=rows.filter(function(r){
  if(q.jenis&&r.jenis!==q.jenis)return false;
  if(q.status&&r.status!==q.status)return false;
  if(from&&String(r.tanggal)<from)return false;
  if(to&&String(r.tanggal)>to)return false;
  if(term&&[r.film,r.kodeDisk,r.ph,r.id].join(' ').toLowerCase().indexOf(term)<0)return false;
  return true;
 }).sort(function(a,b){return (b.tanggal+b.jam).localeCompare(a.tanggal+a.jam)||(Number(b.createdAt)||0)-(Number(a.createdAt)||0);});
 return list.map(function(r){return dcpShape_(r,rows);});
}
function dcpGet_(u,q){
 dcpAdmin_(u);
 dcpCan_(u,'view');
 var rows=dcpRows_('DCP_TRANSAKSI'),r=rows.filter(function(x){return x.id===q.id;})[0];
 if(!r)bad_('Transaksi tidak ditemukan.');
 var fotos=dcpRows_('DCP_FOTO').filter(function(p){return p.transaksiId===r.id&&p.fileId;})
  .sort(function(a,b){return Number(a.uploadedAt)-Number(b.uploadedAt);})
  .map(function(p){return {fotoId:p.fotoId,kind:p.kind,url:'',uploadedAt:Number(p.uploadedAt)||0};});
 var rel=null;
 if(r.relasiId){var k=rows.filter(function(x){return x.id===r.relasiId;})[0];if(k)rel=dcpShape_(k,rows);}
 return {transaksi:dcpShape_(r,rows),foto:fotos,relasi:rel};
}
function dcpDel_(u,q){
 dcpAdmin_(u);
 dcpCan_(u,'delete');
 var rows=dcpRows_('DCP_TRANSAKSI'),r=rows.filter(function(x){return x.id===q.id;})[0];
 if(!r)bad_('Transaksi tidak ditemukan.');
 /* Admin berhak menghapus semua data (juga yang tervalidasi atau tertaut). Tautan MASUK ke baris ini dilepas, data MASUK tetap utuh. */
 rows.forEach(function(x){if(String(x.relasiId)===String(r.id))dcpUpd_('DCP_TRANSAKSI',x._row,{relasiId:''});});
 purgeFotos_('DCP_FOTO','transaksiId',r.id);
 ss_().getSheetByName('DCP_TRANSAKSI').deleteRow(r._row);
 log_(u,'HAPUS_DCP','DCP_TRANSAKSI',r.id,r.jenis+' '+r.film);
 return true;
}
/* foto DCP: maks 30 per transaksi, 6 per kiriman; file ke Drive, tautan ke sheet */
function dcpPhotoSave_(u,q){
 dcpCan_(u,'photo');
 var rows=dcpRows_('DCP_TRANSAKSI'),r=rows.filter(function(x){return x.id===q.id;})[0];
 dcpGuardTr_(u,r);
 if(!r)bad_('Transaksi tidak ditemukan.');
 if(u.role!=='ADMIN'&&String(r.createdByUserId)!==String(u.userId))bad_('Hanya Admin atau pembuat transaksi yang dapat menambah foto.'); /* sebelumnya tidak diperiksa di server */
 if(r.verifiedAt&&r.modul==='PEL'&&u.role!=='ADMIN')bad_('Laporan sudah tervalidasi. Batalkan validasi terlebih dahulu untuk menambah foto.');
 var have=dcpRows_('DCP_FOTO').filter(function(p){return p.transaksiId===r.id;}),n=have.length,add=0,fail=0;
 var fol=dcpSub_(dcpSub_(dcpRoot_(),'DCP'),r.id),ps=ss_().getSheetByName('DCP_FOTO');
 (Array.isArray(q.photos)?q.photos:[]).slice(0,DCP_BATCH).forEach(function(ph){try{
  if(n>=DCP_MAXFOTO)throw new Error('melebihi batas '+DCP_MAXFOTO+' foto');
  var b=Utilities.base64Decode(String(ph&&ph.b64||'').replace(/^data:[^,]*,/,'')),mime=imgKind_(b);
  if(!mime||b.length>DCP_MAXB)throw new Error('format/ukuran tidak valid');
  var kind=s_(ph&&ph.kind,40)||'Foto';
  var f=fol.createFile(Utilities.newBlob(b,mime,r.id+'-'+('0'+(n+1)).slice(-2)+'.'+EXT[mime]));
  n++;add++;
  var up=Date.now();
  dcpAdd_('DCP_FOTO',{fotoId:uid_('DF'),transaksiId:r.id,kind:kind,fileId:f.getId(),uploadedAt:up,capturedAt:Number(ph&&ph.capturedAt)||up,url:''});
 }catch(e){fail++;Logger.log('foto DCP gagal: '+e);}});
 log_(u,'FOTO_DCP','DCP_FOTO',r.id,'+'+add+' foto');
 return {added:add,failed:fail,total:n};
}
function dcpPhotoDel_(u,q){
 dcpCan_(u,'photo');
 var p=dcpRows_('DCP_FOTO').filter(function(x){return x.fotoId===q.fotoId;})[0];
 dcpGuardFoto_(u,p);
 if(!p)bad_('Foto tidak ditemukan.');
 var tr=dcpRows_('DCP_TRANSAKSI').filter(function(x){return x.id===p.transaksiId;})[0];
 if(tr&&u.role!=='ADMIN'&&String(tr.createdByUserId)!==String(u.userId))bad_('Hanya Admin atau pembuat transaksi yang dapat menghapus foto.');
 if(tr&&tr.verifiedAt&&tr.modul==='PEL'&&u.role!=='ADMIN')bad_('Foto laporan tervalidasi tidak dapat dihapus.');
 try{DriveApp.getFileById(String(p.fileId)).setTrashed(true);}catch(e){Logger.log('file foto DCP tidak dipindahkan: '+e);}
 ss_().getSheetByName('DCP_FOTO').deleteRow(p._row);
 log_(u,'HAPUS_FOTO_DCP','DCP_FOTO',p.transaksiId,p.kind);
 return true;
}
function dcpDash_(u,q){
 dcpAdmin_(u);
 dcpCan_(u,'view');
 var rows=dcpRows_('DCP_TRANSAKSI'),linked=dcpLinkedMap_(rows),days=Number(P_().getProperty('DCP_OVERDUE_HARI'))||7;
 var keluar=rows.filter(function(r){return r.jenis==='KELUAR';}),masuk=rows.filter(function(r){return r.jenis==='MASUK';});
 var belum=keluar.filter(function(r){return r.modul!=='PEL'&&!linked[String(r.id)];}); /* KELUAR Pelaporan = kirim ke outlet lain, tidak ditunggu kembali */
 var lastBy={};rows.slice().sort(function(a,b){var x=dcpDay_(a.tanggal)+' '+dcpHm_(a.jam),y=dcpDay_(b.tanggal)+' '+dcpHm_(b.jam);return x<y?-1:x>y?1:(Number(a.createdAt)||0)-(Number(b.createdAt)||0);}).forEach(function(r){lastBy[String(r.kodeDisk).toLowerCase()]=r;});
 var act=Object.keys(lastBy).map(function(k){return lastBy[k];}).filter(function(r){return r.jenis==='MASUK';}); /* DCP aktif = catatan terakhir kode disk adalah MASUK */
 var cutoff=Utilities.formatDate(new Date(Date.now()-days*864e5),TZ,'yyyy-MM-dd');
 var overdue=belum.filter(function(r){return String(r.tanggal)<cutoff;}).map(function(r){return {id:r.id,film:r.film,kodeDisk:r.kodeDisk,tanggal:r.tanggal};});
 return {
  totalKeluar:keluar.length,totalMasuk:masuk.length,belumKembali:belum.length,
  aktif:act.length,dcpAktif:act.length,sudahKeluar:keluar.length, /* dcpAktif = nama lama (kompatibel dengan frontend repo sebelumnya) */
  dalamProses:act.filter(function(r){return r.status==='Diproses';}).length,
  bermasalah:act.filter(function(r){return r.status==='Bermasalah'||dcpIncomplete_(dcpParseJson_(r.kelengkapan,[])).length>0;}).length,
  batasHari:days,overdue:overdue,
  terbaru:rows.slice().sort(function(a,b){return (b.createdAt||0)-(a.createdAt||0);}).slice(0,8).map(function(r){return dcpShape_(r,rows);})
 };
}

/* ---------- Daily Film Testing ---------- */
/* checklist: [{id, val:'PASS'|'FAIL'|'N/A', ket}] -> status dihitung di server. Klien tidak dapat memaksa PASSED. */
function ftChecklist_(raw){
 var src=Array.isArray(raw)?raw:[],ids=ftItemIds_(),out=[];
 ids.forEach(function(it){
  var r=null;src.forEach(function(x){if(x&&String(x.id)===it.id)r=x;});r=r||{};
  var v=String(r.val||'');if(FT_VAL.indexOf(v)<0)bad_('Isi PASS, FAIL, atau N/A untuk setiap item: "'+it.label+'".');
  var ket=s_(r.ket,500);
  if(v==='FAIL'&&!ket)bad_('Item FAIL wajib diberi keterangan masalah: "'+it.label+'".');
  out.push({id:it.id,val:v,ket:ket});
 });
 return out;
}
function ftVerdict_(items,retest){
 if(items.some(function(x){return x.val==='FAIL';}))return 'FAILED';
 if(retest)return 'RETEST';
 return 'PASSED';
}
/* satu studio hanya boleh sekali per laporan (baris awal, bukan pengetesan ulang). key = kunci laporan dari ftKeys_ */
function ftDupStudio_(rows,key,studio,exceptId){
 var keys=ftKeys_(rows),z=String(studio).trim().toLowerCase();
 return rows.some(function(r){return !r.parentId&&r.id!==exceptId&&keys[r.id]===key&&String(r.studio).trim().toLowerCase()===z;});
}
/* ===== Daily Ceklis Konten Film: satu baris FT_LAPORAN (mode HARIAN) per studio per hari; foto memakai FT_FOTO yang sama ===== */
var FTD_ITEMS=['Konten sudah ditransfer','KDM sudah diupload','Testplay 5 menit AWAL','Testplay 5 menit AKHIR']; /* tidak ada batas jam: laporan boleh dibuat kapan saja */
function ftdStudios_(){var p=PropertiesService.getScriptProperties().getProperty('FTD_STUDIOS');var l=String(p||'Studio 1,Studio 2,Studio 3,Studio 4').split(',').map(function(x){return s_(x,40);}).filter(Boolean);return l.length?l:['Studio 1'];}
function ftdToday_(){return Utilities.formatDate(new Date(),TZ,'yyyy-MM-dd');}
function ftdChecks_(r){var a=dcpParseJson_(r.checklist,[]);return FTD_ITEMS.map(function(_,i){return a[i]===1||a[i]===true||a[i]==='1';});}
function ftdList_(u,q){
 dcpCan_(u,'view');
 var tgl=s_(q.tanggal,10)||ftdToday_();
 if(!/^\d{4}-\d{2}-\d{2}$/.test(tgl)||isNaN(Date.parse(tgl+'T00:00:00+07:00')))bad_('Tanggal tidak valid.');
 var rows=dcpRows_('FT_LAPORAN').filter(function(r){return r.mode==='HARIAN'&&dcpDay_(r.tanggal)===tgl;});
 var fl={};dcpRows_('FT_FOTO').forEach(function(p){if(p.fileId)(fl[p.ftId]=fl[p.ftId]||[]).push({fotoId:p.fotoId,kind:p.kind,capturedAt:Number(p.capturedAt)||0});});
 var names=ftdStudios_();rows.forEach(function(r){if(names.indexOf(String(r.studio))<0)names.push(String(r.studio));});
 var today=ftdToday_();
 return {tanggal:tgl,hariIni:today,items:FTD_ITEMS,
  studios:names.map(function(z){var r=rows.filter(function(x){return String(x.studio)===z;})[0];
   if(!r)return {studio:z,id:'',checks:FTD_ITEMS.map(function(){return false;}),pic:'',createdByUserId:'',foto:[],complete:false,editable:u.role==='ADMIN'||tgl===today};
   var c=ftdChecks_(r),f=fl[r.id]||[];
   return {studio:z,id:r.id,checks:c,pic:r.tester||r.createdBy||'',createdByUserId:r.createdByUserId||'',createdAt:Number(r.createdAt)||0,updatedAt:Number(r.updatedAt)||0,jam:r.jam||'',foto:f,
    complete:c.every(Boolean)&&f.length>0,
    editable:u.role==='ADMIN'||(tgl===today&&String(r.createdByUserId)===String(u.userId))};})};
}
function ftdSave_(u,q){
 var tgl=s_(q.tanggal,10),studio=s_(q.studio,40);
 if(!/^\d{4}-\d{2}-\d{2}$/.test(tgl)||isNaN(Date.parse(tgl+'T00:00:00+07:00')))bad_('Tanggal tidak valid.');
 var raw=Array.isArray(q.checks)?q.checks:[],checks=FTD_ITEMS.map(function(_,i){return raw[i]===true||raw[i]===1||raw[i]==='1'?1:0;});
 var rows=dcpRows_('FT_LAPORAN'),old=rows.filter(function(r){return r.mode==='HARIAN'&&dcpDay_(r.tanggal)===tgl&&String(r.studio)===studio;})[0];
 if(!old&&ftdStudios_().indexOf(studio)<0)bad_('Studio tidak dikenal.');
 dcpCan_(u,old?'edit':'create');
 if(u.role!=='ADMIN'){
  if(tgl!==ftdToday_())bad_('Ceklis hanya dapat diisi untuk hari ini. Hubungi Admin untuk tanggal lain.');
  if(old&&String(old.createdByUserId)!==String(u.userId))bad_('Ceklis studio ini sudah diisi oleh '+(old.createdBy||'pengguna lain')+'. Hanya pembuat atau Admin yang dapat mengubah.');
 }
 var now=Date.now();
 if(old){dcpUpd_('FT_LAPORAN',old._row,{checklist:JSON.stringify(checks),updatedBy:u.fullName,updatedAt:now});log_(u,'EDIT_CEKLIS','FT_LAPORAN',old.id,studio+' '+tgl+' '+checks.join(''));return {id:old.id,created:false};}
 var o={tanggal:tgl,jam:Utilities.formatDate(new Date(now),TZ,'HH:mm'),film:'Ceklis harian',studio:studio,format:'',server:'',projector:'',tester:u.fullName,status:'',checklist:JSON.stringify(checks),masalah:'',tindakan:'',hasilUlang:'',catatanTeknisi:'',parentId:'',createdBy:u.fullName,createdByUserId:u.userId,createdAt:now,batchId:'',ketDok:'',mode:'HARIAN'};
 o.id=dcpNext_('FT-'+Utilities.formatDate(new Date(now),TZ,'yyyy')+'-',rows,'id');
 dcpAdd_('FT_LAPORAN',o);log_(u,'BUAT_CEKLIS','FT_LAPORAN',o.id,studio+' '+tgl+' '+checks.join(''));
 return {id:o.id,created:true};
}
function ftSave_(u,q,ctxRows){
 dcpCan_(u,q.id?'edit':'create');
 var rows=ctxRows||dcpRows_('FT_LAPORAN'),now=Date.now();
 var tanggal=s_(q.tanggal,10),jam=s_(q.jam,5);
 if(!/^\d{4}-\d{2}-\d{2}$/.test(tanggal)||isNaN(Date.parse(tanggal+'T00:00:00+07:00')))bad_('Tanggal testing tidak valid.');
 if(!/^([01]\d|2[0-3]):[0-5]\d$/.test(jam))bad_('Jam testing tidak valid.');
 var film=s_(q.film,200),studio=s_(q.studio,40),format=s_(q.format,20),server=s_(q.server,80),projector=s_(q.projector,80),tester=s_(q.tester,120);
 if(!film||!studio||!tester)bad_('Lengkapi judul film, nomor studio, dan nama tester.');
 if(FT_FMT.indexOf(format)<0)bad_('Pilih format film.');
 var items=ftChecklist_(q.checklist),retest=!!q.retest;
 var status=ftVerdict_(items,retest);
 var masalah=s_(q.masalah,2000),tindakan=s_(q.tindakan,2000),hasilUlang=s_(q.hasilUlang,2000),catatan=s_(q.catatanTeknisi,2000),ketDok=s_(q.ketDok,1000);
 if(status==='FAILED'&&!masalah)bad_('Hasil FAILED: isi deskripsi masalah.');
 if(status==='FAILED'&&!tindakan)bad_('Hasil FAILED: isi tindakan perbaikan yang direncanakan.');
 var parentId=s_(q.parentId,40),old=q.id?rows.filter(function(r){return r.id===q.id;})[0]:null;
 if(q.id&&!old)bad_('Laporan testing tidak ditemukan.');
 if(old){
  if(u.role!=='ADMIN'&&String(old.createdByUserId)!==String(u.userId))bad_('Hanya Admin atau pembuat laporan yang dapat mengubah.');
  if(rows.some(function(r){return r.parentId===old.id;}))bad_('Laporan ini sudah memiliki pengetesan ulang. Riwayat tidak dapat diubah.');
 }
 var o={tanggal:tanggal,jam:jam,film:film,studio:studio,format:format,server:server,projector:projector,tester:tester,status:status,checklist:JSON.stringify(items),masalah:masalah,tindakan:tindakan,hasilUlang:hasilUlang,catatanTeknisi:catatan,ketDok:ketDok,updatedBy:u.fullName,updatedAt:now};
 if(old){
  /* ganti nomor studio tidak boleh menabrak studio lain dalam laporan yang sama */
  if(!old.parentId&&String(old.studio).trim().toLowerCase()!==studio.toLowerCase()&&ftDupStudio_(rows,ftKeys_(rows)[old.id],studio,old.id))bad_('Studio '+studio+' sudah ada di laporan ini.');
  dcpUpd_('FT_LAPORAN',old._row,o);log_(u,'EDIT_TESTING','FT_LAPORAN',old.id,film+' → '+status);return old.id;}
 if(parentId){
  var par=rows.filter(function(r){return r.id===parentId;})[0];
  if(!par)bad_('Laporan induk tidak ditemukan.');
  if(par.status==='PASSED')bad_('Film yang sudah PASSED tidak perlu dites ulang.');
  if(rows.some(function(r){return r.parentId===parentId;}))bad_('Pengetesan ulang untuk laporan ini sudah ada.');
  o.parentId=parentId;
 }else{
  o.parentId='';
  var bk=s_(q.batchId,40);
  if(bk&&ftDupStudio_(rows,bk,studio,''))bad_('Studio '+studio+' sudah ada di laporan ini.');
 }
 o.id=dcpNext_('FT-'+Utilities.formatDate(new Date(now),TZ,'yyyy')+'-',rows,'id');
 o.createdBy=u.fullName;o.createdByUserId=u.userId;o.createdAt=now;o.batchId=parentId?'':s_(q.batchId,40);
 dcpAdd_('FT_LAPORAN',o);
 if(ctxRows)ctxRows.push(o);
 log_(u,parentId?'RETEST_TESTING':'BUAT_TESTING','FT_LAPORAN',o.id,film+' → '+status+(parentId?' (retest dari '+parentId+')':''));
 return o.id;
}
/* Simpan hasil testing banyak studio dalam satu laporan (batchId sama). Data tiap studio tetap baris sendiri.
   Atomik: bila satu studio gagal, baris yang sudah tertulis pada permintaan ini dihapus kembali. Berjalan di bawah kunci aksi *_save. */
function ftBatchSave_(u,q){
 var items=q.items;
 if(Object.prototype.toString.call(items)!=='[object Array]'||!items.length)bad_('Pilih minimal satu studio.');
 if(items.length>20)bad_('Maksimal 20 studio dalam satu laporan.');
 var batchId=s_(q.batchId,40)||('FB'+Date.now().toString(36).toUpperCase()),seen={};
 items.forEach(function(it,i){var k=String(it&&it.studio||'').trim().toLowerCase();if(!k)bad_('Studio '+(i+1)+': nomor studio kosong.');if(seen[k])bad_('Studio '+it.studio+' dipilih dua kali.');seen[k]=1;});
 var ids=[],rowsOnce=dcpRows_('FT_LAPORAN');
 try{
  items.forEach(function(it){
   var x={};Object.keys(it||{}).forEach(function(k){x[k]=it[k];});
   x.batchId=batchId;x.id='';x.parentId='';
   try{ids.push(ftSave_(u,x,rowsOnce));}
   catch(e){throw new Error('!Studio '+s_(x.studio,20)+': '+String(e&&e.message||e).replace(/^!/,''));}
  });
 }catch(err){
  var sh=ss_().getSheetByName('FT_LAPORAN');
  ids.slice().reverse().forEach(function(id){var r=dcpRows_('FT_LAPORAN').filter(function(x){return x.id===id;})[0];if(r)sh.deleteRow(r._row);});
  if(ids.length)log_(u,'ROLLBACK_BATCH','FT_LAPORAN',batchId,'Dibatalkan: '+ids.join(', '));
  throw err;
 }
 return {batchId:batchId,ids:ids};
}
function ftShape_(r){
 return {id:r.id,tanggal:r.tanggal,jam:r.jam,film:r.film,studio:r.studio,format:r.format,server:r.server,projector:r.projector,tester:r.tester,status:r.status,checklist:dcpParseJson_(r.checklist,[]),masalah:r.masalah,tindakan:r.tindakan,hasilUlang:r.hasilUlang,catatanTeknisi:r.catatanTeknisi,ketDok:r.ketDok||'',fotoCount:Number(r._fc)||0,parentId:r.parentId||'',batchId:r.batchId||'',createdBy:r.createdBy,createdByUserId:r.createdByUserId,createdAt:Number(r.createdAt)||0,updatedAt:Number(r.updatedAt)||0};
}
function ftList_(u,q){
 dcpCan_(u,'view');
 var rows=dcpRows_('FT_LAPORAN').filter(function(r){return r.mode!=='HARIAN';}),term=s_(q.q,120).toLowerCase();
 /* jumlah foto per laporan (satu kali baca FT_FOTO): ekspor tahu studio mana yang punya foto tanpa memanggil ft_get satu per satu */
 var fc={};dcpRows_('FT_FOTO').forEach(function(p){if(p.fileId)fc[p.ftId]=(fc[p.ftId]||0)+1;});
 rows.forEach(function(r){r._fc=fc[r.id]||0;});
 return rows.filter(function(r){
  if(q.status&&r.status!==q.status)return false;
  if(q.studio&&String(r.studio)!==String(q.studio))return false;
  if(q.tester&&String(r.tester).indexOf(q.tester)<0)return false;
  if(q.from&&String(r.tanggal)<q.from)return false;
  if(q.to&&String(r.tanggal)>q.to)return false;
  if(term&&[r.film,r.id].join(' ').toLowerCase().indexOf(term)<0)return false;
  return true;
 }).sort(function(a,b){return (b.tanggal+b.jam).localeCompare(a.tanggal+a.jam)||(Number(b.createdAt)||0)-(Number(a.createdAt)||0);}).map(ftShape_);
}
function ftGet_(u,q){
 dcpCan_(u,'view');
 var rows=dcpRows_('FT_LAPORAN'),r=rows.filter(function(x){return x.id===q.id;})[0];
 if(!r)bad_('Laporan testing tidak ditemukan.');
 var fotos=dcpRows_('FT_FOTO').filter(function(p){return p.ftId===r.id&&p.fileId;}).sort(function(a,b){return Number(a.uploadedAt)-Number(b.uploadedAt);})
  .map(function(p){return {fotoId:p.fotoId,kind:p.kind,url:'',uploadedAt:Number(p.uploadedAt)||0};});
 /* riwayat: rantai induk -> retest, agar pengetesan ulang tidak menimpa hasil lama */
 var chain=[r],cur=r,guard=0;
 while(cur.parentId&&guard++<50){var p=rows.filter(function(x){return x.id===cur.parentId;})[0];if(!p)break;chain.unshift(p);cur=p;}
 var child=rows.filter(function(x){return x.parentId===r.id;})[0];
 var sib=r.batchId?rows.filter(function(x){return x.batchId===r.batchId&&x.id!==r.id;}).sort(function(a,b){return String(a.studio).localeCompare(String(b.studio),'id',{numeric:true});}).map(ftShape_):[];
 return {laporan:ftShape_(r),foto:fotos,riwayat:chain.map(ftShape_),retestId:child?child.id:'',batch:sib};
}

/* Satu LAPORAN testing = semua studio dalam satu hari/pengiriman (batchId sama) + retest-nya.
   Satu panggilan mengembalikan seluruh studio dan foto-fotonya, sehingga klien tidak perlu ft_get per studio. */
/* kunci laporan tiap baris: batchId akar; data lama tanpa batchId disatukan per hari (studio sama di hari sama = laporan baru).
   Aturan identik dengan ftGroup di modul-dcp.js. */
function ftKeys_(rows){
 var byId={};rows.forEach(function(x){byId[x.id]=x;});
 function root(r){var c=r,g=0;while(c.parentId&&byId[c.parentId]&&g++<50)c=byId[c.parentId];return c;}
 var bk={},legacy={},out={};
 rows.filter(function(r){return root(r)===r&&!r.batchId;})
  .sort(function(a,b){return (Number(a.createdAt)||0)-(Number(b.createdAt)||0)||String(a.id).localeCompare(String(b.id));})
  .forEach(function(r){var d=String(r.tanggal),bs=bk[d]=bk[d]||[],i=0;while(i<bs.length&&bs[i][String(r.studio)])i++;if(i===bs.length)bs.push({});bs[i][String(r.studio)]=1;legacy[r.id]='D'+d+(i?'-'+(i+1):'');});
 rows.forEach(function(r){var rt=root(r);out[r.id]=rt.batchId||legacy[rt.id];});
 return out;
}
function ftReportGet_(u,q){
 dcpCan_(u,'view');
 var rows=dcpRows_('FT_LAPORAN'),byId={};rows.forEach(function(x){byId[x.id]=x;});
 var seed=q.id?byId[q.id]:rows.filter(function(x){return q.batchId&&x.batchId===q.batchId;})[0];
 if(!seed)bad_('Laporan testing tidak ditemukan.');
 var keys=ftKeys_(rows),key=keys[seed.id];
 var grp=rows.filter(function(x){return keys[x.id]===key;});
 var ids={};grp.forEach(function(x){ids[x.id]=1;});
 var foto={};
 dcpRows_('FT_FOTO').filter(function(p){return p.fileId&&ids[p.ftId];}).sort(function(a,b){return Number(a.uploadedAt)-Number(b.uploadedAt);})
  .forEach(function(p){(foto[p.ftId]=foto[p.ftId]||[]).push({fotoId:p.fotoId,kind:p.kind,url:'',uploadedAt:Number(p.uploadedAt)||0});});
 grp.forEach(function(x){x._fc=(foto[x.id]||[]).length;});
 return {key:key,laporan:grp.sort(function(a,b){return String(a.studio).localeCompare(String(b.studio),'id',{numeric:true})||String(a.createdAt).localeCompare(String(b.createdAt));}).map(ftShape_),foto:foto};
}
function ftDel_(u,q){
 dcpCan_(u,'delete');
 var rows=dcpRows_('FT_LAPORAN'),r=rows.filter(function(x){return x.id===q.id;})[0];
 if(!r)bad_('Laporan testing tidak ditemukan.');
 if(rows.some(function(x){return x.parentId===r.id;}))bad_('Laporan ini punya riwayat pengetesan ulang. Hapus dari yang terbaru terlebih dahulu.');
 purgeFotos_('FT_FOTO','ftId',r.id);
 ss_().getSheetByName('FT_LAPORAN').deleteRow(r._row);
 log_(u,'HAPUS_TESTING','FT_LAPORAN',r.id,r.film+' '+r.status);
 return true;
}
function ftPhotoSave_(u,q){
 dcpCan_(u,'photo');
 var rows=dcpRows_('FT_LAPORAN'),r=rows.filter(function(x){return x.id===q.id;})[0];
 if(!r)bad_('Laporan testing tidak ditemukan.');
 if(u.role!=='ADMIN'&&String(r.createdByUserId)!==String(u.userId))bad_('Hanya Admin atau pembuat laporan yang dapat menambah foto.');
 var have=dcpRows_('FT_FOTO').filter(function(p){return p.ftId===r.id;}),n=have.length,add=0,fail=0;
 var fol=dcpSub_(dcpSub_(dcpRoot_(),'TESTING'),r.id);
 (Array.isArray(q.photos)?q.photos:[]).slice(0,DCP_BATCH).forEach(function(ph){try{
  if(n>=FT_MAXFOTO)throw new Error('melebihi batas '+FT_MAXFOTO+' foto');
  var b=Utilities.base64Decode(String(ph&&ph.b64||'').replace(/^data:[^,]*,/,'')),mime=imgKind_(b);
  if(!mime||b.length>DCP_MAXB)throw new Error('format/ukuran tidak valid');
  var kind=s_(ph&&ph.kind,40)||'Dokumentasi';
  var f=fol.createFile(Utilities.newBlob(b,mime,r.id+'-'+('0'+(n+1)).slice(-2)+'.'+EXT[mime]));
  n++;add++;
  var up=Date.now();
  dcpAdd_('FT_FOTO',{fotoId:uid_('FF'),ftId:r.id,kind:kind,fileId:f.getId(),uploadedAt:up,capturedAt:Number(ph&&ph.capturedAt)||up,url:''});
 }catch(e){fail++;Logger.log('foto testing gagal: '+e);}});
 log_(u,'FOTO_TESTING','FT_FOTO',r.id,'+'+add+' foto');
 return {added:add,failed:fail,total:n};
}
function ftPhotoDel_(u,q){
 dcpCan_(u,'photo');
 var p=dcpRows_('FT_FOTO').filter(function(x){return x.fotoId===q.fotoId;})[0];
 if(!p)bad_('Foto tidak ditemukan.');
 var L=dcpRows_('FT_LAPORAN').filter(function(x){return x.id===p.ftId;})[0];
 if(L&&u.role!=='ADMIN'&&String(L.createdByUserId)!==String(u.userId))bad_('Hanya Admin atau pembuat laporan yang dapat menghapus foto.');
 try{DriveApp.getFileById(String(p.fileId)).setTrashed(true);}catch(e){Logger.log('file foto tidak dipindahkan: '+e);}
 ss_().getSheetByName('FT_FOTO').deleteRow(p._row);
 log_(u,'HAPUS_FOTO_TESTING','FT_FOTO',p.ftId,p.kind);
 return true;
}
function ftDash_(u,q){
 dcpCan_(u,'view');
 var rows=dcpRows_('FT_LAPORAN').filter(function(r){return r.mode!=='HARIAN';}),today=dcpToday_();
 var hari=rows.filter(function(r){return String(r.tanggal)===today;});
 var films=all_('ATS_FILMS').filter(function(f){return f.status!=='Completed';});
 var testedToday={};hari.forEach(function(r){testedToday[String(r.film).toLowerCase()]=1;});
 var belum=films.filter(function(f){return !testedToday[String(f.title).toLowerCase()];}).map(function(f){return {title:f.title,status:f.status};});
 var perlu=rows.filter(function(r){return r.status==='FAILED'||r.status==='RETEST';}).filter(function(r){return !rows.some(function(x){return x.parentId===r.id;});})
  .slice(0,10).map(ftShape_);
 return {
  ceklisHariIni:dcpRows_('FT_LAPORAN').filter(function(r){return r.mode==='HARIAN'&&dcpDay_(r.tanggal)===today;}).length,
  hariIni:hari.length,passed:hari.filter(function(r){return r.status==='PASSED';}).length,
  failed:hari.filter(function(r){return r.status==='FAILED';}).length,retest:hari.filter(function(r){return r.status==='RETEST';}).length,
  belumDites:belum,perluTindakan:perlu,
  terbaru:rows.slice().sort(function(a,b){return (b.createdAt||0)-(a.createdAt||0);}).slice(0,8).map(ftShape_)
 };
}

/* baca foto privat: hanya lewat aksi ini (izin view). Tidak ada tautan publik. */
/* ---------- verifikasi DCP (izin: verify) ---------- */
function dcpVerify_(u,q){
 dcpAdmin_(u);
 dcpCan_(u,'verify');
 var rows=dcpRows_('DCP_TRANSAKSI'),r=rows.filter(function(x){return x.id===q.id;})[0];
 if(!r)bad_('Transaksi tidak ditemukan.');
 if(r.verifiedAt)bad_('DCP ini sudah diverifikasi oleh '+r.verifiedBy+'.');
 var kel=dcpParseJson_(r.kelengkapan,[]);
 if(dcpIncomplete_(kel).length&&!String(r.keterangan||'').trim())bad_('Tidak dapat diverifikasi: komponen tidak lengkap tanpa keterangan.');
 dcpUpd_('DCP_TRANSAKSI',r._row,{verifiedBy:u.fullName,verifiedAt:Date.now()});
 log_(u,'VERIFIKASI_DCP','DCP_TRANSAKSI',r.id,r.jenis+' '+r.film);
 return true;
}

/* ---------- izin pengguna (hanya Admin) ---------- */
function permSave_(u,q){
 adm_(u);
 var uid=s_(q.userId,60),users=all_('USERS'),target=users.filter(function(x){return x.userId===uid;})[0];
 if(!target)bad_('Pengguna tidak ditemukan.');
 if(target.role==='ADMIN')bad_('Admin selalu memiliki semua izin.');
 if(target.role!=='USER')bad_('Izin modul ini hanya untuk role Teknisi (USER).');
 var list=(Array.isArray(q.perms)?q.perms:[]).map(function(k){return String(k);}).filter(function(k){return DCP_PERM_ALL.indexOf(k)>=0;});
 if(list.indexOf('view')<0&&list.length)bad_('Izin "view" wajib jika izin lain diberikan.');
 var row=dcpRows_('USER_PERMS').filter(function(r){return String(r.userId)===uid;})[0];
 var o={userId:uid,perms:JSON.stringify(list),updatedAt:Date.now(),updatedBy:u.fullName};
 if(row)dcpUpd_('USER_PERMS',row._row,o);else dcpAdd_('USER_PERMS',o);
 log_(u,'UBAH_IZIN_DCP','USER_PERMS',uid,list.join(', ')||'(tanpa izin)');
 return list;
}
function permGet_(u,q){
 adm_(u);
 var uid=s_(q.userId,60),t=all_('USERS').filter(function(x){return x.userId===uid;})[0];
 if(!t)bad_('Pengguna tidak ditemukan.');
 var list=t.role==='ADMIN'?DCP_PERM_ALL.slice():(dcpPermOf_(uid)||DCP_PERM_DEF.USER);
 return {userId:uid,role:t.role,perms:list,custom:dcpPermOf_(uid)!==null,all:DCP_PERM_ALL};
}

/* ---------- job harian: DCP keluar yang belum kembali lebih dari N hari ---------- */
function dcpCekOverdueJob(){
 var lock=LockService.getScriptLock();lock.waitLock(20000);
 try{return dcpCekOverdueRun_();}finally{lock.releaseLock();}
}
function dcpCekOverdueRun_(){
 dcpSheets_();
 var days=Number(P_().getProperty('DCP_OVERDUE_HARI'))||7;
 var rows=dcpRows_('DCP_TRANSAKSI'),linked=dcpLinkedMap_(rows);
 var cutoff=Utilities.formatDate(new Date(Date.now()-days*864e5),TZ,'yyyy-MM-dd'),sent=0;
 rows.forEach(function(r){
  if(r.jenis!=='KELUAR'||r.modul==='PEL'||linked[String(r.id)]||r.overdueNotifiedAt)return;
  if(String(r.tanggal)>=cutoff)return;
  try{
   osSend_(notifAudience_(['ADMIN','USER'],''),'DCP belum kembali',r.film+' · '+r.kodeDisk+' · keluar '+r.tanggal,{dcpId:r.id});
   dcpUpd_('DCP_TRANSAKSI',r._row,{overdueNotifiedAt:Date.now()});
   sent++;
  }catch(e){Logger.log('notifikasi DCP gagal untuk '+r.id+': '+e);}
 });
 Logger.log('DCP terlambat diberitahu: '+sent);
 return sent;
}
/* jalankan sekali dari editor Apps Script (atau dari aksi Admin) untuk memasang trigger harian pukul 09.00 */
function aktifkanDcpOverdueTrigger(){
 ScriptApp.getProjectTriggers().forEach(function(t){if(t.getHandlerFunction()==='dcpCekOverdueJob')ScriptApp.deleteTrigger(t);});
 ScriptApp.newTrigger('dcpCekOverdueJob').timeBased().everyDays(1).atHour(9).create();
 return true;
}

/* ---------- daftar aksi (dipanggil dari route_ bila aksi tidak ada di H) ---------- */

/* ===== MODUL PELAPORAN DCP (aditif) =====
 * Satu kode disk = satu DCP. Riwayat Masuk/Keluar disimpan di DCP_TRANSAKSI yang sama (kolom modul='PEL').
 * TANPA validasi manual: DCP Masuk dianggap Lengkap bila satu laporan (batchId) punya minimal 1 foto; DCP Keluar hanya didata (foto opsional).
 * Status dihitung di klien dari data foto, sehingga foto tambahan tetap boleh diunggah kapan saja.
 * verifiedBy/verifiedAt hanya tersisa untuk data lama yang sudah divalidasi (tetap terkunci sampai validasi dibatalkan dengan alasan).
 * Setiap aksi dicatat di AUDIT_LOG. Tanggal dan jam diisi otomatis oleh server (zona Asia/Jakarta).
 */
var PEL_OUTLET = ['Sukaramai Trade Center XXI','Pekanbaru Xchange XXI','Ciputra Seraya XXI','SKA XXI'];
/* alur DCP: datang (dari XPDC / outlet XXI lain) > STC XXI > outlet XXI lain. DCP Masuk selalu dicatat di STC. */
var PEL_STC='Sukaramai Trade Center XXI',PEL_PUSAT='Kantor Pusat';
function pelOutlet_(v,lain){
 var o=s_(v,120);
 if(o==='Lainnya'){o=s_(lain,120);if(!o)bad_('Isi nama outlet lain.');return o;}
 if(PEL_OUTLET.indexOf(o)<0)bad_('Pilih outlet dari daftar atau pilih Lainnya.');
 return o;
}
/* keadaan satu kode disk dari riwayat (diurutkan tanggal, jam, lalu waktu input) */
/* tanggal/jam dari sheet bisa berupa objek Date; normalkan ke zona Asia/Jakarta */
function dcpIsDate_(v){return Object.prototype.toString.call(v)==='[object Date]';}
function dcpDay_(v){return dcpIsDate_(v)?Utilities.formatDate(v,TZ,'yyyy-MM-dd'):String(v||'');}
function dcpHm_(v){return dcpIsDate_(v)?Utilities.formatDate(v,TZ,'HH:mm'):String(v||'');}
function pelState_(rows,kd){
 var hist=rows.filter(function(r){return String(r.kodeDisk)===kd;}).sort(function(a,b){
  var ka=dcpDay_(a.tanggal)+' '+dcpHm_(a.jam),kb=dcpDay_(b.tanggal)+' '+dcpHm_(b.jam);
  return ka<kb?-1:ka>kb?1:Number(a.createdAt)-Number(b.createdAt);});
 if(!hist.length)return {none:true};
 var last=hist[hist.length-1],film=hist[0].film,ph=hist[0].ph;
 if(last.jenis==='MASUK'){return {none:false,atOutlet:true,outlet:last.modul==='PEL'?(last.outlet||''):PEL_STC,film:film,ph:ph};} /* baris DCP Management: kolom tujuan = teks bebas, DCP masuk selalu berada di STC */
 return {none:false,atOutlet:false,inTransitTo:last.outletTujuan||last.tujuan||'',film:film,ph:ph};
}
function pelList_(u,q){
 dcpCan_(u,'view');
 var rows=dcpRows_('DCP_TRANSAKSI').filter(function(r){return u.role==='ADMIN'||r.modul==='PEL';}),fotos=dcpRows_('DCP_FOTO'),fb={};
 fotos.forEach(function(p){(fb[p.transaksiId]=fb[p.transaksiId]||[]).push({fotoId:p.fotoId,kind:p.kind,capturedAt:Number(p.capturedAt)||0});});
 /* terbaru di atas: tanggal, jam, lalu waktu input */
 rows=rows.slice().sort(function(a,b){var x=dcpDay_(a.tanggal)+' '+dcpHm_(a.jam),y=dcpDay_(b.tanggal)+' '+dcpHm_(b.jam);return x<y?1:x>y?-1:(Number(b.createdAt)||0)-(Number(a.createdAt)||0);});
 return pelShareFoto_(rows.map(function(r){return {id:r.id,jenis:r.jenis,tanggal:dcpDay_(r.tanggal),jam:dcpHm_(r.jam),film:r.film,kodeDisk:r.kodeDisk,ph:r.ph,
  outlet:r.outlet||'',outletTujuan:r.outletTujuan||(r.jenis==='KELUAR'?r.tujuan:''),penerima:r.penerima||'',
  pelapor:r.createdBy||'',pelaporId:String(r.createdByUserId||''),batchId:r.batchId||'',keterangan:r.keterangan||'',
  kelengkapan:dcpParseJson_(r.kelengkapan,[]),jenisDcp:r.jenisDcp||'Besar',asalDari:r.asalDari||'',namaPengirim:r.namaPengirim||'',verifiedBy:r.verifiedBy||'',verifiedAt:Number(r.verifiedAt)||0,alasanBatal:r.alasanBatal||'',batalBy:r.batalBy||'',batalAt:Number(r.batalAt)||0,
  modul:r.modul||'',createdAt:Number(r.createdAt)||0,tipeKeluar:r.tipeKeluar||(r.jenis==='KELUAR'&&r.modul==='PEL'?'PINDAH':''),noResi:r.noResi||'',tglKirim:dcpDay_(r.tglKirim)||'',foto:fb[r.id]||[]};}));
}
function pelShareFoto_(list){ /* bukti kirim diunggah sekali per pengiriman (di item pertama): ikut tampil di semua item pengiriman yang sama */
 var by={};list.forEach(function(r){if(r.jenis==='KELUAR'&&r.batchId&&r.foto.length&&!by[r.batchId])by[r.batchId]=r.foto;});
 list.forEach(function(r){if(r.jenis==='KELUAR'&&r.batchId&&!r.foto.length&&by[r.batchId]){r.foto=by[r.batchId];r.fotoBersama=true;}});
 return list;
}
function pelSave_(u,q){
 var id=s_(q.id,40),rows=dcpRows_('DCP_TRANSAKSI'),now=Date.now();
 var old=id?rows.filter(function(r){return r.id===id;})[0]:null;
 if(id&&!old)bad_('Transaksi tidak ditemukan.');
 if(old){
  dcpCan_(u,'edit');
  if(old.modul!=='PEL')return dcpSave_(u,q); /* baris DCP Management: aturan Management (hanya Admin) */
  if(old.verifiedAt&&u.role!=='ADMIN')bad_('Laporan sudah tervalidasi dan tidak dapat diubah. Batalkan validasi terlebih dahulu.');
  if(u.role!=='ADMIN'&&String(old.createdByUserId)!==String(u.userId))bad_('Hanya Admin atau pembuat laporan yang dapat mengubah.');
 }else dcpCan_(u,'create');
 var jenis=s_(q.jenis,10);
 if(DCP_JENIS.indexOf(jenis)<0)bad_('Jenis harus MASUK atau KELUAR.');
 if(old&&old.jenis!==jenis)bad_('Jenis transaksi tidak dapat diubah.');
 var kd=s_(q.kodeDisk,60);if(!kd)bad_('Kode disk wajib diisi.');
 var st=pelState_(rows.filter(function(r){return !old||r.id!==old.id;}),kd);
 var ket=s_(q.keterangan,2000),outlet,film,ph,tujuan='',penerima='',pengirim='',tipe='',noResi='',tglKirim='';
 var ukuran=s_(q.jenisDcp,10);if(DCP_UKURAN.indexOf(ukuran)<0)bad_('Pilih jenis DCP: Besar atau Kecil.');
 var kel=dcpCheckList_(q.kelengkapan,ukuran);
 if(jenis==='MASUK'){
  outlet=PEL_STC;
  penerima=s_(u.fullName||u.name||u.username,120);
  var asal=s_(q.asalDari,200),pengirim=s_(q.namaPengirim,120);
  if(!asal)bad_('Isi asal DCP (datang dari mana).');
  if(!pengirim)bad_('Isi nama pengirim DCP.');
  if(st.none){
   film=s_(q.film,200);ph=s_(q.ph,160);
   if(!film||!ph)bad_('Judul film dan PH wajib diisi untuk DCP baru.');
  }else{
   film=st.film;ph=st.ph;
   var fIn=s_(q.film,200);
   if(fIn&&fIn.toLowerCase()!==String(film).toLowerCase())bad_('Judul film tidak sesuai dengan data DCP ('+film+').');
   if(st.atOutlet)bad_('DCP ini sudah tercatat berada di '+st.outlet+'. Catat Keluar terlebih dahulu.');
  }
 }else{
  if(st.none)bad_('DCP belum pernah dicatat masuk. Catat Masuk terlebih dahulu.');
  if(!st.atOutlet)bad_('DCP belum berada di '+PEL_STC+'. Catat DCP Masuk terlebih dahulu.');
  film=st.film;ph=st.ph;
  outlet=pelOutlet_(q.outlet,q.outletLain);
  if(st.outlet&&outlet.toLowerCase()!==String(st.outlet).toLowerCase())bad_('Outlet asal harus sama dengan outlet DCP saat ini ('+st.outlet+').');
  tipe=s_(q.tipeKeluar,10)==='PUSAT'?'PUSAT':'PINDAH';
  if(tipe==='PUSAT'){
   tujuan=PEL_PUSAT;noResi=s_(q.noResi,60);tglKirim=dcpDay_(s_(q.tglKirim,10));
   if(!noResi)bad_('Nomor resi pengiriman ke Kantor Pusat wajib diisi.');
   if(!/^\d{4}-\d{2}-\d{2}$/.test(tglKirim)||isNaN(new Date(tglKirim+'T00:00:00Z').getTime()))bad_('Tanggal pengiriman wajib diisi dengan benar.');
   penerima=s_(q.penerima,120)||PEL_PUSAT;
  }else{
   tujuan=pelOutlet_(q.outletTujuan,q.outletTujuanLain);
   if(tujuan.toLowerCase()===outlet.toLowerCase())bad_('Outlet tujuan tidak boleh sama dengan outlet asal.');
   penerima=s_(q.penerima,120);if(!penerima)bad_('Nama penerima di outlet tujuan wajib diisi.');
  }
  pengirim=s_(q.namaPengirim,120);if(!pengirim)bad_('Isi nama pengirim DCP.');
 }
 var o={jenis:jenis,film:film,kodeDisk:kd,ph:ph,format:'',penyerah:'',penerima:penerima,tujuan:jenis==='KELUAR'?tujuan:outlet,
  keterangan:ket,status:'Diproses',kelengkapan:JSON.stringify(kel),jenisDcp:ukuran,asalDari:jenis==='MASUK'?asal:'',namaPengirim:pengirim,relasiId:'',modul:'PEL',outlet:outlet,outletTujuan:jenis==='KELUAR'?tujuan:'',tipeKeluar:tipe,noResi:noResi,tglKirim:tglKirim,
  updatedBy:u.fullName,updatedAt:now};
 if(old){
  if(!old.batchId&&q.batchId)o.batchId=s_(q.batchId,40);
  dcpUpd_('DCP_TRANSAKSI',old._row,o);
  log_(u,'EDIT_PEL','DCP_TRANSAKSI',old.id,jenis+' '+film+' ('+kd+') di '+outlet);
  return old.id;
 }
 var tz=Utilities.formatDate(new Date(now),TZ,'yyyy-MM-dd'),tj=Utilities.formatDate(new Date(now),TZ,'HH:mm');
 o.tanggal=tz;o.jam=tj;o.batchId=s_(q.batchId,40);
 o.id=dcpNext_('DCP-'+(jenis==='KELUAR'?'OUT':'IN')+'-'+Utilities.formatDate(new Date(now),TZ,'yyyy')+'-',rows.filter(function(r){return r.jenis===jenis;}),'id');
 o.createdBy=u.fullName;o.createdByUserId=u.userId;o.createdAt=now;
 dcpAdd_('DCP_TRANSAKSI',o);
 log_(u,'BUAT_PEL','DCP_TRANSAKSI',o.id,jenis+' '+film+' ('+kd+') PH '+ph+' di '+outlet+(jenis==='KELUAR'?' ke '+tujuan+(tipe==='PUSAT'?' resi '+noResi+' tgl kirim '+tglKirim:'')+' penerima '+penerima:''));
 return o.id;
}
function pelValidasi_(u,q){
 dcpCan_(u,'view');
 var rows=dcpRows_('DCP_TRANSAKSI'),r=rows.filter(function(x){return x.id===q.id;})[0];
 if(!r)bad_('Transaksi tidak ditemukan.');
 if(u.role!=='ADMIN'&&String(r.createdByUserId)!==String(u.userId))bad_('Hanya pembuat laporan atau Admin yang dapat memvalidasi atau membatalkan validasi.');
 var aksi=s_(q.aksi,10);
 if(aksi==='validasi'){
  if(r.jenis==='KELUAR')bad_('DCP keluar tidak memerlukan validasi.');
  if(r.verifiedAt)bad_('Laporan ini sudah tervalidasi.');
  var ids={};ids[r.id]=1;
  if(r.batchId)rows.forEach(function(x){if(String(x.batchId)===String(r.batchId))ids[x.id]=1;});
  var nf=dcpRows_('DCP_FOTO').filter(function(p){return ids[p.transaksiId]&&p.fileId;}).length;
  if(!nf)bad_('Tambahkan minimal satu foto dokumentasi.');
  dcpUpd_('DCP_TRANSAKSI',r._row,{verifiedBy:u.fullName,verifiedAt:Date.now()});
  log_(u,'VALIDASI_PEL','DCP_TRANSAKSI',r.id,r.jenis+' '+r.film+' ('+r.kodeDisk+') oleh '+u.fullName);
  return true;
 }
 if(aksi==='batal'){
  if(!r.verifiedAt)bad_('Laporan ini belum tervalidasi.');
  var alasan=s_(q.alasan,500);
  if(alasan.length<5)bad_('Isi alasan pembatalan validasi (minimal 5 karakter).');
  dcpUpd_('DCP_TRANSAKSI',r._row,{verifiedBy:'',verifiedAt:'',alasanBatal:alasan,batalBy:u.fullName,batalAt:Date.now()});
  log_(u,'BATAL_VALIDASI_PEL','DCP_TRANSAKSI',r.id,'Alasan: '+alasan);
  return true;
 }
 bad_('Aksi tidak dikenal.');
}


/* ===== AUDIT SHEET TIDAK TERPAKAI (tombol web: hanya melaporkan dan menandai merah, tidak menghapus) =====
   Registri sheet yang dipakai = sheetPakai_() di Code.gs (SH, MIR, PHS, dan DCP_SH). Satu sumber untuk web dan editor.
   Pembersihan lewat editor Apps Script (Code.gs): periksaSheetTidakTerpakai > sembunyikanSheetTidakTerpakai > hapusSheetTidakTerpakai. */
var SHEET_MARK='#dc2626';
function sheetsDipakai_(){var o=sheetPakai_(),r={};Object.keys(o).forEach(function(n){r[n]=DCP_SH.hasOwnProperty(n)?'Modul DCP / Film Testing':'Data aplikasi';});return r;}
function sheetAudit_(u,q){
 adm_(u);
 var ss=ss_(),used=sheetsDipakai_(),out=[],unused=[];
 ss.getSheets().forEach(function(s){
  var n=s.getName(),ok=used.hasOwnProperty(n),rows=Math.max(0,s.getLastRow()-1);
  out.push({sheet:n,dipakai:ok,alasan:ok?used[n]:'Tidak dipakai kode aplikasi',baris:rows});
  if(!ok)unused.push(s);
 });
 var marked=0;
 if(q&&q.tandai){
  unused.forEach(function(s){try{s.setTabColor(SHEET_MARK);marked++;}catch(e){Logger.log('tandai gagal '+s.getName()+': '+e);}});
  log_(u,'AUDIT_SHEET','SPREADSHEET','-','Ditandai merah: '+unused.map(function(s){return s.getName();}).join(', '));
 }
 /* sheet dipakai yang belum ada (belum pernah dibuka modulnya) dilaporkan terpisah, bukan sebagai sampah */
 var have={};ss.getSheets().forEach(function(s){have[s.getName()]=1;});
 var belumAda=Object.keys(used).filter(function(n){return !have[n];});
 return {sheets:out,tidakTerpakai:out.filter(function(x){return !x.dipakai;}).map(function(x){return x.sheet;}),belumDibuat:belumAda,ditandai:marked};
}
var DCP_H = {
 dcp_save:function(u,q){return dcpSave_(u,q);},
 dcp_batch_save:function(u,q){return dcpBatchSave_(u,q);},
 dcp_list:function(u,q){return dcpList_(u,q);},
 dcp_get:function(u,q){return dcpGet_(u,q);},
 dcp_del:function(u,q){return dcpDel_(u,q);},
 dcp_photo_save:function(u,q){return dcpPhotoSave_(u,q);},
 dcp_photo_del:function(u,q){return dcpPhotoDel_(u,q);},
 dcp_photo_get:function(u,q){return dcpPhotoGet_(u,q);},
 dcp_photo_batch:function(u,q){return dcpPhotoBatch_(u,q);},
 ft_photo_get:function(u,q){return ftPhotoGet_(u,q);},
 dcp_dash:function(u,q){return dcpDash_(u,q);},
 ft_save:function(u,q){return ftSave_(u,q);},
 ftd_list:function(u,q){return ftdList_(u,q);},
 ftd_save:function(u,q){return ftdSave_(u,q);},
 ft_batch_save:function(u,q){return ftBatchSave_(u,q);},
 ft_list:function(u,q){return ftList_(u,q);},
 ft_get:function(u,q){return ftGet_(u,q);},
 ft_report_get:function(u,q){return ftReportGet_(u,q);},
 sheet_audit:function(u,q){return sheetAudit_(u,q);},
 ft_del:function(u,q){return ftDel_(u,q);},
 ft_photo_save:function(u,q){return ftPhotoSave_(u,q);},
 ft_dash:function(u,q){return ftDash_(u,q);},
 ft_photo_del:function(u,q){return ftPhotoDel_(u,q);},
 dcp_verify_save:function(u,q){return dcpVerify_(u,q);},
 pel_list:function(u,q){return pelList_(u,q);},
 pel_save:function(u,q){return pelSave_(u,q);},
 pel_validasi:function(u,q){return pelValidasi_(u,q);},
 perm_get:function(u,q){return permGet_(u,q);},
 perm_me:function(u,q){return {perms:u.role==='ADMIN'?DCP_PERM_ALL.slice():(dcpPermOf_(u.userId)||DCP_PERM_DEF.USER)};},
 dcp_job_install:function(u,q){adm_(u);return aktifkanDcpOverdueTrigger();},
 perm_save:function(u,q){return permSave_(u,q);},
 perm_list:function(u,q){adm_(u);return dcpRows_('USER_PERMS').map(function(r){return {userId:r.userId,perms:dcpParseJson_(r.perms,[]),updatedBy:r.updatedBy,updatedAt:Number(r.updatedAt)||0};});},
 dcp_meta:function(u,q){dcpCan_(u,'view');return {komponen:DCP_KOMP,komponenStatus:DCP_KOMP_ST,formatDcp:DCP_FMT,statusDcp:DCP_STATUS,formatFt:FT_FMT,checklist:FT_GROUPS,checklistItem:ftItemIds_(),status:FT_STATUS};}
};
