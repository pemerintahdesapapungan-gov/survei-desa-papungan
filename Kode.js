/** Survei Desa Papungan. Fungsi berakhiran _ hanya untuk server/editor. */
function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Survei Desa Papungan').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
function include_(name) { return HtmlService.createHtmlOutputFromFile(name).getContent(); }
function props_() { return PropertiesService.getScriptProperties(); }
function book_() {
  var id = props_().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('Aplikasi belum disiapkan. Hubungi admin.');
  return SpreadsheetApp.openById(id);
}
/** Jalankan dari editor saja sesudah mengisi ADMIN_PASSWORD pada Script Properties. */
function setup_() {
  var p = props_(), pw = p.getProperty('ADMIN_PASSWORD');
  if (!p.getProperty('ADMIN_HASH')) {
    if (!pw || pw.length < 12) throw new Error('Isi Script Property ADMIN_PASSWORD minimal 12 karakter.');
    setPassword_(pw); p.deleteProperty('ADMIN_PASSWORD');
  }
  var b = p.getProperty('SPREADSHEET_ID') ? book_() : SpreadsheetApp.getActiveSpreadsheet();
  if (!b) throw new Error('Buka Apps Script melalui Extensions / Ekstensi di Google Sheets.');
  p.setProperty('SPREADSHEET_ID', b.getId());
  var heads = {
    Config: ['Revisi','Waktu','Pertanyaan_JSON'],
    Jawaban_A: ['ID','Waktu_WIB','Nama','Alamat','Umur','Jenis_Kelamin','Anonim','Revisi','Bintang','Kritik','Saran','Jawaban_JSON','Request_ID'],
    Jawaban_B: ['ID','Waktu_WIB','Nama','Alamat','Umur','Jenis_Kelamin','Anonim','Revisi','Bintang','Kritik','Saran','Jawaban_JSON','Request_ID']
  };
  Object.keys(heads).forEach(function(n) {
    var s = b.getSheetByName(n) || b.insertSheet(n);
    if (!s.getLastRow()) { s.appendRow(heads[n]); s.setFrozenRows(1); s.getRange(1,1,1,heads[n].length).setBackground('#126a87').setFontColor('#ffffff').setFontWeight('bold'); }
  });
  var c = b.getSheetByName('Config');
  if (c.getLastRow() === 1) c.appendRow([1, iso_(), JSON.stringify(defaultQuestions_())]);
  return 'Siap. Deploy sebagai Web app. Jangan bagikan Google Sheets kepada responden.';
}
function iso_() { return Utilities.formatDate(new Date(), 'Asia/Jakarta', "yyyy-MM-dd'T'HH:mm:ssXXX"); }
function hash_(s) { return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,s,Utilities.Charset.UTF_8)); }
function setPassword_(pw) {
  var salt=Utilities.getUuid();
  props_().setProperties({ADMIN_SALT:salt, ADMIN_HASH:hash_(salt+pw), AUTH_VERSION:Utilities.getUuid()});
}
function config_() {
  var s=book_().getSheetByName('Config');
  if (!s || s.getLastRow()<2) throw new Error('Jalankan setup_ terlebih dahulu.');
  var r=s.getRange(s.getLastRow(),1,1,3).getValues()[0];
  return {revision:Number(r[0]),questions:JSON.parse(r[2])};
}
function getPublicConfig() {
  var c=config_();
  return {revision:c.revision,questions:c.questions.filter(function(q){return q.active;}).map(function(q){
    return {id:q.id,survey:q.survey,group:q.group,text:q.text,options:q.options.map(function(o){return {label:o.label};})};
  })};
}
function requireAdmin_(token) {
  if (typeof token!=='string' || token.length>100) throw new Error('Silakan login admin.');
  var value=CacheService.getScriptCache().get('admin:'+hash_(token));
  if (!value || value!==props_().getProperty('AUTH_VERSION')) throw new Error('Sesi admin berakhir. Silakan login kembali.');
}
function adminLogin(password) {
  var lock=LockService.getScriptLock();lock.waitLock(10000);
  try {
    var p=props_(), now=Date.now(), limit=JSON.parse(p.getProperty('LOGIN_LIMIT')||'{"start":0,"count":0}');
    if (now-limit.start>300000) limit={start:now,count:0};
    if (limit.count>=10) throw new Error('Terlalu banyak percobaan. Tunggu 5 menit.');
    if (!p.getProperty('ADMIN_HASH')) throw new Error('Admin belum disiapkan.');
    if (typeof password!=='string' || password.length>200 || hash_(p.getProperty('ADMIN_SALT')+password)!==p.getProperty('ADMIN_HASH')) {
      limit.count++;p.setProperty('LOGIN_LIMIT',JSON.stringify(limit));throw new Error('Kata sandi salah.');
    }
    p.deleteProperty('LOGIN_LIMIT');
    var token=Utilities.getUuid()+Utilities.getUuid();
    CacheService.getScriptCache().put('admin:'+hash_(token),p.getProperty('AUTH_VERSION'),3600);
    return {token:token};
  } finally {lock.releaseLock();}
}
function adminLogout(token) { requireAdmin_(token);CacheService.getScriptCache().remove('admin:'+hash_(token));return true; }
function adminChangePassword(token,password) {
  requireAdmin_(token);
  if (typeof password!=='string'||password.length<12||password.length>200) throw new Error('Kata sandi harus 12–200 karakter.');
  setPassword_(password);return true;
}
function text_(v,max) {
  if(v===undefined||v===null) return '';
  if(typeof v!=='string') throw new Error('Format teks tidak valid.');
  v=v.trim();if(v.length>max) throw new Error('Teks melebihi '+max+' karakter.');return v;
}
function safeCell_(s) { return /^[=+\-@\t\r]/.test(String(s)) ? "'"+s : s; }
function submitSurvey(data) {
  if(!data||!['A','B'].includes(data.survey)) throw new Error('Survei tidak valid.');
  if(typeof data.requestId!=='string'||!/^[a-zA-Z0-9-]{16,80}$/.test(data.requestId)) throw new Error('ID pengiriman tidak valid.');
  var lock=LockService.getScriptLock();lock.waitLock(20000);
  try {
    var s=book_().getSheetByName('Jawaban_'+data.survey);
    if(s.getLastRow()>1) {
      var found=s.getRange(2,13,s.getLastRow()-1,1).createTextFinder(data.requestId).matchEntireCell(true).findNext();
      if(found) return {id:s.getRange(found.getRow(),1).getValue(),duplicate:true};
    }
    var c=config_();if(data.revision!==c.revision) throw new Error('Pertanyaan diperbarui admin. Muat ulang halaman sebelum mengisi kembali.');
    var qs=c.questions.filter(function(q){return q.active&&q.survey===data.survey;});
    if(!qs.length) throw new Error('Survei belum memiliki pertanyaan aktif.');
    var input=data.answers;if(!input||typeof input!=='object'||Array.isArray(input)) throw new Error('Jawaban tidak valid.');
    var answers=qs.map(function(q){
      var ix=input[q.id];
      if(!Number.isInteger(ix)||ix<0||ix>=q.options.length) throw new Error('Lengkapi jawaban: '+q.text);
      var o=q.options[ix];
      return {id:q.id,group:q.group,text:q.text,options:q.options,index:ix,label:o.label,score:o.score};
    });
    var anon=data.anonymous===true, person=data.person||{};
    var name=anon?'':text_(person.name,120),address=anon?'':text_(person.address,300);
    var gender=anon?'':text_(person.gender,30),age=anon?'':text_(person.age,3);
    if(age&&(!/^\d+$/.test(age)||Number(age)<1||Number(age)>120)) throw new Error('Umur harus 1–120 atau dikosongkan.');
    if(gender&&!['Laki-laki','Perempuan','Tidak ingin menyebutkan'].includes(gender)) throw new Error('Jenis kelamin tidak valid.');
    var stars='',crit='',suggest='';
    if(data.survey==='B') {
      if(!Number.isInteger(data.stars)||data.stars<1||data.stars>5) throw new Error('Pilih penilaian 1–5 bintang.');
      stars=data.stars;crit=text_(data.criticism,2000);suggest=text_(data.suggestion,2000);
    }
    if(JSON.stringify(answers).length>45000) throw new Error('Ukuran jawaban terlalu besar. Hubungi admin untuk menyederhanakan pertanyaan.');
    var id='PAP-'+Utilities.getUuid();
    s.appendRow([id,iso_(),safeCell_(name),safeCell_(address),age,gender,anon||(!name&&!address&&!age&&!gender),c.revision,stars,safeCell_(crit),safeCell_(suggest),JSON.stringify(answers),data.requestId]);
    SpreadsheetApp.flush();return {id:id};
  } finally {lock.releaseLock();}
}
function adminGetQuestions(token) {requireAdmin_(token);return config_();}
function validateQuestions_(questions) {
  if(!Array.isArray(questions)||!questions.length||questions.length>100) throw new Error('Jumlah pertanyaan harus 1–100.');
  var seen={};
  var clean=questions.map(function(q){
    if(!q||!['A','B'].includes(q.survey)||typeof q.id!=='string'||!/^[A-Za-z0-9_-]{1,50}$/.test(q.id)||seen[q.id]) throw new Error('ID pertanyaan harus unik.');
    seen[q.id]=true;
    var text=text_(q.text,1000),group=text_(q.group,180);
    if(!text||!group||!Array.isArray(q.options)||q.options.length<2||q.options.length>8) throw new Error('Isi pertanyaan, kelompok, dan 2–8 pilihan jawaban.');
    var labels={};
    var options=q.options.map(function(o){
      var label=text_(o.label,200);
      if(!label||labels[label]) throw new Error('Label pilihan wajib diisi dan berbeda.');labels[label]=true;
      if(o.score!==null&&(!Number.isInteger(o.score)||o.score<1||o.score>5)) throw new Error('Skor harus 1–5 atau tidak dinilai.');
      return {label:label,score:o.score};
    });
    if(!options.some(function(o){return o.score!==null;})) throw new Error('Minimal satu pilihan harus memiliki skor.');
    return {id:q.id,survey:q.survey,group:group,text:text,options:options,active:q.active===true};
  });
  ['A','B'].forEach(function(s){if(!clean.some(function(q){return q.active&&q.survey===s;}))throw new Error('Survei '+s+' harus memiliki pertanyaan aktif.');});
  var json=JSON.stringify(clean);if(json.length>45000) throw new Error('Total konfigurasi terlalu panjang.');return clean;
}
function adminSaveQuestions(token,revision,questions) {
  requireAdmin_(token);var clean=validateQuestions_(questions),lock=LockService.getScriptLock();lock.waitLock(20000);
  try {
    var c=config_();if(c.revision!==revision)throw new Error('Konfigurasi berubah. Muat ulang editor.');
    // ID yang pernah ada tidak boleh dipindahkan ke survei lain.
    var old={};c.questions.forEach(function(q){old[q.id]=q.survey;});
    clean.forEach(function(q){if(old[q.id]&&old[q.id]!==q.survey)throw new Error('Jangan pindahkan ID ke survei lain. Buat pertanyaan baru.');});
    book_().getSheetByName('Config').appendRow([c.revision+1,iso_(),JSON.stringify(clean)]);
    return {revision:c.revision+1};
  } finally {lock.releaseLock();}
}
function dateFilter_(f) {
  f=f||{};var from=text_(f.from,10),to=text_(f.to,10);
  [from,to].forEach(function(d){if(d&&(!/^\d{4}-\d{2}-\d{2}$/.test(d)||isNaN(Date.parse(d))||new Date(d).toISOString().slice(0,10)!==d))throw new Error('Tanggal filter tidak valid.');});
  if(from&&to&&from>to)throw new Error('Tanggal awal melebihi tanggal akhir.');return {from:from,to:to};
}
function rows_(survey,filter) {
  var s=book_().getSheetByName('Jawaban_'+survey);if(s.getLastRow()<2)return [];
  return s.getRange(2,1,s.getLastRow()-1,13).getValues().filter(function(r){var day=String(r[1]).slice(0,10);return (!filter.from||day>=filter.from)&&(!filter.to||day<=filter.to);}).map(function(r){
    function raw(v){v=String(v);return /^'[=+\-@\t\r]/.test(v)?v.slice(1):v;}
    return {id:r[0],time:String(r[1]),name:raw(r[2]),address:raw(r[3]),age:r[4],gender:r[5],anonymous:r[6]===true,revision:Number(r[7]),stars:r[8],criticism:raw(r[9]),suggestion:raw(r[10]),answers:JSON.parse(r[11])};
  });
}
function mean_(values) {return values.length ? values.reduce(function(a,b){return a+b;},0)/values.length : null;}
function summarize_(rows) {
  var variants={},groups={},daily={},starCounts=[0,0,0,0,0],starValues=[],responseMeans=[];
  rows.forEach(function(r){
    var day=r.time.slice(0,10);daily[day]=(daily[day]||0)+1;
    if(Number(r.stars)>=1&&Number(r.stars)<=5){starValues.push(Number(r.stars));starCounts[Number(r.stars)-1]++;}
    var scored=[];
    r.answers.forEach(function(a){
      // Revisi dipisah agar perubahan redaksi/skor tidak mencampur hasil lama.
      var key=r.revision+':'+a.id;
      if(!variants[key])variants[key]={id:a.id,revision:r.revision,text:a.text,group:a.group,options:a.options,counts:a.options.map(function(){return 0;}),values:[],total:0,unknown:0};
      var q=variants[key];q.total++;q.counts[a.index]++;
      if(a.score===null){q.unknown++;return;}
      q.values.push(a.score);scored.push(a.score);
      var gkey=r.revision+':'+a.group;
      if(!groups[gkey])groups[gkey]={revision:r.revision,group:a.group,values:[]};groups[gkey].values.push(a.score);
    });
    var m=mean_(scored);if(m!==null)responseMeans.push(m);
  });
  function metrics(x) {x.mean=mean_(x.values);x.index=x.mean===null?null:(x.mean-1)/4*100;x.valid=x.values.length;delete x.values;return x;}
  var overall=mean_(responseMeans);
  return {count:rows.length,anonymous:rows.filter(function(r){return r.anonymous;}).length,mean:overall,index:overall===null?null:(overall-1)/4*100,
    questions:Object.keys(variants).map(function(k){return metrics(variants[k]);}),groups:Object.keys(groups).map(function(k){return metrics(groups[k]);}),daily:daily,
    starCounts:starCounts,starMean:mean_(starValues),starN:starValues.length,revisions:Array.from(new Set(rows.map(function(r){return r.revision;})))};
}
function adminGetReport(token,survey,filter,page) {
  requireAdmin_(token);if(!['A','B'].includes(survey))throw new Error('Survei tidak valid.');
  var rows=rows_(survey,dateFilter_(filter));rows.sort(function(a,b){return b.time.localeCompare(a.time);});
  page=Number.isInteger(page)&&page>=0?page:0;var size=20;
  return {summary:summarize_(rows),rows:rows.slice(page*size,(page+1)*size),page:page,pages:Math.ceil(rows.length/size)};
}
function adminExportCsv(token,survey,filter) {
  requireAdmin_(token);if(!['A','B'].includes(survey))throw new Error('Survei tidak valid.');
  var rows=rows_(survey,dateFilter_(filter)),out=[['ID','Waktu WIB','Nama','Alamat','Umur','Jenis kelamin','Anonim','Revisi','Bintang','Kritik','Saran','ID pertanyaan','Kelompok','Pertanyaan','Jawaban','Skor']];
  rows.forEach(function(r){r.answers.forEach(function(a){out.push([r.id,r.time,r.name,r.address,r.age,r.gender,r.anonymous?'Ya':'Tidak',r.revision,r.stars,r.criticism,r.suggestion,a.id,a.group,a.text,a.label,a.score===null?'Tidak dinilai':a.score]);});});
  return '\uFEFF'+out.map(function(row){return row.map(function(v){return '"'+safeCell_(String(v===null?'':v)).replace(/"/g,'""')+'"';}).join(',');}).join('\r\n');
}

function setupAdmin() {
  SpreadsheetApp.getUi();
  const hasil = setup_();
  console.log(hasil);
}