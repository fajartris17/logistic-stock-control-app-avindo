const $ = id => document.getElementById(id);
const fmtNumber = v => new Intl.NumberFormat('id-ID').format(Number(v || 0));
const fmtRupiah = v => 'Rp ' + new Intl.NumberFormat('id-ID').format(Number(v || 0));
const safe = v => String(v ?? '').replace(/[&<>"']/g, s => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[s]));

function showToast(msg, type='success'){
  const t=$('toast'); if(!t) return;
  t.textContent=msg; t.className='toast show '+type;
  setTimeout(()=>{t.className='toast';},3600);
}

function openPage(id){
  document.querySelectorAll('.page').forEach(p=>p.classList.remove('active'));
  document.querySelectorAll('.nav').forEach(m=>m.classList.remove('active'));
  $(id)?.classList.add('active');
  document.querySelector(`.nav[data-page="${id}"]`)?.classList.add('active');
  $('sidebar')?.classList.remove('show');
  if(id==='stokBarang') loadStock();
  if(id==='laporan') loadLaporan();
  if(id==='database') loadDatabaseInfo();
}
document.querySelectorAll('.nav').forEach(btn=>btn.addEventListener('click',()=>openPage(btn.dataset.page)));
document.querySelectorAll('[data-target]').forEach(btn=>btn.addEventListener('click',()=>openPage(btn.dataset.target)));
$('hamb')?.addEventListener('click', ()=>$('sidebar').classList.toggle('show'));

function buildCalendar(){
  const now=new Date();
  const monthNames=['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
  $('calendarMonth').textContent=`${monthNames[now.getMonth()]} ${now.getFullYear()}`;
  const days=['Sen','Sel','Rab','Kam','Jum','Sab','Min'];
  const cal=$('calendar'); cal.innerHTML='';
  days.forEach(d=>{const e=document.createElement('div');e.className='calCell head';e.textContent=d;cal.appendChild(e);});
  const first=new Date(now.getFullYear(),now.getMonth(),1);
  const total=new Date(now.getFullYear(),now.getMonth()+1,0).getDate();
  const start=(first.getDay()+6)%7;
  for(let i=0;i<start;i++){const e=document.createElement('div');e.className='calCell empty';cal.appendChild(e);}
  for(let d=1;d<=total;d++){const e=document.createElement('div');e.className='calCell'+(d===now.getDate()?' today':'');e.textContent=d;cal.appendChild(e);}
}

let calcValue='0';
function updateCalc(){ $('calcDisplay').value=calcValue; }
document.querySelectorAll('.calcKeys button').forEach(btn=>btn.addEventListener('click',()=>{
  const k=btn.dataset.key;
  if(k==='C') calcValue='0';
  else if(k==='←') calcValue=calcValue.length>1?calcValue.slice(0,-1):'0';
  else if(k==='='){ try{ calcValue=String(Function(`return (${calcValue.replace(/%/g,'/100')})`)()); }catch(e){ calcValue='Error'; } }
  else { if(calcValue==='0'||calcValue==='Error') calcValue=k; else calcValue+=k; }
  updateCalc();
}));

async function apiJson(url, options={}){
  const res = await fetch(url, options);
  const data = await res.json().catch(()=>({ok:false,message:'Response bukan JSON'}));
  if(!res.ok || data.ok===false) throw new Error(data.message || 'Request gagal');
  return data;
}

async function loadDashboard(){
  const data=await apiJson('/api/dashboard');
  $('todayDay').textContent=data.hari; $('todayDate').textContent=data.tanggal;
  $('statJenis').textContent=fmtNumber(data.total_jenis_barang);
  $('statMasuk').textContent=fmtNumber(data.total_barang_masuk);
  $('statKeluar').textContent=fmtNumber(data.total_barang_keluar);
  $('statStok').textContent=fmtNumber(data.total_stok_tersedia);
  $('statMinimum').textContent=fmtNumber(data.stok_minimum);
  $('statHabis').textContent=fmtNumber(data.stok_habis);
  $('statNilaiMasuk').textContent=data.total_nilai_masuk_fmt;
  $('statNilaiKeluar').textContent=data.total_nilai_keluar_fmt;
  const tb=$('recentStock'); tb.innerHTML='';
  if(!data.recent_stock||!data.recent_stock.length){tb.innerHTML='<tr><td colspan="4">Belum ada data stok.</td></tr>';return;}
  data.recent_stock.slice().reverse().forEach(r=>{
    tb.innerHTML+=`<tr><td>${safe(r['Code Number'])}</td><td>${safe(r['Nama Barang'])}</td><td>${fmtNumber(r['Stok Akhir'])}</td><td>${statusBadge(r['Status Stok'])}</td></tr>`;
  });
}

function statusBadge(v){
  const text=safe(v||'-');
  let cls='ok';
  if(text.includes('MINIMUM')) cls='warn';
  if(text.includes('HABIS')) cls='danger';
  return `<span class="status ${cls}">${text}</span>`;
}

function fileLink(path, name='Buka'){
  if(!path) return '-';
  const label = path.startsWith('/uploads/') ? name : path;
  return `<a class="lampiranTag" href="${safe(path)}" target="_blank">${safe(label)}</a>`;
}

async function loadStock(){
  const q=encodeURIComponent($('searchStok')?.value||'');
  const data=await apiJson('/api/stok?q='+q);
  const tb=$('stokTable'); tb.innerHTML='';
  if(!data.length){tb.innerHTML='<tr><td colspan="10">Belum ada data stok.</td></tr>';return;}
  data.forEach(r=>{
    tb.innerHTML+=`<tr>
      <td>${safe(r['Code Number'])}</td><td>${safe(r['Nama Barang'])}</td><td>${safe(r['Jenis Barang'])}</td><td>${safe(r['Satuan'])}</td>
      <td>${fmtNumber(r['Total QTY Masuk'])}</td><td>${fmtNumber(r['Total QTY Keluar'])}</td><td>${fmtNumber(r['Stok Akhir'])}</td>
      <td>${fmtRupiah(r['Estimasi Nilai Stok'])}</td><td>${statusBadge(r['Status Stok'])}</td><td>${fileLink(r['Foto Barang Terakhir'],'Foto')}</td>
    </tr>`;
  });
}

async function loadLaporan(){
  const params=new URLSearchParams();
  if($('filterStart')?.value) params.set('start',$('filterStart').value);
  if($('filterEnd')?.value) params.set('end',$('filterEnd').value);
  if($('searchLaporan')?.value) params.set('q',$('searchLaporan').value);
  const data=await apiJson('/api/laporan?'+params.toString());
  const tb=$('laporanTable'); tb.innerHTML='';
  if(!data.length){tb.innerHTML='<tr><td colspan="12">Belum ada laporan.</td></tr>';return;}
  data.slice().reverse().forEach(r=>{
    const lamp=[fileLink(r['Foto'],'Foto'), fileLink(r['Bukti'],'Bukti')].filter(x=>x!=='-').join('');
    tb.innerHTML+=`<tr>
      <td>${safe(r['Tanggal'])}</td><td>${safe(r['Jam'])}</td><td>${safe(r['Jenis Transaksi'])}</td><td>${safe(r['ID Transaksi'])}</td>
      <td>${safe(r['Code Number'])}</td><td>${safe(r['Nama Barang'])}</td><td>${fmtNumber(r['Qty'])}</td><td>${safe(r['Satuan'])}</td>
      <td>${fmtRupiah(r['Nilai'])}</td><td>${safe(r['Pihak'])}</td><td>${safe(r['Status Koreksi'])}</td><td>${lamp||'-'}</td>
    </tr>`;
  });
}

async function submitForm(form,url){
  try{
    const body=new FormData(form);
    const data=await apiJson(url,{method:'POST',body});
    showToast(data.message,'success');
    form.reset(); clearPreviews(form); setTodayDefaults(); await loadDashboard();
  }catch(e){ showToast(e.message,'error'); }
}

$('formMasuk')?.addEventListener('submit',e=>{e.preventDefault();submitForm(e.currentTarget,'/api/barang-masuk');});
$('formKeluar')?.addEventListener('submit',e=>{e.preventDefault();submitForm(e.currentTarget,'/api/barang-keluar');});
$('formOcr')?.addEventListener('submit',e=>{e.preventDefault();submitForm(e.currentTarget,'/api/nota-ocr');});

$('formKoreksi')?.addEventListener('submit',async e=>{
  e.preventDefault();
  try{
    const payload=Object.fromEntries(new FormData(e.currentTarget).entries());
    const data=await apiJson('/api/koreksi',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)});
    showToast(data.message,'success'); e.currentTarget.reset();
  }catch(err){showToast(err.message,'error');}
});

$('formEdit')?.addEventListener('submit',async e=>{
  e.preventDefault();
  const fd=new FormData(e.currentTarget);
  const txid=fd.get('id_transaksi');
  const updates={};
  const mapping={
    nama_barang:'Nama Barang',
    jenis_barang:'Jenis Barang',
    satuan:'Satuan',
    harga_satuan:'Harga Satuan',
    status_koreksi:'Status Koreksi',
    keterangan:'Keterangan'
  };
  for(const [k,h] of Object.entries(mapping)){ if(fd.get(k)) updates[h]=fd.get(k); }
  if(fd.get('qty')){
    updates['QTY Masuk']=Number(fd.get('qty'));
    updates['QTY Keluar']=Number(fd.get('qty'));
    updates['Total QTY']=Number(fd.get('qty'));
  }
  try{
    const data=await apiJson('/api/transaksi/'+encodeURIComponent(txid),{method:'PUT',headers:{'Content-Type':'application/json'},body:JSON.stringify({updates,user_edit:fd.get('user_edit'),keterangan_edit:fd.get('keterangan')})});
    showToast(data.message,'success'); e.currentTarget.reset(); await loadDashboard(); await loadLaporan();
  }catch(err){showToast(err.message,'error');}
});

$('refreshStok')?.addEventListener('click',()=>loadStock().catch(e=>showToast(e.message,'error')));
$('searchStok')?.addEventListener('input',()=>loadStock().catch(()=>{}));
$('refreshLaporan')?.addEventListener('click',()=>loadLaporan().catch(e=>showToast(e.message,'error')));
$('printLaporan')?.addEventListener('click',()=>window.print());

function setTodayDefaults(){
  const today=new Date().toISOString().slice(0,10);
  document.querySelectorAll('input[type="date"]').forEach(i=>{if(!i.value)i.value=today;});
}
function clearPreviews(form){form.querySelectorAll('.previewBox').forEach(box=>box.textContent=box.dataset.empty||'Preview file');}
function attachFilePreviews(){
  document.querySelectorAll('input[type="file"][data-preview]').forEach(input=>{
    input.addEventListener('change',()=>{
      const box=$(input.dataset.preview); if(!box)return;
      box.dataset.empty=box.dataset.empty||box.textContent;
      const file=input.files&&input.files[0];
      if(!file){box.textContent=box.dataset.empty;return;}
      if(file.type.startsWith('image/')){
        const url=URL.createObjectURL(file);
        box.innerHTML=`<img src="${url}" alt="preview"><small>${safe(file.name)}</small>`;
      }else{
        box.innerHTML=`<b>File dipilih</b><small>${safe(file.name)}</small>`;
      }
    });
  });
}

async function loadDatabaseInfo(){
  try{
    const data=await apiJson('/api/database-info');
    const el=$('dbStatusText');
    if(el) el.textContent=`Aktif: ${data.file_name} • ${data.sheet_count} sheet • ${data.file_size_kb} KB`;
    const list=$('sheetList');
    if(list) list.innerHTML=(data.sheets||[]).map(s=>`<span>${safe(s)}</span>`).join('');
  }catch(e){}
}

if('serviceWorker' in navigator){navigator.serviceWorker.register('/static/sw.js').catch(()=>{});}
buildCalendar(); attachFilePreviews(); setTodayDefaults();
loadDashboard().catch(e=>showToast(e.message,'error'));
loadDatabaseInfo();
