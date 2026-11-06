'use strict';
const $=s=>document.querySelector(s),$$=s=>[...document.querySelectorAll(s)];
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const sleep=ms=>new Promise(r=>setTimeout(r,ms));
// key, label, header-detection regex (order matters: specific before generic)
const FIELDS=[['sku','SKU',/sku|barcode|باركود|كود|رمز/i],['mfrPrice','سعر المصنع',/manufactur|factory|cost|مصنع|تكلفة/i],['seoTitle','عنوان SEO',/seo/i],['images','روابط الصور',/links|gallery|additional|روابط/i],['image','الصورة الرئيسية',/image|img|صورة/i],['price','السعر (EGP)',/egp|price|سعر/i],['desc','الوصف',/desc|وصف/i],['cat','التصنيفات',/categor|تصنيف|قسم|فئة/i],['name','اسم المنتج العربي',/arabic|name|اسم/i]];
const GEN=[['title','العنوان'],['description','الوصف'],['bullets','النقاط الخمس'],['seoTitle','SEO Title'],['seoDescription','SEO Description'],['keywords','الكلمات المفتاحية'],['features','المميزات'],['brand','العلامة التجارية'],['category','التصنيف المقترح'],['missing','معلومات ناقصة']];
const S={wb:null,headers:[],rows:[],products:[]};
let running=false,stopFlag=false,dlgId=null;
const find=id=>S.products.find(p=>p.id===id);
const isDone=p=>!!(p.gen&&p.gen.title);
const arr=v=>Array.isArray(v)?v:(v?String(v).split('\n').map(x=>x.trim()).filter(Boolean):[]);

function toast(m,err){const t=$('#toast');t.textContent=m;t.className='show'+(err?' err':'');clearTimeout(toast.t);toast.t=setTimeout(()=>t.className='',err?6000:3000)}

// ---------- settings ----------
const LS=localStorage;
function lsGet(k,d=''){try{return LS.getItem(k)??d}catch{return d}}
function lsSet(k,v){try{LS.setItem(k,v)}catch{}}
$('#apiKey').value=lsGet('lahzh_key');
$('#model').value=lsGet('lahzh_model_v2','gemini-3.8-flash');
$('#batchSize').value=lsGet('lahzh_batch','5');
$('#rpm').value=lsGet('lahzh_rpm','8');
$('#useImg').checked=lsGet('lahzh_img')==='1';
$('#apiKey').addEventListener('change',e=>{lsSet('lahzh_key',e.target.value.trim());toast('تم حفظ المفتاح في هذا المتصفح فقط')});
$('#model').addEventListener('change',e=>lsSet('lahzh_model_v2',e.target.value.trim()));
$('#batchSize').addEventListener('change',e=>lsSet('lahzh_batch',e.target.value));
$('#rpm').addEventListener('change',e=>lsSet('lahzh_rpm',e.target.value));
$('#useImg').addEventListener('change',e=>lsSet('lahzh_img',e.target.checked?'1':'0'));
$('#showKey').onclick=()=>{const i=$('#apiKey');i.type=i.type==='password'?'text':'password';$('#showKey').textContent=i.type==='password'?'إظهار':'إخفاء'};
$('#clearKey').onclick=()=>{$('#apiKey').value='';try{LS.removeItem('lahzh_key')}catch{}toast('تم مسح المفتاح')};

// ---------- import ----------
$('#file').addEventListener('change',e=>{
  const f=e.target.files[0];if(!f)return;
  const rd=new FileReader();
  rd.onerror=()=>toast('تعذر قراءة الملف',true);
  rd.onload=()=>{try{
    S.wb=/\.csv$/i.test(f.name)?XLSX.read(rd.result,{type:'string'}):XLSX.read(rd.result,{type:'array'});
    const sh=$('#sheet');sh.innerHTML=S.wb.SheetNames.map(n=>`<option>${esc(n)}</option>`).join('');
    sh.hidden=S.wb.SheetNames.length<2;loadSheet(S.wb.SheetNames[0]);
  }catch(err){toast('الملف غير صالح: '+err.message,true)}};
  /\.csv$/i.test(f.name)?rd.readAsText(f,'UTF-8'):rd.readAsArrayBuffer(f);
});
$('#sheet').onchange=e=>loadSheet(e.target.value);
function loadSheet(name){
  S.rows=XLSX.utils.sheet_to_json(S.wb.Sheets[name],{defval:'',raw:false});
  if(!S.rows.length){toast('الورقة فارغة',true);return}
  const hs=[];S.rows.forEach(r=>Object.keys(r).forEach(k=>{if(!hs.includes(k))hs.push(k)}));
  S.headers=hs;
  const used=new Set(),map={};
  FIELDS.forEach(([k,,re])=>{const h=hs.find(x=>!used.has(x)&&re.test(x));if(h){map[k]=h;used.add(h)}});
  $('#mapping').innerHTML=FIELDS.map(([k,l])=>`<label>${l}<select data-k="${k}"><option value="">— بدون —</option>${hs.map(h=>`<option ${map[k]===h?'selected':''}>${esc(h)}</option>`).join('')}</select></label>`).join('');
  $('#mapBox').hidden=false;
  toast(`تم تحميل ${S.rows.length} صف`);
}
$('#applyMap').onclick=()=>{
  const m={};$$('#mapping select').forEach(s=>m[s.dataset.k]=s.value);
  if(!m.name&&!m.sku)return toast('اربط عمود الاسم أو SKU على الأقل',true);
  S.products=S.rows.map((r,i)=>({id:'p'+i,orig:r,sku:r[m.sku]??'',name:r[m.name]??'',price:r[m.price]??'',mfrPrice:r[m.mfrPrice]??'',image:r[m.image]??'',images:r[m.images]??'',desc:r[m.desc]??'',seoTitle:r[m.seoTitle]??'',cat:r[m.cat]??'',gen:null,status:'pending',sel:false,err:'',upload:null}));
  $('#prodCard').hidden=$('#expCard').hidden=false;buildExpCols();render();
  toast(`تم إنشاء ${S.products.length} منتج`);$('#prodCard').scrollIntoView({behavior:'smooth'});
};

// ---------- table ----------
function visible(){
  const q=$('#search').value.trim().toLowerCase(),f=$('#statusF').value;
  return S.products.filter(p=>{
    if(q&&![p.sku,p.name,p.gen?.title].some(v=>String(v||'').toLowerCase().includes(q)))return false;
    if(f==='missing')return isDone(p)&&arr(p.gen.missing).length>0;
    return !f||p.status===f;
  });
}
const imgSrc=p=>p.upload||String(p.image||'').trim().split(/[\s,;|]+/).find(u=>/^https?:/i.test(u))||'';
const badge=p=>({pending:'معلّق',processing:'جارٍ…',done:'مكتمل',failed:'فشل'}[p.status]);
function render(){
  const v=visible();
  $('#tbody').innerHTML=v.map(p=>{const s=imgSrc(p),miss=isDone(p)&&arr(p.gen.missing).length;
    return `<tr data-id="${p.id}"><td><input type="checkbox" class="sel" ${p.sel?'checked':''}></td>
    <td>${s?`<img class="thumb" src="${esc(s)}" referrerpolicy="no-referrer" loading="lazy" alt="" onerror="this.style.visibility='hidden'">`:''}</td>
    <td dir="ltr">${esc(p.sku)}</td><td dir="auto">${esc(p.name)}</td><td dir="ltr">${esc(p.price)}</td>
    <td><span class="b ${p.status}" title="${esc(p.err)}">${badge(p)}</span>${miss?` <span class="b warn" title="${esc(arr(p.gen.missing).join(' | '))}">ناقص ${miss}</span>`:''}</td>
    <td><input class="ttl" dir="ltr" value="${esc(p.gen?.title||'')}" placeholder="—"></td>
    <td><button class="ghost edit">تفاصيل</button></td></tr>`}).join('');
  stats();
}
function stats(){
  const n=S.products.length,d=S.products.filter(isDone).length,f=S.products.filter(p=>p.status==='failed').length;
  $('#stats').textContent=n?`المنتجات ${n} · مكتمل ${d} · فشل ${f} · معلّق ${n-d-f}`:'';
  try{lsSet('lahzh_products',JSON.stringify(S.products.map(p=>({...p,upload:null,status:p.status==='processing'?'pending':p.status}))))}catch{}
}
$('#tbody').addEventListener('change',e=>{if(e.target.classList.contains('sel'))find(e.target.closest('tr').dataset.id).sel=e.target.checked});
$('#tbody').addEventListener('input',e=>{if(e.target.classList.contains('ttl')){const p=find(e.target.closest('tr').dataset.id);p.gen=p.gen||{};p.gen.title=e.target.value;if(p.gen.title)p.status='done';stats()}});
$('#tbody').addEventListener('click',e=>{if(e.target.classList.contains('edit'))openDlg(e.target.closest('tr').dataset.id)});
$('#search').oninput=render;$('#statusF').onchange=render;
$('#selAll').onchange=e=>{visible().forEach(p=>p.sel=e.target.checked);render()};
$('#selVis').onclick=()=>{visible().forEach(p=>p.sel=true);render()};
$('#selMissing').onclick=()=>{S.products.forEach(p=>p.sel=!isDone(p));render()};
$('#selNone').onclick=()=>{S.products.forEach(p=>p.sel=false);render()};

// ---------- detail dialog ----------
function openDlg(id){
  const p=find(id),g=p.gen||{};dlgId=id;
  $('#dTitle').textContent=`${p.sku} — ${p.name}`;
  const s=imgSrc(p);$('#dImg').src=s;$('#dImg').style.visibility=s?'visible':'hidden';
  $('#fTitle').value=g.title||'';$('#fDesc').value=g.description||'';$('#fBul').value=arr(g.bullets).join('\n');
  $('#fSeoT').value=g.seoTitle||'';$('#fSeoD').value=g.seoDescription||'';$('#fKw').value=arr(g.keywords).join(', ');
  $('#fFeat').value=arr(g.features).join('\n');$('#fBrand').value=g.brand||'';$('#fCat').value=g.category||'';$('#fMiss').value=arr(g.missing).join('\n');
  $('#dlg').showModal();
}
$('#dClose').onclick=()=>$('#dlg').close();
$('#dSave').onclick=()=>{
  const p=find(dlgId);
  p.gen={title:$('#fTitle').value.trim(),description:$('#fDesc').value.trim(),bullets:arr($('#fBul').value).slice(0,5),seoTitle:$('#fSeoT').value.trim(),seoDescription:$('#fSeoD').value.trim(),keywords:$('#fKw').value.split(',').map(x=>x.trim()).filter(Boolean),features:arr($('#fFeat').value),brand:$('#fBrand').value.trim(),category:$('#fCat').value.trim(),missing:arr($('#fMiss').value)};
  if(p.gen.title){p.status='done';p.err=''}
  $('#dlg').close();render();toast('تم حفظ التعديلات');
};
$('#dUp').onchange=e=>{
  const f=e.target.files[0];if(!f)return;
  if(f.size>4e6)return toast('الصورة أكبر من 4MB',true);
  const rd=new FileReader();rd.onload=()=>{const p=find(dlgId);p.upload=rd.result;p._img=null;$('#dImg').src=rd.result;$('#dImg').style.visibility='visible';render()};rd.readAsDataURL(f);
};
$('#dGen').onclick=()=>{$('#dlg').close();run([dlgId])};

// ---------- Gemini ----------
const SYSTEM=`You are an e-commerce catalog specialist writing English marketplace listings (Amazon/Noon) for an Egyptian seller.
STRICT RULES:
- Never invent specifications (size, dimensions, material, capacity, wattage, count, color, compatibility, certifications, warranty, origin). Use only facts present in the input fields or clearly visible in a provided image.
- If key information is missing, list each missing item in "missing" (e.g. "Material not provided", "Size/dimensions not provided", "Brand not provided"). Never fill gaps with guesses.
- Do NOT translate the Arabic name literally. Write a natural commercial English title: [Brand if known] + product type + key known attributes. Max 150 characters.
- "description": 60-150 words, plain sentences, only known facts and safe general uses. No exaggerated or unverifiable claims.
- "bullets": exactly 5 short benefit/feature lines (each under 200 characters), each grounded in known information.
- "seoTitle": max 60 characters. "seoDescription": max 155 characters.
- "keywords": 8-15 relevant English search terms (array of strings).
- "features": array of known features only (may be empty).
- "brand": only if stated in the input or visible in an image; otherwise "".
- "category": a suggested marketplace category path in English.
Return ONLY a JSON array, one object per input product, with keys: id, title, description, bullets, seoTitle, seoDescription, keywords, features, brand, category, missing. Copy each id exactly.`;

async function imgPart(p){
  if(p._img!==undefined&&p._img!==null)return p._img;
  try{
    let blob;
    if(p.upload){blob=await (await fetch(p.upload)).blob()}
    else{const u=imgSrc(p);if(!u)return p._img=null;const r=await fetch(u,{mode:'cors'});if(!r.ok)return p._img=null;blob=await r.blob()}
    if(!/^image\//.test(blob.type)||blob.size>4e6)return p._img=null;
    const data=await new Promise((res,rej)=>{const r=new FileReader();r.onload=()=>res(r.result.split(',')[1]);r.onerror=rej;r.readAsDataURL(blob)});
    return p._img={mimeType:blob.type,data};
  }catch{return p._img=null}
}
async function callGemini(items,key,model,useImg){
  const input=items.map(p=>({id:p.id,sku:p.sku,arabic_name:p.name,existing_description:p.desc,existing_seo_title:p.seoTitle,existing_category:p.cat}));
  const parts=[{text:'Products (JSON):\n'+JSON.stringify(input)}];
  if(useImg)for(const p of items){const im=await imgPart(p);if(im){parts.push({text:`Image for product id ${p.id}:`});parts.push({inlineData:im})}}
  const body={systemInstruction:{parts:[{text:SYSTEM}]},contents:[{role:'user',parts}],generationConfig:{temperature:0.3,responseMimeType:'application/json'}};
  let lastErr;
  for(let a=0;a<4;a++){
    let r,d;
    try{r=await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent`,{method:'POST',headers:{'Content-Type':'application/json','x-goog-api-key':key},body:JSON.stringify(body)});d=await r.json()}
    catch(e){lastErr=new Error('تعذر الاتصال بالشبكة');await sleep(3000*(a+1));continue}
    if(r.ok){
      const txt=(d.candidates?.[0]?.content?.parts||[]).map(x=>x.text||'').join('').replace(/^```json|```$/g,'').trim();
      if(!txt)throw new Error(d.promptFeedback?.blockReason?'تم حظر الطلب: '+d.promptFeedback.blockReason:'رد فارغ من النموذج');
      const out=JSON.parse(txt);return Array.isArray(out)?out:[out];
    }
    const msg=d.error?.message||('HTTP '+r.status);lastErr=new Error(msg);lastErr.status=r.status;
    if(r.status===429||r.status===503){
      if(/per day|daily/i.test(msg))throw lastErr;
      const m=msg.match(/retry in ([\d.]+)s/i);await sleep((m?parseFloat(m[1])+1:15*(a+1))*1000);continue;
    }
    throw lastErr;
  }
  throw lastErr;
}
function progress(done,total){$('#prog').hidden=false;$('#progBar').style.width=(total?done/total*100:0)+'%';$('#progTxt').textContent=`${done} / ${total}`}
async function run(ids){
  if(running)return toast('هناك عملية قيد التنفيذ',true);
  const key=$('#apiKey').value.trim();if(!key)return toast('أدخل مفتاح Gemini API أولاً',true);
  lsSet('lahzh_key',key);
  const model=$('#model').value.trim()||'gemini-3.8-flash',useImg=$('#useImg').checked;
  const size=Math.min(20,Math.max(1,+$('#batchSize').value||5)),gap=60000/Math.min(60,Math.max(1,+$('#rpm').value||8));
  let items=ids.map(find).filter(Boolean);
  if(!$('#redo').checked&&ids.length>1)items=items.filter(p=>!isDone(p));
  if(!items.length)return toast('لا توجد منتجات تحتاج توليداً');
  running=true;stopFlag=false;$('#stop').hidden=false;$$('#prodCard button:not(#stop)').forEach(b=>b.disabled=true);
  let done=0;progress(0,items.length);
  for(let i=0;i<items.length&&!stopFlag;i+=size){
    const chunk=items.slice(i,i+size),t0=Date.now();
    chunk.forEach(p=>{p.status='processing';p.err=''});render();
    try{
      const res=await callGemini(chunk,key,model,useImg);
      chunk.forEach(p=>{
        const r=res.find(x=>String(x.id)===p.id)||(res.length===1&&chunk.length===1?res[0]:null);
        if(r&&r.title){p.gen={title:String(r.title).trim(),description:r.description||'',bullets:arr(r.bullets).slice(0,5),seoTitle:r.seoTitle||'',seoDescription:r.seoDescription||'',keywords:arr(r.keywords),features:arr(r.features),brand:r.brand||'',category:r.category||'',missing:arr(r.missing)};p.status='done'}
        else{p.status='failed';p.err='لم يرجع النموذج نتيجة لهذا المنتج'}
      });
    }catch(e){
      chunk.forEach(p=>{p.status='failed';p.err=e.message});
      if([400,401,403,404].includes(e.status)){toast('خطأ: '+e.message,true);stopFlag=true;
        }
      else toast('فشل جزء من الطلبات: '+e.message,true);
    }
    done+=chunk.length;progress(done,items.length);render();
    const wait=gap-(Date.now()-t0);if(wait>0&&i+size<items.length&&!stopFlag)await sleep(wait);
  }
  items.forEach(p=>{if(p.status==='processing')p.status='pending'});
  running=false;$('#stop').hidden=true;$$('#prodCard button:not(#stop)').forEach(b=>b.disabled=false);render();
  const ok=items.filter(p=>p.status==='done').length;toast(`انتهت المعالجة: نجح ${ok} من ${items.length}`);
}
$('#stop').onclick=()=>{stopFlag=true;toast('سيتوقف بعد الطلب الحالي')};
$('#genSel').onclick=()=>{const ids=S.products.filter(p=>p.sel).map(p=>p.id);ids.length?run(ids):toast('حدد منتجات أولاً',true)};
$('#genAll').onclick=()=>run(S.products.filter(p=>!isDone(p)).map(p=>p.id));
$('#retry').onclick=()=>{const ids=S.products.filter(p=>p.status==='failed').map(p=>p.id);ids.length?run(ids):toast('لا توجد منتجات فاشلة')};

// ---------- export ----------
function buildExpCols(){$('#expCols').innerHTML=GEN.map(([k,l])=>`<label><input type="checkbox" class="expcol" value="${k}" checked>${l}</label>`).join('')}
const fmt=(k,v)=>Array.isArray(v)?v.join(k==='keywords'?', ':k==='missing'?'; ':' | '):(v??'');
function exportList(){
  const s=$('#scope').value;
  const l=S.products.filter(p=>s==='all'||(s==='selected'?p.sel:isDone(p)));
  if(!l.length)toast('لا توجد منتجات مطابقة للتصدير',true);return l;
}
function save(rows,name){
  if(!rows.length)return;
  const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows),'Products');
  XLSX.writeFile(wb,name+'.xlsx');toast(`تم تصدير ${rows.length} منتج`);
}
$('#exFull').onclick=()=>{
  const cols=$$('.expcol:checked').map(c=>c.value);
  save(exportList().map(p=>{const o={...p.orig},g=p.gen||{};
    cols.forEach(k=>{if(k==='bullets')for(let i=0;i<5;i++)o['AI_Bullet_'+(i+1)]=arr(g.bullets)[i]||'';else o['AI_'+k]=fmt(k,g[k])});return o}),'LAHZH_products_full');
};
const urls=p=>[p.image,...String(p.images||'').split(/[\s,;|]+/)].map(u=>String(u||'').trim()).filter((u,i,a)=>/^https?:/i.test(u)&&a.indexOf(u)===i);
function template(kind){
  const l=exportList(),noBrand=l.filter(p=>!p.gen?.brand).length;
  save(l.map(p=>{const g=p.gen||{},b=arr(g.bullets),u=urls(p),o={};
    if(kind==='amazon'){
      Object.assign(o,{item_sku:p.sku,item_name:g.title||'',brand_name:g.brand||'',product_description:g.description||'',product_type:g.category||''});
      for(let i=0;i<5;i++)o['bullet_point'+(i+1)]=b[i]||'';
      Object.assign(o,{generic_keywords:arr(g.keywords).join(' '),standard_price:p.price,currency:'EGP',main_image_url:u[0]||''});
      for(let i=1;i<=8;i++)o['other_image_url'+i]=u[i]||'';
    }else{
      Object.assign(o,{partner_sku:p.sku,product_title:g.title||'',brand:g.brand||'',long_description:g.description||'',product_subtype:g.category||''});
      for(let i=0;i<5;i++)o['feature_bullet_'+(i+1)]=b[i]||'';
      Object.assign(o,{search_keywords:arr(g.keywords).join(', '),price:p.price,currency:'EGP'});
      for(let i=0;i<8;i++)o['image_url_'+(i+1)]=u[i]||'';
    }
    o.review_notes=arr(g.missing).join('; ');return o}),kind==='amazon'?'LAHZH_amazon_template':'LAHZH_noon_template');
  if(l.length&&noBrand)toast(`تنبيه: ${noBrand} منتج بدون علامة تجارية، راجعها قبل الرفع`,true);
}
$('#exAmz').onclick=()=>template('amazon');$('#exNoon').onclick=()=>template('noon');

// ---------- restore previous session ----------
try{
  const saved=JSON.parse(lsGet('lahzh_products','[]'));
  if(saved.length){S.products=saved;S.rows=saved.map(p=>p.orig);$('#prodCard').hidden=$('#expCard').hidden=false;buildExpCols();render();toast('تمت استعادة آخر جلسة عمل')}
}catch{}
