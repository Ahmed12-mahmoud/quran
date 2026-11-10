import React, {useEffect, useMemo, useRef, useState} from 'react';
import {createRoot} from 'react-dom/client';
import {Download, Search, BookOpen, ShieldCheck, ArrowLeft, FileText, Archive, Sparkles, X, Pencil, Save, Upload, Database, LogIn, LogOut, Cloud, LockKeyhole, Printer} from 'lucide-react';
import './styles.css';
import {supabase, supabaseConfigured} from './lib/supabase';
import {Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell, WidthType, ShadingType, VerticalAlign, BorderStyle} from 'docx';
import {zipSync} from 'fflate';
import html2pdf from 'html2pdf.js';
import html2canvas from 'html2canvas';
import {jsPDF} from 'jspdf';

const SURAH_NAMES=['الفاتحة','البقرة','آل عمران','النساء','المائدة','الأنعام','الأعراف','الأنفال','التوبة','يونس','هود','يوسف','الرعد','إبراهيم','الحجر','النحل','الإسراء','الكهف','مريم','طه','الأنبياء','الحج','المؤمنون','النور','الفرقان','الشعراء','النمل','القصص','العنكبوت','الروم','لقمان','السجدة','الأحزاب','سبأ','فاطر','يس','الصافات','ص','الزمر','غافر','فصلت','الشورى','الزخرف','الدخان','الجاثية','الأحقاف','محمد','الفتح','الحجرات','ق','الذاريات','الطور','النجم','القمر','الرحمن','الواقعة','الحديد','المجادلة','الحشر','الممتحنة','الصف','الجمعة','المنافقون','التغابن','الطلاق','التحريم','الملك','القلم','الحاقة','المعارج','نوح','الجن','المزمل','المدثر','القيامة','الإنسان','المرسلات','النبأ','النازعات','عبس','التكوير','الانفطار','المطففين','الانشقاق','البروج','الطارق','الأعلى','الغاشية','الفجر','البلد','الشمس','الليل','الضحى','الشرح','التين','العلق','القدر','البينة','الزلزلة','العاديات','القارعة','التكاثر','العصر','الهمزة','الفيل','قريش','الماعون','الكوثر','الكافرون','النصر','المسد','الإخلاص','الفلق','الناس'];
const surahName=n=>SURAH_NAMES[Number(n)-1]||`السورة ${n}`;

async function loadRangeCatalog(){
  try{
    const response=await fetch('/ranges/index.json',{cache:'no-store'});
    if(!response.ok) throw new Error(`HTTP ${response.status}`);
    const rows=await response.json();
    return Array.isArray(rows)?rows.map(row=>{
      const system=String(row.system||'').toLowerCase()==='supper'?'upper':String(row.system||'').toLowerCase();
      const file=String(row.file||`/ranges/${system}-${String(row.range_id).padStart(2,'0')}.json`).replace(/Supper/gi,'upper').replace(/\/+$/,'');
      return {...row,system,file,range_key:row.range_key||`${system}-${String(row.range_id).padStart(2,'0')}`};
    }):[];
  }catch(error){
    console.error('تعذر تحميل index.json:',error);
    return [];
  }
}
const rangeIndex = await loadRangeCatalog();
async function loadRouteCatalog(){
  try{
    const response=await fetch('/route-catalog.json',{cache:'no-store'});
    if(!response.ok) throw new Error(`HTTP ${response.status}`);
    return await response.json();
  }catch(error){
    console.error('تعذر تحميل كتالوج الفئات والمسارات:',error);
    return {summary:{routes:0,tests:0,models:0,models_per_test:25},routes:[]};
  }
}
const routeCatalog = await loadRouteCatalog();
const systemLabels={lower:'من الناس إلى الفاتحة',upper:'من الفاتحة إلى الناس'};
const EDIT_KEY = 'quranExamEdits';

function App(){
  const [data,setData]=useState([]); const [rangeLoading,setRangeLoading]=useState(false); const [rangeError,setRangeError]=useState('');
  const [query,setQuery]=useState(''); const [selected,setSelected]=useState(null); const [selectedSystem,setSelectedSystem]=useState('lower'); const [selectedRange,setSelectedRange]=useState(1); const [notice,setNotice]=useState('');
  const [catalogDirection,setCatalogDirection]=useState('from_nas'); const [catalogCategory,setCatalogCategory]=useState('all'); const [catalogPath,setCatalogPath]=useState('all'); const [catalogRoute,setCatalogRoute]=useState(null); const [catalogTest,setCatalogTest]=useState(null);
  const [edits,setEdits]=useState(()=>{try{return JSON.parse(localStorage.getItem(EDIT_KEY)||'{}')}catch{return {}}});
  const saveTimers=useRef({}); const originalTexts=useRef({});
  const [editMode,setEditMode]=useState(false); const [user,setUser]=useState(null); const [authOpen,setAuthOpen]=useState(false); const [authLoading,setAuthLoading]=useState(false); const [email,setEmail]=useState(''); const [password,setPassword]=useState('');
  // يجب إعادة بناء النماذج عند وصول ملف النطاق async، وليس عند تعديل الأسئلة فقط.
  const effective=useMemo(()=>data.map(m=>({...m,preview:m.preview.map(q=>edits[q.uid]?{...q,text:edits[q.uid]}:q)})),[data,edits]);
  const rangeOptions=useMemo(()=>rangeIndex.filter(x=>x.system===selectedSystem).map(x=>({id:Number(x.range_id),name:x.name,file:x.file,range_key:x.range_key})).sort((a,b)=>a.id-b.id),[selectedSystem]);
  const catalogRoutes=useMemo(()=>routeCatalog.routes.filter(r=>r.direction===catalogDirection),[catalogDirection]);
  const categoryOrder=['الفئة الأولى','الفئة الثانية','الفئة الثالثة','الفئة الرابعة','التأسيس'];
  const pathOrder=['المسار الأول','المسار الثاني','المسار الثالث','المسار الرابع','المسار الخامس','المسار السادس','مسار التأسيس'];
  const catalogCategories=useMemo(()=>categoryOrder.filter(c=>catalogRoutes.some(r=>r.category_name===c)),[catalogRoutes]);
  const catalogPaths=useMemo(()=>pathOrder.filter(p=>catalogRoutes.some(r=>(catalogCategory==='all'||r.category_name===catalogCategory)&&r.path===p)),[catalogRoutes,catalogCategory]);
  const visibleCatalogRoutes=useMemo(()=>catalogRoutes.filter(r=>(catalogCategory==='all'||r.category_name===catalogCategory)&&(catalogPath==='all'||r.path===catalogPath)),[catalogRoutes,catalogCategory,catalogPath]);
  const selectedCatalogRoute=useMemo(()=>catalogRoute?routeCatalog.routes.find(r=>r.route_number===Number(catalogRoute)):visibleCatalogRoutes[0],[catalogRoute,visibleCatalogRoutes]);
  const selectedCatalogTest=useMemo(()=>selectedCatalogRoute?.tests.find(t=>t.test===Number(catalogTest))||selectedCatalogRoute?.tests[0],[selectedCatalogRoute,catalogTest]);
  const currentEntry=useMemo(()=>rangeIndex.find(x=>x.system===selectedSystem&&Number(x.range_id)===Number(selectedRange)),[selectedRange,selectedSystem]);
  // البيانات محمّلة من الملف المختار نفسه؛ الاعتماد على النظام ورقم الاختبار
  // يمنع اختفاء النماذج إذا كانت قيمة range_key قديمة أو مختلفة في التخزين المؤقت.
  const filtered=useMemo(()=>{const needle=query.trim().toLowerCase();return effective.filter(m=>!needle||`${m.id} ${m.name} ${m.range_name} ${m.model_in_range} ${m.path_name||''}`.toLowerCase().includes(needle))},[effective,query]);
  const showNotice=(t)=>{setNotice(t); setTimeout(()=>setNotice(''),3200)};

  useEffect(()=>{setSelectedRange(1)},[selectedSystem]);
  useEffect(()=>{let active=true;setRangeLoading(true);setRangeError('');setSelected(null);setData([]);if(!selectedCatalogRoute||!selectedCatalogTest){setRangeLoading(false);return}const routeNo=String(selectedCatalogRoute.route_number).padStart(2,'0');const testNo=Number(selectedCatalogTest.test);const file=`/route-files/route-${routeNo}.json`;fetch(file,{cache:'no-store'}).then(r=>{if(!r.ok)throw new Error(`HTTP ${r.status}`);return r.json()}).then(async payload=>{const list=Array.isArray(payload)?payload:(Array.isArray(payload?.chunks)?(await Promise.all(payload.chunks.map(u=>fetch(u,{cache:'no-store'}).then(r=>{if(!r.ok)throw new Error(`HTTP ${r.status}`);return r.json()}))).then(parts=>parts.flat())):[]);const entry=list.find(x=>Number(x.test)===testNo);const rows=entry?.models||list.filter(x=>Number(x.test_number)===testNo);const normalized=(Array.isArray(rows)?rows:[]).map(m=>({...m,system:selectedCatalogRoute.direction==='from_nas'?'lower':'upper',range_id:selectedCatalogRoute.route_number,range_key:`route-${routeNo}-test-${String(testNo).padStart(3,'0')}`,test_name:selectedCatalogTest?.name||m.test_name,name:`${selectedCatalogTest?.name||m.test_name||m.name} — نموذج ${String(m.model_in_range).padStart(2,'0')}`,range_name:selectedCatalogTest?.name||m.range_name||selectedCatalogRoute.path}));if(active)setData(normalized)}).catch(e=>{if(active){setData([]);setRangeError(`${e.message} — تأكد من رفع ملف ${file}`)}}).finally(()=>{if(active)setRangeLoading(false)});return()=>{active=false}},[selectedCatalogRoute,selectedCatalogTest]);

  useEffect(()=>{
    if(!supabaseConfigured) return;
    let active=true;
    loadRemoteEdits();
    supabase.auth.getSession().then(async ({data:{session}})=>{if(active){setUser(session?.user||null)}});
    const {data:{subscription}}=supabase.auth.onAuthStateChange(async (_event,session)=>{if(active){setUser(session?.user||null)}});
    return ()=>{active=false; subscription.unsubscribe()};
  },[]);

  async function loadRemoteEdits(){
    if(!supabase) return;
    const {data:rows,error}=await supabase.from('question_edits').select('uid,text');
    if(error){showNotice(`تعذر تحميل تعديلات Supabase: ${error.message}`);return}
    const remote=Object.fromEntries((rows||[]).map(row=>[row.uid,row.text]));
    setEdits(remote); localStorage.setItem(EDIT_KEY,JSON.stringify(remote));
  }
  const download=(url,label)=>{const a=document.createElement('a');a.href=url;a.download='';document.body.appendChild(a);a.click();a.remove();showNotice(`بدأ تنزيل ${label}`)};
  const modelTitle=(model)=>model.test_name||model.range_name||String(model.name||'').replace(/\s+—\s+نموذج\s+\d+\s*$/,'');
  const modelContext=(model)=>[model.direction_name,model.category_name,model.path_name,modelTitle(model)].filter(Boolean).join(' — ');
  const printMarkup=(model,root)=>{const esc=s=>String(s).replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));const groups=[];for(const q of model.preview){let g=groups.find(x=>x.id===q.part_id);if(!g){g={id:q.part_id,name:q.part_name,rows:[]};groups.push(g)}g.rows.push(q)}let n=0;root.innerHTML=`<section class="print-cover"><div class="document-header">إتحاف الأخيار في الحفظ بالأعشار</div><div class="direct-print-sub">${esc(modelContext(model))}</div><h1>نموذج رقم ${esc(model.model_in_range||model.id)}</h1><div class="print-field">اسم الطالب: ................................................</div><div class="print-field">الدرجة: ................................</div></section>${groups.map(g=>`<section class="print-unit"><h2>${esc(g.name)}</h2>${g.rows.map(q=>`<article><div class="question-number">سؤال ${++n}</div><div class="direct-print-meta">${esc(q.kind)} — سورة ${surahName(q.surah)}، آية ${q.ayah} إلى سورة ${surahName(q.end_surah)}، آية ${q.end_ayah}</div><div class="direct-print-text">${esc(q.text)}</div></article>`).join('')}</section>`).join('')}`};
  const printModel=(model)=>{let root=document.getElementById('direct-print-root');if(!root){root=document.createElement('div');root.id='direct-print-root';document.body.appendChild(root)}printMarkup(model,root);root.style.display='block';const previousTitle=document.title;document.title=model.name;const cleanup=()=>{window.removeEventListener('afterprint',cleanup);setTimeout(()=>{root.remove();document.title=previousTitle},600)};window.addEventListener('afterprint',cleanup);requestAnimationFrame(()=>requestAnimationFrame(()=>setTimeout(()=>window.print(),350)))};
  const buildWordBlob=async(model)=>{
    const rows=model.preview;
    const cell=(text,opts={})=>new TableCell({shading:opts.fill?{fill:opts.fill,type:ShadingType.CLEAR}:undefined,verticalAlign:VerticalAlign.CENTER,margins:{top:130,bottom:130,left:170,right:170},borders:{top:{style:BorderStyle.SINGLE,size:4,color:opts.border||'D8E2DB'},bottom:{style:BorderStyle.SINGLE,size:4,color:opts.border||'D8E2DB'},left:{style:BorderStyle.SINGLE,size:4,color:opts.border||'D8E2DB'},right:{style:BorderStyle.SINGLE,size:16,color:opts.border||'4F8566'}},children:[new Paragraph({bidirectional:true,alignment:opts.center?AlignmentType.CENTER:AlignmentType.RIGHT,spacing:{after:0},children:[new TextRun({text:String(text),bold:!!opts.bold,color:opts.color||'263238',font:{name:'Traditional Arabic',eastAsia:'Traditional Arabic'},size:opts.size||21})]})]});
    const title=(text,size=30,color='164A3B')=>new Paragraph({alignment:AlignmentType.CENTER,bidirectional:true,spacing:{after:180},children:[new TextRun({text,bold:true,color,font:{name:'Traditional Arabic',eastAsia:'Traditional Arabic'},size})]});
    const coverTitle=(text,size=48)=>new Paragraph({alignment:AlignmentType.CENTER,bidirectional:true,spacing:{after:220},border:{bottom:{style:BorderStyle.SINGLE,size:8,color:'B79A55',space:7}},children:[new TextRun({text,bold:true,color:'164A3B',font:{name:'Traditional Arabic',eastAsia:'Traditional Arabic'},size})]});
    const modelNumber=String(model.model_in_range||model.id||'').match(/\d+$/)?.[0]?.padStart(2,'0')||'01';
    const children=[new Paragraph({spacing:{before:2200},children:[]}),coverTitle('إتحاف الأخيار في الحفظ بالأعشار',48),title(modelContext(model),56),title(`نموذج رقم ${modelNumber}`,48),title('اسم الطالب: ................................................',42),title('الدرجة: ................................',42),new Paragraph({pageBreakBefore:true})];
    let questionNo=0;
    for(const partId of [...new Set(rows.map(q=>q.part_id))]){
      const jr=rows.filter(q=>q.part_id===partId); const partName=jr[0]?.part_name||`الوحدة ${partId}`;
      children.push(new Paragraph({alignment:AlignmentType.CENTER,bidirectional:true,keepNext:true,spacing:{before:260,after:120},shading:{fill:'F7F5EF'},children:[new TextRun({text:partName,bold:true,color:'164A3B',font:{name:'Traditional Arabic',eastAsia:'Traditional Arabic'},size:24})]}));
      const tableRows=[];
      for(const q of jr){
        questionNo++;
        const gold=q.kind==='احتياطي';
        const card=new TableCell({shading:{fill:gold?'FFF9E9':'FFFFFF',type:ShadingType.CLEAR},verticalAlign:VerticalAlign.CENTER,margins:{top:75,bottom:75,left:150,right:150},borders:{top:{style:BorderStyle.NONE,size:0,color:'FFFFFF'},bottom:{style:BorderStyle.NONE,size:0,color:'FFFFFF'},left:{style:BorderStyle.NONE,size:0,color:'FFFFFF'},right:{style:BorderStyle.SINGLE,size:18,color:gold?'B79A55':'4F8566'}},children:[new Paragraph({bidirectional:true,alignment:AlignmentType.RIGHT,spacing:{after:20},children:[new TextRun({text:`سؤال ${questionNo}`,bold:true,color:gold?'8A6A23':'286044',font:{name:'Arial',eastAsia:'Arial'},size:17})]}),new Paragraph({bidirectional:true,alignment:AlignmentType.RIGHT,spacing:{after:25},children:[new TextRun({text:`${q.kind}  •  سورة ${surahName(q.surah)}، آية ${q.ayah} إلى سورة ${surahName(q.end_surah)}، آية ${q.end_ayah}`,bold:true,color:gold?'8A6A23':'286044',font:{name:'Arial',eastAsia:'Arial'},size:16})]}),new Paragraph({bidirectional:true,alignment:AlignmentType.RIGHT,spacing:{after:0},children:[new TextRun({text:q.text,color:'17231D',font:{name:'Traditional Arabic',eastAsia:'Traditional Arabic'},size:23})]})]});
        tableRows.push(new TableRow({cantSplit:true,children:[card]}));
      }
      children.push(new Table({width:{size:100,type:WidthType.PERCENTAGE},columnWidths:[10200],borders:{top:{style:BorderStyle.NONE,size:0,color:'FFFFFF'},bottom:{style:BorderStyle.NONE,size:0,color:'FFFFFF'},left:{style:BorderStyle.NONE,size:0,color:'FFFFFF'},right:{style:BorderStyle.NONE,size:0,color:'FFFFFF'},insideHorizontal:{style:BorderStyle.NONE,size:0,color:'FFFFFF'},insideVertical:{style:BorderStyle.NONE,size:0,color:'FFFFFF'}},rows:tableRows}));
    }
    return await Packer.toBlob(new Document({styles:{default:{document:{run:{font:{name:'Arial',eastAsia:'Arial'},size:21},paragraph:{bidirectional:true}}}},sections:[{properties:{page:{size:{width:11906,height:16838},margin:{top:936,bottom:936,left:1008,right:1008}}},children}]}));
  };
  const downloadUpdatedWord=async(model)=>{
    const blob=await buildWordBlob(model);
    const filename=`${modelContext(model)}.docx`;
    try{
      const file=new File([blob],filename,{type:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
      if(navigator.share&&navigator.canShare&&navigator.canShare({files:[file]})){await navigator.share({files:[file],title:filename,text:'ملف Word'});showNotice('اختر حفظ في الملفات أو الطباعة من شاشة المشاركة');return}
    }catch(error){if(error?.name==='AbortError')return}
    const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=filename;a.target='_blank';a.rel='noopener';document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),60000);showNotice('تم تجهيز ملف Word؛ افتح التنزيلات أو اختر حفظ الملف');
  };
  const buildPdfBlob=async(model)=>{
    const root=document.createElement('div');root.id='pdf-export-root';root.className='pdf-export-root';root.style.cssText='position:absolute;left:0;top:0;width:794px;min-height:1123px;box-sizing:border-box;padding:38px 42px;background:#fff;color:#17231d;direction:rtl;font-family:Arial,Tahoma,sans-serif;line-height:1.8;z-index:2147483647;visibility:visible;opacity:1;';document.body.appendChild(root);printMarkup(model,root);
    await new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve)));
    const canvas=await html2canvas(root,{scale:2,backgroundColor:'#ffffff',useCORS:true,logging:false,scrollX:0,scrollY:0,windowWidth:794,windowHeight:Math.max(1123,root.scrollHeight)});
    const pdf=new jsPDF({unit:'mm',format:'a4',orientation:'portrait',compress:true});
    const pageW=210,pageH=297,margin=10,imgW=pageW-margin*2,usableH=pageH-margin*2;
    const sourcePageH=Math.floor(canvas.width*usableH/imgW);const scale=canvas.width/root.getBoundingClientRect().width;const articleStarts=[...root.querySelectorAll('article')].map(el=>Math.floor((el.offsetTop*scale))).filter(y=>y>0);const unitStarts=[...root.querySelectorAll('.print-unit')].map(el=>Math.floor((el.offsetTop*scale))).filter(y=>y>0);let sourceTop=0,page=0;
    while(sourceTop<canvas.height){if(page)pdf.addPage();let next=Math.min(sourceTop+sourcePageH,canvas.height);const nearUnit=unitStarts.filter(y=>y>sourceTop+80&&y<next&&next-y<700);if(next<canvas.height&&nearUnit.length){next=Math.min(...nearUnit)}else{const safeStarts=articleStarts.filter(y=>y>sourceTop+80&&y<=next);if(next<canvas.height&&safeStarts.length)next=Math.max(sourceTop+80,Math.max(...safeStarts));}const slice=document.createElement('canvas');slice.width=canvas.width;slice.height=Math.min(next-sourceTop,canvas.height-sourceTop);const ctx=slice.getContext('2d');ctx.fillStyle='#fff';ctx.fillRect(0,0,slice.width,slice.height);ctx.drawImage(canvas,0,sourceTop,canvas.width,slice.height,0,0,slice.width,slice.height);const sliceH=slice.height*imgW/slice.width;pdf.addImage(slice.toDataURL('image/jpeg',0.94),'JPEG',margin,margin,imgW,sliceH,undefined,'FAST');sourceTop+=slice.height;page++}
    root.remove();return pdf.output('blob');
  };
  const downloadUpdatedPdf=async(model)=>{
      try{const filename=`${modelContext(model)}.pdf`;const blob=await buildPdfBlob(model);const file=new File([blob],filename,{type:'application/pdf'});if(navigator.share&&navigator.canShare&&navigator.canShare({files:[file]})){try{await navigator.share({files:[file],title:filename,text:'ملف PDF للاختبار'});showNotice('اختر Google Drive أو حفظ في الملفات');return}catch(shareError){if(shareError?.name==='AbortError')return;console.warn('PDF share unavailable, falling back to download',shareError)}}const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=filename;a.rel='noopener';document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),60000);showNotice('تم تنزيل ملف PDF')}catch(error){console.error(error);document.getElementById('pdf-export-root')?.remove();showNotice('تعذر إنشاء PDF')}};
  const downloadArchiveBundle=async(kind)=>{
    try{
      if(!filtered.length){showNotice('لا توجد نماذج جاهزة لهذا الاختبار');return}
      const partSize=5; const parts={}; const totalParts=Math.ceil(filtered.length/partSize);
      for(let i=0;i<filtered.length;i+=partSize){
        const partNo=String(Math.floor(i/partSize)+1).padStart(2,'0'); const files={};
        showNotice(`جارٍ تجهيز ${kind==='word'?'Word':'PDF'} — الجزء ${partNo} من ${String(totalParts).padStart(2,'0')}`);
        for(const model of filtered.slice(i,i+partSize)){
          const blob=kind==='word'?await buildWordBlob(model):await buildPdfBlob(model);
          files[`${model.name}.${kind==='word'?'docx':'pdf'}`]=new Uint8Array(await blob.arrayBuffer());
        }
        parts[`الجزء-${partNo}-${kind==='word'?'Word':'PDF'}.zip`]=zipSync(files,{level:6});
      }
      const bundle= new Blob([zipSync(parts,{level:6})],{type:'application/zip'});
      const suffix=kind==='word'?'Word':'PDF'; const bundleName=`${modelContext(filtered[0])||'اختبار القرآن'} - ${suffix} - كل النماذج.zip`;
      const file=new File([bundle],bundleName,{type:'application/zip'});
      if(navigator.share&&navigator.canShare&&navigator.canShare({files:[file]})){
        try{await navigator.share({files:[file],title:bundleName,text:`حزمة ${suffix} كاملة مقسمة إلى أجزاء`});showNotice(`تم تجهيز حزمة ${suffix} الكاملة`);return}catch(error){if(error?.name==='AbortError')return;console.warn('archive share unavailable',error)}
      }
      const url=URL.createObjectURL(bundle);const a=document.createElement('a');a.href=url;a.download=bundleName;a.rel='noopener';document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),60000);showNotice(`تم تنزيل حزمة ${suffix} الكاملة — ${totalParts} أجزاء`);
    }catch(error){console.error(error);document.getElementById('pdf-export-root')?.remove();showNotice('تعذر تجهيز الحزمة المقسمة؛ حاول مرة أخرى')}
  };
  const downloadRangeZip=()=>downloadArchiveBundle('word');
  const downloadRangePdf=()=>downloadArchiveBundle('pdf');
  const updateQuestion=(uid,text)=>{setSelected(s=>{const old=s?.preview.find(q=>q.uid===uid)?.text;if(s&&originalTexts.current[uid]===undefined)originalTexts.current[uid]=old;return s?{...s,preview:s.preview.map(q=>q.uid===uid?{...q,text}:q)}:s});setEdits(prev=>{const next={...prev,[uid]:text};localStorage.setItem(EDIT_KEY,JSON.stringify(next));return next})};
  const saveQuestionWithReason=async(uid,text)=>{if(!user){showNotice('سجّل دخول المشرف أولًا');return}const reason=window.prompt('اكتب سبب التعديل — هذا الحقل إجباري:','');if(reason===null)return;if(!reason.trim()){showNotice('يجب كتابة سبب التعديل قبل الحفظ');return}await persistQuestion(uid,text,reason.trim(),originalTexts.current[uid]);delete originalTexts.current[uid]};
  const persistQuestion=async(uid,text,reason,beforeText)=>{
    if(!user||!supabase){showNotice('تم الحفظ على هذا الجهاز فقط؛ سجّل دخول المشرف للحفظ المركزي');return}
    const q=selected?.preview.find(item=>item.uid===uid); const now=new Date().toISOString();
    const latest=await supabase.from('question_edits').upsert({uid,text,updated_at:now,updated_by:user.id},{onConflict:'uid'});
    if(latest.error){showNotice(`تعذر حفظ السؤال: ${latest.error.message}`);return}
    const history=await supabase.from('question_edit_history').insert({uid,before_text:beforeText||'',after_text:text,reason,model_id:q?.model||selected?.id||null,model_name:selected?.name||'',range_key:q?.range_key||'',range_name:q?.range_name||selected?.range_name||'',surah:q?.surah||null,ayah:q?.ayah||null,updated_at:now,updated_by:user.id});
    if(history.error){showNotice(`تم حفظ السؤال لكن تعذر حفظ سجل السبب: ${history.error.message}`);return}
    showNotice('تم حفظ السؤال وسبب التعديل مركزيًا في Supabase');
  };
  const exportEdits=()=>{const blob=new Blob([JSON.stringify({instructions:'ضع هذا الملف بجانب generate_project.py ثم شغّل python generate_project.py',edits},null,2)],{type:'application/json'});const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download='quran-exam-edits.json';a.click();URL.revokeObjectURL(a.href);showNotice('تم تصدير التعديلات')};
  const importEdits=(e)=>{const f=e.target.files?.[0];if(!f)return;const r=new FileReader();r.onload=()=>{try{const x=JSON.parse(r.result);const next=x.edits||x;setEdits(next);localStorage.setItem(EDIT_KEY,JSON.stringify(next));showNotice('تم استيراد التعديلات محليًا؛ احفظ كل سؤال لتزامنه مع Supabase')}catch{showNotice('تعذر قراءة ملف التعديلات')}};r.readAsText(f)};
  const login=async(e)=>{e.preventDefault();if(!supabase){showNotice('بيانات Supabase غير مضافة إلى بيئة التشغيل');return}setAuthLoading(true);const {data:result,error}=await supabase.auth.signInWithPassword({email:email.trim().toLowerCase(),password});setAuthLoading(false);if(error){showNotice(`خطأ Supabase: ${error.message}`);return}setUser(result.user);setAuthOpen(false);setPassword('');await loadRemoteEdits();showNotice('تم تسجيل دخول المشرف')};
  const logout=async()=>{await supabase?.auth.signOut();setUser(null);setEditMode(false);showNotice('تم تسجيل الخروج')};
  const startEdit=()=>{if(!user){setAuthOpen(true);showNotice('سجّل دخول المشرف أولًا للتعديل');return}setEditMode(!editMode)};
  return <div className="app">
    <header className="topbar"><div className="brand"><div className="brand-mark">۞</div><div><strong>اختبارات القرآن الكريم</strong><span>الحفظ الشفوي • إصدار تعليمي منظم</span></div></div><div className="top-actions"><button className="admin-btn" onClick={()=>user?logout():setAuthOpen(true)}>{user?<><LogOut size={15}/> خروج المشرف</>:<><LockKeyhole size={15}/> دخول المشرف</>}</button><a className="nav-link" href="#models">النماذج <ArrowLeft size={16}/></a></div></header>
    <main>
      <section className="hero"><div className="hero-copy"><div className="eyebrow"><Sparkles size={16}/> أدوات عملية للمحفّظ والممتحن</div><h1>اختبر الحفظ بثقة،<br/><em>واجعل كل جزء حاضرًا.</em></h1><p>منظومة منظمة من 46 مسارًا و1263 اختبارًا، ينتج كل اختبار 25 نموذجًا، مع توزيع الأسئلة على شرائح 7 + 7 + 6 أوجه والاحتياطي في الفاصل الأوسط.</p><div className="hero-actions"><a className="primary" href="#models">استعرض المسارات <ArrowLeft size={18}/></a><button className="secondary" onClick={downloadRangeZip}><Archive size={18}/> تنزيل Word كامل للاختبار المحدد</button><button className="secondary" onClick={downloadRangePdf}><FileText size={18}/> تنزيل PDF كامل للاختبار المحدد</button></div></div><div className="hero-card"><div className="seal">۞</div><div className="hero-card-title">فئات ومسارات دقيقة</div><div className="hero-card-text">25 نموذجًا لكل اختبار<br/>من الفاتحة إلى الناس أو من الناس إلى الفاتحة</div><div className="hero-card-rule"/><div className="hero-card-foot"><span>31,575 نموذجًا</span><span>46 مسارًا</span></div></div></section>
      <section className="stats"><div><b>31,575</b><span>نموذج اختبار</span></div><div><b>1,263</b><span>اختبارًا</span></div><div><b>46</b><span>مسارًا</span></div><div><b>25</b><span>نموذجًا لكل اختبار</span></div></section>
      <section className="section-head" id="models"><div><div className="eyebrow">المكتبة الرقمية</div><h2>اختر الفئة والمسار والاختبار</h2><p>لكل اختبار 25 نموذجًا؛ ويظهر الجديد والقديم بعدد الأسئلة المعتمد لكل نظام.</p></div><div className="search"><Search size={18}/><input value={query} onChange={e=>setQuery(e.target.value)} placeholder="ابحث برقم النموذج أو الجزء…" aria-label="البحث في النماذج"/></div></section>
      <section className="route-catalog-panel"><div className="route-catalog-heading"><div><div className="eyebrow">خريطة الاختبارات الجديدة</div><h3>اختر الاتجاه ثم الفئة والمسار</h3><p>الكتالوج الصحيح يضم 4 فئات و6 مسارات و46 ملفًا، ويعرض اسم الاختبار قبل فتح النماذج.</p></div><span className="catalog-count">{routeCatalog.summary.routes} مسارًا • {routeCatalog.summary.tests} اختبارًا</span></div><div className="route-selectors"><label>الاتجاه<select value={catalogDirection} onChange={e=>{setCatalogDirection(e.target.value);setCatalogCategory('all');setCatalogPath('all');setCatalogRoute(null);setCatalogTest(null)}}><option value="from_nas">من الناس إلى الفاتحة</option><option value="from_fatiha">من الفاتحة إلى الناس</option></select></label><label>الفئة<select value={catalogCategory} onChange={e=>{setCatalogCategory(e.target.value);setCatalogPath('all');setCatalogRoute(null);setCatalogTest(null)}}><option value="all">كل الفئات</option>{catalogCategories.map(x=><option key={x} value={x}>{x}</option>)}</select></label><label>المسار<select value={catalogPath} onChange={e=>{setCatalogPath(e.target.value);setCatalogRoute(null);setCatalogTest(null)}}><option value="all">كل المسارات</option>{catalogPaths.map(x=><option key={x} value={x}>{x}</option>)}</select></label><label>المسار المحدد<select value={selectedCatalogRoute?.route_number||''} onChange={e=>setCatalogRoute(Number(e.target.value))}>{visibleCatalogRoutes.map(r=><option key={r.route_number} value={r.route_number}>{r.path} — {r.test_count} اختبارًا</option>)}</select></label><label>الاختبار<select value={selectedCatalogTest?.test||''} onChange={e=>setCatalogTest(Number(e.target.value))}>{(selectedCatalogRoute?.tests||[]).map(t=><option key={t.test} value={t.test}>{t.name||`اختبار ${String(t.test).padStart(2,'0')}`}</option>)}</select></label></div>{selectedCatalogRoute&&<div className="route-summary"><div><b>{selectedCatalogRoute.direction_name}</b><span>{selectedCatalogRoute.category_name} • {selectedCatalogRoute.path}</span></div><div><strong>{selectedCatalogRoute.test_count}</strong><span>اختبارًا</span></div><div><strong>{selectedCatalogRoute.model_count.toLocaleString('ar-EG')}</strong><span>نموذجًا</span></div><div><small>الاختبار المحدد</small><span>{selectedCatalogTest?.name||selectedCatalogRoute.tests[0]?.name||selectedCatalogRoute.tests[0]?.raw||'—'}</span></div></div>}</section>
      <div className="catalog-active-note">النماذج المعروضة مرتبطة مباشرة بالمسار والاختبار المحددين أعلاه.<span className="catalog-download-actions"><button className="text-btn" onClick={downloadRangeZip}>تنزيل Word كاملًا إلى Drive/الجهاز</button><button className="text-btn" onClick={downloadRangePdf}>تنزيل PDF كاملًا إلى Drive/الجهاز</button></span></div>
      {rangeLoading?<div className="empty"><BookOpen size={30}/><h3>جارٍ تحميل نماذج الاختبار…</h3><p>{systemLabels[selectedSystem]}</p></div>:rangeError?<div className="empty"><Search size={30}/><h3>تعذر تحميل نماذج هذا النطاق</h3><p>{rangeError} — الملف المطلوب: {selectedCatalogRoute?`/route-files/route-${String(selectedCatalogRoute.route_number).padStart(2,'0')}.json`:'غير محدد'}</p></div>:filtered.length?<div className="grid">{filtered.map(m=><article className="model-card" key={m.id}><div className="card-top"><span className="model-no">{String(m.model_in_range).padStart(2,'0')}</span><span className="ready"><ShieldCheck size={15}/> جاهز</span></div><h3>{m.name}</h3><div className="card-meta"><span><BookOpen size={16}/> {m.range_name}</span><span><FileText size={16}/> {m.questions} أسئلة</span></div><div className="card-breakdown"><span>أساسي <b>{m.basic}</b></span><span>احتياطي <b>{m.reserve}</b></span></div><div className="card-actions"><button className="primary small" onClick={()=>downloadUpdatedWord(m)}><Download size={16}/> تنزيل Word</button><button className="print-btn" onClick={()=>downloadUpdatedPdf(m)}><FileText size={16}/> تنزيل PDF</button><button className="print-btn" onClick={()=>printModel(m)}><Printer size={16}/> طباعة</button><button className="text-btn" onClick={()=>{setSelected(m);setEditMode(false)}}>عرض وتعديل <ArrowLeft size={15}/></button></div></article>)}</div>:<div className="empty"><Search size={26}/><h3>لا توجد نماذج في هذا النطاق</h3><p>الملف المحمّل: {selectedCatalogRoute?`/route-files/route-${String(selectedCatalogRoute.route_number).padStart(2,'0')}.json`:'غير محدد'}</p></div>}
      <section className="how"><div className="how-intro"><div className="eyebrow">لوحة المشرف</div><h2>تعديل آمن ومركزي</h2><p>بعد تسجيل دخول المشرف، تُحفظ تعديلات الأسئلة في Supabase وتظهر لكل الزوار.</p></div><div className="steps"><div><span>01</span><h3>دخول المشرف</h3><p>استخدم حساب Supabase الذي أنشأته للمشرف.</p></div><div><span>02</span><h3>عدّل واحفظ</h3><p>كل سؤال يُحفظ مركزيًا عند مغادرة خانة التعديل.</p></div><div><span>03</span><h3>نزّل Word</h3><p>اضغط تنزيل Word من داخل النموذج.</p></div></div></section>
      <section className="source-note"><ShieldCheck size={22}/><div><strong>حالة الربط</strong><p>{supabaseConfigured?'Supabase مهيأ — '+(user?'تم تسجيل دخول المشرف':'بانتظار دخول المشرف'):'Supabase غير مهيأ في بيئة التشغيل'}</p></div><a href="https://tanzil.net/download/" target="_blank">المصدر</a></section>
    </main><footer><span>© اختبارات القرآن الكريم</span><span>صُمّم للمؤسسات التعليمية وحلقات التحفيظ</span></footer>
    {selected&&<div className="modal-backdrop" onClick={()=>setSelected(null)}><div className="modal" onClick={e=>e.stopPropagation()}><button className="close" onClick={()=>setSelected(null)}><X/></button><div className="modal-heading"><span className="model-no">{String(selected.id).padStart(2,'0')}</span><div><div className="eyebrow">معاينة النموذج</div><h2>{selected.name}</h2><p>أسئلة موزعة على نطاق الأجزاء المحدد — {editMode?'وضع التعديل مفعل':'وضع القراءة'}</p></div></div><div className="modal-actions"><button className="secondary" onClick={startEdit}><Pencil size={16}/> {editMode?'إنهاء التعديل':user?'تعديل الأسئلة':'دخول للتعديل'}</button><button className="secondary" onClick={()=>printModel(selected)}><Printer size={16}/> طباعة النموذج</button><button className="secondary" onClick={()=>downloadUpdatedPdf(selected)}><FileText size={16}/> تنزيل PDF</button></div><div className="preview-list">{selected.preview.map((q,i)=><div className={`preview-row ${q.kind==='احتياطي'?'reserve':''}`} key={q.uid}><div><b>{q.part_name}</b><small>{q.kind} • سورة {surahName(q.surah)}، آية {q.ayah}</small></div>{editMode?<textarea value={q.text} onChange={e=>updateQuestion(q.uid,e.target.value)} onBlur={e=>saveQuestionWithReason(q.uid,e.target.value)} aria-label={`تعديل سؤال ${i+1}`}/>:<p>{q.text} …</p>}</div>)}</div><button className="primary" onClick={()=>downloadUpdatedWord(selected)}><Download size={17}/> تنزيل Word</button><p className="edit-hint">عدّل الأسئلة ثم اضغط خارج الخانة للحفظ، وبعدها استخدم زر Word أو PDF.</p></div></div>}
    {authOpen&&<div className="modal-backdrop" onClick={()=>setAuthOpen(false)}><form className="auth-modal" onClick={e=>e.stopPropagation()} onSubmit={login}><button type="button" className="close" onClick={()=>setAuthOpen(false)}><X/></button><div className="auth-icon"><LockKeyhole/></div><div className="eyebrow">صلاحية المشرف</div><h2>تسجيل دخول المشرف</h2><p>هذا الدخول مخصص لتعديل الأسئلة وحفظها مركزيًا.</p><label>البريد الإلكتروني<input type="email" required value={email} onChange={e=>setEmail(e.target.value)} autoComplete="username" autoCapitalize="none"/></label><label>كلمة المرور<input type="password" required value={password} onChange={e=>setPassword(e.target.value)} autoComplete="current-password"/></label><button className="primary" disabled={authLoading}>{authLoading?'جارٍ التحقق…':<><LogIn size={17}/> دخول آمن</>}</button>{!supabaseConfigured&&<small className="auth-error">لم تتم إضافة متغيرات Supabase إلى بيئة التشغيل.</small>}</form></div>}
    {notice&&<div className="toast">{notice}</div>}
  </div>
}
createRoot(document.getElementById('root')).render(<App/>);
