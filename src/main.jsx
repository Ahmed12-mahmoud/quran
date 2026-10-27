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
const systemLabels={lower:'من الناس إلى الفاتحة',upper:'من الفاتحة إلى الناس'};
const EDIT_KEY = 'quranExamEdits';

function App(){
  const [data,setData]=useState([]); const [rangeLoading,setRangeLoading]=useState(false); const [rangeError,setRangeError]=useState('');
  const [query,setQuery]=useState(''); const [selected,setSelected]=useState(null); const [selectedSystem,setSelectedSystem]=useState('lower'); const [selectedRange,setSelectedRange]=useState(1); const [notice,setNotice]=useState('');
  const [edits,setEdits]=useState(()=>{try{return JSON.parse(localStorage.getItem(EDIT_KEY)||'{}')}catch{return {}}});
  const saveTimers=useRef({}); const originalTexts=useRef({});
  const [editMode,setEditMode]=useState(false); const [user,setUser]=useState(null); const [authOpen,setAuthOpen]=useState(false); const [authLoading,setAuthLoading]=useState(false); const [email,setEmail]=useState(''); const [password,setPassword]=useState('');
  // يجب إعادة بناء النماذج عند وصول ملف النطاق async، وليس عند تعديل الأسئلة فقط.
  const effective=useMemo(()=>data.map(m=>({...m,preview:m.preview.map(q=>edits[q.uid]?{...q,text:edits[q.uid]}:q)})),[data,edits]);
  const rangeOptions=useMemo(()=>rangeIndex.filter(x=>x.system===selectedSystem).map(x=>({id:Number(x.range_id),name:x.name,file:x.file,range_key:x.range_key})).sort((a,b)=>a.id-b.id),[selectedSystem]);
  const currentEntry=useMemo(()=>rangeIndex.find(x=>x.system===selectedSystem&&Number(x.range_id)===Number(selectedRange)),[selectedRange,selectedSystem]);
  // البيانات محمّلة من الملف المختار نفسه؛ الاعتماد على النظام ورقم الاختبار
  // يمنع اختفاء النماذج إذا كانت قيمة range_key قديمة أو مختلفة في التخزين المؤقت.
  const filtered=useMemo(()=>{const needle=query.trim().toLowerCase();return effective.filter(m=>m.system===selectedSystem&&Number(m.range_id)===Number(selectedRange)&&(!needle||`${m.id} ${m.name} ${m.range_name} ${m.model_in_range}`.toLowerCase().includes(needle)))},[effective,query,selectedRange,selectedSystem]);
  const showNotice=(t)=>{setNotice(t); setTimeout(()=>setNotice(''),3200)};

  useEffect(()=>{setSelectedRange(1)},[selectedSystem]);
  useEffect(()=>{let active=true;setRangeLoading(true);setRangeError('');setSelected(null);setData([]);const entry=rangeIndex.find(x=>x.system===selectedSystem&&Number(x.range_id)===Number(selectedRange));if(!entry){setRangeLoading(false);return}const expectedKey=`${entry.system}-${String(entry.range_id).padStart(2,'0')}`;fetch(entry.file,{cache:'no-store'}).then(r=>{if(!r.ok)throw new Error(`HTTP ${r.status}`);return r.json()}).then(rows=>{const normalized=(Array.isArray(rows)?rows:[]).map(m=>({...m,system:entry.system,range_id:Number(entry.range_id),range_key:entry.range_key||expectedKey,range_name:m.range_name||entry.name}));if(active)setData(normalized)}).catch(e=>{if(active){setData([]);setRangeError(`${e.message} — تأكد من رفع ملف ${entry.file}`)}}).finally(()=>{if(active)setRangeLoading(false)});return()=>{active=false}},[selectedRange,selectedSystem]);

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
  const printMarkup=(model,root)=>{const esc=s=>String(s).replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));root.innerHTML=`<h1 style="text-align:center;color:#183f2e;font-size:24px;margin:0 0 4px">${esc(model.name)}</h1><div class="direct-print-sub" style="text-align:center;color:#66756d;font-size:13px;margin-bottom:18px;border-bottom:2px solid #b79a55;padding-bottom:10px">أسئلة اختبار شفوي — آخر نسخة معدلة</div>${model.preview.map(q=>`<article style="display:block;border:1px solid #d8e2db;border-right:4px solid #4f8566;border-radius:6px;padding:7px 10px;margin:0 0 6px;break-inside:auto;page-break-inside:auto"><div class="direct-print-meta" style="font-size:11px;font-weight:bold;color:#286044">${esc(q.part_name)} — ${esc(q.kind)} — سورة ${surahName(q.surah)}، آية ${q.ayah}</div><div class="direct-print-text" style="font-size:16px;margin-top:2px">${esc(q.text)}</div></article>`).join('')}`};
  const printModel=(model)=>{let root=document.getElementById('direct-print-root');if(!root){root=document.createElement('div');root.id='direct-print-root';document.body.appendChild(root)}printMarkup(model,root);root.style.display='block';const previousTitle=document.title;document.title=model.name;const cleanup=()=>{window.removeEventListener('afterprint',cleanup);setTimeout(()=>{root.remove();document.title=previousTitle},600)};window.addEventListener('afterprint',cleanup);requestAnimationFrame(()=>requestAnimationFrame(()=>setTimeout(()=>window.print(),350)))};
  const buildWordBlob=async(model)=>{
    const rows=model.preview;
    const cell=(text,opts={})=>new TableCell({shading:opts.fill?{fill:opts.fill,type:ShadingType.CLEAR}:undefined,verticalAlign:VerticalAlign.CENTER,margins:{top:130,bottom:130,left:170,right:170},borders:{top:{style:BorderStyle.SINGLE,size:4,color:opts.border||'D8E2DB'},bottom:{style:BorderStyle.SINGLE,size:4,color:opts.border||'D8E2DB'},left:{style:BorderStyle.SINGLE,size:4,color:opts.border||'D8E2DB'},right:{style:BorderStyle.SINGLE,size:16,color:opts.border||'4F8566'}},children:[new Paragraph({bidirectional:true,alignment:opts.center?AlignmentType.CENTER:AlignmentType.RIGHT,spacing:{after:0},children:[new TextRun({text:String(text),bold:!!opts.bold,color:opts.color||'263238',font:{name:'Al-Mateen',eastAsia:'Al-Mateen'},size:opts.size||21})]})]});
    const title=(text,size=30,color='164A3B')=>new Paragraph({alignment:AlignmentType.CENTER,bidirectional:true,spacing:{after:180},children:[new TextRun({text,bold:true,color,font:{name:'Al-Mateen',eastAsia:'Al-Mateen'},size})]});
    const children=[title('بِسْمِ اللهِ الرَّحْمَنِ الرَّحِيمِ',22),title(model.name,30)];
    for(const partId of [...new Set(rows.map(q=>q.part_id))].sort((a,b)=>a-b)){
      const jr=rows.filter(q=>q.part_id===partId); const partName=jr[0]?.part_name||`الوحدة ${partId}`;
      children.push(new Paragraph({alignment:AlignmentType.RIGHT,bidirectional:true,spacing:{before:160,after:80},children:[new TextRun({text:partName,bold:true,color:'164A3B',font:{name:'Al-Mateen',eastAsia:'Al-Mateen'},size:24})]}));
      const tableRows=[];
      for(const q of jr){
        const gold=q.kind==='احتياطي';
        const card=new TableCell({shading:{fill:gold?'FFF9E9':'FFFFFF',type:ShadingType.CLEAR},verticalAlign:VerticalAlign.CENTER,margins:{top:75,bottom:75,left:150,right:150},borders:{top:{style:BorderStyle.SINGLE,size:4,color:gold?'E4D29A':'D8E2DB'},bottom:{style:BorderStyle.SINGLE,size:4,color:gold?'E4D29A':'D8E2DB'},left:{style:BorderStyle.SINGLE,size:4,color:gold?'E4D29A':'D8E2DB'},right:{style:BorderStyle.SINGLE,size:18,color:gold?'B79A55':'4F8566'}},children:[new Paragraph({bidirectional:true,alignment:AlignmentType.RIGHT,spacing:{after:25},children:[new TextRun({text:`${q.kind}  •  سورة ${surahName(q.surah)}، آية ${q.ayah}`,bold:true,color:gold?'8A6A23':'286044',font:{name:'Arial',eastAsia:'Arial'},size:16})]}),new Paragraph({bidirectional:true,alignment:AlignmentType.RIGHT,spacing:{after:0},children:[new TextRun({text:q.text,color:'17231D',font:{name:'Al-Mateen',eastAsia:'Al-Mateen'},size:20})]})]});
        tableRows.push(new TableRow({cantSplit:true,children:[card]}));
      }
      children.push(new Table({width:{size:100,type:WidthType.PERCENTAGE},columnWidths:[10200],borders:{top:{style:BorderStyle.NONE,size:0,color:'FFFFFF'},bottom:{style:BorderStyle.NONE,size:0,color:'FFFFFF'},left:{style:BorderStyle.NONE,size:0,color:'FFFFFF'},right:{style:BorderStyle.NONE,size:0,color:'FFFFFF'},insideHorizontal:{style:BorderStyle.NONE,size:0,color:'FFFFFF'},insideVertical:{style:BorderStyle.NONE,size:0,color:'FFFFFF'}},rows:tableRows}));
    }
    return await Packer.toBlob(new Document({styles:{default:{document:{run:{font:{name:'Arial',eastAsia:'Arial'},size:21},paragraph:{bidirectional:true}}}},sections:[{properties:{page:{size:{width:11906,height:16838},margin:{top:936,bottom:936,left:1008,right:1008}}},children}]}));
  };
  const downloadUpdatedWord=async(model)=>{
    const blob=await buildWordBlob(model);
    const filename=`${model.name} - محدث.docx`;
    try{
      const file=new File([blob],filename,{type:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
      if(navigator.share&&navigator.canShare&&navigator.canShare({files:[file]})){await navigator.share({files:[file],title:filename,text:'ملف Word المحدث'});showNotice('اختر حفظ في الملفات أو الطباعة من شاشة المشاركة');return}
    }catch(error){if(error?.name==='AbortError')return}
    const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=filename;a.target='_blank';a.rel='noopener';document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),60000);showNotice('تم تجهيز ملف Word؛ افتح التنزيلات أو اختر حفظ الملف');
  };
  const buildPdfBlob=async(model)=>{
    const root=document.createElement('div');root.id='pdf-export-root';root.style.cssText='position:absolute;left:0;top:0;width:794px;min-height:1123px;box-sizing:border-box;padding:38px 42px;background:#fff;color:#17231d;direction:rtl;font-family:Arial,Tahoma,sans-serif;line-height:1.8;z-index:2147483647;visibility:visible;opacity:1;';document.body.appendChild(root);printMarkup(model,root);
    await new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve)));
    const canvas=await html2canvas(root,{scale:2,backgroundColor:'#ffffff',useCORS:true,logging:false,scrollX:0,scrollY:0,windowWidth:794,windowHeight:Math.max(1123,root.scrollHeight)});
    const pdf=new jsPDF({unit:'mm',format:'a4',orientation:'portrait',compress:true});
    const pageW=210,pageH=297,margin=10,imgW=pageW-margin*2,usableH=pageH-margin*2;
    const sourcePageH=Math.floor(canvas.width*usableH/imgW);const scale=canvas.width/root.getBoundingClientRect().width;const cardStarts=[...root.querySelectorAll('article')].map(el=>Math.floor((el.offsetTop*scale))).filter(y=>y>0);let sourceTop=0,page=0;
    while(sourceTop<canvas.height){if(page)pdf.addPage();let next=Math.min(sourceTop+sourcePageH,canvas.height);const safeStarts=cardStarts.filter(y=>y>sourceTop+80&&y<=next);if(next<canvas.height&&safeStarts.length)next=Math.max(sourceTop+80,Math.max(...safeStarts));const slice=document.createElement('canvas');slice.width=canvas.width;slice.height=Math.min(next-sourceTop,canvas.height-sourceTop);const ctx=slice.getContext('2d');ctx.fillStyle='#fff';ctx.fillRect(0,0,slice.width,slice.height);ctx.drawImage(canvas,0,sourceTop,canvas.width,slice.height,0,0,slice.width,slice.height);const sliceH=slice.height*imgW/slice.width;pdf.addImage(slice.toDataURL('image/jpeg',0.94),'JPEG',margin,margin,imgW,sliceH,undefined,'FAST');sourceTop+=slice.height;page++}
    root.remove();return pdf.output('blob');
  };
  const downloadUpdatedPdf=async(model)=>{
    try{const filename=`${model.name} - محدث.pdf`;const blob=await buildPdfBlob(model);const file=new File([blob],filename,{type:'application/pdf'});if(navigator.share&&navigator.canShare&&navigator.canShare({files:[file]})){try{await navigator.share({files:[file],title:filename,text:'ملف PDF المحدث'});showNotice('اختر Google Drive أو حفظ في الملفات');return}catch(shareError){if(shareError?.name==='AbortError')return;console.warn('PDF share unavailable, falling back to download',shareError)}}const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=filename;a.rel='noopener';document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),60000);showNotice('تم تنزيل PDF المحدث بآخر التعديلات')}catch(error){console.error(error);document.getElementById('pdf-export-root')?.remove();showNotice('تعذر إنشاء PDF المحدث')}};
  const downloadRangeZip=async()=>{
    try{
      if(!filtered.length){showNotice('لا توجد نماذج جاهزة لهذا الاختبار');return}
      const files={};
      for(const model of filtered){const blob=await buildWordBlob(model);files[`${model.name} - محدث.docx`]=new Uint8Array(await blob.arrayBuffer())}
      const zipBlob=new Blob([zipSync(files,{level:6})],{type:'application/zip'}); const zipName=`${rangeOptions.find(x=>x.id===selectedRange)?.name||'اختبار القرآن'} - محدث.zip`;
      const file=new File([zipBlob],zipName,{type:'application/zip'});
      if(navigator.share&&navigator.canShare&&navigator.canShare({files:[file]})){
        try{await navigator.share({files:[file],title:zipName,text:'أرشيف نماذج Word المحدث'});showNotice('تم تجهيز أرشيف Word المحدث');return}catch(shareError){if(shareError?.name==='AbortError')return;console.warn('Share unavailable, falling back to download',shareError)}
      }
      const url=URL.createObjectURL(zipBlob);const a=document.createElement('a');a.href=url;a.download=zipName;a.rel='noopener';document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),60000);showNotice('تم تنزيل أرشيف Word المحدث لكل النماذج')
    }catch(error){console.error(error);showNotice('تعذر إنشاء أرشيف Word؛ جرّب بعد اكتمال تحميل النماذج')}
  };
  const downloadRangePdf=async()=>{
    try{if(!filtered.length){showNotice('لا توجد نماذج جاهزة لهذا الاختبار');return}showNotice('جارٍ تجهيز 25 ملف PDF…');const files={};for(const model of filtered){const blob=await buildPdfBlob(model);files[`${model.name} - محدث.pdf`]=new Uint8Array(await blob.arrayBuffer())}const zipBlob=new Blob([zipSync(files,{level:6})],{type:'application/zip'});const zipName=`${rangeOptions.find(x=>x.id===selectedRange)?.name||'اختبار القرآن'} - PDF محدث.zip`;const file=new File([zipBlob],zipName,{type:'application/zip'});if(navigator.share&&navigator.canShare&&navigator.canShare({files:[file]})){try{await navigator.share({files:[file],title:zipName,text:'أرشيف PDF المحدث — 25 نموذجًا'});showNotice('اختر Google Drive أو حفظ في الملفات');return}catch(shareError){if(shareError?.name==='AbortError')return;console.warn('PDF archive share unavailable, falling back to download',shareError)}}const url=URL.createObjectURL(zipBlob);const a=document.createElement('a');a.href=url;a.download=zipName;a.rel='noopener';document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),60000);showNotice('تم تنزيل أرشيف PDF المحدث — 25 نموذجًا')}catch(error){console.error(error);document.getElementById('pdf-export-root')?.remove();showNotice('تعذر إنشاء أرشيف PDF؛ جرّب بعد اكتمال تحميل النماذج')}};
  const updateQuestion=(uid,text)=>{setSelected(s=>{const old=s?.preview.find(q=>q.uid===uid)?.text;if(s&&originalTexts.current[uid]===undefined)originalTexts.current[uid]=old;return s?{...s,preview:s.preview.map(q=>q.uid===uid?{...q,text}:q)}:s});setEdits(prev=>{const next={...prev,[uid]:text};localStorage.setItem(EDIT_KEY,JSON.stringify(next));return next})};
  const saveQuestionWithReason=async(uid,text)=>{if(!user){showNotice('سجّل دخول المشرف أولًا');return}const reason=window.prompt('اكتب سبب التعديل — هذا الحقل إجباري:','');if(reason===null)return;if(!reason.trim()){showNotice('يجب كتابة سبب التعديل قبل الحفظ');return}await persistQuestion(uid,text,reason.trim(),originalTexts.current[uid]);delete originalTexts.current[uid]};
  const persistQuestion=async(uid,text,reason,beforeText)=>{
    if(!user||!supabase){showNotice('تم الحفظ على هذا الجهاز فقط؛ سجّل دخول المشرف للحفظ المركزي');return}
    const q=selected?.preview.find(item=>item.uid===uid); const now=new Date().toISOString();
    const latest=await supabase.from('question_edits').upsert({uid,text,updated_at:now,updated_by:user.id},{onConflict:'uid'});
    const history=await supabase.from('question_edit_history').insert({uid,before_text:beforeText||'',after_text:text,reason,model_id:q?.model||selected?.id||null,model_name:selected?.name||'',range_key:q?.range_key||'',range_name:q?.range_name||selected?.range_name||'',surah:q?.surah||null,ayah:q?.ayah||null,updated_at:now,updated_by:user.id});
    const error=latest.error||history.error;
    showNotice(error?`تعذر الحفظ في Supabase: ${error.message}`:'تم حفظ السؤال مركزيًا في Supabase');
  };
  const exportEdits=()=>{const blob=new Blob([JSON.stringify({instructions:'ضع هذا الملف بجانب generate_project.py ثم شغّل python generate_project.py',edits},null,2)],{type:'application/json'});const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download='quran-exam-edits.json';a.click();URL.revokeObjectURL(a.href);showNotice('تم تصدير التعديلات')};
  const importEdits=(e)=>{const f=e.target.files?.[0];if(!f)return;const r=new FileReader();r.onload=()=>{try{const x=JSON.parse(r.result);const next=x.edits||x;setEdits(next);localStorage.setItem(EDIT_KEY,JSON.stringify(next));showNotice('تم استيراد التعديلات محليًا؛ احفظ كل سؤال لتزامنه مع Supabase')}catch{showNotice('تعذر قراءة ملف التعديلات')}};r.readAsText(f)};
  const login=async(e)=>{e.preventDefault();if(!supabase){showNotice('بيانات Supabase غير مضافة إلى بيئة التشغيل');return}setAuthLoading(true);const {data:result,error}=await supabase.auth.signInWithPassword({email:email.trim().toLowerCase(),password});setAuthLoading(false);if(error){showNotice(`خطأ Supabase: ${error.message}`);return}setUser(result.user);setAuthOpen(false);setPassword('');await loadRemoteEdits();showNotice('تم تسجيل دخول المشرف')};
  const logout=async()=>{await supabase?.auth.signOut();setUser(null);setEditMode(false);showNotice('تم تسجيل الخروج')};
  const startEdit=()=>{if(!user){setAuthOpen(true);showNotice('سجّل دخول المشرف أولًا للتعديل');return}setEditMode(!editMode)};
  return <div className="app">
    <header className="topbar"><div className="brand"><div className="brand-mark">۞</div><div><strong>اختبارات القرآن الكريم</strong><span>الحفظ الشفوي • إصدار تعليمي منظم</span></div></div><div className="top-actions"><button className="admin-btn" onClick={()=>user?logout():setAuthOpen(true)}>{user?<><LogOut size={15}/> خروج المشرف</>:<><LockKeyhole size={15}/> دخول المشرف</>}</button><a className="nav-link" href="#models">النماذج <ArrowLeft size={16}/></a></div></header>
    <main>
      <section className="hero"><div className="hero-copy"><div className="eyebrow"><Sparkles size={16}/> أدوات عملية للمحفّظ والممتحن</div><h1>اختبر الحفظ بثقة،<br/><em>واجعل كل جزء حاضرًا.</em></h1><p>1425 نموذجًا في نظامين: 675 من الناس إلى أعلى المصحف و750 من أعلى المصحف إلى أسفله، مع تباعد ثلاثة أوجه ومنع التشابه داخل النموذج.</p><div className="hero-actions"><a className="primary" href="#models">استعرض النماذج <ArrowLeft size={18}/></a><button className="secondary" onClick={downloadRangeZip}><Archive size={18}/> تنزيل Word للاختبار (25 نموذجًا)</button><button className="secondary" onClick={downloadRangePdf}><FileText size={18}/> تنزيل PDF للاختبار (25 نموذجًا)</button></div></div><div className="hero-card"><div className="seal">۞</div><div className="hero-card-title">نظامان للاختبار</div><div className="hero-card-text">25 نموذجًا لكل اختبار<br/>اختر أي اختبار لعرض نماذجه وتنزيل ملفه</div><div className="hero-card-rule"/><div className="hero-card-foot"><span>1425 نموذجًا</span><span>57 اختبارًا</span></div></div></section>
      <section className="stats"><div><b>1425</b><span>نموذج اختبار</span></div><div><b>57</b><span>اختبارًا متاحًا</span></div><div><b>4 + 3</b><span>حسب ترتيب الجزء</span></div><div><b>100%</b><span>قابل للتحرير والتنزيل</span></div></section>
      <section className="section-head" id="models"><div><div className="eyebrow">المكتبة الرقمية</div><h2>اختر النطاق ثم النموذج</h2><p>لكل اختبار 25 نموذجًا؛ ويظهر الجديد والقديم بعدد الأسئلة المعتمد لكل نظام.</p></div><div className="search"><Search size={18}/><input value={query} onChange={e=>setQuery(e.target.value)} placeholder="ابحث برقم النموذج أو الجزء…" aria-label="البحث في النماذج"/></div></section>
      <div className="system-switch"><button className={selectedSystem==='upper'?'active':''} onClick={()=>setSelectedSystem('upper')}>من الفاتحة إلى الناس</button><button className={selectedSystem==='lower'?'active':''} onClick={()=>setSelectedSystem('lower')}>من الناس إلى الفاتحة</button></div><div className="range-picker"><label htmlFor="range-select">اختر الاختبار</label><select id="range-select" value={selectedRange} onChange={e=>setSelectedRange(Number(e.target.value))}>{rangeOptions.map(r=><option key={r.id} value={r.id}>{r.name}</option>)}</select></div><div className="juz-tabs">{rangeOptions.map(r=><button key={r.id} className={selectedRange===r.id?'active':''} onClick={()=>setSelectedRange(r.id)}>{r.name}</button>)}</div>
      {rangeLoading?<div className="empty"><BookOpen size={30}/><h3>جارٍ تحميل نماذج الاختبار…</h3><p>{systemLabels[selectedSystem]}</p></div>:rangeError?<div className="empty"><Search size={30}/><h3>تعذر تحميل نماذج هذا النطاق</h3><p>{rangeError} — الملف المطلوب: {currentEntry?.file||'غير موجود في ranges/index.json'}</p></div>:filtered.length?<div className="grid">{filtered.map(m=><article className="model-card" key={m.id}><div className="card-top"><span className="model-no">{String(m.model_in_range).padStart(2,'0')}</span><span className="ready"><ShieldCheck size={15}/> جاهز</span></div><h3>{m.name}</h3><div className="card-meta"><span><BookOpen size={16}/> {m.range_name}</span><span><FileText size={16}/> {m.questions} أسئلة</span></div><div className="card-breakdown"><span>أساسي <b>{m.basic}</b></span><span>احتياطي <b>{m.reserve}</b></span></div><div className="card-actions"><button className="primary small" onClick={()=>downloadUpdatedWord(m)}><Download size={16}/> Word محدث</button><button className="print-btn" onClick={()=>downloadUpdatedPdf(m)}><FileText size={16}/> PDF محدث</button><button className="print-btn" onClick={()=>printModel(m)}><Printer size={16}/> طباعة</button><button className="text-btn" onClick={()=>{setSelected(m);setEditMode(false)}}>عرض وتعديل <ArrowLeft size={15}/></button></div></article>)}</div>:<div className="empty"><Search size={26}/><h3>لا توجد نماذج في هذا النطاق</h3><p>الملف المحمّل: {currentEntry?.file||'غير موجود في ranges/index.json'}</p></div>}
      <section className="how"><div className="how-intro"><div className="eyebrow">لوحة المشرف</div><h2>تعديل آمن ومركزي</h2><p>بعد تسجيل دخول المشرف، تُحفظ تعديلات الأسئلة في Supabase وتظهر لكل الزوار.</p></div><div className="steps"><div><span>01</span><h3>دخول المشرف</h3><p>استخدم حساب Supabase الذي أنشأته للمشرف.</p></div><div><span>02</span><h3>عدّل واحفظ</h3><p>كل سؤال يُحفظ مركزيًا عند مغادرة خانة التعديل.</p></div><div><span>03</span><h3>نزّل Word</h3><p>اضغط تنزيل Word المحدث من داخل النموذج.</p></div></div></section>
      <section className="source-note"><ShieldCheck size={22}/><div><strong>حالة الربط</strong><p>{supabaseConfigured?'Supabase مهيأ — '+(user?'تم تسجيل دخول المشرف':'بانتظار دخول المشرف'):'Supabase غير مهيأ في بيئة التشغيل'}</p></div><a href="https://tanzil.net/download/" target="_blank">المصدر</a></section>
    </main><footer><span>© اختبارات القرآن الكريم</span><span>صُمّم للمؤسسات التعليمية وحلقات التحفيظ</span></footer>
    {selected&&<div className="modal-backdrop" onClick={()=>setSelected(null)}><div className="modal" onClick={e=>e.stopPropagation()}><button className="close" onClick={()=>setSelected(null)}><X/></button><div className="modal-heading"><span className="model-no">{String(selected.id).padStart(2,'0')}</span><div><div className="eyebrow">معاينة النموذج</div><h2>{selected.name}</h2><p>أسئلة موزعة على نطاق الأجزاء المحدد — {editMode?'وضع التعديل مفعل':'وضع القراءة'}</p></div></div><div className="modal-actions"><button className="secondary" onClick={startEdit}><Pencil size={16}/> {editMode?'إنهاء التعديل':user?'تعديل الأسئلة':'دخول للتعديل'}</button><button className="secondary" onClick={()=>printModel(selected)}><Printer size={16}/> طباعة النموذج</button><button className="secondary" onClick={()=>downloadUpdatedPdf(selected)}><FileText size={16}/> PDF محدث</button></div><div className="preview-list">{selected.preview.map((q,i)=><div className={`preview-row ${q.kind==='احتياطي'?'reserve':''}`} key={q.uid}><div><b>{q.part_name}</b><small>{q.kind} • سورة {surahName(q.surah)}، آية {q.ayah}</small></div>{editMode?<textarea value={q.text} onChange={e=>updateQuestion(q.uid,e.target.value)} onBlur={e=>saveQuestionWithReason(q.uid,e.target.value)} aria-label={`تعديل سؤال ${i+1}`}/>:<p>{q.text} …</p>}</div>)}</div><button className="primary" onClick={()=>downloadUpdatedWord(selected)}><Download size={17}/> تنزيل Word المحدث</button><p className="edit-hint">عدّل الأسئلة ثم اضغط خارج الخانة للحفظ، وبعدها استخدم زر Word أو PDF المحدث.</p></div></div>}
    {authOpen&&<div className="modal-backdrop" onClick={()=>setAuthOpen(false)}><form className="auth-modal" onClick={e=>e.stopPropagation()} onSubmit={login}><button type="button" className="close" onClick={()=>setAuthOpen(false)}><X/></button><div className="auth-icon"><LockKeyhole/></div><div className="eyebrow">صلاحية المشرف</div><h2>تسجيل دخول المشرف</h2><p>هذا الدخول مخصص لتعديل الأسئلة وحفظها مركزيًا.</p><label>البريد الإلكتروني<input type="email" required value={email} onChange={e=>setEmail(e.target.value)} autoComplete="username" autoCapitalize="none"/></label><label>كلمة المرور<input type="password" required value={password} onChange={e=>setPassword(e.target.value)} autoComplete="current-password"/></label><button className="primary" disabled={authLoading}>{authLoading?'جارٍ التحقق…':<><LogIn size={17}/> دخول آمن</>}</button>{!supabaseConfigured&&<small className="auth-error">لم تتم إضافة متغيرات Supabase إلى بيئة التشغيل.</small>}</form></div>}
    {notice&&<div className="toast">{notice}</div>}
  </div>
}
createRoot(document.getElementById('root')).render(<App/>);
