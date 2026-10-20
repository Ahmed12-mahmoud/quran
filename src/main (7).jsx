import React, {useEffect, useMemo, useRef, useState} from 'react';
import {createRoot} from 'react-dom/client';
import {Download, Search, BookOpen, ShieldCheck, ArrowLeft, FileText, Archive, Sparkles, X, Pencil, Save, Upload, Database, LogIn, LogOut, Cloud, LockKeyhole, Printer} from 'lucide-react';
import './styles.css';
import {supabase, supabaseConfigured} from './lib/supabase';
import {Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell, WidthType, ShadingType, VerticalAlign, BorderStyle} from 'docx';

const data = await fetch('/models.json').then(r=>r.json());
const zipUrl = '/quran-exams.zip';
const EDIT_KEY = 'quranExamEdits';

function App(){
  const [query,setQuery]=useState(''); const [selected,setSelected]=useState(null); const [notice,setNotice]=useState('');
  const [edits,setEdits]=useState(()=>{try{return JSON.parse(localStorage.getItem(EDIT_KEY)||'{}')}catch{return {}}});
  const saveTimers=useRef({});
  const [editMode,setEditMode]=useState(false); const [user,setUser]=useState(null); const [authOpen,setAuthOpen]=useState(false); const [authLoading,setAuthLoading]=useState(false); const [email,setEmail]=useState(''); const [password,setPassword]=useState('');
  const effective=useMemo(()=>data.map(m=>({...m,preview:m.preview.map(q=>edits[q.uid]?{...q,text:edits[q.uid]}:q)})),[edits]);
  const filtered=useMemo(()=>effective.filter(m=>`${m.id} ${m.name}`.includes(query.trim())),[effective,query]);
  const showNotice=(t)=>{setNotice(t); setTimeout(()=>setNotice(''),3200)};

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
  const printModel=(model)=>{const esc=s=>String(s).replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));let root=document.getElementById('direct-print-root');if(!root){root=document.createElement('div');root.id='direct-print-root';document.body.appendChild(root)}root.innerHTML=`<h1>${esc(model.name)}</h1><div class="direct-print-sub">أسئلة اختبار شفوي — نسخة الطباعة</div>${model.preview.map(q=>`<article><div class="direct-print-meta">الجزء ${String(q.juz).padStart(2,'0')} — ${esc(q.kind)} — سورة ${q.surah}، آية ${q.ayah}</div><div class="direct-print-text">${esc(q.text)}</div></article>`).join('')}`;const cleanup=()=>{root.remove();window.removeEventListener('afterprint',cleanup)};window.addEventListener('afterprint',cleanup);window.print()};
  const downloadUpdatedWord=async(model)=>{
    const rows=model.preview;
    const cell=(text,opts={})=>new TableCell({shading:opts.fill?{fill:opts.fill,type:ShadingType.CLEAR}:undefined,verticalAlign:VerticalAlign.CENTER,children:[new Paragraph({bidirectional:true,alignment:opts.center?AlignmentType.CENTER:AlignmentType.RIGHT,spacing:{after:0},children:[new TextRun({text:String(text),bold:!!opts.bold,color:opts.color||'263238',font:{name:'Arial',eastAsia:'Arial'},size:opts.size||21})]})]});
    const children=[new Paragraph({text:'بِسْمِ اللهِ الرَّحْمَنِ الرَّحِيمِ',heading:HeadingLevel.HEADING_1,alignment:AlignmentType.CENTER,bidirectional:true}),new Paragraph({text:model.name,heading:HeadingLevel.HEADING_1,alignment:AlignmentType.CENTER,bidirectional:true}),new Paragraph({text:'أسئلة اختبار شفوي — نسخة محدثة من الموقع',alignment:AlignmentType.CENTER,bidirectional:true,spacing:{after:260}})];
    for(let juz=1;juz<=30;juz++){
      const jr=rows.filter(q=>q.juz===juz);
      children.push(new Paragraph({text:`الجزء ${String(juz).padStart(2,'0')}`,heading:HeadingLevel.HEADING_2,alignment:AlignmentType.RIGHT,bidirectional:true,spacing:{before:160,after:80}}));
      const tableRows=[new TableRow({children:[cell('النوع',{fill:'164A3B',color:'FFFFFF',bold:true,center:true,size:18}),cell('بداية الآية التي يقرأها الممتحن',{fill:'164A3B',color:'FFFFFF',bold:true,center:true,size:18}),cell('الموضع',{fill:'164A3B',color:'FFFFFF',bold:true,center:true,size:18})]})];
      for(const q of jr)tableRows.push(new TableRow({children:[cell(q.kind,{fill:q.kind==='أساسي'?'EAF3EE':'FFF6DF',center:true,size:18}),cell(`${q.text} …`,{size:21}),cell(`سورة ${q.surah} — آية ${q.ayah}`,{center:true,size:17})]}));
      children.push(new Table({width:{size:100,type:WidthType.PERCENTAGE},columnWidths:[1500,6900,1800],borders:{top:{style:BorderStyle.SINGLE,size:4,color:'B7C8BC'},bottom:{style:BorderStyle.SINGLE,size:4,color:'B7C8BC'},left:{style:BorderStyle.SINGLE,size:4,color:'B7C8BC'},right:{style:BorderStyle.SINGLE,size:4,color:'B7C8BC'},insideHorizontal:{style:BorderStyle.SINGLE,size:2,color:'D8E2DB'},insideVertical:{style:BorderStyle.SINGLE,size:2,color:'D8E2DB'}},rows:tableRows}));
    }
    const blob=await Packer.toBlob(new Document({styles:{default:{document:{run:{font:{name:'Arial',eastAsia:'Arial'},size:21},paragraph:{bidirectional:true}}}},sections:[{properties:{page:{size:{width:11906,height:16838},margin:{top:936,bottom:936,left:1008,right:1008}}},children}]}));
    const filename=`${model.name} - محدث.docx`;
    try{
      const file=new File([blob],filename,{type:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
      if(navigator.share&&navigator.canShare&&navigator.canShare({files:[file]})){await navigator.share({files:[file],title:filename,text:'ملف Word المحدث'});showNotice('اختر حفظ في الملفات أو الطباعة من شاشة المشاركة');return}
    }catch(error){if(error?.name==='AbortError')return}
    const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=filename;a.target='_blank';a.rel='noopener';document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),60000);showNotice('تم تجهيز ملف Word؛ افتح التنزيلات أو اختر حفظ الملف');
  };
  const updateQuestion=(uid,text)=>{setEdits(prev=>{const next={...prev,[uid]:text};localStorage.setItem(EDIT_KEY,JSON.stringify(next));setSelected(s=>s?{...s,preview:s.preview.map(q=>q.uid===uid?{...q,text}:q)}:s);return next});if(user&&supabase){clearTimeout(saveTimers.current[uid]);saveTimers.current[uid]=setTimeout(()=>persistQuestion(uid,text),700)}};
  const persistQuestion=async(uid,text)=>{
    if(!user||!supabase){showNotice('تم الحفظ على هذا الجهاز فقط؛ سجّل دخول المشرف للحفظ المركزي');return}
    const {error}=await supabase.from('question_edits').upsert({uid,text,updated_at:new Date().toISOString(),updated_by:user.id},{onConflict:'uid'});
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
      <section className="hero"><div className="hero-copy"><div className="eyebrow"><Sparkles size={16}/> أدوات عملية للمحفّظ والممتحن</div><h1>اختبر الحفظ بثقة،<br/><em>واجعل كل جزء حاضرًا.</em></h1><p>خمسة عشر نموذجًا متكاملًا تغطي الأجزاء الثلاثين، مصممة لتكون واضحة في الجلسة وسهلة في الطباعة والتنزيل.</p><div className="hero-actions"><a className="primary" href="#models">استعرض النماذج <ArrowLeft size={18}/></a><button className="secondary" onClick={()=>download(zipUrl,'ملف النماذج الكامل')}><Archive size={18}/> تنزيل الكل ZIP</button></div></div><div className="hero-card"><div className="seal">۞</div><div className="hero-card-title">من البقرة إلى الناس</div><div className="hero-card-text">أسئلة تبدأ من أوائل آيات صحيحة<br/>ومواضع متباعدة داخل كل جزء</div><div className="hero-card-rule"/><div className="hero-card-foot"><span>15 نموذجًا</span><span>1,800 سؤال فريد</span></div></div></section>
      <section className="stats"><div><b>15</b><span>نموذج اختبار</span></div><div><b>30</b><span>جزءًا في كل نموذج</span></div><div><b>90 + 30</b><span>أساسي واحتياطي</span></div><div><b>100%</b><span>قابل للتحرير والتنزيل</span></div></section>
      <section className="section-head" id="models"><div><div className="eyebrow">المكتبة الرقمية</div><h2>اختر نموذجك</h2><p>كل بطاقة مرتبطة بملف Word حقيقي ومعاينة كاملة للمحتوى.</p></div><div className="search"><Search size={18}/><input value={query} onChange={e=>setQuery(e.target.value)} placeholder="ابحث برقم النموذج…" aria-label="البحث في النماذج"/></div></section>
      {filtered.length?<div className="grid">{filtered.map(m=><article className="model-card" key={m.id}><div className="card-top"><span className="model-no">{String(m.id).padStart(2,'0')}</span><span className="ready"><ShieldCheck size={15}/> جاهز</span></div><h3>{m.name}</h3><div className="card-meta"><span><BookOpen size={16}/> {m.parts} جزءًا</span><span><FileText size={16}/> {m.questions} سؤال</span></div><div className="card-breakdown"><span>أساسي <b>{m.basic}</b></span><span>احتياطي <b>{m.reserve}</b></span></div><div className="card-actions"><button className="primary small" onClick={()=>downloadUpdatedWord(m)}><Download size={16}/> Word محدث</button><button className="print-btn" onClick={()=>printModel(m)}><Printer size={16}/> طباعة</button><button className="text-btn" onClick={()=>{setSelected(m);setEditMode(false)}}>عرض وتعديل <ArrowLeft size={15}/></button></div></article>)}</div>:<div className="empty"><Search size={26}/><h3>لم نعثر على نموذج مطابق</h3><p>جرّب البحث برقم من 01 إلى 15.</p></div>}
      <section className="how"><div className="how-intro"><div className="eyebrow">لوحة المشرف</div><h2>تعديل آمن ومركزي</h2><p>بعد تسجيل دخول المشرف، تُحفظ تعديلات الأسئلة في Supabase وتظهر لكل الزوار.</p></div><div className="steps"><div><span>01</span><h3>دخول المشرف</h3><p>استخدم حساب Supabase الذي أنشأته للمشرف.</p></div><div><span>02</span><h3>عدّل واحفظ</h3><p>كل سؤال يُحفظ مركزيًا عند مغادرة خانة التعديل.</p></div><div><span>03</span><h3>نزّل Word</h3><p>اضغط تنزيل Word المحدث من داخل النموذج.</p></div></div></section>
      <section className="source-note"><ShieldCheck size={22}/><div><strong>حالة الربط</strong><p>{supabaseConfigured?'Supabase مهيأ — '+(user?'تم تسجيل دخول المشرف':'بانتظار دخول المشرف'):'Supabase غير مهيأ في بيئة التشغيل'}</p></div><a href="https://tanzil.net/download/" target="_blank">المصدر</a></section>
    </main><footer><span>© اختبارات القرآن الكريم</span><span>صُمّم للمؤسسات التعليمية وحلقات التحفيظ</span></footer>
    {selected&&<div className="modal-backdrop" onClick={()=>setSelected(null)}><div className="modal" onClick={e=>e.stopPropagation()}><button className="close" onClick={()=>setSelected(null)}><X/></button><div className="modal-heading"><span className="model-no">{String(selected.id).padStart(2,'0')}</span><div><div className="eyebrow">معاينة النموذج</div><h2>{selected.name}</h2><p>120 سؤالًا موزعة على الأجزاء الثلاثين — {editMode?'وضع التعديل مفعل':'وضع القراءة'}</p></div></div><div className="modal-actions"><button className="secondary" onClick={startEdit}><Pencil size={16}/> {editMode?'إنهاء التعديل':user?'تعديل الأسئلة':'دخول للتعديل'}</button><button className="secondary" onClick={()=>printModel(selected)}><Printer size={16}/> طباعة النموذج</button></div><div className="preview-list">{selected.preview.map((q,i)=><div className={`preview-row ${q.kind==='احتياطي'?'reserve':''}`} key={q.uid}><div><b>الجزء {String(q.juz).padStart(2,'0')}</b><small>{q.kind} • سورة {q.surah}، آية {q.ayah}</small></div>{editMode?<textarea value={q.text} onChange={e=>updateQuestion(q.uid,e.target.value)} onBlur={e=>persistQuestion(q.uid,e.target.value)} aria-label={`تعديل سؤال ${i+1}`}/>:<p>{q.text} …</p>}</div>)}</div><button className="primary" onClick={()=>downloadUpdatedWord(selected)}><Download size={17}/> تنزيل Word المحدث</button><p className="edit-hint">عدّل الأسئلة ثم اضغط خارج الخانة للحفظ، وبعدها استخدم زر Word المحدث.</p></div></div>}
    {authOpen&&<div className="modal-backdrop" onClick={()=>setAuthOpen(false)}><form className="auth-modal" onClick={e=>e.stopPropagation()} onSubmit={login}><button type="button" className="close" onClick={()=>setAuthOpen(false)}><X/></button><div className="auth-icon"><LockKeyhole/></div><div className="eyebrow">صلاحية المشرف</div><h2>تسجيل دخول المشرف</h2><p>هذا الدخول مخصص لتعديل الأسئلة وحفظها مركزيًا.</p><label>البريد الإلكتروني<input type="email" required value={email} onChange={e=>setEmail(e.target.value)} autoComplete="username" autoCapitalize="none"/></label><label>كلمة المرور<input type="password" required value={password} onChange={e=>setPassword(e.target.value)} autoComplete="current-password"/></label><button className="primary" disabled={authLoading}>{authLoading?'جارٍ التحقق…':<><LogIn size={17}/> دخول آمن</>}</button>{!supabaseConfigured&&<small className="auth-error">لم تتم إضافة متغيرات Supabase إلى بيئة التشغيل.</small>}</form></div>}
    {notice&&<div className="toast">{notice}</div>}
  </div>
}
createRoot(document.getElementById('root')).render(<App/>);
