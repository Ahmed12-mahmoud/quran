import json, re, zipfile, shutil, os, hashlib, random, collections, difflib
from pathlib import Path
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT = Path(__file__).parent
DOCS = ROOT / 'documents'
PUBLIC_DOCS = ROOT / 'public' / 'documents'
DOCS.mkdir(exist_ok=True); PUBLIC_DOCS.mkdir(parents=True, exist_ok=True)
for old in DOCS.glob('*.docx'): old.unlink()
for old in PUBLIC_DOCS.glob('*.docx'): old.unlink()

SURAH_NAMES = ['الفاتحة','البقرة','آل عمران','النساء','المائدة','الأنعام','الأعراف','الأنفال','التوبة','يونس','هود','يوسف','الرعد','إبراهيم','الحجر','النحل','الإسراء','الكهف','مريم','طه','الأنبياء','الحج','المؤمنون','النور','الفرقان','الشعراء','النمل','القصص','العنكبوت','الروم','لقمان','السجدة','الأحزاب','سبأ','فاطر','يس','الصافات','ص','الزمر','غافر','فصلت','الشورى','الزخرف','الدخان','الجاثية','الأحقاف','محمد','الفتح','الحجرات','ق','الذاريات','الطور','النجم','القمر','الرحمن','الواقعة','الحديد','المجادلة','الحشر','الممتحنة','الصف','الجمعة','المنافقون','التغابن','الطلاق','التحريم','الملك','القلم','الحاقة','المعارج','نوح','الجن','المزمل','المدثر','القيامة','الإنسان','المرسلات','النبأ','النازعات','عبس','التكوير','الانفطار','المطففين','الانشقاق','البروج','الطارق','الأعلى','الغاشية','الفجر','البلد','الشمس','الليل','الضحى','الشرح','التين','العلق','القدر','البينة','الزلزلة','العاديات','القارعة','التكاثر','العصر','الهمزة','الفيل','قريش','الماعون','الكوثر','الكافرون','النصر','المسد','الإخلاص','الفلق','الناس']
JUZ_STARTS = [(1,1),(2,142),(2,253),(3,93),(4,24),(4,148),(5,82),(6,111),(7,88),(8,41),(9,93),(11,6),(12,53),(15,1),(17,1),(18,75),(21,1),(23,1),(25,21),(27,56),(29,46),(33,31),(36,28),(39,32),(41,47),(46,1),(51,31),(58,1),(67,1),(78,1)]

quran = json.loads((ROOT/'quran.json').read_text())
page_map = json.loads((ROOT/'page_map.json').read_text())
verses=[]
for ch in quran:
    for v in ch['verses']:
        verses.append({'surah':ch['id'],'ayah':v['id'],'text':v['text'],'page':page_map.get(f"{ch['id']}:{v['id']}")})
index={(x['surah'],x['ayah']):i for i,x in enumerate(verses)}
# assign each verse to a Juz by the canonical start positions
starts=[index[s] for s in JUZ_STARTS]
for i,x in enumerate(verses):
    x['juz'] = max(0, min(29, next((j for j,s in enumerate(starts) if i < s), 30)-1)) if False else 0
for j,s in enumerate(starts):
    end = starts[j+1] if j<29 else len(verses)
    for x in verses[s:end]: x['juz']=j+1
juz_verses={j:[x for x in verses if x['juz']==j] for j in range(1,31)}

# Mark each verse's position on its standard Mushaf page for QA and reporting.
page_groups={}
for verse in verses:
    page_groups.setdefault(verse.get('page'),[]).append(verse)
for group in page_groups.values():
    for pos,verse in enumerate(group,1):
        verse['page_pos']=pos
        verse['page_total']=len(group)

# Build two cumulative systems: 27 named lower units and 30 standard upper Juz ranges.
def norm_q(t):
    return re.sub(r'[\u064B-\u065F\u0670\u06D6-\u06ED\s\u200f\u200e]', '', t)
def excerpt(t, limit=10):
    words=t.split(); return t if len(words)<=limit else ' '.join(words[:limit])+' …'
def excerpt_end(t, limit=10):
    words=t.split(); return t if len(words)<=limit else '… '+' '.join(words[-limit:])
def qtext(a,b):
    return f"من قوله تعالى: «{excerpt(a['text'])}» إلى ختام الآية: «{excerpt_end(b['text'], 10)}»"
def similar_start(text, previous):
    a=norm_q(text); aw=a.split()[:6]
    for old in previous:
        b=norm_q(old).split()[:6]
        if aw == b: return True
    return False
def slice_between(start_key,end_key):
    return verses[index[start_key]:index[end_key]+1]
def valid_start(start, used_pages, model_starts, model_seen):
    key=norm_q(start['text'])
    return key not in model_seen and all(abs(start['page']-p)>=3 for p in used_pages)
def choose_questions(pool, needed, rng, model_starts, model_seen, usage):
    eligible=[x for x in pool if x['surah']!=1 and x.get('page') and x['ayah']>2 and x['page']<=last_page-2]
    if not eligible: raise RuntimeError('No eligible verses in part')
    first_page=min(x['page'] for x in eligible)
    bands=[[x for x in eligible if x['page']-first_page<7],[x for x in eligible if 7<=x['page']-first_page<14],[x for x in eligible if x['page']-first_page>=14]]
    for _ in range(120):
        pools=[list(x) for x in bands]
        for group in pools: rng.shuffle(group)
        basics=None
        for a in pools[0]:
            if not valid_start(a,[],model_starts,model_seen): continue
            for b in pools[1]:
                if not valid_start(b,[a['page']],model_starts,model_seen): continue
                for c in pools[2]:
                    if valid_start(c,[a['page'],b['page']],model_starts,model_seen): basics=[a,b,c]; break
                if basics: break
            if basics: break
        if not basics: continue
        gaps=[(basics[0]['page']+3,basics[1]['page']-3),(basics[1]['page']+3,basics[2]['page']-3)]
        extra_candidates=[]
        for lo,hi in gaps:
            group=[x for x in eligible if lo<=x['page']<=hi]; rng.shuffle(group); extra_candidates.extend(group)
        extras=[]; used=[x['page'] for x in basics]
        for x in extra_candidates:
            if len(extras)>=needed-3: break
            if valid_start(x,used,model_starts,model_seen): extras.append(x); used.append(x['page'])
        if len(extras)<needed-3: continue
        result=[]; local_keys=[]; okay=True
        for idx,start_q in enumerate(basics+extras):
            ends=[x for x in verses if start_q['page']+2<=x.get('page',0)<=start_q['page']+3]
            rng.shuffle(ends); chosen=None
            for end_q in ends:
                text=qtext(start_q,end_q); key=f"{start_q['surah']}:{start_q['ayah']}-{end_q['surah']}:{end_q['ayah']}"
                if usage[key]<3 and key not in model_seen and key not in local_keys:
                    chosen=(start_q,end_q,text,key,'أساسي' if idx<3 else 'احتياطي'); break
            if not chosen: okay=False; break
            result.append(chosen); local_keys.append(chosen[3])
        if okay:
            for item in result:
                model_starts.append(item[0]['text']); model_seen.add(norm_q(item[0]['text'])); model_seen.add(item[3]); usage[item[3]]+=1
            return result
    raise RuntimeError(f'Unable to distribute {needed} questions across 7+7+6 pages; eligible={len(eligible)}, bands={[len(x) for x in bands]}, first_page={first_page}')

last_page=max(x.get('page',0) for x in verses)
CUSTOM_UNITS=[
 (1,'الفاتحة إلى سيقول السفهاء','الفاتحة إلى سيقول السفهاء',(1,1),(2,141),4,3),
 (2,'باقي البقرة','باقي سورة البقرة',(2,142),(2,286),5,4),
 (3,'آل عمران','سورة آل عمران',(3,1),(3,200),5,4),
 (4,'النساء','سورة النساء',(4,1),(4,176),5,4),
 (5,'المائدة','سورة المائدة',(5,1),(5,120),4,3),
 (6,'الأنعام','سورة الأنعام',(6,1),(6,165),4,3),
 (7,'الأعراف','سورة الأعراف',(7,1),(7,206),4,3),
 (8,'الأنفال إلى التوبة','الأنفال إلى التوبة',(8,1),(9,129),5,4),
 (9,'يونس إلى هود','يونس إلى هود',(10,1),(11,123),5,4),
 (10,'يوسف إلى إبراهيم','يوسف إلى إبراهيم',(12,1),(14,52),5,4),
 (11,'الحجر إلى النحل','الحجر إلى النحل',(15,1),(16,128),4,3),
 (12,'الإسراء إلى الكهف','الإسراء إلى الكهف',(17,1),(18,110),4,3),
 (13,'مريم إلى طه','مريم إلى طه',(19,1),(20,135),4,3),
 (14,'الأنبياء إلى الحج','الأنبياء إلى الحج',(21,1),(22,78),4,3),
 (15,'المؤمنون إلى النور','المؤمنون إلى النور',(23,1),(24,64),4,3),
 (16,'الفرقان إلى الشعراء','الفرقان إلى الشعراء',(25,1),(26,227),4,3),
 (17,'النمل إلى القصص','النمل إلى القصص',(27,1),(28,88),4,3),
 (18,'العنكبوت إلى السجدة','العنكبوت إلى السجدة',(29,1),(32,30),4,3),
 (19,'الأحزاب إلى فاطر','الأحزاب إلى فاطر',(33,1),(35,45),4,3),
 (20,'يس إلى ص','يس إلى ص',(36,1),(38,88),4,3),
 (21,'الزمر إلى فصلت','الزمر إلى فصلت',(39,1),(41,54),4,3),
 (22,'الشورى إلى الجاثية','الشورى إلى الجاثية',(42,1),(45,37),4,3),
 (23,'الأحقاف إلى ق','الأحقاف إلى ق',(46,1),(50,45),4,3),
 (24,'الذاريات إلى الحديد','الذاريات إلى الحديد',(51,1),(57,29),4,3),
 (25,'المجادلة إلى التحريم','المجادلة إلى التحريم',(58,1),(66,12),4,3),
 (26,'الملك إلى المرسلات','الملك إلى المرسلات',(67,1),(77,50),4,3),
 (27,'عم إلى الناس','عم إلى الناس',(78,1),(114,6),4,3),
]
custom_pools={u[0]:slice_between(u[3],u[4]) for u in CUSTOM_UNITS}
custom_by_id={u[0]:u for u in CUSTOM_UNITS}
models=[]; all_questions=[]; usage=collections.Counter(); model_id=0
family_max_reuse={}

def add_family(family, family_title, ranges):
    global model_id
    family_usage=collections.Counter()
    for range_id,parts,range_name,new_part,counts in ranges:
        for local_model in range(1,26):
            model_id+=1; rows=[]; model_seen=set(); model_starts=[]
            for part_id in parts:
                pool = custom_pools[part_id] if family=='lower' else juz_verses[part_id]
                needed = counts[0] if part_id==new_part else counts[1]
                selected=choose_questions(pool,needed,random.Random(20261026+model_id*101+part_id),model_starts,model_seen,family_usage)
                part_name = custom_by_id[part_id][1] if family=='lower' else f'الجزء {part_id:02d}'
                for q_index,(start_q,end_q,text,key,kind) in enumerate(selected):
                    item=dict(start_q); item.update({'end_surah':end_q['surah'],'end_ayah':end_q['ayah'],'end_page':end_q['page'],'text':text,'kind':kind,'model':model_id,'system':family,'system_title':family_title,'range_id':range_id,'range_key':f'{family}-{range_id:02d}','range_name':range_name,'model_in_range':local_model,'part_id':part_id,'part_name':part_name,'is_new':part_id==new_part,'uid':f'{family}-r{range_id}-m{local_model}-p{part_id}-{start_q["surah"]}:{start_q["ayah"]}-{end_q["surah"]}:{end_q["ayah"]}'})
                    rows.append(item); all_questions.append(item)
            models.append(rows)
    family_max_reuse[family]=max(family_usage.values(),default=0)
    usage.update(family_usage)

lower_ranges=[]
for rid in range(1,28):
    start=28-rid; parts=list(range(start,28)); start_name=custom_by_id[start][1]
    start_display=start_name.split(' إلى ')[0].replace('باقي ','')
    label='اختبار من الناس إلى النبأ' if rid==1 else f'اختبار من {start_display} إلى الناس'
    unit=custom_by_id[start]; lower_ranges.append((rid,parts,label,start,(unit[5],unit[6])))
add_family('lower','الاختبارات من الناس إلى الفاتحة',lower_ranges)
upper_ranges=[]
for rid in range(1,31):
    parts=list(range(1,rid+1)); start_name='الفاتحة'
    end_verse=verses[starts[rid]-1] if rid<30 else verses[-1]
    end_name=f"{SURAH_NAMES[end_verse['surah']-1]} حتى الآية {end_verse['ayah']}"
    label=f'اختبار من {start_name} إلى {end_name}'
    upper_ranges.append((rid,parts,label,rid,(4,3)))
add_family('upper','الاختبارات من الفاتحة إلى الناس',upper_ranges)

# Optional edits exported from the browser editor. Place the exported file at
# the project root as quran-exam-edits.json before rerunning this script.
edits_path=ROOT/'quran-exam-edits.json'
if edits_path.exists():
    payload=json.loads(edits_path.read_text(encoding='utf-8'))
    edits=payload.get('edits', payload)
    for q in all_questions:
        if q['uid'] in edits and str(edits[q['uid']]).strip():
            q['text']=str(edits[q['uid']]).strip()
    by_uid={q['uid']:q for q in all_questions}
    for model_rows in models:
        for q in model_rows:
            q['text']=by_uid[q['uid']]['text']

# normalize for QA only (does not alter displayed text)
def norm(t):
    return re.sub(r'[\u064B-\u065F\u0670\u06D6-\u06ED\s\u200f\u200e]', '', t)
assert len(models)==1425
assert len([m for m in models if m[0]['system']=='lower'])==675
assert len([m for m in models if m[0]['system']=='upper'])==750
assert all(q['ayah']>2 and q['surah']!=1 for q in all_questions)
assert max(family_max_reuse.values())<=3
assert all(q['end_page']-q['page']>=2 for q in all_questions), 'Every question must span at least two Mushaf pages'

# helpers
def shade(cell, fill):
    tcPr=cell._tc.get_or_add_tcPr(); shd=OxmlElement('w:shd'); shd.set(qn('w:fill'),fill); tcPr.append(shd)
def set_cell_border(cell, **kwargs):
    tc=cell._tc; tcPr=tc.get_or_add_tcPr(); borders=tcPr.first_child_found_in('w:tcBorders')
    if borders is None: borders=OxmlElement('w:tcBorders'); tcPr.append(borders)
    for edge, data in kwargs.items():
        tag='w:{}'.format(edge); element=borders.find(qn(tag))
        if element is None: element=OxmlElement(tag); borders.append(element)
        for key,val in data.items(): element.set(qn('w:'+key), str(val))
def rtl_paragraph(p):
    pPr=p._p.get_or_add_pPr(); bidi=OxmlElement('w:bidi'); bidi.set(qn('w:val'),'1'); pPr.append(bidi)
def rtl_cell(cell):
    for p in cell.paragraphs: rtl_paragraph(p)
def set_keep(p):
    pPr=p._p.get_or_add_pPr(); keep=OxmlElement('w:keepNext'); pPr.append(keep)
def add_page_field(paragraph):
    run=paragraph.add_run('الصفحة ')
    run.font.name='Arial'; run.font.size=Pt(8); run.font.color.rgb=RGBColor(100,100,100)
    fld_begin=OxmlElement('w:fldChar'); fld_begin.set(qn('w:fldCharType'),'begin')
    instr=OxmlElement('w:instrText'); instr.set(qn('xml:space'),'preserve'); instr.text=' PAGE '
    fld_sep=OxmlElement('w:fldChar'); fld_sep.set(qn('w:fldCharType'),'separate')
    cached=OxmlElement('w:t'); cached.text='1'
    fld_end=OxmlElement('w:fldChar'); fld_end.set(qn('w:fldCharType'),'end')
    for node in (fld_begin,instr,fld_sep,cached,fld_end): run._r.append(node)
def prevent_row_split(row):
    trPr=row._tr.get_or_add_trPr(); cant=OxmlElement('w:cantSplit'); trPr.append(cant)

def set_font(run, name='Al-Mateen', size=21, bold=False, color=None):
    run.font.name=name; run.font.size=Pt(size); run.font.bold=bold
    run._element.rPr.rFonts.set(qn('w:ascii'),name); run._element.rPr.rFonts.set(qn('w:hAnsi'),name); run._element.rPr.rFonts.set(qn('w:eastAsia'),name)
    if color: run.font.color.rgb=RGBColor.from_string(color)
def add_text_paragraph(d,text,size=21,bold=False,color='263238',align=WD_ALIGN_PARAGRAPH.RIGHT,space_before=0,space_after=5):
    p=d.add_paragraph(); p.alignment=align; p.paragraph_format.space_before=Pt(space_before); p.paragraph_format.space_after=Pt(space_after); rtl_paragraph(p)
    r=p.add_run(text); set_font(r,size=size,bold=bold,color=color); return p
def set_font(run, name='Al-Mateen', size=21, bold=False, color=None):
    run.font.name=name; run.font.size=Pt(size); run.font.bold=bold
    run._element.rPr.rFonts.set(qn('w:ascii'),name); run._element.rPr.rFonts.set(qn('w:hAnsi'),name); run._element.rPr.rFonts.set(qn('w:eastAsia'),name)
    if color: run.font.color.rgb=RGBColor.from_string(color)
def add_text_paragraph(d,text,size=21,bold=False,color='263238',align=WD_ALIGN_PARAGRAPH.RIGHT,space_before=0,space_after=5):
    p=d.add_paragraph(); p.alignment=align; p.paragraph_format.space_before=Pt(space_before); p.paragraph_format.space_after=Pt(space_after); rtl_paragraph(p)
    r=p.add_run(text); set_font(r,size=size,bold=bold,color=color); return p
def make_doc(model_no, rows, exam_name):
    d=Document(); sec=d.sections[0]; sec.page_width=Inches(8.27); sec.page_height=Inches(11.69)
    sec.top_margin=Inches(.65); sec.bottom_margin=Inches(.65); sec.left_margin=Inches(.7); sec.right_margin=Inches(.7)
    styles=d.styles
    for name,size,color,bold in [('Normal',12,'263238',False),('Heading 1',18,'164A3B',True),('Heading 2',14,'B38B3D',True)]:
        st=styles[name]; st.font.name='Al-Mateen'; st.font.size=Pt(size); st.font.bold=bold; st.font.color.rgb=RGBColor.from_string(color)
        st._element.rPr.rFonts.set(qn('w:ascii'),'Al-Mateen'); st._element.rPr.rFonts.set(qn('w:hAnsi'),'Al-Mateen'); st._element.rPr.rFonts.set(qn('w:eastAsia'),'Al-Mateen')
    footer=sec.footer.paragraphs[0]; footer.alignment=WD_ALIGN_PARAGRAPH.CENTER; rtl_paragraph(footer)
    r=footer.add_run(f'{exam_name}  •  النموذج {model_no:02d}  •  '); set_font(r,name='Arial',size=8,color='64736b'); add_page_field(footer)
    p=d.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after=Pt(16); rtl_paragraph(p)
    r=p.add_run('بِسْمِ اللهِ الرَّحْمَنِ الرَّحِيمِ'); set_font(r,size=18,bold=True,color='164A3B')
    add_text_paragraph(d,exam_name,size=28,bold=True,color='164A3B',align=WD_ALIGN_PARAGRAPH.CENTER,space_before=30,space_after=12)
    add_text_paragraph(d,f'نموذج {model_no:02d}',size=22,bold=True,color='B38B3D',align=WD_ALIGN_PARAGRAPH.CENTER,space_after=70)
    d.add_page_break()
    add_text_paragraph(d,'إرشادات الممتحن',size=19,bold=True,color='164A3B',space_after=8)
    for text in ['يقرأ الممتحن بداية السؤال، ثم يطلب من الطالب مواصلة التسميع حتى خاتمة الآية المحددة.','تظهر الأسئلة مرتبة حسب وحدات الاختبار، مع تباعد لا يقل عن ثلاثة أوجه بين بدايات الأسئلة داخل الوحدة.','لا تبدأ الأسئلة من الآية الأولى أو الثانية من أي سورة، ولا تشمل سورة الفاتحة.']:
        add_text_paragraph(d,text,size=13,color='4f5f57',space_after=5)
    for part_id in sorted({x['part_id'] for x in rows}):
        part_rows=[x for x in rows if x['part_id']==part_id]; part_name=part_rows[0]['part_name']
        p=add_text_paragraph(d,part_name,size=17,bold=True,color='164A3B',space_before=14,space_after=6); p.paragraph_format.keep_with_next=True
        table=d.add_table(rows=1,cols=3); table.alignment=WD_TABLE_ALIGNMENT.CENTER; table.style='Table Grid'
        hdr=table.rows[0].cells
        for c,txt in zip(hdr,['النوع','السؤال','الموضع']):
            c.text=txt; shade(c,'164A3B'); rtl_cell(c)
            for pp in c.paragraphs:
                pp.alignment=WD_ALIGN_PARAGRAPH.CENTER
                for rr in pp.runs: set_font(rr,name='Arial',size=10,bold=True,color='FFFFFF')
        prevent_row_split(table.rows[0])
        for q in part_rows:
            row=table.add_row(); prevent_row_split(row); cells=row.cells
            cells[0].text='أساسي'; cells[1].text=q['text']; cells[2].text=f"سورة {SURAH_NAMES[q['surah']-1]} — آية {q['ayah']} إلى سورة {SURAH_NAMES[q['end_surah']-1]} — آية {q['end_ayah']}"
            for c in cells: rtl_cell(c); c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
            shade(cells[0],'EAF3EE')
            for pp in cells[0].paragraphs: pp.alignment=WD_ALIGN_PARAGRAPH.CENTER
            for pp in cells[1].paragraphs:
                pp.alignment=WD_ALIGN_PARAGRAPH.RIGHT
                for rr in pp.runs: set_font(rr,name='Al-Mateen',size=13)
            for pp in cells[2].paragraphs:
                pp.alignment=WD_ALIGN_PARAGRAPH.CENTER
                for rr in pp.runs: set_font(rr,name='Arial',size=9,color='5a5a5a')
        d.add_paragraph().paragraph_format.space_after=Pt(2)
    return d

metadata=[]
for model_no,rows in enumerate(models,1):
    first=rows[0]; family=first['system']; rid=first['range_id']; local=first['model_in_range']; exam_name=first['range_name']
    filename=f'{exam_name} - نموذج {local:02d}.docx'
    path=DOCS/filename; make_doc(model_no,rows,exam_name).save(path); shutil.copy2(path,PUBLIC_DOCS/filename)
    metadata.append({'id':model_no,'name':f'{exam_name} — نموذج {local:02d}','filename':filename,'file':f'/documents/{filename}','system':family,'system_title':first['system_title'],'range_id':rid,'range_key':first['range_key'],'range_name':exam_name,'model_in_range':local,'parts':len(set(q['part_id'] for q in rows)),'basic':sum(q['kind']=='أساسي' for q in rows),'reserve':sum(q['kind']=='احتياطي' for q in rows),'questions':len(rows),'preview':rows})
with zipfile.ZipFile(ROOT/'نماذج اختبارات القرآن الكريم.zip','w',zipfile.ZIP_DEFLATED) as z:
    for p in sorted(DOCS.glob('*.docx')): z.write(p,arcname=p.name)
(ROOT/'public'/'models.json').write_text(json.dumps(metadata,ensure_ascii=False,indent=2),encoding='utf-8')
ranges_dir=ROOT/'public'/'ranges'; range_zips_dir=ROOT/'public'/'range-zips'
ranges_dir.mkdir(parents=True,exist_ok=True); range_zips_dir.mkdir(parents=True,exist_ok=True)
for old in list(ranges_dir.glob('range-*.json'))+list(range_zips_dir.glob('range-*.zip')):
    old.unlink()
index_rows=[]
for family,count in [('lower',27),('upper',30)]:
    family_rows=[m for m in metadata if m['system']==family]
    for rid in range(1,count+1):
        rows=[m for m in family_rows if m['range_id']==rid]
        key=f'{family}-{rid:02d}'
        (ranges_dir/f'{key}.json').write_text(json.dumps(rows,ensure_ascii=False,indent=2),encoding='utf-8')
        with zipfile.ZipFile(range_zips_dir/f'{key}.zip','w',zipfile.ZIP_DEFLATED) as z:
            for m in rows: z.write(DOCS/m['filename'],arcname=m['filename'])
        if rows: index_rows.append({'system':family,'range_id':rid,'range_key':key,'name':rows[0]['range_name'],'models':25,'file':f'/ranges/{key}.json','zip':f'/range-zips/{key}.zip'})
(ranges_dir/'index.json').write_text(json.dumps(index_rows,ensure_ascii=False,indent=2),encoding='utf-8')
(ranges_dir/'catalog.json').write_text(json.dumps(index_rows,ensure_ascii=False,indent=2),encoding='utf-8')
report={'total_questions':len(all_questions),'basic_questions':len([q for q in all_questions if q['kind']=='أساسي']),'reserve_questions':len([q for q in all_questions if q['kind']=='احتياطي']),'unique_locations':len({q['uid'] for q in all_questions}),'unique_question_texts':len({norm(q['text']) for q in all_questions}),'models':len(models),'lower_models':675,'upper_models':750,'lower_ranges':27,'upper_ranges':30,'models_per_range':25,'max_exact_question_reuse':max(family_max_reuse.values()),'max_exact_question_reuse_per_system':family_max_reuse,'excluded_surahs':[1],'start_filters':'لا تبدأ الأسئلة من الآية الأولى أو الثانية من أي سورة، ولا تشمل الفاتحة.','minimum_page_gap_within_model':3,'ending_rule':'تنتهي الأسئلة عند خاتمة آية فعلية، ويُعرض مقتطف من آخر كلمات الآية.','systems':{'lower':'27 وحدة مسماة من الناس إلى أعلى المصحف؛ الجديد 5 أو 4 حسب الوحدة، والقديم حسب تعريف الوحدة.','upper':'30 جزءًا قياسيًا من أعلى المصحف إلى أسفله؛ الجديد 4 والقديم 3.'},'source':'Tanzil Quran text via Quran JSON (uthmani), Hafs scripts; standard page numbers from Quran.com verse metadata.'}
(ROOT/'QA_REPORT.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
(ROOT/'questions_database.json').write_text(json.dumps({'source':'Tanzil/Quran JSON','questions':all_questions},ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps({'models':len(models),'lower_models':675,'upper_models':750,'docs':len(list(DOCS.glob('*.docx'))),'questions':len(all_questions),'max_reuse_per_system':family_max_reuse,'zip':str(ROOT/'نماذج اختبارات القرآن الكريم.zip')},ensure_ascii=False))
