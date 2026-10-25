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

# Build 25 cumulative models for each range: [30], [29..30], ... [1..30].
def norm_q(t):
    return re.sub(r'[\u064B-\u065F\u0670\u06D6-\u06ED\s\u200f\u200e]', '', t)
def excerpt(t, limit=10):
    words=t.split(); return t if len(words)<=limit else ' '.join(words[:limit])+' …'
def qtext(a,b):
    return f"من قوله تعالى: «{excerpt(a['text'])}» إلى قوله تعالى: «{excerpt(b['text'], 8)}»"
def similar_start(text, previous):
    a=norm_q(text)
    aw=a.split()
    for old in previous:
        b=norm_q(old); bw=b.split()
        if aw[:6] == bw[:6]: return True
        if difflib.SequenceMatcher(None,a,b).ratio() >= 0.72: return True
    return False
def range_end(start, juz):
    pool=[x for x in juz_verses[juz] if start['page']+2 <= x.get('page',0) <= start['page']+3]
    return random.choice(pool or [x for x in juz_verses[juz] if x.get('page',0)>=start['page']+2] or [start])

last_page=max(x.get('page',0) for x in verses)
eligible={j:[x for x in juz_verses[j] if x['surah']!=1 and x.get('page') and x['ayah']>2 and x['page']<=last_page-2] for j in range(1,31)}
juz_first_page={j:page_map.get(f"{surah}:{ayah}", min(x['page'] for x in juz_verses[j] if x.get('page'))) for j,(surah,ayah) in enumerate(JUZ_STARTS,1)}
def page_band(juz, verse):
    """تقسيم كل جزء إلى 7 أوجه + 7 أوجه + بقية الوجه السادس تقريبًا."""
    offset=verse['page']-juz_first_page[juz]
    return 0 if offset < 7 else 1 if offset < 14 else 2
def in_band(juz, band):
    return [x for x in eligible[juz] if page_band(juz,x)==band]
def valid_start(start, used_pages, model_starts, model_seen):
    key=norm_q(start['text'])
    return key not in model_seen and not similar_start(start['text'], model_starts) and all(abs(start['page']-p)>=3 for p in used_pages)
models=[]; all_questions=[]; usage=collections.Counter(); model_id=0
for top_juz in range(30,0,-1):
    parts=list(range(top_juz,31))
    for local_model in range(1,26):
        model_id+=1; rows=[]; model_seen=set(); model_starts=[]; model_pages=[]
        for juz in parts:
            needed=4 if juz==top_juz else 3
            seed=20261020+top_juz*10000+local_model*101+juz
            rng=random.Random(seed)
            bands=[in_band(juz,b) for b in range(3)]
            for pool in bands: rng.shuffle(pool)
            chosen_basic=None; reserve_start=None
            for a in bands[0]:
                if not valid_start(a,[],model_starts,model_seen): continue
                for b in bands[1]:
                    if not valid_start(b,[a['page']],model_starts,model_seen): continue
                    for c in bands[2]:
                        if not valid_start(c,[a['page'],b['page']],model_starts,model_seen): continue
                        candidate_basic=[a,b,c]
                        if needed==3:
                            chosen_basic=candidate_basic; break
                        lo1,hi1=a['page']+3,b['page']-3; lo2,hi2=b['page']+3,c['page']-3
                        reserve_pool=[x for x in eligible[juz] if (lo1<=x['page']<=hi1) or (lo2<=x['page']<=hi2)]
                        rng.shuffle(reserve_pool)
                        reserve=next((x for x in reserve_pool if valid_start(x,[q['page'] for q in candidate_basic],model_starts,model_seen)),None)
                        if reserve:
                            chosen_basic=candidate_basic; reserve_start=reserve; break
                    if chosen_basic: break
                if chosen_basic: break
            if not chosen_basic:
                raise RuntimeError(f'No 7+7+6 distribution for juz {juz}, range top {top_juz}, model {local_model}')
            selected=chosen_basic[:]
            if needed==4:
                if reserve_start is None:
                    raise RuntimeError(f'No reserve gap of 3 pages for juz {juz}, range top {top_juz}, model {local_model}')
                selected.append(reserve_start)
            for q_index,start_q in enumerate(selected):
                end_pool=[x for x in verses if start_q['page']+2 <= x.get('page',0) <= start_q['page']+3]
                rng.shuffle(end_pool)
                chosen_end=None
                for candidate_end in end_pool or [range_end(start_q,juz)]:
                    candidate_text=qtext(start_q,candidate_end); candidate_key=norm_q(candidate_text)
                    if usage[candidate_key] < 3 and candidate_key not in model_seen:
                        chosen_end,candidate_text,candidate_key=candidate_end,candidate_text,candidate_key; break
                if chosen_end is None:
                    raise RuntimeError(f'No reusable-safe ending for juz {juz}, range top {top_juz}, model {local_model}')
                end_q,text,key=chosen_end,candidate_text,candidate_key
                usage[key]+=1; model_seen.add(norm_q(start_q['text'])); model_seen.add(key); model_starts.append(start_q['text']); model_pages.append(start_q['page'])
                kind='احتياطي' if q_index==3 else 'أساسي'
                item=dict(start_q); item.update({'end_surah':end_q['surah'],'end_ayah':end_q['ayah'],'end_page':end_q['page'],'text':text,'kind':kind,'model':model_id,'range_id':31-top_juz,'top_juz':top_juz,'model_in_range':local_model,'juz':juz,'uid':f"r{31-top_juz}-m{local_model}-j{juz}-{start_q['surah']}:{start_q['ayah']}-{end_q['surah']}:{end_q['ayah']}"})
                rows.append(item); all_questions.append(item)
        models.append(rows)

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
assert len(models)==750
assert len(all_questions)==35625
assert all(q['ayah']>2 and q['surah']!=1 for q in all_questions)
assert max(usage.values())<=3
assert all(q['end_page']-q['page']>=2 for q in all_questions), 'Every question must span at least two Mushaf pages'
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

def make_doc(model_no, rows):
    d=Document(); sec=d.sections[0]; sec.page_width=Inches(8.27); sec.page_height=Inches(11.69)
    sec.top_margin=Inches(.65); sec.bottom_margin=Inches(.65); sec.left_margin=Inches(.7); sec.right_margin=Inches(.7)
    styles=d.styles
    for name,size,color,bold in [('Normal',11,'263238',False),('Heading 1',18,'164A3B',True),('Heading 2',14,'B38B3D',True)]:
        st=styles[name]; st.font.name='Arial'; st.font.size=Pt(size); st.font.bold=bold; st.font.color.rgb=RGBColor.from_string(color)
        st._element.rPr.rFonts.set(qn('w:eastAsia'),'Arial')
    # footer
    footer=sec.footer.paragraphs[0]; footer.alignment=WD_ALIGN_PARAGRAPH.CENTER; rtl_paragraph(footer)
    run=footer.add_run('اختبارات الحفظ الشفوي للقرآن الكريم  •  النموذج %02d  •  ' % model_no); run.font.name='Arial'; run.font.size=Pt(8); run.font.color.rgb=RGBColor(100,100,100)
    add_page_field(footer)
    # cover
    p=d.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.space_after=Pt(16); rtl_paragraph(p)
    r=p.add_run('بِسْمِ اللهِ الرَّحْمَنِ الرَّحِيمِ'); r.bold=True; r.font.size=Pt(17); r.font.color.rgb=RGBColor(22,74,59)
    p=d.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; rtl_paragraph(p); p.space_before=Pt(35)
    r=p.add_run('اختبارات الحفظ الشفوي\nللقرآن الكريم'); r.bold=True; r.font.size=Pt(27); r.font.color.rgb=RGBColor(22,74,59)
    p=d.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; rtl_paragraph(p); p.space_before=Pt(18)
    r=p.add_run('نموذج رقم %02d' % model_no); r.bold=True; r.font.size=Pt(21); r.font.color.rgb=RGBColor(179,139,61)
    p=d.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; rtl_paragraph(p); p.space_before=Pt(40)
    r=p.add_run('نطاق الأجزاء %02d–30 • الجزء الرئيسي: %02d' % (rows[0]['top_juz'], rows[0]['top_juz'])); r.font.size=Pt(12); r.font.color.rgb=RGBColor(75,75,75)
    p=d.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; rtl_paragraph(p); p.space_before=Pt(70)
    r=p.add_run('ورقة أسئلة للممتحن — دون إجابات نموذجية'); r.italic=True; r.font.size=Pt(11); r.font.color.rgb=RGBColor(100,100,100)
    d.add_page_break()
    # intro
    h=d.add_heading('إرشادات الاستخدام', level=1); rtl_paragraph(h); set_keep(h)
    for text in ['يقرأ الممتحن بداية الآية بصوت واضح، ثم يطلب من الطالب مواصلة التسميع شفويًا.','الجزء الأعلى في النطاق له أربعة أسئلة، ولكل جزء أدنى ثلاثة أسئلة أساسية.','لا تبدأ الأسئلة من الآيتين الأولى والثانية، ولا توجد أسئلة متشابهة داخل النموذج، مع تباعد لا يقل عن ثلاثة أوجه بين البدايات.']:
        p=d.add_paragraph(text); rtl_paragraph(p); p.paragraph_format.space_after=Pt(5)
    for juz in sorted({x['juz'] for x in rows}):
        h=d.add_heading('الجزء %02d' % juz, level=1); rtl_paragraph(h); set_keep(h)
        table=d.add_table(rows=1, cols=3); table.alignment=WD_TABLE_ALIGNMENT.CENTER; table.style='Table Grid'
        hdr=table.rows[0].cells
        for c,txt in zip(hdr,['النوع','بداية الآية التي يقرأها الممتحن','الموضع']):
            c.text=txt; shade(c,'164A3B'); rtl_cell(c)
            for p in c.paragraphs:
                p.alignment=WD_ALIGN_PARAGRAPH.CENTER
                for rr in p.runs: rr.font.bold=True; rr.font.color.rgb=RGBColor(255,255,255); rr.font.size=Pt(10)
        prevent_row_split(table.rows[0])
        for q in [x for x in rows if x['juz']==juz]:
            row=table.add_row(); prevent_row_split(row); cells=row.cells; cells[0].text=q['kind']; cells[1].text=q['text']+' …'; cells[2].text=f"سورة {SURAH_NAMES[q['surah']-1]} — آية {q['ayah']}"
            for c in cells: rtl_cell(c); c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
            shade(cells[0], 'EAF3EE' if q['kind']=='أساسي' else 'FFF6DF')
            for p in cells[0].paragraphs: p.alignment=WD_ALIGN_PARAGRAPH.CENTER
            for p in cells[1].paragraphs:
                p.alignment=WD_ALIGN_PARAGRAPH.RIGHT
                for rr in p.runs: rr.font.name='Arial'; rr.font.size=Pt(11)
            for p in cells[2].paragraphs:
                p.alignment=WD_ALIGN_PARAGRAPH.CENTER
                for rr in p.runs: rr.font.size=Pt(9); rr.font.color.rgb=RGBColor(90,90,90)
        d.add_paragraph().paragraph_format.space_after=Pt(2)
    return d

metadata=[]
for model_no,rows in enumerate(models,1):
    local=rows[0]['model_in_range']; top=rows[0]['top_juz']; rid=rows[0]['range_id']
    name=f'النطاق {rid:02d} - من الجزء {top:02d} إلى 30 - نموذج {local:02d}.docx'
    path=DOCS/name; make_doc(model_no,rows).save(path); shutil.copy2(path,PUBLIC_DOCS/name)
    metadata.append({'id':model_no,'name':f'النطاق {rid:02d}: الأجزاء {top:02d}–30 — نموذج {local:02d}','file':f'/documents/{name}','range_id':rid,'top_juz':top,'parts':30-top+1,'model_in_range':local,'basic':3*(30-top)+3,'reserve':1,'questions':3*(30-top)+4,'preview':rows})
with zipfile.ZipFile(ROOT/'نماذج اختبارات القرآن الكريم.zip','w',zipfile.ZIP_DEFLATED) as z:
    for p in sorted(DOCS.glob('*.docx')): z.write(p,arcname=p.name)
# preview JSON: keep all data for browser preview, with source info
(ROOT/'public'/'models.json').write_text(json.dumps(metadata,ensure_ascii=False,indent=2),encoding='utf-8')
# Split browser data and Word archives so every uploaded GitHub file stays small.
ranges_dir=ROOT/'public'/'ranges'; range_zips_dir=ROOT/'public'/'range-zips'
ranges_dir.mkdir(parents=True,exist_ok=True); range_zips_dir.mkdir(parents=True,exist_ok=True)
for rid in range(1,31):
    range_rows=[m for m in metadata if m['range_id']==rid]
    (ranges_dir/f'range-{rid:02d}.json').write_text(json.dumps(range_rows,ensure_ascii=False,indent=2),encoding='utf-8')
    with zipfile.ZipFile(range_zips_dir/f'range-{rid:02d}.zip','w',zipfile.ZIP_DEFLATED) as z:
        for m in range_rows:
            p=DOCS/Path(m['file']).name
            z.write(p,arcname=p.name)
(ranges_dir/'index.json').write_text(json.dumps([{'range_id':rid,'top_juz':31-rid,'parts':rid,'models':25,'file':f'/ranges/range-{rid:02d}.json'} for rid in range(1,31)],ensure_ascii=False,indent=2),encoding='utf-8')
# QA report
report={'total_questions':len(all_questions),'basic_questions':len([q for q in all_questions if q['kind']=='أساسي']),'reserve_questions':len([q for q in all_questions if q['kind']=='احتياطي']),'unique_locations':len({q['uid'] for q in all_questions}),'unique_question_texts':len({norm(q['text']) for q in all_questions}),'models':750,'ranges':30,'models_per_range':25,'max_exact_question_reuse':3,'excluded_surahs':[1],'start_filters':'Never starts at ayah 1 or 2 of any surah, including Juz Amma.','similarity_filter':'No high lexical similarity within a model; sequence ratio and six-word fingerprints are rejected.','minimum_page_gap_within_model':3,'distribution':'Each Juz has three basic questions: one in the first 7 pages, one in the second 7 pages, and one in the third 6-page section. The top Juz has one reserve question placed between basic questions with a minimum 3-page gap; lower Juz have 3 basic questions.','source':'Tanzil Quran text via Quran JSON (uthmani), Hafs scripts; standard page numbers from Quran.com verse metadata.'}
(ROOT/'QA_REPORT.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
(ROOT/'questions_database.json').write_text(json.dumps({'source':'Tanzil/Quran JSON','questions':all_questions},ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps({'models':750,'docs':len(list(DOCS.glob('*.docx'))),'questions':len(all_questions),'max_reuse':max(usage.values()),'zip':str(ROOT/'نماذج اختبارات القرآن الكريم.zip')},ensure_ascii=False))
