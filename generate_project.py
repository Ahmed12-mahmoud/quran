import json, re, zipfile, shutil, os, hashlib, random
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

# select 4 distinct positions per Juz per model, in four spatial quartiles.
# Each quarter is independently shuffled with a stable seed, so adjacent models
# do not receive adjacent verses. Al-Fatiha is excluded from the candidate pool.
def norm_q(t):
    return re.sub(r'[\u064B-\u065F\u0670\u06D6-\u06ED\s\u200f\u200e]', '', t)
used=set(); used_starts=set(); all_questions=[]; models=[]
for model in range(1,16):
    rows=[]
    for juz in range(1,31):
        arr=[x for x in juz_verses[juz] if x['surah'] != 1]
        chunks=[arr[(k*len(arr))//4:((k+1)*len(arr))//4] for k in range(4)]
        chosen=[]
        model_pages=[q['page'] for q in rows]
        for k,chunk in enumerate(chunks):
            order=list(range(len(chunk)))
            random.Random(20261020 + juz*100 + k*17).shuffle(order)
            # Stable shuffled order gives each model a far-apart, non-sequential verse.
            cand=chunk[order[model-1]]
            candidates=[chunk[i] for i in order]
            if (cand['surah'],cand['ayah']) in used or norm_q(cand['text']) in used_starts or (model_pages and abs(cand['page']-model_pages[-1])<3):
                for alt in candidates + arr:
                    if ((alt['surah'],alt['ayah']) not in used and norm_q(alt['text']) not in used_starts and (not model_pages or abs(alt['page']-model_pages[-1])>=3)):
                        cand=alt; break
                else:
                    raise RuntimeError(f'Could not satisfy 3-page spacing for model {model}, juz {juz}, quarter {k}')
            used.add((cand['surah'],cand['ayah']))
            used_starts.add(norm_q(cand['text']))
            item=dict(cand); item['kind']='أساسي' if k<3 else 'احتياطي'; item['model']=model; item['juz']=juz; item['uid']=f"{cand['surah']}:{cand['ayah']}"
            chosen.append(item); all_questions.append(item)
            model_pages.append(cand['page'])
        rows.extend(chosen)
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
assert len(all_questions)==1800
assert len({(q['surah'],q['ayah']) for q in all_questions})==1800
assert len({norm_q(q['text']) for q in all_questions})==1800, 'Edited text creates a duplicate normalized start'
assert all(len([q for q in m if q['kind']=='أساسي'])==90 and len([q for q in m if q['kind']=='احتياطي'])==30 for m in models)
assert all(abs(m[i]['page']-m[i-1]['page'])>=3 for m in models for i in range(1,len(m))), 'Questions are closer than 3 Mushaf pages'

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
    r=p.add_run('يشمل الأجزاء الثلاثين • 90 سؤالًا أساسيًا • 30 سؤالًا احتياطيًا'); r.font.size=Pt(12); r.font.color.rgb=RGBColor(75,75,75)
    p=d.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; rtl_paragraph(p); p.space_before=Pt(70)
    r=p.add_run('ورقة أسئلة للممتحن — دون إجابات نموذجية'); r.italic=True; r.font.size=Pt(11); r.font.color.rgb=RGBColor(100,100,100)
    d.add_page_break()
    # intro
    h=d.add_heading('إرشادات الاستخدام', level=1); rtl_paragraph(h); set_keep(h)
    for text in ['يقرأ الممتحن بداية الآية بصوت واضح، ثم يطلب من الطالب مواصلة التسميع شفويًا.','توزعت الأسئلة على أربعة نطاقات متباعدة داخل كل جزء؛ الأسئلة الثلاثة الأولى أساسية، والسؤال الرابع احتياطي.','النصوص القرآنية من نسخة Uthmani برواية حفص عن عاصم، وقد أُبقيت كما وردت في المصدر.']:
        p=d.add_paragraph(text); rtl_paragraph(p); p.paragraph_format.space_after=Pt(5)
    for juz in range(1,31):
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
    name=f'نموذج اختبار القرآن الكريم {model_no:02d}.docx'
    path=DOCS/name; make_doc(model_no,rows).save(path); shutil.copy2(path,PUBLIC_DOCS/name)
    metadata.append({'id':model_no,'name':f'نموذج اختبار القرآن الكريم {model_no:02d}','file':f'/documents/{name}','parts':30,'basic':90,'reserve':30,'questions':120,'preview':rows})
with zipfile.ZipFile(ROOT/'نماذج اختبارات القرآن الكريم.zip','w',zipfile.ZIP_DEFLATED) as z:
    for p in sorted(DOCS.glob('*.docx')): z.write(p,arcname=p.name)
# preview JSON: keep all data for browser preview, with source info
(ROOT/'public'/'models.json').write_text(json.dumps(metadata,ensure_ascii=False,indent=2),encoding='utf-8')
# QA report
report={'total_questions':len(all_questions),'basic_questions':1350,'reserve_questions':450,'unique_locations':len({q['uid'] for q in all_questions}),'unique_normalized_starts':len({norm(q['text']) for q in all_questions}),'models':15,'parts_per_model':30,'basic_per_model':90,'reserve_per_model':30,'duplicates':0,'duplicate_normalized_starts':0,'excluded_surahs':[1],'fatihah_questions':0,'minimum_page_gap_within_model':min(abs(m[i]['page']-m[i-1]['page']) for m in models for i in range(1,len(m))),'selection_strategy':'stable shuffled selection across four wide spatial quarters per Juz, with at least 3 standard Mushaf pages between consecutive questions inside each model','source':'Tanzil Quran text via Quran JSON (uthmani), Hafs scripts; standard page numbers from Quran.com verse metadata.','normalization':'Used only for QA identity comparison; display text is unchanged.','juz_distribution':{str(j):len(juz_verses[j]) for j in range(1,31)}}
(ROOT/'QA_REPORT.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
(ROOT/'questions_database.json').write_text(json.dumps({'source':'Tanzil/Quran JSON','questions':all_questions},ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps({'models':15,'docs':len(list(DOCS.glob('*.docx'))),'questions':len(all_questions),'unique':len(used),'zip':str(ROOT/'نماذج اختبارات القرآن الكريم.zip')},ensure_ascii=False))
