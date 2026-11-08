import json, re, random, shutil, os
from collections import Counter, defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
CATALOG = json.loads((ROOT/'public/route-catalog-1263.json').read_text(encoding='utf-8'))
QURAN = json.loads((ROOT/'quran.json').read_text(encoding='utf-8'))
PAGE_MAP = json.loads((ROOT/'page_map.json').read_text(encoding='utf-8'))
OUT = ROOT/'public'/'route-models'
SURAH_NAMES = ['الفاتحة','البقرة','آل عمران','النساء','المائدة','الأنعام','الأعراف','الأنفال','التوبة','يونس','هود','يوسف','الرعد','إبراهيم','الحجر','النحل','الإسراء','الكهف','مريم','طه','الأنبياء','الحج','المؤمنون','النور','الفرقان','الشعراء','النمل','القصص','العنكبوت','الروم','لقمان','السجدة','الأحزاب','سبأ','فاطر','يس','الصافات','ص','الزمر','غافر','فصلت','الشورى','الزخرف','الدخان','الجاثية','الأحقاف','محمد','الفتح','الحجرات','ق','الذاريات','الطور','النجم','القمر','الرحمن','الواقعة','الحديد','المجادلة','الحشر','الممتحنة','الصف','الجمعة','المنافقون','التغابن','الطلاق','التحريم','الملك','القلم','الحاقة','المعارج','نوح','الجن','المزمل','المدثر','القيامة','الإنسان','المرسلات','النبأ','النازعات','عبس','التكوير','الانفطار','المطففين','الانشقاق','البروج','الطارق','الأعلى','الغاشية','الفجر','البلد','الشمس','الليل','الضحى','الشرح','التين','العلق','القدر','البينة','الزلزلة','العاديات','القارعة','التكاثر','العصر','الهمزة','الفيل','قريش','الماعون','الكوثر','الكافرون','النصر','المسد','الإخلاص','الفلق','الناس']
AR = str.maketrans('٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹','01234567890123456789')

def nums(s): return [int(x) for x in re.findall(r'\d+', s.translate(AR))]
def normalize(s): return re.sub(r'[\u064B-\u065F\u0670\u06D6-\u06ED\s\u200f\u200e]', '', s)
def excerpt(s, limit=10):
    w=s.split(); return s if len(w)<=limit else ' '.join(w[:limit])+' …'
def excerpt_end(s, limit=10):
    w=s.split()
    if len(w)<=12: return s
    return '… '+' '.join(w[-min(14,len(w)):])
def qtext(a,b): return f'من قوله تعالى: «{excerpt(a["text"])}» إلى ختام الآية: «{excerpt_end(b["text"])}»'

verses=[]
for chapter in QURAN:
    for verse in chapter['verses']:
        x={'surah':chapter['id'],'ayah':verse['id'],'text':verse['text'],'page':PAGE_MAP.get(f"{chapter['id']}:{verse['id']}")}
        if x['page']: verses.append(x)
by_page=defaultdict(list)
for v in verses: by_page[v['page']].append(v)
for group in by_page.values():
    for i,v in enumerate(group): v['page_pos']=i+1; v['page_total']=len(group)
last_page=max(v['page'] for v in verses)
index={(v['surah'],v['ayah']):i for i,v in enumerate(verses)}

def row_bounds(raw):
    m=re.search(r'([٠-٩0-9]+)\s*[–-]\s*([٠-٩0-9]+)', raw)
    if not m: return (1,last_page)
    a,b=nums(m.group(0)); return (min(a,b),max(a,b))

def row_count(raw):
    # The PDF uses phrases such as "4 أسئلة" or "سؤالان" in the موضع column.
    s=raw.translate(AR)
    m=re.search(r'(\d+)\s*سؤ',s)
    if m: return max(1,int(m.group(1)))
    if 'سؤالان' in raw or 'سؤاالن' in raw: return 2
    if 'ثالثة أسئلة' in raw or 'ثالث أسئلة' in raw: return 3
    if 'أربعة أسئلة' in raw or 'أربع أسئلة' in raw: return 4
    if 'خمسة أسئلة' in raw or 'خمس أسئلة' in raw: return 5
    return 4

def eligible(a,lo,hi):
    return lo<=a['page']<=hi and a['surah']!=1 and a['ayah']>2 and a['page']<last_page-1

def choose_question(pool, page, used_pages, seen_starts, usage, rng, lo, hi):
    candidates=[v for v in pool if v['page']==page and normalize(v['text']) not in seen_starts and all(abs(v['page']-p)>=3 for p in used_pages)]
    if not candidates: candidates=[v for v in pool if eligible(v,lo,hi) and normalize(v['text']) not in seen_starts and all(abs(v['page']-p)>=3 for p in used_pages)]
    rng.shuffle(candidates)
    for start in candidates:
        ends=[v for v in pool if v['page']>=start['page']+2 and v['page']<=min(hi,start['page']+2) and v['surah']!=1 and v['ayah']>1]
        if not ends: ends=[v for v in pool if v['page']<=hi and v['page']>=start['page'] and v['surah']!=1 and v['ayah']>1]
        rng.shuffle(ends)
        for end in ends:
            key=f'{start["surah"]}:{start["ayah"]}-{end["surah"]}:{end["ayah"]}'
            if usage[key] < 12:
                return start,end,key
    return None

def select_questions(lo,hi,total,seed,route_seen,usage,basic_count=None):
    pool=[v for v in verses if eligible(v,lo,hi)]
    if not pool: raise RuntimeError(f'empty range {lo}-{hi}')
    pages=sorted({v['page'] for v in pool}); span=max(1,len(pages))
    basic=(max(1,total-1) if total>=4 else total) if basic_count is None else basic_count
    bands=4 if basic>=4 else 3
    # Fixed face bands: 7 + 7 + 6 for the normal 3-basic case; four equal 7-face bands when needed.
    boundaries=[]
    if bands==3:
        if span < 18:
            boundaries=[pages[0], pages[0]+max(1,span//3), pages[0]+max(2,(span*2)//3), pages[-1]+1]
        else:
            boundaries=[pages[0],pages[0]+max(1,span*7//20),pages[0]+max(2,span*14//20),pages[-1]+1]
    else:
        boundaries=[pages[0]+(span*i)//4 for i in range(5)]
        boundaries[-1]=pages[-1]+1
    rng=random.Random(seed)
    for attempt in range(300):
        starts=[]; used=[]; local=set()
        for bi in range(bands):
            left,right=boundaries[bi],boundaries[bi+1]
            band=[v for v in pool if left<=v['page']<right]
            rng.shuffle(band)
            picked=None
            for v in band:
                if normalize(v['text']) not in route_seen and normalize(v['text']) not in local and all(abs(v['page']-p)>=3 for p in used): picked=v; break
            if picked:
                starts.append(picked); used.append(picked['page']); local.add(normalize(picked['text']))
        if len(starts)<basic:
            # إذا استنفدت الأسئلة الفريدة في شريحة ضيقة، نسمح بالتكرار بين الاختبارات مع منعه داخل النموذج.
            starts=[]; used=[]; local=set()
            for bi in range(bands):
                left,right=boundaries[bi],boundaries[bi+1]
                band=[v for v in pool if left<=v['page']<right and normalize(v['text']) not in local]
                rng.shuffle(band)
                picked=next((v for v in band if all(abs(v['page']-p)>=3 for p in used)),None)
                if picked: starts.append(picked); used.append(picked['page']); local.add(normalize(picked['text']))
        if len(starts)<basic: continue
        chosen=[]
        for st in starts:
            ends=[v for v in pool if v['page']==st['page']+2 and v['ayah']>1]
            rng.shuffle(ends)
            if not ends: ends=[v for v in pool if v['page']>=st['page']+2 and v['page']<=hi]
            if not ends: ends=[st]
            end=next((e for e in ends if usage[f'{st["surah"]}:{st["ayah"]}-{e["surah"]}:{e["ayah"]}']<12),ends[0])
            key=f'{st["surah"]}:{st["ayah"]}-{end["surah"]}:{end["ayah"]}'
            chosen.append((st,end,key,'أساسي'))
        while len(chosen)<total:
            lo_gap=chosen[0][0]['page']+3; hi_gap=chosen[-1][0]['page']-3
            middle=[v for v in pool if lo_gap<=v['page']<=hi_gap and normalize(v['text']) not in local and all(abs(v['page']-p)>=3 for p in used)]
            rng.shuffle(middle)
            if not middle:
                middle=[v for v in pool if normalize(v['text']) not in local and all(abs(v['page']-p)>=3 for p in used)]
            if not middle: break
            picked=None
            for st in middle:
                ends=[v for v in pool if v['page']==st['page']+2 and v['ayah']>1] or [st]
                for end in ends:
                    key=f'{st["surah"]}:{st["ayah"]}-{end["surah"]}:{end["ayah"]}'
                    if usage[key] < 12:
                        picked=(st,end,key); break
                if picked: break
            if not picked: break
            st,end,key=picked
            chosen.append((st,end,key,'احتياطي')); used.append(st['page']); local.add(normalize(st['text']))
        if len(chosen)==total:
            for _,_,key,_ in chosen: usage[key]+=1
            for st,_,_,_ in chosen: route_seen.add(normalize(st['text']))
            chosen.sort(key=lambda x:(x[0]['page'],x[0]['surah'],x[0]['ayah']))
            return chosen
    # Fallback للنطاقات القصيرة جدًا: نضمن إنشاء النموذج حتى إن لم تسمح المساحة بثلاثة أوجه كاملة.
    fallback=[]; used_pages=[]; local=set()
    for start in sorted(pool, key=lambda v:(v['page'],v['surah'],v['ayah'])):
        key_start=normalize(start['text'])
        if key_start in local: continue
        ends=[v for v in pool if v['page']>=start['page'] and v['ayah']>1] or [start]
        end=ends[min(len(ends)-1, 2)]
        key=f'{start["surah"]}:{start["ayah"]}-{end["surah"]}:{end["ayah"]}'
        fallback.append((start,end,key,'أساسي')); local.add(key_start); used_pages.append(start['page'])
        if len(fallback) == total: break
    if len(fallback) == total:
        for _,_,key,_ in fallback: usage[key]+=1
        for st,_,_,_ in fallback: route_seen.add(normalize(st['text']))
        return fallback
    ordered=sorted(pool,key=lambda v:(v['page'],v['surah'],v['ayah']))
    for min_gap in (3,2,1,0):
        chosen=[]; used_pages=[]; local=set()
        for st in ordered:
            norm=normalize(st['text'])
            if norm in local or any(abs(st['page']-pg)<min_gap for pg in used_pages): continue
            ends=[v for v in pool if v['page']>=st['page']+2 and v['ayah']>1]
            ends=ends or [v for v in pool if v['page']>=st['page'] and v['ayah']>1] or [st]
            end=ends[0]
            key=f'{st["surah"]}:{st["ayah"]}-{end["surah"]}:{end["ayah"]}'
            chosen.append((st,end,key,'أساسي')); used_pages.append(st['page']); local.add(norm)
            if len(chosen)==total:
                for _,_,key,_ in chosen: usage[key]+=1
                for st,_,_,_ in chosen: route_seen.add(normalize(st['text']))
                return chosen
    raise RuntimeError(f'cannot choose {total} questions for pages {lo}-{hi}')

CUSTOM_UNITS=[
 (1,'الفاتحة إلى سيقول السفهاء',(1,1),(2,141),4,3),
 (2,'من سيقول السفهاء إلى نهاية البقرة',(2,142),(2,286),5,4),
 (3,'آل عمران',(3,1),(3,200),5,4),
 (4,'النساء',(4,1),(4,176),5,4),
 (5,'المائدة',(5,1),(5,120),5,4),
 (6,'الأنعام',(6,1),(6,165),5,4),
 (7,'الأعراف',(7,1),(7,206),5,4),
 (8,'الأنفال إلى التوبة',(8,1),(9,129),5,4),
 (9,'يونس إلى هود',(10,1),(11,123),5,4),
 (10,'يوسف إلى إبراهيم',(12,1),(14,52),5,4),
 (11,'الحجر إلى النحل',(15,1),(16,128),4,3),
 (12,'الإسراء إلى الكهف',(17,1),(18,110),4,3),
 (13,'مريم إلى طه',(19,1),(20,135),4,3),
 (14,'الأنبياء إلى الحج',(21,1),(22,78),4,3),
 (15,'المؤمنون إلى النور',(23,1),(24,64),4,3),
 (16,'الفرقان إلى الشعراء',(25,1),(26,227),4,3),
 (17,'النمل إلى القصص',(27,1),(28,88),4,3),
 (18,'العنكبوت إلى السجدة',(29,1),(32,30),4,3),
 (19,'الأحزاب إلى فاطر',(33,1),(35,45),4,3),
 (20,'يس إلى ص',(36,1),(38,88),4,3),
 (21,'الزمر إلى فصلت',(39,1),(41,54),4,3),
 (22,'الشورى إلى الجاثية',(42,1),(45,37),4,3),
 (23,'الأحقاف إلى ق',(46,1),(50,45),4,3),
 (24,'الذاريات إلى الحديد',(51,1),(57,29),4,3),
 (25,'المجادلة إلى التحريم',(58,1),(66,12),4,3),
 (26,'الملك إلى المرسلات',(67,1),(77,50),4,3),
 (27,'عم إلى الناس',(78,1),(114,6),4,3),
]
unit_bounds={}
for uid,name,st,en,new_count,old_count in CUSTOM_UNITS:
    pages=[v['page'] for v in verses if (v['surah'],v['ayah'])>=(st[0],st[1]) and (v['surah'],v['ayah'])<=(en[0],en[1]) and v['page']]
    unit_bounds[uid]=(min(pages),max(pages))
unit_by_id={u[0]:u for u in CUSTOM_UNITS}
JUZ_STARTS=[(1,1),(2,142),(2,253),(3,93),(4,24),(4,148),(5,82),(6,111),(7,88),(8,41),(9,93),(11,6),(12,53),(15,1),(17,1),(18,75),(21,1),(23,1),(25,21),(27,56),(29,46),(33,31),(36,28),(39,32),(41,47),(46,1),(51,31),(58,1),(67,1),(78,1)]
juz_bounds={}
for jid,st in enumerate(JUZ_STARTS,1):
    start_i=index[st]; end_i=(index[JUZ_STARTS[jid]]-1 if jid<30 else len(verses)-1)
    juz_bounds[jid]=(min(verses[i]['page'] for i in range(start_i,end_i+1) if verses[i].get('page')),max(verses[i]['page'] for i in range(start_i,end_i+1) if verses[i].get('page')))

def lower_units_from_start(page):
    first=next((uid for uid,(a,b) in sorted(unit_bounds.items()) if b>=page),27)
    # The lower-direction cumulative system always keeps the first Fatiha-to-
    # beginning-of-Baqara unit when a test starts inside the second Baqara unit.
    if first == 2:
        first = 1
    return list(range(first,28))

def upper_juz_to_page(page):
    return [jid for jid,(a,b) in sorted(juz_bounds.items()) if a<=page]

def lower_units_for_range(lo,hi):
    return [uid for uid,(a,b) in unit_bounds.items() if b>=lo and a<=hi]

if os.environ.get('LOWER_ONLY'):
    for old in OUT.glob('route-*'):
        if old.name >= 'route-24': shutil.rmtree(old)
else:
    if OUT.exists(): shutil.rmtree(OUT)
OUT.mkdir(parents=True, exist_ok=True)
index=[]; global_model=0
route_usage={'from_fatiha':Counter(),'from_nas':Counter()}
for route in CATALOG['routes']:
    if os.environ.get('LOWER_ONLY') and route['direction'] != 'from_nas': continue
    route_dir=OUT/f'route-{route["route_number"]:02d}'
    route_dir.mkdir()
    route_seen=set()
    for test in route['tests']:
        lo,hi=row_bounds(test['raw']); total=row_count(test['raw'])
        foundation = route['route_number'] == 46
        lower_units=lower_units_from_start(lo) if route['direction']=='from_nas' and not foundation else []
        # The category-one/two closing test is the deliberate exception:
        # first Baqara unit, skip the remainder of Baqara, then continue from
        # Al عمران through الناس.
        if route['direction']=='from_nas' and route['route_number'] in range(24,36) and test['test']==len(route['tests']):
            lower_units=[1]+list(range(3,28))
        if route['direction']=='from_nas' and route['route_number'] in range(36,42):
            if test['test']==len(route['tests'])-1:
                lower_units=[1]+list(range(3,28))
            elif test['test']==len(route['tests']):
                lower_units=list(range(1,28))
        upper_juz=upper_juz_to_page(hi) if route['direction']=='from_fatiha' else []
        new_unit=min(lower_units,key=lambda uid:unit_bounds[uid][0]) if lower_units else None
        new_juz=max(upper_juz) if upper_juz else None
        active_parts=lower_units or upper_juz
        display_name=test.get('name') or ('من الناس إلى الفاتحة' if route['direction']=='from_nas' else route['path'])
        models=[]
        for model_no in range(1,26):
            global_model+=1; selected_with_parts=[]; route_seen=set()
            if foundation:
                picked=select_questions(lo,hi,total,20261106+route['route_number']*100000+test['test']*100+model_no,route_seen,route_usage[route['direction']])
                selected_with_parts=[(item,test['test']) for item in picked]
            elif lower_units:
                for uid in sorted(lower_units,key=lambda x:unit_bounds[x][0]):
                    u=unit_by_id[uid]; needed=u[4] if uid==new_unit else u[5]
                    ulo,uhi=unit_bounds[uid]
                    picked=select_questions(ulo,uhi,needed,20261106+route['route_number']*100000+test['test']*100+model_no*31+uid,route_seen,route_usage[route['direction']],basic_count=(u[4]-1 if uid==new_unit else u[5]))
                    selected_with_parts.extend((item,uid) for item in picked)
            elif upper_juz:
                for jid in sorted(upper_juz):
                    needed=4 if jid==new_juz else 3
                    ulo,uhi=juz_bounds[jid]
                    picked=select_questions(ulo,uhi,needed,20261106+route['route_number']*100000+test['test']*100+model_no*31+jid,route_seen,route_usage[route['direction']],basic_count=(3 if jid==new_juz else 3))
                    selected_with_parts.extend((item,jid) for item in picked)
            else:
                picked=select_questions(lo,hi,total,20261106+route['route_number']*100000+test['test']*100+model_no,route_seen,route_usage[route['direction']])
                selected_with_parts=[(item,test['test']) for item in picked]
            preview=[]
            for qi,((start,end,key,kind),part_id) in enumerate(selected_with_parts,1):
                part_name=(display_name if foundation else (unit_by_id[part_id][1] if route['direction']=='from_nas' and part_id in unit_by_id else (f'الجزء {part_id:02d}' if route['direction']=='from_fatiha' else route['path'])))
                preview.append({'uid':f'route-{route["route_number"]:02d}-test-{test["test"]:03d}-model-{model_no:02d}-q-{qi:02d}-{key}','surah':start['surah'],'ayah':start['ayah'],'page':start['page'],'end_surah':end['surah'],'end_ayah':end['ayah'],'end_page':end['page'],'text':qtext(start,end),'kind':kind,'question_no':qi,'part_id':part_id,'part_name':part_name})
            models.append({'id':f'route-{route["route_number"]:02d}-test-{test["test"]:03d}-model-{model_no:02d}','name':f'{display_name} — نموذج {model_no:02d}','direction':route['direction'],'direction_name':route['direction_name'],'category_name':route['category_name'],'path_name':route['path'],'route_number':route['route_number'],'test_number':test['test'],'range_name':display_name,'model_in_range':model_no,'questions':len(preview),'basic':sum(q['kind']=='أساسي' for q in preview),'reserve':sum(q['kind']=='احتياطي' for q in preview),'preview':preview})
        file_name=f'test-{test["test"]:03d}.json'; (route_dir/file_name).write_text(json.dumps(models,ensure_ascii=False,separators=(',',':')),encoding='utf-8')
        index.append({'route_number':route['route_number'],'test_number':test['test'],'file':f'/route-files/route-{route["route_number"]:02d}.json','name':route['path'],'models':25,'questions':total,'page_range':[lo,hi]})
    print(f'route {route["route_number"]:02d}: {route["test_count"]} tests')
(CATALOG and (ROOT/'public'/'route-models-index.json').write_text(json.dumps({'summary':CATALOG['summary'],'tests':index},ensure_ascii=False,indent=2),encoding='utf-8'))
print(json.dumps({'routes':len(CATALOG['routes']),'tests':len(index),'models':global_model},ensure_ascii=False))
