import json, re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
CATALOG = ROOT / 'public' / 'route-catalog-1263.json'
ROUTE_FILES = ROOT / 'public' / 'route-files'
AR = str.maketrans('٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹', '01234567890123456789')
SURAH_ORDER = {
    name: i + 1 for i, name in enumerate(
        'الفاتحة البقرة آل عمران النساء المائدة الأنعام الأعراف الأنفال التوبة يونس هود يوسف الرعد إبراهيم الحجر النحل الإسراء الكهف مريم طه الأنبياء الحج المؤمنون النور الفرقان الشعراء النمل القصص العنكبوت الروم لقمان السجدة الأحزاب سبأ فاطر يس الصافات ص الزمر غافر فصلت الشورى الزخرف الدخان الجاثية الأحقاف محمد الفتح الحجرات ق الذاريات الطور النجم القمر الرحمن الواقعة الحديد المجادلة الحشر الممتحنة الصف الجمعة المنافقون التغابن الطلاق التحريم الملك القلم الحاقة المعارج نوح الجن المزمل المدثر القيامة الإنسان المرسلات النبأ النازعات عبس التكوير الانفطار المطففين الانشقاق البروج الطارق الأعلى الغاشية الفجر البلد الشمس الليل الضحى الشرح التين العلق القدر البينة الزلزلة العاديات القارعة التكاثر العصر الهمزة الفيل قريش الماعون الكوثر الكافرون النصر المسد الإخلاص الفلق الناس'.split()
    )
}

REPLACEMENTS = {
    'إىل': 'إلى', 'اىل': 'إلى', 'األوىل': 'الأولى', 'األوﻻ': 'الأولى',
    'الثاين': 'الثاني', 'الثالث': 'الثالث', 'الرابع': 'الرابع',
    'اخلامس': 'الخامس', 'السادس': 'السادس', 'األجزاء': 'الأجزاء',
    'االجزاء': 'الأجزاء', 'األحزاب': 'الأحزاب', 'االنفال': 'الأنفال',
    'األنفال': 'الأنفال', 'األنعام': 'الأنعام', 'األعراف': 'الأعراف',
    'اإلرساء': 'الإسراء', 'االرساء': 'الإسراء', 'النبإ': 'النبأ',
    'االختبار': 'الاختبار', 'االختبارات': 'الاختبارات', 'يف': 'في',
    'األخري': 'الأخير', 'األخري': 'الأخير', 'األخرية': 'الأخيرة',
    'األجزاء': 'الأجزاء', 'العرش': 'العشر', 'سؤاالن': 'سؤالان',
    'األقصى': 'الأقصى', 'المرسالت': 'المرسلات', 'المطففني': 'المطففين',
    'االنفطار': 'الانفطار', 'االرساء': 'الإسراء', 'االحقاف': 'الأحقاف',
    'الألنفال': 'الأنفال', 'الألحزاب': 'الأحزاب', 'الألنعام': 'الأنعام',
    'الألعراف': 'الأعراف', 'الألحقاف': 'الأحقاف', 'المطففني': 'المطففين',
    'النبإ': 'النبأ', 'الشوري': 'الشورى', 'الذار يات': 'الذاريات',
    'التحر يم': 'التحريم', 'المؤ منون': 'المؤمنون', 'المرسالت': 'المرسلات',
    'االنفال': 'الأنفال', 'التوبة واألنفال': 'الأنفال والتوبة',
}

def clean(s: str) -> str:
    s = (s or '').translate(AR)
    for a, b in REPLACEMENTS.items():
        s = s.replace(a, b)
    s = re.sub(r'\s+', ' ', s).strip(' -–')
    s = re.sub(r'([ء-ي]+)\s*(\d+)\s*إلى\s*([ء-ي]+)\s*(\d+)', r'\1 آية \2 إلى \3 آية \4', s)
    s = re.sub(r'\s+([،:؛])', r'\1', s)
    return s

def extract_name(route, test):
    raw = clean(test.get('raw', ''))
    # الاسم المعتمد هو عمود «نطاق الاختبار» نفسه في ملف المصدر، وليس اسم
    # المجلد ولا اسمًا مُعادَ صياغته حسب أعلى سورة. نلتقطه قبل عمود الصفحات.
    m = re.search(r'^(من\s+.+?)(?=\s+\d+\s*[–-]\s*\d+\b)', raw)
    name = clean(m.group(1)) if m else clean(raw.split('  ')[0])
    # جعل أرقام الآيات واضحة في الاسم مع الحفاظ على نفس البداية والنهاية.
    name = re.sub(r'^(من\s+.+?\s+إلى\s+[^\d]+?)\s+(\d+)$', r'\1 آية \2', name)
    # في الفئة الرابعة والتأسيس يضع المصدر البداية والنهاية التفصيليتين
    # في عمود اسم المجلد؛ نعتمدهما حتى لا يضيع رقم الآية.
    if int(route['route_number']) >= 42:
        folder = re.search(r'-\s*\d+\s*(من\s+.+?)(?=\s+\d+\s*$)', raw)
        if folder:
            name = clean(folder.group(1))
            name = re.sub(r'\s+(\d+)\s+إلى\s+', r' آية \1 إلى ', name)
            name = re.sub(r'\s+(\d+)$', r' آية \1', name)
            name = name.replace('آية آية', 'آية')
    return name

def main():
    catalog = json.loads(CATALOG.read_text(encoding='utf-8'))
    catalog['summary'] = {'routes': 46, 'tests': 1263, 'models': 31575, 'models_per_test': 25}
    for route in catalog['routes']:
        route['category_name'] = clean(route.get('category_name', ''))
        route['path'] = clean(route.get('path', ''))
        if int(route['route_number']) == 46:
            route['category_name'] = 'فئة التأسيس'
            route['path'] = 'مسار التأسيس'
        route['direction_name'] = ('من الفاتحة إلى الناس' if route['direction'] == 'from_fatiha' else 'من الناس إلى الفاتحة')
        for test in route['tests']:
            name = extract_name(route, test)
            name = re.sub(r'^من\s+من\s+', 'من ', name)
            name = re.sub(r'^من\s+سورة\s+', 'من ', name)
            name = re.sub(r'(آل عمران|البقرة|النساء|المائدة|الأنعام|الأعراف|الأنفال|التوبة|الفجر|النبأ|المطففين)\s*(\d+)\s*إلى\s*\1\s*(\d+)\s*إلى الناس', r'\1 آية \2 إلى \1 آية \3 إلى الناس', name)
            test['name'] = name
            test['display_name'] = name
            test['raw_normalized'] = clean(test.get('raw', ''))
    for filename in ('route-catalog-1263.json', 'route-catalog.json'):
        (ROOT / 'public' / filename).write_text(json.dumps(catalog, ensure_ascii=False, indent=2), encoding='utf-8')
    changed = 0
    for route in catalog['routes']:
        path = ROUTE_FILES / f"route-{int(route['route_number']):02d}.json"
        if not path.exists():
            continue
        tests = json.loads(path.read_text(encoding='utf-8'))
        by_test = {int(t['test']): t for t in route['tests']}
        for entry in tests:
            entry_test = entry.get('test', entry.get('test_number'))
            source = by_test.get(int(entry_test)) if entry_test is not None else None
            if not source:
                continue
            name = source['name']
            models = entry.get('models', []) if isinstance(entry.get('models', []), list) else []
            if 'preview' in entry:
                models = [entry]
            for model in models:
                model['test_name'] = name
                model['range_name'] = name
                model['name'] = f"{name} — نموذج {int(model.get('model_in_range', 0)):02d}"
                model['category_name'] = route['category_name']
                model['path_name'] = route['path']
                model['direction_name'] = route['direction_name']
            changed += 1
        path.write_text(json.dumps(tests, ensure_ascii=False, separators=(',', ':')), encoding='utf-8')
    print(json.dumps({'routes': len(catalog['routes']), 'tests': sum(len(r['tests']) for r in catalog['routes']), 'updated_route_tests': changed}, ensure_ascii=False))

if __name__ == '__main__':
    main()
