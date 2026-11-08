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
    if route['direction'] == 'from_fatiha':
        # The first field in each PDF row is the named cumulative exam range.
        m = re.search(r'(من الفاتحة إلى .*?)(?=\s+\d+[–-]\d+\b)', raw)
        if m:
            return clean(m.group(1))
        m = re.match(r'(من الفاتحة إلى .*?)(?:\s+\d+\s*)?$', raw)
        return clean(m.group(1) if m else raw)
    # For the people-to-fatiha direction, the PDF's folder-name column is the
    # canonical exam name and differs by category/path.
    m = re.search(r'-\s*\d+\s*(.*?)(?:\s+\d+\s*)$', raw)
    if m:
        name = clean(m.group(1))
        if name:
            # In the ascending direction the canonical exam name is the
            # endpoint of the folder range back down to الناس.
            if ' إلى ' in name:
                endpoints = [x.strip() for x in name[3:].split(' إلى ')]
                # أعلى سورة هي الأقرب إلى الفاتحة؛ use the smallest surah number.
                endpoint = min(endpoints, key=lambda x: SURAH_ORDER.get(re.sub(r'\s+آية\s+\d+$', '', x), 999))
                return clean(f'من {endpoint} إلى الناس')
            if ' و' in name or ' و ' in name:
                parts = [x.strip() for x in re.split(r'\s+و\s*', name)]
                endpoint = min(parts, key=lambda x: SURAH_ORDER.get(re.sub(r'\s+آية\s+\d+$', '', x), 999))
                return clean(f'من {endpoint} إلى الناس')
            if name.startswith('جزء '):
                return clean(f'من {name[4:]} إلى الناس')
            return clean(f'من {name} إلى الناس')
    return clean(raw.split('  ')[0])

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
            if int(route['route_number']) in range(24, 36) and int(test['test']) == len(route['tests']):
                name = 'من أول جزء في البقرة إلى الناس'
            if int(route['route_number']) in range(36, 42) and int(test['test']) == len(route['tests']) - 1:
                name = 'من أول جزء في البقرة إلى الناس'
            if int(route['route_number']) in range(36, 42) and int(test['test']) == len(route['tests']):
                name = 'من الفاتحة إلى الناس'
            if int(route['route_number']) == 46 and 'إلى الفاتحة' in clean(test.get('raw', '')):
                name = 'من الفاتحة إلى الناس'
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
