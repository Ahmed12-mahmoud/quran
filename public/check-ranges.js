/*
 * فحص ملفات أسئلة القرآن من نفس الموقع.
 * افتح الموقع ثم افتح أدوات المطور Console والصق هذا الملف أو محتواه.
 */
(async () => {
  const base = location.origin;
  const results = [];

  for (let range = 1; range <= 30; range++) {
    const name = `range-${String(range).padStart(2, '0')}.json`;
    const url = `${base}/ranges/${name}`;

    try {
      const response = await fetch(url, { cache: 'no-store' });
      if (!response.ok) throw new Error(`HTTP ${response.status}`);

      const models = await response.json();
      if (!Array.isArray(models)) throw new Error('الملف ليس قائمة نماذج');

      const juzCounts = {};
      let questionCount = 0;
      for (const model of models) {
        for (const question of model.preview || []) {
          questionCount++;
          juzCounts[question.juz] = (juzCounts[question.juz] || 0) + 1;
        }
      }

      results.push({
        range,
        file: name,
        ok: models.length === 25,
        models: models.length,
        questions: questionCount,
        juz: Object.keys(juzCounts).sort((a, b) => Number(a) - Number(b)).join(', '),
        details: juzCounts
      });
    } catch (error) {
      results.push({ range, file: name, ok: false, error: error.message });
    }
  }

  console.table(results.map(({ details, ...row }) => row));
  const failed = results.filter(row => !row.ok);
  console.log(
    failed.length === 0
      ? '✅ كل ملفات النطاقات الثلاثين تعمل، وكل ملف يحتوي على 25 نموذجًا.'
      : `❌ يوجد ${failed.length} ملف/ملفات بها مشكلة. راجع الجدول أعلاه.`
  );
  window.quranRangeCheck = results;
})();
