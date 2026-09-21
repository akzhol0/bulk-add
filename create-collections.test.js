const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const os = require('os');
const path = require('path');
const writeXlsxFile = require('write-excel-file/node');

const {
  buildManualCollectionUrl,
  buildCourseSearchAttempts,
  loadCollectionsFromWorkbook,
  normalizeCollectionsUrl,
  parseArgs,
  resolveCourseSearchMaxPages,
  resolveCollectionName,
  stripTrailingCourseTypeLabel,
} = require('./create-collections');

test('limits course search pagination to a small configurable number of pages', () => {
  assert.equal(resolveCourseSearchMaxPages(2), 2);
  assert.equal(resolveCourseSearchMaxPages(1), 1);
  assert.equal(resolveCourseSearchMaxPages(0), 2);
  assert.equal(resolveCourseSearchMaxPages(100.5), 2);
  assert.equal(resolveCourseSearchMaxPages(undefined), 2);
});

test('removes trailing course type labels only for fallback search', () => {
  assert.equal(
    stripTrailingCourseTypeLabel('Geographic Information Systems (GIS) Specialization'),
    'Geographic Information Systems (GIS)',
  );
  assert.equal(stripTrailingCourseTypeLabel('Название курса — Специализация'), 'Название курса');
  assert.equal(stripTrailingCourseTypeLabel('Название курса (Course)'), 'Название курса');
  assert.equal(stripTrailingCourseTypeLabel('Название курса Курс'), 'Название курса');
  assert.equal(stripTrailingCourseTypeLabel('Crash Course on Python'), 'Crash Course on Python');
  assert.equal(stripTrailingCourseTypeLabel('Course Design Basics'), 'Course Design Basics');
});

test('tries the Excel title before the title without a trailing type label', () => {
  const attempts = buildCourseSearchAttempts('Example Title Specialization');
  assert.deepEqual(attempts[0], {
    exactName: 'Example Title Specialization',
    queries: ['Example Title Specialization'],
  });
  assert.equal(attempts[1].exactName, 'Example Title');
  assert.deepEqual(attempts[1].queries, ['Example Title', 'Title']);
});

test('normalizes a Coursera program URL to its collections page', () => {
  assert.equal(
    normalizeCollectionsUrl(
      'https://www.coursera.org/o/example/admin/programs/program-123/main?x=1#top',
    ),
    'https://www.coursera.org/o/example/admin/programs/program-123/catalog/collections',
  );
});

test('builds the manual collection form URL used by Coursera', () => {
  assert.equal(
    buildManualCollectionUrl(
      'https://www.coursera.org/o/example/admin/programs/program-123/catalog/collections',
    ),
    'https://www.coursera.org/o/example/program-creation/program-123/new-collection?continueTo=catalog',
  );
});

test('parses collection creation command options', () => {
  assert.deepEqual(
    parseArgs([
      '--url=https://www.coursera.org/o/example/admin/programs/program-123/catalog/collections',
      '--file=source.xlsx',
      '--dry-run',
      '--limit-collections=1',
    ]),
    {
      url: 'https://www.coursera.org/o/example/admin/programs/program-123/catalog/collections',
      file: 'source.xlsx',
      names: null,
      dryRun: true,
      reorderOnly: false,
      limitCollections: 1,
    },
  );

  assert.throws(
    () => parseArgs(['--limit-collections=0']),
    /целым числом от 1/,
  );
});

test('uses an override when available and otherwise accepts a 31-character sheet name', () => {
  const truncated = '1234567890123456789012345678901';
  assert.equal(resolveCollectionName(truncated), truncated);
  assert.equal(
    resolveCollectionName(truncated, { [truncated]: 'Полное название коллекции' }),
    'Полное название коллекции',
  );
});

test('reads every sheet as a collection and uses only the exact course column', async () => {
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'coursera-collections-'));
  const workbookPath = path.join(tempDir, 'input.xlsx');

  try {
    await writeXlsxFile([
      {
        sheet: 'Энергетика',
        data: [
          [{ value: 'Примечание' }, { value: 'Онлайн-курсы Coursera' }],
          [{ value: 'не курс' }, { value: 'Course A' }],
          [{ value: 'не курс' }, { value: 'Course B' }],
        ],
      },
      {
        sheet: 'Геология',
        data: [
          [{ value: 'Онлайн-курсы Coursera' }],
          [{ value: 'Course C' }],
        ],
      },
    ]).toFile(workbookPath);

    assert.deepEqual(await loadCollectionsFromWorkbook(workbookPath), [
      {
        name: 'Энергетика',
        courses: [
          { name: 'Course A', excelRow: 2 },
          { name: 'Course B', excelRow: 3 },
        ],
      },
      {
        name: 'Геология',
        courses: [{ name: 'Course C', excelRow: 2 }],
      },
    ]);
  } finally {
    fs.rmSync(tempDir, { recursive: true, force: true });
  }
});
