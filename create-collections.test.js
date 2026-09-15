const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const os = require('os');
const path = require('path');
const writeXlsxFile = require('write-excel-file/node');

const {
  buildManualCollectionUrl,
  loadCollectionsFromWorkbook,
  normalizeCollectionsUrl,
  parseArgs,
  resolveCollectionName,
} = require('./create-collections');

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
