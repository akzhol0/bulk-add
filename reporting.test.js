const test = require('node:test');
const assert = require('node:assert/strict');
const { spawnSync } = require('node:child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');
const readXlsxFile = require('read-excel-file/node');

const { localStamp, writeExcelReport } = require('./reporting');

test('writes a readable Excel report with Cyrillic text', async () => {
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'coursera-report-'));
  const reportPath = path.join(tempDir, 'report.xlsx');

  try {
    await writeExcelReport(
      reportPath,
      [
        { key: 'name', label: 'Программа' },
        { key: 'status', label: 'Статус' },
      ],
      [{ name: 'Астана', status: 'Проверено' }],
      'Аудит',
    );

    const workbook = await readXlsxFile(reportPath);
    assert.deepEqual(workbook[0].data[0], ['Программа', 'Статус']);
    assert.deepEqual(workbook[0].data[1], ['Астана', 'Проверено']);
  } finally {
    fs.rmSync(tempDir, { recursive: true, force: true });
  }
});

test('writes console transcript as UTF-8 with BOM', () => {
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'coursera-log-'));
  const logPath = path.join(tempDir, 'console.txt');
  const reportingPath = path.join(__dirname, 'reporting.js');

  try {
    const script = [
      `const reporting = require(${JSON.stringify(reportingPath)});`,
      `reporting.installConsoleTranscript(${JSON.stringify(logPath)});`,
      "console.log('Кириллица: аудит и добавление');",
    ].join('');
    const child = spawnSync(process.execPath, ['-e', script], { encoding: 'utf8' });
    assert.equal(child.status, 0, child.stderr);

    const bytes = fs.readFileSync(logPath);
    assert.deepEqual([...bytes.subarray(0, 3)], [0xEF, 0xBB, 0xBF]);
    assert.match(bytes.toString('utf8'), /Кириллица: аудит и добавление/);
  } finally {
    fs.rmSync(tempDir, { recursive: true, force: true });
  }
});

test('uses a Windows-safe local date and time in folder names', () => {
  assert.match(localStamp(new Date(2026, 7, 20, 16, 7, 20, 104)), /^2026-08-20_16-07-20-104$/);
});
