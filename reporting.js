const fs = require('fs');
const path = require('path');
const writeExcelFile = require('write-excel-file/node');

const reportsDir = path.join(__dirname, 'reports');

function localStamp(date = new Date()) {
  const pad = (value, length = 2) => String(value).padStart(length, '0');
  return [
    date.getFullYear(),
    pad(date.getMonth() + 1),
    pad(date.getDate()),
  ].join('-') + '_'
    + [pad(date.getHours()), pad(date.getMinutes()), pad(date.getSeconds())].join('-')
    + `-${pad(date.getMilliseconds(), 3)}`;
}

function createRunArtifacts(kind) {
  const prefix = kind === 'audit'
    ? 'АУДИТ'
    : kind === 'collections'
      ? 'СОЗДАНИЕ_КОЛЛЕКЦИЙ'
      : 'ДОБАВЛЕНИЕ';
  const runDir = path.join(reportsDir, `${prefix}_${localStamp()}`);
  const screenshotsDir = path.join(runDir, 'screenshots');

  fs.mkdirSync(screenshotsDir, { recursive: true });

  return {
    runDir,
    screenshotsDir,
    consoleLogPath: path.join(runDir, 'console.txt'),
    reportPath: path.join(runDir, 'report.xlsx'),
  };
}

function installConsoleTranscript(logPath) {
  // BOM помогает старым версиям Блокнота и другим Windows-программам сразу распознать UTF-8.
  fs.writeFileSync(logPath, '\uFEFF', 'utf8');
  const fd = fs.openSync(logPath, 'a');
  const originalStdoutWrite = process.stdout.write.bind(process.stdout);
  const originalStderrWrite = process.stderr.write.bind(process.stderr);
  const stripAnsi = (text) => text.replace(/\x1B\[[0-?]*[ -/]*[@-~]/g, '');

  function mirror(originalWrite) {
    return (chunk, encoding, callback) => {
      try {
        const value = Buffer.isBuffer(chunk)
          ? chunk.toString(typeof encoding === 'string' ? encoding : 'utf8')
          : String(chunk);
        fs.writeSync(fd, stripAnsi(value), null, 'utf8');
      } catch {
        // Ошибка записи TXT не должна останавливать обработку Coursera.
      }
      return originalWrite(chunk, encoding, callback);
    };
  }

  process.stdout.write = mirror(originalStdoutWrite);
  process.stderr.write = mirror(originalStderrWrite);
  process.once('exit', () => {
    try {
      fs.closeSync(fd);
    } catch {
      // Файл мог быть уже закрыт операционной системой.
    }
  });
}

function excelCell(value, header = false) {
  const normalized = value === undefined || value === null ? '' : value;
  return {
    value: normalized,
    wrap: true,
    verticalAlign: 'top',
    ...(header ? {
      fontWeight: 'bold',
      backgroundColor: '#D9EAF7',
    } : {}),
  };
}

function calculateColumnWidth(header, rows, key) {
  let longest = String(header).length;
  for (const row of rows) {
    const text = String(row[key] ?? '');
    const longestLine = text.split(/\r?\n/).reduce((max, line) => Math.max(max, line.length), 0);
    longest = Math.max(longest, longestLine);
  }
  return Math.max(12, Math.min(60, longest + 2));
}

async function writeExcelReport(reportPath, header, rows, sheetName) {
  const data = [
    header.map((item) => excelCell(item.label, true)),
    ...rows.map((row) => header.map((item) => excelCell(row[item.key] ?? ''))),
  ];
  const columns = header.map((item) => ({
    width: item.width || calculateColumnWidth(item.label, rows, item.key),
  }));

  await writeExcelFile(data, {
    columns,
    sheet: sheetName,
    stickyRowsCount: 1,
  }).toFile(reportPath);
}

module.exports = {
  createRunArtifacts,
  installConsoleTranscript,
  localStamp,
  writeExcelReport,
};
