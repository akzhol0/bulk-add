const { chromium } = require('playwright');
const readXlsxFile = require('read-excel-file/node');
const fs = require('fs');
const path = require('path');
const {
  createRunArtifacts,
  installConsoleTranscript,
  localStamp,
  writeExcelReport,
} = require('./reporting');

const programs = require('./programs.json');
const config = require('./config.json');

const profileDir = path.join(__dirname, '.playwright-profile');
const defaultExpectedFile = path.join(__dirname, 'All courses.xlsx');

const ADD_CONTENT_NAMES = ['Add content', 'Добавить материалы'];

function normalizeText(value) {
  return String(value ?? '').replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim();
}

function courseTitleKey(value) {
  return normalizeText(value)
    .normalize('NFKC')
    .replace(/[-‐‑‒–—―]/g, ' ')
    .replace(/\s+/g, ' ')
    .toLocaleLowerCase();
}

function normalizedTextRegex(text) {
  const parts = normalizeText(text)
    .split(/\s+/)
    .map((part) => part.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'));
  return new RegExp(`^\\s*${parts.join('\\s+')}\\s*$`, 'u');
}

function parseArgs(argv) {
  const opts = { slugs: [], limit: null, file: defaultExpectedFile };
  for (const arg of argv) {
    if (arg.startsWith('--slug=')) opts.slugs.push(arg.slice('--slug='.length));
    else if (arg.startsWith('--limit=')) opts.limit = Number(arg.slice('--limit='.length));
    else if (arg.startsWith('--file=')) opts.file = path.resolve(arg.slice('--file='.length));
  }
  return opts;
}

async function loadExpectedCourses(filePath) {
  if (!fs.existsSync(filePath)) throw new Error(`Не найден файл с эталоном: ${filePath}`);

  const workbook = await readXlsxFile(filePath);
  const rows = Array.isArray(workbook[0]?.data) ? workbook[0].data : workbook;
  if (rows.length < 2) throw new Error(`В файле нет строк с курсами: ${filePath}`);

  const expected = rows.slice(1)
    .map((row, index) => {
      const title = normalizeText(row[0]);
      let provider = normalizeText(row[1]);

      // В текущем Excel последнее слово обрезано на одну букву. На Coursera оно отображается как "project".
      if (provider === 'Participant in the Ministry of Science and Higher Education projec') {
        provider += 't';
      }

      return { title, provider, excelRow: index + 2 };
    })
    .filter((item) => item.title || item.provider);

  const incomplete = expected.filter((item) => !item.title || !item.provider);
  if (incomplete.length > 0) {
    throw new Error(`В Excel есть строки без названия или провайдера: ${incomplete.map((x) => x.excelRow).join(', ')}`);
  }

  const seenTitles = new Set();
  const duplicates = [];
  for (const item of expected) {
    const titleKey = courseTitleKey(item.title);
    if (seenTitles.has(titleKey)) duplicates.push(item.title);
    seenTitles.add(titleKey);
  }
  if (duplicates.length > 0) {
    throw new Error(`В Excel повторяются названия курсов: ${[...new Set(duplicates)].join(' | ')}`);
  }

  return expected;
}

async function firstVisible(locator) {
  const count = await locator.count();
  for (let i = 0; i < count; i++) {
    const item = locator.nth(i);
    if (await item.isVisible().catch(() => false)) return item;
  }
  return null;
}

async function clickExactNamed(page, role, name, timeout) {
  const regex = normalizedTextRegex(name);
  const candidates = [page.getByRole(role, { name: regex }), page.getByText(regex)];
  const start = Date.now();

  while (Date.now() - start < timeout) {
    for (const candidate of candidates) {
      const visible = await firstVisible(candidate);
      if (visible) {
        await visible.click({ timeout: Math.max(1, timeout - (Date.now() - start)) });
        return;
      }
    }
    await page.waitForTimeout(250);
  }

  throw new Error(`Не найден элемент с точным названием: ${name}`);
}

async function waitForAnyNamed(page, roles, names, timeout) {
  const start = Date.now();
  while (Date.now() - start < timeout) {
    for (const role of roles) {
      for (const name of names) {
        const visible = await firstVisible(
          page.getByRole(role, { name: normalizedTextRegex(name) }),
        );
        if (visible) return visible;
      }
    }
    await page.waitForTimeout(250);
  }
  throw new Error(`Не появился элемент: ${names.join(' / ')}`);
}

async function gotoWithRetry(page, url, timeout) {
  let lastError;
  for (let attempt = 1; attempt <= 3; attempt++) {
    try {
      await page.goto(url, { waitUntil: 'domcontentloaded', timeout });
      return;
    } catch (error) {
      lastError = error;
      const retryable = /ERR_ABORTED|frame was detached|Navigation interrupted/i.test(error?.message || '');
      if (!retryable || attempt === 3) throw error;
      await page.waitForTimeout(500 * attempt);
    }
  }
  throw lastError;
}

async function readCollectionCourseCount(page, timeout = 5000) {
  const start = Date.now();
  while (Date.now() - start < timeout) {
    const headings = await page.locator('h2').allTextContents();
    for (const heading of headings) {
      const match = normalizeText(heading).match(/^([\d\s]+)\s+(?:курс(?:ов|а)?|courses?)$/iu);
      if (match) return Number(match[1].replace(/\s/g, ''));
    }
    await page.waitForTimeout(200);
  }
  return null;
}

async function collectCollectionCourses(page) {
  const found = new Map();
  await page.evaluate(() => window.scrollTo(0, 0));

  let previousHeight = 0;
  let stableBottomPasses = 0;

  for (let pass = 0; pass < 200; pass++) {
    const items = page.locator('li[data-e2e="collection-list-item"]');
    const itemCount = await items.count();

    for (let i = 0; i < itemCount; i++) {
      const item = items.nth(i);
      const titleLocator = item.locator('h3').first();
      const title = normalizeText(await titleLocator.evaluate((node) => {
        const clone = node.cloneNode(true);
        clone.querySelectorAll('[data-testid="tag-root"]').forEach((tag) => tag.remove());
        return clone.textContent;
      }).catch(() => ''));
      if (!title) continue;

      const provider = normalizeText(
        await item.locator('p').first().textContent().catch(() => ''),
      );
      found.set(title, { title, provider });
    }

    const state = await page.evaluate(() => ({
      height: document.documentElement.scrollHeight,
      viewport: window.innerHeight,
      y: window.scrollY,
    }));
    const atBottom = state.y + state.viewport >= state.height - 2;

    if (atBottom && state.height === previousHeight) stableBottomPasses += 1;
    else stableBottomPasses = 0;
    if (stableBottomPasses >= 3) break;

    previousHeight = state.height;
    await page.evaluate(() => window.scrollBy(0, Math.max(400, window.innerHeight * 0.8)));
    await page.waitForTimeout(180);
  }

  await page.evaluate(() => window.scrollTo(0, 0));
  return found;
}

function compareCourses(expected, actual, pageCount) {
  const expectedByTitle = new Map(expected.map((item) => [courseTitleKey(item.title), item]));
  const actualByTitle = new Map(
    [...actual.values()].map((item) => [courseTitleKey(item.title), item]),
  );
  const missing = [];
  const extra = [];
  const providerMismatches = [];

  for (const item of expected) {
    const actualItem = actualByTitle.get(courseTitleKey(item.title));
    if (!actualItem) {
      missing.push(`${item.title} — ${item.provider}`);
      continue;
    }

    if (actualItem.provider !== item.provider) {
      providerMismatches.push(
        `${item.title} | ожидался: ${item.provider} | на Coursera: ${actualItem.provider || '[не найден]'}`,
      );
    }
  }

  for (const item of actual.values()) {
    if (!expectedByTitle.has(courseTitleKey(item.title))) {
      extra.push(`${item.title} — ${item.provider || '[провайдер не найден]'}`);
    }
  }

  const matchedCount = expected.length - missing.length - providerMismatches.length;

  const matches = missing.length === 0
    && extra.length === 0
    && providerMismatches.length === 0
    && pageCount === expected.length
    && actual.size === expected.length;

  return {
    status: matches ? 'match' : 'mismatch',
    expected_count: expected.length,
    page_count: pageCount,
    found_unique: actual.size,
    matched_count: matchedCount,
    missing_count: missing.length,
    extra_count: extra.length,
    provider_mismatch_count: providerMismatches.length,
    missing_courses: missing.join(' | '),
    extra_courses: extra.join(' | '),
    provider_mismatches: providerMismatches.join(' || '),
  };
}

async function auditProgram(page, program, expected) {
  const actionTimeout = config.actionTimeoutMs || 20000;
  const collectionsUrl = program.url.replace(/\/main\/?$/, '/catalog/collections');

  await gotoWithRetry(page, collectionsUrl, config.navigationTimeoutMs || 60000);
  if (/login|auth|signin/i.test(page.url())) {
    throw new Error('Coursera запросила повторный вход. Запустите npm run login и войдите снова.');
  }

  await clickExactNamed(page, 'link', config.targetCollection, actionTimeout);
  await waitForAnyNamed(page, ['button', 'link'], ADD_CONTENT_NAMES, actionTimeout);

  const pageCount = await readCollectionCourseCount(page);
  const actual = await collectCollectionCourses(page);
  return compareCourses(expected, actual, pageCount);
}

(async () => {
  const artifacts = createRunArtifacts('audit');
  installConsoleTranscript(artifacts.consoleLogPath);

  const opts = parseArgs(process.argv.slice(2));
  const expected = await loadExpectedCourses(opts.file);

  let selectedPrograms = programs;
  if (opts.slugs.length > 0) {
    const requested = new Set(opts.slugs);
    selectedPrograms = selectedPrograms.filter((program) => requested.has(program.slug));
    const missingSlugs = opts.slugs.filter((slug) => !selectedPrograms.some((p) => p.slug === slug));
    if (missingSlugs.length > 0) throw new Error(`Slug не найден в programs.json: ${missingSlugs.join(', ')}`);
  }
  if (Number.isFinite(opts.limit) && opts.limit > 0) selectedPrograms = selectedPrograms.slice(0, opts.limit);

  console.log(`Запуск аудита: ${new Date().toLocaleString('ru-RU')}`);
  console.log(`Эталон: ${opts.file}`);
  console.log(`Пар «курс + провайдер»: ${expected.length}`);
  console.log(`Коллекция: ${config.targetCollection}`);
  console.log(`Программ к проверке: ${selectedPrograms.length}`);
  console.log('Режим: ТОЛЬКО ЧТЕНИЕ');
  console.log(`Папка запуска: ${artifacts.runDir}`);
  console.log(`Лог консоли: ${artifacts.consoleLogPath}`);

  const context = await chromium.launchPersistentContext(profileDir, {
    channel: 'chrome',
    headless: false,
    viewport: null,
    args: ['--start-maximized'],
  });
  const page = await context.newPage();
  page.setDefaultTimeout(config.actionTimeoutMs || 20000);
  page.setDefaultNavigationTimeout(config.navigationTimeoutMs || 60000);

  const results = [];
  let fatalAuthError = false;

  for (let i = 0; i < selectedPrograms.length; i++) {
    const program = selectedPrograms[i];
    process.stdout.write(`[${i + 1}/${selectedPrograms.length}] ${program.name} ... `);

    try {
      const result = await auditProgram(page, program, expected);
      const row = { ...program, ...result };
      results.push(row);
      console.log(
        `${result.status} | на странице: ${result.page_count ?? '?'} | считано: ${result.found_unique}`
        + ` | совпало: ${result.matched_count}/${result.expected_count}`
        + ` | нет: ${result.missing_count} | лишних: ${result.extra_count}`
        + ` | неверный провайдер: ${result.provider_mismatch_count}`,
      );
    } catch (error) {
      const message = error?.message || String(error);
      const screenshotPath = path.join(
        artifacts.screenshotsDir,
        `${String(i + 1).padStart(3, '0')}-${program.slug}-${localStamp()}.png`,
      );
      await page.screenshot({ path: screenshotPath, fullPage: true }).catch(() => {});
      results.push({ ...program, status: 'error', message, screenshot: screenshotPath });
      console.log(`ERROR: ${message}`);

      if (/повторный вход|login|auth|signin/i.test(message)) {
        fatalAuthError = true;
        break;
      }
    }
  }

  const header = [
    { key: 'name', label: 'Программа', width: 38 },
    { key: 'slug', label: 'Slug', width: 28 },
    { key: 'url', label: 'URL', width: 55 },
    { key: 'status', label: 'Статус', width: 16 },
    { key: 'expected_count', label: 'Ожидалось', width: 14 },
    { key: 'page_count', label: 'Указано на странице', width: 20 },
    { key: 'found_unique', label: 'Считано уникальных', width: 20 },
    { key: 'matched_count', label: 'Совпало', width: 14 },
    { key: 'missing_count', label: 'Не хватает', width: 14 },
    { key: 'extra_count', label: 'Лишних', width: 12 },
    { key: 'provider_mismatch_count', label: 'Неверных провайдеров', width: 22 },
    { key: 'missing_courses', label: 'Недостающие курсы', width: 60 },
    { key: 'extra_courses', label: 'Лишние курсы', width: 60 },
    { key: 'provider_mismatches', label: 'Ошибки провайдеров', width: 60 },
    { key: 'message', label: 'Ошибка', width: 60 },
    { key: 'screenshot', label: 'Скриншот ошибки', width: 55 },
  ];
  await writeExcelReport(artifacts.reportPath, header, results, 'Аудит');

  const totals = results.reduce((acc, row) => {
    acc[row.status] = (acc[row.status] || 0) + 1;
    return acc;
  }, {});

  console.log('\nИтог:', totals);
  console.log(`Отчёт Excel: ${artifacts.reportPath}`);
  console.log(`Лог консоли UTF-8: ${artifacts.consoleLogPath}`);
  if (fatalAuthError) console.log('Остановка: сессия Coursera истекла. Повторите npm run login.');
  if (totals.error || totals.mismatch) process.exitCode = 1;

  await context.close();
})().catch((error) => {
  console.error('\nFATAL:', error);
  process.exitCode = 1;
});
