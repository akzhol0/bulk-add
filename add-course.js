const { chromium } = require('playwright');
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

const UI_TEXT = {
  addContent: ['Add content', '\u0414\u043e\u0431\u0430\u0432\u0438\u0442\u044c \u043c\u0430\u0442\u0435\u0440\u0438\u0430\u043b\u044b'],
  search: ['Search', '\u041f\u043e\u0438\u0441\u043a'],
  select: ['Select', '\u0412\u044b\u0431\u0440\u0430\u0442\u044c'],
  nextPage: ['Go to next page', '\u041f\u0435\u0440\u0435\u0439\u0442\u0438 \u043d\u0430 \u0441\u043b\u0435\u0434\u0443\u044e\u0449\u0443\u044e \u0441\u0442\u0440\u0430\u043d\u0438\u0446\u0443'],
};

function parseArgs(argv) {
  const opts = { dryRun: false, slugs: [], limit: null, course: null };
  for (const arg of argv) {
    if (arg === '--dry-run') opts.dryRun = true;
    else if (arg.startsWith('--slug=')) opts.slugs.push(arg.slice('--slug='.length));
    else if (arg.startsWith('--limit=')) opts.limit = Number(arg.slice('--limit='.length));
    else if (arg.startsWith('--course=')) opts.course = arg.slice('--course='.length);
    else if (!arg.startsWith('--') && !opts.course) opts.course = arg;
  }
  return opts;
}

function normalizedTextRegex(text) {
  const parts = text.trim().split(/\s+/).map((part) => part.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'));
  return new RegExp(`^\\s*${parts.join('\\s+')}\\s*$`, 'u');
}

async function firstVisible(locator) {
  const count = await locator.count();
  for (let i = 0; i < count; i++) {
    const item = locator.nth(i);
    if (await item.isVisible().catch(() => false)) return item;
  }
  return null;
}

async function clickNamed(page, { role, name, timeout }) {
  const regex = normalizedTextRegex(name);
  const candidates = [];
  if (role) candidates.push(page.getByRole(role, { name: regex }));
  candidates.push(page.getByText(regex));

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
  throw new Error(`Не найден кликабельный элемент: ${name}`);
}

async function waitForNamed(page, { role, name, timeout }) {
  const regex = normalizedTextRegex(name);
  const candidates = [];
  if (role) candidates.push(page.getByRole(role, { name: regex }));
  candidates.push(page.getByText(regex));

  const start = Date.now();
  while (Date.now() - start < timeout) {
    for (const candidate of candidates) {
      const visible = await firstVisible(candidate);
      if (visible) return visible;
    }
    await page.waitForTimeout(250);
  }
  throw new Error(`Не появился элемент: ${name}`);
}

async function clickNamedAny(page, { role, names, timeout }) {
  const start = Date.now();
  while (Date.now() - start < timeout) {
    for (const name of names) {
      const regex = normalizedTextRegex(name);
      const candidates = [];
      if (role) candidates.push(page.getByRole(role, { name: regex }));
      candidates.push(page.getByText(regex));

      for (const candidate of candidates) {
        const visible = await firstVisible(candidate);
        if (visible) {
          await visible.click({ timeout: Math.max(1, timeout - (Date.now() - start)) });
          return;
        }
      }
    }
    await page.waitForTimeout(250);
  }
  throw new Error(`Не найден кликабельный элемент: ${names.join(' / ')}`);
}

async function waitForNamedAny(page, { role, names, timeout }) {
  const start = Date.now();
  while (Date.now() - start < timeout) {
    for (const name of names) {
      const regex = normalizedTextRegex(name);
      const candidates = [];
      if (role) candidates.push(page.getByRole(role, { name: regex }));
      candidates.push(page.getByText(regex));

      for (const candidate of candidates) {
        const visible = await firstVisible(candidate);
        if (visible) return visible;
      }
    }
    await page.waitForTimeout(250);
  }
  throw new Error(`Не появился элемент: ${names.join(' / ')}`);
}

async function exactCourseExistsInCollection(page, courseName) {
  const titleRegex = normalizedTextRegex(courseName);
  const titleLocator = page.locator('h3').filter({ hasText: titleRegex });

  await page.evaluate(() => window.scrollTo(0, 0));
  let previousHeight = 0;

  for (let i = 0; i < 200; i++) {
    if ((await titleLocator.count()) > 0) return true;

    const state = await page.evaluate(() => ({
      height: document.documentElement.scrollHeight,
      viewport: window.innerHeight,
      y: window.scrollY,
    }));
    const atBottom = state.y + state.viewport >= state.height - 2;
    if (atBottom && state.height === previousHeight) break;

    previousHeight = state.height;
    await page.evaluate(() => window.scrollBy(0, Math.max(400, window.innerHeight * 0.8)));
    await page.waitForTimeout(150);
  }

  await page.evaluate(() => window.scrollTo(0, 0));
  return false;
}

async function readCollectionCourseCount(page, timeout = 5000) {
  const start = Date.now();
  while (Date.now() - start < timeout) {
    const headings = await page.locator('h2').allTextContents();
    for (const heading of headings) {
      const normalized = heading.replace(/\u00a0/g, ' ').trim();
      const match = normalized.match(/^([\d\s]+)\s+(?:курс(?:ов|а)?|courses?)$/iu);
      if (match) return Number(match[1].replace(/\s/g, ''));
    }
    await page.waitForTimeout(200);
  }
  return null;
}

async function firstExactCourseCard(page, courseName) {
  const titleRegex = normalizedTextRegex(courseName);
  const cards = page.locator('[data-testid="product-card-cds"]');
  const count = await cards.count();

  for (let i = 0; i < count; i++) {
    const card = cards.nth(i);
    if (!await card.isVisible().catch(() => false)) continue;

    const title = await firstVisible(card.locator('h3').filter({ hasText: titleRegex }));
    if (!title) continue;

    const courseLinks = card.locator('a[aria-label$=", COURSE"]');
    const linkCount = await courseLinks.count();
    for (let j = 0; j < linkCount; j++) {
      const ariaLabel = await courseLinks.nth(j).getAttribute('aria-label');
      if (ariaLabel?.startsWith(`${courseName}, `) && ariaLabel.endsWith(', COURSE')) {
        return card;
      }
    }
  }

  return null;
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

async function findCourseCardAndSelect(page, courseName, timeout) {
  const start = Date.now();
  let card = null;
  while (Date.now() - start < timeout) {
    card = await firstExactCourseCard(page, courseName);
    if (card) break;
    await page.waitForTimeout(300);
  }
  if (!card) throw new Error(`Карточка COURSE с точным названием не найдена: ${courseName}`);

  const checkbox = card.getByRole('checkbox').first();
  if ((await checkbox.count()) > 0) {
    await checkbox.check({ timeout });
    return;
  }

  for (const name of UI_TEXT.select) {
    const selectText = card.getByText(normalizedTextRegex(name)).first();
    if ((await selectText.count()) > 0) {
      await selectText.click({ timeout });
      return;
    }
  }

  throw new Error(`У курса не найден элемент Select / «Выбрать»: ${courseName}`);
}

async function findExactCourseAcrossSearchPages(page, courseName, timeout) {
  const currentPageButton = page.locator('button[aria-current="page"]');
  const firstCourseResult = page.locator(
    '[data-testid="product-card-cds"] a[aria-label$=", COURSE"]',
  ).first();

  for (let visitedPages = 0; visitedPages < 100; visitedPages++) {
    await currentPageButton.waitFor({ state: 'attached', timeout });
    await firstCourseResult.waitFor({ state: 'visible', timeout });

    const exactCourseCard = await firstExactCourseCard(page, courseName);
    if (exactCourseCard) return exactCourseCard;

    let next = null;
    for (const name of UI_TEXT.nextPage) {
      next = await firstVisible(page.getByRole('button', { name: normalizedTextRegex(name) }));
      if (next) break;
    }
    if (!next || await next.isDisabled()) break;

    const previousPage = (await currentPageButton.textContent())?.trim();
    const previousFirstResult = await firstCourseResult.getAttribute('aria-label');
    await next.click({ timeout });
    await page.waitForFunction(({ oldPage, oldFirstResult }) => {
      const current = document.querySelector('button[aria-current="page"]');
      const firstResult = document.querySelector(
        '[data-testid="product-card-cds"] a[aria-label$=", COURSE"]',
      );
      return current?.textContent?.trim() !== oldPage
        && firstResult?.getAttribute('aria-label') !== oldFirstResult;
    }, { oldPage: previousPage, oldFirstResult: previousFirstResult }, { timeout });
  }

  return null;
}

function formatCollectionCourseCounts(row) {
  const before = row.collection_courses_before;
  const after = row.collection_courses_after;
  const hasBefore = Number.isFinite(before);
  const hasAfter = Number.isFinite(after);

  if (hasBefore && hasAfter && before !== after) return ` | курсов: ${before} → ${after}`;
  if (hasAfter) return ` | курсов в коллекции: ${after}`;
  if (hasBefore) return ` | курсов в коллекции: ${before}`;
  return '';
}

async function processProgram(page, program, courseName, dryRun, programState) {
  const actionTimeout = config.actionTimeoutMs || 20000;
  const collectionsUrl = program.url.replace(/\/main\/?$/, '/catalog/collections');

  await gotoWithRetry(page, collectionsUrl, config.navigationTimeoutMs || 60000);

  // Если сессия истекла, не продолжаем, чтобы не получить 100 одинаковых ошибок.
  if (/login|auth|signin/i.test(page.url())) {
    throw new Error('Coursera запросила повторный вход. Запустите npm run login и войдите снова.');
  }

  await clickNamed(page, { role: 'link', name: config.targetCollection, timeout: actionTimeout }).catch(async () => {
    await clickNamed(page, { role: null, name: config.targetCollection, timeout: actionTimeout });
  });

  await waitForNamedAny(page, { role: 'button', names: UI_TEXT.addContent, timeout: actionTimeout }).catch(async () => {
    await waitForNamedAny(page, { role: 'link', names: UI_TEXT.addContent, timeout: actionTimeout });
  });

  programState.collection_courses_before = await readCollectionCourseCount(page);

  // Идемпотентность: если курс уже есть в коллекции, пропускаем.
  if (await exactCourseExistsInCollection(page, courseName)) {
    programState.collection_courses_after = programState.collection_courses_before;
    return { status: 'already_exists', message: 'Курс уже есть в коллекции' };
  }

  await clickNamedAny(page, { role: 'button', names: UI_TEXT.addContent, timeout: actionTimeout }).catch(async () => {
    await clickNamedAny(page, { role: 'link', names: UI_TEXT.addContent, timeout: actionTimeout });
  });

  const search = page.locator([
    'input[placeholder="Search catalog"]',
    'input[placeholder="Search the catalog"]',
    'input[placeholder="Поиск в каталоге"]',
    'input[role="searchbox"]',
    'input[type="search"]',
  ].join(', ')).first();
  await search.waitFor({ state: 'visible', timeout: actionTimeout });

  // Ограничиваем выдачу курсами, чтобы совпадающий заголовок видео или урока
  // никогда не считался найденным курсом.
  const courseTypeFilter = page.locator('[data-testid="productTypeDescription:Courses-false"] input[type="checkbox"]');
  if (await courseTypeFilter.isVisible().catch(() => false)) {
    await courseTypeFilter.check({ timeout: actionTimeout });
  }

  const courseWords = courseName.trim().split(/\s+/);
  const searchQueries = [...new Set([
    courseName,
    courseWords.slice(-2).join(' '),
    courseWords.at(-1),
  ].filter(Boolean))];

  let exactCourse = null;
  for (const query of searchQueries) {
    await search.click({ timeout: actionTimeout });
    await search.press(process.platform === 'darwin' ? 'Meta+A' : 'Control+A');
    await search.pressSequentially(query, { delay: 20, timeout: actionTimeout });
    const enteredQuery = await search.inputValue();
    if (enteredQuery !== query) {
      throw new Error(`Поле поиска получило неверное значение: ${enteredQuery}`);
    }

    // В контейнере две кнопки: «Очистить поиск» и «Поиск». Выбираем вторую
    // по точному ARIA-имени, иначе запрос сбросится до пустого.
    const searchContainer = search.locator('xpath=ancestor::*[.//button][1]');
    let searchButton = null;
    for (const name of UI_TEXT.search) {
      searchButton = await firstVisible(
        searchContainer.getByRole('button', { name: normalizedTextRegex(name) }),
      );
      if (searchButton) break;
    }
    if (!searchButton) throw new Error('Не найдена кнопка Search / «Поиск»');
    await searchButton.click({ timeout: actionTimeout });
    await page.waitForURL((url) => url.searchParams.get('query') === query, { timeout: actionTimeout });
    await page.locator('h1, h2, h3, h4').filter({ hasText: query }).first()
      .waitFor({ state: 'visible', timeout: actionTimeout });

    exactCourse = await findExactCourseAcrossSearchPages(page, courseName, actionTimeout);
    if (exactCourse) break;
  }

  if (!exactCourse) {
    throw new Error(`Курс с точным названием не найден во всех результатах поиска: ${courseName}`);
  }

  if (dryRun) {
    programState.collection_courses_after = programState.collection_courses_before;
    return { status: 'dry_run_ok', message: 'Курс найден; выбор и подтверждение не выполнялись' };
  }

  await findCourseCardAndSelect(page, courseName, actionTimeout);

  const confirmRegex = /^\s*(?:Confirm\s+selections|Подтвердить\s+выбор\s+элементов)\s*\(\s*1\s*\)\s*$/iu;
  const confirmButton = page.getByRole('button', { name: confirmRegex }).first();
  await confirmButton.waitFor({ state: 'visible', timeout: actionTimeout });
  await confirmButton.click({ timeout: actionTimeout });

  // После подтверждения Coursera должна вернуть на страницу коллекции.
  await waitForNamedAny(page, { role: 'button', names: UI_TEXT.addContent, timeout: actionTimeout }).catch(async () => {
    await waitForNamedAny(page, { role: 'link', names: UI_TEXT.addContent, timeout: actionTimeout });
  });

  programState.collection_courses_after = await readCollectionCourseCount(page);
  const verified = await exactCourseExistsInCollection(page, courseName);
  return {
    status: verified ? 'success_verified' : 'success_submitted',
    message: verified ? 'Курс добавлен и найден в коллекции' : 'Добавление подтверждено; курс не найден в текущем DOM списка для дополнительной проверки',
  };
}

(async () => {
  const artifacts = createRunArtifacts('addition');
  installConsoleTranscript(artifacts.consoleLogPath);

  const opts = parseArgs(process.argv.slice(2));
  const courseName = (opts.course || config.courseName || '').trim();
  if (!courseName) throw new Error('Не указано название курса. Используйте --course="Название" или config.json.');

  let selectedPrograms = programs;
  if (opts.slugs.length > 0) {
    const requestedSlugs = new Set(opts.slugs);
    selectedPrograms = selectedPrograms.filter((p) => requestedSlugs.has(p.slug));
    const missingSlugs = opts.slugs.filter((slug) => !selectedPrograms.some((p) => p.slug === slug));
    if (missingSlugs.length > 0) throw new Error(`Slug не найден в programs.json: ${missingSlugs.join(', ')}`);
  }
  if (Number.isFinite(opts.limit) && opts.limit > 0) selectedPrograms = selectedPrograms.slice(0, opts.limit);

  console.log(`Запуск: ${new Date().toLocaleString('ru-RU')}`);
  console.log(`Команда: node add-course.js ${process.argv.slice(2).join(' ')}`);
  console.log(`Папка запуска: ${artifacts.runDir}`);
  console.log(`Лог консоли: ${artifacts.consoleLogPath}`);
  console.log(`Курс: ${courseName}`);
  console.log(`Коллекция: ${config.targetCollection}`);
  console.log(`Программ к обработке: ${selectedPrograms.length}`);
  console.log(`Режим: ${opts.dryRun ? 'DRY RUN (без добавления)' : 'ДОБАВЛЕНИЕ'}`);

  const context = await chromium.launchPersistentContext(profileDir, {
    channel: 'chrome',
    headless: false,
    viewport: null,
    args: ['--start-maximized'],
  });
  // Отдельная вкладка не участвует в восстановлении предыдущей Chrome-сессии,
  // которое иногда прерывает первый page.goto с net::ERR_ABORTED.
  const page = await context.newPage();
  page.setDefaultTimeout(config.actionTimeoutMs || 20000);
  page.setDefaultNavigationTimeout(config.navigationTimeoutMs || 60000);

  const results = [];
  let fatalAuthError = false;

  for (let i = 0; i < selectedPrograms.length; i++) {
    const program = selectedPrograms[i];
    const programState = {};
    const prefix = `[${i + 1}/${selectedPrograms.length}] ${program.name}`;
    process.stdout.write(`${prefix} ... `);

    try {
      const result = await processProgram(page, program, courseName, opts.dryRun, programState);
      const row = { ...program, ...programState, ...result };
      results.push(row);
      console.log(`${result.status}${formatCollectionCourseCounts(row)}`);
    } catch (error) {
      const message = error?.message || String(error);
      const screenshotPath = path.join(
        artifacts.screenshotsDir,
        `${String(i + 1).padStart(3, '0')}-${program.slug}-${localStamp()}.png`,
      );
      await page.screenshot({ path: screenshotPath, fullPage: true }).catch(() => {});
      const row = { ...program, ...programState, status: 'error', message, screenshot: screenshotPath };
      results.push(row);
      console.log(`ERROR: ${message}${formatCollectionCourseCounts(row)}`);
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
    { key: 'collection_courses_before', label: 'Курсов до', width: 14 },
    { key: 'collection_courses_after', label: 'Курсов после', width: 16 },
    { key: 'status', label: 'Статус', width: 20 },
    { key: 'message', label: 'Сообщение', width: 60 },
    { key: 'screenshot', label: 'Скриншот ошибки', width: 55 },
  ];
  await writeExcelReport(artifacts.reportPath, header, results, 'Добавление');

  const counts = results.reduce((acc, r) => {
    acc[r.status] = (acc[r.status] || 0) + 1;
    return acc;
  }, {});

  console.log('\nИтог:', counts);
  console.log(`Отчет Excel: ${artifacts.reportPath}`);
  console.log(`Лог консоли UTF-8: ${artifacts.consoleLogPath}`);
  if (fatalAuthError) console.log('Остановка: сессия Coursera истекла. Повторите npm run login.');
  if (counts.error) process.exitCode = 1;

  await context.close();
})().catch((error) => {
  console.error('\nFATAL:', error);
  process.exitCode = 1;
});
