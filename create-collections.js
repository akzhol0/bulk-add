const { chromium } = require('playwright');
const readXlsxFile = require('read-excel-file/node');
const fs = require('fs');
const path = require('path');

const config = require('./config.json');
const {
  createRunArtifacts,
  installConsoleTranscript,
  localStamp,
  writeExcelReport,
} = require('./reporting');

const profileDir = path.join(__dirname, '.playwright-profile');
const courseColumnHeader = 'Онлайн-курсы Coursera';
const defaultCollectionNamesPath = path.join(__dirname, 'collection-names.json');

const UI_TEXT = {
  addCollection: ['Add Collection', 'Добавить коллекцию'],
  continue: ['Continue', 'Продолжить'],
  search: ['Search', 'Поиск'],
  select: ['Select', 'Выбрать'],
  nextPage: ['Go to next page', 'Перейти на следующую страницу'],
};

const reportHeader = [
  { key: 'collection', label: 'Коллекция', width: 38 },
  { key: 'excel_row', label: 'Строка Excel', width: 14 },
  { key: 'course', label: 'Курс', width: 55 },
  { key: 'collection_status', label: 'Статус коллекции', width: 25 },
  { key: 'course_status', label: 'Статус курса', width: 25 },
  { key: 'reorder_status', label: 'Перемещение коллекции', width: 28 },
  { key: 'message', label: 'Сообщение', width: 60 },
  { key: 'program_url', label: 'URL программы', width: 60 },
  { key: 'source_file', label: 'Исходный Excel', width: 55 },
  { key: 'screenshot', label: 'Скриншот ошибки', width: 55 },
];

function normalizeText(value) {
  return String(value ?? '').replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim();
}

function resolveCourseSearchMaxPages(value) {
  return Number.isInteger(value) && value >= 1 ? value : 2;
}

function stripTrailingCourseTypeLabel(value) {
  const courseName = normalizeText(value);
  const stripped = courseName.replace(
    /(?:\s*[-–—:|]\s*|\s+)\(?\s*(?:speciali[sz]ation|course|специализация|курс)\s*\)?\s*$/iu,
    '',
  ).trim();
  return stripped || courseName;
}

function buildCourseSearchQueries(courseName) {
  const words = courseName.split(/\s+/);
  return [...new Set([
    courseName,
    words.slice(-2).join(' '),
    words.at(-1),
  ].filter(Boolean))];
}

function buildCourseSearchAttempts(courseName) {
  const originalName = normalizeText(courseName);
  const strippedName = stripTrailingCourseTypeLabel(originalName);

  if (strippedName === originalName) {
    return [{ exactName: originalName, queries: buildCourseSearchQueries(originalName) }];
  }

  return [
    { exactName: originalName, queries: [originalName] },
    { exactName: strippedName, queries: buildCourseSearchQueries(strippedName) },
  ];
}

function flexibleTextPattern(text) {
  return normalizeText(text)
    .split(/\s+/)
    .map((part) => part.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'))
    .join('\\s+');
}

function normalizedTextRegex(text) {
  return new RegExp(`^\\s*${flexibleTextPattern(text)}\\s*$`, 'u');
}

function parseArgs(argv) {
  const opts = {
    url: null,
    file: null,
    names: null,
    dryRun: false,
    reorderOnly: false,
    limitCollections: null,
  };

  for (const arg of argv) {
    if (arg === '--dry-run') opts.dryRun = true;
    else if (arg === '--reorder-only') opts.reorderOnly = true;
    else if (arg.startsWith('--url=')) opts.url = arg.slice('--url='.length).trim();
    else if (arg.startsWith('--file=')) opts.file = arg.slice('--file='.length).trim();
    else if (arg.startsWith('--names=')) opts.names = arg.slice('--names='.length).trim();
    else if (arg.startsWith('--limit-collections=')) {
      const value = Number(arg.slice('--limit-collections='.length));
      if (!Number.isInteger(value) || value < 1) {
        throw new Error('--limit-collections должен быть целым числом от 1 и больше.');
      }
      opts.limitCollections = value;
    }
  }

  return opts;
}

function loadCollectionNameOverrides(fileOption) {
  const filePath = fileOption
    ? path.resolve(process.cwd(), fileOption)
    : defaultCollectionNamesPath;

  if (!fs.existsSync(filePath)) {
    if (fileOption) throw new Error(`Файл полных названий не найден: ${filePath}`);
    return { filePath: null, names: {} };
  }

  let parsed;
  try {
    parsed = JSON.parse(fs.readFileSync(filePath, 'utf8'));
  } catch (error) {
    throw new Error(`Не удалось прочитать ${filePath}: ${error?.message || String(error)}`);
  }

  if (!parsed || Array.isArray(parsed) || typeof parsed !== 'object') {
    throw new Error(`Файл ${filePath} должен содержать объект: «название вкладки»: «полное название».`);
  }

  const names = {};
  for (const [sheetName, collectionName] of Object.entries(parsed)) {
    const normalizedSheetName = normalizeText(sheetName);
    const normalizedCollectionName = normalizeText(collectionName);
    if (!normalizedSheetName || !normalizedCollectionName) {
      throw new Error(`В ${filePath} обнаружено пустое название вкладки или коллекции.`);
    }
    if (normalizedCollectionName.length > 80) {
      throw new Error(`Название коллекции длиннее 80 символов: ${normalizedCollectionName}`);
    }
    names[normalizedSheetName] = normalizedCollectionName;
  }

  return { filePath, names };
}

function resolveCollectionName(sheetName, nameOverrides = {}) {
  const normalizedSheetName = normalizeText(sheetName);
  const overriddenName = normalizeText(nameOverrides[normalizedSheetName]);
  if (overriddenName) return overriddenName;

  if (normalizedSheetName.length === 31) {
    console.warn(
      `ПРЕДУПРЕЖДЕНИЕ: название вкладки «${normalizedSheetName}» содержит 31 символ `
      + 'и может быть обрезано Excel. Полного варианта нет в collection-names.json; '
      + 'используется название вкладки как есть.',
    );
  }

  return normalizedSheetName;
}

function resolveWorkbookPath(fileOption) {
  if (fileOption) {
    const resolved = path.resolve(process.cwd(), fileOption);
    if (!fs.existsSync(resolved)) throw new Error(`Excel-файл не найден: ${resolved}`);
    if (path.extname(resolved).toLowerCase() !== '.xlsx') {
      throw new Error(`Ожидался файл .xlsx: ${resolved}`);
    }
    return resolved;
  }

  const candidates = fs.readdirSync(__dirname, { withFileTypes: true })
    .filter((entry) => entry.isFile())
    .map((entry) => entry.name)
    .filter((name) => name.toLowerCase().endsWith('.xlsx'))
    .filter((name) => !name.startsWith('~$'))
    .filter((name) => name.toLowerCase() !== 'all courses.xlsx');

  if (candidates.length === 1) return path.join(__dirname, candidates[0]);
  if (candidates.length === 0) {
    throw new Error('В корне проекта не найден Excel-файл для создания коллекций.');
  }
  throw new Error(
    `В корне найдено несколько Excel-файлов. Укажите нужный через --file="имя.xlsx": ${candidates.join(', ')}`,
  );
}

function normalizeCollectionsUrl(value) {
  if (!value) {
    throw new Error('Не указана ссылка. Используйте --url="https://www.coursera.org/.../catalog/collections".');
  }

  let url;
  try {
    url = new URL(value);
  } catch {
    throw new Error(`Некорректная ссылка Coursera: ${value}`);
  }

  if (!/(^|\.)coursera\.org$/i.test(url.hostname)) {
    throw new Error(`Ссылка должна вести на coursera.org: ${value}`);
  }

  const programMatch = url.pathname.match(/^(.*\/admin\/programs\/[^/]+)(?:\/.*)?$/i);
  if (!programMatch) {
    throw new Error('В ссылке не найден путь /admin/programs/<program>.');
  }

  url.pathname = `${programMatch[1]}/catalog/collections`;
  url.search = '';
  url.hash = '';
  return url.toString();
}

function buildManualCollectionUrl(collectionsUrl) {
  const url = new URL(collectionsUrl);
  const match = url.pathname.match(
    /^(\/o\/[^/]+)\/admin\/programs\/([^/]+)\/catalog\/collections\/?$/i,
  );
  if (!match) {
    throw new Error('Не удалось построить адрес ручного создания коллекции из ссылки программы.');
  }

  url.pathname = `${match[1]}/program-creation/${match[2]}/new-collection`;
  url.search = 'continueTo=catalog';
  url.hash = '';
  return url.toString();
}

async function loadCollectionsFromWorkbook(filePath, nameOverrides = {}) {
  const workbook = await readXlsxFile(filePath);
  const sheets = Array.isArray(workbook[0]?.data)
    ? workbook
    : [{ sheet: 'Sheet1', data: workbook }];

  const collections = [];
  for (const sheet of sheets) {
    const sourceSheet = normalizeText(sheet.sheet);
    const collectionName = resolveCollectionName(sourceSheet, nameOverrides);
    const rows = sheet.data || [];
    let headerRow = -1;
    let courseColumn = -1;

    for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
      for (let columnIndex = 0; columnIndex < rows[rowIndex].length; columnIndex++) {
        if (normalizeText(rows[rowIndex][columnIndex]) === courseColumnHeader) {
          headerRow = rowIndex;
          courseColumn = columnIndex;
          break;
        }
      }
      if (headerRow >= 0) break;
    }

    if (headerRow < 0) {
      throw new Error(`Вкладка «${collectionName}»: не найден столбец «${courseColumnHeader}».`);
    }

    const courses = [];
    const seen = new Set();
    for (let rowIndex = headerRow + 1; rowIndex < rows.length; rowIndex++) {
      const courseName = normalizeText(rows[rowIndex][courseColumn]);
      if (!courseName) continue;
      if (seen.has(courseName)) {
        throw new Error(`Вкладка «${collectionName}»: курс повторяется — ${courseName}.`);
      }
      seen.add(courseName);
      courses.push({ name: courseName, excelRow: rowIndex + 1 });
    }

    if (!collectionName) throw new Error('В Excel обнаружена вкладка без названия.');
    if (collectionName.length > 80) {
      throw new Error(`Название коллекции длиннее 80 символов: ${collectionName}`);
    }
    if (courses.length === 0) {
      throw new Error(`Вкладка «${collectionName}»: список курсов пуст.`);
    }

    collections.push({ name: collectionName, courses });
  }

  return collections;
}

async function firstVisible(locator) {
  const count = await locator.count();
  for (let i = 0; i < count; i++) {
    const item = locator.nth(i);
    if (await item.isVisible().catch(() => false)) return item;
  }
  return null;
}

async function findNamedAny(page, roles, names) {
  for (const role of roles) {
    for (const name of names) {
      const visible = await firstVisible(
        page.getByRole(role, { name: normalizedTextRegex(name) }),
      );
      if (visible) return visible;
    }
  }
  for (const name of names) {
    const visible = await firstVisible(page.getByText(normalizedTextRegex(name)));
    if (visible) return visible;
  }
  return null;
}

async function waitForNamedAny(page, roles, names, timeout) {
  const start = Date.now();
  while (Date.now() - start < timeout) {
    const found = await findNamedAny(page, roles, names);
    if (found) return found;
    await page.waitForTimeout(250);
  }
  throw new Error(`Не появился элемент: ${names.join(' / ')}`);
}

async function clickNamedAny(page, roles, names, timeout) {
  const element = await waitForNamedAny(page, roles, names, timeout);
  await element.click({ timeout });
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

function assertAuthenticated(page) {
  if (/login|auth|signin/i.test(page.url())) {
    throw new Error('Coursera запросила повторный вход. Запустите npm run login и войдите снова.');
  }
}

async function openCollectionsPage(page, collectionsUrl, timeout) {
  await gotoWithRetry(page, collectionsUrl, config.navigationTimeoutMs || 60000);
  assertAuthenticated(page);
  const stableButton = page.locator('button[data-pendo="add-collection"]').first();
  if (!await stableButton.isVisible().catch(() => false)) {
    await waitForNamedAny(page, ['button'], UI_TEXT.addCollection, timeout);
  }
}

async function exactCollectionExists(page, collectionName, timeout = 0) {
  const link = page.getByRole('link', { name: normalizedTextRegex(collectionName) });
  const start = Date.now();
  do {
    if (await firstVisible(link)) return true;
    if (Date.now() - start >= timeout) return false;
    await page.waitForTimeout(250);
  } while (true);
}

function collectionRows(page) {
  return page.locator('li').filter({
    has: page.locator('[data-testid="collection-drag-handle-icon"]'),
  });
}

async function collectionRowByName(page, collectionName) {
  const exactLink = page.getByRole('link', { name: normalizedTextRegex(collectionName) });
  const rows = collectionRows(page).filter({ has: exactLink });
  return firstVisible(rows);
}

async function firstCollectionName(page) {
  const firstRow = collectionRows(page).first();
  if (!await firstRow.isVisible().catch(() => false)) return '';
  const link = firstRow.locator('a[href*="/collection/"]').first();
  return normalizeText(await link.innerText().catch(() => ''));
}

async function collectionNamesInOrder(page) {
  const rows = collectionRows(page);
  const count = await rows.count();
  const names = [];
  for (let index = 0; index < count; index++) {
    const link = rows.nth(index).locator('a[href*="/collection/"]').first();
    names.push(normalizeText(await link.innerText().catch(() => '')));
  }
  return names;
}

async function dragRowAbove(page, sourceRow, targetRow) {
  const sourceHandle = sourceRow.locator('[data-testid="collection-drag-handle-icon"]').first();
  const targetHandle = targetRow.locator('[data-testid="collection-drag-handle-icon"]').first();

  await targetRow.scrollIntoViewIfNeeded();
  await sourceRow.scrollIntoViewIfNeeded();
  const sourceBox = await sourceHandle.boundingBox();
  const targetBox = await targetHandle.boundingBox();
  if (!sourceBox || !targetBox) throw new Error('Не удалось определить координаты drag-and-drop.');

  const sourceX = sourceBox.x + sourceBox.width / 2;
  const sourceY = sourceBox.y + sourceBox.height / 2;
  const targetX = targetBox.x + targetBox.width / 2;
  const targetY = targetBox.y + 2;

  await page.mouse.move(sourceX, sourceY);
  await page.mouse.down();
  try {
    // A short initial movement activates Coursera's pointer-based drag sensor.
    await page.mouse.move(sourceX, sourceY - 10, { steps: 5 });
    await page.waitForTimeout(150);
    await page.mouse.move(targetX, targetY, { steps: 20 });
    await page.waitForTimeout(500);
  } finally {
    await page.mouse.up();
  }
  await page.waitForTimeout(1200);
}

async function dragCollectionToTop(page, collectionName, timeout) {
  const initialRow = await collectionRowByName(page, collectionName);
  if (!initialRow) throw new Error(`Коллекция не найдена для перемещения: ${collectionName}`);
  if (await firstCollectionName(page) === collectionName) return 'already_at_top';

  const maximumSteps = Math.max(1, await collectionRows(page).count()) * 2;
  for (let step = 0; step < maximumSteps; step++) {
    const before = await collectionNamesInOrder(page);
    const sourceIndex = before.indexOf(collectionName);
    if (sourceIndex === 0) {
      await page.waitForTimeout(1000);
      return 'moved_to_top';
    }
    if (sourceIndex < 0) throw new Error(`Коллекция исчезла из списка: ${collectionName}`);

    const rows = collectionRows(page);
    const sourceRow = rows.nth(sourceIndex);
    const targetRow = rows.nth(sourceIndex - 1);
    await dragRowAbove(page, sourceRow, targetRow);

    const after = await collectionNamesInOrder(page);
    const newIndex = after.indexOf(collectionName);
    if (newIndex < 0 || newIndex >= sourceIndex) {
      throw new Error(`Порядок не изменился после перетаскивания: ${collectionName}`);
    }
  }

  throw new Error(`Не удалось поднять коллекцию наверх за ${timeout} мс: ${collectionName}`);
}

async function reorderCollectionsToTop(page, collectionsUrl, collections) {
  const timeout = config.actionTimeoutMs || 20000;
  await openCollectionsPage(page, collectionsUrl, timeout);
  await page.waitForTimeout(1500);

  const statuses = new Map();
  // Moving bottom-to-top in reverse gives the final top block the same order
  // as the Excel tabs: first sheet first, second sheet second, and so on.
  for (const collection of [...collections].reverse()) {
    process.stdout.write(`    ${collection.name} ... `);
    try {
      const status = await dragCollectionToTop(page, collection.name, timeout);
      statuses.set(collection.name, status);
      console.log(status);
    } catch (error) {
      const message = error?.message || String(error);
      statuses.set(collection.name, `error: ${message}`);
      console.log(`ERROR: ${message}`);
    }
  }

  return statuses;
}

async function waitForCollectionConfirmation(page, collectionName, timeout) {
  const start = Date.now();
  let collectionsPageReadySince = null;

  while (Date.now() - start < timeout) {
    assertAuthenticated(page);
    const pathname = new URL(page.url()).pathname;

    if (/\/collection\/[^/]+\/?$/i.test(pathname)) {
      const addContent = await findNamedAny(
        page,
        ['button', 'link'],
        ['Add content', 'Добавить материалы'],
      );
      if (addContent) return 'verified_on_collection_page';
    }

    if (/\/catalog\/collections\/?$/i.test(pathname)) {
      if (await exactCollectionExists(page, collectionName)) {
        return 'verified_in_collections_list';
      }

      const addCollection = await firstVisible(
        page.locator('button[data-pendo="add-collection"]'),
      );
      if (addCollection) {
        if (collectionsPageReadySince === null) collectionsPageReadySince = Date.now();
        if (Date.now() - collectionsPageReadySince >= 3000) {
          return 'submitted_to_collections_list';
        }
      }
    }

    await page.waitForTimeout(250);
  }

  throw new Error('После Confirm selections Coursera не открыла коллекцию или список коллекций.');
}

async function startManualCollection(page, collectionsUrl, collectionName, timeout) {
  // Coursera's Add Collection popover can ignore an automated click even though
  // Playwright reports success. The destination below is the exact href exposed
  // by Coursera's "Create Collection manually" link.
  const manualCollectionUrl = buildManualCollectionUrl(collectionsUrl);
  await gotoWithRetry(page, manualCollectionUrl, config.navigationTimeoutMs || 60000);
  assertAuthenticated(page);
  await waitForNamedAny(page, ['heading'], ['Create a collection', 'Создать коллекцию'], timeout);

  // The form is server-rendered first and then hydrated by React. Filling it too
  // early can look successful for a moment and then be reset to an empty value.
  await page.waitForLoadState('networkidle', { timeout: 5000 }).catch(() => {});
  await page.waitForTimeout(750);

  let nameInput = await firstVisible(page.locator('#collection-name'));
  if (!nameInput) {
    nameInput = await firstVisible(
      page.getByRole('textbox', { name: normalizedTextRegex('Collection name') }),
    );
  }
  if (!nameInput) throw new Error('Не найдено поле Collection name.');

  for (let attempt = 1; attempt <= 2; attempt++) {
    await nameInput.click({ timeout });
    await nameInput.fill(collectionName, { timeout });
    await nameInput.blur();
    await page.waitForTimeout(750);
    if (normalizeText(await nameInput.inputValue()) === collectionName) break;
  }

  const enteredName = normalizeText(await nameInput.inputValue());
  if (enteredName !== collectionName) {
    throw new Error(`Поле Collection name получило неверное значение: ${await nameInput.inputValue()}`);
  }

  // Описание, High stakes и остальные настройки намеренно не трогаем.
  let continueButton = await firstVisible(page.locator('button[data-e2e="continue"]'));
  if (!continueButton) {
    continueButton = await waitForNamedAny(page, ['button'], UI_TEXT.continue, timeout);
  }
  await continueButton.click({ timeout });

  try {
    await findCatalogSearch(page, config.navigationTimeoutMs || 60000);
  } catch (error) {
    const currentValue = normalizeText(await nameInput.inputValue().catch(() => ''));
    if (await nameInput.isVisible().catch(() => false)) {
      throw new Error(
        `После Continue форма не перешла к поиску курсов. Значение Collection name: «${currentValue}».`,
      );
    }
    throw error;
  }
}

async function findCatalogSearch(page, timeout) {
  const candidates = page.locator([
    'input[placeholder="Search catalog"]',
    'input[placeholder="Search the catalog"]',
    'input[placeholder="Поиск в каталоге"]',
    'input[role="searchbox"]',
    'input[type="search"]',
  ].join(', '));
  const start = Date.now();
  while (Date.now() - start < timeout) {
    const search = await firstVisible(candidates);
    if (search) return search;
    await page.waitForTimeout(250);
  }
  throw new Error('Не найдено поле Search catalog / «Поиск в каталоге».');
}

async function firstExactCourseCard(page, courseName) {
  const cards = page.locator('[data-testid="product-card-cds"]');
  const count = await cards.count();

  for (let i = 0; i < count; i++) {
    const card = cards.nth(i);
    if (!await card.isVisible().catch(() => false)) continue;

    const headings = card.locator('h2, h3, [data-testid="product-card-title"]');
    const headingCount = await headings.count();
    for (let j = 0; j < headingCount; j++) {
      const title = normalizeText(await headings.nth(j).innerText().catch(() => ''));
      if (title === courseName) return card;
    }
  }

  return null;
}

async function findExactCourseAcrossPages(page, courseName, timeout, maxPages) {
  const firstCourseResult = page.locator('[data-testid="product-card-cds"]').first();

  for (let visitedPages = 0; visitedPages < maxPages; visitedPages++) {
    await firstCourseResult.waitFor({ state: 'visible', timeout: Math.min(timeout, 3000) }).catch(() => {});

    const exactCard = await firstExactCourseCard(page, courseName);
    if (exactCard) return exactCard;

    if (visitedPages + 1 >= maxPages) break;

    let next = null;
    for (const name of UI_TEXT.nextPage) {
      next = await firstVisible(page.getByRole('button', { name: normalizedTextRegex(name) }));
      if (next) break;
    }
    if (!next || await next.isDisabled()) break;

    const previousFirstResult = normalizeText(
      await firstCourseResult.innerText().catch(() => ''),
    );
    await next.click({ timeout });
    const pageChanged = await page.waitForFunction((oldFirstResult) => {
      const firstResult = document.querySelector('[data-testid="product-card-cds"]');
      const current = (firstResult?.innerText || '').replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim();
      return current !== oldFirstResult;
    }, previousFirstResult, { timeout: Math.min(timeout, 5000) })
      .then(() => true)
      .catch(() => false);
    if (!pageChanged) break;
  }

  return null;
}

async function submitSearch(page, search, query, timeout) {
  await search.click({ timeout });
  await search.press(process.platform === 'darwin' ? 'Meta+A' : 'Control+A');
  await search.pressSequentially(query, { delay: 15, timeout });
  if (await search.inputValue() !== query) {
    throw new Error(`Поле поиска получило неверное значение: ${await search.inputValue()}`);
  }

  const searchContainer = search.locator('xpath=ancestor::*[.//button][1]');
  let searchButton = null;
  for (const name of UI_TEXT.search) {
    searchButton = await firstVisible(
      searchContainer.getByRole('button', { name: normalizedTextRegex(name) }),
    );
    if (searchButton) break;
  }
  if (!searchButton) throw new Error('Не найдена кнопка Search / «Поиск».');

  await searchButton.click({ timeout });
  await page.waitForURL((url) => url.searchParams.get('query') === query, { timeout });
  await page.waitForTimeout(1200);
}

async function findCourse(page, search, courseName, timeout) {
  const maxPages = resolveCourseSearchMaxPages(config.courseSearchMaxPages);
  const attempts = buildCourseSearchAttempts(courseName);

  for (const attempt of attempts) {
    for (const query of attempt.queries) {
      await submitSearch(page, search, query, timeout);
      const card = await findExactCourseAcrossPages(
        page,
        attempt.exactName,
        timeout,
        maxPages,
      );
      if (card) return card;
    }
  }
  return null;
}

async function waitForConfirmCount(page, expectedCount, timeout) {
  const regex = new RegExp(
    `^\\s*(?:Confirm\\s+selections|Подтвердить\\s+выбор\\s+элементов)\\s*\\(\\s*${expectedCount}\\s*\\)\\s*$`,
    'iu',
  );
  const button = page.getByRole('button', { name: regex }).first();
  await button.waitFor({ state: 'visible', timeout });
  return button;
}

async function selectCourse(page, card, selectedCount, timeout) {
  const checkbox = card.getByRole('checkbox').first();
  if ((await checkbox.count()) > 0) {
    if (!await checkbox.isChecked()) await checkbox.check({ timeout });
  } else {
    let selectControl = null;
    for (const name of UI_TEXT.select) {
      selectControl = await firstVisible(card.getByText(normalizedTextRegex(name)));
      if (selectControl) break;
    }
    if (!selectControl) throw new Error('В карточке курса не найден элемент Select.');
    await selectControl.click({ timeout });
  }

  await waitForConfirmCount(page, selectedCount + 1, timeout);
}

async function processCollection(page, collectionsUrl, collection, filePath, dryRun) {
  const timeout = config.actionTimeoutMs || 20000;
  await openCollectionsPage(page, collectionsUrl, timeout);

  if (await exactCollectionExists(page, collection.name, Math.min(timeout, 3000))) {
    console.log(`SKIP: коллекция уже существует (${collection.courses.length} курсов в Excel)`);
    return collection.courses.map((course) => ({
      collection: collection.name,
      excel_row: course.excelRow,
      course: course.name,
      collection_status: 'collection_already_exists',
      course_status: 'skipped',
      message: 'Коллекция с точным названием уже существует; изменений не было.',
      program_url: collectionsUrl,
      source_file: filePath,
    }));
  }

  if (dryRun) {
    console.log(`DRY RUN: будет создана коллекция, курсов: ${collection.courses.length}`);
    return collection.courses.map((course) => ({
      collection: collection.name,
      excel_row: course.excelRow,
      course: course.name,
      collection_status: 'dry_run_would_create',
      course_status: 'not_checked',
      message: 'Dry-run: коллекция и курсы не создавались.',
      program_url: collectionsUrl,
      source_file: filePath,
    }));
  }

  await startManualCollection(page, collectionsUrl, collection.name, timeout);
  const search = await findCatalogSearch(page, timeout);
  const rows = [];
  let selectedCount = 0;

  for (let i = 0; i < collection.courses.length; i++) {
    const course = collection.courses[i];
    process.stdout.write(`    [${i + 1}/${collection.courses.length}] ${course.name} ... `);

    try {
      const card = await findCourse(page, search, course.name, timeout);
      if (!card) {
        rows.push({
          collection: collection.name,
          excel_row: course.excelRow,
          course: course.name,
          collection_status: 'creating',
          course_status: 'not_found',
          message: 'Курс с точным названием не найден; добавьте его вручную.',
          program_url: collectionsUrl,
          source_file: filePath,
        });
        console.log('NOT FOUND');
        continue;
      }

      await selectCourse(page, card, selectedCount, timeout);
      selectedCount += 1;
      rows.push({
        collection: collection.name,
        excel_row: course.excelRow,
        course: course.name,
        collection_status: 'creating',
        course_status: 'selected',
        message: '',
        program_url: collectionsUrl,
        source_file: filePath,
      });
      console.log(`selected (${selectedCount})`);
    } catch (error) {
      rows.push({
        collection: collection.name,
        excel_row: course.excelRow,
        course: course.name,
        collection_status: 'creating',
        course_status: 'error',
        message: error?.message || String(error),
        program_url: collectionsUrl,
        source_file: filePath,
      });
      console.log(`ERROR: ${error?.message || String(error)}`);
    }
  }

  if (selectedCount === 0) {
    for (const row of rows) row.collection_status = 'not_created';
    return rows;
  }

  const confirmButton = await waitForConfirmCount(page, selectedCount, timeout);
  await confirmButton.click({ timeout });
  const confirmation = await waitForCollectionConfirmation(
    page,
    collection.name,
    config.navigationTimeoutMs || 60000,
  );

  const notFoundCount = rows.filter((row) => row.course_status === 'not_found').length;
  const errorCount = rows.filter((row) => row.course_status === 'error').length;
  const collectionStatus = notFoundCount || errorCount ? 'created_partial' : 'created';

  for (const row of rows) {
    row.collection_status = collectionStatus;
    if (row.course_status === 'selected') {
      row.course_status = confirmation === 'submitted_to_collections_list'
        ? 'added_submitted'
        : 'added';
    }
  }

  console.log(
    `    Подтверждено: ${selectedCount}/${collection.courses.length}`
    + ` | не найдено: ${notFoundCount} | ошибок: ${errorCount}`
    + ` | проверка: ${confirmation}`,
  );
  return rows;
}

async function main() {
  const opts = parseArgs(process.argv.slice(2));
  if (opts.dryRun && opts.reorderOnly) {
    throw new Error('Нельзя одновременно использовать --dry-run и --reorder-only.');
  }
  const collectionsUrl = normalizeCollectionsUrl(opts.url);
  const filePath = resolveWorkbookPath(opts.file);
  const nameOverrides = loadCollectionNameOverrides(opts.names);
  let collections = await loadCollectionsFromWorkbook(filePath, nameOverrides.names);

  if (Number.isFinite(opts.limitCollections) && opts.limitCollections > 0) {
    collections = collections.slice(0, opts.limitCollections);
  }

  const artifacts = createRunArtifacts('collections');
  installConsoleTranscript(artifacts.consoleLogPath);

  const totalCourses = collections.reduce((sum, collection) => sum + collection.courses.length, 0);
  console.log(`Запуск: ${new Date().toLocaleString('ru-RU')}`);
  console.log(`Excel: ${filePath}`);
  console.log(`Полные названия: ${nameOverrides.filePath || 'не требуются'}`);
  console.log(`Страница программы: ${collectionsUrl}`);
  console.log(`Коллекций: ${collections.length}`);
  console.log(`Курсов в Excel: ${totalCourses}`);
  console.log(
    `Режим: ${opts.dryRun
      ? 'DRY RUN (без создания)'
      : opts.reorderOnly
        ? 'ТОЛЬКО ПЕРЕМЕЩЕНИЕ КОЛЛЕКЦИЙ'
        : 'СОЗДАНИЕ КОЛЛЕКЦИЙ'}`,
  );
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
  let reorderHadErrors = false;

  for (let index = 0; !opts.reorderOnly && index < collections.length; index++) {
    const collection = collections[index];
    console.log(`\n[${index + 1}/${collections.length}] Коллекция: ${collection.name}`);

    try {
      const rows = await processCollection(
        page,
        collectionsUrl,
        collection,
        filePath,
        opts.dryRun,
      );
      results.push(...rows);
    } catch (error) {
      const message = error?.message || String(error);
      const screenshotPath = path.join(
        artifacts.screenshotsDir,
        `${String(index + 1).padStart(2, '0')}-${localStamp()}.png`,
      );
      await page.screenshot({ path: screenshotPath, fullPage: true }).catch(() => {});
      results.push({
        collection: collection.name,
        excel_row: '',
        course: '',
        collection_status: 'error',
        course_status: '',
        message,
        program_url: collectionsUrl,
        source_file: filePath,
        screenshot: screenshotPath,
      });
      console.log(`ERROR COLLECTION: ${message}`);

      if (/повторный вход|login|auth|signin/i.test(message)) {
        fatalAuthError = true;
      }
    }

    await writeExcelReport(artifacts.reportPath, reportHeader, results, 'Создание коллекций');
    console.log(`Промежуточный отчёт сохранён: ${artifacts.reportPath}`);
    if (fatalAuthError) break;
  }

  if (!opts.dryRun && !fatalAuthError) {
    console.log('\nПеремещение коллекций наверх:');
    const reorderStatuses = await reorderCollectionsToTop(page, collectionsUrl, collections);
    if (opts.reorderOnly) {
      for (const collection of collections) {
        results.push({
          collection: collection.name,
          excel_row: '',
          course: '',
          collection_status: 'reorder_only',
          course_status: '',
          program_url: collectionsUrl,
          source_file: filePath,
        });
      }
    }
    for (const row of results) {
      row.reorder_status = reorderStatuses.get(row.collection) || 'not_attempted';
    }
    reorderHadErrors = [...reorderStatuses.values()].some((status) => status.startsWith('error:'));
    await writeExcelReport(artifacts.reportPath, reportHeader, results, 'Создание коллекций');
    console.log(`Отчёт после сортировки сохранён: ${artifacts.reportPath}`);
  }

  const counts = results.reduce((summary, row) => {
    const key = row.course_status || row.collection_status;
    summary[key] = (summary[key] || 0) + 1;
    return summary;
  }, {});

  console.log('\nИтог по строкам Excel:', counts);
  console.log(`Отчёт Excel: ${artifacts.reportPath}`);
  console.log(`Лог консоли UTF-8: ${artifacts.consoleLogPath}`);
  if (fatalAuthError) console.log('Остановка: сессия Coursera истекла. Повторите npm run login.');

  if (
    counts.error
    || counts.not_found
    || reorderHadErrors
    || results.some((row) => row.collection_status === 'error')
  ) {
    process.exitCode = 1;
  }

  await context.close();
}

if (require.main === module) {
  main().catch((error) => {
    console.error('\nFATAL:', error);
    process.exitCode = 1;
  });
}

module.exports = {
  buildManualCollectionUrl,
  buildCourseSearchAttempts,
  loadCollectionsFromWorkbook,
  loadCollectionNameOverrides,
  normalizeCollectionsUrl,
  normalizeText,
  parseArgs,
  resolveCourseSearchMaxPages,
  resolveCollectionName,
  resolveWorkbookPath,
  stripTrailingCourseTypeLabel,
};
