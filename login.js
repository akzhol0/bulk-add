const { chromium } = require('playwright');
const path = require('path');
const readline = require('readline');

const profileDir = path.join(__dirname, '.playwright-profile');
const ADMIN_URL = 'https://www.coursera.org/o/qazaqstan/admin/programs';

function waitForEnter(message) {
  return new Promise((resolve) => {
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
    rl.question(message, () => {
      rl.close();
      resolve();
    });
  });
}

(async () => {
  const context = await chromium.launchPersistentContext(profileDir, {
    channel: 'chrome',
    headless: false,
    viewport: null,
    args: ['--start-maximized'],
  });

  const page = context.pages()[0] || await context.newPage();
  await page.goto(ADMIN_URL, { waitUntil: 'domcontentloaded' });

  console.log('\nВ открывшемся Chrome войдите в Coursera вручную (SSO/MFA тоже вручную).');
  console.log('Когда увидите админ-панель Coursera, вернитесь в терминал.');
  await waitForEnter('Нажмите Enter, чтобы сохранить сессию и закрыть браузер... ');

  await context.close();
  console.log('Сессия сохранена в локальном профиле .playwright-profile.');
})().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
