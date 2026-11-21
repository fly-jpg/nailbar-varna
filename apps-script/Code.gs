// Принимает заявки с сайта и пишет их строкой в первый лист таблицы SHEET_ID.
// Отдельный проект на script.google.com, затем Развернуть → Новое развертывание → Веб-приложение
// (Запуск от: Я, Доступ: Все). Полученный URL вставить в FORM_ENDPOINT в app.js.

const SHEET_ID = "1h5ILzfA-mPfqHF0FCy6MY4hbUy9Oef6KG6kmt6yWLdk";

const HEADERS = ["Дата и время", "Имя", "Телефон", "Услуга", "Сообщение", "Язык сайта"];
const MAX_LENGTH = { name: 100, phone: 40, service: 80, message: 2000, lang: 5 };

function doPost(event) {
  const params = (event && event.parameter) || {};
  // Honeypot: настоящие люди это поле не видят и не заполняют.
  if (params.website) return reply({ ok: true });

  const name = clean(params.name, MAX_LENGTH.name);
  const phone = clean(params.phone, MAX_LENGTH.phone);
  if (!name || !phone) return reply({ ok: false, error: "missing fields" });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
      sheet.setFrozenRows(1);
    }
    sheet.appendRow([
      new Date(),
      name,
      phone,
      clean(params.service, MAX_LENGTH.service),
      clean(params.message, MAX_LENGTH.message),
      clean(params.lang, MAX_LENGTH.lang)
    ]);
  } finally {
    lock.releaseLock();
  }
  return reply({ ok: true });
}

// Апостроф впереди: телефон "+359..." остается текстом, а "=..." не станет формулой.
function clean(value, maxLength) {
  const text = String(value || "").trim().slice(0, maxLength);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function reply(body) {
  return ContentService.createTextOutput(JSON.stringify(body)).setMimeType(ContentService.MimeType.JSON);
}
