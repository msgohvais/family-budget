// Google Apps Script: принимает операции из приложения и пишет их в лист «Операции».
// Установка: в таблице «Семейный бюджет — база данных» → Расширения → Apps Script → вставить код → Развернуть → Веб-приложение
// Выполнять от имени: Я · Доступ: Все, у кого есть ссылка. Полученную ссылку (…/exec) вставить в приложение: Настройки → Google Таблица.

const SHEET_NAME = 'Операции';
const HEADER = ['id', 'дата', 'тип', 'кошелёк', 'категория', 'подкатегория', 'сумма', 'комментарий', 'кошелёк 2'];

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  const rows = (data.txs || []).map(t => [
    t.id,
    t.d,
    t.t === 'exp' ? 'Расход' : t.t === 'inc' ? 'Доход' : 'Перевод',
    t.w || '',
    t.t === 'exp' ? t.c : t.t === 'inc' ? t.src : '',
    '',
    t.a,
    t.n || '',
    t.w2 || ''
  ]);
  sh.clearContents();
  sh.getRange(1, 1, rows.length + 1, HEADER.length).setValues([HEADER, ...rows]);
  return json({ ok: true, count: rows.length });
}

function doGet() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sh) return json({ ok: true, count: 0 });
  return json({ ok: true, rows: sh.getDataRange().getValues() });
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
