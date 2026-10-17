// Дашборд продаж: Google Apps Script веб-приложение.
// Читает таблицу при каждом открытии страницы, поэтому данные всегда свежие.
// Две ссылки с разными ключами: команда (с прогнозами и экономикой) и артист (только продажи).

const SHEET_ID = '1qOM_QjNHqrgn9OJbsI8H05BixEhOxqZnhfv0G93HeG0';
const TEAM_KEY = 'ЗАМЕНИТЕ_КЛЮЧ_КОМАНДЫ';
const ARTIST_KEY = 'ЗАМЕНИТЕ_КЛЮЧ_АРТИСТА';

// Настройки тура (залы, вместимость, цены, затраты). Те же, что были в дашборде.
const CFG = {
  "artist": "Семён Слепаков",
  "tour": "европейский тур",
  "company": "Bard Productions",
  "title": "Слепаков · Продажи Европа",
  "year": 2026,
  "channels": [
    "Ticketshop",
    "Kontramarka",
    "Friday"
  ],
  "price_source": "Ticketshop (отчёты продаж из админки)",
  "price_checked": "14.10.2026",
  "avg_note": "средняя цена билета — взвешенная по зонам выпущенных мест из отчётов Ticketshop (а не середина диапазона); заблокированные места учтены по той же средней",
  "fx": {
    "£": 1.17
  },
  "zone_labels": {},
  "cities": {
    "BARCELONA": {
      "name": "Барселона",
      "venue": "Parallel 62",
      "cost": 71000,
      "pmin": 45,
      "pmax": 200,
      "avg": 92.3,
      "cur": "€",
      "on_sale": 925,
      "open": 925,
      "capacity": 933
    },
    "LONDON": {
      "name": "Лондон",
      "venue": "Clapham Grand",
      "cost": 65000,
      "pmin": 75,
      "pmax": 180,
      "avg": 112.4,
      "cur": "£",
      "on_sale": 720,
      "open": 467,
      "capacity": 873
    },
    "DUBLIN": {
      "name": "Дублин (Дрохеда)",
      "venue": "TLT",
      "cost": 84000,
      "pmin": 70,
      "pmax": 200,
      "avg": 97.7,
      "cur": "€",
      "on_sale": 830,
      "open": 830,
      "capacity": 896
    },
    "VILNUS": {
      "name": "Вильнюс",
      "venue": "Pramogų arena",
      "cost": 84000,
      "pmin": 50,
      "pmax": 120,
      "avg": 80.4,
      "cur": "€",
      "on_sale": 1393,
      "open": 1393,
      "capacity": 1492
    },
    "RIGA": {
      "name": "Рига",
      "venue": "Atta Centre",
      "cost": 60000,
      "pmin": 50,
      "pmax": 190,
      "avg": 78.7,
      "cur": "€",
      "on_sale": 1488,
      "open": 1488,
      "capacity": 2099
    },
    "ALIKANTE": {
      "name": "Аликанте",
      "venue": "VB Spaces",
      "cost": 81000,
      "pmin": 60,
      "pmax": 200,
      "avg": 97.7,
      "cur": "€",
      "on_sale": 1000,
      "open": 607,
      "capacity": 1173
    }
  },
  "capacity_note": "рабочая вместимость задана Павлом: Лондон 720, Аликанте 1000, Дублин 830, Вильнюс 1393, Рига 1488, Барселона 925 (открыто в Ticketshop)"
};

function modeFor_(key) {
  if (key && key === TEAM_KEY) return 'team';
  if (key && key === ARTIST_KEY) return 'artist';
  return null;
}

// Для артиста убираем из настроек всё, что касается денег, чтобы этого не было и в исходном коде страницы.
function cfgFor_(mode) {
  if (mode === 'team') return CFG;
  const c = JSON.parse(JSON.stringify(CFG));
  ['fee', 'fx', 'price_source', 'price_checked', 'avg_note'].forEach(k => delete c[k]);
  Object.values(c.cities).forEach(city => ['cost', 'pmin', 'pmax', 'avg', 'cur'].forEach(k => delete city[k]));
  return c;
}

// Таблица -> двумерный массив, даты превращаем в строки yyyy-MM-dd.
function readGrid_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const tz = ss.getSpreadsheetTimeZone();
  const values = ss.getSheets()[0].getDataRange().getValues();
  return values.map(row => row.map(v => v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : (v === '' ? null : v)));
}

function stamp_() {
  return Utilities.formatDate(new Date(), 'Europe/Moscow', 'dd.MM.yyyy HH:mm') + ' (МСК)';
}

function doGet(e) {
  const key = (e && e.parameter && e.parameter.key) || '';
  const mode = modeFor_(key);
  if (!mode) return HtmlService.createHtmlOutput('<p style="font:16px sans-serif;padding:24px">Нет доступа. Проверьте ссылку.</p>');
  const t = HtmlService.createTemplateFromFile('Index');
  const js = o => JSON.stringify(o).replace(/</g, '\\u003c');
  t.gridJson = js(readGrid_());
  t.keyJson = js(key);
  t.cfgJson = js(cfgFor_(mode));
  t.modeJson = js(mode);
  t.stampJson = js(stamp_());
  return t.evaluate()
    .setTitle(CFG.title || 'Продажи')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Вызывается кнопкой «Обновить» и автообновлением страницы.
function getData(key) {
  if (!modeFor_(key)) throw new Error('Нет доступа');
  return { grid: readGrid_(), stamp: stamp_() };
}
