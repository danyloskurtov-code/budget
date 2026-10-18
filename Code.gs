/**
 * Наш бюджет — бекенд на Google Sheets.
 * 1) Вставте цей код у Розширення → Apps Script.
 * 2) Запустіть setup() один раз → у Журналі виконання буде КЛЮЧ доступу.
 * 3) Розгорніть як веб-застосунок (Виконувати від: мене, Доступ: усі).
 */

const LIMITS = [
  ['Житло', '🏠', 3650],
  ['Продукти', '🛒', 1800],
  ['Авто', '⛽', 850],
  ['Репетитор', '📚', 490],
  ['Кіт', '🐈', 420],
  ['Кафе', '☕', 400],
  ['Краса', '💅', 325],
  ['Інше', '❓', 300],
  ['Підписки', '📱', 172],
  ['Інтернет і телефон', '📶', 152],
  ['Побут', '🧽', 150],
  ['Бабуся', '👵', 70],
];

// назва, емодзі, внесок/міс, ціль, стартовий баланс, тип (goal = ціль, fund = копілка), дата цілі
const FUNDS_HEAD = ['fund', 'emoji', 'monthly', 'target', 'start', 'kind', 'deadline'];
const FUNDS = [
  ['Подушка', '🛡️', 1500, 52700, 36025, 'goal', ''],
  ['Відпустка', '🏖️', 500, 6000, 0, 'goal', ''],
  ['Трихоскоп', '🔬', 0, 0, 0, 'goal', ''],
  ['Авто-резерв', '🔧', 350, 4130, 0, 'fund', ''],
  ['Техніка і дім', '🔌', 150, 1800, 0, 'fund', ''],
  ['Одяг', '👕', 150, 1800, 0, 'fund', ''],
  ["Здоров'я", '🩺', 100, 1200, 0, 'fund', ''],
];

const SETTINGS = [
  ['income_plan', 11850],
  ['cycle_day', 10],
];

const TX_HEAD = ['id', 'created', 'date', 'type', 'amount', 'category', 'method', 'who', 'note'];

function setup() {
  const ss = SpreadsheetApp.getActive();
  const tx = ensure_(ss, 'Tx', TX_HEAD);
  tx.getRange('C:C').setNumberFormat('@');
  seed_(ensure_(ss, 'Limits', ['category', 'emoji', 'limit']), LIMITS);
  const funds = ensure_(ss, 'Funds', FUNDS_HEAD);
  migrateFunds_(funds);
  seed_(funds, FUNDS);
  funds.getRange('G:G').setNumberFormat('@');
  seed_(ensure_(ss, 'Settings', ['key', 'value']), SETTINGS);
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('TOKEN')) {
    props.setProperty('TOKEN', Utilities.getUuid().replace(/-/g, '').slice(0, 10));
  }
  Logger.log('КЛЮЧ ДОСТУПУ: ' + props.getProperty('TOKEN'));
}

function ensure_(ss, name, head) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(head);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, head.length).setFontWeight('bold');
  }
  return sh;
}

// старі таблиці без колонок kind/deadline: доповнюємо і додаємо відсутні цілі
function migrateFunds_(sh) {
  if (sh.getLastRow() <= 1) return;
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  if (head.indexOf('kind') < 0) {
    sh.getRange(1, 6, 1, 2).setValues([['kind', 'deadline']]).setFontWeight('bold');
    const names = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0]));
    sh.getRange(2, 6, names.length, 2).setValues(names.map(n => [/^(Подушка|Відпустка)$/.test(n) ? 'goal' : 'fund', '']));
  }
  const names = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0]));
  FUNDS.forEach(f => { if (f[5] === 'goal' && names.indexOf(f[0]) < 0) sh.appendRow(f); });
}

function seed_(sh, rows) {
  if (sh.getLastRow() === 1) sh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function doGet(e) {
  try {
    const p = (e && e.parameter) || {};
    auth_(p.token);
    if (p.action === 'state') return json_(state_(p.from || ''));
    return json_({ ok: false, error: 'Невідома дія' });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const b = JSON.parse(e.postData.contents);
    auth_(b.token);
    if (b.action === 'add') { addTx_(b.tx); return json_({ ok: true }); }
    if (b.action === 'delete') { delTx_(b.id); return json_({ ok: true }); }
    if (b.action === 'saveGoal') { saveGoal_(b.goal); return json_({ ok: true }); }
    return json_({ ok: false, error: 'Невідома дія' });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function auth_(t) {
  const tok = PropertiesService.getScriptProperties().getProperty('TOKEN');
  if (!tok || String(t) !== tok) throw new Error('Невірний ключ доступу');
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function rows_(name) {
  const v = SpreadsheetApp.getActive().getSheetByName(name).getDataRange().getValues();
  const head = v.shift();
  return v
    .filter(r => r.join('') !== '')
    .map(r => {
      const o = {};
      head.forEach((k, i) => { o[k] = r[i]; });
      return o;
    });
}

function ymd_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v);
}

function state_(from) {
  const tx = rows_('Tx').map(r => ({
    id: String(r.id),
    date: ymd_(r.date),
    type: String(r.type),
    amount: Number(r.amount),
    category: String(r.category),
    method: String(r.method || ''),
    who: String(r.who || ''),
    note: String(r.note || ''),
  }));
  const funds = rows_('Funds').map(f => {
    const name = String(f.fund);
    const moved = tx
      .filter(t => t.category === name)
      .reduce((s, t) => s + (t.type === 'fund_in' ? t.amount : t.type === 'fund_out' ? -t.amount : 0), 0);
    return {
      fund: name,
      emoji: String(f.emoji),
      monthly: Number(f.monthly),
      target: Number(f.target),
      start: Number(f.start),
      balance: Number(f.start) + moved,
      kind: String(f.kind || 'fund'),
      deadline: f.deadline ? ymd_(f.deadline) : '',
    };
  });
  const settings = {};
  rows_('Settings').forEach(r => { settings[r.key] = r.value; });
  return {
    ok: true,
    limits: rows_('Limits').map(l => ({ category: String(l.category), emoji: String(l.emoji), limit: Number(l.limit) })),
    funds: funds,
    settings: settings,
    tx: from ? tx.filter(t => t.date >= from) : tx,
  };
}

function addTx_(t) {
  if (!t || !t.id) throw new Error('Порожній запис');
  const amount = Number(t.amount);
  if (!(amount > 0)) throw new Error('Сума має бути більше нуля');
  if (['expense', 'income', 'fund_in', 'fund_out'].indexOf(t.type) < 0) throw new Error('Невідомий тип запису');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(t.date))) throw new Error('Невірна дата');
  const sh = SpreadsheetApp.getActive().getSheetByName('Tx');
  const n = sh.getLastRow() - 1;
  if (n > 0) {
    const ids = sh.getRange(2, 1, n, 1).getValues().map(r => String(r[0]));
    if (ids.indexOf(String(t.id)) >= 0) return; // повтор із офлайн-черги
  }
  sh.appendRow([String(t.id), new Date(), String(t.date), t.type, amount,
    String(t.category || ''), String(t.method || ''), String(t.who || ''), String(t.note || '')]);
}

// створити або оновити ціль/копілку (назва — ключ, її не змінюємо)
function saveGoal_(g) {
  const name = String((g && g.fund) || '').trim();
  if (!name) throw new Error('Вкажіть назву цілі');
  const target = Number(g.target), monthly = Number(g.monthly), start = Number(g.start || 0);
  if (!(target >= 0) || !(monthly >= 0) || !(start >= 0)) throw new Error('Суми не можуть бути відʼємними');
  const deadline = String(g.deadline || '');
  if (deadline && !/^\d{4}-\d{2}-\d{2}$/.test(deadline)) throw new Error('Невірна дата цілі');
  const kind = g.kind === 'fund' ? 'fund' : 'goal';
  const sh = SpreadsheetApp.getActive().getSheetByName('Funds');
  migrateFunds_(sh);
  const n = sh.getLastRow() - 1;
  const names = n > 0 ? sh.getRange(2, 1, n, 1).getValues().map(r => String(r[0])) : [];
  const i = names.indexOf(name);
  if (i >= 0) {
    sh.getRange(i + 2, 2, 1, 6).setValues([[String(g.emoji || '🎯'), monthly, target, sh.getRange(i + 2, 5).getValue(), kind, deadline]]);
  } else {
    sh.appendRow([name, String(g.emoji || '🎯'), monthly, target, start, kind, deadline]);
  }
}

function delTx_(id) {
  const sh = SpreadsheetApp.getActive().getSheetByName('Tx');
  const n = sh.getLastRow() - 1;
  if (n <= 0) return;
  const ids = sh.getRange(2, 1, n, 1).getValues().map(r => String(r[0]));
  const i = ids.indexOf(String(id));
  if (i >= 0) sh.deleteRow(i + 2);
}
