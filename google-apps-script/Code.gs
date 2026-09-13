const HEADERS = [
  'application_id',
  'submitted_at',
  'route',
  'name',
  'telegram',
  'email',
  'role',
  'recommender',
  'portfolio',
  'about',
  'source',
  'page_url',
  'user_agent',
  'status',
  'admin_comment',
  'notification_status'
];

function doGet() {
  return HtmlService.createHtmlOutput('TUTVSE application endpoint is running');
}

function doPost(e) {
  const requestId = clean_(e && e.parameter && e.parameter.request_id, 100) || Utilities.getUuid();

  try {
    const p = e.parameter || {};

    // Honeypot: отвечаем успехом, но спам не сохраняем.
    if (clean_(p.website, 200)) return response_({ ok: true, request_id: requestId });

    const application = {
      request_id: requestId,
      route: clean_(p.route, 20),
      name: clean_(p.name, 160),
      telegram: normalizeTelegram_(p.telegram),
      email: clean_(p.email, 240).toLowerCase(),
      role: clean_(p.role, 300),
      recommender: clean_(p.recommender, 300),
      portfolio: clean_(p.portfolio, 500),
      about: clean_(p.about, 3000),
      source: clean_(p.source, 500),
      page_url: clean_(p.page_url, 1000),
      user_agent: clean_(p.user_agent, 1000)
    };

    validate_(application);

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);

    let sheet;
    let rowNumber;
    try {
      sheet = getSheet_();
      const duplicate = sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), 1)
        .createTextFinder(requestId)
        .matchEntireCell(true)
        .findNext();

      if (duplicate) {
        return response_({ ok: true, request_id: requestId, application_id: requestId, duplicate: true });
      }

      const row = [
        requestId,
        new Date(),
        application.route,
        application.name,
        application.telegram,
        application.email,
        application.role,
        application.recommender,
        application.portfolio,
        application.about,
        application.source,
        application.page_url,
        application.user_agent,
        'new',
        '',
        'pending'
      ];
      sheet.appendRow(row);
      rowNumber = sheet.getLastRow();
    } finally {
      lock.releaseLock();
    }

    let notificationStatus = 'sent';
    try {
      sendTelegram_(application, requestId);
    } catch (telegramError) {
      notificationStatus = 'error: ' + clean_(telegramError.message, 400);
      console.error(telegramError);
    }
    sheet.getRange(rowNumber, HEADERS.indexOf('notification_status') + 1).setValue(notificationStatus);

    return response_({ ok: true, request_id: requestId, application_id: requestId });
  } catch (error) {
    console.error(error);
    return response_({
      ok: false,
      request_id: requestId,
      message: 'Не удалось сохранить заявку. Попробуйте ещё раз.'
    });
  }
}

function setup() {
  const sheet = getSheet_();
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  sheet.autoResizeColumns(1, HEADERS.length);
  return 'Готово: ' + sheet.getParent().getUrl() + '#gid=' + sheet.getSheetId();
}

function getSheet_() {
  const props = PropertiesService.getScriptProperties();
  const spreadsheetId = props.getProperty('SPREADSHEET_ID');
  const sheetName = props.getProperty('SHEET_NAME') || 'Заявки с сайта';
  if (!spreadsheetId) throw new Error('Не задано свойство SPREADSHEET_ID');

  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) sheet = spreadsheet.insertSheet(sheetName);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  } else {
    const actual = sheet.getRange(1, 1, 1, HEADERS.length).getDisplayValues()[0];
    if (actual.join('|') !== HEADERS.join('|')) {
      throw new Error('Заголовки листа не совпадают. Используйте новый пустой лист или исправьте первую строку.');
    }
  }
  return sheet;
}

function validate_(a) {
  if (!['ref', 'new', 'reco', 'open'].includes(a.route)) throw new Error('Неизвестный маршрут заявки');
  if (a.name.length < 2) throw new Error('Не заполнено имя');
  if (!/^@?[a-zA-Z0-9_]{4,}$/.test(a.telegram)) throw new Error('Некорректный Telegram');
  if (!/^[^@\s]+@[^@\s.]+\.[^@\s]{2,}$/.test(a.email)) throw new Error('Некорректный email');
  if (a.role.length < 2) throw new Error('Не заполнена роль');
  if (['ref', 'reco'].includes(a.route) && a.recommender.length < 2) throw new Error('Не заполнен рекомендатель');
  if (['new', 'open'].includes(a.route) && a.portfolio.length < 4) throw new Error('Не заполнено портфолио');
  if (['new', 'open'].includes(a.route) && a.about.length < 16) throw new Error('Не заполнено описание');
}

function sendTelegram_(a, applicationId) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('BOT_TOKEN');
  const chatId = props.getProperty('CHAT_ID');
  if (!token || !chatId) throw new Error('Не заданы BOT_TOKEN или CHAT_ID');

  const route = ['ref', 'reco'].includes(a.route) ? 'По рекомендации' : 'Новое знакомство';
  const lines = [
    '<b>Новая заявка в «Тут все»</b>',
    '',
    '<b>ID:</b> <code>' + escapeHtml_(applicationId) + '</code>',
    '<b>Маршрут:</b> ' + escapeHtml_(route),
    '<b>Имя:</b> ' + escapeHtml_(a.name),
    '<b>Telegram:</b> ' + escapeHtml_(a.telegram),
    '<b>Email:</b> ' + escapeHtml_(a.email),
    '<b>Компания / роль:</b> ' + escapeHtml_(a.role)
  ];
  if (a.recommender) lines.push('<b>Рекомендует:</b> ' + escapeHtml_(a.recommender));
  if (a.portfolio) lines.push('<b>Портфолио:</b> ' + escapeHtml_(a.portfolio));
  if (a.about) lines.push('<b>О себе:</b> ' + escapeHtml_(a.about));
  if (a.source) lines.push('<b>Источник:</b> ' + escapeHtml_(a.source));

  const response = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      chat_id: chatId,
      text: lines.join('\n'),
      parse_mode: 'HTML',
      disable_web_page_preview: true
    }),
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error('Telegram API: ' + response.getResponseCode() + ' ' + response.getContentText().slice(0, 300));
  }
}

function normalizeTelegram_(value) {
  const telegram = clean_(value, 100).replace(/^https?:\/\/t\.me\//i, '').replace(/^@+/, '');
  return telegram ? '@' + telegram : '';
}

function clean_(value, maxLength) {
  return String(value == null ? '' : value).trim().slice(0, maxLength);
}

function escapeHtml_(value) {
  return String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function response_(result) {
  const payload = JSON.stringify(Object.assign({ type: 'tutvse:application-result' }, result)).replace(/</g, '\\u003c');
  return HtmlService
    .createHtmlOutput('<!doctype html><meta charset="utf-8"><script>parent.postMessage(' + payload + ',"*");<\/script>')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
