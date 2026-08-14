/**
 * 「南相馬を、知りなおす。」投稿・承認システム
 *
 * 使い方
 * 1. Googleスプレッドシートの「拡張機能」→「Apps Script」を開く
 * 2. Code.gsの内容をすべて消し、このファイルを丸ごと貼り付ける
 * 3. setupSheet を実行する
 * 4. testNotification を実行してメールを確認する
 * 5. ウェブアプリとしてデプロイする
 */

var CONFIG = {
  SHEET_NAME: '投稿問題',
  STATUSES: ['承認待ち', '承認', '差し戻し', '非公開'],
  CATEGORIES: ['town', 'nomaoi', 'ancient', 'clan', 'modern', 'future'],
  HEADERS: [
    'ID', '状態', '投稿日', 'テーマ', '形式', '問題文',
    '選択肢1', '選択肢2', '選択肢3', '正解番号', '解説',
    '出典名', '出典URL', '投稿者名', '連絡先（非公開）',
    '管理メモ', '承認日'
  ]
};

/**
 * 最初に一度だけ実行します。
 * 管理シートを整え、スプレッドシートIDを保存します。
 */
function setupSheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error('スプレッドシートからApps Scriptを開いてください。');
  }

  PropertiesService.getScriptProperties().setProperty(
    'SPREADSHEET_ID',
    spreadsheet.getId()
  );

  spreadsheet.setSpreadsheetTimeZone('Asia/Tokyo');

  var sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(CONFIG.SHEET_NAME);
  }

  sheet.getRange(1, 1, 1, CONFIG.HEADERS.length).setValues([CONFIG.HEADERS]);
  sheet.setFrozenRows(1);

  sheet.getRange(1, 1, 1, CONFIG.HEADERS.length)
    .setFontWeight('bold')
    .setBackground('#263A31')
    .setFontColor('#FFFFFF')
    .setVerticalAlignment('middle')
    .setWrap(true);

  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.STATUSES, true)
    .setAllowInvalid(false)
    .build();

  sheet.getRange(2, 2, sheet.getMaxRows() - 1, 1).setDataValidation(statusRule);
  sheet.getRange('C:C').setNumberFormat('yyyy/mm/dd hh:mm');
  sheet.getRange('Q:Q').setNumberFormat('yyyy/mm/dd hh:mm');

  setColumnWidths_(sheet);

  return '初期設定が完了しました。';
}

/**
 * 承認済みの問題をWebサイトへ返します。
 */
function doGet(event) {
  try {
    var action = event && event.parameter ? event.parameter.action : '';
    if (action !== 'approved') {
      return jsonResponse_({
        ok: false,
        message: '未対応の操作です。'
      });
    }

    var sheet = getQuestionSheet_();
    var lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      return jsonResponse_({ ok: true, items: [] });
    }

    var rows = sheet
      .getRange(2, 1, lastRow - 1, CONFIG.HEADERS.length)
      .getValues();

    var items = [];

    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      if (row[1] !== '承認') {
        continue;
      }

      var choices = [row[6], row[7]];
      if (row[8] !== '') {
        choices.push(row[8]);
      }

      items.push({
        id: String(row[0]),
        category: String(row[3]),
        type: String(row[4]),
        question: String(row[5]),
        choices: choices,
        correct_index: Number(row[9]),
        explanation: String(row[10]),
        source_title: String(row[11]),
        source_url: String(row[12])
      });
    }

    return jsonResponse_({ ok: true, items: items });
  } catch (error) {
    console.error(error);
    return jsonResponse_({
      ok: false,
      message: '問題を読み込めませんでした。'
    });
  }
}

/**
 * Webサイトの投稿フォームから問題を受け取ります。
 */
function doPost(event) {
  try {
    var rawBody = event && event.postData ? event.postData.contents : '';
    if (!rawBody || rawBody.length > 15000) {
      throw new Error('送信内容が不正です。');
    }

    var input = JSON.parse(rawBody);
    var questionData = validateSubmission_(input);
    var lock = LockService.getScriptLock();

    lock.waitLock(10000);

    try {
      var sheet = getQuestionSheet_();
      rejectDuplicate_(sheet, questionData.question);

      var questionId = Utilities.getUuid();

      sheet.appendRow([
        questionId,
        '承認待ち',
        new Date(),
        questionData.category,
        questionData.type,
        questionData.question,
        questionData.choices[0],
        questionData.choices[1],
        questionData.choices[2] || '',
        questionData.correctIndex,
        questionData.explanation,
        questionData.sourceTitle,
        questionData.sourceUrl,
        questionData.contributor,
        questionData.contact,
        '',
        ''
      ]);

      sendNewSubmissionEmail_(questionId, questionData);

      return jsonResponse_({
        ok: true,
        id: questionId
      });
    } finally {
      lock.releaseLock();
    }
  } catch (error) {
    console.error(error);
    return jsonResponse_({
      ok: false,
      message: error.message || '投稿を受け付けられませんでした。'
    });
  }
}

/**
 * 状態を「承認」に変更した日時を自動記録します。
 */
function onEdit(event) {
  if (!event || !event.range) {
    return;
  }

  var range = event.range;
  var sheet = range.getSheet();

  if (
    sheet.getName() !== CONFIG.SHEET_NAME ||
    range.getRow() < 2 ||
    range.getColumn() !== 2
  ) {
    return;
  }

  var approvedDateCell = sheet.getRange(range.getRow(), 17);

  if (range.getValue() === '承認') {
    approvedDateCell.setValue(new Date());
  } else {
    approvedDateCell.clearContent();
  }
}

/**
 * 通知メールが届くか確認するためのテスト関数です。
 */
function testNotification() {
  var notificationAddress = getNotificationAddress_();

  if (!notificationAddress) {
    throw new Error(
      '通知先を取得できません。プロジェクトの設定でADMIN_EMAILを登録してください。'
    );
  }

  MailApp.sendEmail({
    to: notificationAddress,
    subject: '【南相馬を、知りなおす。】投稿通知テスト',
    body:
      '投稿申請の通知メールは正常に設定されています。\n\n' +
      getSpreadsheet_().getUrl()
  });

  return 'テストメールを送信しました：' + notificationAddress;
}

/**
 * フォーム入力を検証し、安全な値に整えます。
 */
function validateSubmission_(input) {
  if (!input || typeof input !== 'object' || Array.isArray(input)) {
    throw new Error('送信内容が不正です。');
  }

  if (String(input.website || '').trim() !== '') {
    throw new Error('投稿を受け付けられませんでした。');
  }

  if (input.agreement !== true) {
    throw new Error('投稿内容の確認・編集・公開への同意が必要です。');
  }

  var startedAt = Number(input.started_at);
  var elapsed = Date.now() - startedAt;

  if (!isFinite(startedAt) || elapsed < 2000 || elapsed > 86400000) {
    throw new Error('ページを再読み込みして、もう一度お試しください。');
  }

  var category = cleanText_(input.category, 20);
  var type = cleanText_(input.type, 10);
  var question = cleanText_(input.question, 180);
  var explanation = cleanText_(input.explanation, 600);
  var sourceTitle = cleanText_(input.source_title, 100);
  var sourceUrl = cleanText_(input.source_url, 500);
  var contributor = cleanText_(input.contributor, 40);
  var contact = cleanText_(input.contact, 150);
  var correctIndex = Number(input.correct_index);
  var choices = [];

  if (Array.isArray(input.choices)) {
    for (var i = 0; i < input.choices.length; i++) {
      choices.push(cleanText_(input.choices[i], 100));
    }
  }

  if (CONFIG.CATEGORIES.indexOf(category) === -1) {
    throw new Error('テーマを選び直してください。');
  }

  if (type !== 'ox' && type !== 'three') {
    throw new Error('出題形式を選び直してください。');
  }

  if (question.length < 8) {
    throw new Error('問題文は8文字以上で入力してください。');
  }

  if (explanation.length < 10) {
    throw new Error('解説は10文字以上で入力してください。');
  }

  if (!sourceTitle) {
    throw new Error('出典名を入力してください。');
  }

  if (!/^https:\/\//i.test(sourceUrl)) {
    throw new Error('出典URLは https:// で始めてください。');
  }

  if (contact && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(contact)) {
    throw new Error('連絡先メールを確認してください。');
  }

  if (type === 'ox') {
    if (
      choices.length !== 2 ||
      choices[0] !== '○' ||
      choices[1] !== '×' ||
      (correctIndex !== 0 && correctIndex !== 1)
    ) {
      throw new Error('○×問題の答えが不正です。');
    }
  }

  if (type === 'three') {
    if (
      choices.length !== 3 ||
      !choices[0] ||
      !choices[1] ||
      !choices[2] ||
      choices[0] === choices[1] ||
      choices[0] === choices[2] ||
      choices[1] === choices[2] ||
      (correctIndex !== 0 && correctIndex !== 1 && correctIndex !== 2)
    ) {
      throw new Error('三つの異なる選択肢と正解を入力してください。');
    }
  }

  return {
    category: category,
    type: type,
    question: safeCellText_(question),
    choices: [
      safeCellText_(choices[0]),
      safeCellText_(choices[1]),
      safeCellText_(choices[2] || '')
    ],
    correctIndex: correctIndex,
    explanation: safeCellText_(explanation),
    sourceTitle: safeCellText_(sourceTitle),
    sourceUrl: sourceUrl,
    contributor: safeCellText_(contributor),
    contact: safeCellText_(contact)
  };
}

/**
 * 同じ問題文が登録済みなら投稿を止めます。
 */
function rejectDuplicate_(sheet, question) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return;
  }

  var target = normalizeText_(question);
  var questions = sheet
    .getRange(2, 6, lastRow - 1, 1)
    .getDisplayValues();

  for (var i = 0; i < questions.length; i++) {
    if (normalizeText_(questions[i][0]) === target) {
      throw new Error('同じ問題文がすでに投稿されています。');
    }
  }
}

/**
 * 新しい投稿が届いたときに管理者へメールを送ります。
 */
function sendNewSubmissionEmail_(questionId, questionData) {
  var notificationAddress = getNotificationAddress_();
  if (!notificationAddress) {
    console.error('通知先メールアドレスを取得できませんでした。');
    return;
  }

  try {
    MailApp.sendEmail({
      to: notificationAddress,
      subject: '【南相馬を、知りなおす。】新しい問題が届きました',
      body:
        '承認待ちの問題が1件届きました。\n\n' +
        '問題：' + questionData.question + '\n' +
        '管理番号：' + questionId + '\n\n' +
        'スプレッドシートで内容と出典を確認してください。\n' +
        getSpreadsheet_().getUrl()
    });
  } catch (error) {
    console.error('通知メールを送信できませんでした。', error);
  }
}

/**
 * ADMIN_EMAILが設定されていれば優先し、なければスクリプト所有者を使います。
 */
function getNotificationAddress_() {
  var configuredAddress = PropertiesService
    .getScriptProperties()
    .getProperty('ADMIN_EMAIL');

  if (configuredAddress) {
    return configuredAddress;
  }

  return Session.getEffectiveUser().getEmail();
}

function getSpreadsheet_() {
  var spreadsheetId = PropertiesService
    .getScriptProperties()
    .getProperty('SPREADSHEET_ID');

  if (!spreadsheetId) {
    throw new Error('先に setupSheet を実行してください。');
  }

  return SpreadsheetApp.openById(spreadsheetId);
}

function getQuestionSheet_() {
  var sheet = getSpreadsheet_().getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    throw new Error('投稿問題シートが見つかりません。');
  }

  return sheet;
}

function setColumnWidths_(sheet) {
  var widths = [
    230, 110, 145, 110, 90, 420, 140, 140, 140,
    90, 420, 210, 300, 140, 220, 280, 145
  ];

  for (var i = 0; i < widths.length; i++) {
    sheet.setColumnWidth(i + 1, widths[i]);
  }
}

function cleanText_(value, maxLength) {
  var text = value === null || value === undefined ? '' : String(value);
  return text.trim().slice(0, maxLength);
}

function normalizeText_(value) {
  return String(value)
    .normalize('NFKC')
    .replace(/\s+/g, '')
    .toLowerCase();
}

/**
 * =、+、-、@で始まる投稿をスプレッドシート数式として解釈させません。
 */
function safeCellText_(value) {
  var text = value === null || value === undefined ? '' : String(value);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function jsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
