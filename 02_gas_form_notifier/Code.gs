/**
 * Googleフォームの回答を自動で整理し、担当者へメールで知らせるスクリプト
 *
 * 動き:
 *   1. フォームに回答が届く
 *   2. 「受付一覧」シートに受付番号付きで1行追加する
 *   3. 問い合わせの種類に応じた担当者へ通知メールを送る
 *   4. （設定すれば）回答者へ受付完了の自動返信メールを送る
 *
 * 設置方法は README.md を参照。
 */

// ===== 設定（ここだけ書きかえれば使えます） =====
const CONFIG = {
  SHEET_NAME: '受付一覧',
  TIME_ZONE: 'Asia/Tokyo',
  TICKET_PREFIX: 'Q',

  // フォームの質問名 → このスクリプトで使う名前
  QUESTIONS: {
    name: 'お名前',
    email: 'メールアドレス',
    category: 'お問い合わせの種類',
    message: 'お問い合わせ内容',
  },

  // 問い合わせの種類ごとの通知先（当てはまらなければ DEFAULT_TO へ）
  ROUTING: {
    '商品について': 'sales@example.com',
    '請求・お支払い': 'billing@example.com',
    '不具合の報告': 'support@example.com',
  },
  DEFAULT_TO: 'info@example.com',

  SEND_AUTO_REPLY: true,
  SHOP_NAME: 'サンプル商店',
};

const HEADERS = ['受付番号', '受付日時', 'お名前', 'メールアドレス', '種類', '内容', '担当', '対応状況'];

// ===== 初回だけ実行：シートの準備とトリガーの設定 =====
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#e8f0fe');
  }

  // 同じトリガーが二重に登録されないようにする
  ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === 'onFormSubmit')
    .forEach((t) => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(ss).onFormSubmit().create();
}

// ===== フォーム送信時に自動で動く =====
function onFormSubmit(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30 * 1000); // 同時送信で受付番号が重ならないようにする
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
    const answers = normalizeAnswers(e.namedValues, CONFIG.QUESTIONS);
    const now = new Date();
    const today = formatDate_(now, 'yyyyMMdd');

    const existingIds = sheet.getLastRow() > 1
      ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().map((r) => String(r[0]))
      : [];
    const ticketId = makeTicketId(CONFIG.TICKET_PREFIX, today, existingIds);
    const to = pickRecipient(answers.category, CONFIG.ROUTING, CONFIG.DEFAULT_TO);

    const record = {
      ticketId: ticketId,
      receivedAt: formatDate_(now, 'yyyy/MM/dd HH:mm'),
      ...answers,
      assignee: to,
    };
    sheet.appendRow(toRow(record));

    const notify = buildNotifyMail(record, sheet.getParent().getUrl());
    MailApp.sendEmail(to, notify.subject, notify.body);

    if (CONFIG.SEND_AUTO_REPLY && isValidEmail(answers.email)) {
      const reply = buildReplyMail(record, CONFIG.SHOP_NAME);
      MailApp.sendEmail(answers.email, reply.subject, reply.body);
    }
  } finally {
    lock.releaseLock();
  }
}

// ===== ここから下は、テストしやすいよう Google のサービスを使わない関数 =====

/** フォームの回答（{質問名: [回答]}）を、扱いやすい形に変える */
function normalizeAnswers(namedValues, questions) {
  const get = (title) => {
    const v = namedValues[title];
    return (Array.isArray(v) ? v[0] : v || '').toString().trim();
  };
  return {
    name: get(questions.name).replace(/　/g, ' ').replace(/ +/g, ' '),
    email: get(questions.email).toLowerCase(),
    category: get(questions.category),
    message: get(questions.message),
  };
}

/** その日の連番付きの受付番号を作る（例：Q-20261121-003） */
function makeTicketId(prefix, yyyymmdd, existingIds) {
  const head = prefix + '-' + yyyymmdd + '-';
  const maxSeq = existingIds
    .filter((id) => id.indexOf(head) === 0)
    .map((id) => parseInt(id.slice(head.length), 10))
    .filter((n) => !isNaN(n))
    .reduce((a, b) => Math.max(a, b), 0);
  return head + String(maxSeq + 1).padStart(3, '0');
}

/** 問い合わせの種類から通知先を決める */
function pickRecipient(category, routing, defaultTo) {
  return Object.prototype.hasOwnProperty.call(routing, category) ? routing[category] : defaultTo;
}

function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function toRow(r) {
  return [r.ticketId, r.receivedAt, r.name, r.email, r.category, r.message, r.assignee, '未対応'];
}

function buildNotifyMail(r, sheetUrl) {
  return {
    subject: '【新着問い合わせ】' + r.ticketId + '（' + (r.category || '種類なし') + '）',
    body: [
      '新しいお問い合わせが届きました。',
      '',
      '受付番号：' + r.ticketId,
      '受付日時：' + r.receivedAt,
      'お名前　：' + r.name,
      'メール　：' + r.email,
      '種類　　：' + r.category,
      '',
      '--- 内容 ---',
      r.message,
      '',
      '受付一覧：' + sheetUrl,
    ].join('\n'),
  };
}

function buildReplyMail(r, shopName) {
  return {
    subject: '【' + shopName + '】お問い合わせを受け付けました（' + r.ticketId + '）',
    body: [
      r.name + ' 様',
      '',
      'お問い合わせいただき、ありがとうございます。',
      '以下の内容で受け付けました。担当者より順次ご連絡いたします。',
      '',
      '受付番号：' + r.ticketId,
      '種類　　：' + r.category,
      '',
      '--- お問い合わせ内容 ---',
      r.message,
      '',
      '※このメールは送信専用です。',
      shopName,
    ].join('\n'),
  };
}

function formatDate_(date, pattern) {
  return Utilities.formatDate(date, CONFIG.TIME_ZONE, pattern);
}
