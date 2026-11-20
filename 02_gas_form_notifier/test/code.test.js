// Code.gs を Google のサービスの「にせもの」と一緒に読み込んでテストする
// 実行: node --test test/code.test.js
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function loadGas(mocks) {
  const code = fs.readFileSync(path.join(__dirname, '..', 'Code.gs'), 'utf8');
  const context = vm.createContext({ ...mocks });
  // const で宣言した設定もテストから見えるように、最後に取り出す
  vm.runInContext(code + '\n;this.CONFIG = CONFIG; this.HEADERS = HEADERS;', context);
  return context;
}

function fakeSheet(initialRows) {
  const rows = initialRows.map((r) => r.slice());
  return {
    rows,
    getLastRow: () => rows.length,
    getRange: (row, col, numRows) => ({
      getValues: () => rows.slice(row - 1, row - 1 + numRows).map((r) => [r[col - 1]]),
    }),
    appendRow: (r) => rows.push(r),
    getParent: () => ({ getUrl: () => 'https://docs.google.com/spreadsheets/d/xxx' }),
  };
}

function makeMocks(sheet, sentMails) {
  return {
    SpreadsheetApp: { getActiveSpreadsheet: () => ({ getSheetByName: () => sheet }) },
    MailApp: { sendEmail: (to, subject, body) => sentMails.push({ to, subject, body }) },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    Utilities: {
      formatDate: (d, tz, pattern) => {
        const p = (n) => String(n).padStart(2, '0');
        // テストでは 2026-11-21 10:05 に固定
        const map = { yyyy: '2026', MM: '11', dd: '21', HH: p(10), mm: p(5) };
        return pattern.replace(/yyyy|MM|dd|HH|mm/g, (k) => map[k]);
      },
    },
  };
}

test('受付番号はその日の連番になる', () => {
  const gas = loadGas({});
  assert.equal(gas.makeTicketId('Q', '20261121', []), 'Q-20261121-001');
  assert.equal(
    gas.makeTicketId('Q', '20261121', ['Q-20261120-007', 'Q-20261121-001', 'Q-20261121-002']),
    'Q-20261121-003',
  );
  // 行が削除されて番号が飛んでいても、重ならない
  assert.equal(gas.makeTicketId('Q', '20261121', ['Q-20261121-005']), 'Q-20261121-006');
});

test('回答の空白や大文字をそろえる', () => {
  const gas = loadGas({});
  const a = gas.normalizeAnswers(
    {
      'お名前': ['　山田　　太郎 '],
      'メールアドレス': [' Taro@Example.com'],
      'お問い合わせの種類': ['商品について'],
      'お問い合わせ内容': ['在庫はありますか？'],
    },
    gas.CONFIG.QUESTIONS,
  );
  assert.equal(a.name, '山田 太郎');
  assert.equal(a.email, 'taro@example.com');
});

test('種類に応じて通知先が変わる', () => {
  const gas = loadGas({});
  const { ROUTING, DEFAULT_TO } = gas.CONFIG;
  assert.equal(gas.pickRecipient('不具合の報告', ROUTING, DEFAULT_TO), 'support@example.com');
  assert.equal(gas.pickRecipient('その他', ROUTING, DEFAULT_TO), 'info@example.com');
  assert.equal(gas.pickRecipient('toString', ROUTING, DEFAULT_TO), 'info@example.com');
});

test('フォーム送信で、シートに1行追加され、担当者と回答者にメールが届く', () => {
  const sent = [];
  const sheet = fakeSheet([
    ['受付番号', '受付日時', 'お名前', 'メールアドレス', '種類', '内容', '担当', '対応状況'],
    ['Q-20261121-001', '2026/11/21 09:00', '鈴木 花子', 'hanako@example.com', '商品について', '…', 'sales@example.com', '対応済み'],
  ]);
  const gas = loadGas(makeMocks(sheet, sent));

  gas.onFormSubmit({
    namedValues: {
      'お名前': ['山田 太郎'],
      'メールアドレス': ['taro@example.com'],
      'お問い合わせの種類': ['請求・お支払い'],
      'お問い合わせ内容': ['請求書を再発行してください'],
    },
  });

  assert.equal(sheet.rows.length, 3);
  // vm の中で作られた配列は別環境のオブジェクトなので、コピーしてから比べる
  assert.deepEqual(Array.from(sheet.rows[2]), [
    'Q-20261121-002', '2026/11/21 10:05', '山田 太郎', 'taro@example.com',
    '請求・お支払い', '請求書を再発行してください', 'billing@example.com', '未対応',
  ]);

  assert.equal(sent.length, 2);
  assert.equal(sent[0].to, 'billing@example.com');
  assert.match(sent[0].subject, /Q-20261121-002/);
  assert.equal(sent[1].to, 'taro@example.com');
  assert.match(sent[1].body, /山田 太郎 様/);
});

test('メールアドレスが不正なら自動返信は送らない', () => {
  const sent = [];
  const sheet = fakeSheet([['受付番号']]);
  const gas = loadGas(makeMocks(sheet, sent));
  gas.onFormSubmit({
    namedValues: { 'お名前': ['A'], 'メールアドレス': ['abc'], 'お問い合わせの種類': ['その他'], 'お問い合わせ内容': ['x'] },
  });
  assert.equal(sent.length, 1);
  assert.equal(sent[0].to, 'info@example.com');
  assert.equal(sheet.rows[1][0], 'Q-20261121-001');
});
