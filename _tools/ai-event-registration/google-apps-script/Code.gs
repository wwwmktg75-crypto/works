/**
 * AI交流イベント申込 → Googleスプレッドシート追記
 *
 * 【事前準備】
 * 1. 新しいスプレッドシートを作成し、1行目に次の見出しを入力してください:
 *    送信日時(ISO) | お名前 | メール | 職業 | AI活用について | メッセージ
 * 2. 下の SPREADSHEET_ID を、ブラウザのURLにある ID に置き換えてください。
 *    例: https://docs.google.com/spreadsheets/d/【この部分がID】/edit
 * 3. メニュー「デプロイ」→「新しいデプロイ」→ 種類「ウェブアプリ」
 *    - 次のユーザーとして実行: 自分
 *    - アクセスできるユーザー: 全員（匿名ユーザーを含む）
 * 4. デプロイ後に表示される **ウェブアプリのURL** をコピーし、
 *    Vercel の環境変数 GAS_WEBAPP_URL に設定してください。
 */

/** 未設定のままだとエラーになります（URLの /d/ と /edit の間の文字列） */
var SPREADSHEET_ID = '';

/** データを書き込むシート名（存在しなければ最初のシートを使用） */
var SHEET_NAME = '申込';

/**
 * POST で JSON を受け取り、1行追加する
 * Content-Type: application/json
 * ボディ例: { "fullName":"...", "email":"...", ... }
 */
function doPost(e) {
  var output = { ok: false };
  try {
    if (!SPREADSHEET_ID) {
      output.error = 'Code.gs の SPREADSHEET_ID をスプレッドシートのIDに設定してください';
      return jsonResponse_(output);
    }

    if (!e.postData || !e.postData.contents) {
      output.error = 'postData がありません';
      return jsonResponse_(output);
    }

    var data = JSON.parse(e.postData.contents);
    var row = buildRow_(data);

    var sheet = getTargetSheet_();
    sheet.appendRow(row);

    output.ok = true;
    return jsonResponse_(output);
  } catch (err) {
    output.error = String(err);
    return jsonResponse_(output);
  }
}

/** 動作確認用（ブラウザでURLを開いたとき） */
function doGet() {
  return ContentService
    .createTextOutput('AI event GAS: POST JSON で申込を受け付けます。')
    .setMimeType(ContentService.MimeType.TEXT);
}

function buildRow_(data) {
  var submittedAt = data.submittedAt || new Date().toISOString();
  var occupation = occupationLabel_(data.occupation);
  return [
    submittedAt,
    data.fullName || '',
    data.email || '',
    occupation,
    data.aiUsage || '',
    data.message || ''
  ];
}

function occupationLabel_(value) {
  var map = {
    company: '会社員',
    entrepreneur: '経営者・起業家',
    freelance: 'フリーランス',
    student: '学生',
    public: '公務員',
    other: 'その他'
  };
  return map[value] || value || '';
}

function getTargetSheet_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (sheet) {
    return sheet;
  }
  return ss.getSheets()[0];
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
