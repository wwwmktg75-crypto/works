/**
 * 経営管理ダッシュボードシステム★AI-chan作業中 2025/02/18 (火) 16:00 更新★
 * ※ clasp push 前に上記の更新日時を必ず現在日時に書き換えること
 */

// ダッシュボード用スプレッドシートID
// 別プロジェクト用: GASエディタで「プロジェクトの設定」→「スクリプト プロパティ」に
// SPREADSHEET_ID = (新しいスプレッドシートのID) を追加すると、ここを編集せずに切り替え可能
function getSpreadsheetId() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (id && id.trim()) return id.trim();
  return '1fG_fXkXqY7lQlUyfk4e-ilExr_s0y0cmGvE--TQSmjg'; // デフォルト（未設定時）
}

// ユーザー管理用スプレッドシートID（ユーザー情報を保存するシート）
const USER_SHEET_NAME = 'ユーザー管理';

// 顧客データ用スプレッドシートID（ダッシュボード用と同じ）
// 顧客マスタシート名
const CUSTOMER_MASTER_SHEET_NAME = '顧客マスタ';
// 顧客継続シート名
const CUSTOMER_CONTINUE_SHEET_NAME = '顧客継続';

// スタッフマスタシート名
const STAFF_MASTER_SHEET_NAME = 'スタッフマスタ';

// 実績報告シート名
const PERFORMANCE_REPORT_SHEET_NAME = '実績報告';

/**
 * スプレッドシートの全シート名を取得（デバッグ用）
 */
function getSheetNames() {
  try {
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const sheets = ss.getSheets();
    return sheets.map(s => s.getName());
  } catch (e) {
    return ['エラー: ' + e.toString()];
  }
}

// ========================================
// パフォーマンス最適化: キャッシュユーティリティ
// ========================================

/**
 * キャッシュからデータを取得、なければ取得関数を実行してキャッシュに保存
 * @param {string} key - キャッシュキー
 * @param {Function} fetchFunction - データ取得関数
 * @param {number} expirationSeconds - キャッシュ有効期限（秒）、デフォルト1800秒（30分）
 * @return {*} キャッシュされたデータまたは取得したデータ
 */
function getCachedData(key, fetchFunction, expirationSeconds = 1800) {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get(key);
    
    if (cached) {
      try {
        return JSON.parse(cached);
      } catch (parseError) {
        console.warn(`キャッシュ解析エラー (${key}):`, parseError);
        // キャッシュが壊れている場合は削除して再取得
        cache.remove(key);
      }
    }
    
    const data = fetchFunction();
    try {
      cache.put(key, JSON.stringify(data), expirationSeconds);
    } catch (cacheError) {
      console.warn(`キャッシュ保存エラー (${key}):`, cacheError);
      // キャッシュ保存に失敗してもデータは返す
    }
    return data;
  } catch (error) {
    console.error(`キャッシュ処理エラー (${key}):`, error);
    // エラー時は直接取得関数を実行
    return fetchFunction();
  }
}

/**
 * キャッシュをクリア
 * @param {string} key - キャッシュキー（省略時は全クリア）
 */
function clearCache(key = null) {
  try {
    const cache = CacheService.getScriptCache();
    if (key) {
      cache.remove(key);
    } else {
      cache.removeAll(['store_list', 'performance_reports_*', 'staff_reports_*']);
    }
  } catch (error) {
    console.error('キャッシュクリアエラー:', error);
  }
}

/**
 * Webアプリのエントリーポイント
 */
function doGet(e) {
  try {
    // パラメータを安全に取得
    const params = e && e.parameter ? e.parameter : {};
    let page = params.page || '';
    const store = params.store || '';
    const sessionId = params.sessionId || null;
    
    // WebアプリのベースURLを取得（常に最新のデプロイURL）
    const baseUrl = ScriptApp.getService().getUrl();
    
    // ページ未指定時: 未ログインなら必ずログイン画面へリダイレクト（自社ドメインから開いても常にログインから入る）
    if (page === '' || page === 'index') {
      let session = null;
      try {
        if (sessionId) session = getSession(sessionId);
        if (!session) session = getSession();
      } catch (err) { session = null; }
      if (!session) {
        const loginUrl = baseUrl + '?page=login';
        return HtmlService.createHtmlOutput(
          '<!DOCTYPE html><html><head><meta charset="UTF-8"><meta http-equiv="refresh" content="0;url=' + loginUrl + '"><title>ログインへ</title></head><body><p>ログイン画面へ移動しています...</p><script>location.replace(' + JSON.stringify(loginUrl) + ');</script></body></html>'
        ).setTitle('ログイン');
      }
      page = 'index';
    }
    
    // ログインページの場合は直接表示（セッション確認不要）
    if (page === 'login') {
      try {
        const template = HtmlService.createTemplateFromFile('login');
        template.deployUrl = baseUrl;
        return template.evaluate()
          .setTitle('ログイン')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      } catch (error) {
        console.error('ログインページ読み込みエラー:', error);
        // エラーが発生した場合もログイン画面を返す（簡易版）
        return HtmlService.createHtmlOutput(`
          <!DOCTYPE html>
          <html>
          <head>
            <meta charset="UTF-8">
            <title>ログイン</title>
            <style>
              body { font-family: sans-serif; padding: 20px; text-align: center; }
              .error { color: red; }
            </style>
          </head>
          <body>
            <h1>ログイン</h1>
            <p class="error">ログインページの読み込みに失敗しました。ページを再読み込みしてください。</p>
            <button onclick="window.location.reload()">再読み込み</button>
          </body>
          </html>
        `).setTitle('ログイン');
      }
    }
    
    // ログアウトページ
    if (page === 'logout') {
      try {
        clearSession();
      } catch (error) {
        console.error('セッションクリアエラー:', error);
      }
      const template = HtmlService.createTemplateFromFile('login');
      template.deployUrl = baseUrl;
      return template.evaluate()
        .setTitle('ログイン')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    
    // report-detail はセッション不要でページを表示（クライアント側で認証処理）
    if (page === 'report-detail') {
      const rdTemplate = HtmlService.createTemplateFromFile('report-detail');
      rdTemplate.deployUrl = baseUrl;
      rdTemplate.reportId = params.id || '';
      rdTemplate.reporter = params.reporter || '';
      rdTemplate.reportDateTime = params.reportDateTime || params.timestamp || '';
      rdTemplate.sessionId = params.sessionId || '';
      return rdTemplate
        .evaluate()
        .setTitle('実績報告詳細')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    
    // その他のページはセッション確認
    let session = null;
    try {
      // sessionIdが指定されている場合はそれで取得
      if (sessionId) {
        session = getSession(sessionId);
      }
      // sessionIdで見つからない場合は current_session を試す
      if (!session) {
        session = getSession();
      }
    } catch (error) {
      console.error('セッション確認エラー:', error);
      session = null;
    }
    
    if (!session) {
      // ログインしていない場合はログイン画面へ
      const template = HtmlService.createTemplateFromFile('login');
      template.deployUrl = baseUrl;
      return template.evaluate()
        .setTitle('ログイン')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    
    // セッションがある場合、各ページを表示
    try {
      if (page === 'dashboard') {
        const template = HtmlService.createTemplateFromFile('dashboard');
        template.deployUrl = baseUrl;
        if (store) {
          template.store = store;
        }
        return template
          .evaluate()
          .setTitle('店舗実績ダッシュボード')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      } else if (page === 'staff') {
        const template = HtmlService.createTemplateFromFile('staff-dashboard');
        template.deployUrl = baseUrl;
        return template
          .evaluate()
          .setTitle('スタッフ別ダッシュボード')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      } else if (page === 'report') {
        const template = HtmlService.createTemplateFromFile('report');
        template.deployUrl = baseUrl;
        return template
          .evaluate()
          .setTitle('実績報告')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      } else if (page === 'customer-register') {
        const template = HtmlService.createTemplateFromFile('customer-register');
        template.deployUrl = baseUrl;
        return template
          .evaluate()
          .setTitle('顧客マスタ登録')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      } else if (page === 'customer-list') {
        const template = HtmlService.createTemplateFromFile('customer-list');
        template.deployUrl = baseUrl;
        return template
          .evaluate()
          .setTitle('顧客一覧')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      } else if (page === 'customer-detail') {
        const template = HtmlService.createTemplateFromFile('customer-detail');
        template.deployUrl = baseUrl;
        // URLパラメータから顧客IDを取得してテンプレートに埋め込む
        template.customerId = params.id || '';
        console.log('customer-detail: URLパラメータのid:', params.id);
        return template
          .evaluate()
          .setTitle('顧客詳細')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      } else if (page === 'customer-continue') {
        const template = HtmlService.createTemplateFromFile('customer-continue');
        template.deployUrl = baseUrl;
        return template
          .evaluate()
          .setTitle('顧客継続情報登録')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      } else if (page === 'staff-report-list') {
        const template = HtmlService.createTemplateFromFile('staff-report-list');
        template.deployUrl = baseUrl;
        return template
          .evaluate()
          .setTitle('スタッフ別実績報告一覧')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      } else if (page === 'staff-report-detail') {
        const template = HtmlService.createTemplateFromFile('staff-report-detail');
        template.deployUrl = baseUrl;
        return template
          .evaluate()
          .setTitle('スタッフ別実績報告詳細')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      } else if (page === 'report-list') {
        // 実績報告一覧（FCオーナー＝全件、オーナー＝自店舗のみ、スタッフは閲覧不可）
        if (session.role === 'スタッフ') {
          const template = HtmlService.createTemplateFromFile('index');
          template.deployUrl = baseUrl;
          return template.evaluate()
            .setTitle('トレーナーズジム　経営管理ダッシュボード')
            .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
        }
        const template = HtmlService.createTemplateFromFile('report-list');
        template.deployUrl = baseUrl;
        return template
          .evaluate()
          .setTitle('実績報告一覧')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      } else if (page === 'reward') {
        // 報酬管理ページ（スタッフには非表示）
        if (session.role === 'スタッフ') {
          // スタッフの場合はindexにリダイレクト
          const template = HtmlService.createTemplateFromFile('index');
          template.deployUrl = baseUrl;
          return template.evaluate()
            .setTitle('トレーナーズジム　経営管理ダッシュボード')
            .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
        }
        const template = HtmlService.createTemplateFromFile('reward');
        template.deployUrl = baseUrl;
        return template
          .evaluate()
          .setTitle('報酬管理')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      } else {
        // indexページ
        try {
          const template = HtmlService.createTemplateFromFile('index');
          template.deployUrl = baseUrl;
          const htmlOutput = template.evaluate();
          return htmlOutput
            .setTitle('トレーナーズジム　経営管理ダッシュボード')
            .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
        } catch (templateError) {
          console.error('index.htmlテンプレート評価エラー:', templateError);
          console.error('エラースタック:', templateError.stack);
          // エラーが発生した場合は簡易版のHTMLを返す
          return HtmlService.createHtmlOutput(`
            <!DOCTYPE html>
            <html>
            <head>
              <meta charset="UTF-8">
              <title>エラー</title>
              <style>
                body { font-family: sans-serif; padding: 20px; text-align: center; }
                .error { color: red; }
              </style>
            </head>
            <body>
              <h1>エラー</h1>
              <p class="error">ページの読み込みに失敗しました。</p>
              <p>エラー: ${templateError.message || templateError.toString()}</p>
              <button onclick="window.location.reload()">再読み込み</button>
            </body>
            </html>
          `).setTitle('エラー');
        }
      }
    } catch (error) {
      console.error('ページ読み込みエラー:', error);
      // ページ読み込みに失敗した場合はログイン画面へ
      return HtmlService.createTemplateFromFile('login')
        .evaluate()
        .setTitle('ログイン')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
  } catch (error) {
    console.error('doGet重大エラー:', error);
    console.error('エラースタック:', error.stack);
    // エラーが発生した場合はログイン画面を表示（簡易版）
    return HtmlService.createHtmlOutput(`
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="UTF-8">
        <title>ログイン</title>
        <style>
          body { font-family: sans-serif; padding: 20px; text-align: center; }
          .error { color: red; }
        </style>
      </head>
      <body>
        <h1>ログイン</h1>
        <p class="error">システムエラーが発生しました。ページを再読み込みしてください。</p>
        <button onclick="window.location.href='?page=login'">ログイン画面へ</button>
      </body>
      </html>
    `).setTitle('ログイン');
  }
}

/**
 * HTMLファイルにCSSやJSをインクルードするためのヘルパー関数
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * 年月を正規化する（様々なフォーマットを「YYYY年M月」形式に統一）
 */
function normalizeYearMonth(value) {
  if (!value) return '';
  
  // Dateオブジェクトの場合（スプレッドシートから日付として取得される場合）
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    const year = value.getFullYear();
    const month = value.getMonth() + 1;
    return `${year}年${month}月`;
  }
  
  const str = String(value).trim();
  
  // 空文字の場合
  if (!str) return '';
  
  // 既に「YYYY年M月」形式の場合はそのまま返す
  if (/^\d{4}年\d{1,2}月$/.test(str)) {
    return str;
  }
  
  // 「YYYY/MM/DD」形式の場合
  const slashMatch = str.match(/^(\d{4})\/(\d{1,2})\/\d{1,2}$/);
  if (slashMatch) {
    const year = slashMatch[1];
    const month = parseInt(slashMatch[2], 10);
    return `${year}年${month}月`;
  }
  
  // 「YYYY/MM」形式の場合
  const slashMatch2 = str.match(/^(\d{4})\/(\d{1,2})$/);
  if (slashMatch2) {
    const year = slashMatch2[1];
    const month = parseInt(slashMatch2[2], 10);
    return `${year}年${month}月`;
  }
  
  // 「YYYY-MM-DD」形式の場合
  const dashMatch = str.match(/^(\d{4})-(\d{1,2})-\d{1,2}$/);
  if (dashMatch) {
    const year = dashMatch[1];
    const month = parseInt(dashMatch[2], 10);
    return `${year}年${month}月`;
  }
  
  // 「YYYY-MM」形式の場合
  const dashMatch2 = str.match(/^(\d{4})-(\d{1,2})$/);
  if (dashMatch2) {
    const year = dashMatch2[1];
    const month = parseInt(dashMatch2[2], 10);
    return `${year}年${month}月`;
  }
  
  // 数値（シリアル値）の場合（Excelの日付シリアル値）
  const numValue = Number(str);
  if (!isNaN(numValue) && numValue > 40000 && numValue < 50000) {
    // Excelのシリアル値をDateに変換
    const date = new Date((numValue - 25569) * 86400 * 1000);
    const year = date.getFullYear();
    const month = date.getMonth() + 1;
    return `${year}年${month}月`;
  }
  
  return str;
}

/**
 * 店舗リストを取得（アクセス権限を考慮）
 */
function getStoreList(sessionId = null) {
  // セッションIDが指定されている場合はセッションを設定
  if (sessionId) {
    try {
      setSession(sessionId);
    } catch (error) {
      console.error('セッション設定エラー:', error);
    }
  }
  
  // 権限チェック
  const user = getCurrentUser();
  if (!user) {
    return [];
  }
  
  // アクセス可能な店舗を取得（権限に基づいてフィルタリング）
  const accessibleStores = getAccessibleStores();
  return accessibleStores;
}

/**
 * 店舗の年間売上トレンドを取得
 */
function getStoreYearlyTrend(storeName, sessionId = null) {
  // セッションIDが指定されている場合はセッションを設定
  if (sessionId) {
    try {
      setSession(sessionId);
    } catch (error) {
      console.error('セッション設定エラー:', error);
    }
  }
  
  // 権限チェック
  if (!canAccessStore(storeName)) {
    return {};
  }
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  const sheet = ss.getSheetByName('報告データシート履歴');
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const getIdx = (name) => headers.indexOf(name);
  const idx = {
    month: getIdx('年月'),
    store: getIdx('店舗'),
    amount: getIdx('金額')
  };
  
  const trendData = {};
  
  data.slice(1).forEach(row => {
    const store = String(row[idx.store]);
    const month = normalizeYearMonth(row[idx.month]); // 正規化
    const amount = Number(row[idx.amount]) || 0;
    
    if (store === storeName && month) {
      trendData[month] = (trendData[month] || 0) + amount;
    }
  });
  
  return trendData;
}

/**
 * 月別詳細データを取得
 */
function getMonthlyDetails(storeName, targetMonth, sessionId = null) {
  // セッションIDが指定されている場合はセッションを設定
  if (sessionId) {
    try {
      setSession(sessionId);
    } catch (error) {
      console.error('セッション設定エラー:', error);
    }
  }
  
  // 権限チェック
  if (!canAccessStore(storeName)) {
    return { error: 'この店舗へのアクセス権限がありません', summary: {}, salesDetails: [], inquiryStatus: [] };
  }
  try {
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const sheet = ss.getSheetByName('報告データシート履歴');
    if (!sheet) return { error: '報告データシート履歴が見つかりません', summary: {}, salesDetails: [], inquiryStatus: [] };
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // デバッグログ
    console.log(`getMonthlyDetails called: storeName="${storeName}", targetMonth="${targetMonth}"`);
  
    // 各列のインデックスをヘッダー名から動的に取得
    const getIdx = (name) => headers.indexOf(name);
  
    const idx = {
      month: getIdx('年月'),
      store: getIdx('店舗'),
      type: getIdx('種別'),
      payDate: getIdx('入金日'),
      name: getIdx('お名前'),
      age: getIdx('年齢'),
      course: getIdx('コース'),
      amount: getIdx('金額'),
      note: getIdx('補足/備考'),
      staff: getIdx('スタッフ氏名'),
      purpose: getIdx('目的（新規のみ）'),
      contReason: getIdx('継続理由（継続のみ）'),
      system: getIdx('制度選択'),
      session: getIdx('セッション本数')
    };

    // 年月を正規化して比較
    const filteredRows = data.slice(1).filter(row => {
      const rowMonth = normalizeYearMonth(row[idx.month]);
      return String(row[idx.store]) === storeName && rowMonth === targetMonth;
    });
  
    // デバッグ: フィルタ結果
    console.log(`Filtered rows count: ${filteredRows.length}`);

    const summary = { totalSales: 0, totalSessions: 0, newCount: 0, renewCount: 0 };
    const salesDetails = [];
    const inquiryStatus = [];
    const salesDetailsByStaff = {};
    const staffSet = new Set();

    filteredRows.forEach(row => {
      const amount = Number(row[idx.amount]) || 0;
      const type = String(row[idx.type]);
      const staff = String(row[idx.staff] || '').trim();

      summary.totalSales += amount;
      summary.totalSessions += (Number(row[idx.session]) || 0);
      if (type.includes('新規')) summary.newCount++;
      if (type.includes('継続')) summary.renewCount++;

      // 全ての項目を含む詳細データを作成
      // 注意: すべてのフィールドを明示的にプリミティブ型に変換（Dateオブジェクトはシリアライズ不可）
      let payDateStr = '-';
      try {
        if (row[idx.payDate]) {
          const pd = row[idx.payDate];
          if (Object.prototype.toString.call(pd) === '[object Date]' && !isNaN(pd)) {
            payDateStr = Utilities.formatDate(pd, "JST", "MM/dd");
          } else {
            payDateStr = String(pd);
          }
        }
      } catch (e) {
        payDateStr = '-';
      }
    
      // 年齢フィールドの処理（Dateオブジェクトが入っている場合がある）
      let ageValue = '-';
      const rawAge = row[idx.age];
      if (rawAge !== null && rawAge !== undefined && rawAge !== '') {
        if (Object.prototype.toString.call(rawAge) === '[object Date]') {
          ageValue = '-'; // Dateオブジェクトは無視
        } else if (typeof rawAge === 'number') {
          ageValue = rawAge;
        } else {
          ageValue = String(rawAge);
        }
      }
    
      const detail = {
        type: String(type || '-'),
        payDate: payDateStr,
        name: String(row[idx.name] || '-'),
        age: ageValue,
        course: String(row[idx.course] || '-'),
        amount: Number(amount) || 0,
        note: String(row[idx.note] || '-'),
        staff: staff || '未設定',
        purpose: String(row[idx.purpose] || '-'),
        contReason: String(row[idx.contReason] || '-'),
        system: String(row[idx.system] || '-')
      };

      if (amount > 0 || type.includes('新規') || type.includes('継続')) {
        salesDetails.push(detail);
        
        // スタッフ別にグループ化
        const staffKey = staff || '未設定';
        staffSet.add(staffKey);
        if (!salesDetailsByStaff[staffKey]) {
          salesDetailsByStaff[staffKey] = [];
        }
        salesDetailsByStaff[staffKey].push({
          name: detail.name,
          amount: detail.amount,
          type: detail.type
        });
      }
      if (type.includes('新規')) {
        inquiryStatus.push(detail);
      }
    });

    // スタッフリストをソート
    const staffList = Array.from(staffSet).sort((a, b) => {
      if (a.includes('オーナー') || a.includes('古林')) return -1;
      if (b.includes('オーナー') || b.includes('古林')) return 1;
      return a.localeCompare(b);
    });

    // 返却直前のデバッグログ
    console.log(`Return data - summary.totalSales: ${summary.totalSales}, salesDetails.length: ${salesDetails.length}`);
  
    const result = { summary, salesDetails, inquiryStatus, salesDetailsByStaff, staffList };
    console.log(`Returning result object with keys: ${Object.keys(result).join(', ')}`);
  
    return result;
  } catch (error) {
    console.error(`getMonthlyDetails error: ${error.message}`);
    console.error(`Stack: ${error.stack}`);
    return { error: error.message, summary: { totalSales: 0, totalSessions: 0, newCount: 0, renewCount: 0 }, salesDetails: [], inquiryStatus: [], salesDetailsByStaff: {}, staffList: [] };
  }
}

/**
 * スタッフ別の成果一覧表を取得
 * @param {string} storeName - 店舗名
 * @param {string} targetMonth - 対象月（例: "2025年8月"）
 * @return {Object} スタッフ別の集計データ
 */
function getStaffPerformanceTable(storeName, targetMonth, sessionId = null) {
  // セッションIDが指定されている場合はセッションを設定
  if (sessionId) {
    try {
      setSession(sessionId);
    } catch (error) {
      console.error('セッション設定エラー:', error);
    }
  }
  
  // 権限チェック
  if (!canAccessStore(storeName)) {
    return { error: 'この店舗へのアクセス権限がありません', total: {}, staffData: {}, staffList: [] };
  }
  try {
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const sheet = ss.getSheetByName('報告データシート履歴');
    if (!sheet) return { error: '報告データシート履歴が見つかりません', total: {}, staffData: {}, staffList: [] };
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // 各列のインデックスを取得
    const getIdx = (name) => headers.indexOf(name);
    const idx = {
      month: getIdx('年月'),
      store: getIdx('店舗'),
      type: getIdx('種別'),
      amount: getIdx('金額'),
      staff: getIdx('スタッフ氏名'),
      system: getIdx('制度選択')
    };

    // 年月を正規化してフィルタリング
    const filteredRows = data.slice(1).filter(row => {
      const rowMonth = normalizeYearMonth(row[idx.month]);
      return String(row[idx.store]) === storeName && rowMonth === targetMonth;
    });

    // スタッフ別に集計
    const staffData = {};
    const allStaffSet = new Set();

    filteredRows.forEach(row => {
      const staff = String(row[idx.staff] || '').trim();
      const type = String(row[idx.type] || '').trim();
      const amount = Number(row[idx.amount]) || 0;
      const system = String(row[idx.system] || '').trim();

      // スタッフ名が空の場合は「未設定」として扱う
      const staffKey = staff || '未設定';
      allStaffSet.add(staffKey);

      // スタッフ別データを初期化
      if (!staffData[staffKey]) {
        staffData[staffKey] = {
          sales: 0,              // 売上
          oldSystemReward: 0,    // 報酬（旧制度）
          newCount: 0,           // 新規
          renewCount: 0,         // 継続
          inquiryCount: 0,        // 問合数
          trialCount: 0,          // 体験
          contractCount: 0,       // 成約
          noContractCount: 0      // 不成約
        };
      }

      // 売上を集計
      staffData[staffKey].sales += amount;

      // 報酬（旧制度）を集計
      if (system.includes('旧制度') || system.includes('旧')) {
        staffData[staffKey].oldSystemReward += amount;
      }

      // 種別ごとにカウント
      if (type.includes('新規')) {
        staffData[staffKey].newCount++;
      }
      if (type.includes('継続')) {
        staffData[staffKey].renewCount++;
      }
      if (type.includes('問合') || type.includes('問い合わせ')) {
        staffData[staffKey].inquiryCount++;
      }
      if (type.includes('体験')) {
        staffData[staffKey].trialCount++;
      }
      if (type.includes('成約')) {
        staffData[staffKey].contractCount++;
      }
      if (type.includes('不成約')) {
        staffData[staffKey].noContractCount++;
      }
    });

    // 合計を計算
    const total = {
      sales: 0,
      oldSystemReward: 0,
      newCount: 0,
      renewCount: 0,
      inquiryCount: 0,
      trialCount: 0,
      contractCount: 0,
      noContractCount: 0
    };

    Object.values(staffData).forEach(data => {
      total.sales += data.sales;
      total.oldSystemReward += data.oldSystemReward;
      total.newCount += data.newCount;
      total.renewCount += data.renewCount;
      total.inquiryCount += data.inquiryCount;
      total.trialCount += data.trialCount;
      total.contractCount += data.contractCount;
      total.noContractCount += data.noContractCount;
    });

    // 成約率を計算（各スタッフと合計）
    const calculateContractRate = (contract, noContract) => {
      const total = contract + noContract;
      if (total === 0) return 0.0;
      return (contract / total * 100).toFixed(1);
    };

    // スタッフ別データに成約率を追加
    const staffDataWithRate = {};
    Object.keys(staffData).forEach(staffKey => {
      const data = staffData[staffKey];
      staffDataWithRate[staffKey] = {
        ...data,
        contractRate: parseFloat(calculateContractRate(data.contractCount, data.noContractCount))
      };
    });

    // 合計の成約率を計算
    const totalContractRate = parseFloat(calculateContractRate(total.contractCount, total.noContractCount));

    // スタッフリストをソート（オーナーを先に、その後スタッフ）
    const staffList = Array.from(allStaffSet).sort((a, b) => {
      // 「オーナー」や「古林」などのキーワードで判定（必要に応じて調整）
      if (a.includes('オーナー') || a.includes('古林')) return -1;
      if (b.includes('オーナー') || b.includes('古林')) return 1;
      return a.localeCompare(b);
    });

    return {
      total: {
        ...total,
        contractRate: totalContractRate
      },
      staffData: staffDataWithRate,
      staffList: staffList
    };
  } catch (error) {
    console.error(`getStaffPerformanceTable error: ${error.message}`);
    console.error(`Stack: ${error.stack}`);
    return {
      total: {
        sales: 0,
        oldSystemReward: 0,
        newCount: 0,
        renewCount: 0,
        inquiryCount: 0,
        trialCount: 0,
        contractCount: 0,
        noContractCount: 0,
        contractRate: 0.0
      },
      staffData: {},
      staffList: []
    };
  }
}

/**
 * 店舗の直近3ヶ月の売上データを取得
 * @param {string} storeName - 店舗名
 * @return {Object} 直近3ヶ月の売上データ
 */
function getStoreLast3MonthsTrend(storeName, sessionId = null) {
  // セッションIDが指定されている場合はセッションを設定
  if (sessionId) {
    try {
      setSession(sessionId);
    } catch (error) {
      console.error('セッション設定エラー:', error);
    }
  }
  
  // 権限チェック
  if (!canAccessStore(storeName)) {
    return { months: [], values: [] };
  }
  try {
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const sheet = ss.getSheetByName('報告データシート履歴');
    if (!sheet) return { months: [], values: [] };
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const getIdx = (name) => headers.indexOf(name);
    const idx = {
      month: getIdx('年月'),
      store: getIdx('店舗'),
      amount: getIdx('金額')
    };
    
    // 現在の年月から直近3ヶ月を計算
    const now = new Date();
    const months = [];
    for (let i = 2; i >= 0; i--) {
      const date = new Date(now.getFullYear(), now.getMonth() - i, 1);
      const year = date.getFullYear();
      const month = date.getMonth() + 1;
      months.push(`${year}年${month}月`);
    }
    
    const trendData = {};
    months.forEach(month => {
      trendData[month] = 0;
    });
    
    data.slice(1).forEach(row => {
      const rowStore = String(row[idx.store] || '').trim();
      const rowMonth = normalizeYearMonth(row[idx.month]);
      const amount = Number(row[idx.amount]) || 0;
      
      if (rowStore === storeName && months.includes(rowMonth)) {
        trendData[rowMonth] = (trendData[rowMonth] || 0) + amount;
      }
    });
    
    return {
      months: months,
      values: months.map(m => trendData[m] || 0)
    };
  } catch (error) {
    console.error(`getStoreLast3MonthsTrend error: ${error.message}`);
    const now = new Date();
    const months = [];
    for (let i = 2; i >= 0; i--) {
      const date = new Date(now.getFullYear(), now.getMonth() - i, 1);
      const year = date.getFullYear();
      const month = date.getMonth() + 1;
      months.push(`${year}年${month}月`);
    }
    return {
      months: months,
      values: [0, 0, 0]
    };
  }
}

/**
 * 全店舗の合計と各店舗のサマリを取得（経営者向け）
 * @param {string} sessionId - セッションID（オプション）
 * @return {Object} 全店舗の集計データと各店舗のサマリ
 */
function getAllStoresSummary(sessionId = null) {
  try {
    console.log('getAllStoresSummary呼び出し: sessionId =', sessionId);
    
    // セッションIDが指定されている場合はセッションを設定
    if (sessionId) {
      try {
        const setResult = setSession(sessionId);
        if (!setResult || !setResult.success) {
          console.warn('セッション設定失敗:', setResult);
        }
      } catch (error) {
        console.error('セッション設定エラー:', error);
      }
    }
    
    // 権限チェック
    const user = getCurrentUser();
    console.log('getCurrentUser結果:', user ? 'ユーザー取得成功' : 'ユーザー取得失敗');
    
    if (!user) {
      console.error('ユーザーが取得できませんでした');
      return { 
        error: 'ログインが必要です',
        currentMonth: '',
        totalCurrentMonth: { sales: 0, customerCount: 0, paidCustomerCount: 0, newCount: 0, renewCount: 0, storeCount: 0 },
        totalAll: { sales: 0, customerCount: 0 },
        storeSummaries: [],
        storeCount: 0
      };
    }
    
    const accessibleStores = getAccessibleStores();
    console.log('アクセス可能な店舗:', accessibleStores);
    
    let ss;
    try {
      ss = SpreadsheetApp.openById(getSpreadsheetId());
    } catch (ssError) {
      console.error('スプレッドシート接続エラー:', ssError);
      return { 
        error: 'スプレッドシートに接続できません',
        currentMonth: '',
        totalCurrentMonth: { sales: 0, customerCount: 0, paidCustomerCount: 0, newCount: 0, renewCount: 0, storeCount: 0 },
        totalAll: { sales: 0, customerCount: 0 },
        storeSummaries: [],
        storeCount: 0
      };
    }
    
    // 報告データシート履歴から売上データを取得（存在しない場合はスキップ）
    const reportSheet = ss.getSheetByName('報告データシート履歴');
    let reportData = [];
    let headers = [];
    if (reportSheet) {
      const reportLastRow = reportSheet.getLastRow();
      if (reportLastRow > 0) {
        reportData = reportSheet.getDataRange().getValues();
        headers = reportData[0] || [];
      }
    } else {
      console.log('報告データシート履歴が見つかりません（顧客登録データのみで集計）');
    }
    
    const getIdx = (name) => headers.indexOf(name);
    const idx = {
      month: getIdx('年月'),
      store: getIdx('店舗'),
      type: getIdx('種別'),
      amount: getIdx('金額'),
      name: getIdx('お名前')
    };
    
    // 現在の年月を取得
    const now = new Date();
    const currentYear = now.getFullYear();
    const currentMonth = now.getMonth() + 1;
    const currentMonthStr = `${currentYear}年${currentMonth}月`;
    
    // 当月の開始日と終了日を取得（顧客登録シート用）
    const startOfMonth = new Date(currentYear, currentMonth - 1, 1);
    const endOfMonth = new Date(currentYear, currentMonth, 0, 23, 59, 59, 999);
    
    // 顧客登録シートから当月の顧客数を取得
    const customerSheet = ss.getSheetByName('顧客登録');
    const customerData = customerSheet ? customerSheet.getDataRange().getValues() : [];
    
    // 店舗別の顧客数を集計（顧客登録シートから）
    const customerCountByStore = {};
    const paidCustomerCountByStore = {}; // 決済完了の顧客数
    const salesByStore = {}; // 決済完了の顧客の売上金額
    const allStores = new Set();
    
    if (customerData && customerData.length > 1) {
      // ヘッダー行をスキップしてデータを取得
      // A:ID, B:店舗, C:担当者, D:体験/入会, E:日付, F:名前, ..., U:支払い状況, W:売上, ..., Y:登録日時
      customerData.slice(1).forEach((row) => {
        const store = String(row[1] || '').trim(); // B列: 店舗
        const name = String(row[5] || '').trim(); // F列: 名前
        const registeredAtStr = String(row[24] || '').trim(); // Y列: 登録日時
        const paymentStatus = String(row[20] || '').trim(); // U列: 支払い状況
        const salesStr = String(row[22] || '').trim(); // W列: 売上
        
        if (!name || !store) return; // 名前または店舗がない行はスキップ
        
        // アクセス権限チェック
        if (accessibleStores.length > 0 && !accessibleStores.includes(store)) {
          return;
        }
        
        allStores.add(store);
        
        // 登録日時が当月かどうかをチェック（最適化版）
        let isCurrentMonth = false;
        if (registeredAtStr) {
          try {
            let registeredDate = null;
            // Dateオブジェクトの場合は直接使用
            if (row[24] instanceof Date) {
              registeredDate = row[24];
            } else if (typeof registeredAtStr === 'string') {
              // 文字列の場合は簡易解析
              const dateStr = registeredAtStr.split(' ')[0]; // 日付部分のみ取得
              if (dateStr.includes('-')) {
                const parts = dateStr.split('-');
                if (parts.length === 3) {
                  registeredDate = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
                }
              } else if (dateStr.includes('/')) {
                const parts = dateStr.split('/');
                if (parts.length === 3) {
                  registeredDate = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
                }
              }
            }
            
            // 当月の顧客のみをカウント
            if (registeredDate && !isNaN(registeredDate.getTime())) {
              isCurrentMonth = registeredDate >= startOfMonth && registeredDate <= endOfMonth;
            }
          } catch (dateError) {
            // エラー時はスキップ
            isCurrentMonth = false;
          }
        }
        
        if (isCurrentMonth) {
              // 登録数
              if (!customerCountByStore[store]) {
                customerCountByStore[store] = 0;
              }
              customerCountByStore[store]++;
              
              // 決済完了数と売上（支払い状況が「済」の場合）
              if (paymentStatus === '済') {
                if (!paidCustomerCountByStore[store]) {
                  paidCustomerCountByStore[store] = 0;
                }
                paidCustomerCountByStore[store]++;
                
                // 売上金額を取得（数値に変換）
                let salesAmount = 0;
                if (salesStr) {
                  // 数値以外の文字を削除（カンマ、円記号など）
                  const cleanedSales = salesStr.replace(/[^0-9.-]/g, '');
                  salesAmount = parseFloat(cleanedSales) || 0;
                }
                
                if (!salesByStore[store]) {
                  salesByStore[store] = 0;
                }
                salesByStore[store] += salesAmount;
              }
        }
      });
    }
    
    // 店舗別の集計（顧客登録シートから売上と顧客数を取得）
    const storeSummary = {};
    
    // 全期間のデータも集計（累計）- 報告データシート履歴から
    const storeSummaryAll = {};
    
    // 報告データシート履歴から全期間の累計を取得
    reportData.slice(1).forEach(row => {
      const store = String(row[idx.store] || '').trim();
      if (!store || store === '' || store === 'ー' || store === '-') {
        return;
      }
      
      // アクセス権限チェック
      if (accessibleStores.length > 0 && !accessibleStores.includes(store)) {
        return;
      }
      
      allStores.add(store);
      const amount = Number(row[idx.amount]) || 0;
      
      // 全期間の集計
      if (!storeSummaryAll[store]) {
        storeSummaryAll[store] = {
          sales: 0,
          customerCount: 0
        };
      }
      storeSummaryAll[store].sales += amount;
      if (amount > 0) {
        storeSummaryAll[store].customerCount++;
      }
    });
    
    // 顧客登録シートから取得したデータで店舗別サマリを構築
    Object.keys(customerCountByStore).forEach(store => {
      storeSummary[store] = {
        sales: salesByStore[store] || 0, // 決済完了の顧客の売上金額
        customerCount: customerCountByStore[store] || 0,
        newCount: 0, // 新規・継続の集計は報告データから取得する場合は追加可能
        renewCount: 0
      };
    });
    
    // 売上データがあるが顧客登録にない店舗も追加（売上は0、顧客数も0）
    allStores.forEach(store => {
      if (!storeSummary[store]) {
        storeSummary[store] = {
          sales: 0,
          customerCount: 0,
          newCount: 0,
          renewCount: 0
        };
      }
    });
    
    // 全店舗の合計を計算
    const totalCurrentMonth = {
      sales: 0,
      customerCount: 0,
      paidCustomerCount: 0, // 決済完了の顧客数
      newCount: 0,
      renewCount: 0,
      storeCount: allStores.size
    };
    
    const totalAll = {
      sales: 0,
      customerCount: 0
    };
    
    Object.values(storeSummary).forEach(summary => {
      totalCurrentMonth.sales += summary.sales;
      totalCurrentMonth.customerCount += summary.customerCount;
      totalCurrentMonth.newCount += summary.newCount;
      totalCurrentMonth.renewCount += summary.renewCount;
    });
    
    // 決済完了の顧客数の合計を計算
    Object.values(paidCustomerCountByStore).forEach(count => {
      totalCurrentMonth.paidCustomerCount += count;
    });
    
    Object.values(storeSummaryAll).forEach(summary => {
      totalAll.sales += summary.sales;
      totalAll.customerCount += summary.customerCount;
    });
    
    // 店舗リストを表示順でソート（上段: 駒沢・高円寺・江古田、下段: 曙橋・西荻窪・幡ヶ谷）
    const DASHBOARD_STORE_ORDER = ['駒沢', '高円寺', '江古田', '曙橋', '西荻窪', '幡ヶ谷'];
    const storeList = Array.from(allStores).sort((a, b) => {
      const ia = DASHBOARD_STORE_ORDER.indexOf(a);
      const ib = DASHBOARD_STORE_ORDER.indexOf(b);
      if (ia >= 0 && ib >= 0) return ia - ib;
      if (ia >= 0) return -1;
      if (ib >= 0) return 1;
      return a.localeCompare(b);
    });

    // 各店舗のサマリを配列に変換
    const storeSummaries = storeList.map(store => ({
      storeName: store,
      currentMonth: storeSummary[store] || {
        sales: 0,
        customerCount: 0,
        newCount: 0,
        renewCount: 0
      },
      allTime: storeSummaryAll[store] || {
        sales: 0,
        customerCount: 0
      }
    }));
    
    const result = {
      currentMonth: currentMonthStr,
      totalCurrentMonth: totalCurrentMonth,
      totalAll: totalAll,
      storeSummaries: storeSummaries,
      storeCount: allStores.size
    };
    
    console.log('getAllStoresSummary成功:', {
      currentMonth: result.currentMonth,
      storeCount: result.storeCount,
      summariesCount: result.storeSummaries.length
    });
    
    return result;
  } catch (error) {
    console.error(`getAllStoresSummary error: ${error.message}`);
    console.error(`Stack: ${error.stack}`);
    return {
      error: 'データの取得中にエラーが発生しました: ' + error.message,
      currentMonth: '',
      totalCurrentMonth: {
        sales: 0,
        customerCount: 0,
        paidCustomerCount: 0,
        newCount: 0,
        renewCount: 0,
        storeCount: 0
      },
      totalAll: {
        sales: 0,
        customerCount: 0
      },
      storeSummaries: [],
      storeCount: 0
    };
  }
}

/**
 * 全スタッフの担当顧客数・売上・報酬を取得（全期間または指定期間）
 * @param {string} storeName - 店舗名（空の場合は全店舗）
 * @param {string} targetMonth - 対象月（空の場合は全期間、例: "2025年8月"）
 * @param {string} sessionId - セッションID（オプション）
 * @return {Object} スタッフ別の集計データ
 */
function getAllStaffSummary(storeName = '', targetMonth = '', sessionId = null) {
  // セッションIDが指定されている場合はセッションを設定
  if (sessionId) {
    try {
      setSession(sessionId);
    } catch (error) {
      console.error('セッション設定エラー:', error);
    }
  }
  
  // 権限チェック
  const user = getCurrentUser();
  if (!user) {
    return { error: 'ログインが必要です', total: {}, staffData: {}, staffList: [] };
  }
  
  const accessibleStores = getAccessibleStores();
  
  // 店舗が指定されている場合、アクセス権限をチェック
  if (storeName && !canAccessStore(storeName)) {
    return { error: 'この店舗へのアクセス権限がありません', total: {}, staffData: {}, staffList: [] };
  }
  
  // 店舗が指定されていない場合、アクセス可能な店舗のみを対象にする
  if (!storeName && accessibleStores.length > 0) {
    // 全店舗ではなく、アクセス可能な店舗のみを対象にする
    // この場合は、storeNameを空のままにして、後でフィルタリング
  }
  try {
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const sheet = ss.getSheetByName('報告データシート履歴');
    if (!sheet) return { error: "報告データシート履歴が見つかりません", total: {}, staffData: {}, staffList: [] };
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const getIdx = (name) => headers.indexOf(name);
    const idx = {
      month: getIdx('年月'),
      store: getIdx('店舗'),
      type: getIdx('種別'),
      amount: getIdx('金額'),
      staff: getIdx('スタッフ氏名'),
      system: getIdx('制度選択'),
      name: getIdx('お名前')
    };
    
    // フィルタリング
    let filteredRows = data.slice(1);
    
    if (storeName) {
      filteredRows = filteredRows.filter(row => {
        return String(row[idx.store] || '').trim() === storeName;
      });
    } else if (accessibleStores.length > 0) {
      // アクセス可能な店舗のみをフィルタリング
      filteredRows = filteredRows.filter(row => {
        const rowStore = String(row[idx.store] || '').trim();
        return accessibleStores.includes(rowStore);
      });
    }
    
    if (targetMonth) {
      filteredRows = filteredRows.filter(row => {
        const rowMonth = normalizeYearMonth(row[idx.month]);
        return rowMonth === targetMonth;
      });
    }
    
    // スタッフ別に集計
    const staffData = {};
    const allStaffSet = new Set();
    const customerSetByStaff = {}; // スタッフ別の顧客セット（重複排除用）
    
    filteredRows.forEach(row => {
      const staff = String(row[idx.staff] || '').trim();
      const staffKey = staff || '未設定';
      const type = String(row[idx.type] || '').trim();
      const amount = Number(row[idx.amount]) || 0;
      const system = String(row[idx.system] || '').trim();
      const customerName = String(row[idx.name] || '').trim();
      
      allStaffSet.add(staffKey);
      
      // スタッフ別データを初期化
      if (!staffData[staffKey]) {
        staffData[staffKey] = {
          sales: 0,              // 売上
          oldSystemReward: 0,    // 報酬（旧制度）
          newSystemReward: 0,    // 報酬（新制度）- 計算が必要な場合は追加
          customerCount: 0,      // 担当顧客数
          newCount: 0,           // 新規
          renewCount: 0,         // 継続
          inquiryCount: 0,       // 問合数
          trialCount: 0,         // 体験
          contractCount: 0,      // 成約
          noContractCount: 0     // 不成約
        };
        customerSetByStaff[staffKey] = new Set();
      }
      
      // 売上を集計
      staffData[staffKey].sales += amount;
      
      // 報酬（旧制度）を集計
      if (system.includes('旧制度') || system.includes('旧')) {
        staffData[staffKey].oldSystemReward += amount;
      } else {
        // 新制度の報酬計算（必要に応じて調整）
        // ここでは売上の一部として計算する例
        staffData[staffKey].newSystemReward += amount;
      }
      
      // 顧客数をカウント（重複排除）
      if (customerName && customerName !== '-' && customerName !== '') {
        customerSetByStaff[staffKey].add(customerName);
      }
      
      // 種別ごとにカウント
      if (type.includes('新規')) {
        staffData[staffKey].newCount++;
      }
      if (type.includes('継続')) {
        staffData[staffKey].renewCount++;
      }
      if (type.includes('問合') || type.includes('問い合わせ')) {
        staffData[staffKey].inquiryCount++;
      }
      if (type.includes('体験')) {
        staffData[staffKey].trialCount++;
      }
      if (type.includes('成約')) {
        staffData[staffKey].contractCount++;
      }
      if (type.includes('不成約')) {
        staffData[staffKey].noContractCount++;
      }
    });
    
    // 顧客数をセット
    Object.keys(staffData).forEach(staffKey => {
      staffData[staffKey].customerCount = customerSetByStaff[staffKey] ? customerSetByStaff[staffKey].size : 0;
    });
    
    // 合計を計算
    const total = {
      sales: 0,
      oldSystemReward: 0,
      newSystemReward: 0,
      customerCount: 0,
      newCount: 0,
      renewCount: 0,
      inquiryCount: 0,
      trialCount: 0,
      contractCount: 0,
      noContractCount: 0
    };
    
    Object.values(staffData).forEach(data => {
      total.sales += data.sales;
      total.oldSystemReward += data.oldSystemReward;
      total.newSystemReward += data.newSystemReward;
      total.customerCount += data.customerCount;
      total.newCount += data.newCount;
      total.renewCount += data.renewCount;
      total.inquiryCount += data.inquiryCount;
      total.trialCount += data.trialCount;
      total.contractCount += data.contractCount;
      total.noContractCount += data.noContractCount;
    });
    
    // 成約率を計算
    const calculateContractRate = (contract, noContract) => {
      const sum = contract + noContract;
      if (sum === 0) return 0.0;
      return (contract / sum * 100).toFixed(1);
    };
    
    // スタッフ別データに成約率を追加
    const staffDataWithRate = {};
    Object.keys(staffData).forEach(staffKey => {
      const data = staffData[staffKey];
      staffDataWithRate[staffKey] = {
        ...data,
        contractRate: parseFloat(calculateContractRate(data.contractCount, data.noContractCount)),
        totalReward: data.oldSystemReward + data.newSystemReward // 総報酬
      };
    });
    
    // 合計の成約率を計算
    const totalContractRate = parseFloat(calculateContractRate(total.contractCount, total.noContractCount));
    
    // スタッフリストをソート（オーナーを先に、その後スタッフ）
    const staffList = Array.from(allStaffSet).sort((a, b) => {
      if (a.includes('オーナー') || a.includes('古林')) return -1;
      if (b.includes('オーナー') || b.includes('古林')) return 1;
      return a.localeCompare(b);
    });
    
    return {
      total: {
        ...total,
        contractRate: totalContractRate,
        totalReward: total.oldSystemReward + total.newSystemReward
      },
      staffData: staffDataWithRate,
      staffList: staffList,
      filter: {
        storeName: storeName || '全店舗',
        targetMonth: targetMonth || '全期間'
      }
    };
  } catch (error) {
    console.error(`getAllStaffSummary error: ${error.message}`);
    console.error(`Stack: ${error.stack}`);
    return {
      total: {
        sales: 0,
        oldSystemReward: 0,
        newSystemReward: 0,
        customerCount: 0,
        newCount: 0,
        renewCount: 0,
        inquiryCount: 0,
        trialCount: 0,
        contractCount: 0,
        noContractCount: 0,
        contractRate: 0.0,
        totalReward: 0
      },
      staffData: {},
      staffList: [],
      filter: {
        storeName: storeName || '全店舗',
        targetMonth: targetMonth || '全期間'
      }
    };
  }
}

/**
 * スタッフ別の詳細実績報告を取得
 * @param {string} staffName - スタッフ名
 * @param {string} storeName - 店舗名（空の場合は全店舗）
 * @param {string} targetMonth - 対象月（空の場合は全期間、例: "2025年8月"）
 * @param {string} sessionId - セッションID（オプション）
 * @return {Object} スタッフの詳細実績データ
 */
function getStaffDetailReport(staffName, storeName = '', targetMonth = '', sessionId = null) {
  try {
    console.log('getStaffDetailReport呼び出し:', { staffName, storeName, targetMonth, sessionId });
    
    // セッションIDが指定されている場合はセッションを設定
    if (sessionId) {
      try {
        setSession(sessionId);
      } catch (error) {
        console.error('セッション設定エラー:', error);
      }
    }
    
    // 権限チェック
    const user = getCurrentUser();
    if (!user) {
      console.error('ユーザーが取得できませんでした');
      return { error: 'ログインが必要です', staffName: staffName, records: [], summary: {} };
    }
    
    const accessibleStores = getAccessibleStores();
    console.log('アクセス可能な店舗:', accessibleStores);
    
    // 店舗が指定されている場合、アクセス権限をチェック
    if (storeName && !canAccessStore(storeName)) {
      console.error('店舗アクセス権限なし:', storeName);
      return { error: 'この店舗へのアクセス権限がありません', staffName: staffName, records: [], summary: {} };
    }
    
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const sheet = ss.getSheetByName('報告データシート履歴');
    
    if (!sheet) {
      console.error('報告データシート履歴が見つかりません');
      return { error: '報告データシート履歴が見つかりません', staffName: staffName, records: [], summary: {} };
    }
    
    const data = sheet.getDataRange().getValues();
    
    if (!data || data.length <= 1) {
      console.warn('データがありません');
      return {
        staffName: staffName,
        storeName: storeName || '全店舗',
        targetMonth: targetMonth || '全期間',
        records: [],
        summary: {
          totalSales: 0,
          oldSystemReward: 0,
          newSystemReward: 0,
          totalReward: 0,
          newCount: 0,
          renewCount: 0,
          inquiryCount: 0,
          trialCount: 0,
          contractCount: 0,
          noContractCount: 0,
          customerCount: 0,
          contractRate: '0.0'
        }
      };
    }
    
    const headers = data[0];
    
    const getIdx = (name) => headers.indexOf(name);
    const idx = {
      month: getIdx('年月'),
      store: getIdx('店舗'),
      type: getIdx('種別'),
      amount: getIdx('金額'),
      staff: getIdx('スタッフ氏名'),
      system: getIdx('制度選択'),
      name: getIdx('お名前')
    };
    
    // フィルタリング
    let filteredRows = data.slice(1);
    console.log('フィルタ前の行数:', filteredRows.length);
    
    // スタッフ名でフィルタリング
    if (staffName) {
      const beforeCount = filteredRows.length;
      filteredRows = filteredRows.filter(row => {
        const rowStaff = String(row[idx.staff] || '').trim();
        return rowStaff === staffName;
      });
      console.log('スタッフ名フィルタ後:', filteredRows.length, '件（フィルタ前:', beforeCount, '件）');
    }
    
    // 店舗でフィルタリング
    if (storeName) {
      const beforeCount = filteredRows.length;
      filteredRows = filteredRows.filter(row => {
        const rowStore = String(row[idx.store] || '').trim();
        return rowStore === storeName;
      });
      console.log('店舗フィルタ後:', filteredRows.length, '件（フィルタ前:', beforeCount, '件）');
    } else if (accessibleStores.length > 0) {
      // アクセス可能な店舗のみ
      const beforeCount = filteredRows.length;
      filteredRows = filteredRows.filter(row => {
        const rowStore = String(row[idx.store] || '').trim();
        return accessibleStores.includes(rowStore);
      });
      console.log('アクセス可能店舗フィルタ後:', filteredRows.length, '件（フィルタ前:', beforeCount, '件）');
    }
    
    // 月でフィルタリング
    if (targetMonth) {
      const beforeCount = filteredRows.length;
      filteredRows = filteredRows.filter(row => {
        const rowMonth = normalizeYearMonth(row[idx.month]);
        return rowMonth === targetMonth;
      });
      console.log('月フィルタ後:', filteredRows.length, '件（フィルタ前:', beforeCount, '件）, 対象月:', targetMonth);
    }
    
    console.log('最終フィルタ後の行数:', filteredRows.length);
    
    // 詳細レコードを構築
    const records = filteredRows.map(row => {
      const month = normalizeYearMonth(row[idx.month]);
      const store = String(row[idx.store] || '').trim();
      const type = String(row[idx.type] || '').trim();
      const amount = Number(row[idx.amount]) || 0;
      const system = String(row[idx.system] || '').trim();
      const customerName = String(row[idx.name] || '').trim();
      
      return {
        month: month,
        store: store,
        type: type,
        amount: amount,
        system: system,
        customerName: customerName
      };
    });
    
    // サマリを計算
    const summary = {
      totalSales: 0,
      oldSystemReward: 0,
      newSystemReward: 0,
      totalReward: 0,
      newCount: 0,
      renewCount: 0,
      inquiryCount: 0,
      trialCount: 0,
      contractCount: 0,
      noContractCount: 0,
      customerCount: 0
    };
    
    const customerSet = new Set();
    
    records.forEach(record => {
      summary.totalSales += record.amount;
      
      if (record.system.includes('旧制度') || record.system.includes('旧')) {
        summary.oldSystemReward += record.amount;
      } else {
        summary.newSystemReward += record.amount;
      }
      
      if (record.type.includes('新規')) {
        summary.newCount++;
      }
      if (record.type.includes('継続')) {
        summary.renewCount++;
      }
      if (record.type.includes('問合') || record.type.includes('問い合わせ')) {
        summary.inquiryCount++;
      }
      if (record.type.includes('体験')) {
        summary.trialCount++;
      }
      if (record.type.includes('成約')) {
        summary.contractCount++;
      }
      if (record.type.includes('不成約')) {
        summary.noContractCount++;
      }
      
      if (record.customerName && record.customerName !== '-' && record.customerName !== '') {
        customerSet.add(record.customerName);
      }
    });
    
    summary.totalReward = summary.oldSystemReward + summary.newSystemReward;
    summary.customerCount = customerSet.size;
    
    // 成約率を計算
    const totalCases = summary.contractCount + summary.noContractCount;
    summary.contractRate = totalCases > 0 ? (summary.contractCount / totalCases * 100).toFixed(1) : '0.0';
    
    console.log('getStaffDetailReport成功:', {
      staffName: staffName,
      recordCount: records.length,
      totalSales: summary.totalSales,
      customerCount: summary.customerCount
    });
    
    return {
      staffName: staffName,
      storeName: storeName || '全店舗',
      targetMonth: targetMonth || '全期間',
      records: records,
      summary: summary
    };
  } catch (error) {
    console.error(`getStaffDetailReport error: ${error.message}`);
    console.error(`Stack: ${error.stack}`);
    return {
      error: 'データの取得中にエラーが発生しました: ' + error.message,
      staffName: staffName,
      records: [],
      summary: {}
    };
  }
}

/**
 * ユーザー管理シートを手動で作成する関数
 * GASエディタから実行して、ユーザー管理シートとデフォルトアカウントを作成できます
 * @return {string} 作成結果メッセージ
 */
function createUserSheetManually() {
  try {
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    let userSheet = ss.getSheetByName(USER_SHEET_NAME);
    
    if (userSheet) {
      return '「ユーザー管理」シートは既に存在します。';
    }
    
    // シートを作成
    userSheet = ss.insertSheet(USER_SHEET_NAME);
    const headers = ['メールアドレス', 'パスワード（ハッシュ）', '権限レベル', '店舗名', 'スタッフ名', '所属店舗', '稼働店舗'];
    userSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    userSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    userSheet.setFrozenRows(1);
    
    // デフォルトの経営オーナーアカウントを作成
    try {
      const defaultPasswordHash = hashPassword('admin123');
      userSheet.appendRow([
        'admin@trainers.com',
        defaultPasswordHash,
        '経営オーナー',
        '',
        '管理者',
        '',
        ''
      ]);
      
      return '「ユーザー管理」シートを作成しました。\nデフォルトアカウント:\nメール: admin@trainers.com\nパスワード: admin123';
    } catch (error) {
      console.error('デフォルトユーザー作成エラー:', error);
      return '「ユーザー管理」シートを作成しましたが、デフォルトユーザーの作成に失敗しました。\nエラー: ' + error.toString();
    }
  } catch (error) {
    console.error('ユーザー管理シート作成エラー:', error);
    return 'シートの作成に失敗しました: ' + error.toString();
  }
}

/**
 * ダミーのスタッフアカウントを追加する関数
 * GASエディタから実行して、テスト用のダミーアカウントを追加できます
 * @return {string} 作成結果メッセージ
 */
function addDummyStaffAccounts() {
  try {
    const userSheet = getUserSheet();
    
    // ダミースタッフ情報
    // パスワードは全て「password123」
    const dummyStaffs = [
      // 経営オーナー
      {
        email: 'owner@trainers.com',
        password: 'password123',
        role: '経営オーナー',
        storeName: '',
        staffName: '経営太郎',
        belongStores: '',
        workStores: ''
      },
      // FCオーナー（曙橋店）
      {
        email: 'fc-akebonobashi@trainers.com',
        password: 'password123',
        role: 'FCオーナー',
        storeName: '曙橋',
        staffName: '曙橋オーナー',
        belongStores: '曙橋',
        workStores: '曙橋'
      },
      // FCオーナー（高田馬場店）
      {
        email: 'fc-takadanobaba@trainers.com',
        password: 'password123',
        role: 'FCオーナー',
        storeName: '高田馬場',
        staffName: '高田馬場オーナー',
        belongStores: '高田馬場',
        workStores: '高田馬場'
      },
      // スタッフ（曙橋店所属）
      {
        email: 'staff1@trainers.com',
        password: 'password123',
        role: 'スタッフ',
        storeName: '曙橋',
        staffName: '田中トレーナー',
        belongStores: '曙橋',
        workStores: '曙橋,高田馬場'
      },
      // スタッフ（高田馬場店所属）
      {
        email: 'staff2@trainers.com',
        password: 'password123',
        role: 'スタッフ',
        storeName: '高田馬場',
        staffName: '鈴木トレーナー',
        belongStores: '高田馬場',
        workStores: '高田馬場'
      },
      // スタッフ（複数店舗勤務）
      {
        email: 'staff3@trainers.com',
        password: 'password123',
        role: 'スタッフ',
        storeName: '曙橋',
        staffName: '佐藤トレーナー',
        belongStores: '曙橋',
        workStores: '曙橋,高田馬場,新宿'
      }
    ];
    
    let addedCount = 0;
    const existingData = userSheet.getDataRange().getValues();
    const existingEmails = existingData.slice(1).map(row => String(row[0]).toLowerCase().trim());
    
    const results = [];
    
    for (const staff of dummyStaffs) {
      // 既に存在するメールアドレスはスキップ
      if (existingEmails.includes(staff.email.toLowerCase())) {
        results.push(`${staff.email}: 既に存在します（スキップ）`);
        continue;
      }
      
      try {
        const hashedPassword = hashPassword(staff.password);
        userSheet.appendRow([
          staff.email,
          hashedPassword,
          staff.role,
          staff.storeName,
          staff.staffName,
          staff.belongStores,
          staff.workStores
        ]);
        addedCount++;
        results.push(`${staff.email}: 追加完了（${staff.role}）`);
      } catch (error) {
        results.push(`${staff.email}: 追加失敗 - ${error.toString()}`);
      }
    }
    
    return `ダミースタッフ追加完了\n追加数: ${addedCount}件\n\n詳細:\n${results.join('\n')}\n\n【ログイン情報】\n全アカウント共通パスワード: password123`;
  } catch (error) {
    console.error('ダミースタッフ追加エラー:', error);
    return 'ダミースタッフの追加に失敗しました: ' + error.toString();
  }
}

/**
 * ユーザー管理シートを取得（存在しない場合は作成）
 * @return {Sheet} ユーザー管理シート
 */
function getUserSheet() {
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  let userSheet = ss.getSheetByName(USER_SHEET_NAME);
  
  if (!userSheet) {
    userSheet = ss.insertSheet(USER_SHEET_NAME);
    const headers = ['メールアドレス', 'パスワード（ハッシュ）', '権限レベル', '店舗名', 'スタッフ名', '所属店舗', '稼働店舗'];
    userSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    userSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    userSheet.setFrozenRows(1);
    
    // デフォルトの経営オーナーアカウントを作成（初期設定用）
    // パスワードは後で変更してください
    try {
      const defaultPasswordHash = hashPassword('admin123');
      userSheet.appendRow([
        'admin@trainers.com',
        defaultPasswordHash,
        '経営オーナー',
        '',
        '管理者',
        '',
        ''
      ]);
    } catch (error) {
      console.error('デフォルトユーザー作成エラー:', error);
      // エラーが発生してもシートは作成されているので続行
    }
  }
  
  return userSheet;
}

/**
 * パスワードをハッシュ化
 * @param {string} password - パスワード
 * @return {string} ハッシュ化されたパスワード
 */
function hashPassword(password) {
  if (!password || password === null || password === undefined || password === '') {
    throw new Error('パスワードが空です');
  }
  
  try {
    const hash = Utilities.computeDigest(
      Utilities.DigestAlgorithm.MD5,
      String(password),
      Utilities.Charset.UTF_8
    );
    return Utilities.base64Encode(hash);
  } catch (error) {
    console.error('パスワードハッシュ化エラー:', error);
    throw new Error('パスワードのハッシュ化に失敗しました: ' + error.toString());
  }
}

/**
 * ログイン処理（スタッフマスタシート参照）
 * スタッフマスタの列構成:
 *   A列: 店舗名, B列: ログインID, C列: パスワード,
 *   D列: 権限（"管理者"=経営オーナー, "オーナー"=FCオーナー, 空=スタッフ）,
 *   E列: 名前
 * 同じ名前が複数行にある場合、複数店舗へのアクセスを許可
 * @param {string} loginId - ログインID（スタッフID）
 * @param {string} password - パスワード
 * @return {Object} ログイン結果
 */
function login(loginId, password) {
  try {
    // 入力値の検証
    if (!loginId || loginId === null || loginId === undefined || String(loginId).trim() === '') {
      return {
        success: false,
        message: 'ログインIDを入力してください'
      };
    }
    
    if (!password || password === null || password === undefined || password === '') {
      return {
        success: false,
        message: 'パスワードを入力してください'
      };
    }
    
    // スタッフマスタシートを取得
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const sheet = ss.getSheetByName(STAFF_MASTER_SHEET_NAME);
    if (!sheet) {
      return {
        success: false,
        message: 'スタッフマスタシートが見つかりません'
      };
    }
    
    const data = sheet.getDataRange().getValues();
    if (!data || data.length <= 1) {
      return {
        success: false,
        message: 'スタッフマスタにデータがありません'
      };
    }
    
    const headers = data[0];
    
    // ヘッダーから列インデックスを動的に取得
    const findIdx = (candidates, fallback) => {
      for (const c of candidates) {
        const idx = headers.indexOf(c);
        if (idx >= 0) return idx;
      }
      return fallback;
    };
    
    const storeIdx = findIdx(['店舗名', '店舗', '所属店舗'], 0);
    const loginIdIdx = findIdx(['ログインID', 'スタッフID', 'ID', 'ログイン'], 1);
    const passwordIdx = findIdx(['パスワード', 'PW', 'PASS', 'pass'], 2);
    const roleIdx = findIdx(['権限', 'アクセス権', 'ロール', '役割'], 3);
    const nameIdx = findIdx(['名前', '氏名', 'スタッフ名'], 4);
    
    console.log('スタッフマスタ列マッピング - 店舗:', storeIdx, 'ログインID:', loginIdIdx, 'パスワード:', passwordIdx, '権限:', roleIdx, '名前:', nameIdx);
    
    // ログインIDとパスワードで認証
    const normalizedLoginId = String(loginId).trim();
    const normalizedPassword = String(password).trim();
    
    let matchedRow = null;
    let matchedRowIndex = -1;
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowLoginId = String(row[loginIdIdx] || '').trim();
      const rowPassword = String(row[passwordIdx] || '').trim();
      
      if (rowLoginId === normalizedLoginId && rowPassword === normalizedPassword) {
        matchedRow = row;
        matchedRowIndex = i;
        break;
      }
    }
    
    if (!matchedRow) {
      return {
        success: false,
        message: 'ログインIDまたはパスワードが正しくありません'
      };
    }
    
    // ログイン成功 - ユーザー情報を取得
    const staffName = String(matchedRow[nameIdx] || '').trim();
    const rawRole = String(matchedRow[roleIdx] || '').trim();
    const storeName = String(matchedRow[storeIdx] || '').trim();
    
    // 権限マッピング: "管理者"→経営オーナー, "オーナー"→FCオーナー, 空→スタッフ
    let role;
    if (rawRole === '管理者') {
      role = '経営オーナー';
    } else if (rawRole === 'オーナー') {
      role = 'FCオーナー';
    } else {
      role = 'スタッフ';
    }
    
    // 同じ名前の人が複数行にある場合、全店舗を収集（複数店舗アクセス）
    const allStores = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowName = String(row[nameIdx] || '').trim();
      const rowStore = String(row[storeIdx] || '').trim();
      if (rowName === staffName && rowStore) {
        if (!allStores.includes(rowStore)) {
          allStores.push(rowStore);
        }
      }
    }
    
    console.log(`ログイン成功: ${staffName} (${role}), 店舗: ${allStores.join(', ')}`);
    
    // セッションを作成
    const sessionId = Utilities.getUuid();
    const session = {
      sessionId: sessionId,
      userId: normalizedLoginId,
      role: role,
      storeName: storeName,
      staffName: staffName,
      belongStores: allStores,
      workStores: allStores, // 後方互換性のため同じ値を設定
      loginTime: new Date().toISOString()
    };
    
    // セッションを保存（PropertiesServiceを使用）
    const properties = PropertiesService.getScriptProperties();
    const sessionJson = JSON.stringify(session);
    
    // 古いセッションをクリーンアップ（PropertiesService容量確保）
    try {
      const allKeys = properties.getKeys();
      const sessionKeys = allKeys.filter(k => k.startsWith('session_') || k.startsWith('user_'));
      // 50件以上のセッションがある場合、古いものを削除
      if (sessionKeys.length > 50) {
        console.log('セッションクリーンアップ: ' + sessionKeys.length + '件の古いセッションを削除');
        sessionKeys.forEach(k => properties.deleteProperty(k));
      }
    } catch (cleanupError) {
      console.error('セッションクリーンアップエラー:', cleanupError);
    }
    
    properties.setProperty('session_' + sessionId, sessionJson);
    properties.setProperty('user_' + normalizedLoginId, sessionId);
    properties.setProperty('current_session', sessionJson);
    
    // WebアプリのベースURLを取得してリダイレクトURLを生成
    const baseUrl = ScriptApp.getService().getUrl();
    const redirectUrl = baseUrl + '?sessionId=' + sessionId;
    
    return {
      success: true,
      session: session,
      redirectUrl: redirectUrl
    };
  } catch (error) {
    console.error('ログインエラー:', error);
    console.error('エラースタック:', error.stack);
    return {
      success: false,
      message: 'ログイン処理中にエラーが発生しました: ' + (error.message || error.toString())
    };
  }
}

/**
 * セッションを取得
 * @param {string} sessionId - セッションID（オプション、クライアント側から送信される場合）
 * @return {Object|null} セッション情報
 */
function getSession(sessionId = null) {
  try {
    const properties = PropertiesService.getScriptProperties();
    
    // セッションIDが指定されている場合はそれを使用
    if (sessionId) {
      const sessionData = properties.getProperty('session_' + sessionId);
      if (sessionData) {
        const session = JSON.parse(sessionData);
        // 現在のセッションとしても保存
        properties.setProperty('current_session', sessionData);
        return session;
      }
    }
    
    // セッションIDが指定されていない場合は、current_sessionを取得
    const sessionData = properties.getProperty('current_session');
    if (sessionData) {
      return JSON.parse(sessionData);
    }
    
    return null;
  } catch (error) {
    console.error('セッション取得エラー:', error);
    return null;
  }
}

/**
 * WebアプリのベースURLを取得
 * @return {string} WebアプリのベースURL
 */
function getWebAppUrl() {
  try {
    return ScriptApp.getService().getUrl();
  } catch (error) {
    console.error('WebアプリURL取得エラー:', error);
    return '';
  }
}

/**
 * セッションを保存（クライアント側から呼び出し）
 * @param {string} sessionId - セッションID
 * @return {Object} 結果
 */
function setSession(sessionId) {
  try {
    const properties = PropertiesService.getScriptProperties();
    const sessionData = properties.getProperty('session_' + sessionId);
    
    if (sessionData) {
      properties.setProperty('current_session', sessionData);
      return { success: true };
    }
    
    return { success: false, message: 'セッションが見つかりません' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * セッションをクリア
 */
function clearSession() {
  try {
    const properties = PropertiesService.getScriptProperties();
    properties.deleteProperty('current_session');
  } catch (error) {
    console.error('セッションクリアエラー:', error);
  }
}

/**
 * 現在のユーザーの権限を取得
 * @return {Object} ユーザー情報と権限
 */
function getCurrentUser() {
  const session = getSession();
  if (!session) {
    return null;
  }
  return session;
}

/**
 * 店舗へのアクセス権限をチェック
 * スタッフマスタで同じ名前が複数店舗にある場合、全店舗にアクセス可能
 * @param {string} storeName - 店舗名
 * @return {boolean} アクセス可能かどうか
 */
function canAccessStore(storeName) {
  const user = getCurrentUser();
  if (!user) {
    return false;
  }
  
  // 経営オーナーは全店舗にアクセス可能
  if (user.role === '経営オーナー') {
    return true;
  }
  
  // FCオーナー・スタッフはスタッフマスタに登録された全店舗にアクセス可能
  const allStores = [...new Set([
    ...(user.belongStores || []),
    ...(user.workStores || []),
    ...(user.storeName ? [user.storeName] : [])
  ])];
  return allStores.includes(storeName);
}

/**
 * アクセス可能な店舗リストを取得
 * @return {Array} アクセス可能な店舗名の配列
 */
function getAccessibleStores() {
  const user = getCurrentUser();
  if (!user) {
    return [];
  }
  
  // 経営オーナーは全店舗
  if (user.role === '経営オーナー') {
    try {
      // 無限ループを避けるため、セッションIDなしで直接店舗リストを取得
      return getStoreListInternal();
    } catch (error) {
      console.error('getAccessibleStores: 店舗リスト取得エラー:', error);
      return [];
    }
  }
  
  // FCオーナー・スタッフはスタッフマスタに登録された全店舗
  if (user.role === 'FCオーナー' || user.role === 'スタッフ') {
    const allStores = [...new Set([
      ...(user.belongStores || []),
      ...(user.workStores || []),
      ...(user.storeName ? [user.storeName] : [])
    ])];
    return allStores;
  }
  
  return [];
}

/**
 * 店舗リストを取得（内部実装、セッションID不要）
 * @return {Array} 店舗名の配列
 */
function getStoreListInternal() {
  // キャッシュを活用（1時間キャッシュ）
  return getCachedData('store_list', function() {
    try {
      const ss = SpreadsheetApp.openById(getSpreadsheetId());
      const storeSet = new Set();
      
      // 1. 報告データシート履歴から店舗名を取得
      const sheet = ss.getSheetByName('報告データシート履歴');
      if (sheet) {
        const lastRow = sheet.getLastRow();
        const lastCol = sheet.getLastColumn();
        if (lastRow > 0 && lastCol > 0) {
          const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
          if (data && data.length > 1) {
            const headers = data[0];
            const storeIdx = headers.indexOf('店舗');
            if (storeIdx >= 0) {
              data.slice(1).forEach(row => {
                const store = String(row[storeIdx] || '').trim();
                if (store && store !== '' && store !== 'ー' && store !== '-') {
                  storeSet.add(store);
                }
              });
            }
          }
        }
      }
      
      // 2. 顧客登録シートからも店舗名を取得（フォールバック）
      const customerSheet = ss.getSheetByName('顧客登録');
      if (customerSheet) {
        const custLastRow = customerSheet.getLastRow();
        if (custLastRow > 1) {
          const custData = customerSheet.getDataRange().getValues();
          // B列（インデックス1）が店舗
          custData.slice(1).forEach(row => {
            const store = String(row[1] || '').trim();
            if (store && store !== '' && store !== 'ー' && store !== '-') {
              storeSet.add(store);
            }
          });
        }
      }
      
      // 3. スタッフマスタからも店舗名を取得
      const staffSheet = ss.getSheetByName('スタッフマスタ');
      if (staffSheet) {
        const staffLastRow = staffSheet.getLastRow();
        if (staffLastRow > 1) {
          const staffData = staffSheet.getDataRange().getValues();
          // A列（インデックス0）が店舗名
          staffData.slice(1).forEach(row => {
            const store = String(row[0] || '').trim();
            if (store && store !== '' && store !== 'ー' && store !== '-') {
              storeSet.add(store);
            }
          });
        }
      }
      
      // 配列に変換してソート
      return Array.from(storeSet).sort();
    } catch (error) {
      console.error('getStoreListInternal error:', error);
      return [];
    }
  }, 3600); // 1時間キャッシュ
}

/**
 * スタッフマスタシートを取得（存在しない場合は作成）
 * @return {Sheet} スタッフマスタシート
 */
function getStaffMasterSheet() {
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  let sheet = ss.getSheetByName(STAFF_MASTER_SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(STAFF_MASTER_SHEET_NAME);
    const headers = ['スタッフ名', 'メールアドレス', '所属店舗', '稼働店舗', '登録日', '更新日'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  
  return sheet;
}

/**
 * コースマスタからコースリストを取得
 * A列からコース名を取得
 * @param {string} sessionId - セッションID（オプション）
 * @return {Array} コース名の配列
 */
function getCourseList(sessionId = null) {
  // セッションIDが指定されている場合はセッションを設定
  if (sessionId) {
    try {
      setSession(sessionId);
    } catch (error) {
      console.error('セッション設定エラー:', error);
    }
  }
  
  try {
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const sheet = ss.getSheetByName('コースマスタ');
    
    if (!sheet) {
      console.error('コースマスタシートが見つかりません');
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    
    if (!data || data.length <= 1) {
      console.log('コースマスタにデータがありません');
      return [];
    }
    
    // A列（インデックス0）からコース名を取得（ヘッダー行を除く）
    const courseList = [];
    
    data.slice(1).forEach((row, index) => {
      const courseName = String(row[0] || '').trim();
      if (courseName) {
        courseList.push(courseName);
      }
    });
    
    // 重複を除去
    const uniqueCourseList = [...new Set(courseList)];
    
    console.log(`getCourseList: ${uniqueCourseList.length}件のコースを取得`);
    
    return uniqueCourseList;
  } catch (error) {
    console.error('getCourseList error:', error);
    console.error('エラースタック:', error.stack);
    return [];
  }
}

/**
 * 店舗名に基づいてスタッフリストを取得
 * スタッフマスタシートのA列が店舗名、E列（名前）がスタッフ名
 * @param {string} storeName - 店舗名
 * @param {string} sessionId - セッションID（オプション）
 * @return {Array} スタッフ名の配列
 */
function getStaffByStore(storeName, sessionId = null) {
  // セッションIDが指定されている場合はセッションを設定
  if (sessionId) {
    try {
      setSession(sessionId);
    } catch (error) {
      console.error('セッション設定エラー:', error);
    }
  }
  
  try {
    const sheet = getStaffMasterSheet();
    const data = sheet.getDataRange().getValues();
    
    if (!data || data.length <= 1) {
      console.log('スタッフマスタにデータがありません');
      return [];
    }
    
    const headers = data[0];
    // ヘッダーから動的に列インデックスを取得
    const findIdx = (candidates, fallback) => {
      for (const c of candidates) {
        const idx = headers.indexOf(c);
        if (idx >= 0) return idx;
      }
      return fallback;
    };
    
    const storeIdx = findIdx(['店舗名', '店舗', '所属店舗'], 0);  // A列
    const nameIdx = findIdx(['名前', '氏名', 'スタッフ名'], 4);   // E列（名前）
    
    const staffList = [];
    
    console.log(`getStaffByStore: 検索店舗「${storeName}」、名前列: ${nameIdx}、データ行数: ${data.length - 1}`);
    
    data.slice(1).forEach((row) => {
      const rowStore = String(row[storeIdx] || '').trim();
      const staffName = String(row[nameIdx] || '').trim();
      
      if (!staffName) return;
      
      // 店舗名が指定されていない場合は全スタッフを返す
      if (!storeName) {
        staffList.push(staffName);
        return;
      }
      
      // 店舗名が一致するかチェック
      if (rowStore === storeName) {
        staffList.push(staffName);
      }
    });
    
    // 重複を除去してソート
    const uniqueStaffList = [...new Set(staffList)].sort();
    
    console.log(`getStaffByStore: 店舗「${storeName}」のスタッフ ${uniqueStaffList.length}名を取得: ${uniqueStaffList.join(', ')}`);
    
    return uniqueStaffList;
  } catch (error) {
    console.error('getStaffByStore error:', error);
    console.error('エラースタック:', error.stack);
    return [];
  }
}

/**
 * 実績報告シートを取得（存在しない場合は作成）
 * @return {Sheet} 実績報告シート
 */
function getPerformanceReportSheet() {
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  let sheet = ss.getSheetByName(PERFORMANCE_REPORT_SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(PERFORMANCE_REPORT_SHEET_NAME);
    const headers = [
      'ID',
      '報告年月',
      'スタッフ名',
      '所属店舗',
      'セッション本数',
      '出勤回数',
      '翌月終了予定者数',
      '現在の担当人数',
      '追加希望担当人数',
      '売上顧客数',
      '報告日時'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  } else {
    // 既存シートにIDカラムがない場合は追加
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (headers.indexOf('ID') < 0) {
      // IDカラムを先頭に追加
      sheet.insertColumnBefore(1);
      sheet.getRange(1, 1).setValue('ID');
      sheet.getRange(1, 1).setFontWeight('bold');
      
      // 既存データにIDを付与
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        for (let i = 2; i <= lastRow; i++) {
          sheet.getRange(i, 1).setValue(Utilities.getUuid());
        }
      }
    }
  }
  
  return sheet;
}

/**
 * 売上顧客シートを取得（存在しない場合は作成）
 * @return {Sheet} 売上顧客シート
 */
function getSalesCustomerSheet() {
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  let sheet = ss.getSheetByName('売上顧客');
  
  if (!sheet) {
    sheet = ss.insertSheet('売上顧客');
    const headers = [
      '報告年月',
      'スタッフ名',
      '顧客名',
      '種別',
      '店舗名',
      '報告日時'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  
  return sheet;
}

/**
 * SNS・記事投稿シートを取得（なければ作成）
 */
function getSnsPostSheet() {
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  let sheet = ss.getSheetByName('SNS投稿');
  
  if (!sheet) {
    sheet = ss.insertSheet('SNS投稿');
    const headers = [
      '実績報告ID',
      '報告年月',
      'スタッフ名',
      '種別',
      'URL',
      'メモ',
      '画像URL',
      '登録日時'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#667eea').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  
  return sheet;
}

/**
 * 指定年月の第3日曜日を返す（その日の0時0分0秒）
 * @param {number} year - 年
 * @param {number} month - 月（1-12）
 * @return {Date} 第3日曜日のDate
 */
function getThirdSunday(year, month) {
  // 月の1日
  const first = new Date(year, month - 1, 1);
  // 1日の曜日（0=日曜）
  const dayOfWeek = first.getDay();
  // 1日から最初の日曜までの日数（1日が日曜なら0、月曜なら1...土曜なら6）
  const daysToFirstSunday = (7 - dayOfWeek) % 7;
  // 第1日曜 = 1 + daysToFirstSunday（1日が日曜の場合は1日）
  const firstSunday = dayOfWeek === 0 ? 1 : 1 + daysToFirstSunday;
  const thirdSunday = firstSunday + 14;
  return new Date(year, month - 1, thirdSunday, 0, 0, 0, 0);
}

/**
 * 実績報告の集計期間を返す（前月の第3日曜の翌日 〜 当月の第3日曜の終日）
 * @param {number} targetYear - 報告年月の年
 * @param {number} targetMonth - 報告年月の月（1-12）
 * @return {{ start: Date, end: Date }} 集計期間
 */
function getReportPeriodStartEnd(targetYear, targetMonth) {
  const prevYear = targetMonth === 1 ? targetYear - 1 : targetYear;
  const prevMonth = targetMonth === 1 ? 12 : targetMonth - 1;
  const thirdSunPrev = getThirdSunday(prevYear, prevMonth);
  const thirdSunCurr = getThirdSunday(targetYear, targetMonth);
  const start = new Date(thirdSunPrev);
  start.setDate(start.getDate() + 1);
  start.setHours(0, 0, 0, 0);
  const end = new Date(thirdSunCurr);
  end.setHours(23, 59, 59, 999);
  return { start: start, end: end };
}

/**
 * 日付が該当月かどうかをチェック
 * @param {*} dateValue - 日付値（Dateオブジェクト、文字列、シリアル値など）
 * @param {number} targetYear - 対象年
 * @param {number} targetMonth - 対象月
 * @return {boolean} 該当月かどうか
 */
function isDateInTargetMonth(dateValue, targetYear, targetMonth) {
  if (!dateValue) return false;
  
  try {
    let date;
    if (Object.prototype.toString.call(dateValue) === '[object Date]') {
      date = dateValue;
    } else {
      const dateStr = String(dateValue).trim();
      // 様々な日付形式を試す
      if (dateStr.match(/^\d{4}\/\d{1,2}\/\d{1,2}/)) {
        date = new Date(dateStr.replace(/\//g, '-'));
      } else if (dateStr.match(/^\d{4}-\d{1,2}-\d{1,2}/)) {
        date = new Date(dateStr);
      } else {
        // スプレッドシートのシリアル値の可能性
        const serial = parseFloat(dateStr);
        if (!isNaN(serial) && serial > 40000) {
          date = new Date((serial - 25569) * 86400 * 1000);
        } else {
          return false;
        }
      }
    }
    
    if (date && !isNaN(date.getTime())) {
      const dateYear = date.getFullYear();
      const dateMonth = date.getMonth() + 1;
      return (dateYear === targetYear && dateMonth === targetMonth);
    }
  } catch (e) {
    // 日付のパースに失敗した場合は無視
  }
  
  return false;
}

/**
 * 店舗名に基づいて顧客データを取得（集計期間内の顧客）
 * 集計期間＝前月の第3日曜の翌日〜当月の第3日曜の終日
 * 「顧客マスタ」シートから入会、「顧客継続」シートから継続を取得
 * @param {string} storeName - 店舗名
 * @param {string} yearMonth - 年月（例: "2025年1月"）
 * @param {string} sessionId - セッションID（オプション）
 * @return {Array} 顧客データの配列（コース、支払い状況、売上を含む）
 */
function getCustomersByStore(storeName, yearMonth, sessionId = null) {
  // セッションIDが指定されている場合はセッションを設定
  if (sessionId) {
    try {
      setSession(sessionId);
    } catch (error) {
      console.error('セッション設定エラー:', error);
    }
  }
  
  // 権限チェック
  if (!canAccessStore(storeName)) {
    return [];
  }
  
  try {
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    
    // 年月を正規化（例: "2025年1月" -> {year: 2025, month: 1}）
    let targetYear, targetMonth;
    if (yearMonth) {
      const match = yearMonth.match(/(\d{4})年(\d{1,2})月/);
      if (match) {
        targetYear = parseInt(match[1], 10);
        targetMonth = parseInt(match[2], 10);
      }
    }
    
    // 集計期間: 前月の第3日曜の翌日 〜 当月の第3日曜の終日
    let periodStart = null;
    let periodEnd = null;
    if (targetYear && targetMonth) {
      const period = getReportPeriodStartEnd(targetYear, targetMonth);
      periodStart = period.start;
      periodEnd = period.end;
    }
    
    function isDateInReportPeriod(dateValue) {
      if (!periodStart || !periodEnd || !dateValue) return false;
      let date = null;
      if (dateValue instanceof Date && !isNaN(dateValue)) {
        date = dateValue;
      } else if (typeof dateValue === 'string') {
        const str = String(dateValue).trim();
        if (str.includes('-')) {
          const parts = str.split(/[\s-]/);
          if (parts.length >= 3) {
            date = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
          }
        } else if (str.includes('/')) {
          const parts = str.split(/[\s\/]/);
          if (parts.length >= 3) {
            date = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
          }
        }
      } else if (typeof dateValue === 'number') {
        date = new Date((dateValue - 25569) * 86400 * 1000);
      }
      if (date && !isNaN(date.getTime())) {
        return date >= periodStart && date <= periodEnd;
      }
      return false;
    }
    
    const customers = [];
    
    // 1. 「顧客マスタ」シートから入会ステータスの顧客を取得
    const masterSheet = ss.getSheetByName(CUSTOMER_MASTER_SHEET_NAME);
    if (masterSheet) {
      const masterData = masterSheet.getDataRange().getValues();
      if (masterData && masterData.length > 1) {
        // 列のインデックス（0ベース）
        const statusIdx = 2; // C列：ステータス
        const dateIdx = 3;   // D列：日付（集計期間チェック用）
        const courseIdx = 18; // S列：コース
        const paymentIdx = 19; // T列：支払い状況
        const salesIdx = 21;   // V列：売上
        
        // 店舗名列のインデックスを取得
        let storeIdx = 0; // A列をデフォルト
        if (masterData.length > 0) {
          const headers = masterData[0];
          const storeHeaderIdx = headers.findIndex(h => {
            const headerStr = String(h || '').trim().toLowerCase();
            return headerStr === '店舗名' || headerStr === '店舗' || headerStr === 'store' || headerStr === 'storename';
          });
          if (storeHeaderIdx >= 0) {
            storeIdx = storeHeaderIdx;
          }
        }
        
        // 顧客名列のインデックスを取得
        let nameIdx = 1; // B列をデフォルト
        if (masterData.length > 0) {
          const headers = masterData[0];
          const nameHeaderIdx = headers.findIndex(h => {
            const headerStr = String(h || '').trim().toLowerCase();
            return headerStr === '名前' || headerStr === 'name' || headerStr === '顧客名' || headerStr === '氏名';
          });
          if (nameHeaderIdx >= 0) {
            nameIdx = nameHeaderIdx;
          }
        }
        
        masterData.slice(1).forEach((row, index) => {
          const rowStore = String(row[storeIdx] || '').trim();
          const status = String(row[statusIdx] || '').trim();
          const dateValue = row[dateIdx];
          
          // 店舗名が一致しない場合はスキップ
          if (storeName && rowStore !== storeName) {
            return;
          }
          
          // ステータスが「入会」でない場合はスキップ
          if (status !== '入会') {
            return;
          }
          
          // D列が集計期間内かチェック（前月第3日曜翌日〜当月第3日曜）
          if (periodStart && periodEnd && !isDateInReportPeriod(dateValue)) {
            return;
          }
          
          const name = String(row[nameIdx] || '').trim();
          if (!name || name === '' || name === '入会' || name === '継続') {
            return; // 名前が空、またはステータス値の場合はスキップ
          }
          
          customers.push({
            name: name,
            storeName: rowStore || storeName,
            status: '入会',
            course: String(row[courseIdx] || '').trim(),
            paymentStatus: String(row[paymentIdx] || '').trim(),
            sales: String(row[salesIdx] || '').trim(),
            rowIndex: index + 2,
            sheetName: CUSTOMER_MASTER_SHEET_NAME
          });
        });
      }
    }
    
    // 2. 「顧客継続」シートから継続ステータスの顧客を取得
    const continueSheet = ss.getSheetByName(CUSTOMER_CONTINUE_SHEET_NAME);
    if (continueSheet) {
      const continueData = continueSheet.getDataRange().getValues();
      if (continueData && continueData.length > 1) {
        // 列のインデックス（0ベース）
        const statusIdx = 2; // C列：ステータス
        const dateIdx = 3;   // D列：日付（集計期間チェック用）
        const courseIdx = 18; // S列：コース
        const paymentIdx = 19; // T列：支払い状況
        const salesIdx = 21;   // V列：売上
        
        // 店舗名列のインデックスを取得
        let storeIdx = 0; // A列をデフォルト
        if (continueData.length > 0) {
          const headers = continueData[0];
          const storeHeaderIdx = headers.findIndex(h => {
            const headerStr = String(h || '').trim().toLowerCase();
            return headerStr === '店舗名' || headerStr === '店舗' || headerStr === 'store' || headerStr === 'storename';
          });
          if (storeHeaderIdx >= 0) {
            storeIdx = storeHeaderIdx;
          }
        }
        
        // 顧客名列のインデックスを取得
        let nameIdx = 1; // B列をデフォルト
        if (continueData.length > 0) {
          const headers = continueData[0];
          const nameHeaderIdx = headers.findIndex(h => {
            const headerStr = String(h || '').trim().toLowerCase();
            return headerStr === '名前' || headerStr === 'name' || headerStr === '顧客名' || headerStr === '氏名';
          });
          if (nameHeaderIdx >= 0) {
            nameIdx = nameHeaderIdx;
          }
        }
        
        continueData.slice(1).forEach((row, index) => {
          const rowStore = String(row[storeIdx] || '').trim();
          const status = String(row[statusIdx] || '').trim();
          const dateValue = row[dateIdx];
          
          // 店舗名が一致しない場合はスキップ
          if (storeName && rowStore !== storeName) {
            return;
          }
          
          // ステータスが「継続」でない場合はスキップ
          if (status !== '継続') {
            return;
          }
          
          // D列が集計期間内かチェック（前月第3日曜翌日〜当月第3日曜）
          if (periodStart && periodEnd && !isDateInReportPeriod(dateValue)) {
            return;
          }
          
          const name = String(row[nameIdx] || '').trim();
          if (!name || name === '' || name === '入会' || name === '継続') {
            return; // 名前が空、またはステータス値の場合はスキップ
          }
          
          customers.push({
            name: name,
            storeName: rowStore || storeName,
            status: '継続',
            course: String(row[courseIdx] || '').trim(),
            paymentStatus: String(row[paymentIdx] || '').trim(),
            sales: String(row[salesIdx] || '').trim(),
            rowIndex: index + 2,
            sheetName: CUSTOMER_CONTINUE_SHEET_NAME
          });
        });
      }
    }
    
    console.log(`getCustomersByStore: ${customers.length}件の顧客を取得しました（店舗: ${storeName}, 年月: ${yearMonth}）`);
    
    return customers;
  } catch (error) {
    console.error('getCustomersByStore error:', error);
    console.error('エラースタック:', error.stack);
    return [];
  }
}

/**
 * ログイン中のスタッフに紐づく顧客で、選択月に新規または継続登録された顧客を取得
 * @param {string} storeName - 店舗名
 * @param {string} yearMonth - 年月（例: "2025年1月"）
 * @param {string} sessionId - セッションID（オプション）
 * @return {Array} 顧客データの配列
 */
/**
 * シート構造の診断用関数（デバッグ用）
 * スプレッドシートの顧客登録・顧客継続履歴のヘッダーとサンプルデータを返す
 */
function debugSheetStructure(sessionId) {
  if (sessionId) {
    try { setSession(sessionId); } catch (e) {}
  }
  const user = getCurrentUser();
  if (!user) return { error: 'ログインが必要です' };

  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  const result = { staffName: user.staffName, userId: user.userId };

  // 顧客登録シート
  const masterSheet = ss.getSheetByName('顧客登録');
  if (masterSheet) {
    const data = masterSheet.getDataRange().getValues();
    result.customerRegistration = {
      headers: data[0].map((h, i) => `[${i}]${h}`),
      sampleRows: data.slice(1, 6).map(row => row.map((v, i) => `[${i}]${v}`)),
      totalRows: data.length - 1
    };
  }

  // 顧客継続履歴シート
  const continueSheet = ss.getSheetByName('顧客継続履歴');
  if (continueSheet) {
    const data = continueSheet.getDataRange().getValues();
    result.customerContinueHistory = {
      headers: data[0].map((h, i) => `[${i}]${h}`),
      sampleRows: data.slice(1, 6).map(row => row.map((v, i) => `[${i}]${v}`)),
      totalRows: data.length - 1
    };
  }

  return result;
}

function getStaffCustomersByStore(storeName, yearMonth, sessionId = null) {
  // ユーザー取得（sessionIdがある場合は直接そのセッションから取得）
  let user = null;
  if (sessionId) {
    try {
      user = getSession(sessionId);
      console.log('getStaffCustomersByStore: sessionIdからユーザー取得:', sessionId);
    } catch (error) {
      console.error('セッション取得エラー:', error);
    }
  }
  
  // sessionIdからユーザーが取得できない場合はcurrent_sessionを使用
  if (!user) {
    user = getCurrentUser();
  }
  
  if (!user) {
    console.error('ユーザーが取得できませんでした');
    return [];
  }
  
  // 現在のスタッフ名を取得（複数の候補を試す）
  const staffName = user.staffName || user.userId || '';
  const staffNameVariants = [];
  if (staffName) {
    staffNameVariants.push(staffName.trim());
    // 全角スペースを半角に変換
    staffNameVariants.push(staffName.replace(/　/g, ' ').trim());
    // 半角スペースを全角に変換
    staffNameVariants.push(staffName.replace(/ /g, '　').trim());
    // スペースなし
    staffNameVariants.push(staffName.replace(/[\s　]/g, '').trim());
  }
  // 重複を除去
  const uniqueVariants = [...new Set(staffNameVariants)];
  
  if (!staffName) {
    console.error('スタッフ名が取得できませんでした。ユーザー情報:', JSON.stringify(user));
    return [];
  }
  
  console.log('getStaffCustomersByStore呼び出し:', { 
    storeName, 
    yearMonth, 
    staffName: staffName,
    staffNameVariants: uniqueVariants,
    user: { staffName: user.staffName, userId: user.userId }
  });
  
  // 店舗のアクセス権限チェック
  if (storeName && !canAccessStore(storeName)) {
    console.error('店舗へのアクセス権限がありません:', storeName);
    return [];
  }
  
  try {
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const customers = [];
    
    // 年月を正規化（例: "2025年1月" -> {year: 2025, month: 1}）
    let targetYear, targetMonth;
    if (yearMonth) {
      const match = yearMonth.match(/(\d{4})年(\d{1,2})月/);
      if (match) {
        targetYear = parseInt(match[1], 10);
        targetMonth = parseInt(match[2], 10);
      }
    }
    
    // 集計期間: 前月の第3日曜の翌日 〜 当月の第3日曜の終日
    let periodStart = null;
    let periodEnd = null;
    if (targetYear && targetMonth) {
      const period = getReportPeriodStartEnd(targetYear, targetMonth);
      periodStart = period.start;
      periodEnd = period.end;
    }
    
    console.log('集計期間（前月第3日曜翌日〜当月第3日曜）:', periodStart, '〜', periodEnd);
    
    // 日付が集計期間内かチェックするヘルパー関数
    function isDateInReportPeriod(dateValue) {
      if (!periodStart || !periodEnd || !dateValue) return false;
      
      let date = null;
      if (dateValue instanceof Date && !isNaN(dateValue)) {
        date = dateValue;
      } else if (typeof dateValue === 'string') {
        const str = dateValue.trim();
        if (str.includes('-')) {
          const parts = str.split(/[\s-]/);
          if (parts.length >= 3) {
            date = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
          }
        } else if (str.includes('/')) {
          const parts = str.split(/[\s\/]/);
          if (parts.length >= 3) {
            date = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
          }
        }
      } else if (typeof dateValue === 'number') {
        // Excelシリアル値の場合
        date = new Date((dateValue - 25569) * 86400 * 1000);
      }
      
      if (date && !isNaN(date.getTime())) {
        return date >= periodStart && date <= periodEnd;
      }
      return false;
    }
    
    // ヘッダーから列を動的に検出するヘルパー関数
    // 複数の候補名で検索し、見つからなければフォールバックインデックスを使用
    function findColumnIndex(headers, candidateNames, fallbackIdx) {
      for (const name of candidateNames) {
        const idx = headers.indexOf(name);
        if (idx >= 0) {
          console.log(`列検出: ヘッダー「${name}」→ ${idx}列目（${String.fromCharCode(65 + idx)}列）`);
          return idx;
        }
      }
      if (fallbackIdx >= 0 && fallbackIdx < headers.length) {
        console.log(`列がヘッダーから見つかりません（候補: ${candidateNames.join(',')}）。フォールバック: ${fallbackIdx}列目（${String.fromCharCode(65 + fallbackIdx)}列）`);
        return fallbackIdx;
      }
      console.log(`列が見つかりません（候補: ${candidateNames.join(',')}）: -1`);
      return -1;
    }
    
    // 担当者名マッチング関数
    function isStaffMatch(rowStaffRaw) {
      if (!staffName || uniqueVariants.length === 0) return true; // 条件なしなら全て一致
      const rowStaff = String(rowStaffRaw || '').trim();
      if (!rowStaff) return false; // 担当者名が空の行は不一致
      
      // スペースなしバージョンも比較対象に含める
      const rowStaffNoSpace = rowStaff.replace(/[\s　]/g, '');
      
      return uniqueVariants.some(variant => {
        const v = variant.trim();
        const vNoSpace = v.replace(/[\s　]/g, '');
        return rowStaff === v || 
               rowStaffNoSpace === vNoSpace ||
               rowStaff.includes(v) || 
               v.includes(rowStaff);
      });
    }
    
    // === 1. 顧客登録シートから取得 ===
    const masterSheet = ss.getSheetByName('顧客登録');
    if (masterSheet) {
      const masterData = masterSheet.getDataRange().getValues();
      const masterHeaders = masterData[0].map(h => String(h).trim());
      
      // ヘッダー名の揺れに対応（複数の候補名で検索）
      const storeIdx = findColumnIndex(masterHeaders, ['店舗名', '店舗'], -1);
      const nameIdx = findColumnIndex(masterHeaders, ['名前', '顧客名', '氏名'], -1);
      const dateIdx = findColumnIndex(masterHeaders, ['日付', '登録日', '入会日', '更新日付'], -1);
      const courseIdx = findColumnIndex(masterHeaders, ['コース'], -1);
      const paymentIdx = findColumnIndex(masterHeaders, ['支払い状況', '支払状況', '決済状況'], -1);
      const salesIdx = findColumnIndex(masterHeaders, ['売上', '売上金額'], -1);
      const statusIdx = findColumnIndex(masterHeaders, ['ステータス', '体験/入会', 'ステータス区分', '区分'], -1);
      // 担当者列: ヘッダーから動的検出、見つからなければC列（インデックス2）をフォールバック
      const staffIdx = findColumnIndex(masterHeaders, ['担当者', '担当者名', '担当', 'スタッフ', 'スタッフ名'], 2);
      
      console.log('=== 顧客登録シート ===');
      console.log('ヘッダー:', masterHeaders.map((h, i) => `[${i}]${h}`).join(', '));
      console.log('データ件数:', masterData.length - 1);
      console.log('担当者列:', staffIdx, `(${String.fromCharCode(65 + staffIdx)}列)`);
      console.log('検索条件: 店舗=' + storeName + ', スタッフ名=' + staffName + ', 年月=' + yearMonth);
      
      let matchCount = 0;
      let skipReasons = { noName: 0, storeMismatch: 0, staffMismatch: 0, statusExcluded: 0, dateMismatch: 0 };
      
      masterData.slice(1).forEach((row, idx) => {
        const rowStore = storeIdx >= 0 ? String(row[storeIdx] || '').trim() : '';
        const rowStaff = String(row[staffIdx] || '').trim();
        const name = nameIdx >= 0 ? String(row[nameIdx] || '').trim() : '';
        const dateValue = dateIdx >= 0 ? row[dateIdx] : null;
        const status = statusIdx >= 0 ? String(row[statusIdx] || '').trim() : '';
        
        // 最初の5件は常にデバッグ出力
        if (idx < 5) {
          console.log(`顧客登録行${idx + 1}: 店舗=[${rowStore}], 担当者(${String.fromCharCode(65 + staffIdx)}列)=[${rowStaff}], 名前=[${name}], 日付=[${dateValue}], ステータス=[${status}]`);
        }
        
        if (!name) { skipReasons.noName++; return; }
        // 店舗列がある場合のみ店舗フィルタを適用
        if (storeIdx >= 0 && storeName && rowStore && rowStore !== storeName) { skipReasons.storeMismatch++; return; }
        
        if (!isStaffMatch(row[staffIdx])) {
          skipReasons.staffMismatch++;
          if (idx < 10) {
            console.log(`顧客登録: 担当者不一致 - 検索=[${staffName}], シート値=[${rowStaff}]`);
          }
          return;
        }
        
        // ステータスが「入会」「体験」「継続」のいずれか、または空の場合は含める
        // 「体験/入会」列の場合は「体験」「入会」の値がそのまま入る
        if (status && status !== '入会' && status !== '体験' && status !== '継続') {
          skipReasons.statusExcluded++;
          return;
        }
        
        // 集計期間内に登録された顧客のみ（前月第3日曜翌日〜当月第3日曜）
        if (!isDateInReportPeriod(dateValue)) {
          skipReasons.dateMismatch++;
          if (matchCount === 0 && idx < 10) {
            console.log(`顧客登録: 日付不一致 - 名前=[${name}], 日付=[${dateValue}], 範囲=${periodStart}〜${periodEnd}`);
          }
          return;
        }
        
        matchCount++;
        const course = courseIdx >= 0 ? String(row[courseIdx] || '').trim() : '';
        const paymentStatus = paymentIdx >= 0 ? String(row[paymentIdx] || '').trim() : '';
        const salesStr = salesIdx >= 0 ? String(row[salesIdx] || '').trim() : '';
        
        let salesAmount = '';
        if (salesStr) {
          const cleanedSales = salesStr.replace(/[^0-9.-]/g, '');
          const numSales = parseFloat(cleanedSales);
          if (!isNaN(numSales)) {
            salesAmount = numSales.toString();
          }
        }
        
        customers.push({
          name: name,
          storeName: rowStore || storeName,
          status: status || '入会',
          staff: rowStaff,
          course: course,
          paymentStatus: paymentStatus,
          sales: salesAmount
        });
      });
      
      console.log('顧客登録シート結果: マッチ=' + matchCount + '件, スキップ理由:', JSON.stringify(skipReasons));
    } else {
      console.log('顧客登録シートが見つかりません');
    }
    
    // === 2. 顧客継続履歴シートから取得 ===
    const continueSheet = ss.getSheetByName('顧客継続履歴');
    if (continueSheet) {
      const continueData = continueSheet.getDataRange().getValues();
      const continueHeaders = continueData[0].map(h => String(h).trim());
      
      // ヘッダー名の揺れに対応（複数の候補名で検索）
      const storeIdx = findColumnIndex(continueHeaders, ['店舗名', '店舗'], -1);
      const nameIdx = findColumnIndex(continueHeaders, ['名前', '顧客名', '氏名'], -1);
      const dateIdx = findColumnIndex(continueHeaders, ['継続日', '更新日付', '日付', '登録日'], -1);
      const courseIdx = findColumnIndex(continueHeaders, ['コース'], -1);
      const paymentIdx = findColumnIndex(continueHeaders, ['支払い状況', '支払状況', '決済状況'], -1);
      const salesIdx = findColumnIndex(continueHeaders, ['売上', '売上金額'], -1);
      // 担当者列: ヘッダーから動的検出、見つからなければD列（インデックス3）をフォールバック
      const staffIdx = findColumnIndex(continueHeaders, ['担当者', '担当者名', '担当', 'スタッフ', 'スタッフ名'], 3);
      
      const actualDateIdx = dateIdx;
      
      console.log('=== 顧客継続履歴シート ===');
      console.log('ヘッダー:', continueHeaders.map((h, i) => `[${i}]${h}`).join(', '));
      console.log('データ件数:', continueData.length - 1);
      console.log('担当者列:', staffIdx, `(${String.fromCharCode(65 + staffIdx)}列)`);
      console.log('日付列:', actualDateIdx >= 0 ? actualDateIdx : 'なし');
      
      let matchCount = 0;
      let skipReasons = { noName: 0, storeMismatch: 0, staffMismatch: 0, dateMismatch: 0, duplicate: 0 };
      
      continueData.slice(1).forEach((row, idx) => {
        const rowStore = storeIdx >= 0 ? String(row[storeIdx] || '').trim() : '';
        const rowStaff = String(row[staffIdx] || '').trim();
        const name = nameIdx >= 0 ? String(row[nameIdx] || '').trim() : '';
        const dateValue = actualDateIdx >= 0 ? row[actualDateIdx] : null;
        
        // 最初の5件は常にデバッグ出力
        if (idx < 5) {
          console.log(`継続履歴行${idx + 1}: 店舗=[${rowStore}], 担当者(${String.fromCharCode(65 + staffIdx)}列)=[${rowStaff}], 名前=[${name}], 日付=[${dateValue}]`);
        }
        
        if (!name) { skipReasons.noName++; return; }
        // 店舗列がある場合のみ店舗フィルタを適用（顧客継続履歴には店舗列がない場合がある）
        if (storeIdx >= 0 && storeName && rowStore && rowStore !== storeName) { skipReasons.storeMismatch++; return; }
        
        if (!isStaffMatch(row[staffIdx])) {
          skipReasons.staffMismatch++;
          if (idx < 10) {
            console.log(`継続履歴: 担当者不一致 - 検索=[${staffName}], シート値=[${rowStaff}]`);
          }
          return;
        }
        
        // 集計期間内に継続登録された顧客のみ（前月第3日曜翌日〜当月第3日曜）
        if (!isDateInReportPeriod(dateValue)) {
          skipReasons.dateMismatch++;
          if (matchCount === 0 && idx < 10) {
            console.log(`継続履歴: 日付不一致 - 名前=[${name}], 日付=[${dateValue}], 範囲=${periodStart}〜${periodEnd}`);
          }
          return;
        }
        
        // 既に同じ顧客が追加されている場合はスキップ（重複防止）
        const existing = customers.find(c => c.name === name && c.storeName === (rowStore || storeName));
        if (existing) { skipReasons.duplicate++; return; }
        
        matchCount++;
        const course = courseIdx >= 0 ? String(row[courseIdx] || '').trim() : '';
        const paymentStatus = paymentIdx >= 0 ? String(row[paymentIdx] || '').trim() : '';
        const salesStr = salesIdx >= 0 ? String(row[salesIdx] || '').trim() : '';
        
        let salesAmount = '';
        if (salesStr) {
          const cleanedSales = salesStr.replace(/[^0-9.-]/g, '');
          const numSales = parseFloat(cleanedSales);
          if (!isNaN(numSales)) {
            salesAmount = numSales.toString();
          }
        }
        
        customers.push({
          name: name,
          storeName: rowStore || storeName,
          status: '継続',
          staff: rowStaff,
          course: course,
          paymentStatus: paymentStatus,
          sales: salesAmount
        });
      });
      
      console.log('顧客継続履歴シート結果: マッチ=' + matchCount + '件, スキップ理由:', JSON.stringify(skipReasons));
    } else {
      console.log('顧客継続履歴シートが見つかりません');
    }
    
    console.log(`getStaffCustomersByStore: 合計${customers.length}件の顧客を取得（店舗: ${storeName}, 年月: ${yearMonth}, スタッフ: ${staffName}）`);
    
    return customers;
  } catch (error) {
    console.error('getStaffCustomersByStore error:', error);
    console.error('エラースタック:', error.stack);
    return [];
  }
}

/**
 * スタッフが所属・稼働している店舗リストを取得
 * @param {string} sessionId - セッションID（オプション）
 * @return {Array} 店舗名の配列
 */
function getStaffStores(sessionId = null) {
  // ユーザー取得（sessionIdがある場合は直接そのセッションから取得）
  let user = null;
  if (sessionId) {
    try {
      user = getSession(sessionId);
      console.log('getStaffStores: sessionIdからユーザー取得:', sessionId);
    } catch (error) {
      console.error('セッション取得エラー:', error);
    }
  }
  
  // sessionIdからユーザーが取得できない場合はcurrent_sessionを使用
  if (!user) {
    user = getCurrentUser();
  }
  
  if (!user) {
    console.error('getStaffStores: ユーザーが取得できません');
    return [];
  }
  
  console.log('getStaffStores: user =', JSON.stringify(user));
  
  // 経営オーナーの場合は全店舗
  if (user.role === '経営オーナー') {
    return getStoreListInternal();
  }
  
  const stores = [];
  
  // 店舗名（必ず追加）
  if (user.storeName) {
    stores.push(user.storeName);
  }
  
  // 所属店舗を追加
  if (user.belongStores && Array.isArray(user.belongStores) && user.belongStores.length > 0) {
    stores.push(...user.belongStores);
  }
  
  // 稼働店舗を追加
  if (user.workStores && Array.isArray(user.workStores) && user.workStores.length > 0) {
    stores.push(...user.workStores);
  }
  
  // 重複を除去してソート
  const uniqueStores = [...new Set(stores)].filter(s => s).sort();
  console.log('getStaffStores: 結果 =', uniqueStores);
  
  return uniqueStores;
}

/**
 * 実績報告を保存
 * @param {Object} reportData - 実績報告データ
 * @param {string} sessionId - セッションID（オプション）
 * @return {Object} 保存結果
 */
function savePerformanceReport(reportData, sessionId = null) {
  // セッションIDが指定されている場合はセッションを設定
  if (sessionId) {
    try {
      setSession(sessionId);
    } catch (error) {
      console.error('セッション設定エラー:', error);
    }
  }
  
  const user = getCurrentUser();
  if (!user) {
    return { success: false, message: 'ログインが必要です' };
  }
  
  try {
    const reportSheet = getPerformanceReportSheet();
    const salesCustomerSheet = getSalesCustomerSheet();
    
    const now = new Date();
    const reportDateTime = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    
    // 実績報告IDを生成（UUID）
    const reportId = Utilities.getUuid();
    
    // 実績報告を保存
    const reportRow = [
      reportId,                        // ID（UUID）
      reportData.yearMonth,           // 報告年月
      user.staffName || user.userId,  // スタッフ名
      reportData.storeName || '',     // 所属店舗
      reportData.sessionCount || 0,   // セッション本数
      reportData.attendanceCount || 0, // 出勤回数
      reportData.endNextMonthCount || 0, // 翌月終了予定者数
      reportData.currentCustomerCount || 0, // 現在の担当人数
      reportData.desiredCustomerCount || 0, // 追加希望担当人数
      reportData.salesCustomers ? reportData.salesCustomers.length : 0, // 売上顧客数
      reportDateTime                   // 報告日時
    ];
    
    reportSheet.appendRow(reportRow);
    
    // 売上顧客を保存
    if (reportData.salesCustomers && Array.isArray(reportData.salesCustomers)) {
      reportData.salesCustomers.forEach(customer => {
        const salesRow = [
          reportData.yearMonth,           // 報告年月
          user.staffName || user.userId,  // スタッフ名
          customer.name || '',            // 顧客名
          customer.type || '',            // 種別（入会・継続・退会）
          customer.storeName || '',       // 店舗名
          reportDateTime                  // 報告日時
        ];
        salesCustomerSheet.appendRow(salesRow);
      });
    }
    
    // SNS・記事投稿を保存（URL登録のみ。ファイルアップロードは廃止）
    if (reportData.snsPosts && Array.isArray(reportData.snsPosts) && reportData.snsPosts.length > 0) {
      const snsSheet = getSnsPostSheet();
      reportData.snsPosts.forEach(function(post) {
        snsSheet.appendRow([
          reportId,                         // 実績報告ID
          reportData.yearMonth,             // 報告年月
          user.staffName || user.userId,    // スタッフ名
          post.type || '',                  // 種別
          post.url || '',                   // URL
          post.memo || '',                  // メモ
          '',                               // 画像URL（URL登録のみのため空）
          reportDateTime                    // 登録日時
        ]);
      });
    }
    
    return { success: true, message: '実績報告を保存しました', reportId: reportId };
  } catch (error) {
    console.error('savePerformanceReport error:', error);
    return { success: false, message: '実績報告の保存に失敗しました: ' + error.toString() };
  }
}

/**
 * 実績報告を取得
 * @param {string} yearMonth - 年月（例: "2025年1月"）
 * @param {string} sessionId - セッションID（オプション）
 * @return {Object} 実績報告データ
 */
function getPerformanceReport(yearMonth, sessionId = null) {
  // セッションIDが指定されている場合はセッションを設定
  if (sessionId) {
    try {
      setSession(sessionId);
    } catch (error) {
      console.error('セッション設定エラー:', error);
    }
  }
  
  const user = getCurrentUser();
  if (!user) {
    return { success: false, message: 'ログインが必要です' };
  }
  
  try {
    const reportSheet = getPerformanceReportSheet();
    const salesCustomerSheet = getSalesCustomerSheet();
    
    const reportData = reportSheet.getDataRange().getValues();
    const salesData = salesCustomerSheet.getDataRange().getValues();
    
    if (!reportData || reportData.length <= 1) {
      return { success: true, data: null };
    }
    
    const headers = reportData[0];
    const yearMonthIdx = headers.indexOf('報告年月');
    const staffNameIdx = headers.indexOf('スタッフ名');
    
    const staffName = user.staffName || user.userId;
    
    // 該当する実績報告を検索
    for (let i = 1; i < reportData.length; i++) {
      const row = reportData[i];
      const rowYearMonth = String(row[yearMonthIdx] || '').trim();
      const rowStaffName = String(row[staffNameIdx] || '').trim();
      
      if (rowYearMonth === yearMonth && rowStaffName === staffName) {
        // 実績報告データを構築
        const report = {
          yearMonth: rowYearMonth,
          staffName: rowStaffName,
          storeName: String(row[headers.indexOf('所属店舗')] || ''),
          sessionCount: Number(row[headers.indexOf('セッション本数')] || 0),
          attendanceCount: Number(row[headers.indexOf('出勤回数')] || 0),
          endNextMonthCount: Number(row[headers.indexOf('翌月終了予定者数')] || 0),
          currentCustomerCount: Number(row[headers.indexOf('現在の担当人数')] || 0),
          desiredCustomerCount: Number(row[headers.indexOf('追加希望担当人数')] || 0),
          salesCustomers: []
        };
        
        // 売上顧客データを取得
        const salesHeaders = salesData[0];
        const salesYearMonthIdx = salesHeaders.indexOf('報告年月');
        const salesStaffNameIdx = salesHeaders.indexOf('スタッフ名');
        const salesCustomerNameIdx = salesHeaders.indexOf('顧客名');
        const salesTypeIdx = salesHeaders.indexOf('種別');
        const salesStoreNameIdx = salesHeaders.indexOf('店舗名');
        
        for (let j = 1; j < salesData.length; j++) {
          const salesRow = salesData[j];
          const salesRowYearMonth = String(salesRow[salesYearMonthIdx] || '').trim();
          const salesRowStaffName = String(salesRow[salesStaffNameIdx] || '').trim();
          
          if (salesRowYearMonth === yearMonth && salesRowStaffName === staffName) {
            report.salesCustomers.push({
              name: String(salesRow[salesCustomerNameIdx] || ''),
              type: String(salesRow[salesTypeIdx] || ''),
              storeName: String(salesRow[salesStoreNameIdx] || '')
            });
          }
        }
        
        return { success: true, data: report };
      }
    }
    
    return { success: true, data: null };
  } catch (error) {
    console.error('getPerformanceReport error:', error);
    return { success: false, message: '実績報告の取得に失敗しました: ' + error.toString() };
  }
}

/**
 * ログイン中のスタッフの過去の実績報告一覧を取得
 * @param {string} sessionId - セッションID（オプション）
 * @return {Array} 実績報告一覧データの配列
 */
function getStaffPerformanceReports(sessionId = null) {
  // セッションIDが指定されている場合はセッションを設定
  if (sessionId) {
    try {
      setSession(sessionId);
    } catch (error) {
      console.error('セッション設定エラー:', error);
    }
  }
  
  const user = getCurrentUser();
  if (!user) {
    return [];
  }
  
  const staffName = user.staffName || user.userId || '';
  if (!staffName) {
    return [];
  }
  
  try {
    // キャッシュを一時的に無効化（デバッグ用）
    // const cacheKey = `staff_reports_${staffName.replace(/[^a-zA-Z0-9]/g, '_')}`;
    // return getCachedData(cacheKey, function() {
    
    // 直接データを取得（キャッシュなし）
    return (function() {
      const reportSheet = getPerformanceReportSheet();
      const lastRow = reportSheet.getLastRow();
      const lastCol = reportSheet.getLastColumn();
      
      // 早期リターン: データがない場合
      if (lastRow === 0 || lastCol === 0 || lastRow <= 1) {
        return [];
      }
      
      // 必要な範囲のみ取得（最適化）
      const reportData = reportSheet.getRange(1, 1, lastRow, lastCol).getValues();
      
      if (!reportData || reportData.length <= 1) {
        return [];
      }
      
      const headers = reportData[0];
      const idIdx = headers.indexOf('ID');
      const yearMonthIdx = headers.indexOf('報告年月');
      const staffNameIdx = headers.indexOf('スタッフ名');
      const storeNameIdx = headers.indexOf('所属店舗');
      const reportDateTimeIdx = headers.indexOf('報告日時');
      
      // 早期リターン: 必要な列が見つからない場合
      if (staffNameIdx < 0) {
        return [];
      }
      
      const reports = [];
      const staffNameVariants = [
        staffName.trim(),
        staffName.replace(/　/g, ' ').trim(),
        staffName.replace(/ /g, '　').trim()
      ];
      
      // 該当する実績報告を検索（バッチ処理）
      for (let i = 1; i < reportData.length; i++) {
        const row = reportData[i];
        const rowStaffName = String(row[staffNameIdx] || '').trim();
        
        // 担当者名が一致するかチェック
        const rowStaffTrimmed = rowStaffName.trim();
        const isMatch = staffNameVariants.some(variant => {
          const variantTrimmed = variant.trim();
          return rowStaffTrimmed === variantTrimmed || 
                 rowStaffTrimmed.includes(variantTrimmed) || 
                 variantTrimmed.includes(rowStaffTrimmed);
        });
        
        if (isMatch) {
          // IDカラムがある場合はIDを使用、ない場合や空の場合は行番号を使用（後方互換性）
          const rowId = idIdx >= 0 && row[idIdx] ? String(row[idIdx]).trim() : '';
          const reportId = rowId || String(i - 1); // IDが空なら行番号を使用
          
          console.log('getStaffPerformanceReports: 報告発見 - row:', i, 'id:', reportId, 'rowId:', rowId);
          
          reports.push({
            id: reportId,
            yearMonth: yearMonthIdx >= 0 ? String(row[yearMonthIdx] || '').trim() : '',
            staffName: rowStaffName,
            storeName: storeNameIdx >= 0 ? String(row[storeNameIdx] || '').trim() : '',
            reportDateTime: reportDateTimeIdx >= 0 ? String(row[reportDateTimeIdx] || '').trim() : ''
          });
        }
      }
      
      // 報告日時の降順でソート（新しい順）
      reports.sort((a, b) => {
        if (!a.reportDateTime) return 1;
        if (!b.reportDateTime) return -1;
        return b.reportDateTime.localeCompare(a.reportDateTime);
      });
      
      return reports;
    })(); // キャッシュなし（直接実行）
  } catch (error) {
    console.error('getStaffPerformanceReports error:', error);
    return [];
  }
}

/**
 * 実績報告一覧を取得（FCオーナー＝全件、オーナー＝自店舗のみ、スタッフ＝閲覧不可）
 * @param {string} sessionId - セッションID（オプション）
 * @return {{ error?: string, reports: Array }} エラー時は error、成功時は reports 配列
 */
function getPerformanceReportListForOwner(sessionId = null) {
  if (sessionId) {
    try {
      setSession(sessionId);
    } catch (error) {
      console.error('セッション設定エラー:', error);
    }
  }

  const user = getCurrentUser();
  if (!user) {
    return { error: 'ログインが必要です', reports: [] };
  }

  // スタッフは閲覧不可
  if (user.role === 'スタッフ') {
    return { error: 'このページを閲覧する権限がありません', reports: [] };
  }

  try {
    const reportSheet = getPerformanceReportSheet();
    const lastRow = reportSheet.getLastRow();
    const lastCol = reportSheet.getLastColumn();
    if (lastRow === 0 || lastCol === 0 || lastRow <= 1) {
      return { reports: [] };
    }

    const reportData = reportSheet.getRange(1, 1, lastRow, lastCol).getValues();
    if (!reportData || reportData.length <= 1) {
      return { reports: [] };
    }

    const headers = reportData[0];
    const idIdx = headers.indexOf('ID');
    const yearMonthIdx = headers.indexOf('報告年月');
    const staffNameIdx = headers.indexOf('スタッフ名');
    const storeNameIdx = headers.indexOf('所属店舗');
    const reportDateTimeIdx = headers.indexOf('報告日時');
    if (staffNameIdx < 0) {
      return { reports: [] };
    }

    // FCオーナー（経営オーナー）は全店舗、オーナーはアクセス可能店舗のみ
    const allowedStores = getAccessibleStores();
    const isAllStores = user.role === '経営オーナー';

    const reports = [];
    for (let i = 1; i < reportData.length; i++) {
      const row = reportData[i];
      const rowStore = storeNameIdx >= 0 ? String(row[storeNameIdx] || '').trim() : '';
      if (!isAllStores && allowedStores.length > 0 && !allowedStores.includes(rowStore)) {
        continue;
      }
      const rowId = idIdx >= 0 && row[idIdx] ? String(row[idIdx]).trim() : '';
      const reportId = rowId || String(i - 1);
      reports.push({
        id: reportId,
        yearMonth: yearMonthIdx >= 0 ? String(row[yearMonthIdx] || '').trim() : '',
        staffName: String(row[staffNameIdx] || '').trim(),
        storeName: rowStore,
        reportDateTime: reportDateTimeIdx >= 0 ? String(row[reportDateTimeIdx] || '').trim() : ''
      });
    }

    reports.sort((a, b) => {
      if (!a.reportDateTime) return 1;
      if (!b.reportDateTime) return -1;
      return b.reportDateTime.localeCompare(a.reportDateTime);
    });

    return { reports: reports };
  } catch (error) {
    console.error('getPerformanceReportListForOwner error:', error);
    return { error: '一覧の取得に失敗しました: ' + error.toString(), reports: [] };
  }
}

/**
 * 報告日時を比較用に正規化（Dateは文字列に、文字列はトリム）
 * @param {*} val - 報告日時（Dateまたは文字列）
 * @return {string}
 */
function normalizeReportDateTime(val) {
  if (val === null || val === undefined) return '';
  if (Object.prototype.toString.call(val) === '[object Date]' && !isNaN(val.getTime())) {
    return Utilities.formatDate(val, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  }
  return String(val).trim().replace(/\s+/g, ' ');
}

/**
 * 報告者（スタッフ名）と報告日時でスプレッドシートから実績報告を検索し、詳細を返す
 * 詳細ページを「報告者＋タイムスタンプ」で動的に開くために使用
 * @param {string} reporter - 報告者（スタッフ名）
 * @param {string} reportDateTime - 報告日時（例: "2025-01-18 10:30:00"）
 * @param {string} sessionId - セッションID（オプション）
 * @return {Object} getPerformanceReportById と同じ形式（success, message, data）
 */
function getPerformanceReportByReporterAndTimestamp(reporter, reportDateTime, sessionId = null) {
  if (sessionId) {
    try {
      setSession(sessionId);
    } catch (error) {
      console.error('セッション設定エラー:', error);
    }
  }

  const reporterTrim = String(reporter || '').trim();
  const dateTimeTrim = String(reportDateTime || '').trim().replace(/\s+/g, ' ');
  if (!reporterTrim || !dateTimeTrim) {
    return { success: false, message: '報告者と報告日時を指定してください' };
  }

  try {
    const reportSheet = getPerformanceReportSheet();
    const lastRow = reportSheet.getLastRow();
    const lastCol = reportSheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) {
      return { success: false, message: '実績報告が見つかりません' };
    }

    const reportData = reportSheet.getRange(1, 1, lastRow, lastCol).getValues();
    const headers = reportData[0];
    const idIdx = headers.indexOf('ID');
    const staffNameIdx = headers.indexOf('スタッフ名');
    const reportDateTimeIdx = headers.indexOf('報告日時');
    if (staffNameIdx < 0 || reportDateTimeIdx < 0) {
      return { success: false, message: '実績報告シートの形式が不正です' };
    }

    const reporterVariants = [
      reporterTrim,
      reporterTrim.replace(/　/g, ' ').trim(),
      reporterTrim.replace(/ /g, '　').trim()
    ];
    const targetNorm = dateTimeTrim;

    for (let i = 1; i < reportData.length; i++) {
      const row = reportData[i];
      const rowStaff = String(row[staffNameIdx] || '').trim();
      const rowDt = normalizeReportDateTime(row[reportDateTimeIdx]);

      const nameMatch = reporterVariants.some(function(v) {
        const vTrim = v.trim();
        return rowStaff === vTrim || rowStaff.includes(vTrim) || vTrim.includes(rowStaff);
      });
      if (!nameMatch) continue;
      if (rowDt !== targetNorm) continue;

      const reportId = idIdx >= 0 && row[idIdx] ? String(row[idIdx]).trim() : String(i - 1);
      return getPerformanceReportById(reportId, sessionId);
    }

    return { success: false, message: '指定の報告者・報告日時と一致する実績報告が見つかりません' };
  } catch (error) {
    console.error('getPerformanceReportByReporterAndTimestamp error:', error);
    return { success: false, message: '実績報告の取得に失敗しました: ' + error.toString() };
  }
}

/**
 * 実績報告の詳細を取得（IDで指定）
 * @param {string} reportId - 実績報告のID（UUID）または行番号（後方互換性のため）
 * @param {string} sessionId - セッションID（オプション）
 * @return {Object} 実績報告詳細データ
 */
function getPerformanceReportById(reportId, sessionId = null) {
  // セッションIDから直接ユーザー情報を取得（current_sessionに依存しない）
  let user = null;
  if (sessionId) {
    try {
      user = getSession(sessionId);
    } catch (error) {
      console.error('セッション取得エラー:', error);
    }
  }
  // フォールバック: current_sessionから取得
  if (!user) {
    try {
      user = getCurrentUser();
    } catch (error) {
      console.error('getCurrentUser エラー:', error);
    }
  }
  // セッションがなくてもレポートIDがあれば閲覧を許可（UUIDで推測困難）
  
  const staffName = user ? (user.staffName || user.userId || '') : '';
  
  if (!reportId) {
    return { success: false, message: '実績報告IDが指定されていません' };
  }
  
  try {
    const reportSheet = getPerformanceReportSheet();
    const salesCustomerSheet = getSalesCustomerSheet();
    
    // 早期リターン: 範囲チェック
    const reportLastRow = reportSheet.getLastRow();
    const reportLastCol = reportSheet.getLastColumn();
    if (reportLastRow === 0 || reportLastCol === 0) {
      return { success: false, message: '実績報告が見つかりません' };
    }
    
    // 必要な範囲のみ取得（最適化）
    const reportData = reportSheet.getRange(1, 1, reportLastRow, reportLastCol).getValues();
    const salesLastRow = salesCustomerSheet.getLastRow();
    const salesLastCol = salesCustomerSheet.getLastColumn();
    const salesData = (salesLastRow > 0 && salesLastCol > 0) 
      ? salesCustomerSheet.getRange(1, 1, salesLastRow, salesLastCol).getValues()
      : [];
    
    if (!reportData || reportData.length <= 1) {
      return { success: false, message: '実績報告が見つかりません' };
    }
    
    const headers = reportData[0];
    const idIdx = headers.indexOf('ID');
    const reportIdStr = String(reportId).trim();
    
    // IDで検索
    let targetRow = null;
    let targetRowIndex = -1;
    
    console.log('getPerformanceReportById: 検索開始 - reportIdStr:', reportIdStr, 'idIdx:', idIdx);
    
    for (let i = 1; i < reportData.length; i++) {
      const row = reportData[i];
      const rowId = idIdx >= 0 && row[idIdx] ? String(row[idIdx]).trim() : '';
      const rowIndex = i - 1; // 0-indexed の行番号
      
      // IDで一致するか、または行番号で一致するか（後方互換性）
      // IDが空の場合も行番号で検索
      const matchById = rowId && rowId === reportIdStr;
      const matchByIndex = String(rowIndex) === reportIdStr;
      
      if (matchById || matchByIndex) {
        console.log('getPerformanceReportById: 一致発見 - row:', i, 'rowId:', rowId, 'matchById:', matchById, 'matchByIndex:', matchByIndex);
        targetRow = row;
        targetRowIndex = i;
        break;
      }
    }
    
    if (!targetRow) {
      console.log('getPerformanceReportById: 実績報告が見つかりません - reportIdStr:', reportIdStr);
      return { success: false, message: '実績報告が見つかりません（ID: ' + reportIdStr + '）' };
    }
    
    const rowStaffName = String(targetRow[headers.indexOf('スタッフ名')] || '').trim();
    const rowStoreName = String(targetRow[headers.indexOf('所属店舗')] || '').trim();

    // 権限チェック：セッションがある場合のみ所有者確認（セッション切れでも閲覧可能にする）
    if (user && staffName) {
      // 経営オーナーは全てのレポートを閲覧可能
      if (user.role === '経営オーナー') {
        // 許可
      } else if (user.role === 'FCオーナー') {
        // FCオーナーは自店舗の報告のみ閲覧可能
        if (!canAccessStore(rowStoreName)) {
          return { success: false, message: 'この実績報告へのアクセス権限がありません' };
        }
      } else {
        // スタッフは自分の報告のみ閲覧可能
        const staffNameVariants = [
          staffName.trim(),
          staffName.replace(/　/g, ' ').trim(),
          staffName.replace(/ /g, '　').trim()
        ];
        const rowStaffTrimmed = rowStaffName.trim();
        const isMatch = staffNameVariants.some(variant => {
          const variantTrimmed = variant.trim();
          return rowStaffTrimmed === variantTrimmed ||
                 rowStaffTrimmed.includes(variantTrimmed) ||
                 variantTrimmed.includes(rowStaffTrimmed);
        });
        if (!isMatch) {
          return { success: false, message: 'この実績報告へのアクセス権限がありません' };
        }
      }
    }
    
    // 実績報告データを構築
    const reportIdValue = idIdx >= 0 && targetRow[idIdx] ? String(targetRow[idIdx]).trim() : reportIdStr;
    const report = {
      id: reportIdValue,
      yearMonth: String(targetRow[headers.indexOf('報告年月')] || ''),
      staffName: rowStaffName,
      storeName: String(targetRow[headers.indexOf('所属店舗')] || ''),
      sessionCount: Number(targetRow[headers.indexOf('セッション本数')] || 0),
      attendanceCount: Number(targetRow[headers.indexOf('出勤回数')] || 0),
      endNextMonthCount: Number(targetRow[headers.indexOf('翌月終了予定者数')] || 0),
      currentCustomerCount: Number(targetRow[headers.indexOf('現在の担当人数')] || 0),
      desiredCustomerCount: Number(targetRow[headers.indexOf('追加希望担当人数')] || 0),
      salesCustomerCount: Number(targetRow[headers.indexOf('売上顧客数')] || 0),
      reportDateTime: String(targetRow[headers.indexOf('報告日時')] || ''),
      salesCustomers: []
    };
    
    // 売上顧客データを取得
    const salesHeaders = salesData[0];
    const salesYearMonthIdx = salesHeaders.indexOf('報告年月');
    const salesStaffNameIdx = salesHeaders.indexOf('スタッフ名');
    const salesCustomerNameIdx = salesHeaders.indexOf('顧客名');
    const salesTypeIdx = salesHeaders.indexOf('種別');
    const salesStoreNameIdx = salesHeaders.indexOf('店舗名');
    
    for (let j = 1; j < salesData.length; j++) {
      const salesRow = salesData[j];
      const salesRowYearMonth = String(salesRow[salesYearMonthIdx] || '').trim();
      const salesRowStaffName = String(salesRow[salesStaffNameIdx] || '').trim();
      
      if (salesRowYearMonth === report.yearMonth && salesRowStaffName === rowStaffName) {
        report.salesCustomers.push({
          name: String(salesRow[salesCustomerNameIdx] || ''),
          type: String(salesRow[salesTypeIdx] || ''),
          storeName: String(salesRow[salesStoreNameIdx] || '')
        });
      }
    }
    
    return { success: true, data: report };
  } catch (error) {
    console.error('getPerformanceReportById error:', error);
    return { success: false, message: '実績報告の取得に失敗しました: ' + error.toString() };
  }
}

/**
 * 実績報告を更新（編集保存）
 * @param {string} reportId - 実績報告のID
 * @param {Object} updateData - 更新データ
 * @param {string} sessionId - セッションID
 * @return {Object} 更新結果
 */
function updatePerformanceReport(reportId, updateData, sessionId) {
  try {
    // セッションからユーザー情報を取得
    let user = null;
    if (sessionId) {
      user = getSession(sessionId);
    }
    if (!user) {
      user = getCurrentUser();
    }
    if (!user) {
      return { success: false, message: 'ログインが必要です' };
    }

    const staffName = user.staffName || user.userId || '';
    const reportSheet = getPerformanceReportSheet();
    const reportLastRow = reportSheet.getLastRow();
    const reportLastCol = reportSheet.getLastColumn();

    if (reportLastRow <= 1) {
      return { success: false, message: '実績報告データがありません' };
    }

    const reportData = reportSheet.getRange(1, 1, reportLastRow, reportLastCol).getValues();
    const headers = reportData[0];
    const idIdx = headers.indexOf('ID');
    const reportIdStr = String(reportId).trim();

    // 対象行を検索
    let targetRowNum = -1;
    for (let i = 1; i < reportData.length; i++) {
      const row = reportData[i];
      const rowId = idIdx >= 0 && row[idIdx] ? String(row[idIdx]).trim() : '';
      const rowIndex = i - 1;
      if ((rowId && rowId === reportIdStr) || String(rowIndex) === reportIdStr) {
        // 権限チェック（過去データ・担当者変更にも対応するため名前表記のゆらぎを許容）
        const rowStaffName = String(row[headers.indexOf('スタッフ名')] || '').trim();
        const staffNameVariants = [
          staffName.trim(),
          staffName.replace(/　/g, ' ').trim(),
          staffName.replace(/ /g, '　').trim()
        ];
        const rowStaffTrimmed = rowStaffName.trim();
        const nameMatch = staffNameVariants.some(function(variant) {
          const v = variant.trim();
          return rowStaffTrimmed === v || rowStaffTrimmed.includes(v) || v.includes(rowStaffTrimmed);
        });
        // 経営オーナーは全て編集可能。スタッフは自分の報告または名前表記が一致する報告を編集可能
        if (!nameMatch && user.role !== '経営オーナー') {
          return { success: false, message: 'この実績報告の編集権限がありません' };
        }
        targetRowNum = i + 1; // シートの行番号（1-indexed）
        break;
      }
    }

    if (targetRowNum < 0) {
      return { success: false, message: '実績報告が見つかりません' };
    }

    // 更新可能なフィールドを反映
    const updateFields = {
      'セッション本数': updateData.sessionCount,
      '出勤回数': updateData.attendanceCount,
      '翌月終了予定者数': updateData.endNextMonthCount,
      '現在の担当人数': updateData.currentCustomerCount,
      '追加希望担当人数': updateData.desiredCustomerCount
    };

    for (const [headerName, value] of Object.entries(updateFields)) {
      if (value !== undefined && value !== null) {
        const colIdx = headers.indexOf(headerName);
        if (colIdx >= 0) {
          reportSheet.getRange(targetRowNum, colIdx + 1).setValue(Number(value) || 0);
        }
      }
    }

    console.log('実績報告更新完了 - reportId:', reportIdStr, 'row:', targetRowNum);
    return { success: true, message: '実績報告を更新しました' };
  } catch (error) {
    console.error('updatePerformanceReport error:', error);
    return { success: false, message: '更新に失敗しました: ' + error.toString() };
  }
}

/**
 * 顧客マスタシートを取得（存在しない場合は作成）
 * @return {Sheet} 顧客マスタシート
 */
function getCustomerMasterSheet() {
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  let sheet = ss.getSheetByName(CUSTOMER_MASTER_SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(CUSTOMER_MASTER_SHEET_NAME);
    // ヘッダー行を設定（必要に応じて調整）
    const headers = [
      '店舗', '名前', 'ステータス', '入会日・継続日・体験日', '性別', '年齢', '生年月日', 
      '住所', '家族構成', '入会目的・継続理由', '何で知ったか', 'ポスティング日', 
      'お問い合わせ日時', 'エリア', '封筒', '決めて（未入会の場合は不成約理由）', 
      '担当者', '担当', 'コース', '支払い状況', '補足', '売上', '値段/回'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  
  return sheet;
}

/**
 * 顧客登録シートを取得（存在しない場合は作成）
 * @return {Sheet} 顧客登録シート
 */
function getCustomerRegisterSheet() {
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  const sheetName = '顧客登録';
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    // ヘッダー行を設定（A列に会員IDを追加）
    const headers = [
      '会員ID', '店舗', '担当者', '体験/入会', '日付', '名前', '性別', '年齢', '生年月日', 
      '住所', '家族構成', '入会目的・継続理由', '何で知ったか', 'ポスティング日', 
      'お問い合わせ日時', 'エリア', '封筒', '決めて（未入会の場合は不成約理由）', 
      '担当', 'コース', '支払い状況', '補足', '売上', '値段/回', '登録日時'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    
    console.log('顧客登録シートを新規作成しました');
  }
  
  return sheet;
}

/**
 * 店舗名から店舗コードを取得
 * @param {string} storeName - 店舗名
 * @return {string} 店舗コード
 */
function getStoreCode(storeName) {
  const storeCodeMap = {
    '幡ヶ谷': 'HTY',
    '曙橋': 'AB',
    '高円寺': 'KG',
    '西荻窪': 'NG',
    '江古田': 'ECD',
    '駒沢': 'KMZ'
  };
  
  // 完全一致
  if (storeCodeMap[storeName]) {
    return storeCodeMap[storeName];
  }
  
  // 部分一致
  for (const [name, code] of Object.entries(storeCodeMap)) {
    if (storeName.includes(name)) {
      return code;
    }
  }
  
  // 該当なしの場合は店舗名の最初の3文字を大文字で
  return storeName.substring(0, 3).toUpperCase();
}

/**
 * 次の顧客IDを生成
 * @param {string} storeCode - 店舗コード
 * @return {string} 顧客ID（例: HTY-001）
 */
function generateCustomerId(storeCode) {
  const sheet = getCustomerRegisterSheet();
  const data = sheet.getDataRange().getValues();
  
  // 同じ店舗コードの最大番号を探す
  let maxNumber = 0;
  const prefix = storeCode + '-';
  
  data.slice(1).forEach(row => {
    const id = String(row[0] || '').trim();
    if (id.startsWith(prefix)) {
      const numPart = parseInt(id.substring(prefix.length), 10);
      if (!isNaN(numPart) && numPart > maxNumber) {
        maxNumber = numPart;
      }
    }
  });
  
  // 次の番号を3桁でフォーマット
  const nextNumber = String(maxNumber + 1).padStart(3, '0');
  return prefix + nextNumber;
}

/**
 * 顧客マスタに顧客を登録
 * 実際のシート構造: 店舗, 担当者, (空), 入会日・継続日・体験日, 名前, 性別, 年齢, 生年月日, 住所, 家族構成, 入会目的・継続理由, 何で知ったか, ポスティング日, お問い合わせ日時, エリア, 封筒, 決めて（未入会の場合は不成約理由）, 担当, コース, 支払い状況, 補足, 売上, 値段/回
 * @param {Object} customerData - 顧客データオブジェクト
 * @param {string} sessionId - セッションID（オプション）
 * @return {Object} 登録結果
 */
function registerCustomerToMaster(customerData, sessionId = null) {
  try {
    console.log('registerCustomerToMaster呼び出し開始');
    console.log('customerData:', JSON.stringify(customerData));
    
    // セッションIDが指定されている場合はセッションを設定
    if (sessionId) {
      try {
        const setResult = setSession(sessionId);
        if (!setResult || !setResult.success) {
          console.warn('セッション設定失敗:', setResult);
        }
      } catch (error) {
        console.error('セッション設定エラー:', error);
        // セッション設定エラーがあっても続行
      }
    }
    
    // 権限チェック
    const user = getCurrentUser();
    if (!user) {
      console.error('ユーザーが取得できませんでした');
      return { success: false, message: 'ログインが必要です' };
    }
    
    console.log('ユーザー確認完了:', user.userId || user.email);
  
  try {
    // 新しい「顧客登録」シートを使用
    const sheet = getCustomerRegisterSheet();
    
    // 登録日時を追加
    const now = new Date();
    const registeredAt = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    
    // 店舗コードを取得してIDを生成
    const storeCode = getStoreCode(customerData.store || '');
    const customerId = generateCustomerId(storeCode);
    
    console.log('生成されたID:', customerId);
    
    // シート構造に合わせてデータを作成
    // A:ID, B:店舗, C:担当者, D:体験/入会, E:日付, F:名前, G:性別, H:年齢, I:生年月日, 
    // J:住所, K:家族構成, L:入会目的・継続理由, M:何で知ったか, N:ポスティング日, O:お問い合わせ日時, 
    // P:エリア, Q:封筒, R:決めて（未入会の場合は不成約理由）, S:担当, T:コース, U:支払い状況, V:補足, W:売上, X:値段/回, Y:登録日時
    const row = [
      customerId,                                  // A: ID
      customerData.store || '',                    // B: 店舗
      customerData.staffPerson || '',              // C: 担当者
      customerData.status || '',                   // D: 体験/入会
      customerData.date || '',                     // E: 日付
      customerData.name || '',                     // F: 名前
      customerData.gender || '',                   // G: 性別
      customerData.age || '',                      // H: 年齢
      customerData.birthday || '',                 // I: 生年月日
      customerData.address || '',                  // J: 住所
      customerData.family || '',                   // K: 家族構成
      customerData.purpose || '',                  // L: 入会目的・継続理由
      customerData.source || '',                   // M: 何で知ったか
      customerData.postingDate || '',              // N: ポスティング日
      customerData.inquiryDateTime || '',          // O: お問い合わせ日時
      customerData.area || '',                     // P: エリア
      customerData.envelope || '',                 // Q: 封筒
      customerData.decision || '',                 // R: 決めて（未入会の場合は不成約理由）
      customerData.staff || '',                    // S: 担当
      customerData.course || '',                   // T: コース
      customerData.paymentStatus || '',            // U: 支払い状況
      customerData.supplement || '',               // V: 補足
      customerData.sales || '',                    // W: 売上
      customerData.unitPrice || '',                // X: 値段/回
      registeredAt                                 // Y: 登録日時
    ];
    
    console.log('データ行を構築完了。列数:', row.length);
    console.log('データ内容:', JSON.stringify(row));
    
    // データを追加
    try {
      sheet.appendRow(row);
      console.log('データ行を追加しました');
    } catch (appendError) {
      console.error('データ追加エラー:', appendError);
      throw new Error('データの追加に失敗しました: ' + appendError.toString());
    }
    
    // スプレッドシートの更新を確実にするため、flushを実行
    try {
      SpreadsheetApp.flush();
      console.log('スプレッドシートをフラッシュしました');
    } catch (flushError) {
      console.warn('フラッシュエラー（無視）:', flushError);
      // フラッシュエラーは無視（データは追加されている可能性が高い）
    }
    
    const customerName = customerData.name || customerData.名前 || '';
    console.log('顧客マスタに登録しました:', customerName);
    
    return {
      success: true,
      message: '顧客データを登録しました',
      customerName: customerName,
      customerId: customerId
    };
  } catch (innerError) {
    console.error('registerCustomerToMaster内側エラー:', innerError);
    console.error('エラースタック:', innerError.stack);
    throw innerError; // 外側のcatchで処理
  }
  } catch (error) {
    console.error('registerCustomerToMaster外側エラー:', error);
    console.error('エラータイプ:', typeof error);
    console.error('エラーメッセージ:', error.message || error.toString());
    console.error('エラースタック:', error.stack);
    
    // エラーメッセージを構築
    let errorMessage = '登録に失敗しました';
    if (error && error.message) {
      errorMessage += ': ' + error.message;
    } else if (error && error.toString) {
      errorMessage += ': ' + error.toString();
    } else {
      errorMessage += ': 不明なエラー';
    }
    
    return {
      success: false,
      message: errorMessage,
      error: error.toString(),
      errorType: typeof error
    };
  }
}

/**
 * 登録済み顧客一覧を取得
 * @param {string} sessionId - セッションID（オプション）
 * @return {Array} 顧客データの配列
 */
function getRegisteredCustomers(sessionId = null) {
  // セッションIDが指定されている場合はセッションを設定
  if (sessionId) {
    try {
      setSession(sessionId);
    } catch (error) {
      console.error('セッション設定エラー:', error);
    }
  }
  
  // ログインユーザーのアクセス可能店舗を取得
  const user = getCurrentUser();
  const accessibleStores = getAccessibleStores();
  const isAdmin = user && user.role === '経営オーナー';
  
  try {
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const sheet = ss.getSheetByName('顧客登録');
    
    if (!sheet) {
      console.log('顧客登録シートが見つかりません');
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    
    if (!data || data.length <= 1) {
      return [];
    }
    
    const customers = [];
    
    // ヘッダー行をスキップしてデータを取得（全会員を表示・当月フィルタなし）
    // A:ID, B:店舗, C:担当者, D:体験/入会, E:日付, F:名前, G:性別, H:年齢, I:生年月日, 
    // J:住所, K:家族構成, L:入会目的・継続理由, M:何で知ったか, N:ポスティング日, O:お問い合わせ日時, 
    // P:エリア, Q:封筒, R:決めて, S:担当, T:コース, U:支払い状況, V:補足, W:売上, X:値段/回, Y:登録日時
    data.slice(1).forEach((row, index) => {
      const name = String(row[5] || '').trim();
      if (!name) return; // 名前がない行はスキップ
      
      // アクセス可能店舗でフィルタリング（経営オーナー以外）
      // 完全一致に加え、店舗名の部分一致も許可（例: アクセスに「駒沢」があれば「駒沢大学」の顧客も表示）
      const customerStore = String(row[1] || '').trim();
      if (!isAdmin && accessibleStores.length > 0 && customerStore) {
        const canAccess = accessibleStores.some(store => {
          const s = String(store).trim();
          return s === customerStore || customerStore.indexOf(s) >= 0 || s.indexOf(customerStore) >= 0;
        });
        if (!canAccess) return; // アクセス権のない店舗の顧客はスキップ
      }
      
      const registeredAtStr = String(row[24] || '').trim();
      // 登録日時はあれば表示用に使用（空でも行は表示する）
      
      customers.push({
        rowIndex: index + 2, // シートの行番号（1-indexed、ヘッダー行を除く）
        id: String(row[0] || '').trim(),
        store: String(row[1] || '').trim(),
        staffPerson: String(row[2] || '').trim(),
        status: String(row[3] || '').trim(),
        date: formatDateValue(row[4]),
        name: name,
        gender: String(row[6] || '').trim(),
        age: String(row[7] || '').trim(),
        birthday: formatDateValue(row[8]),
        address: String(row[9] || '').trim(),
        family: String(row[10] || '').trim(),
        purpose: String(row[11] || '').trim(),
        source: String(row[12] || '').trim(),
        postingDate: formatDateValue(row[13]),
        inquiryDateTime: String(row[14] || '').trim(),
        area: String(row[15] || '').trim(),
        envelope: String(row[16] || '').trim(),
        decision: String(row[17] || '').trim(),
        staff: String(row[18] || '').trim(),
        course: String(row[19] || '').trim(),
        paymentStatus: String(row[20] || '').trim(),
        supplement: String(row[21] || '').trim(),
        sales: String(row[22] || '').trim(),
        unitPrice: String(row[23] || '').trim(),
        registeredAt: registeredAtStr
      });
    });
    
    console.log(`getRegisteredCustomers: ${customers.length}件の顧客を取得（顧客登録シート）`);
    
    // 継続シートからも顧客を取得
    const continueSheet = ss.getSheetByName(CUSTOMER_CONTINUE_SHEET_NAME);
    if (continueSheet) {
      const continueData = continueSheet.getDataRange().getValues();
      if (continueData && continueData.length > 1) {
        const continueHeaders = continueData[0];
        const nameIdx = continueHeaders.indexOf('名前');
        const storeIdx = continueHeaders.indexOf('店舗名');
        const dateIdx = continueHeaders.indexOf('継続日') >= 0 ? continueHeaders.indexOf('継続日') : continueHeaders.indexOf('日付');
        const staffIdx = continueHeaders.indexOf('担当');
        const courseIdx = continueHeaders.indexOf('コース');
        const paymentIdx = continueHeaders.indexOf('支払い状況');
        const salesIdx = continueHeaders.indexOf('売上');
        
        continueData.slice(1).forEach((row, index) => {
          const name = nameIdx >= 0 ? String(row[nameIdx] || '').trim() : '';
          if (!name) return;
          
          // アクセス可能店舗でフィルタリング（経営オーナー以外）
          const contStore = storeIdx >= 0 ? String(row[storeIdx] || '').trim() : '';
          if (!isAdmin && accessibleStores.length > 0 && contStore && !accessibleStores.includes(contStore)) {
            return;
          }
          
          const dateValue = dateIdx >= 0 ? row[dateIdx] : null;
          const registeredAtStr = dateValue ? formatDateValue(dateValue) : '';
          
          // 登録日時が当月かどうかをチェック
          let isCurrentMonth = false;
          if (registeredAtStr) {
            try {
              let registeredDate = null;
              if (registeredAtStr.includes('-')) {
                registeredDate = new Date(registeredAtStr.replace(/-/g, '/'));
              } else if (registeredAtStr.includes('/')) {
                registeredDate = new Date(registeredAtStr);
              }
              
              if (registeredDate && !isNaN(registeredDate.getTime())) {
                isCurrentMonth = registeredDate >= startOfMonth && registeredDate <= endOfMonth;
              }
            } catch (e) {
              console.warn('継続日時の解析エラー:', registeredAtStr, e);
            }
          }
          
          // 当月の継続のみ追加
          if (isCurrentMonth) {
            // 既に同じ顧客が追加されているかチェック（重複防止）
            const existing = customers.find(c => c.name === name && c.store === String(row[storeIdx] || '').trim());
            if (!existing) {
              customers.push({
                rowIndex: index + 2,
                id: `CONTINUE_${index + 1}`,
                store: storeIdx >= 0 ? String(row[storeIdx] || '').trim() : '',
                staffPerson: staffIdx >= 0 ? String(row[staffIdx] || '').trim() : '',
                status: '継続',
                date: registeredAtStr,
                name: name,
                gender: '',
                age: '',
                birthday: '',
                address: '',
                family: '',
                purpose: '',
                source: '',
                postingDate: '',
                inquiryDateTime: '',
                area: '',
                envelope: '',
                decision: '',
                staff: staffIdx >= 0 ? String(row[staffIdx] || '').trim() : '',
                course: courseIdx >= 0 ? String(row[courseIdx] || '').trim() : '',
                paymentStatus: paymentIdx >= 0 ? String(row[paymentIdx] || '').trim() : '',
                supplement: '',
                sales: salesIdx >= 0 ? String(row[salesIdx] || '').trim() : '',
                unitPrice: '',
                registeredAt: registeredAtStr
              });
            }
          }
        });
      }
    }
    
    console.log(`getRegisteredCustomers: 合計${customers.length}件の顧客を取得（顧客マスタ + 継続）`);
    
    // 登録日時の降順でソート（新しい順）
    customers.sort((a, b) => {
      if (!a.registeredAt) return 1;
      if (!b.registeredAt) return -1;
      return b.registeredAt.localeCompare(a.registeredAt);
    });
    
    return customers;
  } catch (error) {
    console.error('getRegisteredCustomers error:', error);
    console.error('エラースタック:', error.stack);
    return [];
  }
}

/**
 * 行番号で顧客詳細を取得
 * @param {number} rowIndex - 行番号
 * @param {string} sessionId - セッションID（オプション）
 * @return {Object} 顧客データ
 */
function getCustomerByRowIndex(rowIndex, sessionId = null) {
  // セッションIDが指定されている場合はセッションを設定
  if (sessionId) {
    try {
      setSession(sessionId);
    } catch (error) {
      console.error('セッション設定エラー:', error);
    }
  }
  
  try {
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const sheet = ss.getSheetByName('顧客登録');
    
    if (!sheet) {
      console.log('顧客登録シートが見つかりません');
      return null;
    }
    
    const lastRow = sheet.getLastRow();
    if (rowIndex < 2 || rowIndex > lastRow) {
      console.log('無効な行番号:', rowIndex);
      return null;
    }
    
    const row = sheet.getRange(rowIndex, 1, 1, 25).getValues()[0];
    
    const customer = {
      rowIndex: rowIndex,
      id: String(row[0] || '').trim(),
      store: String(row[1] || '').trim(),
      staffPerson: String(row[2] || '').trim(),
      status: String(row[3] || '').trim(),
      date: formatDateValue(row[4]),
      name: String(row[5] || '').trim(),
      gender: String(row[6] || '').trim(),
      age: String(row[7] || '').trim(),
      birthday: formatDateValue(row[8]),
      address: String(row[9] || '').trim(),
      family: String(row[10] || '').trim(),
      purpose: String(row[11] || '').trim(),
      source: String(row[12] || '').trim(),
      postingDate: formatDateValue(row[13]),
      inquiryDateTime: String(row[14] || '').trim(),
      area: String(row[15] || '').trim(),
      envelope: String(row[16] || '').trim(),
      decision: String(row[17] || '').trim(),
      staff: String(row[18] || '').trim(),
      course: String(row[19] || '').trim(),
      paymentStatus: String(row[20] || '').trim(),
      supplement: String(row[21] || '').trim(),
      sales: String(row[22] || '').trim(),
      unitPrice: String(row[23] || '').trim(),
      registeredAt: String(row[24] || '').trim()
    };
    
    console.log(`getCustomerByRowIndex: 行${rowIndex}の顧客「${customer.name}」(ID: ${customer.id})を取得`);
    
    return customer;
  } catch (error) {
    console.error('getCustomerByRowIndex error:', error);
    console.error('エラースタック:', error.stack);
    return null;
  }
}

/**
 * 日付値をフォーマット
 * @param {*} value - 日付値
 * @return {string} フォーマットされた日付文字列
 */
function formatDateValue(value) {
  if (!value) return '';
  
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return Utilities.formatDate(value, 'Asia/Tokyo', 'yyyy-MM-dd');
  }
  
  return String(value).trim();
}

/**
 * 既存の顧客データにIDを割り振る（一度だけ実行）
 * GASエディタから手動で実行してください
 */
function assignIdsToExistingCustomers() {
  try {
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const sheet = ss.getSheetByName('顧客登録');
    
    if (!sheet) {
      console.log('顧客登録シートが見つかりません');
      return { success: false, message: '顧客登録シートが見つかりません' };
    }
    
    const data = sheet.getDataRange().getValues();
    
    if (!data || data.length <= 1) {
      return { success: false, message: 'データがありません' };
    }
    
    // ヘッダーを確認
    const headers = data[0];
    console.log('ヘッダー:', headers);
    
    // A列がIDかどうか確認
    if (headers[0] !== 'ID' && headers[0] !== '会員ID') {
      // A列にID列を挿入
      sheet.insertColumnBefore(1);
      sheet.getRange(1, 1).setValue('会員ID');
      sheet.getRange(1, 1).setFontWeight('bold');
      console.log('会員ID列を挿入しました');
      
      // データを再取得
      const newData = sheet.getDataRange().getValues();
      
      // 各行にIDを割り振る
      let updatedCount = 0;
      for (let i = 1; i < newData.length; i++) {
        const row = newData[i];
        const storeName = String(row[1] || '').trim(); // B列が店舗名
        const name = String(row[5] || '').trim(); // F列が名前
        
        if (name && !row[0]) { // 名前があってIDがない場合
          const storeCode = getStoreCode(storeName);
          const customerId = generateCustomerIdForRow(sheet, storeCode, i);
          sheet.getRange(i + 1, 1).setValue(customerId);
          console.log(`行${i + 1}: ${name}さんにID「${customerId}」を割り振りました`);
          updatedCount++;
        }
      }
      
      return { success: true, message: `${updatedCount}件の顧客にIDを割り振りました` };
    } else {
      // A列がすでにID列の場合、IDがない行にのみ割り振る
      let updatedCount = 0;
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const existingId = String(row[0] || '').trim();
        const storeName = String(row[1] || '').trim(); // B列が店舗名
        const name = String(row[5] || '').trim(); // F列が名前
        
        if (name && !existingId) { // 名前があってIDがない場合
          const storeCode = getStoreCode(storeName);
          const customerId = generateCustomerIdForRow(sheet, storeCode, i);
          sheet.getRange(i + 1, 1).setValue(customerId);
          console.log(`行${i + 1}: ${name}さんにID「${customerId}」を割り振りました`);
          updatedCount++;
        }
      }
      
      return { success: true, message: `${updatedCount}件の顧客にIDを割り振りました` };
    }
  } catch (error) {
    console.error('assignIdsToExistingCustomers error:', error);
    return { success: false, message: 'エラー: ' + error.toString() };
  }
}

/**
 * 特定の行用にIDを生成（既存データ用）
 */
function generateCustomerIdForRow(sheet, storeCode, currentRowIndex) {
  const data = sheet.getDataRange().getValues();
  
  // 同じ店舗コードの最大番号を探す（現在の行より前の行のみ）
  let maxNumber = 0;
  const prefix = storeCode + '-';
  
  for (let i = 1; i <= currentRowIndex; i++) {
    const id = String(data[i][0] || '').trim();
    if (id.startsWith(prefix)) {
      const numPart = parseInt(id.substring(prefix.length), 10);
      if (!isNaN(numPart) && numPart > maxNumber) {
        maxNumber = numPart;
      }
    }
  }
  
  // 次の番号を3桁でフォーマット
  const nextNumber = String(maxNumber + 1).padStart(3, '0');
  return prefix + nextNumber;
}

/**
 * IDで顧客詳細を取得
 * @param {string} customerId - 顧客ID（例: HTY-001）
 * @param {string} sessionId - セッションID（オプション）
 * @return {Object} 顧客データ
 */
function getCustomerById(customerId, sessionId = null) {
  try {
    console.log('getCustomerById呼び出し:', { customerId, sessionId });
    
    // セッションIDが指定されている場合はセッションを設定
    if (sessionId) {
      try {
        setSession(sessionId);
      } catch (error) {
        console.error('セッション設定エラー:', error);
      }
    }
    
    // 権限チェック（セッションがない場合でも顧客データを取得できるようにする）
    const user = getCurrentUser();
    if (!user) {
      console.log('セッションなしで顧客データを取得します');
      // セッションがない場合でも続行（認証なしでも確認できるようにする）
    }
    
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const sheet = ss.getSheetByName('顧客登録');
    
    if (!sheet) {
      console.error('顧客登録シートが見つかりません');
      return null;
    }
    
    const data = sheet.getDataRange().getValues();
    
    if (!data || data.length <= 1) {
      console.warn('顧客データがありません');
      return null;
    }
    
    console.log('顧客登録シートのデータ行数:', data.length);
    console.log('検索対象の顧客ID:', customerId);
    console.log('検索対象の顧客ID（型）:', typeof customerId);
    
    // IDで検索（A列）
    let foundMatch = false;
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const id = String(row[0] || '').trim();
      
      // デバッグ用：最初の10件のIDをログ出力
      if (i <= 10) {
        console.log(`行${i}のID: "${id}" (検索対象: "${customerId}") 一致: ${id === customerId}`);
      }
      
      // 完全一致チェック
      if (id === customerId) {
        foundMatch = true;
        const customer = {
          rowIndex: i + 1,
          id: id,
          store: String(row[1] || '').trim(),
          staffPerson: String(row[2] || '').trim(),
          status: String(row[3] || '').trim(),
          date: formatDateValue(row[4]),
          name: String(row[5] || '').trim(),
          gender: String(row[6] || '').trim(),
          age: String(row[7] || '').trim(),
          birthday: formatDateValue(row[8]),
          address: String(row[9] || '').trim(),
          family: String(row[10] || '').trim(),
          purpose: String(row[11] || '').trim(),
          source: String(row[12] || '').trim(),
          postingDate: formatDateValue(row[13]),
          inquiryDateTime: String(row[14] || '').trim(),
          area: String(row[15] || '').trim(),
          envelope: String(row[16] || '').trim(),
          decision: String(row[17] || '').trim(),
          staff: String(row[18] || '').trim(),
          course: String(row[19] || '').trim(),
          paymentStatus: String(row[20] || '').trim(),
          supplement: String(row[21] || '').trim(),
          sales: String(row[22] || '').trim(),
          unitPrice: String(row[23] || '').trim(),
          registeredAt: String(row[24] || '').trim()
        };
        
        console.log(`getCustomerById: ID「${customerId}」の顧客「${customer.name}」を取得`);
        
        return customer;
      }
    }
    
    console.log(`getCustomerById: ID「${customerId}」の顧客が見つかりません`);
    console.log(`検索結果: ${foundMatch ? '一致するIDが見つかりました' : '一致するIDが見つかりませんでした'}`);
    console.log(`データ行数: ${data.length - 1}件（ヘッダー除く）`);
    
    // 部分一致や類似IDを探す（デバッグ用）
    const similarIds = [];
    for (let i = 1; i < Math.min(data.length, 50); i++) {
      const row = data[i];
      const id = String(row[0] || '').trim();
      if (id && (id.includes(customerId) || customerId.includes(id))) {
        similarIds.push({ row: i + 1, id: id });
      }
    }
    if (similarIds.length > 0) {
      console.log('類似するIDが見つかりました:', similarIds);
    }
    
    return null;
  } catch (error) {
    console.error('getCustomerById error:', error);
    console.error('エラースタック:', error.stack);
    return null;
  }
}

/**
 * 顧客継続履歴シートを取得（存在しない場合は作成）
 * @return {Sheet} 顧客継続履歴シート
 */
function getCustomerContinueHistorySheet() {
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  const sheetName = '顧客継続履歴';
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    // ヘッダー行を設定
    const headers = [
      '顧客ID', '顧客名', '更新日付', '担当', 'コース', '支払い状況', 
      '補足', '売上', '値段/回', '登録日時'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    
    console.log('顧客継続履歴シートを新規作成しました');
  }
  
  return sheet;
}

/**
 * 顧客の継続情報を保存
 * @param {string} customerId - 顧客ID
 * @param {Object} continueData - 継続情報データ
 * @param {string} sessionId - セッションID（オプション）
 * @return {Object} 保存結果
 */
function saveCustomerContinue(customerId, continueData, sessionId = null) {
  try {
    console.log('saveCustomerContinue呼び出し:', { customerId, continueData });
    
    // セッションIDが指定されている場合はセッションを設定
    if (sessionId) {
      try {
        setSession(sessionId);
      } catch (error) {
        console.error('セッション設定エラー:', error);
      }
    }
    
    // 権限チェック
    const user = getCurrentUser();
    if (!user) {
      return { success: false, message: 'ログインが必要です' };
    }
    
    // 顧客情報を取得
    const customer = getCustomerById(customerId, sessionId);
    if (!customer) {
      return { success: false, message: '顧客が見つかりません' };
    }
    
    // 顧客継続履歴シートを取得
    const sheet = getCustomerContinueHistorySheet();
    
    // 登録日時を追加
    const now = new Date();
    const registeredAt = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    
    // データ行を作成
    // A:顧客ID, B:顧客名, C:更新日付, D:担当, E:コース, F:支払い状況, G:補足, H:売上, I:値段/回, J:登録日時
    const row = [
      customerId,                                    // A: 顧客ID
      customer.name || '',                           // B: 顧客名
      continueData.updateDate || '',                 // C: 更新日付
      continueData.staff || '',                      // D: 担当
      continueData.course || '',                     // E: コース
      continueData.paymentStatus || '',              // F: 支払い状況
      continueData.supplement || '',                 // G: 補足
      continueData.sales || '',                      // H: 売上
      continueData.unitPrice || '',                  // I: 値段/回
      registeredAt                                  // J: 登録日時
    ];
    
    // データを追加
    sheet.appendRow(row);
    SpreadsheetApp.flush();
    
    // 顧客登録シートのステータスを「継続」に更新
    const customerSheet = getCustomerRegisterSheet();
    const data = customerSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const id = String(row[0] || '').trim();
      
      if (id === customerId) {
        // ステータス（D列、インデックス3）を「継続」に更新
        customerSheet.getRange(i + 1, 4).setValue('継続');
        // コース、支払い状況、補足、売上、値段/回も更新
        if (continueData.course) {
          customerSheet.getRange(i + 1, 20).setValue(continueData.course); // T列: コース
        }
        if (continueData.paymentStatus) {
          customerSheet.getRange(i + 1, 21).setValue(continueData.paymentStatus); // U列: 支払い状況
        }
        if (continueData.supplement) {
          customerSheet.getRange(i + 1, 22).setValue(continueData.supplement); // V列: 補足
        }
        if (continueData.sales) {
          customerSheet.getRange(i + 1, 23).setValue(continueData.sales); // W列: 売上
        }
        if (continueData.unitPrice) {
          customerSheet.getRange(i + 1, 24).setValue(continueData.unitPrice); // X列: 値段/回
        }
        break;
      }
    }
    
    SpreadsheetApp.flush();
    
    console.log('顧客継続情報を保存しました:', customerId);
    
    return {
      success: true,
      message: '継続情報を登録しました',
      customerId: customerId
    };
  } catch (error) {
    console.error('saveCustomerContinue error:', error);
    console.error('エラースタック:', error.stack);
    return {
      success: false,
      message: '継続情報の保存に失敗しました: ' + (error.message || error.toString())
    };
  }
}

/**
 * 顧客の継続履歴を取得
 * @param {string} customerId - 顧客ID
 * @param {string} sessionId - セッションID（オプション）
 * @return {Array} 継続履歴の配列
 */
function getCustomerContinueHistory(customerId, sessionId = null) {
  try {
    console.log('getCustomerContinueHistory呼び出し:', { customerId });
    
    // セッションIDが指定されている場合はセッションを設定
    if (sessionId) {
      try {
        setSession(sessionId);
      } catch (error) {
        console.error('セッション設定エラー:', error);
      }
    }
    
    // 権限チェック（セッションがない場合でも継続履歴を取得できるようにする）
    const user = getCurrentUser();
    if (!user) {
      console.log('セッションなしで継続履歴を取得します');
      // セッションがない場合でも続行（認証なしでも確認できるようにする）
    }
    
    const sheet = getCustomerContinueHistorySheet();
    const data = sheet.getDataRange().getValues();
    
    if (!data || data.length <= 1) {
      return [];
    }
    
    const history = [];
    
    // ヘッダー行をスキップしてデータを取得
    // A:顧客ID, B:顧客名, C:更新日付, D:担当, E:コース, F:支払い状況, G:補足, H:売上, I:値段/回, J:登録日時
    data.slice(1).forEach((row) => {
      const rowCustomerId = String(row[0] || '').trim();
      
      if (rowCustomerId === customerId) {
        history.push({
          customerId: rowCustomerId,
          customerName: String(row[1] || '').trim(),
          updateDate: formatDateValue(row[2]),
          staff: String(row[3] || '').trim(),
          course: String(row[4] || '').trim(),
          paymentStatus: String(row[5] || '').trim(),
          supplement: String(row[6] || '').trim(),
          sales: String(row[7] || '').trim(),
          unitPrice: String(row[8] || '').trim(),
          registeredAt: String(row[9] || '').trim()
        });
      }
    });
    
    // 登録日時の降順でソート（新しい順）
    history.sort((a, b) => {
      if (!a.registeredAt) return 1;
      if (!b.registeredAt) return -1;
      return b.registeredAt.localeCompare(a.registeredAt);
    });
    
    console.log(`getCustomerContinueHistory: ${history.length}件の継続履歴を取得しました（顧客ID: ${customerId}）`);
    
    return history;
  } catch (error) {
    console.error('getCustomerContinueHistory error:', error);
    console.error('エラースタック:', error.stack);
    return [];
  }
}

// ===== 口コミ取得数 関連 =====

/**
 * 口コミシートを取得（なければ作成）
 * @return {Sheet} 口コミシート
 */
function getReviewSheet() {
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  let sheet = ss.getSheetByName('口コミ取得数');
  if (!sheet) {
    sheet = ss.insertSheet('口コミ取得数');
    // ヘッダーを作成
    sheet.getRange(1, 1, 1, 7).setValues([
      ['顧客ID', '顧客名', '年月', '口コミ数', '画像URL', '登録者', '登録日時']
    ]);
    // ヘッダーの書式設定
    sheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#667eea').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * 口コミ情報を保存
 * @param {Object} reviewData - 口コミデータ { customerId, customerName, yearMonth, reviewCount, imageBase64 }
 * @param {string} sessionId - セッションID
 * @return {Object} 結果
 */
function saveCustomerReview(reviewData, sessionId) {
  try {
    if (sessionId) {
      try { setSession(sessionId); } catch (e) {}
    }
    
    const user = getCurrentUser();
    if (!user) {
      return { success: false, message: 'ログインが必要です' };
    }
    
    if (!reviewData || !reviewData.customerId) {
      return { success: false, message: '顧客IDが指定されていません' };
    }
    
    const sheet = getReviewSheet();
    const now = new Date();
    const registeredAt = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
    
    // 画像をGoogle Driveに保存（Base64データがある場合）
    let imageUrl = '';
    if (reviewData.imageBase64 && reviewData.imageBase64.startsWith('data:image')) {
      try {
        // data:image/png;base64,... の形式からデータ部分を抽出
        const parts = reviewData.imageBase64.split(',');
        const mimeType = parts[0].match(/:(.*?);/)[1];
        const extension = mimeType.split('/')[1] || 'png';
        const base64Data = parts[1];
        const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, 
          `review_${reviewData.customerId}_${Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMMdd_HHmmss')}.${extension}`);
        
        // 指定のGoogle Driveフォルダに保存
        let folder;
        try {
          folder = DriveApp.getFolderById('1o1u7-wN7hUU9SZvfqXuZqXsIUfMyizYT');
        } catch (folderErr) {
          console.error('口コミフォルダアクセスエラー（フォルダID確認要）:', folderErr);
          // フォールバック: ルートにフォルダ作成
          const folders = DriveApp.getFoldersByName('口コミ画像');
          folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('口コミ画像');
        }
        
        const file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        imageUrl = file.getUrl();
        
        console.log('口コミ画像を保存しました:', imageUrl);
      } catch (imgError) {
        console.error('画像保存エラー:', imgError);
        // 画像保存に失敗しても口コミデータは保存する
      }
    }
    
    // スプレッドシートに保存
    const row = [
      reviewData.customerId,
      reviewData.customerName || '',
      reviewData.yearMonth || '',
      reviewData.reviewCount || 0,
      imageUrl,
      user.staffName || user.userId || '',
      registeredAt
    ];
    
    sheet.appendRow(row);
    
    console.log('口コミ情報を保存しました:', reviewData.customerId, reviewData.yearMonth, reviewData.reviewCount);
    
    return { success: true, message: '口コミ情報を保存しました' };
  } catch (error) {
    console.error('saveCustomerReview error:', error);
    return { success: false, message: '口コミ情報の保存に失敗しました: ' + error.toString() };
  }
}

/**
 * 顧客の口コミ履歴を取得
 * @param {string} customerId - 顧客ID
 * @param {string} sessionId - セッションID
 * @return {Array} 口コミ履歴データ
 */
function getCustomerReviews(customerId, sessionId) {
  try {
    if (sessionId) {
      try { setSession(sessionId); } catch (e) {}
    }
    
    if (!customerId) return [];
    
    const sheet = getReviewSheet();
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) return []; // ヘッダーのみ
    
    const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    const reviews = [];
    
    data.forEach(function(row) {
      const rowCustomerId = String(row[0] || '').trim();
      if (rowCustomerId === String(customerId).trim()) {
        reviews.push({
          customerId: rowCustomerId,
          customerName: String(row[1] || '').trim(),
          yearMonth: String(row[2] || '').trim(),
          reviewCount: row[3] || 0,
          imageUrl: String(row[4] || '').trim(),
          registeredBy: String(row[5] || '').trim(),
          registeredAt: String(row[6] || '').trim()
        });
      }
    });
    
    // 登録日時の降順でソート
    reviews.sort(function(a, b) {
      return (b.registeredAt || '').localeCompare(a.registeredAt || '');
    });
    
    return reviews;
  } catch (error) {
    console.error('getCustomerReviews error:', error);
    return [];
  }
}

/**
 * 顧客の口コミ件数を一括取得（顧客名をキーにした集計）
 * @param {string} yearMonth - 年月（例: "2026年1月"）
 * @param {string} sessionId - セッションID
 * @return {Object} { "顧客名": 件数, ... }
 */
function getReviewCountsByYearMonth(yearMonth, sessionId) {
  try {
    if (sessionId) {
      try { setSession(sessionId); } catch (e) {}
    }
    
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const sheet = ss.getSheetByName('口コミ取得数');
    if (!sheet) return {};
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return {};
    
    const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    const counts = {};
    
    data.forEach(function(row) {
      const customerName = String(row[1] || '').trim();
      const rowYearMonth = String(row[2] || '').trim();
      const reviewCount = parseInt(row[3]) || 0;
      
      if (!customerName) return;
      
      // 年月フィルタ（指定されている場合）
      if (yearMonth && rowYearMonth !== yearMonth) return;
      
      if (!counts[customerName]) {
        counts[customerName] = 0;
      }
      counts[customerName] += reviewCount;
    });
    
    return counts;
  } catch (error) {
    console.error('getReviewCountsByYearMonth error:', error);
    return {};
  }
}

// ========================================
// 報酬計算関連関数
// ========================================

/**
 * スタッフマスタから報酬制度情報を取得
 * スプレッドシートの「スタッフマスタ」シート:
 *   A:店舗名, B:オーナー/スタッフ, C:スタッフ名, D:報酬制度(旧制度/新制度),
 *   E:パーセンテージ, F:新制度の金額, G:(空き), H:オーナー
 * @param {string} sessionId - セッションID
 * @return {Array} スタッフの報酬制度情報一覧
 */
function getStaffRewardMaster(sessionId) {
  try {
    if (sessionId) {
      try { setSession(sessionId); } catch (e) {}
    }

    const user = getCurrentUser();
    if (!user) return { error: 'ログインが必要です', data: [] };

    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const sheet = ss.getSheetByName(STAFF_MASTER_SHEET_NAME);
    if (!sheet) return { error: 'スタッフマスタが見つかりません', data: [] };

    const data = sheet.getDataRange().getValues();
    if (!data || data.length <= 1) return { data: [] };

    const headers = data[0];
    // ヘッダーから列インデックスを動的に取得
    const findIdx = (candidates, fallback) => {
      for (const c of candidates) {
        const idx = headers.indexOf(c);
        if (idx >= 0) return idx;
      }
      return fallback;
    };

    const storeIdx = findIdx(['店舗名', '店舗', '所属店舗'], 0);
    const roleIdx = findIdx(['オーナー', '役職', '職種'], 1);
    const nameIdx = findIdx(['スタッフ名', '氏名', '名前'], 2);
    const systemIdx = findIdx(['報酬制度', '制度'], 3);
    const percentIdx = findIdx(['パーセンテージ', '歩合率', '%'], 4);
    const newAmountIdx = findIdx(['新制度の金額', '新制度金額', '金額'], 5);
    const ownerIdx = findIdx(['オーナー'], 7);

    const accessibleStores = getAccessibleStores();
    const result = [];

    data.slice(1).forEach((row, i) => {
      const store = String(row[storeIdx] || '').trim();
      const role = String(row[roleIdx] || '').trim();
      const name = String(row[nameIdx] || '').trim();
      const system = String(row[systemIdx] || '').trim();
      const percent = String(row[percentIdx] || '').trim();
      const newAmount = Number(row[newAmountIdx]) || 0;
      const isOwner = String(row[ownerIdx] || '').trim();

      if (!name) return;

      // アクセス制限
      if (user.role !== '経営オーナー' && accessibleStores.length > 0 && !accessibleStores.includes(store)) {
        return;
      }

      result.push({
        store: store,
        role: role,
        name: name,
        system: system,
        percent: percent,
        newAmount: newAmount,
        isOwner: isOwner === 'オーナー'
      });
    });

    return { data: result };
  } catch (error) {
    console.error('getStaffRewardMaster error:', error);
    return { error: error.message, data: [] };
  }
}

/**
 * 報酬計算：指定年月のスタッフ別報酬を算出
 * 「報告データシート履歴」から売上データを取得し、制度に基づいて計算
 * @param {string} yearMonth - 対象年月（例: "2026年1月"）
 * @param {string} sessionId - セッションID
 * @return {Object} 報酬計算結果
 */
function calculateRewards(yearMonth, sessionId, storeName) {
  try {
    if (sessionId) {
      try { setSession(sessionId); } catch (e) {}
    }

    const user = getCurrentUser();
    if (!user) return { error: 'ログインが必要です' };
    // スタッフは報酬管理にアクセス不可
    if (user.role === 'スタッフ') return { error: 'アクセス権限がありません' };

    const ss = SpreadsheetApp.openById(getSpreadsheetId());

    // 1. スタッフマスタから報酬制度情報を取得
    const staffMaster = getStaffRewardMaster(sessionId);
    if (staffMaster.error) return staffMaster;

    const staffMap = {};
    staffMaster.data.forEach(s => {
      staffMap[s.name] = s;
    });

    // 2. 報告データシート履歴から売上データを取得
    const reportSheet = ss.getSheetByName('報告データシート履歴');
    if (!reportSheet) return { error: '報告データシート履歴が見つかりません' };

    const reportData = reportSheet.getDataRange().getValues();
    if (!reportData || reportData.length <= 1) return { error: 'データがありません', results: [] };

    const headers = reportData[0];
    const getIdx = (name) => headers.indexOf(name);
    const idx = {
      month: getIdx('年月'),
      store: getIdx('店舗'),
      type: getIdx('種別'),
      amount: getIdx('金額'),
      staff: getIdx('スタッフ氏名'),
      system: getIdx('制度選択'),
      name: getIdx('お名前'),
      rewardRate: getIdx('報酬率'),
      sessionCount: getIdx('セッション数'),
      reviewCount: getIdx('口コミ数'),
      article: getIdx('記事'),
      insta: getIdx('インスタ')
    };

    // +0%〜+15%セッション数列も取得試行
    const plus0Idx = getIdx('＋0%セッション数');
    const plus5Idx = getIdx('＋5%セッション数');
    const plus10Idx = getIdx('＋10%セッション数');
    const plus15Idx = getIdx('＋15%セッション数');
    const attendanceNewIdx = getIdx('出勤回数(新制度)');
    const transportIdx = getIdx('交通費');

    const accessibleStores = getAccessibleStores();

    // 3. 対象年月・店舗のデータを抽出
    const filteredRows = reportData.slice(1).filter(row => {
      const rowMonth = normalizeYearMonth(row[idx.month]);
      if (yearMonth && rowMonth !== yearMonth) return false;

      const store = String(row[idx.store] || '').trim();
      if (accessibleStores.length > 0 && !accessibleStores.includes(store)) return false;

      // 店舗フィルタ
      if (storeName && store !== storeName) return false;

      return true;
    });

    // コース回数列（H列）を取得試行
    const courseIdx = getIdx('コース');
    // 出勤回数列（S列）
    const attendanceIdx = getIdx('出勤回数');

    // 4. スタッフ別に集計
    const staffResults = {};

    filteredRows.forEach(row => {
      const staffName = String(row[idx.staff] || '').trim();
      const store = String(row[idx.store] || '').trim();
      const type = String(row[idx.type] || '').trim();
      const amount = Number(row[idx.amount]) || 0;
      const systemType = String(row[idx.system] || '').trim();
      const customerName = String(row[idx.name] || '').trim();

      if (!staffName) return;

      // 初期化
      if (!staffResults[staffName]) {
        const masterInfo = staffMap[staffName] || {};
        staffResults[staffName] = {
          staffName: staffName,
          store: store,
          system: masterInfo.system || systemType || '未設定',
          role: masterInfo.role || '',
          isOwner: masterInfo.isOwner || false,
          // 旧制度
          oldSystemSales: 0,
          oldSystemReward: 0,
          oldRewardRate: 0,
          // 新制度
          newSystemBase: 0,
          sessionUnitPrice: masterInfo.newAmount || 0,
          totalCourseSessions: 0,
          plus0Sessions: 0,
          plus5Sessions: 0,
          plus10Sessions: 0,
          plus15Sessions: 0,
          attendanceNew: 0,
          transportation: 0,
          newSystemReward: 0,
          // 共通集計
          totalSales: 0,
          totalReward: 0,
          newCount: 0,
          renewCount: 0,
          cancelCount: 0,
          customerCount: 0,
          // 活動（旧制度の報酬率に影響）
          hasArticle: false,
          hasInsta: false,
          articleCount: 0,
          instaCount: 0,
          reviewCount: 0,
          // オーナー報酬
          ownerSelfReward: 0,
          facilityManagementFee: 0,
          hqBilling: 0,
          rent: 0,
          hasRent: false,
          ownerRewardRate: 80,
          facilityFeeRate: 80,
          // 明細
          details: []
        };
      }

      const result = staffResults[staffName];
      result.totalSales += amount;

      // 区分別カウント
      if (type.includes('新規')) result.newCount++;
      if (type.includes('継続')) result.renewCount++;
      if (type.includes('退会')) result.cancelCount++;

      // 顧客数
      if (customerName && customerName !== '-' && customerName !== '' && !type.includes('(0件)')) {
        result.customerCount++;
      }

      // 活動実績（〇/✕判定）
      const articleVal = String(row[idx.article] || '').trim();
      const instaVal = String(row[idx.insta] || '').trim();
      const reviewVal = Number(row[idx.reviewCount]) || 0;
      if (articleVal === '〇' || articleVal === '○') {
        result.articleCount++;
        result.hasArticle = true;
      }
      if (instaVal === '〇' || instaVal === '○') {
        result.instaCount++;
        result.hasInsta = true;
      }
      result.reviewCount += reviewVal;

      // ===== 旧制度の売上集計 =====
      if (systemType.includes('旧')) {
        result.oldSystemSales += amount;
      }

      // ===== 新制度のセッション数・コース回数集計 =====
      if (systemType.includes('新')) {
        // コース回数（H列）からセッション数を取得
        if (courseIdx >= 0) {
          const courseVal = String(row[courseIdx] || '');
          // コース名から回数を抽出（例: "16回コース" → 16）
          const courseMatch = courseVal.match(/(\d+)/);
          if (courseMatch) {
            result.totalCourseSessions += parseInt(courseMatch[1]) || 0;
          }
        }
        // 出勤回数（S列）
        if (attendanceIdx >= 0) {
          result.attendanceNew += (Number(row[attendanceIdx]) || 0);
        }
      }

      // +0%〜+15%セッション数列（フォームの段階別入力がある場合）
      if (plus0Idx >= 0) result.plus0Sessions += (Number(row[plus0Idx]) || 0);
      if (plus5Idx >= 0) result.plus5Sessions += (Number(row[plus5Idx]) || 0);
      if (plus10Idx >= 0) result.plus10Sessions += (Number(row[plus10Idx]) || 0);
      if (plus15Idx >= 0) result.plus15Sessions += (Number(row[plus15Idx]) || 0);
      if (attendanceNewIdx >= 0 && !systemType.includes('新')) {
        result.attendanceNew += (Number(row[attendanceNewIdx]) || 0);
      }
      if (transportIdx >= 0) result.transportation += (Number(row[transportIdx]) || 0);

      // 明細に追加
      result.details.push({
        customer: customerName,
        type: type,
        amount: amount,
        system: systemType,
        store: store,
        article: articleVal,
        insta: instaVal,
        review: reviewVal
      });
    });

    // 5. 報酬計算（制度別）
    Object.keys(staffResults).forEach(staffName => {
      const result = staffResults[staffName];
      const masterInfo = staffMap[staffName] || {};

      // ===== 旧制度：60%ベースの減収方式 =====
      if (result.system.includes('旧') && result.oldSystemSales > 0) {
        // ベース報酬率 60%
        let rewardRate = 60;
        // 記事✕ → -10%
        if (!result.hasArticle) rewardRate -= 10;
        // インスタ✕ → -10%
        if (!result.hasInsta) rewardRate -= 10;
        // 最低0%に制限
        if (rewardRate < 0) rewardRate = 0;

        result.oldRewardRate = rewardRate;
        result.oldSystemReward = Math.round(result.oldSystemSales * (rewardRate / 100));
        // 旧制度は交通費なし
        result.transportation = 0;
      }

      // ===== 新制度：セッション単価 × コース回数 + 交通費 =====
      if (result.system.includes('新')) {
        const unitPrice = result.sessionUnitPrice || masterInfo.newAmount || 0;
        // セッション数（コース回数 or 段階別合計のどちらか大きい方を使用）
        const stageTotal = result.plus0Sessions + result.plus5Sessions + result.plus10Sessions + result.plus15Sessions;
        const sessions = result.totalCourseSessions > 0 ? result.totalCourseSessions : stageTotal;

        result.newSystemBase = Math.round(sessions * unitPrice);
        // 交通費 = 500円 × 出勤回数
        if (result.transportation === 0 && result.attendanceNew > 0) {
          result.transportation = result.attendanceNew * 500;
        }
        result.newSystemReward = result.newSystemBase + result.transportation;
      }

      // 合計スタッフ報酬
      result.totalReward = result.oldSystemReward + result.newSystemReward;
    });

    // 6. 家賃データを取得（「家賃」シート: A列=店舗名、B列=家賃額）
    const rentMap = {};
    try {
      const rentSheet = ss.getSheetByName('家賃');
      if (rentSheet && rentSheet.getLastRow() > 1) {
        const rentData = rentSheet.getRange(2, 1, rentSheet.getLastRow() - 1, 2).getValues();
        rentData.forEach(row => {
          const store = String(row[0] || '').trim();
          const amount = Number(row[1]) || 0;
          if (store) rentMap[store] = amount;
        });
      }
    } catch (e) {
      console.warn('家賃シート取得エラー:', e);
    }

    // 7. オーナー報酬の計算
    // オーナーを特定
    const ownerResults = {};
    const staffOnlyResults = {};
    Object.keys(staffResults).forEach(name => {
      if (staffResults[name].isOwner) {
        ownerResults[name] = staffResults[name];
      } else {
        staffOnlyResults[name] = staffResults[name];
      }
    });

    // 店舗ごとにオーナー報酬を計算
    Object.keys(ownerResults).forEach(ownerName => {
      const owner = ownerResults[ownerName];
      const ownerStore = owner.store;

      // 例外処理: 黒津オーナーは自活動報酬率 90%
      if (ownerName.includes('黒津')) {
        owner.ownerRewardRate = 90;
      }
      // 例外処理: 濱田オーナー（江古田店）は施設管理料率 40%
      if (ownerName.includes('濱田')) {
        owner.facilityFeeRate = 40;
      }

      // オーナー自活動報酬 = 自身の売上 × 報酬率
      owner.ownerSelfReward = Math.round(owner.totalSales * (owner.ownerRewardRate / 100));
      const hqFromOwner = owner.totalSales - owner.ownerSelfReward;

      // 同店舗のスタッフの売上・報酬を集計して施設管理料を計算
      let totalStaffSales = 0;
      let totalStaffReward = 0;
      Object.keys(staffOnlyResults).forEach(staffName => {
        const staff = staffOnlyResults[staffName];
        if (staff.store === ownerStore) {
          totalStaffSales += staff.totalSales;
          totalStaffReward += staff.totalReward;
        }
      });

      // 施設管理料 = (スタッフ売上 − スタッフ報酬) × 施設管理料率
      const staffProfit = totalStaffSales - totalStaffReward;
      owner.facilityManagementFee = Math.round(staffProfit * (owner.facilityFeeRate / 100));

      // 濱田オーナーは施設管理料に交通費を加算（交通費は別途設定）
      if (ownerName.includes('濱田') && owner.transportation > 0) {
        owner.facilityManagementFee += owner.transportation;
      }

      // オーナー本部請求額 = 自活動売上 × (100% - 報酬率) + 家賃
      const hqBase = owner.totalSales - owner.ownerSelfReward;

      // 家賃: 「家賃」シートから取得（A列: 店舗名、B列: 家賃額）
      // 江古田店のみ家賃請求なし
      if (ownerStore.includes('江古田')) {
        owner.hasRent = false;
        owner.rent = 0;
      } else {
        const rentAmount = rentMap[ownerStore] || 0;
        owner.hasRent = true;
        owner.rent = rentAmount;
      }

      owner.hqBilling = hqBase + owner.rent;

      // オーナーの合計報酬 = 自活動報酬 + 施設管理料
      owner.totalReward = owner.ownerSelfReward + owner.facilityManagementFee;

      // 各スタッフの本部請求額を計算（売上 − スタッフ報酬）
      Object.keys(staffOnlyResults).forEach(staffName => {
        const staff = staffOnlyResults[staffName];
        if (staff.store === ownerStore) {
          staff.hqBilling = staff.totalSales - staff.totalReward;
        }
      });
    });

    // 7. 結果を配列に変換
    const results = Object.values(staffResults).map(r => ({
      staffName: r.staffName,
      store: r.store,
      system: r.system,
      role: r.role,
      isOwner: r.isOwner,
      totalSales: r.totalSales,
      // 旧制度
      oldSystemSales: r.oldSystemSales,
      oldRewardRate: r.oldRewardRate,
      oldSystemReward: r.oldSystemReward,
      // 新制度
      sessionUnitPrice: r.sessionUnitPrice,
      totalCourseSessions: r.totalCourseSessions,
      newSystemBase: r.newSystemBase,
      newSystemReward: r.newSystemReward,
      plus0Sessions: r.plus0Sessions,
      plus5Sessions: r.plus5Sessions,
      plus10Sessions: r.plus10Sessions,
      plus15Sessions: r.plus15Sessions,
      attendanceNew: r.attendanceNew,
      transportation: r.transportation,
      // オーナー
      ownerSelfReward: r.ownerSelfReward,
      ownerRewardRate: r.ownerRewardRate,
      facilityManagementFee: r.facilityManagementFee,
      facilityFeeRate: r.facilityFeeRate,
      hqBilling: r.hqBilling,
      rent: r.rent,
      hasRent: r.hasRent,
      // 共通
      totalReward: r.totalReward,
      newCount: r.newCount,
      renewCount: r.renewCount,
      cancelCount: r.cancelCount,
      customerCount: r.customerCount,
      // 活動
      hasArticle: r.hasArticle,
      hasInsta: r.hasInsta,
      articleCount: r.articleCount,
      instaCount: r.instaCount,
      reviewCount: r.reviewCount,
      details: r.details
    }));

    // 合計の計算
    const totals = {
      totalSales: results.reduce((s, r) => s + r.totalSales, 0),
      totalReward: results.reduce((s, r) => s + r.totalReward, 0),
      oldSystemReward: results.reduce((s, r) => s + r.oldSystemReward, 0),
      newSystemReward: results.reduce((s, r) => s + r.newSystemReward, 0),
      totalHqBilling: results.reduce((s, r) => s + (r.hqBilling || 0), 0),
      totalFacilityFee: results.filter(r => r.isOwner).reduce((s, r) => s + (r.facilityManagementFee || 0), 0),
      newCount: results.reduce((s, r) => s + r.newCount, 0),
      renewCount: results.reduce((s, r) => s + r.renewCount, 0),
      cancelCount: results.reduce((s, r) => s + r.cancelCount, 0),
      staffCount: results.length,
      ownerCount: results.filter(r => r.isOwner).length
    };

    return {
      yearMonth: yearMonth,
      storeName: storeName || '全店舗',
      results: results,
      totals: totals
    };
  } catch (error) {
    console.error('calculateRewards error:', error);
    return { error: error.message };
  }
}

/**
 * 報酬計算で使用可能な年月リストを取得
 * 報告データシート履歴から重複なしの年月を抽出
 * @param {string} sessionId - セッションID
 * @return {Array} 年月の配列（降順）
 */
function getAvailableRewardMonths(sessionId) {
  try {
    if (sessionId) {
      try { setSession(sessionId); } catch (e) {}
    }

    const user = getCurrentUser();
    if (!user) return [];

    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const sheet = ss.getSheetByName('報告データシート履歴');
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    if (!data || data.length <= 1) return [];

    const headers = data[0];
    const monthIdx = headers.indexOf('年月');
    if (monthIdx < 0) return [];

    const accessibleStores = getAccessibleStores();
    const storeIdx = headers.indexOf('店舗');

    const monthSet = new Set();
    data.slice(1).forEach(row => {
      // アクセス可能な店舗のみ
      if (storeIdx >= 0 && accessibleStores.length > 0) {
        const store = String(row[storeIdx] || '').trim();
        if (!accessibleStores.includes(store)) return;
      }

      const month = normalizeYearMonth(row[monthIdx]);
      if (month) monthSet.add(month);
    });

    // 降順にソート（最新月が先）
    const months = Array.from(monthSet).sort((a, b) => {
      const matchA = a.match(/(\d{4})年(\d{1,2})月/);
      const matchB = b.match(/(\d{4})年(\d{1,2})月/);
      if (!matchA || !matchB) return 0;
      const numA = parseInt(matchA[1]) * 100 + parseInt(matchA[2]);
      const numB = parseInt(matchB[1]) * 100 + parseInt(matchB[2]);
      return numB - numA;
    });

    return months;
  } catch (error) {
    console.error('getAvailableRewardMonths error:', error);
    return [];
  }
}

/**
 * 特定スタッフの報酬明細を取得
 * @param {string} staffName - スタッフ名
 * @param {string} yearMonth - 対象年月
 * @param {string} sessionId - セッションID
 * @return {Object} 報酬明細データ
 */
function getStaffRewardDetail(staffName, yearMonth, sessionId, storeName) {
  try {
    if (sessionId) {
      try { setSession(sessionId); } catch (e) {}
    }

    const user = getCurrentUser();
    if (!user) return { error: 'ログインが必要です' };
    if (user.role === 'スタッフ') return { error: 'アクセス権限がありません' };

    // calculateRewardsの結果からスタッフを絞り込む
    const allRewards = calculateRewards(yearMonth, sessionId, storeName);
    if (allRewards.error) return allRewards;

    const staffReward = allRewards.results.find(r => r.staffName === staffName);
    if (!staffReward) {
      return { error: `${staffName}のデータが見つかりません`, staffName: staffName };
    }

    // マスタ情報も付加
    const masterInfo = getStaffRewardMaster(sessionId);
    const masterData = (masterInfo.data || []).find(d => d.name === staffName) || {};

    return {
      staffName: staffName,
      yearMonth: yearMonth,
      reward: staffReward,
      master: masterData
    };
  } catch (error) {
    console.error('getStaffRewardDetail error:', error);
    return { error: error.message };
  }
}

/**
 * 報酬計算結果をスプレッドシートに保存
 * @param {string} yearMonth - 対象年月
 * @param {string} sessionId - セッションID
 * @return {Object} 保存結果
 */
function saveRewardCalculation(yearMonth, sessionId, storeName) {
  try {
    if (sessionId) {
      try { setSession(sessionId); } catch (e) {}
    }

    const user = getCurrentUser();
    if (!user) return { error: 'ログインが必要です' };
    if (user.role === 'スタッフ') return { error: 'アクセス権限がありません' };

    // 報酬計算を実行
    const rewards = calculateRewards(yearMonth, sessionId, storeName);
    if (rewards.error) return rewards;

    const ss = SpreadsheetApp.openById(getSpreadsheetId());

    // 報酬計算結果シートを取得/作成
    let rewardSheet = ss.getSheetByName('報酬計算結果');
    if (!rewardSheet) {
      rewardSheet = ss.insertSheet('報酬計算結果');
      const rewardHeaders = [
        'ID', '計算日時', '対象年月', 'スタッフ名', '店舗',
        '報酬制度', '売上合計', '旧制度報酬', '新制度報酬',
        '報酬合計', '新規数', '継続数', '退会数',
        '＋0%セッション数', '＋5%セッション数', '＋10%セッション数', '＋15%セッション数',
        '出勤回数(新制度)', '交通費', '記事', 'インスタ', '口コミ数'
      ];
      rewardSheet.getRange(1, 1, 1, rewardHeaders.length).setValues([rewardHeaders]);
      rewardSheet.getRange(1, 1, 1, rewardHeaders.length).setFontWeight('bold');
      rewardSheet.getRange(1, 1, 1, rewardHeaders.length).setBackground('#e3f2fd');
      rewardSheet.setFrozenRows(1);
    }

    const now = new Date();

    // 各スタッフの結果を書き込み
    rewards.results.forEach(r => {
      const rowData = [
        Utilities.getUuid(),
        now,
        yearMonth,
        r.staffName,
        r.store,
        r.system,
        r.totalSales,
        r.oldSystemReward,
        r.newSystemReward,
        r.totalReward,
        r.newCount,
        r.renewCount,
        r.cancelCount,
        r.plus0Sessions,
        r.plus5Sessions,
        r.plus10Sessions,
        r.plus15Sessions,
        r.attendanceNew,
        r.transportation,
        r.articleCount,
        r.instaCount,
        r.reviewCount
      ];
      rewardSheet.appendRow(rowData);
    });

    return {
      status: 'OK',
      message: `${yearMonth}の報酬計算結果を保存しました（${rewards.results.length}件）`,
      yearMonth: yearMonth,
      count: rewards.results.length
    };
  } catch (error) {
    console.error('saveRewardCalculation error:', error);
    return { error: error.message };
  }
}

/**
 * 保存済みの報酬計算結果を取得
 * @param {string} yearMonth - 対象年月（空の場合は全期間）
 * @param {string} sessionId - セッションID
 * @return {Object} 保存済み報酬データ
 */
function getSavedRewardResults(yearMonth, sessionId) {
  try {
    if (sessionId) {
      try { setSession(sessionId); } catch (e) {}
    }

    const user = getCurrentUser();
    if (!user) return { error: 'ログインが必要です' };
    if (user.role === 'スタッフ') return { error: 'アクセス権限がありません' };

    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const sheet = ss.getSheetByName('報酬計算結果');
    if (!sheet) return { data: [] };

    const data = sheet.getDataRange().getValues();
    if (!data || data.length <= 1) return { data: [] };

    const headers = data[0];
    const accessibleStores = getAccessibleStores();
    const storeIdx = headers.indexOf('店舗');
    const monthIdx = headers.indexOf('対象年月');

    const results = [];
    data.slice(1).forEach(row => {
      const rowMonth = String(row[monthIdx] || '').trim();
      if (yearMonth && rowMonth !== yearMonth) return;

      const store = String(row[storeIdx] || '').trim();
      if (accessibleStores.length > 0 && !accessibleStores.includes(store)) return;

      const record = {};
      headers.forEach((h, i) => {
        record[h] = row[i];
      });
      results.push(record);
    });

    return { data: results };
  } catch (error) {
    console.error('getSavedRewardResults error:', error);
    return { error: error.message, data: [] };
  }
}
