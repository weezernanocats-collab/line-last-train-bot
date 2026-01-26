/**
 * 株式ポートフォリオ管理 - Google Apps Script
 *
 * 使用方法:
 * 1. このコードをGoogleスプレッドシートのスクリプトエディタにコピー
 * 2. 「銘柄マスター」シートに銘柄一覧を作成
 * 3. 「ポートフォリオ」シートでコードを入力
 *
 * データソース:
 * - 現在値: Google Finance
 * - PER/PBR/EPS/BPS: みんかぶ（グレーゾーン、負荷をかけない範囲で使用）
 * - 業種別適正PER/PBR: 設定シートで管理
 */

// ==================== 設定 ====================

const CONFIG = {
  // シート名
  SHEET_PORTFOLIO: 'ポートフォリオ',
  SHEET_MASTER: '銘柄マスター',
  SHEET_SETTINGS: '設定',

  // キャッシュ有効期間（秒）
  CACHE_DURATION: 3600, // 1時間

  // 適正PER/PBRのデフォルト値
  DEFAULT_TARGET_PER: 15,  // 日経平均基準
  DEFAULT_TARGET_PBR: 1.0, // 解散価値基準
};

// ==================== メイン関数 ====================

/**
 * スプレッドシート初期化
 */
function initializeSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ポートフォリオシート作成
  createPortfolioSheet(ss);

  // 銘柄マスターシート作成
  createMasterSheet(ss);

  // 設定シート作成
  createSettingsSheet(ss);

  SpreadsheetApp.getUi().alert('初期化が完了しました。\n「銘柄マスター」シートに銘柄情報を追加してください。');
}

/**
 * ポートフォリオシート作成
 */
function createPortfolioSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEET_PORTFOLIO);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_PORTFOLIO);
  } else {
    sheet.clear();
  }

  // ヘッダー設定
  const headers = [
    'コード',      // A: 手動入力
    '種別',        // B: 自動取得（マスターから）
    '銘柄',        // C: 自動取得
    '数量',        // D: 手動入力
    '取得単価',    // E: 手動入力
    '買い',        // F: 条件判定（理論株価 > 現在値）
    '理論株価(PER)', // G: 計算
    '理論株価(PBR)', // H: 計算
    'PER予想',     // I: 手動入力（アナリスト予想）
    'PBR予想',     // J: 手動入力（アナリスト予想）
    '現在値',      // K: 自動取得
    '取得額',      // L: 計算
    '評価額',      // M: 計算
    '損益(円)',    // N: 計算
    '損益(%)',     // O: 計算
    'EPS',         // P: 自動取得（補助列）
    'BPS',         // Q: 自動取得（補助列）
    'PER',         // R: 自動取得（補助列）
    'PBR',         // S: 自動取得（補助列）
    '適正PER',     // T: 設定から取得
    '適正PBR',     // U: 設定から取得
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // ヘッダースタイル
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');

  // 列幅設定
  sheet.setColumnWidth(1, 80);   // コード
  sheet.setColumnWidth(2, 100);  // 種別
  sheet.setColumnWidth(3, 150);  // 銘柄
  sheet.setColumnWidth(4, 80);   // 数量
  sheet.setColumnWidth(5, 100);  // 取得単価
  sheet.setColumnWidth(6, 50);   // 買い
  sheet.setColumnWidth(7, 120);  // 理論株価(PER)
  sheet.setColumnWidth(8, 120);  // 理論株価(PBR)
  sheet.setColumnWidth(9, 80);   // PER予想
  sheet.setColumnWidth(10, 80);  // PBR予想
  sheet.setColumnWidth(11, 100); // 現在値
  sheet.setColumnWidth(12, 120); // 取得額
  sheet.setColumnWidth(13, 120); // 評価額
  sheet.setColumnWidth(14, 120); // 損益(円)
  sheet.setColumnWidth(15, 100); // 損益(%)

  // 補助列を非表示
  sheet.hideColumns(16, 6); // P〜U列を非表示

  // データ行のフォーマット設定（2行目以降、100行分）
  const dataStartRow = 2;
  const dataRows = 100;

  // 数式を設定
  for (let row = dataStartRow; row < dataStartRow + dataRows; row++) {
    setPortfolioFormulas(sheet, row);
  }

  // 条件付き書式（買いシグナル）
  const buyRange = sheet.getRange('F2:F101');
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('○')
    .setBackground('#ffcdd2')
    .setFontColor('#c62828')
    .setBold(true)
    .setRanges([buyRange])
    .build();
  sheet.setConditionalFormatRules([rule]);

  // 損益の条件付き書式
  const profitRange = sheet.getRange('N2:O101');
  const profitRulePositive = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setFontColor('#1b5e20')
    .setRanges([profitRange])
    .build();
  const profitRuleNegative = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setFontColor('#c62828')
    .setRanges([profitRange])
    .build();

  const rules = sheet.getConditionalFormatRules();
  rules.push(profitRulePositive);
  rules.push(profitRuleNegative);
  sheet.setConditionalFormatRules(rules);

  // 数値フォーマット
  sheet.getRange('E2:E101').setNumberFormat('#,##0');
  sheet.getRange('G2:H101').setNumberFormat('#,##0');
  sheet.getRange('K2:N101').setNumberFormat('#,##0');
  sheet.getRange('O2:O101').setNumberFormat('0.00%');

  // ウィンドウ枠の固定
  sheet.setFrozenRows(1);
}

/**
 * ポートフォリオ行に数式を設定
 */
function setPortfolioFormulas(sheet, row) {
  const masterSheet = CONFIG.SHEET_MASTER;
  const settingsSheet = CONFIG.SHEET_SETTINGS;

  // B列: 種別（銘柄マスターからVLOOKUP）
  sheet.getRange(row, 2).setFormula(
    `=IFERROR(VLOOKUP(A${row},'${masterSheet}'!A:C,2,FALSE),"")`
  );

  // C列: 銘柄名（銘柄マスターからVLOOKUP）
  sheet.getRange(row, 3).setFormula(
    `=IFERROR(VLOOKUP(A${row},'${masterSheet}'!A:C,3,FALSE),"")`
  );

  // F列: 買いシグナル
  sheet.getRange(row, 6).setFormula(
    `=IF(OR(A${row}="",K${row}=""),"",IF(OR(AND(G${row}<>"",G${row}>K${row}),AND(H${row}<>"",H${row}>K${row})),"○",""))`
  );

  // G列: 理論株価(PER基準) = 適正PER × EPS
  sheet.getRange(row, 7).setFormula(
    `=IFERROR(IF(OR(A${row}="",P${row}=""),"",T${row}*P${row}),"")`
  );

  // H列: 理論株価(PBR基準) = 適正PBR × BPS
  sheet.getRange(row, 8).setFormula(
    `=IFERROR(IF(OR(A${row}="",Q${row}=""),"",U${row}*Q${row}),"")`
  );

  // K列: 現在値（カスタム関数で取得）
  sheet.getRange(row, 11).setFormula(
    `=IF(A${row}="","",getStockPrice(A${row}))`
  );

  // L列: 取得額 = 取得単価 × 数量
  sheet.getRange(row, 12).setFormula(
    `=IF(OR(D${row}="",E${row}=""),"",D${row}*E${row})`
  );

  // M列: 評価額 = 現在値 × 数量
  sheet.getRange(row, 13).setFormula(
    `=IF(OR(D${row}="",K${row}=""),"",D${row}*K${row})`
  );

  // N列: 損益(円) = 評価額 - 取得額
  sheet.getRange(row, 14).setFormula(
    `=IF(OR(L${row}="",M${row}=""),"",M${row}-L${row})`
  );

  // O列: 損益(%) = 損益(円) / 取得額
  sheet.getRange(row, 15).setFormula(
    `=IF(OR(L${row}="",N${row}=""),"",N${row}/L${row})`
  );

  // P列: EPS（カスタム関数）
  sheet.getRange(row, 16).setFormula(
    `=IF(A${row}="","",getStockEPS(A${row}))`
  );

  // Q列: BPS（カスタム関数）
  sheet.getRange(row, 17).setFormula(
    `=IF(A${row}="","",getStockBPS(A${row}))`
  );

  // R列: PER（カスタム関数）
  sheet.getRange(row, 18).setFormula(
    `=IF(A${row}="","",getStockPER(A${row}))`
  );

  // S列: PBR（カスタム関数）
  sheet.getRange(row, 19).setFormula(
    `=IF(A${row}="","",getStockPBR(A${row}))`
  );

  // T列: 適正PER（設定シートから種別でVLOOKUP）
  sheet.getRange(row, 20).setFormula(
    `=IFERROR(VLOOKUP(B${row},'${settingsSheet}'!A:B,2,FALSE),${CONFIG.DEFAULT_TARGET_PER})`
  );

  // U列: 適正PBR（設定シートから種別でVLOOKUP）
  sheet.getRange(row, 21).setFormula(
    `=IFERROR(VLOOKUP(B${row},'${settingsSheet}'!A:C,3,FALSE),${CONFIG.DEFAULT_TARGET_PBR})`
  );
}

/**
 * 銘柄マスターシート作成
 */
function createMasterSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEET_MASTER);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_MASTER);
  } else {
    sheet.clear();
  }

  // ヘッダー
  const headers = ['コード', '種別', '銘柄名'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // ヘッダースタイル
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#34a853');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');

  // サンプルデータ
  const sampleData = [
    ['7203', '自動車', 'トヨタ自動車'],
    ['9984', '通信', 'ソフトバンクグループ'],
    ['6758', '電気機器', 'ソニーグループ'],
    ['8306', '銀行', '三菱UFJフィナンシャル・グループ'],
    ['9432', '通信', '日本電信電話'],
    ['6861', '電気機器', 'キーエンス'],
    ['4063', '化学', '信越化学工業'],
    ['6501', '電気機器', '日立製作所'],
    ['7974', 'その他製品', '任天堂'],
    ['8035', '電気機器', '東京エレクトロン'],
  ];

  sheet.getRange(2, 1, sampleData.length, 3).setValues(sampleData);

  // 列幅
  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 250);
}

/**
 * 設定シート作成
 */
function createSettingsSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEET_SETTINGS);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_SETTINGS);
  } else {
    sheet.clear();
  }

  // ヘッダー
  const headers = ['種別', '適正PER', '適正PBR', '備考'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // ヘッダースタイル
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#fbbc04');
  headerRange.setFontColor('#000000');
  headerRange.setFontWeight('bold');

  // 業種別適正PER/PBRサンプル
  const settingsData = [
    ['自動車', 10, 0.8, '製造業は低めのPER'],
    ['通信', 12, 1.2, '安定成長株'],
    ['電気機器', 18, 2.0, '成長期待が高い'],
    ['銀行', 8, 0.5, '金融業は低PBR'],
    ['化学', 12, 1.0, ''],
    ['その他製品', 15, 1.5, ''],
    ['医薬品', 20, 2.5, '高成長期待'],
    ['小売', 15, 1.5, ''],
    ['建設', 10, 0.8, ''],
    ['不動産', 12, 1.0, ''],
  ];

  sheet.getRange(2, 1, settingsData.length, 4).setValues(settingsData);

  // 列幅
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 80);
  sheet.setColumnWidth(3, 80);
  sheet.setColumnWidth(4, 200);
}

// ==================== 株価データ取得関数 ====================

/**
 * 株価取得（Google Finance経由）
 * @param {string} code 証券コード
 * @return {number} 現在株価
 * @customfunction
 */
function getStockPrice(code) {
  if (!code) return '';

  const cache = CacheService.getScriptCache();
  const cacheKey = `price_${code}`;
  const cached = cache.get(cacheKey);

  if (cached) {
    return parseFloat(cached);
  }

  try {
    // Google Finance から取得
    const url = `https://www.google.com/finance/quote/${code}:TYO?hl=ja`;
    const response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
    const html = response.getContentText();

    // 株価を抽出（YMlKec fxKbKc クラスを持つ要素）
    const priceMatch = html.match(/class="YMlKec fxKbKc"[^>]*>([^<]+)</);
    if (priceMatch) {
      let price = priceMatch[1]
        .replace(/[￥¥,\s]/g, '')
        .replace(/,/g, '');
      price = parseFloat(price);

      if (!isNaN(price)) {
        cache.put(cacheKey, price.toString(), CONFIG.CACHE_DURATION);
        return price;
      }
    }

    return 'N/A';
  } catch (e) {
    console.error(`株価取得エラー (${code}): ${e.message}`);
    return 'Error';
  }
}

/**
 * EPS取得（みんかぶから）
 * @param {string} code 証券コード
 * @return {number} EPS（1株当たり利益）
 * @customfunction
 */
function getStockEPS(code) {
  if (!code) return '';

  const data = getMinkabuData(code);
  return data.eps || '';
}

/**
 * BPS取得（みんかぶから）
 * @param {string} code 証券コード
 * @return {number} BPS（1株当たり純資産）
 * @customfunction
 */
function getStockBPS(code) {
  if (!code) return '';

  const data = getMinkabuData(code);
  return data.bps || '';
}

/**
 * PER取得
 * @param {string} code 証券コード
 * @return {number} PER
 * @customfunction
 */
function getStockPER(code) {
  if (!code) return '';

  const data = getMinkabuData(code);
  return data.per || '';
}

/**
 * PBR取得
 * @param {string} code 証券コード
 * @return {number} PBR
 * @customfunction
 */
function getStockPBR(code) {
  if (!code) return '';

  const data = getMinkabuData(code);
  return data.pbr || '';
}

/**
 * みんかぶからデータを一括取得（キャッシュ付き）
 * @param {string} code 証券コード
 * @return {object} {eps, bps, per, pbr}
 */
function getMinkabuData(code) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `minkabu_${code}`;
  const cached = cache.get(cacheKey);

  if (cached) {
    return JSON.parse(cached);
  }

  const data = {
    eps: null,
    bps: null,
    per: null,
    pbr: null
  };

  try {
    const url = `https://minkabu.jp/stock/${code}`;
    const response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
    const html = response.getContentText();

    // PER抽出
    const perMatch = html.match(/PER[（(]調整後[)）]<\/th>\s*<td[^>]*>([0-9,.]+)/);
    if (perMatch) {
      data.per = parseFloat(perMatch[1].replace(/,/g, ''));
    } else {
      // 代替パターン
      const perMatch2 = html.match(/PER<\/th>\s*<td[^>]*>([0-9,.]+)/);
      if (perMatch2) {
        data.per = parseFloat(perMatch2[1].replace(/,/g, ''));
      }
    }

    // PBR抽出
    const pbrMatch = html.match(/PBR[（(]実績[)）]<\/th>\s*<td[^>]*>([0-9,.]+)/);
    if (pbrMatch) {
      data.pbr = parseFloat(pbrMatch[1].replace(/,/g, ''));
    } else {
      const pbrMatch2 = html.match(/PBR<\/th>\s*<td[^>]*>([0-9,.]+)/);
      if (pbrMatch2) {
        data.pbr = parseFloat(pbrMatch2[1].replace(/,/g, ''));
      }
    }

    // EPS抽出（1株益）
    const epsMatch = html.match(/1株益[（(]円[)）]<\/th>\s*<td[^>]*>([0-9,.]+)/);
    if (epsMatch) {
      data.eps = parseFloat(epsMatch[1].replace(/,/g, ''));
    } else {
      const epsMatch2 = html.match(/EPS<\/th>\s*<td[^>]*>([0-9,.]+)/);
      if (epsMatch2) {
        data.eps = parseFloat(epsMatch2[1].replace(/,/g, ''));
      }
    }

    // BPS抽出（1株純資産）
    const bpsMatch = html.match(/1株純資産[（(]円[)）]<\/th>\s*<td[^>]*>([0-9,.]+)/);
    if (bpsMatch) {
      data.bps = parseFloat(bpsMatch[1].replace(/,/g, ''));
    } else {
      const bpsMatch2 = html.match(/BPS<\/th>\s*<td[^>]*>([0-9,.]+)/);
      if (bpsMatch2) {
        data.bps = parseFloat(bpsMatch2[1].replace(/,/g, ''));
      }
    }

    // キャッシュに保存
    cache.put(cacheKey, JSON.stringify(data), CONFIG.CACHE_DURATION);

  } catch (e) {
    console.error(`みんかぶデータ取得エラー (${code}): ${e.message}`);
  }

  return data;
}

// ==================== 一括更新機能 ====================

/**
 * 全銘柄のデータを一括更新
 */
function refreshAllData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_PORTFOLIO);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('ポートフォリオシートが見つかりません。');
    return;
  }

  // キャッシュをクリア
  const cache = CacheService.getScriptCache();

  // コード列を取得
  const codes = sheet.getRange('A2:A101').getValues().flat().filter(c => c);

  // 各コードのキャッシュを削除
  codes.forEach(code => {
    cache.remove(`price_${code}`);
    cache.remove(`minkabu_${code}`);
  });

  // シートを再計算
  SpreadsheetApp.flush();

  SpreadsheetApp.getUi().alert(`${codes.length}銘柄のデータを更新しました。`);
}

/**
 * サマリーを表示
 */
function showPortfolioSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_PORTFOLIO);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('ポートフォリオシートが見つかりません。');
    return;
  }

  const data = sheet.getDataRange().getValues();

  let totalCost = 0;      // 総取得額
  let totalValue = 0;     // 総評価額
  let stockCount = 0;     // 銘柄数
  let buySignalCount = 0; // 買いシグナル数

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0]) { // コードがある行
      stockCount++;

      const cost = parseFloat(row[11]) || 0;     // L列: 取得額
      const value = parseFloat(row[12]) || 0;    // M列: 評価額
      const buySignal = row[5];                   // F列: 買いシグナル

      totalCost += cost;
      totalValue += value;

      if (buySignal === '○') {
        buySignalCount++;
      }
    }
  }

  const totalProfit = totalValue - totalCost;
  const profitRate = totalCost > 0 ? (totalProfit / totalCost * 100).toFixed(2) : 0;

  const message = `
【ポートフォリオサマリー】

保有銘柄数: ${stockCount}銘柄
買いシグナル: ${buySignalCount}銘柄

総取得額: ¥${totalCost.toLocaleString()}
総評価額: ¥${totalValue.toLocaleString()}
総損益: ¥${totalProfit.toLocaleString()} (${profitRate}%)
  `;

  SpreadsheetApp.getUi().alert(message);
}

// ==================== メニュー ====================

/**
 * メニューを追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('株式管理')
    .addItem('初期化（シート作成）', 'initializeSpreadsheet')
    .addSeparator()
    .addItem('全データ更新', 'refreshAllData')
    .addItem('サマリー表示', 'showPortfolioSummary')
    .addToUi();
}

// ==================== IMPORTXML代替（シンプル版） ====================

/**
 * IMPORTXML関数の代替（Google Finance用）
 * スプレッドシートの数式として使用: =STOCK_PRICE("7203")
 *
 * @param {string} code 証券コード
 * @return {number} 株価
 * @customfunction
 */
function STOCK_PRICE(code) {
  return getStockPrice(code);
}

/**
 * @param {string} code 証券コード
 * @return {number} EPS
 * @customfunction
 */
function STOCK_EPS(code) {
  return getStockEPS(code);
}

/**
 * @param {string} code 証券コード
 * @return {number} BPS
 * @customfunction
 */
function STOCK_BPS(code) {
  return getStockBPS(code);
}

/**
 * @param {string} code 証券コード
 * @return {number} PER
 * @customfunction
 */
function STOCK_PER(code) {
  return getStockPER(code);
}

/**
 * @param {string} code 証券コード
 * @return {number} PBR
 * @customfunction
 */
function STOCK_PBR(code) {
  return getStockPBR(code);
}
