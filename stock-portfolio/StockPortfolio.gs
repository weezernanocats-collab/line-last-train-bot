/**
 * 株式ポートフォリオ管理 - Google Apps Script（修正版）
 *
 * 修正内容:
 * - EPS/BPSは銘柄マスターで手動管理（スクレイピング不安定のため）
 * - 目標株価を追加（アナリスト予想の代替）
 * - 現在値はIMPORTXML関数で取得（より安定）
 *
 * 使用方法:
 * 1. このコードをGoogleスプレッドシートのスクリプトエディタにコピー
 * 2. 「株式管理」メニュー → 「初期化」を実行
 * 3. 「銘柄マスター」シートに銘柄情報を入力
 * 4. 「ポートフォリオ」シートでコードを入力
 */

// ==================== 設定 ====================

const CONFIG = {
  SHEET_PORTFOLIO: 'ポートフォリオ',
  SHEET_MASTER: '銘柄マスター',
  SHEET_SETTINGS: '設定',
  DEFAULT_TARGET_PER: 15,
  DEFAULT_TARGET_PBR: 1.0,
};

// ==================== メイン関数 ====================

function initializeSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  createPortfolioSheet(ss);
  createMasterSheet(ss);
  createSettingsSheet(ss);
  SpreadsheetApp.getUi().alert(
    '初期化が完了しました。\n\n' +
    '【次のステップ】\n' +
    '1.「銘柄マスター」シートに保有銘柄を登録\n' +
    '2.「ポートフォリオ」シートでコードを入力\n' +
    '3. 数量と取得単価を入力'
  );
}

// ==================== ポートフォリオシート ====================

function createPortfolioSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEET_PORTFOLIO);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_PORTFOLIO);
  } else {
    sheet.clear();
    sheet.clearConditionalFormatRules();
  }

  // ヘッダー設定（ユーザー要望通りの列構成）
  const headers = [
    'コード',           // A: 手動入力
    '種別',             // B: 自動（マスターから）
    '銘柄',             // C: 自動（マスターから）
    '数量',             // D: 手動入力
    '取得単価',         // E: 手動入力
    '買い',             // F: 自動判定
    '理論株価(PER)',    // G: 自動計算
    '理論株価(PBR)',    // H: 自動計算
    '目標株価',         // I: マスターから（アナリスト予想代替）
    '乖離率',           // J: 目標株価との乖離率
    '現在値',           // K: IMPORTXML
    '取得額',           // L: 自動計算
    '評価額',           // M: 自動計算
    '損益(円)',         // N: 自動計算
    '損益(%)',          // O: 自動計算
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // ヘッダースタイル
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');

  // 列幅設定
  const columnWidths = [80, 100, 150, 80, 100, 50, 110, 110, 100, 80, 100, 110, 110, 110, 90];
  columnWidths.forEach((width, i) => sheet.setColumnWidth(i + 1, width));

  // データ行に数式を設定（50行分）
  for (let row = 2; row <= 51; row++) {
    setPortfolioFormulas(sheet, row);
  }

  // 条件付き書式
  setConditionalFormatting(sheet);

  // 数値フォーマット
  sheet.getRange('E2:E51').setNumberFormat('#,##0');
  sheet.getRange('G2:I51').setNumberFormat('#,##0');
  sheet.getRange('J2:J51').setNumberFormat('0.0%');
  sheet.getRange('K2:N51').setNumberFormat('#,##0');
  sheet.getRange('O2:O51').setNumberFormat('0.00%');

  // ウィンドウ枠の固定
  sheet.setFrozenRows(1);
}

function setPortfolioFormulas(sheet, row) {
  const master = CONFIG.SHEET_MASTER;
  const settings = CONFIG.SHEET_SETTINGS;

  // B列: 種別（銘柄マスターから）
  sheet.getRange(row, 2).setFormula(
    `=IFERROR(VLOOKUP(A${row},'${master}'!A:G,2,FALSE),"")`
  );

  // C列: 銘柄名（銘柄マスターから）
  sheet.getRange(row, 3).setFormula(
    `=IFERROR(VLOOKUP(A${row},'${master}'!A:G,3,FALSE),"")`
  );

  // F列: 買いシグナル（理論株価または目標株価 > 現在値）
  sheet.getRange(row, 6).setFormula(
    `=IF(OR(A${row}="",K${row}="",K${row}=0),"",` +
    `IF(OR(AND(G${row}<>"",G${row}>K${row}),AND(H${row}<>"",H${row}>K${row}),AND(I${row}<>"",I${row}>K${row})),"○",""))`
  );

  // G列: 理論株価(PER基準) = 適正PER × EPS
  // 適正PERは設定シートから、EPSは銘柄マスターから
  sheet.getRange(row, 7).setFormula(
    `=IFERROR(IF(A${row}="","",` +
    `IFERROR(VLOOKUP(B${row},'${settings}'!A:B,2,FALSE),${CONFIG.DEFAULT_TARGET_PER})` +
    `*VLOOKUP(A${row},'${master}'!A:G,4,FALSE)),"")`
  );

  // H列: 理論株価(PBR基準) = 適正PBR × BPS
  sheet.getRange(row, 8).setFormula(
    `=IFERROR(IF(A${row}="","",` +
    `IFERROR(VLOOKUP(B${row},'${settings}'!A:C,3,FALSE),${CONFIG.DEFAULT_TARGET_PBR})` +
    `*VLOOKUP(A${row},'${master}'!A:G,5,FALSE)),"")`
  );

  // I列: 目標株価（銘柄マスターから）
  sheet.getRange(row, 9).setFormula(
    `=IFERROR(VLOOKUP(A${row},'${master}'!A:G,6,FALSE),"")`
  );

  // J列: 乖離率 = (目標株価 - 現在値) / 現在値
  sheet.getRange(row, 10).setFormula(
    `=IFERROR(IF(OR(I${row}="",K${row}="",K${row}=0),"",(I${row}-K${row})/K${row}),"")`
  );

  // K列: 現在値（IMPORTXML使用 - Google Finance）
  sheet.getRange(row, 11).setFormula(
    `=IF(A${row}="","",IFERROR(VALUE(SUBSTITUTE(SUBSTITUTE(` +
    `IMPORTXML("https://www.google.com/finance/quote/"&A${row}&":TYO","//div[@class='YMlKec fxKbKc']"),` +
    `"￥",""),",","")),"取得中..."))`
  );

  // L列: 取得額 = 取得単価 × 数量
  sheet.getRange(row, 12).setFormula(
    `=IF(OR(D${row}="",E${row}=""),"",D${row}*E${row})`
  );

  // M列: 評価額 = 現在値 × 数量
  sheet.getRange(row, 13).setFormula(
    `=IF(OR(D${row}="",K${row}="",NOT(ISNUMBER(K${row}))),"",D${row}*K${row})`
  );

  // N列: 損益(円) = 評価額 - 取得額
  sheet.getRange(row, 14).setFormula(
    `=IF(OR(L${row}="",M${row}=""),"",M${row}-L${row})`
  );

  // O列: 損益(%) = 損益(円) / 取得額
  sheet.getRange(row, 15).setFormula(
    `=IF(OR(L${row}="",L${row}=0,N${row}=""),"",N${row}/L${row})`
  );
}

function setConditionalFormatting(sheet) {
  const rules = [];

  // 買いシグナル（○で赤背景）
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('○')
    .setBackground('#ffcdd2')
    .setFontColor('#c62828')
    .setBold(true)
    .setRanges([sheet.getRange('F2:F51')])
    .build());

  // 損益プラス（緑）
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setFontColor('#1b5e20')
    .setRanges([sheet.getRange('N2:O51')])
    .build());

  // 損益マイナス（赤）
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setFontColor('#c62828')
    .setRanges([sheet.getRange('N2:O51')])
    .build());

  // 乖離率プラス（緑）- 割安
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setFontColor('#1b5e20')
    .setRanges([sheet.getRange('J2:J51')])
    .build());

  // 乖離率マイナス（赤）- 割高
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setFontColor('#c62828')
    .setRanges([sheet.getRange('J2:J51')])
    .build());

  sheet.setConditionalFormatRules(rules);
}

// ==================== 銘柄マスターシート ====================

function createMasterSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEET_MASTER);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_MASTER);
  } else {
    sheet.clear();
  }

  // ヘッダー（EPS、BPS、目標株価を手動管理）
  const headers = [
    'コード',     // A
    '種別',       // B
    '銘柄名',     // C
    'EPS',        // D: 手動入力（1株当たり利益）
    'BPS',        // E: 手動入力（1株当たり純資産）
    '目標株価',   // F: 手動入力（アナリストコンセンサス）
    '更新日',     // G: 最終更新日
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // ヘッダースタイル
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#34a853');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');

  // サンプルデータ（実際の数値は確認して更新してください）
  const sampleData = [
    ['7203', '自動車', 'トヨタ自動車', 285, 3200, 3500, '2026/01/26'],
    ['9984', '通信', 'ソフトバンクグループ', 450, 5800, 12000, '2026/01/26'],
    ['6758', '電気機器', 'ソニーグループ', 700, 6500, 18000, '2026/01/26'],
    ['8306', '銀行', '三菱UFJフィナンシャル', 150, 1800, 2000, '2026/01/26'],
    ['9432', '通信', '日本電信電話', 350, 2800, 5000, '2026/01/26'],
    ['6861', '電気機器', 'キーエンス', 2500, 15000, 75000, '2026/01/26'],
    ['4063', '化学', '信越化学工業', 750, 7500, 7000, '2026/01/26'],
    ['6501', '電気機器', '日立製作所', 600, 5500, 15000, '2026/01/26'],
    ['7974', 'その他製品', '任天堂', 400, 6800, 10000, '2026/01/26'],
    ['8035', '電気機器', '東京エレクトロン', 2000, 12000, 28000, '2026/01/26'],
  ];

  if (sampleData.length > 0) {
    sheet.getRange(2, 1, sampleData.length, 7).setValues(sampleData);
  }

  // 列幅
  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 80);
  sheet.setColumnWidth(5, 80);
  sheet.setColumnWidth(6, 100);
  sheet.setColumnWidth(7, 100);

  // 数値フォーマット
  sheet.getRange('D2:F100').setNumberFormat('#,##0');

  // 入力ガイド
  sheet.getRange('A1').setNote('4桁の証券コード');
  sheet.getRange('D1').setNote('EPS: 1株当たり利益\n会社四季報やみんかぶで確認');
  sheet.getRange('E1').setNote('BPS: 1株当たり純資産\n会社四季報やみんかぶで確認');
  sheet.getRange('F1').setNote('目標株価: アナリストコンセンサス\nみんかぶ・Yahoo!ファイナンスで確認');
}

// ==================== 設定シート ====================

function createSettingsSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEET_SETTINGS);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_SETTINGS);
  } else {
    sheet.clear();
  }

  const headers = ['種別', '適正PER', '適正PBR', '備考'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#fbbc04');
  headerRange.setFontColor('#000000');
  headerRange.setFontWeight('bold');

  // 業種別適正PER/PBR
  const settingsData = [
    ['自動車', 10, 0.8, '景気敏感・低PER'],
    ['通信', 12, 1.2, '安定配当'],
    ['電気機器', 18, 2.0, '成長期待'],
    ['銀行', 8, 0.5, '低PBR傾向'],
    ['化学', 12, 1.0, ''],
    ['その他製品', 15, 1.5, ''],
    ['医薬品', 20, 2.5, '高成長期待'],
    ['小売', 15, 1.5, ''],
    ['建設', 10, 0.8, ''],
    ['不動産', 12, 1.0, ''],
    ['食品', 18, 1.8, 'ディフェンシブ'],
    ['サービス', 20, 2.0, ''],
  ];

  sheet.getRange(2, 1, settingsData.length, 4).setValues(settingsData);

  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 80);
  sheet.setColumnWidth(3, 80);
  sheet.setColumnWidth(4, 200);

  // 説明追加
  sheet.getRange('A1').setNote('銘柄マスターの「種別」と一致させる');
  sheet.getRange('B1').setNote('理論株価(PER) = 適正PER × EPS');
  sheet.getRange('C1').setNote('理論株価(PBR) = 適正PBR × BPS');
}

// ==================== メニュー ====================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('株式管理')
    .addItem('初期化（シート作成）', 'initializeSpreadsheet')
    .addSeparator()
    .addItem('サマリー表示', 'showPortfolioSummary')
    .addItem('使い方ガイド', 'showGuide')
    .addToUi();
}

function showPortfolioSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_PORTFOLIO);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('ポートフォリオシートが見つかりません。');
    return;
  }

  const data = sheet.getDataRange().getValues();

  let totalCost = 0;
  let totalValue = 0;
  let stockCount = 0;
  let buySignalCount = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0]) {
      stockCount++;
      const cost = parseFloat(row[11]) || 0;
      const value = parseFloat(row[12]) || 0;
      totalCost += cost;
      totalValue += value;
      if (row[5] === '○') buySignalCount++;
    }
  }

  const totalProfit = totalValue - totalCost;
  const profitRate = totalCost > 0 ? (totalProfit / totalCost * 100).toFixed(2) : 0;

  SpreadsheetApp.getUi().alert(
    `【ポートフォリオサマリー】\n\n` +
    `保有銘柄数: ${stockCount}銘柄\n` +
    `買いシグナル: ${buySignalCount}銘柄\n\n` +
    `総取得額: ¥${totalCost.toLocaleString()}\n` +
    `総評価額: ¥${totalValue.toLocaleString()}\n` +
    `総損益: ¥${totalProfit.toLocaleString()} (${profitRate}%)`
  );
}

function showGuide() {
  SpreadsheetApp.getUi().alert(
    `【使い方ガイド】\n\n` +
    `1. 銘柄マスターに保有銘柄を登録\n` +
    `   - コード: 4桁の証券コード\n` +
    `   - 種別: 業種（設定シートと一致させる）\n` +
    `   - EPS/BPS: 会社四季報やみんかぶで確認\n` +
    `   - 目標株価: アナリストコンセンサス\n\n` +
    `2. ポートフォリオでコードを入力\n` +
    `   - 数量と取得単価を入力\n` +
    `   - 現在値は自動取得されます\n\n` +
    `3. 「買い」シグナル\n` +
    `   - 理論株価または目標株価が\n` +
    `     現在値より高い場合に「○」表示\n\n` +
    `【EPS/BPS確認先】\n` +
    `・みんかぶ: minkabu.jp\n` +
    `・Yahoo!ファイナンス\n` +
    `・会社四季報オンライン`
  );
}
