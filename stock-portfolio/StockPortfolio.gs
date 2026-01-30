/**
 * 株式ポートフォリオ管理 - Google Apps Script（v3）
 *
 * 改善点:
 * - 理論株価の計算を修正
 * - 実績PER/PBR、配当利回り、ROEを追加
 * - 割安判定を複合指標で判定
 * - 銘柄マスターの入力ガイドを強化
 */

const CONFIG = {
  SHEET_PORTFOLIO: 'ポートフォリオ',
  SHEET_MASTER: '銘柄マスター',
  SHEET_SETTINGS: '設定',
  DEFAULT_TARGET_PER: 15,
  DEFAULT_TARGET_PBR: 1.0,
};

// ==================== 初期化 ====================

function initializeSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  createMasterSheet(ss);
  createSettingsSheet(ss);
  createPortfolioSheet(ss);

  SpreadsheetApp.getUi().alert(
    '初期化が完了しました。\n\n' +
    '【使い方】\n' +
    '1.「銘柄マスター」に保有銘柄を登録\n' +
    '   ※EPS/BPSは「円」単位で入力\n' +
    '2.「ポートフォリオ」でコード・数量・取得単価を入力\n\n' +
    '【EPS/BPSの確認先】\n' +
    '・みんかぶ: minkabu.jp/stock/[コード]\n' +
    '・Yahoo!ファイナンス\n' +
    '・会社四季報'
  );
}

// ==================== 銘柄マスターシート ====================

function createMasterSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEET_MASTER);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_MASTER);
  } else {
    sheet.clear();
  }

  // ヘッダー
  const headers = [
    'コード',       // A: 4桁証券コード
    '種別',         // B: 業種
    '銘柄名',       // C: 会社名
    'EPS(円)',      // D: 1株当たり利益（円）
    'BPS(円)',      // E: 1株当たり純資産（円）
    '配当(円)',     // F: 1株当たり年間配当（円）
    'ROE(%)',       // G: 自己資本利益率（%）
    '目標株価',     // H: アナリストコンセンサス
    '更新日',       // I: 最終更新日
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // ヘッダースタイル
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#34a853')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // サンプルデータ（実際の概算値 - 要確認・更新）
  // 注: EPS/BPSは「円」単位、ROEは「%」単位
  const sampleData = [
    // [コード, 種別, 銘柄名, EPS(円), BPS(円), 配当(円), ROE(%), 目標株価, 更新日]
    ['7203', '自動車', 'トヨタ自動車', 280, 3200, 75, 8.8, 3200, '2026/01/30'],
    ['9984', '通信', 'ソフトバンクG', 650, 7500, 44, 8.7, 11000, '2026/01/30'],
    ['6758', '電気機器', 'ソニーG', 750, 7200, 85, 10.4, 3800, '2026/01/30'],
    ['8306', '銀行', '三菱UFJFG', 130, 1650, 50, 7.9, 1900, '2026/01/30'],
    ['9432', '通信', 'NTT', 14, 170, 5.2, 8.2, 180, '2026/01/30'],
    ['6861', '電気機器', 'キーエンス', 2400, 14000, 300, 17.1, 72000, '2026/01/30'],
    ['4063', '化学', '信越化学', 750, 7800, 200, 9.6, 6500, '2026/01/30'],
    ['6501', '電気機器', '日立製作所', 580, 5200, 170, 11.2, 4500, '2026/01/30'],
    ['7974', 'その他製品', '任天堂', 380, 6500, 197, 5.8, 9500, '2026/01/30'],
    ['8035', '電気機器', '東京エレクトロン', 2100, 11500, 390, 18.3, 27000, '2026/01/30'],
  ];

  sheet.getRange(2, 1, sampleData.length, 9).setValues(sampleData);

  // 列幅
  const widths = [70, 90, 180, 80, 90, 80, 70, 90, 90];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  // 数値フォーマット
  sheet.getRange('D2:F100').setNumberFormat('#,##0');
  sheet.getRange('G2:G100').setNumberFormat('0.0');
  sheet.getRange('H2:H100').setNumberFormat('#,##0');

  // 入力ガイド（ノート）
  sheet.getRange('D1').setNote(
    'EPS（1株当たり利益）\n' +
    '単位: 円\n' +
    '例: トヨタ 約280円\n\n' +
    '【確認先】\n' +
    'みんかぶ → 業績・財務 → 1株益'
  );
  sheet.getRange('E1').setNote(
    'BPS（1株当たり純資産）\n' +
    '単位: 円\n' +
    '例: トヨタ 約3,200円\n\n' +
    '【確認先】\n' +
    'みんかぶ → 業績・財務 → 1株純資産'
  );
  sheet.getRange('F1').setNote(
    '年間配当金（1株当たり）\n' +
    '単位: 円\n' +
    '例: トヨタ 約75円'
  );
  sheet.getRange('G1').setNote(
    'ROE（自己資本利益率）\n' +
    '単位: %（数値のみ入力）\n' +
    '例: 10.5 と入力\n\n' +
    '計算式: EPS ÷ BPS × 100\n' +
    '一般に10%以上が優良'
  );
  sheet.getRange('H1').setNote(
    '目標株価（アナリストコンセンサス）\n' +
    '【確認先】\n' +
    'みんかぶ → 株価予想 → 目標株価'
  );

  // ヘッダー行を固定
  sheet.setFrozenRows(1);
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

  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#fbbc04')
    .setFontColor('#000000')
    .setFontWeight('bold');

  // 業種別の適正PER/PBR（一般的な目安）
  const data = [
    ['自動車', 10, 0.9, '景気敏感・低PER傾向'],
    ['通信', 12, 1.2, '安定配当'],
    ['電気機器', 18, 2.0, '成長期待で高PER'],
    ['銀行', 8, 0.5, '低PBR傾向'],
    ['化学', 12, 1.0, ''],
    ['その他製品', 15, 1.5, ''],
    ['医薬品', 20, 2.5, '高成長期待'],
    ['小売', 15, 1.5, ''],
    ['建設', 10, 0.8, ''],
    ['不動産', 12, 1.0, ''],
    ['食品', 18, 1.8, 'ディフェンシブ'],
    ['サービス', 20, 2.0, ''],
    ['機械', 12, 1.0, ''],
    ['情報通信', 22, 3.0, '高成長'],
  ];

  sheet.getRange(2, 1, data.length, 4).setValues(data);

  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 80);
  sheet.setColumnWidth(3, 80);
  sheet.setColumnWidth(4, 200);

  sheet.getRange('B1').setNote('理論株価(PER) = 適正PER × EPS');
  sheet.getRange('C1').setNote('理論株価(PBR) = 適正PBR × BPS');
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

  // ヘッダー
  const headers = [
    'コード',           // A: 手動入力
    '種別',             // B: 自動
    '銘柄',             // C: 自動
    '数量',             // D: 手動入力
    '取得単価',         // E: 手動入力
    '現在値',           // F: 自動取得
    '買い',             // G: 自動判定
    '理論株価(PER)',    // H: 計算
    '理論株価(PBR)',    // I: 計算
    '目標株価',         // J: マスターから
    '実績PER',          // K: 計算
    '実績PBR',          // L: 計算
    '配当利回り',       // M: 計算
    'ROE',              // N: マスターから
    '取得額',           // O: 計算
    '評価額',           // P: 計算
    '損益(円)',         // Q: 計算
    '損益(%)',          // R: 計算
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // ヘッダースタイル
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#4285f4')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // 列幅
  const widths = [70, 90, 140, 60, 90, 90, 40, 100, 100, 90, 80, 80, 80, 60, 100, 100, 100, 80];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  // データ行に数式を設定（30行分）
  for (let row = 2; row <= 31; row++) {
    setPortfolioFormulas(sheet, row);
  }

  // 条件付き書式
  setConditionalFormatting(sheet);

  // 数値フォーマット
  sheet.getRange('E2:E31').setNumberFormat('#,##0');
  sheet.getRange('F2:F31').setNumberFormat('#,##0');
  sheet.getRange('H2:J31').setNumberFormat('#,##0');
  sheet.getRange('K2:L31').setNumberFormat('0.00');
  sheet.getRange('M2:M31').setNumberFormat('0.00%');
  sheet.getRange('N2:N31').setNumberFormat('0.0%');
  sheet.getRange('O2:Q31').setNumberFormat('#,##0');
  sheet.getRange('R2:R31').setNumberFormat('0.00%');

  // ウィンドウ枠の固定
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(3);

  // 説明行を追加
  addSummaryRow(sheet);
}

function setPortfolioFormulas(sheet, row) {
  const master = CONFIG.SHEET_MASTER;
  const settings = CONFIG.SHEET_SETTINGS;

  // B列: 種別
  sheet.getRange(row, 2).setFormula(
    `=IFERROR(VLOOKUP(A${row},'${master}'!A:I,2,FALSE),"")`
  );

  // C列: 銘柄名
  sheet.getRange(row, 3).setFormula(
    `=IFERROR(VLOOKUP(A${row},'${master}'!A:I,3,FALSE),"")`
  );

  // F列: 現在値（IMPORTXML - Google Finance）
  sheet.getRange(row, 6).setFormula(
    `=IF(A${row}="","",IFERROR(VALUE(SUBSTITUTE(SUBSTITUTE(` +
    `IMPORTXML("https://www.google.com/finance/quote/"&A${row}&":TYO","//div[@class='YMlKec fxKbKc']"),` +
    `"￥",""),",","")),"..."))`
  );

  // G列: 買いシグナル（複合判定）
  // 条件: (実績PER < 適正PER) OR (実績PBR < 適正PBR) OR (目標株価 > 現在値)
  sheet.getRange(row, 7).setFormula(
    `=IF(OR(A${row}="",F${row}="",NOT(ISNUMBER(F${row}))),"",` +
    `IF(OR(` +
    `AND(K${row}<>"",ISNUMBER(K${row}),K${row}<IFERROR(VLOOKUP(B${row},'${settings}'!A:B,2,FALSE),${CONFIG.DEFAULT_TARGET_PER})),` +
    `AND(L${row}<>"",ISNUMBER(L${row}),L${row}<IFERROR(VLOOKUP(B${row},'${settings}'!A:C,3,FALSE),${CONFIG.DEFAULT_TARGET_PBR})),` +
    `AND(J${row}<>"",ISNUMBER(J${row}),J${row}>F${row})` +
    `),"○",""))`
  );

  // H列: 理論株価(PER基準) = 適正PER × EPS
  sheet.getRange(row, 8).setFormula(
    `=IFERROR(IF(A${row}="","",` +
    `IFERROR(VLOOKUP(B${row},'${settings}'!A:B,2,FALSE),${CONFIG.DEFAULT_TARGET_PER})*` +
    `VLOOKUP(A${row},'${master}'!A:I,4,FALSE)),"")`
  );

  // I列: 理論株価(PBR基準) = 適正PBR × BPS
  sheet.getRange(row, 9).setFormula(
    `=IFERROR(IF(A${row}="","",` +
    `IFERROR(VLOOKUP(B${row},'${settings}'!A:C,3,FALSE),${CONFIG.DEFAULT_TARGET_PBR})*` +
    `VLOOKUP(A${row},'${master}'!A:I,5,FALSE)),"")`
  );

  // J列: 目標株価（マスターから）
  sheet.getRange(row, 10).setFormula(
    `=IFERROR(VLOOKUP(A${row},'${master}'!A:I,8,FALSE),"")`
  );

  // K列: 実績PER = 現在値 / EPS
  sheet.getRange(row, 11).setFormula(
    `=IFERROR(IF(OR(A${row}="",F${row}="",NOT(ISNUMBER(F${row}))),"",` +
    `F${row}/VLOOKUP(A${row},'${master}'!A:I,4,FALSE)),"")`
  );

  // L列: 実績PBR = 現在値 / BPS
  sheet.getRange(row, 12).setFormula(
    `=IFERROR(IF(OR(A${row}="",F${row}="",NOT(ISNUMBER(F${row}))),"",` +
    `F${row}/VLOOKUP(A${row},'${master}'!A:I,5,FALSE)),"")`
  );

  // M列: 配当利回り = 配当 / 現在値
  sheet.getRange(row, 13).setFormula(
    `=IFERROR(IF(OR(A${row}="",F${row}="",NOT(ISNUMBER(F${row})),F${row}=0),"",` +
    `VLOOKUP(A${row},'${master}'!A:I,6,FALSE)/F${row}),"")`
  );

  // N列: ROE（マスターから）
  sheet.getRange(row, 14).setFormula(
    `=IFERROR(IF(A${row}="","",VLOOKUP(A${row},'${master}'!A:I,7,FALSE)/100),"")`
  );

  // O列: 取得額 = 取得単価 × 数量
  sheet.getRange(row, 15).setFormula(
    `=IF(OR(D${row}="",E${row}=""),"",D${row}*E${row})`
  );

  // P列: 評価額 = 現在値 × 数量
  sheet.getRange(row, 16).setFormula(
    `=IF(OR(D${row}="",F${row}="",NOT(ISNUMBER(F${row}))),"",D${row}*F${row})`
  );

  // Q列: 損益(円) = 評価額 - 取得額
  sheet.getRange(row, 17).setFormula(
    `=IF(OR(O${row}="",P${row}=""),"",P${row}-O${row})`
  );

  // R列: 損益(%) = 損益(円) / 取得額
  sheet.getRange(row, 18).setFormula(
    `=IF(OR(O${row}="",O${row}=0,Q${row}=""),"",Q${row}/O${row})`
  );
}

function addSummaryRow(sheet) {
  const summaryRow = 33;

  // ラベル
  sheet.getRange(summaryRow, 1).setValue('【合計】');
  sheet.getRange(summaryRow, 1).setFontWeight('bold');
  sheet.getRange(summaryRow, 1).setBackground('#e8eaf6');

  // 合計: 取得額
  sheet.getRange(summaryRow, 15).setFormula('=SUM(O2:O31)');
  sheet.getRange(summaryRow, 15).setNumberFormat('#,##0');
  sheet.getRange(summaryRow, 15).setFontWeight('bold');

  // 合計: 評価額
  sheet.getRange(summaryRow, 16).setFormula('=SUM(P2:P31)');
  sheet.getRange(summaryRow, 16).setNumberFormat('#,##0');
  sheet.getRange(summaryRow, 16).setFontWeight('bold');

  // 合計: 損益(円)
  sheet.getRange(summaryRow, 17).setFormula('=SUM(Q2:Q31)');
  sheet.getRange(summaryRow, 17).setNumberFormat('#,##0');
  sheet.getRange(summaryRow, 17).setFontWeight('bold');

  // 合計: 損益(%)
  sheet.getRange(summaryRow, 18).setFormula('=IF(O33=0,"",Q33/O33)');
  sheet.getRange(summaryRow, 18).setNumberFormat('0.00%');
  sheet.getRange(summaryRow, 18).setFontWeight('bold');

  // 背景色
  sheet.getRange(summaryRow, 1, 1, 18).setBackground('#e8eaf6');
}

function setConditionalFormatting(sheet) {
  const rules = [];

  // 買いシグナル（○で赤背景・太字）
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('○')
    .setBackground('#ffcdd2')
    .setFontColor('#c62828')
    .setBold(true)
    .setRanges([sheet.getRange('G2:G31')])
    .build());

  // 損益プラス（緑）
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setFontColor('#1b5e20')
    .setRanges([sheet.getRange('Q2:R31')])
    .build());

  // 損益マイナス（赤）
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setFontColor('#c62828')
    .setRanges([sheet.getRange('Q2:R31')])
    .build());

  // 実績PER: 低い（割安）= 緑、高い（割高）= 赤
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(12)
    .setFontColor('#1b5e20')
    .setRanges([sheet.getRange('K2:K31')])
    .build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(25)
    .setFontColor('#c62828')
    .setRanges([sheet.getRange('K2:K31')])
    .build());

  // 実績PBR: 低い（割安）= 緑、高い（割高）= 赤
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(1)
    .setFontColor('#1b5e20')
    .setRanges([sheet.getRange('L2:L31')])
    .build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(3)
    .setFontColor('#c62828')
    .setRanges([sheet.getRange('L2:L31')])
    .build());

  // 配当利回り: 高い = 緑
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0.03)
    .setFontColor('#1b5e20')
    .setRanges([sheet.getRange('M2:M31')])
    .build());

  // ROE: 高い（優良）= 緑
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0.10)
    .setFontColor('#1b5e20')
    .setRanges([sheet.getRange('N2:N31')])
    .build());

  sheet.setConditionalFormatRules(rules);
}

// ==================== メニュー ====================

function onOpen() {
  SpreadsheetApp.getUi().createMenu('株式管理')
    .addItem('初期化（シート作成）', 'initializeSpreadsheet')
    .addSeparator()
    .addItem('サマリー表示', 'showSummary')
    .addItem('割安銘柄を表示', 'showUndervalued')
    .addItem('使い方ガイド', 'showGuide')
    .addToUi();
}

function showSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_PORTFOLIO);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  let count = 0, buyCount = 0, totalCost = 0, totalValue = 0;

  for (let i = 1; i < data.length - 2; i++) {
    if (data[i][0]) {
      count++;
      if (data[i][6] === '○') buyCount++;
      totalCost += parseFloat(data[i][14]) || 0;
      totalValue += parseFloat(data[i][15]) || 0;
    }
  }

  const profit = totalValue - totalCost;
  const rate = totalCost > 0 ? (profit / totalCost * 100).toFixed(2) : 0;

  SpreadsheetApp.getUi().alert(
    `【ポートフォリオサマリー】\n\n` +
    `保有銘柄数: ${count}\n` +
    `割安シグナル: ${buyCount}銘柄\n\n` +
    `総取得額: ¥${totalCost.toLocaleString()}\n` +
    `総評価額: ¥${totalValue.toLocaleString()}\n` +
    `総損益: ¥${profit.toLocaleString()} (${rate}%)`
  );
}

function showUndervalued() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_PORTFOLIO);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  let undervalued = [];

  for (let i = 1; i < data.length - 2; i++) {
    if (data[i][0] && data[i][6] === '○') {
      undervalued.push(`${data[i][2]}（${data[i][0]}）`);
    }
  }

  if (undervalued.length === 0) {
    SpreadsheetApp.getUi().alert('現在、割安シグナルが出ている銘柄はありません。');
  } else {
    SpreadsheetApp.getUi().alert(
      `【割安シグナル銘柄】\n\n${undervalued.join('\n')}\n\n` +
      `判定基準:\n` +
      `・実績PER < 適正PER\n` +
      `・実績PBR < 適正PBR\n` +
      `・目標株価 > 現在値`
    );
  }
}

function showGuide() {
  SpreadsheetApp.getUi().alert(
    `【株式ポートフォリオ 使い方】\n\n` +
    `■ 銘柄マスターの入力\n` +
    `・EPS: 1株利益（円）例: 280\n` +
    `・BPS: 1株純資産（円）例: 3200\n` +
    `・配当: 年間配当（円）例: 75\n` +
    `・ROE: %で入力 例: 10.5\n\n` +
    `■ 指標の見方\n` +
    `・実績PER: 低いほど割安（12以下=緑）\n` +
    `・実績PBR: 低いほど割安（1以下=緑）\n` +
    `・配当利回り: 3%以上=緑\n` +
    `・ROE: 10%以上=優良企業\n\n` +
    `■ 買いシグナル（○）\n` +
    `以下のいずれかを満たす:\n` +
    `・実績PER < 業種適正PER\n` +
    `・実績PBR < 業種適正PBR\n` +
    `・目標株価 > 現在値`
  );
}
