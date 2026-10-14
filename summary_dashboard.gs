/**
 * 売上ダッシュボード
 * 「売上データ」シートを月ごとに集計し、「月次サマリー」シートに集計表と月次推移の棒グラフを出力する。
 */

const SOURCE_SHEET_NAME = '売上データ';
const SUMMARY_SHEET_NAME = '月次サマリー';
const TRIGGER_HOUR = 9;

/**
 * メイン処理：売上データを月ごとに集計し、「月次サマリー」シートとグラフを作り直す
 */
function runDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = ss.getSheetByName(SOURCE_SHEET_NAME);
  if (!source) {
    throw new Error(`「${SOURCE_SHEET_NAME}」シートが見つかりません`);
  }
  const summary = ss.getSheetByName(SUMMARY_SHEET_NAME) || ss.insertSheet(SUMMARY_SHEET_NAME);

  const rows = source.getDataRange().getValues();
  const monthly = aggregateByMonth_(rows, ss.getSpreadsheetTimeZone());

  writeSummary_(summary, monthly);
  rebuildChart_(summary, monthly.length);
}

/**
 * runDashboard を毎朝9時台に自動実行するトリガーを登録する（最初に一度だけ手動で実行する）
 */
function setDailyTrigger() {
  // 何度実行してもトリガーが重複しないよう、既存の runDashboard 用トリガーを先に削除する
  ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction() === 'runDashboard')
    .forEach(trigger => ScriptApp.deleteTrigger(trigger));

  ScriptApp.newTrigger('runDashboard')
    .timeBased()
    .everyDays(1)
    .atHour(TRIGGER_HOUR)
    .inTimezone(SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone())
    .create();
}

/**
 * 売上データの各行を月ごとに集計する
 * 日付または金額が読み取れない行（見出し行・空行など）は集計対象外とする
 * @param {Array<Array<*>>} rows 「売上データ」シートの全行
 * @param {string} timeZone スプレッドシートのタイムゾーン
 * @return {Array<{label: string, total: number, count: number}>} 月の昇順に並べた集計結果
 */
function aggregateByMonth_(rows, timeZone) {
  const totals = {};
  rows.forEach(row => {
    const monthKey = toMonthKey_(row[0], timeZone);
    const amount = toAmount_(row[3]);
    if (!monthKey || amount === null) return;

    if (!totals[monthKey]) totals[monthKey] = { total: 0, count: 0 };
    totals[monthKey].total += amount;
    totals[monthKey].count += 1;
  });

  // キーは「yyyy-MM」形式なので、文字列のソートで時系列順になる
  return Object.keys(totals).sort().map(key => {
    const [year, month] = key.split('-');
    return {
      label: `${year}年${Number(month)}月`,
      total: totals[key].total,
      count: totals[key].count,
    };
  });
}

/**
 * 日付セルの値を「yyyy-MM」形式のキーに変換する。変換できない場合は null を返す
 */
function toMonthKey_(value, timeZone) {
  if (value instanceof Date) {
    if (isNaN(value.getTime())) return null;
    // スクリプトとスプレッドシートのタイムゾーンが異なると月初の日付が前月にずれるため、スプレッドシート側に合わせる
    return Utilities.formatDate(value, timeZone, 'yyyy-MM');
  }

  const match = String(value).trim().match(/^(\d{4})[\/-](\d{1,2})[\/-]\d{1,2}/);
  if (!match) return null;
  const month = Number(match[2]);
  if (month < 1 || month > 12) return null;
  return `${match[1]}-${String(month).padStart(2, '0')}`;
}

/**
 * 金額セルの値を数値に変換する。数値として読めない場合は null を返す
 */
function toAmount_(value) {
  if (typeof value === 'number') return isFinite(value) ? value : null;

  const text = String(value).replace(/,/g, '').trim();
  if (text === '') return null;
  const amount = Number(text);
  return isFinite(amount) ? amount : null;
}

/**
 * 「月次サマリー」シートをクリアし、見出しと集計結果を書き込む
 */
function writeSummary_(sheet, monthly) {
  sheet.clear();

  const values = [['月', '合計売上', '件数']].concat(
    monthly.map(m => [m.label, m.total, m.count])
  );
  // 「2026年1月」が日付として自動変換されないよう、A列を書式なしテキストにしてから書き込む
  sheet.getRange(1, 1, values.length, 1).setNumberFormat('@');
  sheet.getRange(1, 1, values.length, 3).setValues(values);
}

/**
 * 「月次サマリー」シートに月次売上推移の棒グラフを作り直す
 */
function rebuildChart_(sheet, monthCount) {
  // clear() ではグラフが消えないため、前回作成したグラフを明示的に削除する
  sheet.getCharts().forEach(chart => sheet.removeChart(chart));
  if (monthCount === 0) return;

  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(sheet.getRange(1, 1, monthCount + 1, 2))
    .setNumHeaders(1)
    .setPosition(1, 5, 0, 0)
    .setOption('title', '月次売上推移')
    .setOption('legend', { position: 'none' })
    .build();
  sheet.insertChart(chart);
}
