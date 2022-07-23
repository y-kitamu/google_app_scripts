// 売買損益計算に含めないシート名のリスト
const excludeSheetName: string[] = ["summary", "template", "worksheet"];

/**
 * summaryシートにスクリプト実行日時点の損益を追記する
 * 1列目 : 日付
 * 2列目 : totalの損益
 * 3列目~ : 個別銘柄（各シート）の損益
 */
const updateSummarySheet = () => {
  // 実行日の曜日を確認し、週末（株式市場が休み）の場合はスキップ
  const date = new Date();
  if (date.getDay() === 0 || date.getDay() === 1) {
    console.log("Skip update summary sheet because of weekend.");
    return;
  }

  const ss = SpreadsheetApp.getActive();
  const allSheets = ss.getSheets();

  const summarySheet = ss.getSheetByName("summary");
  const targetRow = getTargetRow(summarySheet);

  // 実行日の行が存在しない場合は追記
  if (summarySheet.getRange(targetRow, 1).getValue() === "") {
    const dateStr = getCurrentDateStr();
    summarySheet.getRange(targetRow, 1).setValue(dateStr);
    console.log(`Add row to summary sheet: ${dateStr}`);
  }

  // シートごとに損益を取得
  let totalProfitLoss = 0;
  for (const sheet of allSheets) {
    if (excludeSheetName.includes(sheet.getName())) {
      continue;
    }

    const profitLoss = getProfitLoss(sheet);
    totalProfitLoss += profitLoss;

    const targetColumn = getTargetColumn(summarySheet, sheet);
    summarySheet.getRange(targetRow, targetColumn).setValue(profitLoss);
  }

  summarySheet.getRange(targetRow, 2).setValue(totalProfitLoss);
};

const getProfitLoss = (sheet: GoogleAppsScript.Spreadsheet.Sheet): number => {
  try {
    return sheet.getRange("M1").getValue();
  } catch (e) {
    console.log("Failed to get profit loss from sheet: " + sheet.getName());
    console.log(e);
  }
  return 0;
};

const getTargetRow = (sheet: GoogleAppsScript.Spreadsheet.Sheet): number => {
  const lastRow = sheet.getLastRow();
  const lastRowValue = sheet.getRange(lastRow, 1).getRichTextValue();
  if (lastRowValue === null) {
    return lastRow;
  } else if (lastRowValue.getText() === getCurrentDateStr()) {
    return lastRow;
  }
  return lastRow + 1;
};

const getTargetColumn = (
  summarySheet: GoogleAppsScript.Spreadsheet.Sheet,
  sheet: GoogleAppsScript.Spreadsheet.Sheet
): number => {
  const lastColumn = summarySheet.getLastColumn();
  const headerValues = summarySheet.getRange(1, 1, 1, lastColumn).getValues();
  const stockName = sheet.getName();

  // 対象銘柄の列がすでに存在する場合はその列のindexを返す
  for (let i = 0; i < headerValues[0].length; i++) {
    if (stockName === headerValues[0][i]) {
      return i + 1;
    }
  }

  // 存在しない場合は、対象銘柄の列を末尾に追加しその列のindexを返す
  summarySheet.getRange(1, lastColumn + 1).setValue(stockName);
  return lastColumn + 1;
};

const getCurrentDateStr = (): string => {
  const date = new Date();
  const dateStr = `${date.getFullYear()}/${
    date.getMonth() + 1
  }/${date.getDate()}`;
  return dateStr;
};
