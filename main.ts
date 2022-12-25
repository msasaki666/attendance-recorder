const getFormattedDate = (): string => {
  const now = new Date();
  return Utilities.formatDate(now, "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
};

const getNextRowIndex = (sheet: GoogleAppsScript.Spreadsheet.Sheet): number => {
  const lastRowIndex = sheet.getLastRow();
  return lastRowIndex + 1;
};

export const hello = () => {
  // idを指定していないstopリクエストの時は、usernameから最後の行を探し出して追加する
  // idを指定してeditできる
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const attendanceSheet = ss.getSheetByName("attendance");
  if (!attendanceSheet) {
    throw new Error("can't get attendance sheet");
  }
  const lastColumnIndex = attendanceSheet.getLastColumn();
  const firstRow = attendanceSheet.getRange(1, 1, 1, lastColumnIndex);
  const id = firstRow.createTextFinder("id").findAll()[0];
  if (!id) {
    throw new Error("can't get id inputted range");
  }
  const startedAt = firstRow.createTextFinder("started_at").findAll()[0];
  if (!startedAt) {
    throw new Error("can't get started_at inputted range");
  }
  const stoppedAt = firstRow.createTextFinder("stopped_at").findAll()[0];
  if (!stoppedAt) {
    throw new Error("can't get stopped_at inputted range");
  }
  const nextRowIndex = getNextRowIndex(attendanceSheet);
  // 値のセット
  // slackに返す
  const rowID = Utilities.getUuid();
  attendanceSheet.getRange(nextRowIndex, id.getColumn()).setValue(rowID);
  const formattedDate = getFormattedDate();
  attendanceSheet
    .getRange(nextRowIndex, startedAt.getColumn())
    .setValue(formattedDate);
};
