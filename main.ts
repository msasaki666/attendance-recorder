import type { SlashCommand } from "@slack/bolt";

const commandMappings = {
  start: "/start",
  stop: "/stop",
  edit: "/edit",
  show: "/show",
} as const;

function getFormattedCurrentDateTime(timezone: string): string {
  const now = new Date();
  return Utilities.formatDate(now, timezone, "yyyy/MM/dd HH:mm:ss");
}

function getNextRowIndex(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  header: GoogleAppsScript.Spreadsheet.Range
): number {
  const offset = 2;
  const idColumnRange = sheet.getRange(
    offset,
    header.getColumn(),
    sheet.getLastRow(),
    1
  );
  return (
    idColumnRange
      .getValues()
      .flatMap((id) => {
        return id;
      })
      .findIndex((id) => {
        return id === "";
      }) + offset
  );
}

function getLastInsertedRowIndex(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  header: GoogleAppsScript.Spreadsheet.Range
): number {
  return getNextRowIndex(sheet, header) - 1;
}

function getVerificationToken(): string | null {
  return PropertiesService.getScriptProperties().getProperty(
    "VERIFICATION_TOKEN"
  );
}

function getSlackBotToken(): string | null {
  return PropertiesService.getScriptProperties().getProperty("SLACK_BOT_TOKEN");
}

export function doPost(
  e: GoogleAppsScript.Events.DoPost
): GoogleAppsScript.Content.TextOutput {
  const body = decodeURIComponent(e.postData.contents)
    .split("&")
    .reduce((acm, kv: string) => {
      const [k, v] = kv.split("=");
      acm[k] = v;
      return acm;
    }, {} as Record<string, string>) as SlashCommand;
  const verificationToken = getVerificationToken();
  if (!verificationToken) {
    throw new Error("VERIFICATION_TOKEN not found");
  }
  if (verificationToken != body.token) {
    return ContentService.createTextOutput("verificationToken is invalid");
  }
  const slackBotToken = getSlackBotToken();
  if (!slackBotToken) {
    throw new Error("SLACK_BOT_TOKEN not found");
  }

  const availableCommands = Object.values(commandMappings as Record<string, string>)
  if (
    !availableCommands.includes(
      body.command
    )
  ) {
    return ContentService.createTextOutput(
      `command: ${body.command} is not implemented. available commands are ${availableCommands}`
    );
  }

  // idを指定していないstopリクエストの時は、usernameから最後の行を探し出して追加する
  // idを指定してeditできる
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userMappingSheet = ss.getSheetByName("user_mapping");
  if (!userMappingSheet) {
    return ContentService.createTextOutput("can't get user_mapping sheet");
  }
  const userMappingSheetLastColumnIndex = userMappingSheet.getLastColumn();
  const userMappingSheetFirstRow = userMappingSheet.getRange(
    1,
    1,
    1,
    userMappingSheetLastColumnIndex
  );
  const slackUserIDHeader = userMappingSheetFirstRow
    .createTextFinder("slack_user_id")
    .findAll()[0];
  if (!slackUserIDHeader) {
    return ContentService.createTextOutput("can't get slackUserID header");
  }
  const userMappingSheetLastRowIndex = userMappingSheet.getLastRow();
  const slackUserIDsRange = userMappingSheet.getRange(
    1,
    slackUserIDHeader.getColumn(),
    userMappingSheetLastRowIndex,
    1
  );
  const slackUserID = slackUserIDsRange
    .createTextFinder(body.user_id)
    .findAll()[0];
  if (!slackUserID) {
    return ContentService.createTextOutput(
      "can't get slackUserID inputted range"
    );
  }
  const attendanceSheetName = `${slackUserID.getValue()}_attendance`;
  const attendanceSheet = ss.getSheetByName(attendanceSheetName);
  if (!attendanceSheet) {
    return ContentService.createTextOutput("can't get attendance sheet");
  }

  const attendanceSheetLastColumnIndex = attendanceSheet.getLastColumn();
  const firstRow = attendanceSheet.getRange(
    1,
    1,
    1,
    attendanceSheetLastColumnIndex
  );
  const idHeader = firstRow.createTextFinder("id").findAll()[0];
  if (!idHeader) {
    return ContentService.createTextOutput("can't get id header");
  }
  const startedAtHeader = firstRow.createTextFinder("started_at").findAll()[0];
  if (!startedAtHeader) {
    return ContentService.createTextOutput("can't get started_at header");
  }
  const stoppedAtHeader = firstRow.createTextFinder("stopped_at").findAll()[0];
  if (!stoppedAtHeader) {
    return ContentService.createTextOutput("can't get stopped_at header");
  }

  switch (body.command) {
    case commandMappings.start:
      const nextRowIndex = getNextRowIndex(attendanceSheet, idHeader);
      // 値のセット
      // slackに返す
      attendanceSheet
        .getRange(nextRowIndex, idHeader.getColumn())
        .setValue(Utilities.getUuid());
      attendanceSheet
        .getRange(nextRowIndex, startedAtHeader.getColumn())
        .setValue(getFormattedCurrentDateTime(ss.getSpreadsheetTimeZone()));
      break;
    case commandMappings.stop:
      // const lastInsertedRowIndex = getLastInsertedRowIndex(
      //   attendanceSheet,
      //   idHeader
      // );
      // const lastStartedAtRange = attendanceSheet.getRange(
      //   lastInsertedRowIndex,
      //   startedAtHeader.getColumn()
      // );
      // if (lastStartedAtRange.isBlank()) {
      //   return ContentService.createTextOutput("not started");
      // }
      // const nextInsertStoppedAtIndex = attendanceSheet.getRange(
      //   lastInsertedRowIndex,
      //   stoppedAtHeader.getColumn()
      // );
      // nextInsertStoppedAtIndex.setValue(
      //   getFormattedCurrentDateTime(ss.getSpreadsheetTimeZone())
      // );
      break;
    case commandMappings.edit:
      break;
    case commandMappings.show:
      break;
    default:
      throw new Error("unknown command")
  }

  return ContentService.createTextOutput("success");
}
