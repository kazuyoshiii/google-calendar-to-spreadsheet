// 設定オブジェクト
const CONFIG = {
  calendarId: "YOUR_EMAIL_ADDRESS",
  spreadsheetId: "YOUR_SPREADSHEET_ID",
  excludeKeywords: ["Home", "KEYWORD2"],
  year: 2024,
  month: 12,
  dateRange: {
    // 月初から月末の1ヶ月
    start: function (config) {
      return new Date(config.year, config.month - 1, 1);
    },
    end: function (config) {
      const year = config.month === 12 ? config.year + 1 : config.year;
      return new Date(year, config.month, 0);
    },
  },
  headers: [
    "Event ID",
    "Month",
    "Title",
    "Start Date",
    "End Date",
    "Duration (hours)",
    "Description",
  ],
};

/**
 * メイン実行関数
 */
function updateSpreadsheetFromCalendar() {
  try {
    const events = getFilteredCalendarEvents();
    const sheet = initializeSpreadsheet();
    writeEventsToSheet(events, sheet);
    Logger.log("スプレッドシートの更新が完了しました。");
  } catch (error) {
    Logger.error("エラーが発生しました: " + error.toString());
    throw error;
  }
}

/**
 * フィルター済みのカレンダーイベントを取得
 */
function getFilteredCalendarEvents() {
  const calendar = CalendarApp.getCalendarById(CONFIG.calendarId);
  if (!calendar) {
    throw new Error("カレンダーが見つかりません");
  }

  const events = calendar.getEvents(
    CONFIG.dateRange.start(CONFIG),
    CONFIG.dateRange.end(CONFIG)
  );
  return events.filter(
    (event) =>
      !CONFIG.excludeKeywords.some((keyword) =>
        event.getTitle().includes(keyword)
      )
  );
}

/**
 * スプレッドシートの初期化
 */
function initializeSpreadsheet() {
  const spreadsheet = SpreadsheetApp.openById(CONFIG.spreadsheetId);
  if (!spreadsheet) {
    throw new Error("スプレッドシートが見つかりません");
  }

  const sheet = spreadsheet.getSheets()[0];
  sheet.clear();
  sheet.appendRow(CONFIG.headers);
  return sheet;
}

/**
 * イベントデータをスプレッドシートに書き込み
 */
function writeEventsToSheet(events, sheet) {
  events.forEach((event) => {
    const startTime = event.getStartTime();
    const endTime = event.getEndTime();
    const duration =
      (endTime.getTime() - startTime.getTime()) / (1000 * 60 * 60);

    sheet.appendRow([
      event.getId(),
      `${CONFIG.year}年${CONFIG.month}月`,
      event.getTitle(),
      startTime,
      endTime,
      duration,
      event.getDescription(),
    ]);
  });
}
