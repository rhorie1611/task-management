function calendarOnEdit_(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== "管理シート") return;

  const row = e.range.getRow();
  if (row < 7) return;

  const calendar = CalendarApp.getCalendarById("916b37fd8ea730ceae1f54f1ea729ac6f96c033fdfec9304ae3bccfed3af6579@group.calendar.google.com");

  const checkbox1 = sheet.getRange(row, 2).getValue(); // B列
  const checkbox2 = sheet.getRange(row, 3).getValue(); // C列
  const date      = sheet.getRange(row, 5).getValue(); // E列
  const time      = sheet.getRange(row, 6).getValue(); // F列
  const task      = sheet.getRange(row, 7).getValue(); // G列
  const note      = sheet.getRange(row, 9).getValue(); // I列
  const idCell    = sheet.getRange(row, 10);           // J列
  const eventId   = idCell.getValue();
  const skipCreate = sheet.getRange(row, 11).getValue(); // K列

  // --- 削除 ---
  if (checkbox1 === true && eventId) {
    try {
      const event = calendar.getEventById(eventId);
      if (event) event.deleteEvent();
      idCell.clearContent();
    } catch (e) {
      Logger.log("削除エラー: " + e);
    }
    return;
  }

  // --- 作成 ---
  if (checkbox2 === true && date && time && task && !eventId && !skipCreate) {
    const start = new Date(date);
    start.setHours(4, 0, 0);
    const end = new Date(date);
    end.setHours(6, 0, 0);

    const timeText = Utilities.formatDate(
      new Date(time),
      Session.getScriptTimeZone(),
      "HH:mm"
    );

    const title = `【${timeText}締切】${task}`;
    const event = calendar.createEvent(title, start, end, {
      description: note
    });

    idCell.setValue(event.getId());
    SpreadsheetApp.flush();
  }
}
