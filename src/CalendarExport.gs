/**
 * CalendarExport.gs - カレンダー連携 (iCal + Google Calendar)
 * シフトデータをiCal形式で出力し、カレンダーアプリに登録可能にする
 */

/**
 * iCal (VCALENDAR) 文字列を生成
 * @param {string} employeeNo - 社員番号
 * @param {string} targetMonth - 対象年月 (例: "2026-03")
 * @return {string} RFC 5545準拠のiCal文字列
 */
function generateICal(employeeNo, targetMonth) {
  const aggregated = aggregateByEmployee(targetMonth);
  const emp = aggregated.find(function(e) { return e.employeeNo === employeeNo; });

  if (!emp || emp.shifts.length === 0) {
    return '';
  }

  var lines = [];
  lines.push('BEGIN:VCALENDAR');
  lines.push('VERSION:2.0');
  lines.push('PRODID:-//ShiftNotification//JP');
  lines.push('X-WR-CALNAME:シフト予定 - ' + emp.name);
  lines.push('X-WR-TIMEZONE:Asia/Tokyo');
  lines.push('BEGIN:VTIMEZONE');
  lines.push('TZID:Asia/Tokyo');
  lines.push('BEGIN:STANDARD');
  lines.push('DTSTART:19700101T000000');
  lines.push('TZOFFSETFROM:+0900');
  lines.push('TZOFFSETTO:+0900');
  lines.push('END:STANDARD');
  lines.push('END:VTIMEZONE');

  for (var i = 0; i < emp.shifts.length; i++) {
    var shift = emp.shifts[i];
    var parsed = parseTimeSlotToHours(shift.timeSlot);
    if (!parsed) continue;

    var dtStart = formatICalDateTime(shift.date, parsed.startHour);
    var dtEnd;
    if (parsed.overnight) {
      // 翌日の終了時刻を計算
      dtEnd = formatICalDateTimeNextDay(shift.date, parsed.endHour);
    } else {
      dtEnd = formatICalDateTime(shift.date, parsed.endHour);
    }

    var uid = 'shift-' + employeeNo + '-' + dtStart + '@shift-notification';

    lines.push('BEGIN:VEVENT');
    lines.push('DTSTART;TZID=Asia/Tokyo:' + dtStart);
    lines.push('DTEND;TZID=Asia/Tokyo:' + dtEnd);
    lines.push('SUMMARY:' + shift.facility + ' (' + shift.timeSlot + ')');
    lines.push('LOCATION:' + shift.facility);
    lines.push('UID:' + uid);
    lines.push('END:VEVENT');
  }

  lines.push('END:VCALENDAR');
  return lines.join('\r\n');
}

/**
 * 時間帯文字列をパースして開始・終了時刻を取得
 * @param {string} timeSlot - 時間帯 (例: "6時～9時", "17時～22時", "22時～")
 * @return {Object|null} {startHour, endHour, overnight}
 */
function parseTimeSlotToHours(timeSlot) {
  if (!timeSlot) return null;

  // "XX時～YY時" パターン
  var match = timeSlot.match(/(\d+)時～(\d+)時/);
  if (match) {
    var start = parseInt(match[1], 10);
    var end = parseInt(match[2], 10);
    return { startHour: start, endHour: end, overnight: end < start };
  }

  // "XX時～" パターン (終了時刻なし → 翌朝6時まで)
  var matchOpen = timeSlot.match(/(\d+)時～$/);
  if (matchOpen) {
    var startHour = parseInt(matchOpen[1], 10);
    return { startHour: startHour, endHour: 6, overnight: true };
  }

  return null;
}

/**
 * 日付と時刻からiCal形式の日時文字列を生成
 * @param {string} dateStr - 日付 (例: "2026/03/01")
 * @param {number} hour - 時刻 (0-23)
 * @return {string} iCal形式日時 (例: "20260301T170000")
 */
function formatICalDateTime(dateStr, hour) {
  var parts = dateStr.split('/');
  var year = parts[0];
  var month = String(parts[1]).padStart(2, '0');
  var day = String(parts[2]).padStart(2, '0');
  var hourStr = String(hour).padStart(2, '0');
  return year + month + day + 'T' + hourStr + '0000';
}

/**
 * 翌日の日時文字列を生成 (夜勤の終了時刻用)
 * @param {string} dateStr - 日付 (例: "2026/03/01")
 * @param {number} hour - 時刻 (0-23)
 * @return {string} iCal形式日時 (翌日)
 */
function formatICalDateTimeNextDay(dateStr, hour) {
  var parts = dateStr.split('/');
  var date = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
  date.setDate(date.getDate() + 1);

  var year = String(date.getFullYear());
  var month = String(date.getMonth() + 1).padStart(2, '0');
  var day = String(date.getDate()).padStart(2, '0');
  var hourStr = String(hour).padStart(2, '0');
  return year + month + day + 'T' + hourStr + '0000';
}

/**
 * iCalダウンロード用の署名トークンを生成
 * HMAC-SHA256(ICAL_SECRET, empNo + ":" + month) の先頭16文字
 * @param {string} empNo - 社員番号
 * @param {string} month - 対象年月 (例: "2026-03")
 * @return {string} 署名トークン
 */
function generateICalToken(empNo, month) {
  var secret = PropertiesService.getScriptProperties().getProperty('ICAL_SECRET');
  if (!secret) {
    // ICAL_SECRET未設定時は自動生成して保存
    secret = Utilities.getUuid();
    PropertiesService.getScriptProperties().setProperty('ICAL_SECRET', secret);
    Logger.log('ICAL_SECRET を自動生成しました');
  }
  var signature = Utilities.computeHmacSha256Signature(empNo + ':' + month, secret);
  return Utilities.base64EncodeWebSafe(signature).substring(0, 16);
}

/**
 * iCalダウンロード用の署名トークンを検証
 * @param {string} empNo - 社員番号
 * @param {string} month - 対象年月
 * @param {string} token - 検証するトークン
 * @return {boolean} 検証結果
 */
function verifyICalToken(empNo, month, token) {
  if (!token) return false;
  var expected = generateICalToken(empNo, month);
  return expected === token;
}

/**
 * iCalダウンロードURLを構築
 * @param {string} empNo - 社員番号
 * @param {string} targetMonth - 対象年月
 * @return {string|null} iCal URL (構築失敗時はnull)
 */
function buildICalUrl(empNo, targetMonth) {
  try {
    var baseUrl = ScriptApp.getService().getUrl();
    var token = generateICalToken(empNo, targetMonth);
    return baseUrl + '?action=ical&empNo=' + empNo + '&month=' + targetMonth + '&token=' + token;
  } catch (e) {
    Logger.log('iCal URL構築エラー: ' + e.toString());
    return null;
  }
}

/**
 * 管理者向け: Google Calendarにシフトを一括登録
 * CalendarAppを使用して「シフト予定」カレンダーに全員分のシフトを登録
 * @param {string} targetMonth - 対象年月 (例: "2026-03")
 */
function syncToGoogleCalendar(targetMonth) {
  var calendarName = 'シフト予定';
  var calendars = CalendarApp.getCalendarsByName(calendarName);
  var calendar;

  if (calendars.length === 0) {
    calendar = CalendarApp.createCalendar(calendarName, {
      timeZone: 'Asia/Tokyo',
      color: CalendarApp.Color.GREEN
    });
    Logger.log('カレンダー「' + calendarName + '」を作成しました');
  } else {
    calendar = calendars[0];
  }

  var aggregated = aggregateByEmployee(targetMonth);
  var eventCount = 0;

  for (var i = 0; i < aggregated.length; i++) {
    var emp = aggregated[i];

    for (var j = 0; j < emp.shifts.length; j++) {
      var shift = emp.shifts[j];
      var parsed = parseTimeSlotToHours(shift.timeSlot);
      if (!parsed) continue;

      var parts = shift.date.split('/');
      var year = parseInt(parts[0], 10);
      var month = parseInt(parts[1], 10) - 1;
      var day = parseInt(parts[2], 10);

      var startTime = new Date(year, month, day, parsed.startHour, 0, 0);
      var endTime;
      if (parsed.overnight) {
        endTime = new Date(year, month, day + 1, parsed.endHour, 0, 0);
      } else {
        endTime = new Date(year, month, day, parsed.endHour, 0, 0);
      }

      var title = emp.name + ' - ' + shift.facility + ' (' + shift.timeSlot + ')';

      calendar.createEvent(title, startTime, endTime, {
        location: shift.facility,
        description: '社員番号: ' + emp.employeeNo + '\n施設: ' + shift.facility + '\n時間帯: ' + shift.timeSlot
      });

      eventCount++;
    }
  }

  Logger.log('Googleカレンダー登録完了: ' + eventCount + '件のイベントを登録');
  return eventCount;
}
