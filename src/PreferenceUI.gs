/**
 * PreferenceUI.gs - シフト希望入力 LINE Flex Message UI
 * 職員がLINEからシフト希望を入力するためのUI生成
 *
 * フロー:
 * 1. pref_start → カレンダー表示（月の日付ボタン一覧）
 * 2. pref_date  → 希望種別選択（希望/NG/どちらでも）
 * 3. pref_type  → 時間帯選択（任意）
 * 4. pref_confirm → 確認画面→保存
 * 5. pref_done  → 完了 or 続けて入力
 */

/**
 * シフト希望入力開始メッセージを生成
 * @param {string} yearMonth - 対象年月 (2026-03)
 * @param {string} employeeName - 従業員名
 * @param {Array<Object>} existingPrefs - 既存の希望データ
 * @return {Object} LINE Flex Message
 */
function buildPreferenceStartMessage(yearMonth, employeeName, existingPrefs) {
  var parts = yearMonth.split('-');
  var year = parseInt(parts[0], 10);
  var month = parseInt(parts[1], 10);
  var displayMonth = year + '年' + month + '月';
  var daysInMonth = new Date(year, month, 0).getDate();

  // 既存希望をDateでマップ化
  var prefMap = {};
  if (existingPrefs) {
    existingPrefs.forEach(function(p) {
      prefMap[p.date] = p.type;
    });
  }

  // カレンダーグリッド生成（7列×最大6行）
  var firstDow = new Date(year, month - 1, 1).getDay(); // 0=日
  var calendarRows = buildCalendarGrid_(year, month, daysInMonth, firstDow, prefMap, yearMonth);

  var bubbles = [];

  // バブル1: カレンダー
  bubbles.push({
    type: 'bubble',
    size: 'mega',
    header: {
      type: 'box',
      layout: 'vertical',
      contents: [
        {
          type: 'text',
          text: displayMonth + ' シフト希望入力',
          weight: 'bold',
          size: 'md',
          color: '#FFFFFF'
        },
        {
          type: 'text',
          text: employeeName + 'さん',
          size: 'sm',
          color: '#FFFFFF',
          margin: 'xs'
        }
      ],
      backgroundColor: '#1DB446',
      paddingAll: '12px'
    },
    body: {
      type: 'box',
      layout: 'vertical',
      contents: [
        buildCalendarHeader_(),
        { type: 'separator', margin: 'sm' },
        ...calendarRows
      ],
      paddingAll: '8px',
      spacing: 'none'
    },
    footer: {
      type: 'box',
      layout: 'vertical',
      contents: [
        {
          type: 'text',
          text: '日付をタップして希望を入力',
          size: 'xs',
          color: '#999999',
          align: 'center'
        },
        {
          type: 'box',
          layout: 'horizontal',
          contents: [
            { type: 'text', text: '🟢希望', size: 'xxs', color: '#27AE60', flex: 1, align: 'center' },
            { type: 'text', text: '🔴NG', size: 'xxs', color: '#E74C3C', flex: 1, align: 'center' },
            { type: 'text', text: '⚪未入力', size: 'xxs', color: '#999999', flex: 1, align: 'center' }
          ],
          margin: 'sm'
        }
      ],
      paddingAll: '8px'
    }
  });

  // バブル2: 入力済み一覧 + 送信ボタン
  var existingCount = existingPrefs ? existingPrefs.length : 0;
  var summaryContents = buildExistingSummary_(existingPrefs, yearMonth);

  bubbles.push({
    type: 'bubble',
    size: 'kilo',
    header: {
      type: 'box',
      layout: 'vertical',
      contents: [{
        type: 'text',
        text: '入力済み: ' + existingCount + '件',
        weight: 'bold',
        size: 'md',
        color: '#FFFFFF'
      }],
      backgroundColor: '#34495E',
      paddingAll: '12px'
    },
    body: {
      type: 'box',
      layout: 'vertical',
      contents: summaryContents,
      paddingAll: '12px'
    },
    footer: {
      type: 'box',
      layout: 'vertical',
      contents: [
        {
          type: 'button',
          action: {
            type: 'postback',
            label: '全てクリアして再入力',
            data: 'action=pref_clear&month=' + yearMonth
          },
          style: 'secondary',
          height: 'sm',
          margin: 'sm'
        }
      ],
      paddingAll: '12px'
    }
  });

  // バブル3: LIFF Webアプリ入力へ誘導
  try {
    var liffId = getSettingValue(SETTING_KEYS.LIFF_ID);
    if (liffId) {
      bubbles.push({
        type: 'bubble',
        size: 'kilo',
        header: {
          type: 'box',
          layout: 'vertical',
          contents: [{
            type: 'text',
            text: 'Webアプリで入力',
            weight: 'bold',
            size: 'md',
            color: '#FFFFFF'
          }],
          backgroundColor: '#2980B9',
          paddingAll: '12px'
        },
        body: {
          type: 'box',
          layout: 'vertical',
          contents: [
            {
              type: 'text',
              text: 'より詳しく入力したい場合はWebアプリからどうぞ',
              wrap: true,
              size: 'sm',
              color: '#555555'
            },
            {
              type: 'text',
              text: '・カレンダーで一覧入力\n・時間帯の指定\n・一括提出',
              wrap: true,
              size: 'sm',
              color: '#888888',
              margin: 'md'
            }
          ],
          paddingAll: '12px'
        },
        footer: {
          type: 'box',
          layout: 'vertical',
          contents: [{
            type: 'button',
            action: {
              type: 'uri',
              label: 'Webアプリで入力する',
              uri: 'https://liff.line.me/' + liffId + '/preference/?month=' + yearMonth
            },
            style: 'primary',
            color: '#2980B9',
            height: 'sm'
          }],
          paddingAll: '12px'
        }
      });
    }
  } catch (e) {
    // LIFF_ID が取得できなければスキップ
    Logger.log('LIFF bubble skipped: ' + e.toString());
  }

  return {
    type: 'flex',
    altText: displayMonth + ' シフト希望入力',
    contents: {
      type: 'carousel',
      contents: bubbles
    }
  };
}

/**
 * 日付選択後の希望種別選択メッセージ
 * @param {string} yearMonth - 対象年月
 * @param {string} dateStr - 選択日 (YYYY/MM/DD)
 * @return {Object} LINE Flex Message
 */
function buildTypeSelectionMessage(yearMonth, dateStr) {
  var dateParts = dateStr.split('/');
  var dayNum = parseInt(dateParts[2], 10);
  var dow = getDayOfWeek(parseInt(dateParts[0]), parseInt(dateParts[1]), dayNum);
  var displayDate = parseInt(dateParts[1], 10) + '月' + dayNum + '日(' + dow + ')';

  return {
    type: 'flex',
    altText: displayDate + ' の希望を選択',
    contents: {
      type: 'bubble',
      size: 'kilo',
      header: {
        type: 'box',
        layout: 'vertical',
        contents: [{
          type: 'text',
          text: displayDate,
          weight: 'bold',
          size: 'lg',
          color: '#FFFFFF',
          align: 'center'
        }],
        backgroundColor: '#1DB446',
        paddingAll: '12px'
      },
      body: {
        type: 'box',
        layout: 'vertical',
        contents: [
          {
            type: 'button',
            action: {
              type: 'postback',
              label: '出勤希望',
              data: 'action=pref_type&month=' + yearMonth + '&date=' + dateStr + '&type=' + PREF_TYPE.WANT,
              displayText: displayDate + ' → 出勤希望'
            },
            style: 'primary',
            color: '#27AE60',
            height: 'sm'
          },
          {
            type: 'button',
            action: {
              type: 'postback',
              label: 'NG（出勤不可）',
              data: 'action=pref_type&month=' + yearMonth + '&date=' + dateStr + '&type=' + PREF_TYPE.NG,
              displayText: displayDate + ' → NG'
            },
            style: 'primary',
            color: '#E74C3C',
            height: 'sm',
            margin: 'sm'
          },
          {
            type: 'button',
            action: {
              type: 'postback',
              label: 'どちらでも',
              data: 'action=pref_type&month=' + yearMonth + '&date=' + dateStr + '&type=' + PREF_TYPE.EITHER,
              displayText: displayDate + ' → どちらでも'
            },
            style: 'secondary',
            height: 'sm',
            margin: 'sm'
          },
          { type: 'separator', margin: 'md' },
          {
            type: 'button',
            action: {
              type: 'postback',
              label: 'この日の希望を削除',
              data: 'action=pref_delete&month=' + yearMonth + '&date=' + dateStr,
              displayText: displayDate + ' の希望を削除'
            },
            style: 'secondary',
            color: '#95A5A6',
            height: 'sm',
            margin: 'md'
          }
        ],
        paddingAll: '12px',
        spacing: 'none'
      }
    }
  };
}

/**
 * 希望保存完了メッセージ
 * @param {string} dateStr - 日付
 * @param {string} type - 希望種別
 * @param {string} yearMonth - 対象年月
 * @return {Object} LINE message with Quick Reply
 */
function buildPreferenceSavedMessage(dateStr, type, yearMonth) {
  var dateParts = dateStr.split('/');
  var dayNum = parseInt(dateParts[2], 10);
  var dow = getDayOfWeek(parseInt(dateParts[0]), parseInt(dateParts[1]), dayNum);
  var displayDate = parseInt(dateParts[1], 10) + '月' + dayNum + '日(' + dow + ')';

  var typeEmoji = type === PREF_TYPE.WANT ? '🟢' : type === PREF_TYPE.NG ? '🔴' : '⚪';
  var typeLabel = type === PREF_TYPE.WANT ? '出勤希望' : type === PREF_TYPE.NG ? 'NG' : 'どちらでも';

  return {
    type: 'text',
    text: typeEmoji + ' ' + displayDate + ' → ' + typeLabel + '\n保存しました！',
    quickReply: {
      items: [
        {
          type: 'action',
          action: {
            type: 'postback',
            label: '続けて入力',
            data: 'action=pref_start&month=' + yearMonth,
            displayText: 'シフト希望入力を続ける'
          }
        },
        {
          type: 'action',
          action: {
            type: 'postback',
            label: '入力完了',
            data: 'action=pref_finish&month=' + yearMonth,
            displayText: 'シフト希望入力を完了'
          }
        }
      ]
    }
  };
}

/**
 * 希望入力完了メッセージ
 * @param {string} yearMonth - 対象年月
 * @param {number} totalCount - 入力件数
 * @return {Object} LINE Flex Message
 */
function buildPreferenceCompleteMessage(yearMonth, totalCount) {
  var parts = yearMonth.split('-');
  var displayMonth = parts[0] + '年' + parseInt(parts[1], 10) + '月';

  return {
    type: 'flex',
    altText: displayMonth + ' シフト希望入力完了',
    contents: {
      type: 'bubble',
      size: 'kilo',
      body: {
        type: 'box',
        layout: 'vertical',
        contents: [
          {
            type: 'text',
            text: '✅ 入力完了',
            weight: 'bold',
            size: 'lg',
            align: 'center',
            color: '#27AE60'
          },
          {
            type: 'text',
            text: displayMonth + 'のシフト希望を' + totalCount + '件提出しました。',
            wrap: true,
            size: 'sm',
            margin: 'md',
            align: 'center',
            color: '#555555'
          },
          {
            type: 'text',
            text: '変更がある場合はもう一度入力できます。',
            wrap: true,
            size: 'xs',
            margin: 'md',
            align: 'center',
            color: '#999999'
          }
        ],
        paddingAll: '20px'
      },
      footer: {
        type: 'box',
        layout: 'vertical',
        contents: [{
          type: 'button',
          action: {
            type: 'postback',
            label: '希望を修正する',
            data: 'action=pref_start&month=' + yearMonth,
            displayText: 'シフト希望を修正'
          },
          style: 'secondary',
          height: 'sm'
        }],
        paddingAll: '12px'
      }
    }
  };
}

// ===== Private helpers =====

/**
 * カレンダーヘッダー（曜日）行を生成
 */
function buildCalendarHeader_() {
  var days = ['日', '月', '火', '水', '木', '金', '土'];
  var colors = ['#E74C3C', '#333', '#333', '#333', '#333', '#333', '#3498DB'];

  return {
    type: 'box',
    layout: 'horizontal',
    contents: days.map(function(d, i) {
      return {
        type: 'text',
        text: d,
        size: 'xs',
        weight: 'bold',
        color: colors[i],
        align: 'center',
        flex: 1
      };
    }),
    margin: 'md'
  };
}

/**
 * カレンダーグリッドを生成
 */
function buildCalendarGrid_(year, month, daysInMonth, firstDow, prefMap, yearMonth) {
  var rows = [];
  var day = 1;

  for (var week = 0; week < 6; week++) {
    var cells = [];

    for (var dow = 0; dow < 7; dow++) {
      if ((week === 0 && dow < firstDow) || day > daysInMonth) {
        cells.push({
          type: 'text',
          text: ' ',
          size: 'sm',
          align: 'center',
          flex: 1
        });
      } else {
        var dateStr = year + '/' + String(month).padStart(2, '0') + '/' + String(day).padStart(2, '0');
        var prefType = prefMap[dateStr];

        // ボタンの色を希望種別で変える
        var bgColor = '#F0F0F0'; // 未入力
        var textColor = '#333333';
        if (prefType === PREF_TYPE.WANT) {
          bgColor = '#D5F5E3';
          textColor = '#27AE60';
        } else if (prefType === PREF_TYPE.NG) {
          bgColor = '#FADBD8';
          textColor = '#E74C3C';
        } else if (prefType === PREF_TYPE.EITHER) {
          bgColor = '#FDEBD0';
          textColor = '#F39C12';
        }

        // 土日は文字色を変える（未入力の場合のみ）
        if (!prefType) {
          if (dow === 0) textColor = '#E74C3C';
          else if (dow === 6) textColor = '#3498DB';
        }

        cells.push({
          type: 'box',
          layout: 'vertical',
          contents: [{
            type: 'text',
            text: String(day),
            size: 'sm',
            align: 'center',
            color: textColor,
            weight: prefType ? 'bold' : 'regular'
          }],
          flex: 1,
          backgroundColor: bgColor,
          cornerRadius: '4px',
          action: {
            type: 'postback',
            data: 'action=pref_date&month=' + yearMonth + '&date=' + dateStr,
            displayText: parseInt(month, 10) + '月' + day + '日'
          },
          paddingAll: '4px'
        });

        day++;
      }
    }

    if (day > daysInMonth && week > 0 && cells.every(function(c) { return c.type === 'text' && c.text === ' '; })) {
      break; // 空行はスキップ
    }

    rows.push({
      type: 'box',
      layout: 'horizontal',
      contents: cells,
      margin: 'xs',
      spacing: 'xs'
    });
  }

  return rows;
}

/**
 * 既存希望のサマリー表示を生成
 */
function buildExistingSummary_(existingPrefs, yearMonth) {
  if (!existingPrefs || existingPrefs.length === 0) {
    return [{
      type: 'text',
      text: 'まだ希望が入力されていません。\n左のカレンダーから日付をタップして入力してください。',
      wrap: true,
      size: 'sm',
      color: '#999999'
    }];
  }

  // 種別ごとの件数
  var wantCount = 0, ngCount = 0, eitherCount = 0;
  existingPrefs.forEach(function(p) {
    if (p.type === PREF_TYPE.WANT) wantCount++;
    else if (p.type === PREF_TYPE.NG) ngCount++;
    else eitherCount++;
  });

  var contents = [
    {
      type: 'box',
      layout: 'horizontal',
      contents: [
        { type: 'text', text: '🟢 出勤希望', size: 'sm', flex: 3 },
        { type: 'text', text: wantCount + '日', size: 'sm', align: 'end', weight: 'bold', flex: 1 }
      ]
    },
    {
      type: 'box',
      layout: 'horizontal',
      contents: [
        { type: 'text', text: '🔴 NG', size: 'sm', flex: 3 },
        { type: 'text', text: ngCount + '日', size: 'sm', align: 'end', weight: 'bold', flex: 1 }
      ],
      margin: 'sm'
    },
    {
      type: 'box',
      layout: 'horizontal',
      contents: [
        { type: 'text', text: '⚪ どちらでも', size: 'sm', flex: 3 },
        { type: 'text', text: eitherCount + '日', size: 'sm', align: 'end', weight: 'bold', flex: 1 }
      ],
      margin: 'sm'
    }
  ];

  // NG日の一覧（重要なので表示）
  if (ngCount > 0) {
    var ngDates = existingPrefs
      .filter(function(p) { return p.type === PREF_TYPE.NG; })
      .map(function(p) {
        var parts = p.date.split('/');
        return parseInt(parts[2], 10) + '日';
      })
      .join(', ');

    contents.push({ type: 'separator', margin: 'md' });
    contents.push({
      type: 'text',
      text: 'NG日: ' + ngDates,
      wrap: true,
      size: 'xs',
      color: '#E74C3C',
      margin: 'md'
    });
  }

  return contents;
}
