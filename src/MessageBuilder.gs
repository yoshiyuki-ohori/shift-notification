/**
 * MessageBuilder.gs - LINE Flex Message生成
 * 従業員ごとのシフト情報をFlex Messageに変換
 */

/**
 * シフト通知のFlex Messageを生成
 * @param {string} targetMonth - 対象年月 (例: "2026-03")
 * @param {string} employeeName - 従業員名
 * @param {Array<Object>} shifts - シフト配列 [{date, timeSlot, facility}]
 * @param {string} [employeeNo] - 社員番号 (カレンダーURL用、省略可)
 * @return {Object} LINE Flex Messageオブジェクト
 */
function buildShiftFlexMessage(targetMonth, employeeName, shifts, employeeNo) {
  const [year, month] = targetMonth.split('-').map(Number);
  const displayMonth = year + '年' + month + '月';

  // iCal URL構築 (HMAC署名付き)
  var icalUrl = employeeNo ? buildICalUrl(employeeNo, targetMonth) : null;

  // シフトを日付でグループ化して表示行を作成
  const bodyContents = buildShiftRows(shifts);

  // Flex Messageが大きすぎる場合は分割
  if (shifts.length > 60) {
    return buildMultiBubbleMessage(displayMonth, employeeName, shifts, employeeNo);
  }

  return {
    type: 'flex',
    altText: displayMonth + ' シフト予定 - ' + employeeName + 'さん',
    contents: {
      type: 'bubble',
      size: 'mega',
      header: buildHeader(displayMonth, employeeName),
      body: {
        type: 'box',
        layout: 'vertical',
        contents: [
          buildTableHeader(),
          { type: 'separator', margin: 'sm' },
          ...bodyContents
        ],
        paddingAll: '12px',
        spacing: 'none'
      },
      footer: buildFooter(shifts.length, icalUrl)
    }
  };
}

/**
 * ヘッダー部分を生成
 * @param {string} displayMonth - 表示用年月
 * @param {string} employeeName - 従業員名
 * @return {Object} ヘッダーコンポーネント
 */
function buildHeader(displayMonth, employeeName) {
  return {
    type: 'box',
    layout: 'vertical',
    contents: [
      {
        type: 'text',
        text: displayMonth + ' シフト予定',
        weight: 'bold',
        size: 'lg',
        color: '#FFFFFF'
      },
      {
        type: 'text',
        text: employeeName + ' さん',
        size: 'md',
        color: '#FFFFFF',
        margin: 'sm'
      }
    ],
    backgroundColor: '#1DB446',
    paddingAll: '15px'
  };
}

/**
 * テーブルヘッダー行を生成
 * @return {Object} テーブルヘッダーコンポーネント
 */
function buildTableHeader() {
  return {
    type: 'box',
    layout: 'horizontal',
    contents: [
      {
        type: 'text',
        text: '日付',
        size: 'xs',
        color: '#888888',
        weight: 'bold',
        flex: 2
      },
      {
        type: 'text',
        text: '時間帯',
        size: 'xs',
        color: '#888888',
        weight: 'bold',
        flex: 3
      },
      {
        type: 'text',
        text: '施設',
        size: 'xs',
        color: '#888888',
        weight: 'bold',
        flex: 2
      }
    ],
    margin: 'md'
  };
}

/**
 * 時間帯に応じた色を返す
 * @param {string} timeSlot - 時間帯文字列
 * @return {string} HEXカラーコード
 */
function getTimeSlotColor(timeSlot) {
  if (!timeSlot) return '#555555';
  if (timeSlot.match(/^6時/)) return '#E67E22';   // 早朝: オレンジ
  if (timeSlot.match(/^17時/)) return '#8E44AD';   // 夕方: パープル
  if (timeSlot.match(/^22時/)) return '#2C3E80';   // 夜間: ダークブルー
  return '#555555';
}

/**
 * 曜日に応じた行背景色を返す
 * @param {string} dayOfWeek - 曜日
 * @return {string|undefined} HEXカラーコード (平日はundefined)
 */
function getRowBackground(dayOfWeek) {
  if (dayOfWeek === '日') return '#FFF0F0';
  if (dayOfWeek === '土') return '#F0F0FF';
  return undefined;
}

/**
 * シフト行を生成
 * @param {Array<Object>} shifts - シフト配列
 * @return {Array<Object>} Flex Messageコンポーネント配列
 */
function buildShiftRows(shifts) {
  const rows = [];
  let prevDate = '';

  for (const shift of shifts) {
    // 日付から日部分を抽出して表示
    const dateParts = shift.date.split('/');
    const dayDisplay = parseInt(dateParts[2], 10) + '日';
    const dayOfWeek = getDayOfWeek(parseInt(dateParts[0]), parseInt(dateParts[1]), parseInt(dateParts[2]));

    // 同日の2行目以降は日付を省略
    const showDate = shift.date !== prevDate;
    prevDate = shift.date;

    const dateText = showDate ? dayDisplay + '(' + dayOfWeek + ')' : '';
    const dateColor = showDate ? (dayOfWeek === '日' ? '#FF0000' : dayOfWeek === '土' ? '#0000FF' : '#333333') : '#333333';
    const bgColor = getRowBackground(dayOfWeek);
    const timeSlotColor = getTimeSlotColor(shift.timeSlot);

    const row = {
      type: 'box',
      layout: 'horizontal',
      contents: [
        {
          type: 'text',
          text: dateText || ' ',
          size: 'xs',
          color: dateColor,
          flex: 2,
          weight: showDate ? 'bold' : 'regular'
        },
        {
          type: 'text',
          text: shift.timeSlot,
          size: 'xs',
          color: timeSlotColor,
          weight: 'bold',
          flex: 3
        },
        {
          type: 'text',
          text: shortenFacilityName(shift.facility),
          size: 'xs',
          color: '#555555',
          flex: 2
        }
      ],
      margin: showDate ? 'md' : 'sm'
    };

    // 土日は背景色を追加
    if (bgColor) {
      row.backgroundColor = bgColor;
      row.cornerRadius = '4px';
      row.paddingAll = '4px';
    }

    rows.push(row);
  }

  return rows;
}

/**
 * フッター部分を生成
 * @param {number} shiftCount - シフト件数
 * @param {string} [icalUrl] - iCalダウンロードURL (省略可)
 * @return {Object} フッターコンポーネント
 */
function buildFooter(shiftCount, icalUrl) {
  var contents = [
    { type: 'separator' },
    {
      type: 'text',
      text: '合計: ' + shiftCount + '件のシフト',
      size: 'sm',
      color: '#666666',
      margin: 'md',
      weight: 'bold'
    },
    {
      type: 'text',
      text: '変更がある場合は管理者にご連絡ください',
      size: 'xs',
      color: '#999999',
      margin: 'sm',
      wrap: true
    }
  ];

  if (icalUrl) {
    contents.push({
      type: 'button',
      action: {
        type: 'uri',
        label: 'カレンダーに追加',
        uri: icalUrl
      },
      style: 'primary',
      color: '#1DB446',
      margin: 'md',
      height: 'sm'
    });
  }

  return {
    type: 'box',
    layout: 'vertical',
    contents: contents,
    paddingAll: '15px'
  };
}

/**
 * 複数バブルに分割したカルーセルメッセージを生成
 * (シフトが多い場合、1バブルに収まらないため)
 * @param {string} displayMonth - 表示用年月
 * @param {string} employeeName - 従業員名
 * @param {Array<Object>} shifts - シフト配列
 * @param {string} [employeeNo] - 社員番号 (カレンダーURL用、省略可)
 * @return {Object} LINE Flex Message (carousel)
 */
function buildMultiBubbleMessage(displayMonth, employeeName, shifts, employeeNo) {
  const SHIFTS_PER_BUBBLE = 30;
  const bubbles = [];

  // iCal URL構築 (HMAC署名付き、最後のバブルのフッターに表示)
  var icalUrl = null;
  if (employeeNo) {
    // displayMonthから targetMonth を復元 (例: "2026年3月" → "2026-03")
    var monthMatch = displayMonth.match(/(\d{4})年(\d{1,2})月/);
    if (monthMatch) {
      var targetMonth = monthMatch[1] + '-' + String(monthMatch[2]).padStart(2, '0');
      icalUrl = buildICalUrl(employeeNo, targetMonth);
    }
  }

  for (let i = 0; i < shifts.length; i += SHIFTS_PER_BUBBLE) {
    const chunk = shifts.slice(i, i + SHIFTS_PER_BUBBLE);
    const pageNum = Math.floor(i / SHIFTS_PER_BUBBLE) + 1;
    const totalPages = Math.ceil(shifts.length / SHIFTS_PER_BUBBLE);

    const bodyContents = buildShiftRows(chunk);

    bubbles.push({
      type: 'bubble',
      size: 'mega',
      header: buildHeader(
        displayMonth + ' シフト予定 (' + pageNum + '/' + totalPages + ')',
        employeeName
      ),
      body: {
        type: 'box',
        layout: 'vertical',
        contents: [
          buildTableHeader(),
          { type: 'separator', margin: 'sm' },
          ...bodyContents
        ],
        paddingAll: '12px',
        spacing: 'none'
      },
      footer: pageNum === totalPages ? buildFooter(shifts.length, icalUrl) : undefined
    });
  }

  return {
    type: 'flex',
    altText: displayMonth + ' シフト予定 - ' + employeeName + 'さん',
    contents: {
      type: 'carousel',
      contents: bubbles.slice(0, 12) // LINE Flex Messageの最大12バブル
    }
  };
}

/**
 * 曜日を取得
 * @param {number} year - 年
 * @param {number} month - 月
 * @param {number} day - 日
 * @return {string} 曜日 (日,月,火,水,木,金,土)
 */
function getDayOfWeek(year, month, day) {
  const days = ['日', '月', '火', '水', '木', '金', '土'];
  const date = new Date(year, month - 1, day);
  return days[date.getDay()];
}

/**
 * 施設名を短縮表示
 * @param {string} name - 施設名
 * @return {string} 短縮施設名
 */
function shortenFacilityName(name) {
  if (!name) return '';
  // 長い施設名は10文字に制限
  if (name.length > 10) {
    return name.substring(0, 9) + '…';
  }
  return name;
}
