/**
 * 【システムの機能】
 *
 * ①練習日程のお知らせ（既存）
 * ②出欠登録のリマインド（既存）
 * ③練習日に参加する保護者のお知らせ（既存）
 * ④子どものチーム分け（既存）
 * ⑤試合記録の保存・集計（既存 saveMatch + 新規 saveExternalMatch）
 * ⑥成績取得（既存 + 種別フィルタ追加）
 */

/**************************************************************************************
 * ***  時間トリガー起動  ****************************************************************
 *  notifyMonthlyPracticeSchedule  毎月1日 8:00-9:00
 *  remindAllIfNeeded              毎週金曜 8:00-9:00
 *  runLotteryIfNeeded             毎週土曜 0:00-1:00
 *  notifyDutyParentsIfNeeded      毎週土曜 9:00-10:00
 */

function notifyDutyParentsIfNeeded() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) {
    Logger.log('🔒 ロック取得失敗');
    return;
  }

  Logger.log('🚀 notifyDutyParentsIfNeeded start');

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const historySheet = ss.getSheetByName('当番履歴');
    const values = historySheet.getDataRange().getValues();

    const now = new Date();
    Logger.log(`⏰ 現在時刻: ${now}`);

    const notifyMap = {};

    for (let i = 1; i < values.length; i++) {
      const practiceDate = values[i][0];
      const name         = values[i][1] + 'の保護者の方';
      const reason       = values[i][2];
      const notified     = values[i][4];

      if (!practiceDate || notified === true) continue;

      const key = Utilities.formatDate(
        new Date(practiceDate),
        Session.getScriptTimeZone(),
        'yyyy/MM/dd'
      );

      if (!notifyMap[key]) notifyMap[key] = [];
      notifyMap[key].push({ rowIndex: i + 1, name, reason });
    }

    Logger.log(`📦 通知対象まとめ: ${JSON.stringify(notifyMap)}`);

    for (const dateKey in notifyMap) {
      const targets = notifyMap[dateKey];

      let text =
        `【当番のお願い】\n` +
        `${Utilities.formatDate(new Date(dateKey), Session.getScriptTimeZone(), 'M/d')} の練習について\n\n`;

      targets.forEach(t => { text += `・${t.name}\n`; });

      text +=
        `\n※抽選または事前の出席回答により当番対象となっています。\n` +
        `よろしくお願いします！`;

      Logger.log(`📨 送信メッセージ:\n${text}`);

      pushMessage(text);

      targets.forEach(t => {
        historySheet.getRange(t.rowIndex, 5).setValue(true);
      });

      Logger.log(`✅ 通知済み更新完了: ${dateKey}`);
    }

  } catch (err) {
    Logger.log('❌ エラー発生');
    Logger.log(err.stack || err);
    throw err;
  } finally {
    lock.releaseLock();
    Logger.log('🔓 ロック解放');
    Logger.log('🏁 notifyDutyParentsIfNeeded end');
  }
}

function runLotteryIfNeeded() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) {
    Logger.log('他の処理が実行中のためスキップ');
    return;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const practiceSheet = ss.getSheetByName('練習日マスタ');
    const historySheet = ss.getSheetByName('当番履歴');

    const practices = practiceSheet.getDataRange().getValues();
    const histories = historySheet.getDataRange().getValues();

    const now = new Date();

    const alreadyDrawn = new Set();
    for (let i = 1; i < histories.length; i++) {
      alreadyDrawn.add(
        Utilities.formatDate(
          new Date(histories[i][0]),
          Session.getScriptTimeZone(),
          'yyyy/MM/dd'
        )
      );
    }

    for (let i = 1; i < practices.length; i++) {
      const practiceDate = practices[i][0];

      if (!practiceDate) continue;

      const schedule = calculateScheduleDates(practiceDate);
      const key = Utilities.formatDate(
        new Date(practiceDate),
        Session.getScriptTimeZone(),
        'yyyy/MM/dd'
      );

      if (!alreadyDrawn.has(key) && now >= schedule.draw) {
        Logger.log(`抽選を実行します：${key}`);
        decideDutyParentsFully(practiceDate);
      }
    }
  } finally {
    lock.releaseLock();
  }
}

function remindAllIfNeeded() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) {
    Logger.log('🔒 ロック取得失敗');
    return;
  }

  Logger.log('🚀 remindAllIfNeeded start');

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const practiceSheet = ss.getSheetByName('練習日マスタ');
    if (!practiceSheet) throw new Error('練習日マスタ が見つかりません');

    const HISTORY_SHEET_NAME = 'リマインド履歴';
    let historySheet = ss.getSheetByName(HISTORY_SHEET_NAME);
    if (!historySheet) {
      historySheet = ss.insertSheet(HISTORY_SHEET_NAME);
      historySheet.appendRow(['練習日', '送信日時', '種別', 'メモ']);
    }

    const practices = practiceSheet.getDataRange().getValues();
    const histories = historySheet.getDataRange().getValues();

    const now = new Date();
    Logger.log(`⏰ 現在時刻: ${now}`);

    const remindedSet = new Set();
    for (let i = 1; i < histories.length; i++) {
      const practiceDate = histories[i][0];
      const type = histories[i][2];
      if (!practiceDate || type !== 'REMIND') continue;

      const key = Utilities.formatDate(
        new Date(practiceDate),
        Session.getScriptTimeZone(),
        'yyyy/MM/dd'
      );
      remindedSet.add(key);
    }

    Logger.log(`📌 既送信(REMIND)件数: ${remindedSet.size}`);

    for (let i = 1; i < practices.length; i++) {
      const practiceDate = practices[i][0];

      if (!practiceDate) continue;

      const key = Utilities.formatDate(
        new Date(practiceDate),
        Session.getScriptTimeZone(),
        'yyyy/MM/dd'
      );

      const schedule = calculateScheduleDates(practiceDate);

      Logger.log('---');
      Logger.log(`📅 練習日: ${key}`);
      Logger.log(`🔔 リマインド時刻: ${schedule.remind}`);

      if (now < schedule.remind) {
        Logger.log('⏳ リマインド時刻未到達 → スキップ');
        continue;
      }

      if (remindedSet.has(key)) {
        Logger.log('✅ 既にリマインド済み → スキップ');
        continue;
      }

      const md = Utilities.formatDate(new Date(practiceDate), Session.getScriptTimeZone(), 'M/d');
      const text =
`【出欠入力リマインド】
${md} の練習について、出欠入力の締め切り日です。

まだ入力していない方は、本日中にフォーム入力をお願いします！`;

      Logger.log(`📨 送信メッセージ:\n${text}`);

      pushMessage(text);

      historySheet.appendRow([key, new Date(), 'REMIND', '']);

      Logger.log(`✅ リマインド送信＆履歴記録: ${key}`);
    }

  } catch (err) {
    Logger.log('❌ エラー発生');
    Logger.log(err.stack || err);
    throw err;
  } finally {
    lock.releaseLock();
    Logger.log('🔓 ロック解放');
    Logger.log('🏁 remindAllIfNeeded end');
  }
}

function notifyMonthlyPracticeSchedule() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const practiceSheet = ss.getSheetByName('練習日マスタ');
  const values = practiceSheet.getDataRange().getValues();

  const now = new Date();
  const tz = Session.getScriptTimeZone();

  const monthStart = new Date(now.getFullYear(), now.getMonth(), 1, 0, 0, 0, 0);

  const scheduleMap = {};

  for (let i = 1; i < values.length; i++) {
    const practiceDate = values[i][0];
    const timeText    = values[i][1];

    if (!practiceDate) continue;
    if (!(practiceDate instanceof Date)) continue;
    if (practiceDate < monthStart) continue;

    const key = Utilities.formatDate(practiceDate, tz, 'yyyy-MM');

    if (!scheduleMap[key]) scheduleMap[key] = [];
    scheduleMap[key].push({
      date: new Date(practiceDate),
      time: timeText || '時間未定'
    });
  }

  if (Object.keys(scheduleMap).length === 0) {
    Logger.log('📭 通知対象の練習日なし');
    return;
  }

  let message = '【今月以降の練習予定】\n\n';

  Object.keys(scheduleMap).sort().forEach(key => {
    const [, month] = key.split('-');
    message += `■ ${Number(month)}月\n`;

    scheduleMap[key].sort((a, b) => a.date - b.date).forEach(item => {
      const d = item.date;
      const dateStr = Utilities.formatDate(d, tz, 'M/d');
      const week = ['日','月','火','水','木','金','土'][d.getDay()];
      message += `・${dateStr}(${week}) ${item.time}\n`;
    });

    message += '\n';
  });

  message += '※変更があれば別途連絡します。';

  Logger.log(`📨 月初予定通知:\n${message}`);

  pushMessage(message);
}


/**************************************************************************************
 * ***  スプレッドシート編集トリガー起動  **************************************************
 */
function onEditHandler(e) {
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  if (sheetName === '名簿') {
    updateFormParentList();
  }

  if (sheetName === '練習日マスタ') {
    updateFormPracticeDateList();
    updateFormPracticeQuestionTitle();
  }
}

function updateFormParentList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName('名簿');
  const masters = masterSheet.getDataRange().getValues();

  const parentNames = [];
  for (let i = 1; i < masters.length; i++) {
    const name = masters[i][0];
    if (name) parentNames.push(name);
  }

  const form = FormApp.openById('1ZYxbcJGznnwsfo-ed4vpR7J6E1Vq9zO0BDBc2xO86EM');
  const items = form.getItems(FormApp.ItemType.LIST);
  const parentItem = items.find(item => item.getTitle() === '子どものお名前');

  if (!parentItem) throw new Error('「子どものお名前」の質問が見つかりません');

  parentItem.asListItem().setChoiceValues(parentNames);
}

function updateFormPracticeDateList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const practiceSheet = ss.getSheetByName('練習日マスタ');
  const values = practiceSheet.getDataRange().getValues();

  const dateList = [];
  const tz = Session.getScriptTimeZone();
  const week = ['日', '月', '火', '水', '木', '金', '土'];

  for (let i = 1; i < values.length; i++) {
    const practiceDate = values[i][0];
    if (practiceDate) {
      const date = new Date(practiceDate);
      const formatted = Utilities.formatDate(date, tz, 'yyyy/MM/dd');
      const dayOfWeek = week[date.getDay()];
      dateList.push(`${formatted}（${dayOfWeek}）`);
    }
  }

  const form = FormApp.openById('1ZYxbcJGznnwsfo-ed4vpR7J6E1Vq9zO0BDBc2xO86EM');
  const items = form.getItems(FormApp.ItemType.LIST);
  const dateItem = items.find(item => item.getTitle().includes('練習日'));

  if (!dateItem) throw new Error('フォームに「練習日」が見つかりません');

  dateItem.asListItem().setChoiceValues(dateList);
}

function updateFormPracticeQuestionTitle() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('練習日マスタ');
  const values = sheet.getDataRange().getValues();

  const tz = Session.getScriptTimeZone();
  const week = ['日','月','火','水','木','金','土'];

  const lines = [];

  for (let i = 1; i < values.length; i++) {
    const date = values[i][0];
    const time = values[i][1];
    if (!date) continue;

    const d = new Date(date);
    const dateStr = Utilities.formatDate(d, tz, 'M/d');
    const dayStr = week[d.getDay()];
    const timeStr = time ? ` ${time}` : '';

    lines.push(`・${dateStr}(${dayStr})${timeStr}`);
  }

  if (lines.length === 0) {
    Logger.log('練習日なし');
    return;
  }

  const title = `練習日\n${lines.join('\n')}`;

  const form = FormApp.openById('1ZYxbcJGznnwsfo-ed4vpR7J6E1Vq9zO0BDBc2xO86EM');
  const items = form.getItems(FormApp.ItemType.LIST);

  if (items.length === 0) throw new Error('プルダウン質問が見つかりません');

  const item = items[1];
  item.setTitle(title);

  Logger.log('質問文更新完了');
}


/**************************************************************************************
 * ***  部品  **************************************************************************
 */

function normalizePracticeDate(value) {
  if (!value) return null;

  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy/MM/dd');
  }

  const match = value.toString().match(/\d{4}\/\d{2}\/\d{2}/);
  return match ? match[0] : null;
}

function calculateScheduleDates(practiceDate) {
  const practice = new Date(practiceDate);

  const day = practice.getDay();
  const diffToMonday = (day === 0) ? -6 : 1 - day;
  const mondayOfPracticeWeek = new Date(practice);
  mondayOfPracticeWeek.setDate(practice.getDate() + diffToMonday);

  const mondayOneWeekBefore = new Date(mondayOfPracticeWeek);
  mondayOneWeekBefore.setDate(mondayOfPracticeWeek.getDate() - 7);

  const fridayDeadline = new Date(mondayOneWeekBefore);
  fridayDeadline.setDate(mondayOneWeekBefore.getDate() + 4);

  const deadline = new Date(fridayDeadline);
  deadline.setHours(23, 59, 0, 0);

  const remind = new Date(fridayDeadline);
  remind.setHours(9, 0, 0, 0);

  const draw = new Date(deadline);

  const notify = new Date(fridayDeadline);
  notify.setDate(fridayDeadline.getDate() + 1);
  notify.setHours(9, 0, 0, 0);

  return { practiceDate: practice, deadline, remind, draw, notify };
}

function decideDutyParentsFully(practiceDate) {
  const REQUIRED_PARENT_COUNT = 4;
  const REASON_DRAW   = '親人数不足のため当番抽選';
  const REASON_ATTEND = 'フォームで親出席と回答';

  const targetDate = Utilities.formatDate(
    new Date(practiceDate), Session.getScriptTimeZone(), 'yyyy/MM/dd'
  );

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const answerSheet  = ss.getSheetByName('フォームの回答 1');
  const masterSheet  = ss.getSheetByName('名簿');
  const historySheet = ss.getSheetByName('当番履歴');

  const answers = answerSheet.getDataRange().getValues();
  const masters = masterSheet.getDataRange().getValues();

  const dutyCountMap = {};
  const activeParentSet = new Set();

  for (let i = 1; i < masters.length; i++) {
    const name  = masters[i][0];
    const count = Number(masters[i][2] || 0);

    if (name) {
      dutyCountMap[name] = count;
      activeParentSet.add(name);
    }
  }

  const attendParents = new Set();

  for (let i = 1; i < answers.length; i++) {
    const row = answers[i];
    const rowDate = normalizePracticeDate(row[2]);

    if (rowDate === targetDate && row[4] === '出席') {
      attendParents.add(row[1]);
    }
  }

  const currentParentCount = attendParents.size;
  const lackCount = REQUIRED_PARENT_COUNT - currentParentCount;

  Logger.log(`練習日: ${targetDate}`);
  Logger.log(`親出席人数: ${currentParentCount}`);
  Logger.log(`不足人数: ${lackCount}`);

  const now = new Date();
  const historyRows = [];
  const incrementSet = new Set();

  attendParents.forEach(name => {
    historyRows.push([targetDate, name, REASON_ATTEND, now, false]);
    incrementSet.add(name);
  });

  if (lackCount <= 0) {
    writeHistoryRowsAndIncrement(historyRows, incrementSet);
    Logger.log('抽選不要（親人数充足）');
    return;
  }

  const candidates = [];

  for (let i = 1; i < answers.length; i++) {
    const row = answers[i];
    const rowDate = normalizePracticeDate(row[2]);
    const name = row[1];

    if (rowDate === targetDate && row[3] === '出席' && row[4] === '欠席' && activeParentSet.has(name)) {
      candidates.push({ name, dutyCount: dutyCountMap[name] ?? 0 });
    }
  }

  if (candidates.length === 0) {
    Logger.log('当番候補者がいません');
    writeHistoryRowsAndIncrement(historyRows, incrementSet);
    return;
  }

  candidates.sort((a, b) => a.dutyCount - b.dutyCount);

  const selectedParents = [];
  let index = 0;

  while (selectedParents.length < lackCount && index < candidates.length) {
    const currentCount = candidates[index].dutyCount;

    const sameGroup = candidates
      .filter(c => c.dutyCount === currentCount)
      .sort(() => Math.random() - 0.5);

    for (const p of sameGroup) {
      if (selectedParents.length < lackCount) {
        selectedParents.push(p.name);
      }
    }

    index += sameGroup.length;
  }

  Logger.log('今回の当番者（抽選）');
  selectedParents.forEach(name => Logger.log(name));

  selectedParents.forEach(name => {
    historyRows.push([targetDate, name, REASON_DRAW, now, false]);
    incrementSet.add(name);
  });

  writeHistoryRowsAndIncrement(historyRows, incrementSet);
}

function writeHistoryRowsAndIncrement(historyRows, incrementSet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const historySheet = ss.getSheetByName('当番履歴');
  const masterSheet  = ss.getSheetByName('名簿');

  if (historyRows.length > 0) {
    historySheet.getRange(
      historySheet.getLastRow() + 1, 1,
      historyRows.length, historyRows[0].length
    ).setValues(historyRows);
  }

  const masters = masterSheet.getDataRange().getValues();

  for (let i = 1; i < masters.length; i++) {
    const name = masters[i][0];
    if (incrementSet.has(name)) {
      masters[i][2] = Number(masters[i][2] || 0) + 1;
    }
  }

  masterSheet.getDataRange().setValues(masters);
}


/**************************************************************************************
 * ***  LINEメッセージ受信トリガー  *******************************************************
 */

function doPost(e) {
  const SPREADSHEET_ID = '1I6hFACp7uovYjONS0SKIOY2nJ8RHQUELEiA81cuwalI';
  const SHEET_NAME = 'LINE_WEBHOOK_LOG';

  try {
    const json = JSON.parse(e.postData.contents);
    const text = json.events?.[0]?.message?.text || '';

    if (!json.events || json.events.length === 0) {
      return ContentService.createTextOutput('OK');
    }

    const event = json.events[0];
    const source = event.source || {};

    const groupId = source.groupId || '';
    const roomId  = source.roomId  || '';
    const userId  = source.userId  || '';
    const type    = source.type    || '';

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(['timestamp','type','groupId','roomId','userId']);
    }

    sheet.appendRow([new Date(), type, groupId, roomId, userId, text]);

    return ContentService.createTextOutput('OK');

  } catch (err) {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

    sheet.appendRow([new Date(), 'ERROR', err.message, '', '']);

    return ContentService.createTextOutput('OK');
  }
}

function getNextPracticeParticipants() {
  const ss = SpreadsheetApp.openById('1I6hFACp7uovYjONS0SKIOY2nJ8RHQUELEiA81cuwalI');
  const sheet = ss.getSheetByName('フォームの回答 1');
  const values = sheet.getDataRange().getValues();

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const map = {};

  for (let i = 1; i < values.length; i++) {
    const childName = values[i][1];
    const practiceDateRaw = values[i][2];
    const attend = values[i][3];

    const practiceDate = parsePracticeDate(practiceDateRaw);
    if (!practiceDate) {
      Logger.log(`⚠️ 練習日パース失敗: raw=${practiceDateRaw}`);
      continue;
    }

    practiceDate.setHours(0, 0, 0, 0);

    if (!childName || attend !== '出席') continue;
    if (practiceDate < today) continue;

    const key = Utilities.formatDate(practiceDate, Session.getScriptTimeZone(), 'yyyy/MM/dd');

    if (!map[key]) map[key] = new Set();
    map[key].add(childName);
  }

  const sortedDates = Object.keys(map).sort();
  if (sortedDates.length === 0) return { date: null, members: [] };

  const targetDate = sortedDates[0];
  return { date: targetDate, members: Array.from(map[targetDate]) };
}

function parsePracticeDate(value) {
  if (!value) return null;

  if (value instanceof Date && !isNaN(value.getTime())) return value;

  const m = value.toString().match(/\d{4}\/\d{1,2}\/\d{1,2}/);
  if (!m) return null;

  const [y, mo, d] = m[0].split('/').map(n => parseInt(n, 10));
  return new Date(y, mo - 1, d);
}

function splitTeamsEvenly(members) {
  const shuffled = [...members].sort(() => Math.random() - 0.5);
  const mid = Math.ceil(shuffled.length / 2);
  return { teamA: shuffled.slice(0, mid), teamB: shuffled.slice(mid) };
}

function runTeamRoulette() {
  const { date, members } = getNextPracticeParticipants();

  if (!date || members.length === 0) return '参加者がいません';

  const { teamA, teamB } = splitTeamsEvenly(members);

  let message =
`🎯 チーム分けルーレット結果
対象日：${date}

【チームA】
${teamA.join('\n')}

【チームB】
${teamB.join('\n')}
`;

  return message;
}

function handleLineMessage() {
  const message = runTeamRoulette();
  sendLineMessageToGroup(message);
}

function pushMessage(text) {
  const url = 'https://api.line.me/v2/bot/message/push';

  const payload = {
    to: GROUP_ID,
    messages: [{ type: 'text', text: text }]
  };

  UrlFetchApp.fetch(url, {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN
    },
    payload: JSON.stringify(payload),
  });
}

function getAllMembers() {
  const ss = SpreadsheetApp.openById('1I6hFACp7uovYjONS0SKIOY2nJ8RHQUELEiA81cuwalI');
  const sheet = ss.getSheetByName('名簿');
  const values = sheet.getDataRange().getValues();
  const members = [];
  for (let i = 1; i < values.length; i++) {
    if (values[i][0]) members.push(values[i][0]);
  }
  return members;
}


/**************************************************************************************
 * ***  Web API (doGet)  ***
 *
 * 既存 + 種別フィルタ対応 + 他校試合保存
 *
 * パラメータ:
 *   action: 上記のいずれか
 *   type:   'all' | 'practice' | 'external'  （取得系で使用、デフォルト 'all'）
 *   data:   JSON文字列（saveMatch / saveExternalMatch で使用）
 */

function doGet(e) {
  const action = e.parameter.action;
  const type   = e.parameter.type || 'all';
  let result;

  // 子ども名簿全員を返す
  if (action === 'getAllMembers') {
    const members = getAllMembers();
    result = JSON.stringify({ members });

  // フォーム回答から次の練習の出席者を返す
  } else if (action === 'getMembers') {
    const { members } = getNextPracticeParticipants();
    result = JSON.stringify({ members });

  // 保護者名簿を返す
  } else if (action === 'getParents') {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName('保護者名簿');
    const values = sheet.getDataRange().getValues();
    const parents = [];
    for (let i = 1; i < values.length; i++) {
      if (!values[i][1]) continue;
      parents.push({ childName: values[i][0], parentName: values[i][1] });
    }
    result = JSON.stringify({ parents });

  // 練習試合の保存
  } else if (action === 'saveMatch') {
    const data = JSON.parse(e.parameter.data);
    saveMatchResult(data);
    result = JSON.stringify({ success: true });

  // 他校試合の保存
  } else if (action === 'saveExternalMatch') {
    const data = JSON.parse(e.parameter.data);
    saveExternalMatchResult(data);
    result = JSON.stringify({ success: true });

  // 子どもの個人成績（種別フィルタ対応）
  } else if (action === 'getPersonalStats') {
    const stats = computeStatsByType('退場記録', type);
    result = JSON.stringify({ stats });

  // 子どもの退場記録（種別フィルタ対応）
  } else if (action === 'getMatchDetail') {
    const records = readDetailByType('退場記録', type);
    result = JSON.stringify({ records });

  // 保護者の個人成績（種別フィルタ対応）
  } else if (action === 'getParentStats') {
    const stats = computeStatsByType('保護者退場記録', type);
    result = JSON.stringify({ stats });

  // 保護者の退場記録（種別フィルタ対応）
  } else if (action === 'getParentDetail') {
    const records = readDetailByType('保護者退場記録', type);
    result = JSON.stringify({ records });

  } else {
    result = JSON.stringify({ error: 'unknown action' });
  }

  return ContentService
    .createTextOutput(result)
    .setMimeType(ContentService.MimeType.JSON);
}


/**************************************************************************************
 * ***  試合保存（練習試合）  ***
 *
 * data: { matchNum, teamA, teamB, outA, outB, winner, hits, parents }
 *
 * 退場記録の H列に type='practice' を記録（後で他校試合と区別するため）
 */
function saveMatchResult(data) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const now = new Date();
  const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/MM/dd');

  const parentNames = new Set(data.parents || []);

  // 試合履歴
  let matchSheet = ss.getSheetByName('試合履歴');
  if (!matchSheet) {
    matchSheet = ss.insertSheet('試合履歴');
    matchSheet.appendRow(['日付','試合No','チームA','チームB','退場A','退場B','勝者','種別','対戦相手','詳細']);
  }
  matchSheet.appendRow([
    dateStr, data.matchNum,
    data.teamA.join(','), data.teamB.join(','),
    data.outA.length, data.outB.length, data.winner,
    'practice', '', ''
  ]);

  const resultA = data.winner === 'A' ? '勝ち' : data.winner === 'B' ? '負け' : '引き分け';
  const resultB = data.winner === 'B' ? '勝ち' : data.winner === 'A' ? '負け' : '引き分け';
  const hits = data.hits || {};

  // 子どもの退場記録
  let detailSheet = ss.getSheetByName('退場記録');
  if (!detailSheet) {
    detailSheet = ss.insertSheet('退場記録');
    detailSheet.appendRow(['日付','試合No','名前','チーム','被ヒット','勝敗','与ヒット数','種別']);
  }
  const childRows = [];
  data.teamA.filter(n => !parentNames.has(n)).forEach(name => {
    childRows.push([dateStr, data.matchNum, name, 'A', data.outA.includes(name), resultA, hits[name] || 0, 'practice']);
  });
  data.teamB.filter(n => !parentNames.has(n)).forEach(name => {
    childRows.push([dateStr, data.matchNum, name, 'B', data.outB.includes(name), resultB, hits[name] || 0, 'practice']);
  });
  if (childRows.length > 0) {
    detailSheet.getRange(detailSheet.getLastRow() + 1, 1, childRows.length, 8).setValues(childRows);
  }

  // 保護者の退場記録
  let parentSheet = ss.getSheetByName('保護者退場記録');
  if (!parentSheet) {
    parentSheet = ss.insertSheet('保護者退場記録');
    parentSheet.appendRow(['日付','試合No','名前','チーム','被ヒット','勝敗','与ヒット数','種別']);
  }
  const parentRows = [];
  data.teamA.filter(n => parentNames.has(n)).forEach(name => {
    parentRows.push([dateStr, data.matchNum, name, 'A', data.outA.includes(name), resultA, hits[name] || 0, 'practice']);
  });
  data.teamB.filter(n => parentNames.has(n)).forEach(name => {
    parentRows.push([dateStr, data.matchNum, name, 'B', data.outB.includes(name), resultB, hits[name] || 0, 'practice']);
  });
  if (parentRows.length > 0) {
    parentSheet.getRange(parentSheet.getLastRow() + 1, 1, parentRows.length, 8).setValues(parentRows);
  }

  // 個人成績シート（キャッシュ）も更新（全件で再集計）
  updatePersonalStats(ss);
  updateParentStats(ss);
}


/**************************************************************************************
 * ***  試合保存（他校試合）  ***
 *
 * data: {
 *   type: 'external',
 *   opponent, date,
 *   teamMembers: [...自チーム子ども],
 *   parents:     [...保護者],
 *   starters:    { in: [...], out: [...] },
 *   halves: [
 *     { num: 1, ourBest, theirBest, hits: {name: count}, ourZones: {name: 'in'|'out'|'eliminated'|null} },
 *     { num: 2, ... }
 *   ],
 *   totalOurBest, totalTheirBest,
 *   winner: 'us' | 'them' | 'draw',
 *   substitutions: [...]
 * }
 *
 * 既存スキーマと統一するために：
 *   - 退場記録に1人1行（試合単位）で書き込む
 *   - team列は 'us' で記録（既存の 'A'/'B' と区別できる）
 *   - matchNum列は 'vs ◯◯小' のラベル
 *   - 種別列は 'external'
 *   - 被ヒット: 後半終了時に 'eliminated' であれば true
 *   - 与ヒット数: 前半+後半 の合計
 */
function saveExternalMatchResult(data) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const dateStr = data.date
    ? Utilities.formatDate(new Date(data.date), tz, 'yyyy/MM/dd')
    : Utilities.formatDate(now, tz, 'yyyy/MM/dd');
  const opponent = data.opponent || '相手校';
  const matchLabel = 'vs ' + opponent;

  // 勝者を A/B/draw に変換（互換性のため）
  const winnerCode =
    data.winner === 'us'   ? 'A' :
    data.winner === 'them' ? 'B' :
    'draw';
  const result =
    data.winner === 'us'   ? '勝ち' :
    data.winner === 'them' ? '負け' :
    '引き分け';

  const halves = data.halves || [];
  const lastHalf = halves[halves.length - 1] || {};
  const finalZones = lastHalf.ourZones || {};

  // 全半合計のヒット数
  const totalHits = {};
  halves.forEach(h => {
    const hits = h.hits || {};
    Object.keys(hits).forEach(name => {
      totalHits[name] = (totalHits[name] || 0) + Number(hits[name] || 0);
    });
  });

  const parentSet = new Set(data.parents || []);

  // 退場（eliminated）になった子の数（後半終了時）
  const eliminatedAll = Object.keys(finalZones).filter(n => finalZones[n] === 'eliminated');
  const eliminatedChildren = eliminatedAll.filter(n => !parentSet.has(n));

  /* =========================
     ① 試合履歴
     ========================= */
  let matchSheet = ss.getSheetByName('試合履歴');
  if (!matchSheet) {
    matchSheet = ss.insertSheet('試合履歴');
    matchSheet.appendRow(['日付','試合No','チームA','チームB','退場A','退場B','勝者','種別','対戦相手','詳細']);
  }
  matchSheet.appendRow([
    dateStr,
    matchLabel,
    (data.teamMembers || []).concat(Array.from(parentSet)).join(','),
    opponent,
    eliminatedAll.length,
    /* 相手の退場数は累計で把握できないので空 */ '',
    winnerCode,
    'external',
    opponent,
    JSON.stringify({
      halves: halves.map(h => ({ num: h.num, ourBest: h.ourBest, theirBest: h.theirBest })),
      totalOurBest: data.totalOurBest,
      totalTheirBest: data.totalTheirBest,
      substitutions: data.substitutions || []
    })
  ]);

  /* =========================
     ② 子ども退場記録
     ========================= */
  let detailSheet = ss.getSheetByName('退場記録');
  if (!detailSheet) {
    detailSheet = ss.insertSheet('退場記録');
    detailSheet.appendRow(['日付','試合No','名前','チーム','被ヒット','勝敗','与ヒット数','種別']);
  }
  const childRows = [];
  (data.teamMembers || []).filter(n => !parentSet.has(n)).forEach(name => {
    const isOut = finalZones[name] === 'eliminated';
    childRows.push([dateStr, matchLabel, name, 'us', isOut, result, totalHits[name] || 0, 'external']);
  });
  if (childRows.length > 0) {
    detailSheet.getRange(detailSheet.getLastRow() + 1, 1, childRows.length, 8).setValues(childRows);
  }

  /* =========================
     ③ 保護者退場記録
     ========================= */
  let parentSheet = ss.getSheetByName('保護者退場記録');
  if (!parentSheet) {
    parentSheet = ss.insertSheet('保護者退場記録');
    parentSheet.appendRow(['日付','試合No','名前','チーム','被ヒット','勝敗','与ヒット数','種別']);
  }
  const parentRows = [];
  Array.from(parentSet).forEach(name => {
    const isOut = finalZones[name] === 'eliminated';
    parentRows.push([dateStr, matchLabel, name, 'us', isOut, result, totalHits[name] || 0, 'external']);
  });
  if (parentRows.length > 0) {
    parentSheet.getRange(parentSheet.getLastRow() + 1, 1, parentRows.length, 8).setValues(parentRows);
  }

  // 個人成績キャッシュも更新（全件で再集計）
  updatePersonalStats(ss);
  updateParentStats(ss);
}


/**************************************************************************************
 * ***  集計（種別フィルタ）  ***
 *
 * 退場記録 / 保護者退場記録 を読み、type で絞り込んで個人成績を計算する
 */
function computeStatsByType(sheetName, type) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  const values = sheet.getDataRange().getValues();
  const statsMap = {};

  for (let i = 1; i < values.length; i++) {
    const name = values[i][2];
    if (!name) continue;

    const rowType = values[i][7] || 'practice'; // 過去データは practice 扱い
    if (type && type !== 'all' && rowType !== type) continue;

    const isOut    = values[i][4];
    const matchRes = values[i][5];
    const hitCount = Number(values[i][6] || 0);

    if (!statsMap[name]) statsMap[name] = { total: 0, win: 0, survive: 0, hits: 0 };
    statsMap[name].total++;
    if (matchRes === '勝ち') statsMap[name].win++;
    if (!isOut) statsMap[name].survive++;
    statsMap[name].hits += hitCount;
  }

  return Object.keys(statsMap).map(name => {
    const s = statsMap[name];
    return {
      name,
      total:       s.total,
      win:         s.win,
      winRate:     s.total > 0 ? Math.round(s.win / s.total * 100) : 0,
      survive:     s.survive,
      surviveRate: s.total > 0 ? Math.round(s.survive / s.total * 100) : 0,
      hits:        s.hits
    };
  });
}

/**
 * 退場記録 / 保護者退場記録 の生レコードを type でフィルタして返す
 * （相性分析などで使用）
 */
function readDetailByType(sheetName, type) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  const values = sheet.getDataRange().getValues();
  const records = [];

  for (let i = 1; i < values.length; i++) {
    if (!values[i][0]) continue;

    const rowType = values[i][7] || 'practice';
    if (type && type !== 'all' && rowType !== type) continue;

    records.push({
      date:     Utilities.formatDate(new Date(values[i][0]), Session.getScriptTimeZone(), 'yyyy/MM/dd'),
      matchNum: values[i][1],
      name:     values[i][2],
      team:     values[i][3],
      isOut:    values[i][4],
      result:   values[i][5]
    });
  }
  return records;
}


/**
 * 子どもの個人成績シート（キャッシュ）を再集計
 * 全件（type=all 相当）で書き出し
 */
function updatePersonalStats(ss) {
  const stats = computeStatsByType('退場記録', 'all');
  let statsSheet = ss.getSheetByName('個人成績');
  if (!statsSheet) statsSheet = ss.insertSheet('個人成績');
  statsSheet.clearContents();
  statsSheet.appendRow(['名前','試合数','勝利数','勝率','生存数','生存率','与ヒット数']);
  stats.forEach(s => {
    statsSheet.appendRow([s.name, s.total, s.win, s.winRate, s.survive, s.surviveRate, s.hits]);
  });
}

/**
 * 保護者の個人成績シート（キャッシュ）を再集計
 */
function updateParentStats(ss) {
  const stats = computeStatsByType('保護者退場記録', 'all');
  let statsSheet = ss.getSheetByName('保護者成績');
  if (!statsSheet) statsSheet = ss.insertSheet('保護者成績');
  statsSheet.clearContents();
  statsSheet.appendRow(['名前','試合数','勝利数','勝率','生存数','生存率','与ヒット数']);
  stats.forEach(s => {
    statsSheet.appendRow([s.name, s.total, s.win, s.winRate, s.survive, s.surviveRate, s.hits]);
  });
}


/**************************************************************************************
 * ***  テスト用  ***
 */
function testPushMessage() {
  pushMessage('テスト通知です。GASから送っています。');
}

function testDecideDutyParentsFully() {
  decideDutyParentsFully(new Date('2026/1/3'));
}

function testCalculateScheduleDates() {
  ['2025/12/27','2025/12/08','2025/12/14'].forEach(d => {
    const r = calculateScheduleDates(new Date(d));
    Logger.log(`練習日: ${d}`);
    Logger.log(`締切: ${r.deadline}`);
    Logger.log('---');
  });
}

function testSaveExternalMatch() {
  const data = {
    type: 'external',
    opponent: 'テスト小学校',
    date: '2026/06/07',
    teamMembers: ['山田 太郎','鈴木 次郎','佐藤 三郎','田中 四郎','高橋 五郎','伊藤 花子','渡辺 陽子','中村 美智子'],
    parents: [],
    starters: { in: ['山田 太郎','鈴木 次郎','佐藤 三郎','田中 四郎','高橋 五郎'], out: ['伊藤 花子','渡辺 陽子','中村 美智子'] },
    halves: [
      { num: 1, ourBest: 6, theirBest: 4, hits: { '山田 太郎': 2, '伊藤 花子': 1 }, ourZones: { '山田 太郎':'in','鈴木 次郎':'in','佐藤 三郎':'out','田中 四郎':'eliminated','高橋 五郎':'eliminated','伊藤 花子':'out','渡辺 陽子':'out','中村 美智子':'in' } },
      { num: 2, ourBest: 5, theirBest: 3, hits: { '鈴木 次郎': 1 }, ourZones: { '山田 太郎':'in','鈴木 次郎':'in','佐藤 三郎':'eliminated','田中 四郎':'eliminated','高橋 五郎':'eliminated','伊藤 花子':'out','渡辺 陽子':'in','中村 美智子':'out' } }
    ],
    totalOurBest: 11,
    totalTheirBest: 7,
    winner: 'us',
    substitutions: []
  };
  saveExternalMatchResult(data);
  Logger.log('✅ テスト保存完了');
}
