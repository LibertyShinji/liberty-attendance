// ============================================================
// 出勤・有給管理システム（LINE LIFF + GAS + スプレッドシート）
// Google Apps Script (GAS)
// ============================================================

// スクリプトプロパティで管理（設定方法はsetup-guide.txtを参照）
function getConfig() {
  const props = PropertiesService.getScriptProperties();
  return {
    SPREADSHEET_ID:   props.getProperty('SPREADSHEET_ID'),
    LINE_TOKEN:       props.getProperty('LINE_CHANNEL_ACCESS_TOKEN'),
    LIFF_LEAVE_ID:       props.getProperty('LIFF_LEAVE_ID'),
    LIFF_ADMIN_ID:       props.getProperty('LIFF_ADMIN_ID'),
    LIFF_CALENDAR_ID:    props.getProperty('LIFF_CALENDAR_ID'),
    LIFF_ATTENDANCE_ID:  props.getProperty('LIFF_ATTENDANCE_ID'),
    LIFF_KIBOU_ID:       props.getProperty('LIFF_KIBOU_ID'),
    LIFF_TEAM_ID:        props.getProperty('LIFF_TEAM_ID'),
    LIFF_SHIFT_ID:       props.getProperty('LIFF_SHIFT_ID'),
    REPORT_EMAIL:        props.getProperty('REPORT_EMAIL'),  // 月次集計CSVの送信先（未設定ならGAS所有者）
  };
}

const SHEETS = {
  EMPLOYEES:     '社員マスタ',
  ATTENDANCE:    '打刻記録',
  LEAVE:         '有給申請',
  BALANCE:       '有給残日数',
  KOUKYUU:       '公休設定',
  SHIFT:         'シフト設定',
  COREMINDER_LOG: '退勤通知ログ',
  CIREMINDER_LOG: '出勤通知ログ',
  SUDDEN_ABSENCE: '急な休み申請',
  LEAVE_LOG:      '有給変更履歴',
};

// 退勤忘れ通知の設定（フェーズC）
const COREMINDER = {
  START_MIN: 19 * 60,   // 19:00 開始
  END_MIN:   22 * 60,   // 22:00 で停止（この時刻以降は送らない）
  GAP_MS:    28 * 60 * 1000,  // 前回通知から28分以上空けて再送（≒30分毎・15分ポーリングのゆらぎ吸収）
};

// 出勤忘れ通知の設定：シフト制のため固定時間帯ではなく「各自のシフト開始時刻」基準（2026-08〜）
// シフト開始時刻から WINDOW_MIN 分の間、未打刻なら対象（15分毎トリガーでポーリング）
const CIREMINDER = {
  WINDOW_MIN: 60,             // シフト開始から60分間は出勤忘れ通知の対象
  GAP_MS:    13 * 60 * 1000,  // 前回通知から13分以上空けて再送（15分ポーリングのゆらぎ吸収）
};

// "HH:mm" を当日0時からの分数に変換。不正な形式はnull。
function hhmmToMinutes_(s) {
  if (typeof s !== 'string') return null;
  const m = /^([01]\d|2[0-3]):([0-5]\d)$/.exec(s);
  if (!m) return null;
  return parseInt(m[1], 10) * 60 + parseInt(m[2], 10);
}

// ============================================================
// エントリーポイント
// ============================================================

function doGet(e) {
  // TimeTree等の外部カレンダーアプリ向け：確定シフトのiCal配信（閲覧専用・鍵付きURL）
  if (e && e.parameter && e.parameter.feed === 'shifts') {
    return handleShiftsFeed_(e);
  }
  const page = (e && e.parameter && e.parameter.page) || 'leave';
  const pageMap = {
    admin:      ['AdminDashboard', '管理者ダッシュボード'],
    calendar:   ['Calendar',       '出勤カレンダー'],
    leave:      ['LeaveRequest',   '有給申請'],
    attendance: ['Attendance',     '出勤・退勤打刻'],
    kibou:      ['KibouRequest',   '休み希望申請'],
    team:       ['TeamCalendar',   '全員カレンダー'],
    shift:      ['ShiftRequest',   'シフト希望申告'],
  };
  const [file, title] = pageMap[page] || pageMap.leave;
  return HtmlService.createHtmlOutputFromFile(file)
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);

    // LIFFからのAPI呼び出し（typeフィールドで判別）
    if (body.type) {
      const result = handleLiffApi(body);
      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // LINEからのWebhook
    if (body.events) {
      body.events.forEach(handleEvent);
    }
  } catch (err) {
    Logger.log('doPost error: ' + err.toString());
  }
  return ContentService
    .createTextOutput(JSON.stringify({ success: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// LIFF API ルーター
// ============================================================

function handleLiffApi(body) {
  switch (body.type) {
    // 有給
    case 'submit_leave':         return submitLeaveRequest(body.data);
    case 'get_my_balance':       return getMyBalance(body.userId);
    case 'get_my_history':       return getMyAttendanceHistory(body.userId, body.year, body.month);
    // 管理者ダッシュボード
    case 'get_dashboard':        return getAdminDashboardData(body.userId);
    case 'approve_leave':        return processLeaveApproval(body.applicationId, body.approverId, '承認');
    case 'reject_leave':         return processLeaveApproval(body.applicationId, body.approverId, '否認');
    // 打刻（LIFF）
    case 'liff_clock_in':        return liffClockIn(body.userId, body.loc);
    case 'liff_clock_out':       return liffClockOut(body.userId, body.loc);
    case 'get_today_status':     return getTodayStatus(body.userId);
    // 急な休み（本人が打刻画面から本日を休みにする。理由を送ると出勤予定スタッフ＋役員へLINE通知）
    case 'submit_sudden_absence': return submitSuddenAbsence(body.userId, body.reason);
    // カレンダー
    case 'get_calendar':         return getCalendarData(body.userId, body.year, body.month);
    // 公休（シフト）管理：管理者画面で入力
    case 'get_koukyuu_admin':    return getKoukyuuForAdmin(body.userId, body.year, body.month);
    case 'save_koukyuu':         return saveKoukyuu(body.userId, body.employeeId, body.dates, body.year, body.month);
    // 休み希望（社員→申請、管理者→依頼/確定）
    case 'get_my_kibou':         return getMyKibou(body.userId, body.year, body.month);
    case 'submit_kibou':         return submitKibou(body.userId, body.dates, body.year, body.month, body.cancels);
    case 'request_kibou_all':    return requestKibouFromAll(body.userId, body.year, body.month);
    // シフト（出勤予定）：社員→希望申告、管理者→確定（2026-08 追加）
    case 'get_my_shift':         return getMyShift(body.userId, body.year, body.month);
    case 'submit_shift':         return submitShiftKibou(body.userId, body.entries, body.year, body.month);
    case 'get_shift_admin':      return getShiftForAdmin(body.userId, body.year, body.month);
    case 'save_shift':           return saveShift(body.userId, body.employeeId, body.entries, body.year, body.month);
    case 'request_shift_all':    return requestShiftFromAll(body.userId, body.year, body.month);
    // 社員設定（雇用形態・時給）・給与集計
    case 'get_employees_admin':  return getEmployeesAdmin(body.userId);
    case 'save_employee_settings': return saveEmployeeSettings(body.userId, body.employeeId, body.empType, body.wageWeekday, body.wageSat, body.wageSunHol, body.deductRate, body.fixedDeduct);
    case 'get_employees_full':   return getEmployeesFull(body.userId);
    case 'add_employee':         return addEmployeeAdmin(body.userId, body.name, body.empType, body.hireDate, body.annualDays);
    case 'retire_employee':      return setEmployeeActive(body.userId, body.employeeId, false);
    case 'reactivate_employee':  return setEmployeeActive(body.userId, body.employeeId, true);
    case 'get_unbound_employees': return getUnboundEmployees();
    case 'bind_my_line_id':      return bindMyLineId(body.userId, body.employeeId);
    // 管理者：出勤・退勤リマインド通知の当日ログを取得（読み取り専用）
    case 'get_reminder_logs':    return getReminderLogs(body.userId);
    case 'get_payroll':          return getPayrollSummary(body.userId, body.year, body.month);
    // 本人：自分の今月の給与（現在＋見込。パートのみ）
    case 'get_my_payroll':       return getMyPayroll(body.userId, body.year, body.month);
    // 本人：自分の月額目標を設定（パートのみ）
    case 'set_my_target':        return setMyTarget(body.userId, body.amount);
    // 管理者：月次サマリーをCSVでメール送信
    case 'email_monthly_summary': return emailMonthlySummary(body.userId, body.year, body.month);
    // 管理者：管理者全員にプッシュ通知が届くか診断（友だち状態・UserID有効性）
    case 'check_admins_reachable': return checkAdminsReachable(body.userId);
    // 管理者：6ボタンのリッチメニューを再設定（画像はGitHub Pagesから取得）
    case 'setup_richmenu7':       return setupRichMenu7(body.userId);
    // 管理者：有給残の手動訂正（used を指定値に。remaining は total-used で再計算）
    case 'adjust_leave_used':     return adjustLeaveUsed(body.userId, body.employeeId, body.year, body.used);
    // 管理者：有給申請の取消（誤申請の削除。承認済みなら残日数を戻す）
    case 'delete_leave':          return deleteLeaveRequest(body.userId, body.applicationId);
    // 管理者：有給変更履歴（監査ログ）を取得（読み取り専用。employeeId省略で全社員分）
    case 'get_leave_log':         return getLeaveLog(body.userId, body.employeeId);
    case 'get_employee_leave':    return getEmployeeLeaveDates(body.userId, body.employeeId);
    // 管理者：勤怠の代理打刻・修正（スマホ故障・打刻忘れ・ミス対応）
    case 'get_attendance_admin':  return getAttendanceForAdmin(body.userId, body.employeeId, body.year, body.month);
    case 'get_month_overview':    return getMonthOverview(body.userId, body.year, body.month);
    case 'save_attendance_admin': return saveAttendanceAdmin(body.userId, body.employeeId, body.date, body.clockIn, body.clockOut);
    // 全員カレンダー（誰がいつ出勤か）
    case 'get_team_calendar':    return getTeamCalendar(body.userId, body.year, body.month);
    default:                     return { success: false, message: '不明なAPIタイプです' };
  }
}

// ============================================================
// LINE Webhook イベントハンドラ
// ============================================================

function handleEvent(event) {
  if (event.type === 'message' && event.message.type === 'text') {
    handleTextMessage(event);
  } else if (event.type === 'postback') {
    handlePostback(event);
  }
}

function handleTextMessage(event) {
  const userId  = event.source.userId;
  const text    = event.message.text.trim();
  const cache   = CacheService.getScriptCache();
  const stateKey = 'state_' + userId;
  const stateRaw = cache.get(stateKey);

  // PIN入力待ち状態の処理
  if (stateRaw) {
    const state = JSON.parse(stateRaw);
    handlePinInput(event, userId, text, state, cache, stateKey);
    return;
  }

  // コマンドルーティング
  const config = getConfig();
  switch (text) {
    case '出勤':
    case '🟢 出勤':
      cache.put(stateKey, JSON.stringify({ action: 'clock_in' }), 300);
      reply(event.replyToken, 'PINコードを入力してください（4〜6桁）');
      break;

    case '退勤':
    case '🔴 退勤':
      cache.put(stateKey, JSON.stringify({ action: 'clock_out' }), 300);
      reply(event.replyToken, 'PINコードを入力してください（4〜6桁）');
      break;

    case '有給残確認':
    case '📅 有給残確認':
      cache.put(stateKey, JSON.stringify({ action: 'check_leave' }), 300);
      reply(event.replyToken, 'PINコードを入力してください（4〜6桁）');
      break;

    case '有給申請':
    case '📝 有給申請':
      reply(event.replyToken,
        '有給申請フォームはこちらから開いてください。\n' +
        'https://liff.line.me/' + config.LIFF_LEAVE_ID
      );
      break;

    case '管理者':
    case '⚙️ 管理者':
      reply(event.replyToken,
        '管理者ダッシュボードはこちらから開いてください。\n' +
        'https://liff.line.me/' + config.LIFF_ADMIN_ID
      );
      break;

    default:
      reply(event.replyToken, 'メニューから操作を選択してください。');
  }
}

function handlePinInput(event, userId, text, state, cache, stateKey) {
  cache.remove(stateKey);

  if (!/^\d{4,6}$/.test(text)) {
    reply(event.replyToken, 'PINコードは4〜6桁の数字で入力してください。');
    return;
  }

  const employee = getEmployeeByPin(text);
  if (!employee) {
    reply(event.replyToken, 'PINコードが正しくありません。');
    return;
  }

  switch (state.action) {
    case 'clock_in':    clockIn(event.replyToken, employee);         break;
    case 'clock_out':   clockOut(event.replyToken, employee);        break;
    case 'check_leave': showLeaveBalance(event.replyToken, employee); break;
  }
}

function handlePostback(event) {
  try {
    const data     = JSON.parse(event.postback.data);
    const approverId = event.source.userId;

    if (data.action === 'approve_leave') {
      const result = processLeaveApproval(data.applicationId, approverId, '承認');
      reply(event.replyToken, result.success
        ? '✅ 有給を承認しました（' + result.employeeName + '、' + result.days + '日）'
        : 'エラー：' + result.message
      );
    } else if (data.action === 'reject_leave') {
      const result = processLeaveApproval(data.applicationId, approverId, '否認');
      reply(event.replyToken, result.success
        ? '有給を否認しました（' + result.employeeName + '）'
        : 'エラー：' + result.message
      );
    }
  } catch (err) {
    Logger.log('handlePostback error: ' + err.toString());
  }
}

// ============================================================
// 出勤打刻
// ============================================================

function clockIn(replyToken, employee) {
  const ss    = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.ATTENDANCE);
  const now   = new Date();
  const today = fmt(now, 'yyyy/MM/dd');
  const time  = fmt(now, 'HH:mm:ss');

  // 当日の二重打刻チェック
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (fmt(new Date(rows[i][0]), 'yyyy/MM/dd') === today &&
        rows[i][1] === employee.id && rows[i][3] !== '') {
      reply(replyToken,
        employee.name + 'さん、本日はすでに出勤済みです。\n出勤時刻：' + rows[i][3]
      );
      return;
    }
  }

  sheet.appendRow([today, employee.id, employee.name, time, '', '']);
  reply(replyToken,
    '✅ 出勤を記録しました\n\n' +
    '社員：' + employee.name + '\n' +
    '日付：' + today + '\n' +
    '時刻：' + time
  );
}

// ============================================================
// 退勤打刻
// ============================================================

function clockOut(replyToken, employee) {
  const ss    = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.ATTENDANCE);
  const now   = new Date();
  const today = fmt(now, 'yyyy/MM/dd');
  const time  = fmt(now, 'HH:mm:ss');

  const rows = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (fmt(new Date(rows[i][0]), 'yyyy/MM/dd') === today && rows[i][1] === employee.id) {
      if (rows[i][3] === '') {
        reply(replyToken,
          employee.name + 'さん、本日の出勤打刻が見つかりません。\nまず出勤打刻を行ってください。'
        );
        return;
      }
      if (rows[i][4] !== '') {
        reply(replyToken,
          employee.name + 'さん、本日はすでに退勤済みです。\n退勤時刻：' + rows[i][4]
        );
        return;
      }

      const clockInTime  = parseTime(today, rows[i][3]);
      const clockOutTime = new Date();
      const hours        = Math.round((clockOutTime - clockInTime) / 36000) / 100;

      sheet.getRange(i + 1, 5).setValue(time);
      sheet.getRange(i + 1, 6).setValue(hours);

      reply(replyToken,
        '✅ 退勤を記録しました\n\n' +
        '社員：' + employee.name + '\n' +
        '日付：' + today + '\n' +
        '退勤時刻：' + time + '\n' +
        '勤務時間：' + hours + '時間'
      );
      return;
    }
  }

  reply(replyToken,
    employee.name + 'さん、本日の出勤記録が見つかりません。\nまず出勤打刻を行ってください。'
  );
}

// ============================================================
// 有給残確認
// ============================================================

function showLeaveBalance(replyToken, employee) {
  const year    = new Date().getFullYear();
  const balance = getLeaveBalance(employee.id, year);
  reply(replyToken,
    '📅 有給残日数\n\n' +
    '社員：' + employee.name + '\n' +
    year + '年度\n' +
    '付与日数：' + balance.total + '日\n' +
    '使用済み：' + balance.used + '日\n' +
    '残り：' + balance.remaining + '日'
  );
}

// ============================================================
// 有給申請・承認（LIFFから呼び出し）
// ============================================================

function submitLeaveRequest(data) {
  const employee = getEmployeeByLineId(data.userId);
  if (!employee) return { success: false, message: '登録されていないアカウントです。管理者にご確認ください。' };

  const ss   = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID);
  const now  = new Date();
  const appId = 'L' + fmt(now, 'yyyyMMddHHmmss');

  // 半休区分：全休 / 午前(10-14時) / 午後(14-18時)
  const halfType = (data.halfType === '午前' || data.halfType === '午後') ? data.halfType : '全休';
  let startDate = data.startDate;
  let endDate   = data.endDate;
  let days;
  if (halfType === '全休') {
    const start = new Date(startDate);
    const end   = new Date(endDate);
    if (end < start) return { success: false, message: '終了日は開始日以降を指定してください' };
    days = Math.round((end - start) / 86400000) + 1;
  } else {
    endDate = startDate;   // 半休は単日のみ
    days = 0.5;
  }

  // 同一期間で解消されていない承認済み記録が履歴上に残っていないかチェック
  // （シート上の行が手動削除されても履歴には残るため、2重承認を未然に防げる）
  const dupAppId = findUnresolvedApproval_(employee.id, startDate, endDate, halfType);
  if (dupAppId) {
    return {
      success: false,
      message: 'この期間（' + startDate + (halfType !== '全休' ? '・' + halfType : (endDate !== startDate ? '〜' + endDate : '')) +
        '）は履歴上すでに承認済みです（申請ID：' + dupAppId + '）。カレンダー等に表示が無い場合は再申請せず、管理者に確認してください。'
    };
  }

  const year  = fiscalYear(new Date(startDate));
  const balance = getLeaveBalance(employee.id, year);

  if (balance.remaining < days) {
    return {
      success: false,
      message: '有給残日数が不足しています（残り' + balance.remaining + '日）'
    };
  }

  const sheet = ss.getSheetByName(SHEETS.LEAVE);
  sheet.appendRow([
    appId,
    employee.id,
    employee.name,
    startDate,
    endDate,
    days,
    data.reason || '',
    '申請中',
    fmt(now, 'yyyy/MM/dd HH:mm'),
    '',
    '',
    halfType
  ]);

  logLeaveChange('申請', employee.id, employee.name, startDate, endDate, halfType, days,
    balance.used, balance.used, balance.remaining, balance.remaining, employee.id, appId, '');

  notifyAdminsOfLeaveRequest(employee, appId, startDate, endDate, days, data.reason, halfType);

  return { success: true, message: '有給申請を受け付けました（申請ID：' + appId + '）' };
}

// 同一社員・同一期間・同一半休区分で、履歴上「承認」されたまま「取消」されていない記録があればその申請IDを返す
function findUnresolvedApproval_(employeeId, startDate, endDate, halfType) {
  const sheet = getOrCreateLeaveLogSheet();
  const rows  = sheet.getDataRange().getValues();
  const toStr = v => (v instanceof Date) ? fmt(v, 'yyyy/MM/dd') : String(v || '');
  const approvedAppIds = [];
  const cancelledAppIds = {};
  rows.slice(1).forEach(r => {
    if (r[1] !== employeeId) return;
    const type = r[3], appId = r[13];
    if (type === '承認' && toStr(r[4]) === startDate && toStr(r[5]) === endDate && (r[6] || '全休') === halfType) {
      approvedAppIds.push(appId);
    }
    if (type === '取消' && appId) cancelledAppIds[appId] = true;
  });
  return approvedAppIds.find(id => !cancelledAppIds[id]) || null;
}

function processLeaveApproval(applicationId, approverId, status) {
  // 複数申請をほぼ同時に承認すると有給残の読み書きが競合し、加算が消える（lost update）。
  // スクリプトロックで直列化して防ぐ（最大30秒待機）。
  const lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (e) {
    return { success: false, message: '処理が混み合っています。少し待って再度お試しください。' };
  }
  try {
  const ss    = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.LEAVE);
  const rows  = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] !== applicationId) continue;
    if (rows[i][7] !== '申請中') {
      return { success: false, message: 'この申請はすでに処理済みです' };
    }

    const now = fmt(new Date(), 'yyyy/MM/dd HH:mm');
    sheet.getRange(i + 1, 8).setValue(status);
    sheet.getRange(i + 1, 10).setValue(now);
    sheet.getRange(i + 1, 11).setValue(approverId);

    const employeeId  = rows[i][1];
    const employeeName = rows[i][2];
    const startDate   = rows[i][3];
    const endDate     = rows[i][4];
    const days        = rows[i][5];
    const halfType    = rows[i][11] || '全休';
    const year        = fiscalYear(new Date(startDate));
    const emp         = getEmployeeById(employeeId);
    const periodText  = (halfType === '午前') ? (startDate + '（午前半休 10:00-14:00）')
                      : (halfType === '午後') ? (startDate + '（午後半休 14:00-18:00）')
                      : (startDate + '〜' + endDate);

    if (status === '承認') {
      const beforeBal = getLeaveBalance(employeeId, year);
      deductLeaveBalance(employeeId, days, year);
      const afterBal = getLeaveBalance(employeeId, year);
      logLeaveChange('承認', employeeId, employeeName, startDate, endDate, halfType, days,
        beforeBal.used, afterBal.used, beforeBal.remaining, afterBal.remaining, approverId, applicationId, '');
      if (emp && emp.lineUserId) {
        sendMessage(emp.lineUserId,
          '✅ 有給が承認されました\n\n' +
          '期間：' + periodText + '\n' +
          '日数：' + days + '日'
        );
      }
    } else {
      const bal = getLeaveBalance(employeeId, year);
      logLeaveChange('否認', employeeId, employeeName, startDate, endDate, halfType, days,
        bal.used, bal.used, bal.remaining, bal.remaining, approverId, applicationId, '');
      if (emp && emp.lineUserId) {
        sendMessage(emp.lineUserId,
          '❌ 有給申請が否認されました\n\n' +
          '期間：' + periodText + '\n\n' +
          'ご不明な点は管理者にご確認ください。'
        );
      }
    }

    return { success: true, employeeName, days, startDate, endDate };
  }

  return { success: false, message: '申請が見つかりません' };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// 管理者ダッシュボードデータ（LIFFから呼び出し）
// ============================================================

function getAdminDashboardData(userId) {
  const admin = getEmployeeByLineId(userId);
  if (!admin || !admin.isAdmin) {
    return { success: false, message: '管理者権限がありません' };
  }

  const ss    = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID);
  const today = fmt(new Date(), 'yyyy/MM/dd');
  const year  = fiscalYear();

  // 本日の打刻一覧
  const attRows = ss.getSheetByName(SHEETS.ATTENDANCE).getDataRange().getValues();
  const todayAtt = attRows.slice(1)
    .filter(r => r[0] && fmt(new Date(r[0]), 'yyyy/MM/dd') === today)
    .map(r => ({
      employeeId:  r[1],
      name:        r[2],
      clockIn:     toHHmm(r[3]),
      clockOut:    toHHmm(r[4]),
      hours:       r[5],
      clockInLoc:  r[6] || '',
      clockOutLoc: r[7] || '',
      flag:        r[8] || ''
    }));

  // 申請中の有給
  const leaveRows = ss.getSheetByName(SHEETS.LEAVE).getDataRange().getValues();
  const pendingLeave = leaveRows.slice(1)
    .filter(r => r[7] === '申請中')
    .map(r => ({
      id:          r[0],
      employeeId:  r[1],
      name:        r[2],
      startDate:   r[3],
      endDate:     r[4],
      days:        r[5],
      reason:      r[6],
      appliedAt:   r[8],
      halfType:    r[11] || '全休'
    }));

  // 有給残日数（今年度）
  const balRows = ss.getSheetByName(SHEETS.BALANCE).getDataRange().getValues();
  const leaveBalance = balRows.slice(1)
    .filter(r => r[2] === year)
    .map(r => ({
      employeeId: r[0],
      name:       r[1],
      total:      r[3],
      used:       r[4],
      remaining:  r[5]
    }));

  // 全社員の本日出勤状況（未打刻も含む）
  const empRows = ss.getSheetByName(SHEETS.EMPLOYEES).getDataRange().getValues();
  const allEmployees = empRows.slice(1)
    .filter(r => r[7] !== false && r[7] !== 'FALSE')      // 有効
    .filter(r => !(r[6] === true || r[6] === 'TRUE'))      // 役員（管理者）は打刻対象外なので除外
    .map(r => {
      const att = todayAtt.find(a => a.employeeId === r[0]);
      return {
        id:          r[0],
        name:        r[1],
        status:      att ? (att.clockOut ? '退勤済' : '出勤中') : '未出勤',
        clockIn:     att ? att.clockIn     : '',
        clockOut:    att ? att.clockOut    : '',
        hours:       att ? att.hours       : '',
        clockInLoc:  att ? att.clockInLoc  : '',
        clockOutLoc: att ? att.clockOutLoc : '',
        flag:        att ? att.flag        : ''
      };
    });

  return { success: true, today, allEmployees, pendingLeave, leaveBalance };
}

function getMyBalance(userId) {
  const emp = getEmployeeByLineId(userId);
  if (!emp) return { success: false, message: '登録されていないアカウントです。管理者にご確認ください。' };
  const year    = fiscalYear();
  const balance = getLeaveBalance(emp.id, year);
  return { success: true, name: emp.name, year, ...balance };
}

function getMyAttendanceHistory(userId, year, month) {
  const emp = getEmployeeByLineId(userId);
  if (!emp) return { success: false, message: '登録されていないアカウントです。管理者にご確認ください。' };

  const ss   = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID);
  const rows = ss.getSheetByName(SHEETS.ATTENDANCE).getDataRange().getValues();
  const target = year + '/' + String(month).padStart(2, '0');

  const history = rows.slice(1)
    .filter(r => {
      if (r[1] !== emp.id) return false;
      const d = fmt(new Date(r[0]), 'yyyy/MM');
      return d === target;
    })
    .map(r => ({
      date:     fmt(new Date(r[0]), 'yyyy/MM/dd'),
      clockIn:  toHHmm(r[3]),
      clockOut: toHHmm(r[4]),
      hours:    r[5]
    }));

  return { success: true, name: emp.name, year, month, history };
}

// ============================================================
// 社員設定（雇用形態・時給）／給与集計：管理者用
// ============================================================

function getEmployeesAdmin(adminUserId) {
  const admin = getEmployeeByLineId(adminUserId);
  if (!admin || !admin.isAdmin) return { success: false, message: '管理者権限がありません' };

  const rows = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID)
    .getSheetByName(SHEETS.EMPLOYEES).getDataRange().getValues();
  const employees = rows.slice(1)
    .filter(r => r[7] !== false && r[7] !== 'FALSE')
    .map(r => ({
      id:         r[0],
      name:       r[1],
      isAdmin:    r[6] === true || r[6] === 'TRUE',
      empType:    r[8] || '正社員',
      wageWeekday: Number(r[9])  || 0,
      wageSat:     Number(r[10]) || 0,
      wageSunHol:  Number(r[11]) || 0,
    }));
  return { success: true, employees };
}

function saveEmployeeSettings(adminUserId, employeeId, empType, wageWeekday, wageSat, wageSunHol, deductRate, fixedDeduct) {
  const admin = getEmployeeByLineId(adminUserId);
  if (!admin || !admin.isAdmin) return { success: false, message: '管理者権限がありません' };

  // 時給・控除率・固定控除額は「値が渡された時だけ」更新する（未指定なら既存値を維持）。
  // ← 雇用形態だけ保存する操作で時給が0に上書きされる事故を防ぐ。
  const hasVal = v => v !== undefined && v !== null && v !== '' && !isNaN(Number(v));

  const sheet = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID).getSheetByName(SHEETS.EMPLOYEES);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === employeeId) {
      sheet.getRange(i + 1, 9).setValue(empType || '正社員');   // I 雇用形態
      if (hasVal(wageWeekday)) sheet.getRange(i + 1, 10).setValue(Number(wageWeekday)); // J 平日時給
      if (hasVal(wageSat))     sheet.getRange(i + 1, 11).setValue(Number(wageSat));     // K 土曜時給
      if (hasVal(wageSunHol))  sheet.getRange(i + 1, 12).setValue(Number(wageSunHol));  // L 日祝時給
      if (hasVal(deductRate))  sheet.getRange(i + 1, 13).setValue(Number(deductRate));  // M 控除率(%)
      if (hasVal(fixedDeduct)) sheet.getRange(i + 1, 15).setValue(Number(fixedDeduct)); // O 固定控除額(円)
      return { success: true };
    }
  }
  return { success: false, message: '社員が見つかりません' };
}

// ============================================================
// 従業員マスタ（入社追加・退職・新人の本人スマホ紐付け）
// ============================================================

// 退職者も含む全社員（マスタ管理画面用）
function getEmployeesFull(adminUserId) {
  const admin = getEmployeeByLineId(adminUserId);
  if (!admin || !admin.isAdmin) return { success: false, message: '管理者権限がありません' };

  const rows = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID)
    .getSheetByName(SHEETS.EMPLOYEES).getDataRange().getValues();
  const employees = rows.slice(1).map(r => ({
    id:        r[0],
    name:      r[1],
    hireDate:  r[3] ? fmt(new Date(r[3]), 'yyyy/MM/dd') : '',
    annualDays: Number(r[4]) || 0,
    isAdmin:   r[6] === true || r[6] === 'TRUE',
    active:    !(r[7] === false || r[7] === 'FALSE'),
    empType:   r[8] || '正社員',
    wageWeekday: Number(r[9])  || 0,
    wageSat:     Number(r[10]) || 0,
    wageSunHol:  Number(r[11]) || 0,
    deductRate:  Number(r[12]) || 0,
    monthlyTarget: Number(r[13]) || 0,
    fixedDeduct: Number(r[14]) || 0,
    hasLineId: !!(r[5] && String(r[5]).trim()),
  }));
  return { success: true, employees };
}

// 新規社員を追加（管理者）。社員IDは自動採番（E0xx）。LINE IDは本人が後で紐付け。
function addEmployeeAdmin(adminUserId, name, empType, hireDate, annualDays) {
  const admin = getEmployeeByLineId(adminUserId);
  if (!admin || !admin.isAdmin) return { success: false, message: '管理者権限がありません' };
  name = (name || '').trim();
  if (!name) return { success: false, message: '氏名を入力してください' };
  empType = (empType === 'パート') ? 'パート' : '正社員';
  annualDays = Number(annualDays) || 0;
  const hd = (hireDate || '').trim();

  const ss = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID);
  const empSheet = ss.getSheetByName(SHEETS.EMPLOYEES);
  const rows = empSheet.getDataRange().getValues();

  // 同名の有効社員がいたら警告（重複防止）
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][1] === name && !(rows[i][7] === false || rows[i][7] === 'FALSE')) {
      return { success: false, message: '同じ氏名の社員が既にいます（' + rows[i][0] + '）。別名にするか確認してください。' };
    }
  }

  // 次の社員IDを採番（E + 3桁）
  let maxNum = 0;
  for (let i = 1; i < rows.length; i++) {
    const m = String(rows[i][0]).match(/^E(\d+)$/);
    if (m) maxNum = Math.max(maxNum, Number(m[1]));
  }
  const newId = 'E' + String(maxNum + 1).padStart(3, '0');

  // [ID, 氏名, PIN, 入社日, 年間有給日数, LINE UserID, 管理者, 有効, 雇用形態, 平日, 土曜, 日祝]
  empSheet.appendRow([newId, name, '', hd, annualDays, '', false, true, empType, 0, 0, 0]);

  // 有給残（今年度）も作成
  if (annualDays > 0) {
    ss.getSheetByName(SHEETS.BALANCE).appendRow([newId, name, fiscalYear(), annualDays, 0, annualDays]);
  }

  return { success: true, id: newId, name: name, empType: empType };
}

// 在籍フラグ（有効）を切り替え。退職＝false（履歴は残す）、復帰＝true。
function setEmployeeActive(adminUserId, employeeId, active) {
  const admin = getEmployeeByLineId(adminUserId);
  if (!admin || !admin.isAdmin) return { success: false, message: '管理者権限がありません' };

  const sheet = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID).getSheetByName(SHEETS.EMPLOYEES);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === employeeId) {
      sheet.getRange(i + 1, 8).setValue(active);   // H 有効
      return { success: true, id: employeeId, name: rows[i][1], active: active };
    }
  }
  return { success: false, message: '社員が見つかりません' };
}

// 新人の本人スマホ紐付け用：LINE未紐付けの有効社員（役員除く）一覧
function getUnboundEmployees() {
  const rows = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID)
    .getSheetByName(SHEETS.EMPLOYEES).getDataRange().getValues();
  const list = rows.slice(1)
    .filter(r => !(r[7] === false || r[7] === 'FALSE'))          // 有効
    .filter(r => !(r[6] === true || r[6] === 'TRUE'))            // 役員除く
    .filter(r => !(r[5] && String(r[5]).trim()))                // LINE未紐付け
    .map(r => ({ id: r[0], name: r[1] }));
  return { success: true, employees: list };
}

// 本人がLINEで自分の社員レコードに自分のUserIDを紐付ける
function bindMyLineId(userId, employeeId) {
  if (!userId) return { success: false, message: 'LINEの情報が取得できませんでした' };
  // 既にこのUserIDが登録済みなら多重紐付けを防ぐ
  if (getEmployeeByLineId(userId)) {
    return { success: false, message: 'このLINEアカウントは既に登録済みです。打刻メニューをご利用ください。' };
  }
  const sheet = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID).getSheetByName(SHEETS.EMPLOYEES);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] !== employeeId) continue;
    if (rows[i][7] === false || rows[i][7] === 'FALSE') return { success: false, message: '退職済みの社員です。管理者にご確認ください。' };
    if (rows[i][5] && String(rows[i][5]).trim()) return { success: false, message: 'この社員は既に別のLINEで登録済みです。管理者にご確認ください。' };
    sheet.getRange(i + 1, 6).setValue(userId);   // F LINE UserID
    return { success: true, name: rows[i][1] };
  }
  return { success: false, message: '社員が見つかりません' };
}

// ============================================================
// パート給与の計算（現在の給与＝実打刻×時給／月末見込＝出勤日平均×予定出勤日数）
// ============================================================
// filterId を渡すとその社員だけ、未指定なら全パートを返す。
// ・現在の給与  = Σ（区分別の実勤務時間 × 区分別時給）… 実際に打刻した分だけ（額面・目安）
// ・見込        = （現在の給与 ÷ これまでの出勤日数）×（当月の予定出勤日数：公休/有給以外・半休0.5）
//                 下限は現在の給与（休みが未登録等で予定日数<出勤日数でも減らさない）
function computePartPayroll_(year, month, filterId) {
  const ss = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID);
  const ym = year + '/' + String(month).padStart(2, '0');
  const daysInMonth = new Date(year, month, 0).getDate();
  const round = x => Math.round(x * 100) / 100;
  const yen   = x => Math.round(x);
  const wHalf = h => (h === '午前' || h === '午後') ? 0.5 : 1;

  // 対象＝有効なパート社員
  const empRows = ss.getSheetByName(SHEETS.EMPLOYEES).getDataRange().getValues();
  let parts = empRows.slice(1)
    .filter(r => (r[7] !== false && r[7] !== 'FALSE') && r[8] === 'パート')
    .map(r => ({
      id: r[0], name: r[1],
      wageWeekday: Number(r[9]) || 0, wageSat: Number(r[10]) || 0, wageSunHol: Number(r[11]) || 0,
      deductRate: Number(r[12]) || 0,      // M 控除率(%)：雇用保険など収入比例分
      monthlyTarget: Number(r[13]) || 0,   // N 月額目標(円)：本人が設定
      fixedDeduct: Number(r[14]) || 0      // O 固定控除額(円)：社保など毎月固定分
    }));
  if (filterId) parts = parts.filter(p => p.id === filterId);
  if (!parts.length) return [];

  // 当月の打刻（出勤日セット・区分別時間）
  const att = {}; // empId -> { days:{}, 平日, 土曜, 日祝, total }
  ss.getSheetByName(SHEETS.ATTENDANCE).getDataRange().getValues().slice(1).forEach(r => {
    if (!r[0] || !r[3]) return; // 出勤時刻なしは無視
    const d = fmt(new Date(r[0]), 'yyyy/MM/dd');
    if (!d.startsWith(ym)) return;
    const id = r[1];
    const hours = Number(r[5]) || 0;
    const type = r[9] || '平日';
    if (!att[id]) att[id] = { days: {}, '平日': 0, '土曜': 0, '日祝': 0, total: 0 };
    att[id].days[d] = true;
    att[id][type] = (att[id][type] || 0) + hours;
    att[id].total += hours;
  });

  // 確定公休（半休区分つき） empId -> { date: 全休/午前/午後 }
  const kou = {};
  const kSheet = ss.getSheetByName(SHEETS.KOUKYUU);
  if (kSheet) kSheet.getDataRange().getValues().slice(1).forEach(r => {
    if (!r[0] || !r[1] || (r[4] || '確定') !== '確定') return;
    const d = fmt(new Date(r[1]), 'yyyy/MM/dd');
    if (!d.startsWith(ym)) return;
    (kou[r[0]] = kou[r[0]] || {})[d] = normKoukyuuHalf(r[5]);
  });

  return parts.map(p => {
    const a = att[p.id] || { days: {}, '平日': 0, '土曜': 0, '日祝': 0, total: 0 };
    const hw = round(a['平日']), hs = round(a['土曜']), hh = round(a['日祝']);
    const totalHours = round(hw + hs + hh);

    // 現在の給与（額面・目安）
    const currentPay = yen(a['平日'] * p.wageWeekday + a['土曜'] * p.wageSat + a['日祝'] * p.wageSunHol);

    // 予定出勤日数（当月・公休/承認有給以外、半休0.5）
    const kmap = kou[p.id] || {};
    const lmap = collectApprovedLeaveDates(ss, p.id, ym); // date -> 全休/午前/午後
    let offDays = 0;
    Object.keys(kmap).forEach(d => { offDays += wHalf(kmap[d]); });
    Object.keys(lmap).forEach(d => { if (!kmap[d]) offDays += wHalf(lmap[d]); });
    const plannedWorkDays = Math.max(0, daysInMonth - offDays);

    // 見込＝出勤日平均 × 予定出勤日数（下限＝現在給与）
    const workedDays = Object.keys(a.days).length;
    const avgPerDay = workedDays > 0 ? currentPay / workedDays : 0;
    let projectedPay = currentPay;
    if (workedDays > 0) {
      projectedPay = Math.max(currentPay, yen(avgPerDay * plannedWorkDays));
    }

    // 手取り目安＝額面 −（固定控除額：社保など）−（額面×控除率：雇用保険など）
    // 社保は毎月ほぼ固定なので固定額で引く。マイナスにはしない。
    const rate = Math.min(100, Math.max(0, p.deductRate)) / 100;
    const fixedDed = Math.max(0, p.fixedDeduct);
    const netCurrentPay   = Math.max(0, yen(currentPay   - fixedDed - currentPay   * rate));
    const netProjectedPay = Math.max(0, yen(projectedPay - fixedDed - projectedPay * rate));

    // 月額目標（本人設定）に対する進捗と「あと何日で達成」
    const target = p.monthlyTarget || 0;
    const remainingToGoal = target > 0 ? Math.max(0, target - currentPay) : 0;
    const daysToGoal = (target > 0 && avgPerDay > 0 && remainingToGoal > 0)
      ? Math.ceil(remainingToGoal / avgPerDay) : 0;
    const goalAchieved     = target > 0 && currentPay   >= target;   // 既に達成
    const goalWillAchieve  = target > 0 && projectedPay >= target;   // このままなら達成見込み

    return {
      id: p.id, name: p.name,
      wageWeekday: p.wageWeekday, wageSat: p.wageSat, wageSunHol: p.wageSunHol,
      deductRate: p.deductRate, fixedDeduct: p.fixedDeduct,
      hoursWeekday: hw, hoursSat: hs, hoursSunHol: hh, totalHours: totalHours,
      currentPay: currentPay, projectedPay: projectedPay,
      netCurrentPay: netCurrentPay, netProjectedPay: netProjectedPay,
      workedDays: workedDays, plannedWorkDays: round(plannedWorkDays),
      monthlyTarget: target,
      remainingToGoal: remainingToGoal, daysToGoal: daysToGoal,
      goalAchieved: goalAchieved, goalWillAchieve: goalWillAchieve
    };
  });
}

// 管理者：全パートの給与集計（現在＋見込）
function getPayrollSummary(adminUserId, year, month) {
  const admin = getEmployeeByLineId(adminUserId);
  if (!admin || !admin.isAdmin) return { success: false, message: '管理者権限がありません' };
  return { success: true, year, month, payroll: computePartPayroll_(year, month, null) };
}

// 本人：自分の今月の給与（パートのみ。正社員は isPart:false を返しカード非表示）
function getMyPayroll(userId, year, month) {
  const emp = getEmployeeByLineId(userId);
  if (!emp) return { success: false, message: '社員情報が取得できませんでした' };
  if (emp.empType !== 'パート') return { success: true, isPart: false };
  const rows = computePartPayroll_(year, month, emp.id);
  return { success: true, isPart: true, year: year, month: month, payroll: rows[0] || null };
}

// 本人：自分の月額目標を設定（パート本人が自分の休み画面から。N列 月額目標に保存）
function setMyTarget(userId, amount) {
  const emp = getEmployeeByLineId(userId);
  if (!emp) return { success: false, message: '社員情報が取得できませんでした' };
  if (emp.empType !== 'パート') return { success: false, message: 'この機能はパートさん向けです' };
  let amt = Math.round(Number(amount) || 0);
  if (amt < 0) amt = 0;
  if (amt > 9999999) amt = 9999999; // 上限（入力ミス対策）
  const sheet = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID).getSheetByName(SHEETS.EMPLOYEES);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === emp.id) {
      sheet.getRange(i + 1, 14).setValue(amt);   // N 月額目標
      return { success: true, monthlyTarget: amt };
    }
  }
  return { success: false, message: '社員が見つかりません' };
}

// ============================================================
// 月次サマリー（全社員・1行=1社員）→ CSVでメール送信
// ============================================================

// 月次サマリーの行データを作る（役員除く・有効社員）。
function buildMonthlySummaryRows_(year, month) {
  const ss = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID);
  const ym = year + '/' + String(month).padStart(2, '0');
  const daysInMonth = new Date(year, month, 0).getDate();

  // 対象社員（役員除く・有効）
  const empRows = ss.getSheetByName(SHEETS.EMPLOYEES).getDataRange().getValues();
  const employees = empRows.slice(1)
    .filter(r => (r[7] !== false && r[7] !== 'FALSE') && !(r[6] === true || r[6] === 'TRUE'))
    .map(r => ({ id: r[0], name: r[1], empType: r[8] || '正社員' }));

  // 打刻集計：出勤日(セット)・総時間・区分別時間
  const att = {}; // empId -> {days:{}, total, 平日, 土曜, 日祝}
  ss.getSheetByName(SHEETS.ATTENDANCE).getDataRange().getValues().slice(1).forEach(r => {
    if (!r[0] || !r[3]) return; // 出勤時刻なしは無視
    const d = fmt(new Date(r[0]), 'yyyy/MM/dd');
    if (!d.startsWith(ym)) return;
    const id = r[1];
    const hours = Number(r[5]) || 0;
    const type = r[9] || '平日';
    if (!att[id]) att[id] = { days: {}, total: 0, '平日': 0, '土曜': 0, '日祝': 0 };
    att[id].days[d] = true;
    att[id].total += hours;
    att[id][type] = (att[id][type] || 0) + hours;
  });

  // 確定公休（半休区分つき） empId -> { date: half }
  const kou = {};
  const kSheet = ss.getSheetByName(SHEETS.KOUKYUU);
  if (kSheet) kSheet.getDataRange().getValues().slice(1).forEach(r => {
    if (!r[0] || !r[1] || (r[4] || '確定') !== '確定') return;
    const d = fmt(new Date(r[1]), 'yyyy/MM/dd');
    if (!d.startsWith(ym)) return;
    (kou[r[0]] = kou[r[0]] || {})[d] = normKoukyuuHalf(r[5]);
  });

  const today = new Date(); today.setHours(0, 0, 0, 0);
  const todayStr = fmt(today, 'yyyy/MM/dd');
  const round = x => Math.round(x * 100) / 100;
  const wHalf = h => (h === '午前' || h === '午後') ? 0.5 : 1;

  return employees.map(e => {
    const a = att[e.id] || { days: {}, total: 0, '平日': 0, '土曜': 0, '日祝': 0 };
    const kmap = kou[e.id] || {};
    const lmap = collectApprovedLeaveDates(ss, e.id, ym); // date -> 全休/午前/午後

    let koukyuuDays = 0; Object.keys(kmap).forEach(d => koukyuuDays += wHalf(kmap[d]));
    let leaveDays = 0;   Object.keys(lmap).forEach(d => leaveDays += wHalf(lmap[d]));

    // 欠勤：過去日(当月は今日まで)で 打刻なし・公休なし・有給なし
    let absent = 0;
    for (let day = 1; day <= daysInMonth; day++) {
      const ds = year + '/' + String(month).padStart(2, '0') + '/' + String(day).padStart(2, '0');
      if (ds > todayStr) break;
      if (a.days[ds] || kmap[ds] || lmap[ds]) continue;
      absent++;
    }

    return {
      id: e.id, name: e.name, empType: e.empType,
      workDays: Object.keys(a.days).length,
      totalHours: round(a.total),
      hoursWeekday: round(a['平日']), hoursSat: round(a['土曜']), hoursSunHol: round(a['日祝']),
      koukyuuDays: koukyuuDays, leaveDays: leaveDays, absentDays: absent
    };
  });
}

// 管理者用：社員1名の今年度の有給消化日一覧（承認済みのみ・残日数タブの名前タップ用）
// 同一期間・同一半休区分の重複行は1件に集約して返す（過去の2重承認行の見た目対策）。
function getEmployeeLeaveDates(userId, employeeId) {
  const admin = getEmployeeByLineId(userId);
  if (!admin || !admin.isAdmin) return { success: false, message: '管理者権限がありません' };

  const ss = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID);
  const fy = fiscalYear();
  const fyStart = fy + '/04/01', fyEnd = (fy + 1) + '/03/31';

  const rows = ss.getSheetByName(SHEETS.LEAVE).getDataRange().getValues();
  const seen = {};
  const items = [];
  rows.slice(1).forEach(r => {
    if (r[1] !== employeeId || r[7] !== '承認' || !r[3]) return;
    const start = fmt(new Date(r[3]), 'yyyy/MM/dd');
    const end   = fmt(new Date(r[4] || r[3]), 'yyyy/MM/dd');
    if (start < fyStart || start > fyEnd) return;  // 今年度分のみ
    const half = r[11] || '全休';
    const key = start + '|' + end + '|' + half;
    if (seen[key]) return;
    seen[key] = true;
    items.push({ start: start, end: end, days: Number(r[5]) || 0, half: half });
  });
  items.sort((a, b) => a.start < b.start ? -1 : 1);
  const listDays = items.reduce((s, x) => s + x.days, 0);

  // 残日数シートの「使用」（Excel引継ぎ等、アプリ申請以外の消化を含む合計）
  let used = null;
  ss.getSheetByName(SHEETS.BALANCE).getDataRange().getValues().slice(1).forEach(r => {
    if (r[0] === employeeId && r[2] === fy) used = Number(r[4]) || 0;
  });

  return { success: true, fiscalYear: fy, items: items, listDays: listDays, used: used };
}

// CSVメールの送信先（REPORT_EMAIL未設定なら社長のGmailを既定に）
function reportRecipient_() {
  return getConfig().REPORT_EMAIL || 'your-report@example.com';
}

// ★一度だけ手動実行：メール送信権限を承認するための関数（GASエディタで実行→権限を許可）
function setupEmailAuth() {
  const to = reportRecipient_();
  MailApp.sendEmail(to, '【勤怠】メール送信権限の確認',
    'この通知が届いていれば、月次集計CSVのメール送信機能が使えます。\n送信先：' + to);
  Logger.log('テストメール送信先：' + to);
}

// CSV用に1セルをエスケープ
function csvCell_(v) {
  const s = (v === null || v === undefined) ? '' : String(v);
  return /[",\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
}

// 管理者：当月の全社員サマリーをCSVにしてメール送信
function emailMonthlySummary(adminUserId, year, month) {
  try {
  const admin = getEmployeeByLineId(adminUserId);
  if (!admin || !admin.isAdmin) return { success: false, message: '管理者権限がありません' };

  const rows = buildMonthlySummaryRows_(year, month);
  const header = ['社員ID', '氏名', '雇用形態', '出勤日数', '総勤務時間',
                  '平日勤務時間', '土曜勤務時間', '日祝勤務時間', '公休日数', '有給取得日数', '欠勤日数'];
  const lines = [header.map(csvCell_).join(',')];
  rows.forEach(r => {
    lines.push([r.id, r.name, r.empType, r.workDays, r.totalHours,
                r.hoursWeekday, r.hoursSat, r.hoursSunHol, r.koukyuuDays, r.leaveDays, r.absentDays]
                .map(csvCell_).join(','));
  });
  // Excelで文字化けしないようUTF-8 BOM + CRLF
  const BOM = String.fromCharCode(0xFEFF);
  const csv = BOM + lines.join('\r\n') + '\r\n';
  const fname = '勤怠集計_' + year + '年' + String(month).padStart(2, '0') + '月.csv';
  const blob = Utilities.newBlob(csv, 'text/csv', fname);

  const to = reportRecipient_();
  if (!to) return { success: false, message: '送信先メールが未設定です（スクリプトプロパティ REPORT_EMAIL を設定してください）' };

  const subject = '【勤怠】' + year + '年' + month + '月 月次集計';
  const body = year + '年' + month + '月の勤怠月次集計を添付します（対象：打刻対象社員、役員除く）。\n'
             + '公休・有給は半休を0.5日として集計しています。\n'
             + '※欠勤＝過去日で打刻なし・公休なし・有給なしの日数。公休未設定の休業日は欠勤に計上されます。\n';
  MailApp.sendEmail({ to: to, subject: subject, body: body, attachments: [blob] });

  return { success: true, year, month, sentTo: to, count: rows.length };
  } catch (err) {
    return { success: false, message: 'メール送信エラー：' + err.toString() };
  }
}

// ============================================================
// 管理者：勤怠の代理打刻・修正
// （スマホ故障・打刻忘れ・打刻ミスの救済。本人IDが無くても管理者が記録できる）
// ============================================================

// 指定社員の当月の打刻記録を取得（修正画面のプリロード用）
function getAttendanceForAdmin(adminUserId, employeeId, year, month) {
  const admin = getEmployeeByLineId(adminUserId);
  if (!admin || !admin.isAdmin) return { success: false, message: '管理者権限がありません' };
  const emp = getEmployeeById(employeeId);
  if (!emp) return { success: false, message: '社員が見つかりません' };

  const ss = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID);
  const ym = year + '/' + String(month).padStart(2, '0');
  const rows = ss.getSheetByName(SHEETS.ATTENDANCE).getDataRange().getValues();
  const records = {}; // 'yyyy/MM/dd' -> {clockIn, clockOut, hours, note}
  rows.slice(1).forEach(r => {
    if (r[1] !== employeeId || !r[0]) return;
    const d = fmt(new Date(r[0]), 'yyyy/MM/dd');
    if (!d.startsWith(ym)) return;
    records[d] = {
      clockIn:  toHHmm(r[3]),
      clockOut: toHHmm(r[4]),
      hours:    (r[5] !== '' && r[5] != null) ? r[5] : '',
      note:     r[8] || ''
    };
  });
  return { success: true, employeeId: emp.id, name: emp.name, year: year, month: month, records: records };
}

// 全社員 × 指定月の出勤状況を一覧で返す（読み取り専用・付け忘れ可視化用）。
// 各日のステータス：done=正常 / no_out=退勤なし / koukyuu=公休 / leave=有給 /
//                   missing=過去で打刻も休みも無い（付け忘れ/欠勤疑い） / none=未到来・当日
function getMonthOverview(adminUserId, year, month) {
  const admin = getEmployeeByLineId(adminUserId);
  if (!admin || !admin.isAdmin) return { success: false, message: '管理者権限がありません' };

  const ss  = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID);
  const ym  = year + '/' + String(month).padStart(2, '0');
  const dim = new Date(year, month, 0).getDate();
  const today = fmt(new Date(), 'yyyy/MM/dd');

  // 有効・役員以外の社員
  const empRows = ss.getSheetByName(SHEETS.EMPLOYEES).getDataRange().getValues();
  const emps = empRows.slice(1)
    .filter(r => r[7] !== false && r[7] !== 'FALSE')
    .filter(r => !(r[6] === true || r[6] === 'TRUE'))
    .map(r => ({ id: r[0], name: r[1] }));

  // 当月の打刻：empId -> { 'yyyy/MM/dd' -> {in, out} }
  const att = {};
  ss.getSheetByName(SHEETS.ATTENDANCE).getDataRange().getValues().slice(1).forEach(r => {
    if (!r[0] || !r[1]) return;
    const d = fmt(new Date(r[0]), 'yyyy/MM/dd');
    if (!d.startsWith(ym)) return;
    (att[r[1]] || (att[r[1]] = {}))[d] = { in: toHHmm(r[3]), out: toHHmm(r[4]) };
  });

  // 当月の確定公休：empId -> { 'yyyy/MM/dd' -> '全休'|'午前'|'午後' }
  const kou = {};
  const kSheet = ss.getSheetByName(SHEETS.KOUKYUU);
  if (kSheet) {
    kSheet.getDataRange().getValues().slice(1).forEach(r => {
      if (!r[0] || !r[1]) return;
      const status = r[4] || '確定';
      if (status !== '確定') return;
      const d = fmt(new Date(r[1]), 'yyyy/MM/dd');
      if (!d.startsWith(ym)) return;
      (kou[r[0]] || (kou[r[0]] = {}))[d] = r[5] || '全休';
    });
  }

  // 当月の確定シフト：empId -> { 'yyyy/MM/dd' -> {start,end} }
  const shiftMonthMap = getConfirmedShiftMapForMonth_(ss, ym);

  const employees = emps.map(e => {
    const aMap = att[e.id] || {};
    const kMap = kou[e.id] || {};
    const lMap = collectApprovedLeaveDates(ss, e.id, ym); // ds -> 全休/午前/午後
    const sMap = shiftMonthMap[e.id] || {};
    let noOut = 0, missing = 0;
    const days = [];
    for (let d = 1; d <= dim; d++) {
      const ds = ym + '/' + String(d).padStart(2, '0');
      const rec = aMap[ds];
      const kh = kMap[ds], lh = lMap[ds];
      const hasShift = !!sMap[ds];
      let st;
      if (rec && rec.in && rec.out) st = 'done';
      else if (rec && rec.in && !rec.out) st = 'no_out';
      else if (kh === '全休') st = 'koukyuu';
      else if (lh === '全休') st = 'leave';
      else if (!hasShift) st = 'off';             // シフト未確定＝非勤務日（付け忘れではない）
      else if (ds >= today) st = 'none';          // シフトはあるが当日・未来は判定しない
      else st = 'missing';                        // シフト確定日なのに過去で打刻が無い
      if (st === 'no_out') noOut++;
      if (st === 'missing') missing++;
      const half = (kh && kh !== '全休') ? ('公' + kh) : (lh && lh !== '全休') ? ('有' + lh) : '';
      days.push({ d, st, in: rec ? rec.in : '', out: rec ? rec.out : '', half, shift: sMap[ds] || null });
    }
    return { id: e.id, name: e.name, noOut, missing, days };
  });

  // 曜日（1日始まりの並び用）
  const dow = [];
  for (let d = 1; d <= dim; d++) dow.push(new Date(year, month - 1, d).getDay());

  return { success: true, year, month, daysInMonth: dim, today, dow, employees };
}

// 管理者が指定社員・指定日の出勤/退勤を作成または修正。
// clockIn / clockOut は "HH:mm"。両方空なら当該日の行を削除（誤登録の取消）。
function saveAttendanceAdmin(adminUserId, employeeId, date, clockIn, clockOut) {
  const admin = getEmployeeByLineId(adminUserId);
  if (!admin || !admin.isAdmin) return { success: false, message: '管理者権限がありません' };
  const emp = getEmployeeById(employeeId);
  if (!emp) return { success: false, message: '社員が見つかりません' };

  const dateStr = fmt(new Date(date), 'yyyy/MM/dd');
  clockIn  = (clockIn  || '').trim();
  clockOut = (clockOut || '').trim();

  const re = /^([01]\d|2[0-3]):([0-5]\d)$/;
  if (clockIn  && !re.test(clockIn))  return { success: false, message: '出勤時刻は HH:mm 形式で入力してください' };
  if (clockOut && !re.test(clockOut)) return { success: false, message: '退勤時刻は HH:mm 形式で入力してください' };
  if (clockOut && !clockIn) return { success: false, message: '退勤だけの登録はできません。出勤時刻も入れてください' };

  // 勤務時間（両方あるときのみ計算）
  let hours = '';
  if (clockIn && clockOut) {
    const ci = parseTime(dateStr, clockIn);
    const co = parseTime(dateStr, clockOut);
    if (co <= ci) return { success: false, message: '退勤時刻は出勤時刻より後にしてください' };
    hours = Math.round((co - ci) / 36000) / 100;
  }

  const ss    = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.ATTENDANCE);
  const rows  = sheet.getDataRange().getValues();
  const note  = (isKoukyuu(employeeId, dateStr) ? '休日打刻 ' : '') +
                '管理者修正(' + fmt(new Date(), 'MM/dd HH:mm') + ' ' + admin.name + ')';
  const dayType = getDayType(parseTime(dateStr, '00:00'));

  // 既存行を探す
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][1] === employeeId && rows[i][0] &&
        fmt(new Date(rows[i][0]), 'yyyy/MM/dd') === dateStr) {
      if (!clockIn && !clockOut) {       // 両方空 → 行削除
        sheet.deleteRow(i + 1);
        return { success: true, deleted: true, date: dateStr, name: emp.name };
      }
      sheet.getRange(i + 1, 4).setValue(clockIn);
      sheet.getRange(i + 1, 5).setValue(clockOut);
      sheet.getRange(i + 1, 6).setValue(hours);
      sheet.getRange(i + 1, 9).setValue(note);
      sheet.getRange(i + 1, 10).setValue(dayType);
      return { success: true, updated: true, date: dateStr, name: emp.name, clockIn, clockOut, hours };
    }
  }

  if (!clockIn && !clockOut) return { success: false, message: '時刻が未入力です' };
  // 新規行 [日付,社員ID,氏名,出勤,退勤,勤務時間,出勤場所,退勤場所,備考,区分]
  sheet.appendRow([dateStr, employeeId, emp.name, clockIn, clockOut, hours, '', '', note, dayType]);
  return { success: true, created: true, date: dateStr, name: emp.name, clockIn, clockOut, hours };
}

// ============================================================
// ヘルパー関数
// ============================================================

function getEmployeeByPin(pin) {
  const rows = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID)
    .getSheetByName(SHEETS.EMPLOYEES).getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][2]) === String(pin) &&
        rows[i][7] !== false && rows[i][7] !== 'FALSE') {
      return {
        id:         rows[i][0],
        name:       rows[i][1],
        hireDate:   rows[i][3],
        annualDays: rows[i][4],
        lineUserId: rows[i][5],
        isAdmin:    rows[i][6] === true || rows[i][6] === 'TRUE',
      };
    }
  }
  return null;
}

// LINE UserID で社員を判別（PIN廃止・各自スマホ運用）
function getEmployeeByLineId(userId) {
  if (!userId) return null;
  const rows = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID)
    .getSheetByName(SHEETS.EMPLOYEES).getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][5]) === String(userId) &&
        rows[i][7] !== false && rows[i][7] !== 'FALSE') {
      return {
        id:         rows[i][0],
        name:       rows[i][1],
        hireDate:   rows[i][3],
        annualDays: rows[i][4],
        lineUserId: rows[i][5],
        isAdmin:    rows[i][6] === true || rows[i][6] === 'TRUE',
        empType:    rows[i][8] || '正社員',
        wageWeekday: Number(rows[i][9])  || 0,
        wageSat:     Number(rows[i][10]) || 0,
        wageSunHol:  Number(rows[i][11]) || 0,
      };
    }
  }
  return null;
}

function getEmployeeById(id) {
  const rows = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID)
    .getSheetByName(SHEETS.EMPLOYEES).getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === id) {
      return {
        id:         rows[i][0],
        name:       rows[i][1],
        lineUserId: rows[i][5],
        isAdmin:    rows[i][6] === true || rows[i][6] === 'TRUE',
        empType:    rows[i][8] || '正社員',
        wageWeekday: Number(rows[i][9])  || 0,
        wageSat:     Number(rows[i][10]) || 0,
        wageSunHol:  Number(rows[i][11]) || 0,
      };
    }
  }
  return null;
}

function getLeaveBalance(employeeId, year) {
  const rows = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID)
    .getSheetByName(SHEETS.BALANCE).getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === employeeId && rows[i][2] === year) {
      return { total: rows[i][3], used: rows[i][4], remaining: rows[i][5] };
    }
  }
  return { total: 0, used: 0, remaining: 0 };
}

function deductLeaveBalance(employeeId, days, year) {
  const sheet = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID)
    .getSheetByName(SHEETS.BALANCE);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === employeeId && rows[i][2] === year) {
      const used      = rows[i][4] + days;
      const remaining = rows[i][3] - used;
      sheet.getRange(i + 1, 5).setValue(used);
      sheet.getRange(i + 1, 6).setValue(Math.max(0, remaining));
      return;
    }
  }
}

// ============================================================
// 有給変更履歴（監査ログ）
// 有給に関わる全ての変更（申請・承認・否認・取消・手動訂正）を追記専用で記録する。
// 目的：LEAVEシートの行が消えても「いつ・誰の・何日分が承認済みだったか」を追跡できるようにする
//（同一申請が2重承認され残日数が2重に減る事故の再発防止）。
// ============================================================
function getOrCreateLeaveLogSheet() {
  const ss = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEETS.LEAVE_LOG);
  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.LEAVE_LOG);
    sheet.appendRow([
      '日時', '社員ID', '氏名', '種別', '開始日', '終了日', '半休区分', '日数',
      '変更前使用', '変更後使用', '変更前残', '変更後残', '操作者', '申請ID', '備考'
    ]);
    sheet.getRange(1, 1, 1, 15).setBackground('#37474F').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// type: '申請'|'承認'|'否認'|'取消'|'手動訂正'
function logLeaveChange(type, employeeId, employeeName, startDate, endDate, halfType, days,
                         beforeUsed, afterUsed, beforeRemaining, afterRemaining, actorId, appId, note) {
  try {
    getOrCreateLeaveLogSheet().appendRow([
      fmt(new Date(), 'yyyy/MM/dd HH:mm:ss'), employeeId, employeeName || '', type,
      startDate || '', endDate || '', halfType || '', days || 0,
      beforeUsed, afterUsed, beforeRemaining, afterRemaining, actorId || '', appId || '', note || ''
    ]);
  } catch (err) {
    Logger.log('logLeaveChange error: ' + err.toString());
  }
}

// 管理者：有給変更履歴を取得（新しい順・最大200件。employeeId指定で絞り込み可）
function getLeaveLog(adminUserId, employeeId) {
  const admin = getEmployeeByLineId(adminUserId);
  if (!admin || !admin.isAdmin) return { success: false, message: '管理者権限がありません' };

  const sheet = getOrCreateLeaveLogSheet();
  const rows = sheet.getDataRange().getValues();
  let logs = rows.slice(1).map(r => ({
    at: r[0] instanceof Date ? fmt(r[0], 'yyyy/MM/dd HH:mm:ss') : r[0],
    employeeId: r[1], name: r[2], type: r[3],
    startDate: r[4], endDate: r[5], halfType: r[6], days: r[7],
    beforeUsed: r[8], afterUsed: r[9], beforeRemaining: r[10], afterRemaining: r[11],
    actorId: r[12], appId: r[13], note: r[14]
  }));
  if (employeeId) logs = logs.filter(l => l.employeeId === employeeId);
  logs.reverse();
  return { success: true, logs: logs.slice(0, 200) };
}

// 管理者：有給の使用日数を指定値に訂正（残＝付与−使用で再計算）。
// 競合などで残がずれた場合の手動修正に使う。
function adjustLeaveUsed(adminUserId, employeeId, year, used) {
  const admin = getEmployeeByLineId(adminUserId);
  if (!admin || !admin.isAdmin) return { success: false, message: '管理者権限がありません' };
  used = Number(used);
  if (isNaN(used) || used < 0) return { success: false, message: '使用日数が不正です' };

  const sheet = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID).getSheetByName(SHEETS.BALANCE);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === employeeId && rows[i][2] === year) {
      const total = rows[i][3];
      const beforeUsed = rows[i][4], beforeRemaining = rows[i][5];
      const afterRemaining = Math.max(0, total - used);
      sheet.getRange(i + 1, 5).setValue(used);
      sheet.getRange(i + 1, 6).setValue(afterRemaining);
      logLeaveChange('手動訂正', employeeId, rows[i][1], '', '', '', used - beforeUsed,
        beforeUsed, used, beforeRemaining, afterRemaining, adminUserId, '', '管理者による手動訂正');
      return { success: true, employeeId, year, total, used, remaining: afterRemaining };
    }
  }
  return { success: false, message: '該当年度の有給残データが見つかりません' };
}

// 管理者：7ボタンのリッチメニューを作成し直す（シフト希望申告を追加した2026-08版）。
// 画像はGASに内蔵したbase64（RichMenuImage.gs）を使用。完了後に旧メニューを削除しデフォルト設定。
function setupRichMenu7(adminUserId) {
  const admin = getEmployeeByLineId(adminUserId);
  if (!admin || !admin.isAdmin) return { success: false, message: '管理者権限がありません' };

  const token = getConfig().LINE_TOKEN;
  const cfg = getConfig();
  const liff = function(id) { return 'https://liff.line.me/' + id; };
  const log = [];

  // 2500x1686 を 4列×2段（8セル・最後の1枠はブランド表示のみでタップ無効）。
  const CW = 625, CH = 843;
  const cell = (col, row) => ({ x: col * CW, y: row * CH, width: CW, height: CH });
  const richmenu = {
    size: { width: 2500, height: 1686 },
    selected: true,
    name: '出勤管理メニュー7',
    chatBarText: 'メニュー',
    areas: [
      { bounds: cell(0, 0), action: { type: 'uri', uri: liff(cfg.LIFF_ATTENDANCE_ID) } }, // 打刻
      { bounds: cell(1, 0), action: { type: 'uri', uri: liff(cfg.LIFF_LEAVE_ID) } },      // 有給申請
      { bounds: cell(2, 0), action: { type: 'uri', uri: liff(cfg.LIFF_SHIFT_ID) } },      // シフト希望申告
      { bounds: cell(3, 0), action: { type: 'uri', uri: liff(cfg.LIFF_KIBOU_ID) } },      // 休み希望
      { bounds: cell(0, 1), action: { type: 'uri', uri: liff(cfg.LIFF_CALENDAR_ID) } },   // 自分の休み
      { bounds: cell(1, 1), action: { type: 'uri', uri: liff(cfg.LIFF_TEAM_ID) } },       // 全員カレンダー
      { bounds: cell(2, 1), action: { type: 'uri', uri: liff(cfg.LIFF_ADMIN_ID) } },      // 管理者
      // (3,1) はブランド表示セル。タップ領域は定義しない（無反応）。
    ]
  };

  // 1) メニュー作成
  const createRes = UrlFetchApp.fetch('https://api.line.me/v2/bot/richmenu', {
    method: 'post', contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify(richmenu), muteHttpExceptions: true
  });
  if (createRes.getResponseCode() !== 200) {
    return { success: false, step: 'create', code: createRes.getResponseCode(), message: createRes.getContentText() };
  }
  const newId = JSON.parse(createRes.getContentText()).richMenuId;
  log.push('created ' + newId);

  // 2) 画像アップロード（GASに内蔵したbase64画像を使用）
  const imgBlob = Utilities.newBlob(Utilities.base64Decode(getRichMenu7Base64()), 'image/jpeg', 'richmenu7.jpg');
  const upRes = UrlFetchApp.fetch('https://api-data.line.me/v2/bot/richmenu/' + newId + '/content', {
    method: 'post', contentType: 'image/jpeg',
    headers: { Authorization: 'Bearer ' + token },
    payload: imgBlob.getBytes(), muteHttpExceptions: true
  });
  if (upRes.getResponseCode() !== 200) {
    return { success: false, step: 'upload', code: upRes.getResponseCode(), message: upRes.getContentText(), richMenuId: newId };
  }
  log.push('image uploaded');

  // 3) デフォルト設定（全友だち）
  const defRes = UrlFetchApp.fetch('https://api.line.me/v2/bot/user/all/richmenu/' + newId, {
    method: 'post', headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true
  });
  if (defRes.getResponseCode() !== 200) {
    return { success: false, step: 'setDefault', code: defRes.getResponseCode(), message: defRes.getContentText(), richMenuId: newId };
  }
  log.push('set as default');

  // 4) 旧メニューを削除（新メニュー以外を全削除）
  const listRes = UrlFetchApp.fetch('https://api.line.me/v2/bot/richmenu/list', {
    headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true
  });
  if (listRes.getResponseCode() === 200) {
    const menus = JSON.parse(listRes.getContentText()).richmenus || [];
    menus.forEach(m => {
      if (m.richMenuId !== newId) {
        UrlFetchApp.fetch('https://api.line.me/v2/bot/richmenu/' + m.richMenuId, {
          method: 'delete', headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true
        });
        log.push('deleted old ' + m.richMenuId);
      }
    });
  }

  return { success: true, richMenuId: newId, log: log };
}

// GASエディタから引数なしで実行するための補助関数：社員マスタの最初の管理者を使ってsetupRichMenu7を呼ぶ。
function debugSetupRichMenu7() {
  const ss = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID);
  const rows = ss.getSheetByName(SHEETS.EMPLOYEES).getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const isAdmin = rows[i][6] === true || rows[i][6] === 'TRUE';
    const lineUserId = rows[i][5];
    if (isAdmin && lineUserId) {
      return setupRichMenu7(String(lineUserId).trim());
    }
  }
  return { success: false, message: 'LINE紐付け済みの管理者が見つかりません' };
}

// 管理者全員にプッシュ通知が届くか診断。
// LINEプロフィール取得API（GET /v2/bot/profile/{userId}）でメッセージを送らずに友だち状態を確認。
// 200=友だちで届く / 404=未友だち or 無効ID。
function checkAdminsReachable(adminUserId) {
  const admin = getEmployeeByLineId(adminUserId);
  if (!admin || !admin.isAdmin) return { success: false, message: '管理者権限がありません' };
  const token = getConfig().LINE_TOKEN;
  const result = getAdmins().map(a => {
    if (!a.lineUserId) return { name: a.name, hasId: false, reachable: false, note: 'UserID未登録' };
    const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/profile/' + a.lineUserId, {
      headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    let displayName = '';
    if (code === 200) { try { displayName = JSON.parse(res.getContentText()).displayName; } catch (e) {} }
    return {
      name: a.name, hasId: true, code: code,
      reachable: code === 200,
      displayName: displayName,
      note: code === 200 ? '友だち登録あり・通知届く' : (code === 404 ? '友だち未登録 or UserID不一致→通知届かない' : 'エラー' + code)
    };
  });
  return { success: true, admins: result };
}

// 管理者：有給申請を取り消す（行削除）。承認済みだった場合は残日数を戻す（usedから日数を引く）。
function deleteLeaveRequest(adminUserId, applicationId) {
  const admin = getEmployeeByLineId(adminUserId);
  if (!admin || !admin.isAdmin) return { success: false, message: '管理者権限がありません' };

  const lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (e) { return { success: false, message: '処理が混み合っています。' }; }
  try {
    const ss = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.LEAVE);
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] !== applicationId) continue;
      const wasApproved = rows[i][7] === '承認';
      const employeeId   = rows[i][1];
      const employeeName = rows[i][2];
      const startDate     = rows[i][3];
      const endDate       = rows[i][4];
      const halfType      = rows[i][11] || '全休';
      const days = Number(rows[i][5]) || 0;
      const year = fiscalYear(new Date(rows[i][3]));
      sheet.deleteRow(i + 1);
      // 承認済みの取消は残日数を戻す（used -= days）
      let beforeUsed = null, afterUsed = null, beforeRemaining = null, afterRemaining = null;
      if (wasApproved) {
        const balSheet = ss.getSheetByName(SHEETS.BALANCE);
        const brows = balSheet.getDataRange().getValues();
        for (let j = 1; j < brows.length; j++) {
          if (brows[j][0] === employeeId && brows[j][2] === year) {
            beforeUsed = Number(brows[j][4]) || 0;
            const newUsed = Math.max(0, beforeUsed - days);
            afterUsed = newUsed;
            beforeRemaining = Number(brows[j][5]) || 0;
            afterRemaining  = (Number(brows[j][3]) || 0) - newUsed;
            balSheet.getRange(j + 1, 5).setValue(newUsed);
            balSheet.getRange(j + 1, 6).setValue(afterRemaining);
            break;
          }
        }
      }
      logLeaveChange('取消', employeeId, employeeName, startDate, endDate, halfType, wasApproved ? -days : 0,
        beforeUsed, afterUsed, beforeRemaining, afterRemaining, adminUserId, applicationId,
        wasApproved ? '承認済みの取消（残日数を復元）' : '未承認の取消');
      return { success: true, deleted: true, applicationId, restoredDays: wasApproved ? days : 0 };
    }
    return { success: false, message: '申請が見つかりません' };
  } finally { lock.releaseLock(); }
}

function getAdmins() {
  const rows = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID)
    .getSheetByName(SHEETS.EMPLOYEES).getDataRange().getValues();
  return rows.slice(1)
    .filter(r => (r[6] === true || r[6] === 'TRUE') &&
                 r[7] !== false && r[7] !== 'FALSE')
    .map(r => ({ id: r[0], name: r[1], lineUserId: r[5] }));
}

function notifyAdminsOfLeaveRequest(employee, appId, startDate, endDate, days, reason, halfType) {
  const halfLabel = (halfType === '午前') ? '（午前半休 10:00-14:00）' : (halfType === '午後') ? '（午後半休 14:00-18:00）' : '';
  const periodText = (halfType === '午前' || halfType === '午後') ? (startDate + ' ' + halfLabel) : (startDate + '〜' + endDate);
  const admins = getAdmins();
  const flexMessage = {
    type: 'flex',
    altText: '有給申請：' + employee.name,
    contents: {
      type: 'bubble',
      header: {
        type: 'box',
        layout: 'vertical',
        backgroundColor: '#2196F3',
        contents: [{
          type: 'text',
          text: '📝 有給申請',
          weight: 'bold',
          size: 'lg',
          color: '#ffffff'
        }]
      },
      body: {
        type: 'box',
        layout: 'vertical',
        spacing: 'sm',
        contents: [
          { type: 'text', text: '社員：'    + employee.name,                weight: 'bold' },
          { type: 'text', text: '期間：'    + periodText, wrap: true },
          { type: 'text', text: '日数：'    + days + '日' },
          { type: 'text', text: '理由：'    + (reason || 'なし'), wrap: true },
          { type: 'text', text: '申請ID：'  + appId, size: 'xs', color: '#888888' }
        ]
      },
      footer: {
        type: 'box',
        layout: 'vertical',
        spacing: 'sm',
        contents: [
          {
            type: 'button',
            style: 'primary',
            color: '#2196F3',
            action: {
              type: 'uri',
              label: '管理者ダッシュボードを開く',
              uri: 'https://liff.line.me/' + getConfig().LIFF_ADMIN_ID
            }
          },
          {
            type: 'text',
            text: '「📝有給」タブで承認・否認できます',
            size: 'xs',
            color: '#888888',
            align: 'center',
            wrap: true
          }
        ]
      }
    }
  };

  admins.forEach(admin => {
    if (admin.lineUserId) sendMessage(admin.lineUserId, flexMessage);
  });
}

// 範囲外打刻を管理者へ個別通知（DM・グループには送らない）
function notifyAdminsOutOfHours(name, date, time, kind, loc) {
  const admins = getAdmins();
  const mapUrl = (loc && loc.lat) ? '\n📍 https://maps.google.com/?q=' + loc.lat + ',' + loc.lng : '';
  const msg = '⏰ 時間外打刻のお知らせ\n\n' +
    '社員：' + name + '\n' +
    kind + '：' + date + ' ' + time + '\n' +
    '（通常の 9:00〜20:00 の範囲外です）' + mapUrl;
  admins.forEach(a => { if (a.lineUserId) sendMessage(a.lineUserId, msg); });
}

// 公休日の打刻を管理者へ個別通知
function notifyAdminsHolidayPunch(name, date, time, loc) {
  const admins = getAdmins();
  const mapUrl = (loc && loc.lat) ? '\n📍 https://maps.google.com/?q=' + loc.lat + ',' + loc.lng : '';
  const msg = '⚠️ 公休日の打刻\n\n' +
    '社員：' + name + '\n' +
    '出勤：' + date + ' ' + time + '\n' +
    '（公休に設定された日に打刻されました）' + mapUrl;
  admins.forEach(a => { if (a.lineUserId) sendMessage(a.lineUserId, msg); });
}

// ============================================================
// 急な休み（本人が打刻画面から本日を欠勤にする）
// 打刻画面の「出勤／退勤」ボタンとは別の目立たない場所に置き、理由入力＋確認ダイアログを経て送信。
// 送信すると：①急な休み申請シートに記録（＝当日の出勤忘れ通知の対象から除外）
//            ②本日が確定公休/承認済み有給ではない有効スタッフ全員＋役員へLINE DM通知。
// 使えるのは「当日・まだ出勤打刻していない」場合のみ（バックエンドでも再チェック）。
// ============================================================

function getSuddenAbsenceSheet(ss) {
  let sh = ss.getSheetByName(SHEETS.SUDDEN_ABSENCE);
  if (!sh) {
    sh = ss.insertSheet(SHEETS.SUDDEN_ABSENCE);
    sh.appendRow(['日時', '日付', '社員ID', '氏名', '理由']);
    sh.getRange(1, 1, 1, 5).setBackground('#37474F').setFontColor('#ffffff').setFontWeight('bold');
  }
  return sh;
}

// 指定社員の当日分の急な休み理由（無ければnull）
function getSuddenAbsenceToday(ss, empId, today) {
  const sh = ss.getSheetByName(SHEETS.SUDDEN_ABSENCE);
  if (!sh) return null;
  const rows = sh.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (rows[i][2] === empId && rows[i][1] === today) return rows[i][4] || '';
  }
  return null;
}

// 当日が確定公休（全休/午前）または承認済み有給（全休/午前）の社員IDの集合（＝本日休みの人）
function getDayOffEmployeeIdsToday(ss, today) {
  const off = {};
  const kSheet = ss.getSheetByName(SHEETS.KOUKYUU);
  if (kSheet) kSheet.getDataRange().getValues().slice(1).forEach(r => {
    if (!r[0] || !r[1]) return;
    if ((r[4] || '確定') !== '確定') return;
    if (fmt(new Date(r[1]), 'yyyy/MM/dd') !== today) return;
    const half = r[5] || '全休';
    if (half === '全休' || half === '午前') off[r[0]] = true;
  });
  const lSheet = ss.getSheetByName(SHEETS.LEAVE);
  if (lSheet) lSheet.getDataRange().getValues().slice(1).forEach(r => {
    if (!r[1] || r[7] !== '承認' || !r[3]) return;
    const half = r[11] || '全休';
    if (half === '午後') return;
    const startStr = fmt(new Date(r[3]), 'yyyy/MM/dd');
    const endStr   = fmt(new Date(r[4] || r[3]), 'yyyy/MM/dd');
    if (startStr <= today && today <= endStr) off[r[1]] = true;
  });
  return off;
}

function submitSuddenAbsence(userId, reason) {
  const emp = getEmployeeByLineId(userId);
  if (!emp) return { success: false, message: '登録されていないアカウントです。管理者にご確認ください。' };
  if (emp.isAdmin) return { success: false, message: '役員は打刻対象外です。' };

  reason = (reason || '').trim();
  if (!reason) return { success: false, message: '理由を入力してください' };

  const ss    = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID);
  const now   = new Date();
  const today = fmt(now, 'yyyy/MM/dd');

  // 既に出勤打刻済みなら使えない
  const attRows = ss.getSheetByName(SHEETS.ATTENDANCE).getDataRange().getValues();
  for (let i = 1; i < attRows.length; i++) {
    if (attRows[i][1] === emp.id && attRows[i][0] && fmt(new Date(attRows[i][0]), 'yyyy/MM/dd') === today && attRows[i][3]) {
      return { success: false, message: 'すでに出勤打刻済みのため、休みにはできません' };
    }
  }

  // 二重送信防止：本日分が既にあれば再通知せず終了
  if (getSuddenAbsenceToday(ss, emp.id, today) !== null) {
    return { success: true, alreadyDone: true, message: '本日はすでにお休み登録済みです' };
  }

  getSuddenAbsenceSheet(ss).appendRow([now, today, emp.id, emp.name, reason]);

  // 公休設定へ本日の「希望」行を自動起票（月次一覧・CSVで欠勤計上のまま残らないように。
  // 管理者が公休タブの「希望を反映」→「確定保存」で公休として確定する）
  try { autoRegisterSuddenKoukyuu_(ss, emp, today); } catch (e) { Logger.log('autoRegisterSuddenKoukyuu_ error: ' + e); }

  notifySuddenAbsence(ss, emp, today, reason);

  return { success: true, message: 'お休み登録を送信しました', name: emp.name, reason: reason };
}

// 急な休み→公休設定に本日の「希望」行を自動起票（2026-07-29）
// 既に本日が確定公休/承認済み有給、またはこの日の公休行が何かしらある場合は起票しない。
function autoRegisterSuddenKoukyuu_(ss, emp, today) {
  if (getDayOffEmployeeIdsToday(ss, today)[emp.id]) return;
  const sheet = getOrCreateKoukyuuSheet();
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === emp.id && rows[i][1] &&
        fmt(new Date(rows[i][1]), 'yyyy/MM/dd') === today) return;
  }
  sheet.appendRow([emp.id, today, emp.name, fmt(new Date(), 'yyyy/MM/dd HH:mm'), '希望', '全休']);
}

// 出勤予定スタッフ（本日が確定公休/承認済み有給ではない有効スタッフ）＋役員へDM
function notifySuddenAbsence(ss, emp, today, reason) {
  const dayOff = getDayOffEmployeeIdsToday(ss, today);
  const empRows = ss.getSheetByName(SHEETS.EMPLOYEES).getDataRange().getValues().slice(1);

  const staffTargets = empRows
    .filter(r => !(r[7] === false || r[7] === 'FALSE'))       // 有効
    .filter(r => !(r[6] === true || r[6] === 'TRUE'))          // 役員は下のgetAdmins()側で別途送るので除く
    .filter(r => r[0] !== emp.id)                               // 本人には送らない
    .filter(r => !dayOff[r[0]])                                 // 本日休みの人には送らない
    .filter(r => r[5] && String(r[5]).trim())                   // LINE紐付け済
    .map(r => String(r[5]).trim());

  const adminTargets = getAdmins().filter(a => a.lineUserId).map(a => a.lineUserId);

  const targets = Array.from(new Set(staffTargets.concat(adminTargets)));
  const msg = emp.name + 'さん本日（' + reason + '）のため、お休みとなりました。';
  targets.forEach(uid => sendMessage(uid, msg));
}

// ============================================================
// 退勤忘れ自動通知（フェーズC）
// 19:00〜22:00、出勤済かつ退勤なしの社員“本人”へ約30分毎にLINE DM。
// ドライラン中（COREMINDER_DRYRUN≠'false'）はログのみで送信しない。
// ============================================================

// 時間主導トリガー（15分毎）から呼ばれる本体。
function checkClockOutReminders(opts) {
  opts = opts || {};
  // 出勤忘れ通知（朝10時台）も同じ15分毎トリガーで動かす（時間帯判定はcheckClockInReminders側）。
  // devルートからの呼び出し（ignoreWindow/preview指定）では実行しない。
  if (!opts.ignoreWindow && !opts.preview) {
    try { checkClockInReminders(); } catch (e) { Logger.log('checkClockInReminders error: ' + e); }
    // 前日以前の退勤打刻もれを役員のみに通知（朝9時以降・1日1回・関数内でガード）
    try { notifyAdminsUnclosedAttendance(); } catch (e) { Logger.log('notifyAdminsUnclosedAttendance error: ' + e); }
  }
  const now = new Date();
  const mod = now.getHours() * 60 + now.getMinutes();

  // 通知時間帯の判定（preview/ignoreWindow時はスキップ）
  if (!opts.ignoreWindow && (mod < COREMINDER.START_MIN || mod >= COREMINDER.END_MIN)) {
    return { skipped: 'outside_window' };
  }

  const ss = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID);
  const today = fmt(now, 'yyyy/MM/dd');

  // 対象社員：有効・役員除く・LINE紐付け済
  const emps = ss.getSheetByName(SHEETS.EMPLOYEES).getDataRange().getValues().slice(1)
    .filter(r => !(r[7] === false || r[7] === 'FALSE'))
    .filter(r => !(r[6] === true || r[6] === 'TRUE'))
    .filter(r => r[5] && String(r[5]).trim())
    .map(r => ({ id: r[0], name: r[1], lineUserId: String(r[5]).trim() }));

  // 本日の打刻：出勤あり・退勤なし＝対象
  const att = {};
  ss.getSheetByName(SHEETS.ATTENDANCE).getDataRange().getValues().slice(1).forEach(r => {
    if (!r[0] || !r[1]) return;
    if (fmt(new Date(r[0]), 'yyyy/MM/dd') !== today) return;
    att[r[1]] = { clockIn: toHHmm(r[3]), clockOut: toHHmm(r[4]) };
  });
  const targets = emps.filter(e => att[e.id] && att[e.id].clockIn && !att[e.id].clockOut)
    .map(e => ({ id: e.id, name: e.name, lineUserId: e.lineUserId, clockIn: att[e.id].clockIn }));

  // プレビュー（テスト用）：送信もログもせず対象だけ返す
  if (opts.preview) return { success: true, now: fmt(now, 'HH:mm'), targets: targets.map(t => ({ id: t.id, name: t.name, clockIn: t.clockIn })) };

  const dryRun = (opts.dryRun != null)
    ? opts.dryRun
    : (PropertiesService.getScriptProperties().getProperty('COREMINDER_DRYRUN') !== 'false');

  // 本日の通知ログ → 直近時刻と通算回数
  const logSheet = getReminderLogSheet(ss);
  const last = {}; // empId -> { time:Date, count:int }
  logSheet.getDataRange().getValues().slice(1).forEach(r => {
    if (!r[0] || !r[1]) return;
    const t = new Date(r[0]);
    if (fmt(t, 'yyyy/MM/dd') !== today) return;
    const cur = last[r[1]];
    if (!cur || t > cur.time) last[r[1]] = { time: t, count: (r[4] || 0) };
    last[r[1]].count = Math.max(last[r[1]].count, (r[4] || 0));
  });

  let sent = 0, throttled = 0, failed = 0;
  targets.forEach(t => {
    const prev = last[t.id];
    if (prev && (now - prev.time) < COREMINDER.GAP_MS) { throttled++; return; }
    const count = (prev ? prev.count : 0) + 1;
    let result;
    if (dryRun) {
      result = 'DRYRUN';
    } else {
      try { sendClockOutReminder(t, count); result = 'SENT'; sent++; }
      catch (e) { result = 'FAIL:' + (e && e.message ? e.message : e); failed++; }
    }
    if (dryRun) sent++; // ドライランも「送ったつもり」件数に計上（ログ確認用）
    logSheet.appendRow([now, t.id, t.name, fmt(now, 'HH:mm'), count, dryRun ? 'DRYRUN' : '本番', result]);
  });

  return { success: true, time: fmt(now, 'HH:mm'), dryRun: dryRun, targets: targets.length, sent: sent, throttled: throttled, failed: failed };
}

// 管理者：出勤・退勤リマインド通知の当日ログを取得（読み取り専用）。
// 通知が実際に送られたか（結果=SENT）をスプレッドシートを開かずに確認できる。
function getReminderLogs(adminUserId) {
  const admin = getEmployeeByLineId(adminUserId);
  if (!admin || !admin.isAdmin) return { success: false, message: '管理者権限がありません' };
  const ss = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID);
  const today = fmt(new Date(), 'yyyy/MM/dd');
  const readLog = sh => sh.getDataRange().getValues().slice(1)
    .filter(r => r[0] && fmt(new Date(r[0]), 'yyyy/MM/dd') === today)
    .map(r => ({ time: fmt(new Date(r[0]), 'HH:mm'), id: r[1], name: r[2], count: r[4], mode: r[5], result: r[6] }));
  const props = PropertiesService.getScriptProperties();
  return {
    success: true, today: today,
    clockIn:  { dryRun: props.getProperty('CIREMINDER_DRYRUN') !== 'false', rows: readLog(getClockInLogSheet(ss)) },
    clockOut: { dryRun: props.getProperty('COREMINDER_DRYRUN') !== 'false', rows: readLog(getReminderLogSheet(ss)) },
  };
}

function getReminderLogSheet(ss) {
  let sh = ss.getSheetByName(SHEETS.COREMINDER_LOG);
  if (!sh) {
    sh = ss.insertSheet(SHEETS.COREMINDER_LOG);
    sh.appendRow(['日時', '社員ID', '氏名', '通知時刻', '当日通算', 'モード', '結果']);
    sh.getRange(1, 1, 1, 7).setBackground('#37474F').setFontColor('#ffffff').setFontWeight('bold');
  }
  return sh;
}

// 本人へ退勤打刻リマインドのDM
function sendClockOutReminder(target, count) {
  const msg = '🔴 退勤打刻のお知らせ\n\n' +
    target.name + ' さん、本日の退勤打刻がまだのようです。\n' +
    'お仕事終わりに、LINEメニューの「打刻」→「退勤」を押してください。\n' +
    '（本日の出勤打刻：' + target.clockIn + '）\n\n' +
    '※すでに退勤済み・本日お休みの場合は、このメッセージは無視してください。';
  sendMessage(target.lineUserId, msg);
}

// トリガー設定（1回だけ実行）：15分毎ポーリング。既存の同名トリガーは作り直す。
function setupClockOutReminderTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'checkClockOutReminders') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('checkClockOutReminders').timeBased().everyMinutes(15).create();
  return ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'checkClockOutReminders').length;
}

// ============================================================
// 出勤忘れ自動通知
// 10:00〜11:00、出勤日なのに出勤打刻がない社員“本人”へ約15分毎にLINE DM。
// 併せて管理者へ未打刻者一覧を1日1回DM。
// 休み（確定公休の全休/午前半休、承認済み有給の全休/午前半休）の社員は対象外。
// 午後半休は午前勤務のため対象に含める。
// ドライラン中（CIREMINDER_DRYRUN≠'false'）はログのみで送信しない。
// トリガーは新設せず、checkClockOutReminders（15分毎）の冒頭から呼ばれる。
// ============================================================

function checkClockInReminders(opts) {
  opts = opts || {};
  const now = new Date();
  const nowMin = now.getHours() * 60 + now.getMinutes();

  const ss = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID);
  const today = fmt(now, 'yyyy/MM/dd');
  const ym = fmt(now, 'yyyy/MM');

  // 対象社員：有効・役員除く・LINE紐付け済（退勤忘れ通知と同じ条件）
  const emps = ss.getSheetByName(SHEETS.EMPLOYEES).getDataRange().getValues().slice(1)
    .filter(r => !(r[7] === false || r[7] === 'FALSE'))
    .filter(r => !(r[6] === true || r[6] === 'TRUE'))
    .filter(r => r[5] && String(r[5]).trim())
    .map(r => ({ id: r[0], name: r[1], lineUserId: String(r[5]).trim() }));

  // 本日の確定シフト：empId -> {start,end}（シフトが無い日はそもそも対象外＝勤務予定なし）
  const shiftMonthMap = getConfirmedShiftMapForMonth_(ss, ym);
  const shiftToday = {};
  Object.keys(shiftMonthMap).forEach(empId => {
    const s = shiftMonthMap[empId][today];
    if (s && s.start) shiftToday[empId] = s;
  });

  // 本日出勤打刻済みの社員
  const clockedIn = {};
  ss.getSheetByName(SHEETS.ATTENDANCE).getDataRange().getValues().slice(1).forEach(r => {
    if (!r[0] || !r[1]) return;
    if (fmt(new Date(r[0]), 'yyyy/MM/dd') !== today) return;
    if (toHHmm(r[3])) clockedIn[r[1]] = true;
  });

  // 本日が休みの社員（確定公休の全休/午前、承認済み有給の全休/午前）。
  // シフト確定後に公休・有給が後追いで入った場合の取りこぼし防止のため引き続き除外する。
  const dayOff = {};
  const kSheet = ss.getSheetByName(SHEETS.KOUKYUU);
  if (kSheet) kSheet.getDataRange().getValues().slice(1).forEach(r => {
    if (!r[0] || !r[1]) return;
    if ((r[4] || '確定') !== '確定') return;
    if (fmt(new Date(r[1]), 'yyyy/MM/dd') !== today) return;
    const half = r[5] || '全休';
    if (half === '全休' || half === '午前') dayOff[r[0]] = true;
  });
  const lSheet = ss.getSheetByName(SHEETS.LEAVE);
  if (lSheet) lSheet.getDataRange().getValues().slice(1).forEach(r => {
    if (!r[1] || r[7] !== '承認' || !r[3]) return;
    const half = r[11] || '全休';
    if (half === '午後') return;
    const startStr = fmt(new Date(r[3]), 'yyyy/MM/dd');
    const endStr   = fmt(new Date(r[4] || r[3]), 'yyyy/MM/dd');
    if (startStr <= today && today <= endStr) dayOff[r[1]] = true;
  });

  // 急な休みを本人が申請済みの社員も対象外
  const aSheet = ss.getSheetByName(SHEETS.SUDDEN_ABSENCE);
  if (aSheet) aSheet.getDataRange().getValues().slice(1).forEach(r => {
    if (r[1] === today && r[2]) dayOff[r[2]] = true;
  });

  // 対象：本日確定シフトあり・未打刻・休みでない・シフト開始からWINDOW_MIN分以内
  const targets = emps.filter(e => {
    const shift = shiftToday[e.id];
    if (!shift || clockedIn[e.id] || dayOff[e.id]) return false;
    const startMin = hhmmToMinutes_(shift.start);
    if (startMin == null) return false;
    if (!opts.ignoreWindow && (nowMin < startMin || nowMin >= startMin + CIREMINDER.WINDOW_MIN)) return false;
    return true;
  }).map(e => ({ id: e.id, name: e.name, lineUserId: e.lineUserId, shiftStart: shiftToday[e.id].start }));

  // プレビュー（テスト用）：送信もログもせず対象だけ返す
  if (opts.preview) return { success: true, now: fmt(now, 'HH:mm'), targets: targets.map(t => ({ id: t.id, name: t.name, shiftStart: t.shiftStart })) };

  const dryRun = (opts.dryRun != null)
    ? opts.dryRun
    : (PropertiesService.getScriptProperties().getProperty('CIREMINDER_DRYRUN') !== 'false');

  // 本日の通知ログ → 直近時刻と通算回数・管理者通知済みか
  const logSheet = getClockInLogSheet(ss);
  const last = {}; // empId -> { time:Date, count:int }
  let adminNotified = false;
  logSheet.getDataRange().getValues().slice(1).forEach(r => {
    if (!r[0] || !r[1]) return;
    const t = new Date(r[0]);
    if (fmt(t, 'yyyy/MM/dd') !== today) return;
    if (r[1] === 'ADMIN') { adminNotified = true; return; }
    const cur = last[r[1]];
    if (!cur || t > cur.time) last[r[1]] = { time: t, count: (r[4] || 0) };
    last[r[1]].count = Math.max(last[r[1]].count, (r[4] || 0));
  });

  let sent = 0, throttled = 0, failed = 0;
  targets.forEach(t => {
    const prev = last[t.id];
    if (prev && (now - prev.time) < CIREMINDER.GAP_MS) { throttled++; return; }
    const count = (prev ? prev.count : 0) + 1;
    let result;
    if (dryRun) {
      result = 'DRYRUN'; sent++;
    } else {
      try { sendClockInReminder(t, count); result = 'SENT'; sent++; }
      catch (e) { result = 'FAIL:' + (e && e.message ? e.message : e); failed++; }
    }
    logSheet.appendRow([now, t.id, t.name, fmt(now, 'HH:mm'), count, dryRun ? 'DRYRUN' : '本番', result]);
  });

  // 管理者への未打刻者一覧（1日1回・最初の検知時のみ）
  if (targets.length && !adminNotified) {
    let result = 'DRYRUN';
    if (!dryRun) {
      try { notifyAdminsClockInMissing(targets, now); result = 'SENT'; }
      catch (e) { result = 'FAIL:' + (e && e.message ? e.message : e); }
    }
    logSheet.appendRow([now, 'ADMIN', '管理者一覧', fmt(now, 'HH:mm'), 1, dryRun ? 'DRYRUN' : '本番', result]);
  }

  return { success: true, time: fmt(now, 'HH:mm'), dryRun: dryRun, targets: targets.length, sent: sent, throttled: throttled, failed: failed };
}

function getClockInLogSheet(ss) {
  let sh = ss.getSheetByName(SHEETS.CIREMINDER_LOG);
  if (!sh) {
    sh = ss.insertSheet(SHEETS.CIREMINDER_LOG);
    sh.appendRow(['日時', '社員ID', '氏名', '通知時刻', '当日通算', 'モード', '結果']);
    sh.getRange(1, 1, 1, 7).setBackground('#37474F').setFontColor('#ffffff').setFontWeight('bold');
  }
  return sh;
}

// 本人へ出勤打刻リマインドのDM
function sendClockInReminder(target, count) {
  const msg = '🟠 出勤打刻のお知らせ\n\n' +
    target.name + ' さん、本日' + target.shiftStart + '〜のシフトですが、出勤打刻がまだのようです。\n' +
    '出勤されている場合は、LINEメニューの「打刻」→「出勤」を押してください。\n\n' +
    '※すでに打刻済みの場合は、このメッセージは無視してください。';
  sendMessage(target.lineUserId, msg);
}

// 管理者へ未打刻者一覧のDM（1日1回）
function notifyAdminsClockInMissing(targets, now) {
  const admins = getAdmins();
  const msg = '🟠 出勤打刻もれ一覧（' + fmt(now, 'HH:mm') + '時点）\n\n' +
    targets.map(t => '・' + t.name + '（シフト' + t.shiftStart + '〜）').join('\n') + '\n\n' +
    '本日シフト確定済みなのに出勤打刻がないスタッフです。\n' +
    '（公休・有給・午前半休の方は除外済み）\n' +
    '本人にもLINEでお知らせしています。';
  admins.forEach(a => { if (a.lineUserId) sendMessage(a.lineUserId, msg); });
}

// ============================================================
// 前日以前の退勤打刻もれを「役員のみ」に通知
// 出勤打刻はあるのに退勤打刻がない“過去の日”の記録を拾い、役員へLINE DM。
// 防犯カメラ映像で退勤時刻を確認し、管理画面から手動で退勤時刻を入力する運用のための通知。
// ・朝9時以降・1日1回だけ送信（checkClockOutReminders の15分毎トリガーから呼ばれる）
// ・当日分は既存の本人向けリマインドが担当するため対象外（過去日のみ）
// ・過去14日分をチェック（古すぎる記録は無視）
// ・opts.preview=true で送信せず対象一覧だけ返す（テスト用）
// ============================================================
function notifyAdminsUnclosedAttendance(opts) {
  opts = opts || {};
  const ss = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID);
  const now = new Date();
  const today = fmt(now, 'yyyy/MM/dd');
  const props = PropertiesService.getScriptProperties();

  // 朝9時より前は送らない（プレビュー時は時刻ゲート無視）
  if (!opts.preview && now.getHours() < 9) return { skipped: 'before_morning' };
  // 1日1回ガード（プレビュー時はスキップせず・フラグも立てない）
  if (!opts.preview && props.getProperty('UNCLOSED_NOTIFIED_DATE') === today) return { skipped: 'already_today' };

  // 社員マスタ：ID→氏名
  const empName = {};
  ss.getSheetByName(SHEETS.EMPLOYEES).getDataRange().getValues().slice(1).forEach(r => {
    if (r[0]) empName[r[0]] = r[1];
  });

  // 打刻記録：出勤あり・退勤なし・日付が今日より前（過去14日以内）
  const LOOKBACK_DAYS = 14;
  const minMs = now.getTime() - LOOKBACK_DAYS * 24 * 60 * 60 * 1000;
  const unclosed = [];
  ss.getSheetByName(SHEETS.ATTENDANCE).getDataRange().getValues().slice(1).forEach(r => {
    if (!r[0] || !r[1]) return;                 // 日付・社員ID必須
    const d = new Date(r[0]);
    const ds = fmt(d, 'yyyy/MM/dd');
    if (ds >= today) return;                    // 当日以降は対象外（当日は本人リマインドが担当）
    if (d.getTime() < minMs) return;            // 古すぎる記録は無視
    const cin = toHHmm(r[3]);
    const cout = toHHmm(r[4]);
    if (cin && !cout) unclosed.push({ id: r[1], name: empName[r[1]] || r[1], date: ds, in: cin });
  });

  // 日付の新しい順に並べる
  unclosed.sort((a, b) => (a.date < b.date ? 1 : (a.date > b.date ? -1 : 0)));

  if (opts.preview) return { success: true, preview: true, count: unclosed.length, unclosed: unclosed };

  if (!unclosed.length) {
    props.setProperty('UNCLOSED_NOTIFIED_DATE', today);
    return { success: true, unclosed: 0 };
  }

  const msg = '🕵️ 退勤打刻もれ（前日以前）\n\n' +
    unclosed.map(u => '・' + u.date + '　' + u.name + '（出勤 ' + u.in + '／退勤 未打刻）').join('\n') +
    '\n\n出勤はあるが退勤が未打刻の記録です。\n' +
    '防犯カメラ等で退勤時刻を確認し、管理画面から手動で退勤時刻を入力してください。';

  const admins = getAdmins();
  let sent = 0;
  admins.forEach(a => { if (a.lineUserId) { sendMessage(a.lineUserId, msg); sent++; } });

  props.setProperty('UNCLOSED_NOTIFIED_DATE', today);
  return { success: true, unclosed: unclosed.length, notifiedAdmins: sent };
}

// テスト用：送信せず対象一覧だけログに出す（GASエディタから ▶実行 しても絶対に送信しない）。
// 実行後、表示 → ログ で対象を確認できる。
function previewUnclosedAttendance() {
  const r = notifyAdminsUnclosedAttendance({ preview: true });
  Logger.log('退勤打刻もれ（前日以前）対象: ' + r.count + '件\n' + JSON.stringify(r.unclosed, null, 2));
  return r;
}


// ============================================================
// LINE API
// ============================================================

function reply(replyToken, message) {
  if (!replyToken) return;
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    method:  'post',
    headers: {
      'Content-Type':  'application/json',
      'Authorization': 'Bearer ' + getConfig().LINE_TOKEN
    },
    payload: JSON.stringify({
      replyToken,
      messages: [toLineMessage(message)]
    }),
    muteHttpExceptions: true
  });
}

function sendMessage(userId, message) {
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method:  'post',
    headers: {
      'Content-Type':  'application/json',
      'Authorization': 'Bearer ' + getConfig().LINE_TOKEN
    },
    payload: JSON.stringify({
      to:       userId,
      messages: [toLineMessage(message)]
    }),
    muteHttpExceptions: true
  });
}

function toLineMessage(message) {
  return typeof message === 'string' ? { type: 'text', text: message } : message;
}

// ============================================================
// ユーティリティ
// ============================================================

function fmt(date, format) {
  return Utilities.formatDate(date, 'Asia/Tokyo', format);
}

// スプレッドシートに時刻が Date 型（時刻のみセル）で保存される場合に "HH:mm" 文字列へ正規化。
// 時刻のみセルはスプレッドシートのTZ(Asia/Tokyo)基準のDateとして読まれるため、同TZで整形する。
// （以前GMT整形にしていたが、実打刻が-9時間ズレて表示・勤務時間も誤算出する不具合があったため修正）
function toHHmm(v) {
  if (v === '' || v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, 'Asia/Tokyo', 'HH:mm');
  }
  return String(v);
}

function parseTime(dateStr, timeStr) {
  return new Date(dateStr.replace(/\//g, '-') + 'T' + timeStr + '+09:00');
}

// 位置情報を保存用文字列へ（記録のみ・ブロックしない）
function formatLoc(loc) {
  if (loc && (loc.lat || loc.lat === 0) && (loc.lng || loc.lng === 0)) {
    return loc.lat + ',' + loc.lng;
  }
  return '位置不明';
}

// その日が当該社員の「確定の全休公休」かどうか（希望・半休はカウントしない）。
// 半休公休は半日出勤するため、休日打刻フラグの対象外（=通常の出勤日として扱う）。
function isKoukyuu(empId, dateStr) {
  const ss = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.KOUKYUU);
  if (!sheet) return false;
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const status = rows[i][4] || '確定';
    const half   = rows[i][5] || '全休';
    if (rows[i][0] === empId && rows[i][1] && status === '確定' && half === '全休' &&
        fmt(new Date(rows[i][1]), 'yyyy/MM/dd') === dateStr) {
      return true;
    }
  }
  return false;
}

// 2026年 日本の祝日（年ごとに更新）。MM/dd形式。
var JP_HOLIDAYS_2026 = [
  '01/01','01/12','02/11','02/23','03/20','04/29','05/03','05/04','05/05','05/06',
  '07/20','08/11','09/21','09/23','10/12','11/03','11/23'
];

function isJapaneseHoliday(date) {
  const md = fmt(date, 'MM/dd');
  const y  = date.getFullYear();
  if (y === 2026) return JP_HOLIDAYS_2026.indexOf(md) >= 0;
  // 他年は祝日表未登録（土日のみ判定）。年が変わったらJP_HOLIDAYSを追加。
  return false;
}

// 曜日区分：平日／土曜／日祝（日曜・祝日は「日祝」、土曜は「土曜」）
function getDayType(date) {
  const dow = date.getDay();
  if (dow === 0 || isJapaneseHoliday(date)) return '日祝';
  if (dow === 6) return '土曜';
  return '平日';
}

// 有給の年度（4月始まり）。例：2026/6→2026年度、2027/2→2026年度
function fiscalYear(date) {
  date = date || new Date();
  const y = date.getFullYear();
  const m = date.getMonth() + 1;
  return m < 4 ? y - 1 : y;
}

// 打刻時刻が通常範囲(9:00-20:00)外かどうか
function isOutOfHours(date) {
  const h = Number(fmt(date, 'HH'));
  const m = Number(fmt(date, 'mm'));
  const minutes = h * 60 + m;
  return minutes < 9 * 60 || minutes > 20 * 60;
}

// ============================================================
// セットアップ関数（初回のみ実行）
// ============================================================

function setupSpreadsheet() {
  const ss = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID);

  function initSheet(name, headers) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length)
        .setBackground('#37474F')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
    }
    return sheet;
  }

  initSheet(SHEETS.EMPLOYEES, [
    '社員ID', '氏名', 'PINコード', '入社日', '年間有給日数',
    'LINE UserID', '管理者', '有効',
    '雇用形態', '平日時給', '土曜時給', '日祝時給', '控除率', '月額目標', '固定控除額'
  ]);
  initSheet(SHEETS.ATTENDANCE, [
    '日付', '社員ID', '氏名', '出勤時刻', '退勤時刻', '勤務時間（時間）',
    '出勤場所', '退勤場所', '備考', '区分'
  ]);
  initSheet(SHEETS.KOUKYUU, [
    '社員ID', '日付', '氏名', '登録日時', 'ステータス', '半休区分'
  ]);
  initSheet(SHEETS.LEAVE, [
    '申請ID', '社員ID', '氏名', '開始日', '終了日', '日数',
    '理由', 'ステータス', '申請日時', '処理日時', '承認者'
  ]);
  initSheet(SHEETS.LEAVE_LOG, [
    '日時', '社員ID', '氏名', '種別', '開始日', '終了日', '半休区分', '日数',
    '変更前使用', '変更後使用', '変更前残', '変更後残', '操作者', '申請ID', '備考'
  ]);
  initSheet(SHEETS.BALANCE, [
    '社員ID', '氏名', '年度', '付与日数', '使用済み', '残日数'
  ]);
  initSheet(SHEETS.SUDDEN_ABSENCE, [
    '日時', '日付', '社員ID', '氏名', '理由'
  ]);

  Logger.log('スプレッドシート初期設定完了（シフト管理シート含む）');
}

// 社員追加（スクリプトエディタから直接実行するか、以下の形式で呼び出す）
function addEmployee_(id, name, pin, hireDate, annualDays, lineUserId, isAdmin) {
  const ss   = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID);
  const year = new Date().getFullYear();

  ss.getSheetByName(SHEETS.EMPLOYEES).appendRow([
    id, name, pin, hireDate, annualDays, lineUserId || '', isAdmin || false, true
  ]);
  ss.getSheetByName(SHEETS.BALANCE).appendRow([
    id, name, year, annualDays, 0, annualDays
  ]);

  Logger.log('社員追加完了：' + name);
}

// 使用例：addEmployee_('E001','田中太郎','1234','2023/04/01',10,'Uxxxxxxxxxx',true)

// ============================================================
// 初期社員一括登録（1回だけ実行）
// PIN廃止・LINE UserIDで判別。入社日・有給日数は後日設定（暫定0）。
// 役員は isAdmin=true・打刻対象外（承認とダッシュボードのみ）。
// ============================================================
function registerInitialEmployees() {
  const ss = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID);
  const empSheet = ss.getSheetByName(SHEETS.EMPLOYEES);
  const balSheet = ss.getSheetByName(SHEETS.BALANCE);
  // 既存のデータ行（2行目以降）をクリアして冪等にする
  if (empSheet.getLastRow() > 1) {
    empSheet.deleteRows(2, empSheet.getLastRow() - 1);
  }
  if (balSheet.getLastRow() > 1) {
    balSheet.deleteRows(2, balSheet.getLastRow() - 1);
  }
  // [社員ID, 氏名, LINE UserID, 管理者] ※自社のスタッフに書き換えて1回だけ実行
  // LINE UserIDは後から本人スマホの初回登録（bindMyLineId）で紐付けもできるので、最初は空''でもよい
  const list = [
    ['E001', 'サンプル太郎', '', false],
    ['E002', 'サンプル役員', 'LINE_USER_ID_OF_ADMIN', true ], // 役員（管理者・承認DMの宛先になるためUserID必須）
  ];
  list.forEach(r => addEmployee_(r[0], r[1], '', '', 0, r[2], r[3]));
  Logger.log('社員' + list.length + '名を登録しました（入社日・有給日数は後日更新）。');
  return list.length + '名を登録しました';
}

// 有給データ取込（エクセル 有給管理_2026_最終版 より。1回実行）
// 2026年度（2026/04/01〜2028/03/31）の残高をセットし、年間付与日数も登録。
function loadLeave2026() {
  const ss = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID);
  const balSheet = ss.getSheetByName(SHEETS.BALANCE);
  const empSheet = ss.getSheetByName(SHEETS.EMPLOYEES);
  const year = 2026; // 2026年度
  // [社員ID, 氏名, 利用可能合計(期首残+付与), 使用済み, 残, 年間付与日数]
  // ※自社の有給残に書き換えて使う（不要なら実行しない）
  const data = [
    ['E001', 'サンプル太郎', 20, 0, 20, 20],
  ];
  const brows = balSheet.getDataRange().getValues();
  const erows = empSheet.getDataRange().getValues();
  data.forEach(d => {
    // 有給残日数シート
    let found = false;
    for (let i = 1; i < brows.length; i++) {
      if (brows[i][0] === d[0] && brows[i][2] === year) {
        balSheet.getRange(i + 1, 4).setValue(d[2]);
        balSheet.getRange(i + 1, 5).setValue(d[3]);
        balSheet.getRange(i + 1, 6).setValue(d[4]);
        found = true; break;
      }
    }
    if (!found) balSheet.appendRow([d[0], d[1], year, d[2], d[3], d[4]]);
    // 社員マスタの年間付与日数
    for (let i = 1; i < erows.length; i++) {
      if (erows[i][0] === d[0]) { empSheet.getRange(i + 1, 5).setValue(d[5]); break; }
    }
  });
  Logger.log('有給残を取込みました');
  return 'done';
}

// フェーズ2セットアップ（1回だけ実行）：新LIFF IDの登録＋パート3名の設定
function applyPhase2Setup() {
  // 1. 新LIFFアプリのIDをスクリプトプロパティに登録
  PropertiesService.getScriptProperties().setProperties({
    LIFF_KIBOU_ID: 'YOUR_KIBOU_LIFF_ID',
    LIFF_TEAM_ID:  'YOUR_TEAM_LIFF_ID'
  }, false);

  // 2. パート3名を雇用形態「パート」に設定（時給は後で管理者画面から入力）
  const sheet = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID).getSheetByName(SHEETS.EMPLOYEES);
  const rows = sheet.getDataRange().getValues();
  const partIds = []; // パート社員のIDに書き換え（例 ['E003']）
  for (let i = 1; i < rows.length; i++) {
    if (partIds.indexOf(rows[i][0]) >= 0) {
      sheet.getRange(i + 1, 9).setValue('パート'); // I列 雇用形態
    }
  }
  Logger.log('フェーズ2セットアップ完了：LIFF ID登録＋パート3名設定');
  return 'done';
}

// テスト用打刻記録を全消去（打刻記録シートのデータ行をクリア）
function clearAttendanceData() {
  const sheet = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID)
    .getSheetByName(SHEETS.ATTENDANCE);
  if (sheet.getLastRow() > 1) {
    sheet.deleteRows(2, sheet.getLastRow() - 1);
  }
  Logger.log('打刻記録のデータ行をクリアしました。');
  return 'cleared';
}

// 年度更新（毎年1月1日に実行するトリガーを設定する）
function updateAnnualLeave() {
  const ss      = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID);
  const empRows = ss.getSheetByName(SHEETS.EMPLOYEES).getDataRange().getValues();
  const balSheet = ss.getSheetByName(SHEETS.BALANCE);
  const year    = fiscalYear();  // 年度（4月始まり）

  // 前年度の残（最大20日まで繰越）を引き継ぐ
  const balRows = balSheet.getDataRange().getValues();
  const prevRemaining = {};
  balRows.slice(1).forEach(r => { if (r[2] === year - 1) prevRemaining[r[0]] = Number(r[5]) || 0; });

  for (let i = 1; i < empRows.length; i++) {
    if (empRows[i][7] === false || empRows[i][7] === 'FALSE') continue;
    const id   = empRows[i][0];
    const name = empRows[i][1];
    let annualDays = Number(empRows[i][4]) || 0;
    const carry = Math.min(prevRemaining[id] || 0, 20); // 繰越上限20日
    const total = carry + annualDays;
    balSheet.appendRow([id, name, year, total, 0, total]);
  }
  Logger.log(year + '年度の有給日数を更新しました');
}

// ============================================================
// LIFF打刻（メッセージを送らずウェブ画面で完結）
// ============================================================

function getTodayStatus(userId) {
  const emp = getEmployeeByLineId(userId);
  if (!emp) return { success: false, notRegistered: true, message: '登録されていないアカウントです。管理者にご確認ください。' };
  if (emp.isAdmin) return { success: true, name: emp.name, noPunch: true, message: '役員のため打刻対象外です。' };

  const ss     = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID);
  const sheet  = ss.getSheetByName(SHEETS.ATTENDANCE);
  const today  = fmt(new Date(), 'yyyy/MM/dd');
  const rows   = sheet.getDataRange().getValues();
  const absenceReason = getSuddenAbsenceToday(ss, emp.id, today);

  for (let i = rows.length - 1; i >= 1; i--) {
    if (rows[i][1] === emp.id && rows[i][0] && fmt(new Date(rows[i][0]), 'yyyy/MM/dd') === today) {
      return {
        success:  true,
        name:     emp.name,
        date:     today,
        clockIn:  toHHmm(rows[i][3]) || null,
        clockOut: toHHmm(rows[i][4]) || null,
        hours:    rows[i][5] !== '' ? rows[i][5] : null,
        status:   rows[i][4] ? 'clocked_out' : rows[i][3] ? 'clocked_in' : 'not_clocked_in',
        absenceReason: absenceReason
      };
    }
  }
  return { success: true, name: emp.name, date: today, clockIn: null, clockOut: null, hours: null, status: 'not_clocked_in', absenceReason: absenceReason };
}

function liffClockIn(userId, loc) {
  const emp = getEmployeeByLineId(userId);
  if (!emp) return { success: false, message: '登録されていないアカウントです。管理者にご確認ください。' };
  if (emp.isAdmin) return { success: false, message: '役員は打刻対象外です。' };

  const ss    = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.ATTENDANCE);
  const now   = new Date();
  const today = fmt(now, 'yyyy/MM/dd');
  const time  = fmt(now, 'HH:mm');

  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][1] === emp.id && rows[i][0] && fmt(new Date(rows[i][0]), 'yyyy/MM/dd') === today && rows[i][3]) {
      return { success: false, message: 'すでに出勤済みです（' + toHHmm(rows[i][3]) + '）', alreadyDone: true, time: toHHmm(rows[i][3]) };
    }
  }

  // 休日（公休）打刻フラグ
  const flag = isKoukyuu(emp.id, today) ? '休日打刻' : '';
  const dayType = getDayType(now);
  // [日付, 社員ID, 氏名, 出勤時刻, 退勤時刻, 勤務時間, 出勤場所, 退勤場所, 備考, 区分]
  sheet.appendRow([today, emp.id, emp.name, time, '', '', formatLoc(loc), '', flag, dayType]);

  // 範囲外（9:00-20:00外）打刻は管理者へ個別通知
  if (isOutOfHours(now)) {
    notifyAdminsOutOfHours(emp.name, today, time, '出勤', loc);
  }
  if (flag) notifyAdminsHolidayPunch(emp.name, today, time, loc);

  return { success: true, name: emp.name, date: today, time: time, action: 'clock_in', flag: flag };
}

function liffClockOut(userId, loc) {
  const emp = getEmployeeByLineId(userId);
  if (!emp) return { success: false, message: '登録されていないアカウントです。管理者にご確認ください。' };
  if (emp.isAdmin) return { success: false, message: '役員は打刻対象外です。' };

  const ss    = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.ATTENDANCE);
  const now   = new Date();
  const today = fmt(now, 'yyyy/MM/dd');
  const time  = fmt(now, 'HH:mm');

  const rows = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (rows[i][1] !== emp.id) continue;
    if (!rows[i][0] || fmt(new Date(rows[i][0]), 'yyyy/MM/dd') !== today) continue;
    if (!rows[i][3]) return { success: false, message: '出勤打刻がありません。先に出勤を押してください。' };
    if (rows[i][4])  return { success: false, message: 'すでに退勤済みです（' + toHHmm(rows[i][4]) + '）', alreadyDone: true, time: toHHmm(rows[i][4]) };

    const ciStr = toHHmm(rows[i][3]);
    const clockInTime = parseTime(today, ciStr);
    const hours = Math.round((now - clockInTime) / 36000) / 100;
    sheet.getRange(i + 1, 5).setValue(time);
    sheet.getRange(i + 1, 6).setValue(hours);
    sheet.getRange(i + 1, 8).setValue(formatLoc(loc)); // 退勤場所
    if (isOutOfHours(now)) {
      notifyAdminsOutOfHours(emp.name, today, time, '退勤', loc);
    }
    return { success: true, name: emp.name, date: today, clockIn: ciStr, clockOut: time, hours: hours, action: 'clock_out' };
  }
  return { success: false, message: '本日の出勤記録が見つかりません。' };
}

// ============================================================
// 初回セットアップ：スクリプトプロパティを設定（1回だけ実行）
function initProperties() {
  PropertiesService.getScriptProperties().setProperties({
    SPREADSHEET_ID:           '1wk2yG4qxcbPUg3VGyuOVeaqheC0yYzCskj4uMAQDZTA',
    LINE_CHANNEL_ACCESS_TOKEN:'PENDING',              // ← 秘密情報のため、コードではなくスクリプトプロパティのUIから直接設定する
    LIFF_ATTENDANCE_ID:       '2010927150-jRsKoCGX',
    LIFF_LEAVE_ID:            '2010927150-3L8ATtLP',
    LIFF_ADMIN_ID:            '2010927150-G3v0pMLx',
    LIFF_CALENDAR_ID:         '2010927150-Fshq2q0M',
    LIFF_KIBOU_ID:            '2010927150-ZVsHVBWW',
    LIFF_TEAM_ID:             '2010927150-9HJyfw05',
    LIFF_SHIFT_ID:            '2010927150-A7us46OW',
  }, false);
  Logger.log('プロパティ設定完了');
}

// ============================================================
// 確定シフトのiCal配信（TimeTree等の「外部カレンダーの取り込み」用・閲覧専用）
// 2026-08 追加：入力はLiberty勤怠システム側のみ、TimeTree側では見るだけという運用のため、
// 片方向（本システム→外部カレンダー）のみ対応。逆方向（TimeTreeの入力を本システムに取り込む）は非対応。
// ============================================================

// URLに使うランダムな鍵。スクリプトプロパティに保存し、無ければ初回アクセス時に自動生成する。
function getFeedKey_() {
  const props = PropertiesService.getScriptProperties();
  let key = props.getProperty('FEED_KEY');
  if (!key) {
    key = Utilities.getUuid().replace(/-/g, '');
    props.setProperty('FEED_KEY', key);
  }
  return key;
}

function handleShiftsFeed_(e) {
  const key = (e.parameter && e.parameter.key) || '';
  if (key !== getFeedKey_()) {
    return ContentService.createTextOutput('Forbidden: invalid key').setMimeType(ContentService.MimeType.TEXT);
  }
  return ContentService.createTextOutput(buildShiftsIcs_()).setMimeType(ContentService.MimeType.ICAL);
}

// GASエディタから実行して、TimeTreeに登録するフィードURLをログに出す（引数なしで実行できる）。
function debugPrintShiftsFeedUrl() {
  const url = ScriptApp.getService().getUrl() + '?feed=shifts&key=' + getFeedKey_();
  Logger.log(url);
  return url;
}

function buildShiftsIcs_() {
  const ss = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID);

  const empName = {};
  ss.getSheetByName(SHEETS.EMPLOYEES).getDataRange().getValues().slice(1).forEach(r => {
    if (r[0]) empName[r[0]] = r[1];
  });

  const today = new Date();
  const rangeStart = new Date(today.getTime() - 30 * 86400000);
  const rangeEnd   = new Date(today.getTime() + 90 * 86400000);
  const startStr = fmt(rangeStart, 'yyyy/MM/dd');
  const endStr   = fmt(rangeEnd, 'yyyy/MM/dd');

  const lines = [
    'BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Liberty//AttendanceShift//JA',
    'CALSCALE:GREGORIAN', 'X-WR-CALNAME:Liberty 確定シフト',
  ];

  const sheet = ss.getSheetByName(SHEETS.SHIFT);
  const rows = sheet ? sheet.getDataRange().getValues().slice(1) : [];
  const nowUtc = Utilities.formatDate(new Date(), 'Etc/UTC', "yyyyMMdd'T'HHmmss'Z'");

  rows.forEach(r => {
    const empId = r[0], dateVal = r[1], start = r[3], end = r[4], status = r[5] || '確定';
    if (!empId || !dateVal || status !== '確定') return;
    const ds = fmt(new Date(dateVal), 'yyyy/MM/dd');
    if (ds < startStr || ds > endStr) return;
    if (!isValidHHmm_(start) || !isValidHHmm_(end)) return;

    const [y, m, d] = ds.split('/').map(Number);
    const [sh, sm] = start.split(':').map(Number);
    const [eh, em] = end.split(':').map(Number);
    const dtStart = Utilities.formatDate(new Date(y, m - 1, d, sh, sm), 'Etc/UTC', "yyyyMMdd'T'HHmmss'Z'");
    const dtEnd   = Utilities.formatDate(new Date(y, m - 1, d, eh, em), 'Etc/UTC', "yyyyMMdd'T'HHmmss'Z'");
    const name = empName[empId] || empId;

    lines.push('BEGIN:VEVENT');
    lines.push('UID:shift-' + empId + '-' + ds.replace(/\//g, '') + '@liberty-attendance');
    lines.push('DTSTAMP:' + nowUtc);
    lines.push('DTSTART:' + dtStart);
    lines.push('DTEND:' + dtEnd);
    lines.push('SUMMARY:' + name + ' 出勤（' + start + '-' + end + '）');
    lines.push('END:VEVENT');
  });

  lines.push('END:VCALENDAR');
  return lines.join('\r\n');
}
