// ============================================================
// シフト管理・カレンダー機能
// (Code.gs と同じ GAS プロジェクトに追加するファイル)
// ============================================================

// ============================================================
// カレンダーデータ取得（LIFF から呼び出し）
// ============================================================

function getCalendarData(userId, year, month) {
  const emp = getEmployeeByLineId(userId);
  if (!emp) return { success: false, message: '登録されていないアカウントです。管理者にご確認ください。' };

  const ss      = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID);
  const ym      = year + '/' + String(month).padStart(2, '0');
  const today   = new Date();
  today.setHours(0, 0, 0, 0);

  // --- 打刻記録 ---
  const attRows = ss.getSheetByName(SHEETS.ATTENDANCE).getDataRange().getValues();
  const attMap  = {};
  attRows.slice(1).forEach(r => {
    if (r[1] !== emp.id) return;
    const d = r[0] ? fmt(new Date(r[0]), 'yyyy/MM/dd') : '';
    if (d.startsWith(ym)) attMap[d] = { clockIn: toHHmm(r[3]), clockOut: toHHmm(r[4]), hours: r[5] };
  });

  // --- 公休設定（確定／希望／取消希望／半休区分） ---
  const koukyuuSet = {};  // date -> { status:'確定'|'希望', half:'全休'|'午前'|'午後' }
  const koukyuuCancelSet = {};  // date -> true（確定公休を「出勤に戻す」申請中。確定行は別に残っている）
  const kSheet = ss.getSheetByName(SHEETS.KOUKYUU);
  if (kSheet) {
    kSheet.getDataRange().getValues().slice(1).forEach(r => {
      if (r[0] !== emp.id || !r[1]) return;
      const d = fmt(new Date(r[1]), 'yyyy/MM/dd');
      if (!d.startsWith(ym)) return;
      const st = r[4] || '確定';
      if (st === '取消希望') { koukyuuCancelSet[d] = true; return; }  // 確定行を上書きしない
      koukyuuSet[d] = { status: st, half: normKoukyuuHalf(r[5]) };
    });
  }

  // --- 承認済み有給（期間内の各日を休みとしてマーク） ---
  const leaveSet = collectApprovedLeaveDates(ss, emp.id, ym);

  // --- 確定シフト（出勤予定の日・時間帯） ---
  const shiftMap = (getConfirmedShiftMapForMonth_(ss, ym)[emp.id]) || {};

  // --- 月の全日付を生成 ---
  const daysInMonth = new Date(year, month, 0).getDate();
  const days = [];
  for (let d = 1; d <= daysInMonth; d++) {
    const dateObj = new Date(year, month - 1, d);
    const dateStr = fmt(dateObj, 'yyyy/MM/dd');
    const isFuture  = dateObj > today;
    const isToday   = fmt(dateObj, 'yyyy/MM/dd') === fmt(today, 'yyyy/MM/dd');
    const attendance = attMap[dateStr]   || null;
    const koukyuuInfo   = koukyuuSet[dateStr] || null;
    const koukyuuStatus = koukyuuInfo ? koukyuuInfo.status : '';
    const koukyuuHalf   = koukyuuInfo ? koukyuuInfo.half : '全休';
    const leaveType = leaveSet[dateStr] || '';  // '全休'|'午前'|'午後'|''
    const shift = shiftMap[dateStr] || null;    // {start,end}｜null＝シフト未確定

    days.push({
      date:       dateStr,
      day:        d,
      dayOfWeek:  dateObj.getDay(),
      isToday,
      isFuture,
      attendance,
      shift,
      koukyuu:    koukyuuStatus === '確定',
      koukyuuCancel: !!koukyuuCancelSet[dateStr],  // 「出勤に戻す」申請中
      koukyuuHalf: (koukyuuStatus && koukyuuHalf !== '全休') ? koukyuuHalf : false,  // 公休が半休なら '午前'|'午後'
      leave:      leaveType || false,
      status:     calcDayStatus(koukyuuStatus, attendance, isFuture, isToday, leaveType, koukyuuHalf, !!shift)
    });
  }

  return { success: true, name: emp.name, isAdmin: emp.isAdmin, year, month, days };
}

// シフト制モデル：出勤予定は「確定シフトがある日」のみ。シフトが無い日は公休/有給でなくても
// 「off（非勤務日）」として扱う（旧・公休モデルの「シフトなし＝欠勤」誤判定を解消、2026-08）。
// 公休（月上限日数の管理）は既存どおり別枠で運用し、シフトの有無とは独立して判定する。
// leaveType: '全休'=有給全休 / '午前'・'午後'=半休（半休は半日勤務するためstatusは通常、day.leaveで「半」表示）
function calcDayStatus(koukyuuStatus, attendance, isFuture, isToday, leaveType, koukyuuHalf, hasShift) {
  const isKoukyuuHalf = (koukyuuHalf === '午前' || koukyuuHalf === '午後');
  if (attendance && attendance.clockIn) {
    return attendance.clockOut ? 'attended' : 'clocked_in';   // 緑 / 薄緑（半休公休の日も半日出勤すれば打刻＝出勤扱い）
  }
  if (leaveType === '全休') return 'leave';                     // 有給全休（承認済み）
  if (leaveType === '午前' || leaveType === '午後') return 'leave_half'; // 有給半休（打刻なし時）
  if (koukyuuStatus === '確定') return isKoukyuuHalf ? 'koukyuu_half' : 'koukyuu';            // 灰（確定公休／半休）
  if (koukyuuStatus === '希望') return isKoukyuuHalf ? 'koukyuu_request_half' : 'koukyuu_request'; // 薄灰（休み希望・未確定）
  if (hasShift) {
    if (isFuture || isToday) return 'scheduled';                // 青（出勤予定＝シフト確定済み）
    return 'absent';                                            // 赤（欠勤＝シフト確定日なのに打刻なし）
  }
  return 'off';                                                 // シフト未確定＝非勤務日（欠勤扱いにしない）
}

// 承認済み有給の各日付→区分（全休/午前/午後）のマップを返す（期間 start〜end を1日ずつ展開）。
// 日付セルのTZズレを避けるため、開始/終了を文字列(yyyy/MM/dd)化してから日次ループ。
function collectApprovedLeaveDates(ss, empId, ym) {
  const set = {}; // date -> '全休'|'午前'|'午後'
  const sheet = ss.getSheetByName(SHEETS.LEAVE);
  if (!sheet) return set;
  const rows = sheet.getDataRange().getValues();
  // 有給申請列: [appId, empId, name, startDate, endDate, days, reason, status, 申請日時, 処理日時, 承認者, 半休区分]
  rows.slice(1).forEach(r => {
    if (r[1] !== empId || r[7] !== '承認' || !r[3]) return;
    const half = r[11] || '全休'; // 半休は単日のみ
    const startStr = fmt(new Date(r[3]), 'yyyy/MM/dd');
    const endStr   = fmt(new Date(r[4] || r[3]), 'yyyy/MM/dd');
    let cur = new Date(startStr.replace(/\//g, '-') + 'T00:00:00+09:00');
    const end = new Date(endStr.replace(/\//g, '-') + 'T00:00:00+09:00');
    let guard = 0;
    while (cur <= end && guard < 400) {
      const ds = fmt(cur, 'yyyy/MM/dd');
      if (ds.startsWith(ym)) set[ds] = half;
      cur = new Date(cur.getTime() + 86400000);
      guard++;
    }
  });
  return set;
}

// ============================================================
// 公休（月6日休み）管理：管理者画面から入力
// ============================================================

// 管理者用：当月の社員一覧＋各社員の公休日を取得
function getKoukyuuForAdmin(adminUserId, year, month) {
  const admin = getEmployeeByLineId(adminUserId);
  if (!admin || !admin.isAdmin) return { success: false, message: '管理者権限がありません' };

  const ss = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID);

  // 打刻対象社員（役員以外・有効）
  const empRows = ss.getSheetByName(SHEETS.EMPLOYEES).getDataRange().getValues();
  const employees = empRows.slice(1)
    .filter(r => (r[7] !== false && r[7] !== 'FALSE') && !(r[6] === true || r[6] === 'TRUE'))
    .map(r => ({ id: r[0], name: r[1] }));

  // 当月の公休（確定）と休み希望（希望）
  const ym = year + '/' + String(month).padStart(2, '0');
  const kSheet = ss.getSheetByName(SHEETS.KOUKYUU);
  const koukyuu = {};    // 確定：empId -> [{d,h}]
  const kibou   = {};    // 希望：empId -> [{d,h}]
  const cancelReq = {};  // 出勤に戻す希望（取消希望）：empId -> [{d,h}]
  if (kSheet) {
    kSheet.getDataRange().getValues().slice(1).forEach(r => {
      if (!r[0] || !r[1]) return;
      const d = fmt(new Date(r[1]), 'yyyy/MM/dd');
      if (!d.startsWith(ym)) return;
      const status = r[4] || '確定';
      const item = { d: d, h: normKoukyuuHalf(r[5]) };
      if (status === '希望')          (kibou[r[0]]     = kibou[r[0]]     || []).push(item);
      else if (status === '取消希望') (cancelReq[r[0]] = cancelReq[r[0]] || []).push(item);
      else                            (koukyuu[r[0]]   = koukyuu[r[0]]   || []).push(item);
    });
  }

  // 雇用形態（公休上限の表示用：正社員6・パート15）
  const types = {};
  empRows.slice(1).forEach(r => { types[r[0]] = r[8] || '正社員'; });

  // 承認済み有給（社員×日付→区分）：公休を確定するとき有給も同じ画面で見えるように（2026-07-30）
  const leave = {};  // empId -> [{d,h}]
  const lSheet = ss.getSheetByName(SHEETS.LEAVE);
  if (lSheet) {
    lSheet.getDataRange().getValues().slice(1).forEach(r => {
      if (!r[1] || r[7] !== '承認' || !r[3]) return;
      const half = r[11] || '全休';
      const startStr = fmt(new Date(r[3]), 'yyyy/MM/dd');
      const endStr   = fmt(new Date(r[4] || r[3]), 'yyyy/MM/dd');
      let cur = new Date(startStr.replace(/\//g, '-') + 'T00:00:00+09:00');
      const end = new Date(endStr.replace(/\//g, '-') + 'T00:00:00+09:00');
      let guard = 0;
      while (cur <= end && guard < 400) {
        const ds = fmt(cur, 'yyyy/MM/dd');
        if (ds.startsWith(ym)) (leave[r[1]] = leave[r[1]] || []).push({ d: ds, h: half });
        cur = new Date(cur.getTime() + 86400000);
        guard++;
      }
    });
  }

  const daysInMonth = new Date(year, month, 0).getDate();
  return { success: true, year, month, daysInMonth, employees, koukyuu, kibou, cancelReq, leave, types };
}

// 管理者用：ある社員の当月の公休日をまとめて保存（既存を置き換え）
function saveKoukyuu(adminUserId, employeeId, dates, year, month) {
  const admin = getEmployeeByLineId(adminUserId);
  if (!admin || !admin.isAdmin) return { success: false, message: '管理者権限がありません' };

  const sheet = getOrCreateKoukyuuSheet();
  const emp = getEmployeeById(employeeId);
  const ym  = year + '/' + String(month).padStart(2, '0');

  // 当該社員・当月の既存行を削除（下から）。希望・確定とも置き換える。
  const rows = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (rows[i][0] === employeeId && rows[i][1] &&
        fmt(new Date(rows[i][1]), 'yyyy/MM/dd').startsWith(ym)) {
      sheet.deleteRow(i + 1);
    }
  }

  // 管理者保存＝確定公休（半休区分つき）
  const now  = fmt(new Date(), 'yyyy/MM/dd HH:mm');
  const list = normKoukyuuDates(dates);
  list.forEach(x => {
    sheet.appendRow([employeeId, x.d, emp ? emp.name : '', now, '確定', x.h]);
  });

  return { success: true, count: list.length };
}

function getOrCreateKoukyuuSheet() {
  const ss = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEETS.KOUKYUU);
  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.KOUKYUU);
    sheet.appendRow(['社員ID', '日付', '氏名', '登録日時', 'ステータス', '半休区分']);
    sheet.getRange(1, 1, 1, 6).setBackground('#37474F').setFontColor('#ffffff').setFontWeight('bold');
  } else if (sheet.getLastColumn() < 6) {
    // 既存シートにF列（半休区分）が無ければ追加
    sheet.getRange(1, 6).setValue('半休区分').setBackground('#37474F').setFontColor('#ffffff').setFontWeight('bold');
  }
  return sheet;
}

// 公休の半休区分を正規化（全休/午前/午後）。それ以外は全休。
function normKoukyuuHalf(h) {
  return (h === '午前' || h === '午後') ? h : '全休';
}
// 公休1件の消費日数（全休=1.0・半休=0.5）
function koukyuuWeight(h) {
  return (h === '午前' || h === '午後') ? 0.5 : 1.0;
}
// dates引数を正規化：文字列なら {d, h:'全休'}、オブジェクトなら {d, h}
function normKoukyuuDates(dates) {
  return (dates || []).map(function(x) {
    if (x && typeof x === 'object') return { d: x.d, h: normKoukyuuHalf(x.h) };
    return { d: x, h: '全休' };
  }).filter(function(x) { return x.d; });
}

// ============================================================
// 休み希望（社員が来月の休みたい日を申請 → 管理者が確定）
// ============================================================

// 社員：自分の当月の希望（と確定）を取得（LIFFのプリロード用）
function getMyKibou(userId, year, month) {
  const emp = getEmployeeByLineId(userId);
  if (!emp) return { success: false, message: '登録されていないアカウントです。' };

  const ss = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID);
  const kSheet = ss.getSheetByName(SHEETS.KOUKYUU);
  const ym = year + '/' + String(month).padStart(2, '0');
  const kibou = [], confirmed = [], cancelReq = [];  // [{d,h}]（cancelReq＝確定公休を出勤に戻す希望）
  if (kSheet) {
    kSheet.getDataRange().getValues().slice(1).forEach(r => {
      if (r[0] !== emp.id || !r[1]) return;
      const d = fmt(new Date(r[1]), 'yyyy/MM/dd');
      if (!d.startsWith(ym)) return;
      const st = r[4] || '確定';
      const item = { d: d, h: normKoukyuuHalf(r[5]) };
      if (st === '希望') kibou.push(item);
      else if (st === '取消希望') cancelReq.push(item);
      else confirmed.push(item);
    });
  }
  const limit = emp.empType === 'パート' ? 15 : 6;
  const daysInMonth = new Date(year, month, 0).getDate();
  // 消費日数（全休1.0・半休0.5）
  const confirmedDays = confirmed.reduce((s, x) => s + koukyuuWeight(x.h), 0);
  const kibouDays     = kibou.reduce((s, x) => s + koukyuuWeight(x.h), 0);
  return { success: true, name: emp.name, empType: emp.empType, limit, year, month, daysInMonth,
           kibou, confirmed, cancelReq, confirmedDays, kibouDays };
}

// 社員：休み希望を保存（当月・翌月以降ともOK。既存の希望を置き換え、確定済みは触らない）
// cancels＝確定公休を「出勤に戻したい」日（yyyy/MM/dd の配列）。ステータス「取消希望」で保存し、
// 管理者が公休タブの「希望を反映」→「確定保存」で成立する（成立まで確定公休のまま）。
function submitKibou(userId, dates, year, month, cancels) {
  const emp = getEmployeeByLineId(userId);
  if (!emp) return { success: false, message: '登録されていないアカウントです。' };
  const list = normKoukyuuDates(dates);  // [{d,h}]
  const cancelList = (cancels || []).map(function(x) {
    return (x && typeof x === 'object') ? x.d : x;
  }).filter(function(d) { return d; });

  // 休み希望は明日以降のみ（当日の休みは打刻画面の「急な休み」に一本化・2026-07-29）
  const todayStr = fmt(new Date(), 'yyyy/MM/dd');
  const past = list.filter(x => x.d <= todayStr).map(x => x.d);
  if (past.length) {
    return { success: false, message: '当日・過去の日付は休み希望に出せません（' + past.join('、') + '）。今日の休みは打刻画面の「今日は急に休みます」から送ってください。' };
  }
  const pastCancel = cancelList.filter(d => d <= todayStr);
  if (pastCancel.length) {
    return { success: false, message: '当日・過去の公休はここでは出勤に戻せません（' + pastCancel.join('、') + '）' };
  }

  const limit = emp.empType === 'パート' ? 15 : 6;
  const sheet = getOrCreateKoukyuuSheet();
  const ym = year + '/' + String(month).padStart(2, '0');

  // 当月の「確定」公休（日付→半休区分）
  const rows = sheet.getDataRange().getValues();
  const confirmedMap = {};
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === emp.id && rows[i][1] &&
        fmt(new Date(rows[i][1]), 'yyyy/MM/dd').startsWith(ym) &&
        (rows[i][4] || '確定') === '確定') {
      confirmedMap[fmt(new Date(rows[i][1]), 'yyyy/MM/dd')] = normKoukyuuHalf(rows[i][5]);
    }
  }

  // 「出勤に戻す」は確定公休の日しか出せない
  const badCancel = cancelList.filter(d => !(d in confirmedMap));
  if (badCancel.length) {
    return { success: false, message: '確定公休でない日は出勤に戻せません（' + badCancel.join('、') + '）' };
  }

  // 上限判定：確定 − 出勤に戻す ＋ 希望（半休=0.5）。戻す分を引くので「休みの入れ替え」も出せる
  let confirmedDays = 0;
  Object.keys(confirmedMap).forEach(d => confirmedDays += koukyuuWeight(confirmedMap[d]));
  const cancelDays = cancelList.reduce((s, d) => s + koukyuuWeight(confirmedMap[d]), 0);
  const kibouDays  = list.reduce((s, x) => s + koukyuuWeight(x.h), 0);
  if (confirmedDays - cancelDays + kibouDays > limit) {
    return { success: false, message: '公休の上限は月' + limit + '日です（確定済' + confirmedDays + '日−出勤に戻す' + cancelDays + '日＋希望' + kibouDays + '日＝' + (confirmedDays - cancelDays + kibouDays) + '日）' };
  }

  // 当該社員・当月の「明日以降の希望・取消希望」のみ削除（確定は残す。
  // 当日以前の希望行＝急な休みの自動起票分は消さない）
  for (let i = rows.length - 1; i >= 1; i--) {
    const st = rows[i][4] || '確定';
    if (rows[i][0] === emp.id && rows[i][1] &&
        fmt(new Date(rows[i][1]), 'yyyy/MM/dd').startsWith(ym) &&
        fmt(new Date(rows[i][1]), 'yyyy/MM/dd') > todayStr &&
        (st === '希望' || st === '取消希望')) {
      sheet.deleteRow(i + 1);
    }
  }

  const now = fmt(new Date(), 'yyyy/MM/dd HH:mm');
  list.forEach(x => sheet.appendRow([emp.id, x.d, emp.name, now, '希望', x.h]));
  cancelList.forEach(d => sheet.appendRow([emp.id, d, emp.name, now, '取消希望', confirmedMap[d]]));

  // 管理者へ通知（DM）
  const admins = getAdmins();
  const cancelNote = cancelList.length ? '\n出勤に戻す希望：' + cancelDays + '日（' + cancelList.join('、') + '）' : '';
  admins.forEach(a => {
    if (a.lineUserId) sendMessage(a.lineUserId,
      '🗓 休み希望が届きました\n\n社員：' + emp.name + '\n対象：' + year + '年' + month + '月\n希望日数：' + kibouDays + '日' + cancelNote + '\n\n管理者ダッシュボードの「公休」タブで確認・確定してください。');
  });

  return { success: true, count: list.length, days: kibouDays, cancels: cancelList.length };
}

// 管理者：来月の休み希望を全員に個別LINEで依頼
function requestKibouFromAll(adminUserId, year, month) {
  const admin = getEmployeeByLineId(adminUserId);
  if (!admin || !admin.isAdmin) return { success: false, message: '管理者権限がありません' };

  const liffId = getConfig().LIFF_KIBOU_ID;
  const link = liffId ? ('https://liff.line.me/' + liffId) : '';

  const ss = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID);
  const empRows = ss.getSheetByName(SHEETS.EMPLOYEES).getDataRange().getValues();
  const targets = empRows.slice(1)
    .filter(r => (r[7] !== false && r[7] !== 'FALSE') && !(r[6] === true || r[6] === 'TRUE'))
    .filter(r => r[5]); // LINE UserIDがある人

  let sent = 0;
  targets.forEach(r => {
    sendMessage(r[5],
      '🗓 ' + year + '年' + month + '月の休み希望をお願いします\n\n' +
      '下のリンクから希望のお休みを選んで送信してください。\n' + link);
    sent++;
  });
  return { success: true, sent };
}

// ============================================================
// 全員カレンダー（誰がいつ出勤・休みか）：全社員が閲覧可
// ============================================================
function getTeamCalendar(userId, year, month) {
  const emp = getEmployeeByLineId(userId);
  if (!emp) return { success: false, message: '登録されていないアカウントです。' };

  const ss = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID);
  const ym = year + '/' + String(month).padStart(2, '0');

  // 打刻対象社員（役員以外・有効）
  const empRows = ss.getSheetByName(SHEETS.EMPLOYEES).getDataRange().getValues();
  const employees = empRows.slice(1)
    .filter(r => (r[7] !== false && r[7] !== 'FALSE') && !(r[6] === true || r[6] === 'TRUE'))
    .map(r => ({ id: r[0], name: r[1] }));

  // 確定公休（社員×日付）。半休区分つき。
  const kSheet = ss.getSheetByName(SHEETS.KOUKYUU);
  const offMap = {}; // 'empId|date' -> '全休'|'午前'|'午後'
  if (kSheet) {
    kSheet.getDataRange().getValues().slice(1).forEach(r => {
      if (!r[0] || !r[1] || (r[4] || '確定') !== '確定') return;
      const d = fmt(new Date(r[1]), 'yyyy/MM/dd');
      if (d.startsWith(ym)) offMap[r[0] + '|' + d] = normKoukyuuHalf(r[5]);
    });
  }

  // 承認済み有給（社員×日付→区分）。期間を1日ずつ展開。
  const leaveMap = {}; // 'empId|date' -> '全休'|'午前'|'午後'
  const lSheet = ss.getSheetByName(SHEETS.LEAVE);
  if (lSheet) {
    lSheet.getDataRange().getValues().slice(1).forEach(r => {
      if (!r[1] || r[7] !== '承認' || !r[3]) return;
      const half = r[11] || '全休';
      const startStr = fmt(new Date(r[3]), 'yyyy/MM/dd');
      const endStr   = fmt(new Date(r[4] || r[3]), 'yyyy/MM/dd');
      let cur = new Date(startStr.replace(/\//g, '-') + 'T00:00:00+09:00');
      const end = new Date(endStr.replace(/\//g, '-') + 'T00:00:00+09:00');
      let guard = 0;
      while (cur <= end && guard < 400) {
        const ds = fmt(cur, 'yyyy/MM/dd');
        if (ds.startsWith(ym)) leaveMap[r[1] + '|' + ds] = half;
        cur = new Date(cur.getTime() + 86400000);
        guard++;
      }
    });
  }

  // 確定シフト（社員×日付→{start,end}）
  const shiftMonthMap = getConfirmedShiftMapForMonth_(ss, ym);

  const daysInMonth = new Date(year, month, 0).getDate();
  const days = [];
  for (let d = 1; d <= daysInMonth; d++) {
    const dateObj = new Date(year, month - 1, d);
    const dateStr = fmt(dateObj, 'yyyy/MM/dd');
    const working = [], off = [], onLeave = [], halfLeave = [], halfKoukyuu = [];
    employees.forEach(e => {
      const key = e.id + '|' + dateStr;
      const lv = leaveMap[key];
      const kk = offMap[key];
      const shift = (shiftMonthMap[e.id] || {})[dateStr];
      if (lv === '全休') { off.push(e.name); onLeave.push(e.name); }
      else if (lv === '午前' || lv === '午後') { working.push(e.name + (shift ? '（' + shift.start + '-' + shift.end + '）' : '')); halfLeave.push(e.name + '（' + lv + '半休）'); } // 有給半休＝半日勤務扱い
      else if (kk === '午前' || kk === '午後') { working.push(e.name + (shift ? '（' + shift.start + '-' + shift.end + '）' : '')); halfKoukyuu.push(e.name + '（' + kk + '半休）'); } // 公休半休＝半日勤務扱い
      else if (kk) off.push(e.name); // 全休公休
      else if (shift) working.push(e.name + '（' + shift.start + '-' + shift.end + '）'); // シフト確定＝出勤
      else off.push(e.name); // シフト未確定＝非勤務日
    });
    days.push({
      day: d, date: dateStr, dayOfWeek: dateObj.getDay(),
      dayType: getDayType(dateObj),
      workingCount: working.length, working, off, onLeave, halfLeave, halfKoukyuu
    });
  }
  return { success: true, year, month, daysInMonth, employees, days };
}

// ============================================================
// シフト（出勤予定）管理：予約対応のシフト制向け（2026-08 追加）
// 「全日出勤がデフォルト」の公休モデルとは逆に、シフトは「出勤する日・時間帯を
// 社員が申告→管理者が確定」の積み上げ方式。公休（月上限日数の管理）は既存の
// 公休設定シート／getKoukyuuForAdmin等をそのまま併用する（本機能とは独立）。
//
// シート「シフト設定」列: A社員ID B日付 C氏名 D開始時刻(HH:mm) E終了時刻(HH:mm)
//                        Fステータス(希望/確定) G登録日時
// ============================================================

function getOrCreateShiftSheet() {
  const ss = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEETS.SHIFT);
  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.SHIFT);
    sheet.appendRow(['社員ID', '日付', '氏名', '開始時刻', '終了時刻', 'ステータス', '登録日時']);
    sheet.getRange(1, 1, 1, 7).setBackground('#37474F').setFontColor('#ffffff').setFontWeight('bold');
  }
  // D・E列（開始/終了時刻）はプレーンテキスト固定にする。
  // これが無いと、Googleスプレッドシートが "10:00" のような文字列を自動的に「時刻」型に
  // 変換してしまい、後で読み込むと 1899-12-30T... のような壊れた値になる（2026-08判明の不具合）。
  sheet.getRange(1, 4, Math.max(sheet.getMaxRows(), 1000), 2).setNumberFormat('@');
  fixCorruptedShiftTimeCells_(sheet);
  return sheet;
}

// 既存行のD/E列がスプレッドシート側で日時型に化けてしまっている場合、HH:mmの文字列に戻す（自己修復）。
function fixCorruptedShiftTimeCells_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const range = sheet.getRange(2, 4, lastRow - 1, 2);
  const values = range.getValues();
  let changed = false;
  const fixed = values.map(row => row.map(v => {
    if (v instanceof Date) { changed = true; return normHHmm_(v); }
    return v;
  }));
  if (changed) range.setValues(fixed);
}

// GASエディタから手動実行して、既存のシフト設定シートを即座に修復する（引数なし）。
function debugFixShiftTimes() {
  const sheet = getOrCreateShiftSheet(); // 呼ぶだけで書式修正＋自己修復が走る
  return 'シフト設定シートの時刻データを修復しました（行数: ' + Math.max(0, sheet.getLastRow() - 1) + '）';
}

// D/E列から読んだ値をHH:mm文字列に正規化する（Date型で入っていても文字列でもOK）。
function normHHmm_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Etc/UTC', 'HH:mm');
  return String(v || '');
}

// HH:mm形式チェック（簡易）
function isValidHHmm_(s) {
  return typeof s === 'string' && /^([01]\d|2[0-3]):[0-5]\d$/.test(s);
}

// 社員：自分の当月のシフト希望・確定を取得
function getMyShift(userId, year, month) {
  const emp = getEmployeeByLineId(userId);
  if (!emp) return { success: false, message: '登録されていないアカウントです。' };

  const sheet = getOrCreateShiftSheet();
  const ym = year + '/' + String(month).padStart(2, '0');
  const kibou = [], confirmed = []; // [{d, start, end}]
  sheet.getDataRange().getValues().slice(1).forEach(r => {
    if (r[0] !== emp.id || !r[1]) return;
    const d = fmt(new Date(r[1]), 'yyyy/MM/dd');
    if (!d.startsWith(ym)) return;
    const item = { d: d, start: normHHmm_(r[3]), end: normHHmm_(r[4]) };
    if ((r[5] || '確定') === '希望') kibou.push(item); else confirmed.push(item);
  });
  const daysInMonth = new Date(year, month, 0).getDate();
  return { success: true, name: emp.name, year, month, daysInMonth, kibou, confirmed };
}

// 社員：出勤希望（シフト）を申告。entries=[{d,start,end}]（明日以降のみ）。
// 当月・翌月ともOK。同月内の「明日以降の希望」を置き換える（確定済みは変更しない＝管理者確定分は上書きしない）。
function submitShiftKibou(userId, entries, year, month) {
  const emp = getEmployeeByLineId(userId);
  if (!emp) return { success: false, message: '登録されていないアカウントです。' };

  const list = (entries || []).filter(x => x && x.d && x.start && x.end);
  const badTime = list.filter(x => !isValidHHmm_(x.start) || !isValidHHmm_(x.end) || x.start >= x.end);
  if (badTime.length) {
    return { success: false, message: '開始・終了時刻の形式が正しくないか、開始が終了以降になっています（' + badTime.map(x => x.d).join('、') + '）' };
  }

  const todayStr = fmt(new Date(), 'yyyy/MM/dd');
  const past = list.filter(x => x.d <= todayStr).map(x => x.d);
  if (past.length) {
    return { success: false, message: '当日・過去の日付はシフト希望に出せません（' + past.join('、') + '）。当日の急な変更は管理者にご連絡ください。' };
  }

  const sheet = getOrCreateShiftSheet();
  const ym = year + '/' + String(month).padStart(2, '0');

  // 当該社員・当月の「明日以降の希望」行のみ削除（確定済みは残す）
  const rows = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    const st = rows[i][5] || '確定';
    if (rows[i][0] === emp.id && rows[i][1] &&
        fmt(new Date(rows[i][1]), 'yyyy/MM/dd').startsWith(ym) &&
        fmt(new Date(rows[i][1]), 'yyyy/MM/dd') > todayStr &&
        st === '希望') {
      sheet.deleteRow(i + 1);
    }
  }

  const now = fmt(new Date(), 'yyyy/MM/dd HH:mm');
  list.forEach(x => sheet.appendRow([emp.id, x.d, emp.name, x.start, x.end, '希望', now]));

  const admins = getAdmins();
  admins.forEach(a => {
    if (a.lineUserId) sendMessage(a.lineUserId,
      '📅 シフト希望が届きました\n\n社員：' + emp.name + '\n対象：' + year + '年' + month + '月\n件数：' + list.length + '日\n\n管理者ダッシュボードの「シフト」タブで確認・確定してください。');
  });

  return { success: true, count: list.length };
}

// 管理者用：当月の社員一覧＋各社員のシフト希望・確定を取得
function getShiftForAdmin(adminUserId, year, month) {
  const admin = getEmployeeByLineId(adminUserId);
  if (!admin || !admin.isAdmin) return { success: false, message: '管理者権限がありません' };

  const ss = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID);
  const empRows = ss.getSheetByName(SHEETS.EMPLOYEES).getDataRange().getValues();
  const employees = empRows.slice(1)
    .filter(r => (r[7] !== false && r[7] !== 'FALSE') && !(r[6] === true || r[6] === 'TRUE'))
    .map(r => ({ id: r[0], name: r[1] }));

  const ym = year + '/' + String(month).padStart(2, '0');
  const sheet = getOrCreateShiftSheet();
  const kibou = {}, confirmed = {}; // empId -> [{d,start,end}]
  sheet.getDataRange().getValues().slice(1).forEach(r => {
    if (!r[0] || !r[1]) return;
    const d = fmt(new Date(r[1]), 'yyyy/MM/dd');
    if (!d.startsWith(ym)) return;
    const item = { d: d, start: normHHmm_(r[3]), end: normHHmm_(r[4]) };
    if ((r[5] || '確定') === '希望') (kibou[r[0]] = kibou[r[0]] || []).push(item);
    else (confirmed[r[0]] = confirmed[r[0]] || []).push(item);
  });

  const daysInMonth = new Date(year, month, 0).getDate();
  return { success: true, year, month, daysInMonth, employees, kibou, confirmed };
}

// 管理者用：ある社員の当月のシフトをまとめて保存（既存の確定・希望を置き換え）。entries=[{d,start,end}]
function saveShift(adminUserId, employeeId, entries, year, month) {
  const admin = getEmployeeByLineId(adminUserId);
  if (!admin || !admin.isAdmin) return { success: false, message: '管理者権限がありません' };

  const list = (entries || []).filter(x => x && x.d && x.start && x.end);
  const badTime = list.filter(x => !isValidHHmm_(x.start) || !isValidHHmm_(x.end) || x.start >= x.end);
  if (badTime.length) {
    return { success: false, message: '開始・終了時刻が不正です（' + badTime.map(x => x.d).join('、') + '）' };
  }

  const sheet = getOrCreateShiftSheet();
  const emp = getEmployeeById(employeeId);
  const ym  = year + '/' + String(month).padStart(2, '0');

  const rows = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (rows[i][0] === employeeId && rows[i][1] &&
        fmt(new Date(rows[i][1]), 'yyyy/MM/dd').startsWith(ym)) {
      sheet.deleteRow(i + 1);
    }
  }

  const now = fmt(new Date(), 'yyyy/MM/dd HH:mm');
  list.forEach(x => sheet.appendRow([employeeId, x.d, emp ? emp.name : '', x.start, x.end, '確定', now]));

  return { success: true, count: list.length };
}

// 管理者：来月のシフト希望を全員に個別LINEで依頼
function requestShiftFromAll(adminUserId, year, month) {
  const admin = getEmployeeByLineId(adminUserId);
  if (!admin || !admin.isAdmin) return { success: false, message: '管理者権限がありません' };

  const liffId = getConfig().LIFF_SHIFT_ID;
  const link = liffId ? ('https://liff.line.me/' + liffId) : '';

  const ss = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID);
  const empRows = ss.getSheetByName(SHEETS.EMPLOYEES).getDataRange().getValues();
  const targets = empRows.slice(1)
    .filter(r => (r[7] !== false && r[7] !== 'FALSE') && !(r[6] === true || r[6] === 'TRUE'))
    .filter(r => r[5]);

  let sent = 0;
  targets.forEach(r => {
    sendMessage(r[5],
      '📅 ' + year + '年' + month + '月のシフト希望をお願いします\n\n' +
      '出勤できる日と時間帯を選んで送信してください。\n' + link);
    sent++;
  });
  return { success: true, sent };
}

// 指定月の確定シフトを empId -> { date -> {start,end} } の形で一括取得（ループ内での逐次シート読みを避ける）
function getConfirmedShiftMapForMonth_(ss, ym) {
  const map = {};
  const sheet = ss.getSheetByName(SHEETS.SHIFT);
  if (!sheet) return map;
  sheet.getDataRange().getValues().slice(1).forEach(r => {
    if (!r[0] || !r[1] || (r[5] || '確定') !== '確定') return;
    const d = fmt(new Date(r[1]), 'yyyy/MM/dd');
    if (!d.startsWith(ym)) return;
    (map[r[0]] = map[r[0]] || {})[d] = { start: normHHmm_(r[3]), end: normHHmm_(r[4]) };
  });
  return map;
}
