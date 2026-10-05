/**
 * 丙級練習教室登記系統 V3
 * Google Apps Script 後端
 *
 * 工作表名稱：報名資料
 * 欄位：
 * A 報名編號
 * B 登記時間
 * C 項目代碼
 * D 項目
 * E 日期
 * F 姓名
 * G 手機
 * H 備註
 * I 費用
 * J 報名狀態
 * K 付款狀態
 * L 管理備註
 * M 最後更新
 */

const CONFIG = {
  SHEET_NAME: '報名資料',
  MAX_PEOPLE: 6,
  PRICE: 1500,
  COURSES: {
    cake: { title: '蛋糕丙級', date: '10/21（三）' },
    cook: { title: '中餐丙級', date: '11/4（三）' }
  }
};

const HEADERS = [
  '報名編號','登記時間','項目代碼','項目','日期',
  '姓名','手機','備註','費用','報名狀態',
  '付款狀態','管理備註','最後更新'
];

function doGet(e) {
  try {
    ensureSheet_();

    const action = (e && e.parameter && e.parameter.action) || 'publicStatus';

    if (action === 'publicStatus') {
      return json_({
        success: true,
        max: CONFIG.MAX_PEOPLE,
        price: CONFIG.PRICE,
        courses: getPublicStatus_()
      });
    }

    return json_({ success:false, message:'不支援的操作' });
  } catch (err) {
    return json_({ success:false, message:String(err.message || err) });
  }
}

function doPost(e) {
  try {
    ensureSheet_();
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const action = data.action || 'register';

    if (action === 'register') return register_(data);

    // 以下均為管理員功能，密碼只在伺服器端驗證
    if (!checkAdminPassword_(data.adminPassword)) {
      return json_({ success:false, message:'管理員密碼錯誤' });
    }

    if (action === 'adminList') return adminList_();
    if (action === 'updateBooking') return updateBooking_(data);
    if (action === 'cancelBooking') return cancelBooking_(data);

    return json_({ success:false, message:'不支援的操作' });
  } catch (err) {
    return json_({ success:false, message:String(err.message || err) });
  }
}

function register_(data) {
  const courseId = String(data.courseId || '').trim();
  const name = String(data.name || '').trim();
  const phone = normalizePhone_(data.phone);
  const note = String(data.note || '').trim();

  if (!CONFIG.COURSES[courseId]) return json_({success:false,message:'請選擇正確的練習項目'});
  if (!name) return json_({success:false,message:'請填寫姓名'});
  if (!/^09\d{8}$/.test(phone)) return json_({success:false,message:'手機格式不正確，請輸入 09xxxxxxxx'});
  if (name.length > 30) return json_({success:false,message:'姓名過長'});
  if (note.length > 300) return json_({success:false,message:'備註最多 300 字'});

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const sheet = getSheet_();
    const rows = getRows_();

    const duplicate = rows.some(r =>
      r.courseId === courseId &&
      normalizePhone_(r.phone) === phone &&
      !['已取消'].includes(r.bookingStatus)
    );

    if (duplicate) {
      return json_({success:false,message:'這支手機已經登記過同一場練習'});
    }

    const activeCount = rows.filter(r =>
      r.courseId === courseId &&
      ['待確認','已確認'].includes(r.bookingStatus)
    ).length;

    const bookingStatus = activeCount < CONFIG.MAX_PEOPLE ? '待確認' : '候補';
    const bookingId = makeBookingId_(courseId);
    const now = new Date();
    const c = CONFIG.COURSES[courseId];

    sheet.appendRow([
      bookingId,
      now,
      courseId,
      c.title,
      c.date,
      name,
      phone,
      note,
      CONFIG.PRICE,
      bookingStatus,
      '未付款',
      '',
      now
    ]);

    SpreadsheetApp.flush();

    const current = getRows_().filter(r =>
      r.courseId === courseId &&
      ['待確認','已確認'].includes(r.bookingStatus)
    ).length;

    const waitlistPosition = bookingStatus === '候補'
      ? getRows_().filter(r => r.courseId === courseId && r.bookingStatus === '候補').length
      : 0;

    return json_({
      success:true,
      message: bookingStatus === '候補' ? '本場已額滿，已加入候補名單' : '登記成功',
      bookingId,
      bookingStatus,
      waitlistPosition,
      count: current,
      remaining: Math.max(0, CONFIG.MAX_PEOPLE - current)
    });
  } finally {
    lock.releaseLock();
  }
}

function adminList_() {
  const rows = getRows_();
  return json_({
    success:true,
    max:CONFIG.MAX_PEOPLE,
    price:CONFIG.PRICE,
    courses:getPublicStatus_(),
    bookings:rows
  });
}

function updateBooking_(data) {
  const bookingId = String(data.bookingId || '').trim();
  if (!bookingId) return json_({success:false,message:'缺少報名編號'});

  const allowedBooking = ['待確認','已確認','候補','已取消'];
  const allowedPayment = ['未付款','已付款','退款'];

  const bookingStatus = String(data.bookingStatus || '').trim();
  const paymentStatus = String(data.paymentStatus || '').trim();
  const adminNote = String(data.adminNote || '').trim();

  if (bookingStatus && !allowedBooking.includes(bookingStatus)) {
    return json_({success:false,message:'報名狀態不正確'});
  }
  if (paymentStatus && !allowedPayment.includes(paymentStatus)) {
    return json_({success:false,message:'付款狀態不正確'});
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const sheet = getSheet_();
    const rows = getRows_();
    const target = rows.find(r => r.bookingId === bookingId);
    if (!target) return json_({success:false,message:'找不到這筆報名'});

    // 若要把候補/取消改成正式名額，先檢查是否還有空位
    const nextStatus = bookingStatus || target.bookingStatus;
    const becomingActive =
      ['待確認','已確認'].includes(nextStatus) &&
      !['待確認','已確認'].includes(target.bookingStatus);

    if (becomingActive) {
      const activeCount = rows.filter(r =>
        r.courseId === target.courseId &&
        r.bookingId !== bookingId &&
        ['待確認','已確認'].includes(r.bookingStatus)
      ).length;

      if (activeCount >= CONFIG.MAX_PEOPLE) {
        return json_({success:false,message:'本場目前已滿 6 人，無法改為正式名額'});
      }
    }

    const row = target.rowNumber;
    if (bookingStatus) sheet.getRange(row, 10).setValue(bookingStatus);
    if (paymentStatus) sheet.getRange(row, 11).setValue(paymentStatus);
    sheet.getRange(row, 12).setValue(adminNote);
    sheet.getRange(row, 13).setValue(new Date());

    // 如果這次是取消正式名額，嘗試自動遞補第一位候補
    let promoted = null;
    if (
      bookingStatus === '已取消' &&
      ['待確認','已確認'].includes(target.bookingStatus)
    ) {
      promoted = promoteFirstWaitlist_(target.courseId);
    }

    SpreadsheetApp.flush();

    return json_({
      success:true,
      message:'更新成功',
      promoted
    });
  } finally {
    lock.releaseLock();
  }
}

function cancelBooking_(data) {
  data.bookingStatus = '已取消';
  return updateBooking_(data);
}

function promoteFirstWaitlist_(courseId) {
  const sheet = getSheet_();
  const rows = getRows_();

  const activeCount = rows.filter(r =>
    r.courseId === courseId &&
    ['待確認','已確認'].includes(r.bookingStatus)
  ).length;

  if (activeCount >= CONFIG.MAX_PEOPLE) return null;

  const waiting = rows
    .filter(r => r.courseId === courseId && r.bookingStatus === '候補')
    .sort((a,b) => new Date(a.createdAt) - new Date(b.createdAt));

  if (!waiting.length) return null;

  const first = waiting[0];
  sheet.getRange(first.rowNumber, 10).setValue('待確認');
  sheet.getRange(first.rowNumber, 12).setValue(
    (first.adminNote ? first.adminNote + '；' : '') + '系統自動由候補遞補'
  );
  sheet.getRange(first.rowNumber, 13).setValue(new Date());

  return {
    bookingId:first.bookingId,
    name:first.name,
    phone:first.phone
  };
}

function getPublicStatus_() {
  const rows = getRows_();
  const result = {};

  Object.keys(CONFIG.COURSES).forEach(courseId => {
    const c = CONFIG.COURSES[courseId];
    const active = rows.filter(r =>
      r.courseId === courseId &&
      ['待確認','已確認'].includes(r.bookingStatus)
    ).length;
    const waitlist = rows.filter(r =>
      r.courseId === courseId &&
      r.bookingStatus === '候補'
    ).length;

    result[courseId] = {
      title:c.title,
      date:c.date,
      count:active,
      remaining:Math.max(0, CONFIG.MAX_PEOPLE - active),
      waitlist,
      full:active >= CONFIG.MAX_PEOPLE
    };
  });

  return result;
}

function ensureSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
  }

  const current = sheet.getRange(1,1,1,HEADERS.length).getValues()[0];
  const needsHeader = HEADERS.some((h,i) => current[i] !== h);

  if (needsHeader && sheet.getLastRow() <= 1) {
    sheet.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1,1,1,HEADERS.length).setFontWeight('bold');
    sheet.autoResizeColumns(1, HEADERS.length);
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
  }

  return sheet;
}

function getSheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
}

function getRows_() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2,1,lastRow-1,HEADERS.length).getValues();

  return values
    .map((r,i) => ({
      rowNumber:i+2,
      bookingId:String(r[0] || ''),
      createdAt:r[1] || '',
      courseId:String(r[2] || ''),
      courseTitle:String(r[3] || ''),
      date:String(r[4] || ''),
      name:String(r[5] || ''),
      phone:String(r[6] || ''),
      note:String(r[7] || ''),
      price:Number(r[8] || CONFIG.PRICE),
      bookingStatus:String(r[9] || ''),
      paymentStatus:String(r[10] || ''),
      adminNote:String(r[11] || ''),
      updatedAt:r[12] || ''
    }))
    .filter(r => r.bookingId);
}

function checkAdminPassword_(password) {
  const saved = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if (!saved) throw new Error('尚未設定 ADMIN_PASSWORD，請先到指令碼屬性設定管理員密碼');
  return String(password || '') === saved;
}

function normalizePhone_(value) {
  return String(value || '').replace(/\D/g,'');
}

function makeBookingId_(courseId) {
  const tz = Session.getScriptTimeZone() || 'Asia/Taipei';
  const stamp = Utilities.formatDate(new Date(), tz, 'yyyyMMddHHmmss');
  const rnd = Math.floor(100 + Math.random()*900);
  return `${courseId.toUpperCase()}-${stamp}-${rnd}`;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
