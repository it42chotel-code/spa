/**
 * 42C Spa Queue V2
 * Google Apps Script backend
 *
 * Staff enter data only in the Queue sheet.
 * Public API returns only therapist names, photos, working state, status and times.
 */
const SPAQ = Object.freeze({
  VERSION: '2.0.3',
  TIMEZONE: 'Asia/Bangkok',
  QUEUE_SPREADSHEET_ID: '1fQ2ieIc0qBhrgx6LwlPQ53HASxvS3QBMs04gcwTfMb8',
  REGISTRY_SPREADSHEET_ID: '1UM-6JfkCp3DJPwaT3Zg1kRX85Psm1epY',
  QUEUE_SHEET: 'Queue',
  THERAPIST_SHEET: 'ตารางเวลาหมอนวด',
  CONFIG_SHEET: 'SPA_CONFIG',
  LOG_SHEET: 'SPA_SYSTEM_LOG',
  AUDIT_SHEET: 'SPA_EXPORT_AUDIT',
  QUEUE_HEADER_ROW: 1,
  REGISTRY_HEADER_ROW: 2,
  FIRST_DATA_ROW: 2,
  REGISTRY_FIRST_DATA_ROW: 3,
  START_HOUR: 14,
  END_HOUR: 23,
  SLOT_TIMES: ['14:00', '15:00', '16:00', '17:00', '18:00', '19:00', '20:00', '21:00', '22:00', '23:00'],
  STATUS: Object.freeze({
    AVAILABLE: 'ว่าง',
    BUSY: 'ไม่ว่าง',
    LAST: 'คิวสุดท้ายของวัน',
    OPEN: 'เปิดห้องสปา',
    CLOSED: 'ปิดห้องสปา'
  }),
  PROPERTY_ACTIVE_DATE: 'SPAQ_ACTIVE_BUSINESS_DATE',
  PROPERTY_PREFIX_RUN: 'SPAQ_RUN_'
});

const SPAQ_THAI_MONTHS = Object.freeze([
  'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
  'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'
]);

const SPAQ_QUEUE_HEADERS = Object.freeze({
  therapistId: ['Therapist_ID', 'Therapist ID'],
  timeSlot: ['Time_Slot', 'Time Slot'],
  status: ['Status', 'สถานะ'],
  customer: ['Customer_Note', 'Customer Note', 'ชื่อลูกค้า', 'ชื่อ'],
  room: ['ห้อง', 'Room'],
  phone: ['เบอร์โทร', 'โทรศัพท์', 'Phone'],
  course: ['คอร์สการนวด', 'คอร์ส', 'Course'],
  courseHours: ['ชั่วโมง', 'Hours'],
  price: ['ราคา', 'Price'],
  customerType: ['ประเภทลูกค้า', 'Customer Type'],
  note: ['หมายเหตุ', 'Note'],
  actualStart: ['เวลาเริ่มจริง', 'Actual Start'],
  actualEnd: ['เวลาสิ้นสุดจริง', 'Actual End'],
  salesperson: ['พนักงานขาย', 'Salesperson']
});

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('42C Spa System')
    .addItem('ติดตั้ง/ซ่อมแซมระบบ', 'installSpaQueueSystem')
    .addSeparator()
    .addItem('ส่งออกข้อมูลตอนนี้', 'manualExportNow')
    .addItem('ตรวจสอบและรีเซ็ตตอนนี้', 'manualVerifyAndResetNow')
    .addItem('ตรวจสอบระบบ', 'validateSpaQueueSystem')
    .addItem('จัดแถวหมอนวดให้ครบ', 'syncQueueLayout')
    .addToUi();
}

function doGet() {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get('SPAQ_PUBLIC_STATE_V2');
    if (cached) {
      return spaJson_(JSON.parse(cached));
    }

    const state = buildPublicState_();
    cache.put('SPAQ_PUBLIC_STATE_V2', JSON.stringify(state), 8);
    return spaJson_(state);
  } catch (error) {
    logEvent_('PUBLIC_API_ERROR', 'ERROR', error && error.stack ? error.stack : String(error));
    return spaJson_({
      ok: false,
      version: SPAQ.VERSION,
      generatedAt: spaNowIso_(),
      timezone: SPAQ.TIMEZONE,
      message: 'ระบบกำลังตรวจสอบข้อมูล / Checking data'
    });
  }
}

function spaJson_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function installSpaQueueSystem() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const queueSS = SpreadsheetApp.openById(SPAQ.QUEUE_SPREADSHEET_ID);
    const registrySS = SpreadsheetApp.openById(SPAQ.REGISTRY_SPREADSHEET_ID);
    queueSS.setSpreadsheetTimeZone(SPAQ.TIMEZONE);
    registrySS.setSpreadsheetTimeZone(SPAQ.TIMEZONE);

    validateRequiredSheets_(queueSS);
    ensureHelperSheets_(queueSS);
    syncTherapistConfig_(queueSS, registrySS);
    ensureQueueRows_(queueSS);
    installMinuteTrigger_();

    const props = PropertiesService.getScriptProperties();
    if (!props.getProperty(SPAQ.PROPERTY_ACTIVE_DATE)) {
      props.setProperty(SPAQ.PROPERTY_ACTIVE_DATE, initialBusinessDateKey_());
    }

    logEvent_('INSTALL', 'OK', 'Installed Spa Queue V2 ' + SPAQ.VERSION);
    SpreadsheetApp.getUi().alert(
      'ติดตั้งระบบ 42C Spa Queue V2 สำเร็จ\n\n' +
      'กรุณาตรวจชื่อเต็มของหมอนวดในแท็บ SPA_CONFIG และ Deploy เป็น Web app'
    );
  } finally {
    lock.releaseLock();
  }
}

function installMinuteTrigger_() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'scheduledDispatcher') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('scheduledDispatcher')
    .timeBased()
    .everyMinutes(1)
    .create();
}

function scheduledDispatcher() {
  const now = new Date();
  const dateKey = Utilities.formatDate(now, SPAQ.TIMEZONE, 'yyyy-MM-dd');
  const minutes = Number(Utilities.formatDate(now, SPAQ.TIMEZONE, 'H')) * 60 +
    Number(Utilities.formatDate(now, SPAQ.TIMEZONE, 'm'));

  if (minutes >= 540 && minutes < 570) {
    runScheduledOnce_('EXPORT_0900', dateKey, function() {
      return runExport_('09:00');
    });
  }

  if (minutes >= 570 && minutes < 600) {
    runScheduledOnce_('EXPORT_0930', dateKey, function() {
      return runExport_('09:30');
    });
  }

  if (minutes >= 600 && minutes < 720) {
    runResetWithRetryWindow_(dateKey);
  }
}

function runScheduledOnce_(taskName, dateKey, callback) {
  const props = PropertiesService.getScriptProperties();
  const key = SPAQ.PROPERTY_PREFIX_RUN + taskName + '_' + dateKey;
  if (props.getProperty(key)) return;

  try {
    const result = callback();
    props.setProperty(key, JSON.stringify({
      at: spaNowIso_(),
      ok: Boolean(result && result.ok)
    }));
  } catch (error) {
    logEvent_(taskName, 'ERROR', error && error.stack ? error.stack : String(error));
    sendAlert_(taskName + ' ทำงานไม่สำเร็จ', String(error));
  }
}

function runResetWithRetryWindow_(dateKey) {
  const props = PropertiesService.getScriptProperties();
  const doneKey = SPAQ.PROPERTY_PREFIX_RUN + 'RESET_1000_' + dateKey;
  if (props.getProperty(doneKey)) return;

  const throttleKey = SPAQ.PROPERTY_PREFIX_RUN + 'RESET_ATTEMPT_' + dateKey;
  const lastAttempt = Number(props.getProperty(throttleKey) || 0);
  if (Date.now() - lastAttempt < 5 * 60 * 1000) return;
  props.setProperty(throttleKey, String(Date.now()));

  try {
    const result = verifyAndReset_('10:00');
    if (result.ok && result.reset) {
      props.setProperty(doneKey, JSON.stringify({ at: spaNowIso_(), ok: true }));
    }
  } catch (error) {
    logEvent_('RESET_1000', 'ERROR', error && error.stack ? error.stack : String(error));
    sendAlert_('Reset 10:00 ทำงานไม่สำเร็จ', String(error));
  }
}

function manualExportNow() {
  const result = runExport_('MANUAL');
  SpreadsheetApp.getUi().alert(JSON.stringify(result, null, 2));
}

function manualVerifyAndResetNow() {
  const result = verifyAndReset_('MANUAL');
  SpreadsheetApp.getUi().alert(JSON.stringify(result, null, 2));
}

function validateSpaQueueSystem() {
  const queueSS = SpreadsheetApp.openById(SPAQ.QUEUE_SPREADSHEET_ID);
  validateRequiredSheets_(queueSS);
  const result = inspectSource_();
  SpreadsheetApp.getUi().alert(JSON.stringify({
    ok: result.invalid.length === 0,
    businessDate: result.businessDateKey,
    bookingCount: result.bookings.length,
    invalid: result.invalid
  }, null, 2));
}

function syncQueueLayout() {
  const queueSS = SpreadsheetApp.openById(SPAQ.QUEUE_SPREADSHEET_ID);
  ensureQueueRows_(queueSS);
  SpreadsheetApp.getUi().alert('จัดแถวหมอนวดและช่วงเวลาเรียบร้อยแล้ว');
}

function runExport_(runLabel) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const source = inspectSource_();
    if (source.invalid.length) {
      const detail = source.invalid.map(function(item) {
        return 'แถว ' + item.row + ': ' + item.errors.join(', ');
      }).join(' | ');
      logEvent_('EXPORT_' + runLabel, 'BLOCKED', detail);
      sendAlert_('Export ถูกระงับ', detail);
      return {
        ok: false,
        exported: 0,
        updated: 0,
        invalid: source.invalid
      };
    }

    if (!source.bookings.length) {
      writeAudit_([{
        businessDate: source.businessDateKey,
        bookingId: '',
        sourceRow: '',
        action: 'NO_BOOKINGS',
        targetSheet: '',
        targetRow: '',
        details: runLabel
      }]);
      logEvent_('EXPORT_' + runLabel, 'OK', 'No bookings');
      return { ok: true, exported: 0, updated: 0, total: 0 };
    }

    const registrySS = SpreadsheetApp.openById(SPAQ.REGISTRY_SPREADSHEET_ID);
    const businessDate = dateFromKey_(source.businessDateKey);
    const targetSheet = getOrCreateMonthlyRegistrySheet_(registrySS, businessDate);
    const existing = getRegistryIdRows_(targetSheet);
    let nextNumber = getNextRegistryNumber_(targetSheet);
    let exported = 0;
    let updated = 0;
    const auditRows = [];

    source.bookings.forEach(function(booking) {
      let targetRow = existing[booking.bookingId] || 0;
      const existed = Boolean(targetRow);
      let sequence;

      if (targetRow) {
        sequence = targetSheet.getRange(targetRow, 1).getValue() || nextNumber++;
        updated++;
      } else {
        targetRow = Math.max(targetSheet.getLastRow() + 1, SPAQ.REGISTRY_FIRST_DATA_ROW);
        sequence = nextNumber++;
        exported++;
      }

      const row = [[
        sequence,
        dateFromKey_(booking.businessDate),
        booking.customer,
        booking.room,
        booking.phone,
        booking.course,
        formatRegistryTimeRange_(booking.startMinutes, booking.endMinutes),
        booking.durationHours,
        booking.registryTherapistName,
        booking.price,
        booking.salesperson,
        combineRegistryNote_(booking.customerType, booking.note),
        booking.bookingId
      ]];

      targetSheet.getRange(targetRow, 1, 1, row[0].length).setValues(row);
      targetSheet.getRange(targetRow, 2).setNumberFormat('d/m/yyyy');
      targetSheet.getRange(targetRow, 7).setNumberFormat('@');
      targetSheet.getRange(targetRow, 8).setNumberFormat('0.##');
      targetSheet.getRange(targetRow, 13).setNumberFormat('@');
      existing[booking.bookingId] = targetRow;

      auditRows.push({
        businessDate: booking.businessDate,
        bookingId: booking.bookingId,
        sourceRow: booking.sourceRow,
        action: existed ? 'UPDATE' : 'INSERT',
        targetSheet: targetSheet.getName(),
        targetRow: targetRow,
        details: runLabel
      });
    });

    ensureRegistryTechnicalColumn_(targetSheet);
    writeAudit_(auditRows);
    SpreadsheetApp.flush();

    const verification = verifyExport_(source, registrySS);
    const ok = verification.missing.length === 0;
    logEvent_('EXPORT_' + runLabel, ok ? 'OK' : 'INCOMPLETE',
      'total=' + source.bookings.length +
      ', exported=' + exported +
      ', updated=' + updated +
      ', missing=' + verification.missing.length);

    if (!ok) {
      sendAlert_('Export ข้อมูลไม่ครบ', 'Booking IDs ที่ขาด: ' + verification.missing.join(', '));
    }

    return {
      ok: ok,
      exported: exported,
      updated: updated,
      total: source.bookings.length,
      missing: verification.missing
    };
  } finally {
    lock.releaseLock();
  }
}

function verifyAndReset_(runLabel) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    let source = inspectSource_();

    if (source.invalid.length) {
      const invalidDetail = source.invalid.map(function(item) {
        return 'แถว ' + item.row + ': ' + item.errors.join(', ');
      }).join(' | ');
      blockReset_(runLabel, 'ข้อมูลคิวไม่สมบูรณ์: ' + invalidDetail);
      return { ok: false, reset: false, reason: 'INVALID_SOURCE', invalid: source.invalid };
    }

    let verification = verifyExport_(source);
    if (verification.missing.length) {
      logEvent_('RESET_' + runLabel, 'RETRY_EXPORT',
        'Missing before reset: ' + verification.missing.join(', '));

      // Release the reset lock before the export routine acquires its own lock.
      // This prevents a nested-lock timeout while retaining a single writer.
      lock.releaseLock();
      try {
        runExport_('RESET_RETRY');
      } finally {
        lock.waitLock(30000);
      }

      source = inspectSource_();
      verification = verifyExport_(source);
    }

    if (verification.missing.length) {
      blockReset_(runLabel, 'พบ Booking ตกหล่น: ' + verification.missing.join(', '));
      return {
        ok: false,
        reset: false,
        reason: 'MISSING_EXPORT',
        missing: verification.missing
      };
    }

    const queueSS = SpreadsheetApp.openById(SPAQ.QUEUE_SPREADSHEET_ID);
    createQueueBackup_(queueSS, source.businessDateKey);
    resetQueueData_(queueSS);
    PropertiesService.getScriptProperties()
      .setProperty(SPAQ.PROPERTY_ACTIVE_DATE, Utilities.formatDate(new Date(), SPAQ.TIMEZONE, 'yyyy-MM-dd'));
    CacheService.getScriptCache().remove('SPAQ_PUBLIC_STATE_V2');

    writeAudit_([{
      businessDate: source.businessDateKey,
      bookingId: '',
      sourceRow: '',
      action: 'RESET_OK',
      targetSheet: SPAQ.QUEUE_SHEET,
      targetRow: '',
      details: runLabel + '; verified=' + source.bookings.length
    }]);
    logEvent_('RESET_' + runLabel, 'OK',
      'Verified ' + source.bookings.length + ' bookings before reset');

    return {
      ok: true,
      reset: true,
      verified: source.bookings.length,
      businessDate: source.businessDateKey
    };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}
function blockReset_(runLabel, reason) {
  writeAudit_([{
    businessDate: getBusinessDateKey_(),
    bookingId: '',
    sourceRow: '',
    action: 'RESET_BLOCKED',
    targetSheet: SPAQ.QUEUE_SHEET,
    targetRow: '',
    details: reason
  }]);
  logEvent_('RESET_' + runLabel, 'BLOCKED', reason);
  sendAlert_('ห้าม Reset ตาราง Spa Queue', reason);
}

function inspectSource_() {
  const queueSS = SpreadsheetApp.openById(SPAQ.QUEUE_SPREADSHEET_ID);
  const registrySS = SpreadsheetApp.openById(SPAQ.REGISTRY_SPREADSHEET_ID);
  syncTherapistConfig_(queueSS, registrySS);

  const sheet = requireSheet_(queueSS, SPAQ.QUEUE_SHEET);
  const range = sheet.getDataRange();
  const values = range.getValues();
  const display = range.getDisplayValues();
  const headers = mapRequiredHeaders_(display[0], SPAQ_QUEUE_HEADERS);
  const therapists = readTherapists_(queueSS);
  const config = getTherapistConfig_(queueSS, therapists);
  const businessDateKey = getBusinessDateKey_();
  const bookings = [];
  const invalid = [];

  for (let r = 1; r < values.length; r++) {
    const therapistId = cleanText_(display[r][headers.therapistId]);
    if (!therapistId) continue;

    const raw = {
      sourceRow: r + 1,
      therapistId: therapistId,
      timeSlot: cleanText_(display[r][headers.timeSlot]),
      status: normalizeStatus_(display[r][headers.status]),
      customer: cleanText_(display[r][headers.customer]),
      room: cleanText_(display[r][headers.room]),
      phone: cleanText_(display[r][headers.phone]),
      course: cleanText_(display[r][headers.course]),
      courseHours: values[r][headers.courseHours],
      price: values[r][headers.price],
      customerType: cleanText_(display[r][headers.customerType]),
      note: cleanText_(display[r][headers.note]),
      actualStartValue: values[r][headers.actualStart],
      actualStartDisplay: display[r][headers.actualStart],
      actualEndValue: values[r][headers.actualEnd],
      actualEndDisplay: display[r][headers.actualEnd],
      salesperson: cleanText_(display[r][headers.salesperson])
    };

    const hasBookingDetails = Boolean(
      raw.customer || raw.room || raw.phone || raw.course ||
      raw.customerType || raw.note || raw.actualStartDisplay ||
      raw.actualEndDisplay || cleanText_(raw.price)
    );

    if (!hasBookingDetails) {
      continue;
    }

    const errors = [];
    if (raw.status === SPAQ.STATUS.AVAILABLE ||
        raw.status === SPAQ.STATUS.OPEN ||
        raw.status === SPAQ.STATUS.CLOSED) {
      errors.push('สถานะไม่ใช่คิวที่กำลังใช้งาน');
    }

    const therapistConfig = config[therapistId];
    if (!therapistConfig || !therapistConfig.registryName) {
      errors.push('ไม่พบชื่อหมอนวดสำหรับทะเบียนของ Therapist_ID ' + therapistId);
    }
    if (!raw.customer) errors.push('ไม่มีชื่อลูกค้า');
    if (!raw.course) errors.push('ไม่มีคอร์สการนวด');

    const startMinutes = parseTimeMinutes_(raw.actualStartValue, raw.actualStartDisplay);
    const endClockMinutes = parseTimeMinutes_(raw.actualEndValue, raw.actualEndDisplay);
    if (startMinutes === null) errors.push('ไม่มีหรืออ่านเวลาเริ่มจริงไม่ได้');
    if (endClockMinutes === null) errors.push('ไม่มีหรืออ่านเวลาสิ้นสุดจริงไม่ได้');

    let endMinutes = endClockMinutes;
    if (startMinutes !== null && endMinutes !== null && endMinutes <= startMinutes) {
      endMinutes += 24 * 60;
    }

    if (startMinutes !== null && endMinutes !== null) {
      const durationMinutes = endMinutes - startMinutes;
      if (durationMinutes <= 0 || durationMinutes > 12 * 60) {
        errors.push('ช่วงเวลาจริงไม่สมเหตุสมผล');
      }
      if (startMinutes < SPAQ.START_HOUR * 60 || startMinutes >= SPAQ.END_HOUR * 60) {
        errors.push('เวลาเริ่มจริงต้องอยู่ระหว่าง 14:00–ก่อน 23:00');
      }
      if (endMinutes > SPAQ.END_HOUR * 60) {
        errors.push('เวลาสิ้นสุดจริงเกินเวลาปิด 23:00');
      }
    }

    if (errors.length) {
      invalid.push({ row: raw.sourceRow, errors: errors });
      continue;
    }

    const booking = {
      businessDate: businessDateKey,
      sourceRow: raw.sourceRow,
      therapistId: therapistId,
      registryTherapistName: therapistConfig.registryName,
      customer: raw.customer,
      room: raw.room,
      phone: raw.phone,
      course: raw.course,
      price: raw.price,
      customerType: raw.customerType,
      note: raw.note,
      salesperson: raw.salesperson,
      startMinutes: startMinutes,
      endMinutes: endMinutes,
      durationHours: roundHours_((endMinutes - startMinutes) / 60),
      status: raw.status || SPAQ.STATUS.BUSY
    };
    booking.bookingId = buildBookingId_(booking);
    bookings.push(booking);
  }

  const unique = {};
  bookings.forEach(function(booking) {
    if (!unique[booking.bookingId]) {
      unique[booking.bookingId] = booking;
    }
  });

  return {
    businessDateKey: businessDateKey,
    bookings: Object.keys(unique).map(function(id) { return unique[id]; }),
    invalid: invalid,
    therapists: therapists
  };
}

function verifyExport_(source, registrySS) {
  registrySS = registrySS || SpreadsheetApp.openById(SPAQ.REGISTRY_SPREADSHEET_ID);
  if (!source.bookings.length) {
    return { missing: [], found: [] };
  }

  const businessDate = dateFromKey_(source.businessDateKey);
  const sheetName = thaiMonthSheetName_(businessDate);
  const sheet = registrySS.getSheetByName(sheetName);
  if (!sheet) {
    return {
      missing: source.bookings.map(function(item) { return item.bookingId; }),
      found: []
    };
  }

  const existing = getRegistryIdRows_(sheet);
  const missing = [];
  const found = [];
  source.bookings.forEach(function(booking) {
    if (existing[booking.bookingId]) found.push(booking.bookingId);
    else missing.push(booking.bookingId);
  });
  return { missing: missing, found: found };
}

function buildPublicState_() {
  const queueSS = SpreadsheetApp.openById(SPAQ.QUEUE_SPREADSHEET_ID);
  const sheet = requireSheet_(queueSS, SPAQ.QUEUE_SHEET);
  const range = sheet.getDataRange();
  const values = range.getValues();
  const display = range.getDisplayValues();
  const headers = mapRequiredHeaders_(display[0], SPAQ_QUEUE_HEADERS);
  const therapists = readTherapists_(queueSS);
  const now = new Date();
  const nowMinutes = Number(Utilities.formatDate(now, SPAQ.TIMEZONE, 'H')) * 60 +
    Number(Utilities.formatDate(now, SPAQ.TIMEZONE, 'm'));
  const dayKey = getBusinessDateKey_();

  const publicTherapists = therapists
    .map(function(therapist) {
      const items = [];
      let hasLastQueue = false;
      let lastQueueEnd = null;
      let explicitClosedAt = null;
      let explicitOpen = false;

      for (let r = 1; r < values.length; r++) {
        const rowId = cleanText_(display[r][headers.therapistId]);
        if (rowId !== therapist.id) continue;

        const status = normalizeStatus_(display[r][headers.status]);
        const slotMinutes = parseTimeMinutes_(values[r][headers.timeSlot], display[r][headers.timeSlot]);
        if (status === SPAQ.STATUS.CLOSED && slotMinutes !== null) {
          explicitClosedAt = explicitClosedAt === null ? slotMinutes : Math.min(explicitClosedAt, slotMinutes);
        }
        if (status === SPAQ.STATUS.OPEN) explicitOpen = true;

        const hasDetails = Boolean(
          cleanText_(display[r][headers.customer]) ||
          cleanText_(display[r][headers.course]) ||
          cleanText_(display[r][headers.actualStart]) ||
          cleanText_(display[r][headers.actualEnd])
        );
        if (!hasDetails) continue;

        const start = parseTimeMinutes_(values[r][headers.actualStart], display[r][headers.actualStart]);
        let end = parseTimeMinutes_(values[r][headers.actualEnd], display[r][headers.actualEnd]);
        if (start === null || end === null) continue;
        if (end <= start) end += 24 * 60;

        if (status === SPAQ.STATUS.LAST) {
          hasLastQueue = true;
          lastQueueEnd = Math.max(lastQueueEnd || 0, end);
        }

        items.push({
          start: minutesToClock_(start),
          end: minutesToClock_(end),
          startMinutes: start,
          endMinutes: end,
          status: status || SPAQ.STATUS.BUSY,
          statusEn: publicStatusEnglish_(status || SPAQ.STATUS.BUSY)
        });
      }

      let currentStatus = therapist.workingToday ? SPAQ.STATUS.AVAILABLE : SPAQ.STATUS.CLOSED;
      const adjustedNow = nowMinutes < SPAQ.START_HOUR * 60 ? nowMinutes + 24 * 60 : nowMinutes;
      const activeBooking = items.find(function(item) {
        return adjustedNow >= item.startMinutes && adjustedNow < item.endMinutes;
      });

      if (!therapist.workingToday) {
        currentStatus = SPAQ.STATUS.CLOSED;
      } else if ((explicitClosedAt !== null && adjustedNow >= explicitClosedAt) ||
          (hasLastQueue && lastQueueEnd !== null && adjustedNow >= lastQueueEnd)) {
        currentStatus = SPAQ.STATUS.CLOSED;
      } else if (activeBooking) {
        currentStatus = activeBooking.status === SPAQ.STATUS.LAST ?
          SPAQ.STATUS.LAST : SPAQ.STATUS.BUSY;
      } else if (explicitOpen) {
        currentStatus = SPAQ.STATUS.OPEN;
      } else if (nowMinutes < SPAQ.START_HOUR * 60 || nowMinutes >= SPAQ.END_HOUR * 60) {
        currentStatus = SPAQ.STATUS.CLOSED;
      }

      return {
        id: therapist.id,
        nameTh: therapist.nameTh,
        nameEn: therapist.nameEn,
        photoUrl: therapist.photoUrl,
        workingToday: therapist.workingToday,
        currentStatus: currentStatus,
        currentStatusEn: publicStatusEnglish_(currentStatus),
        bookings: items
      };
    });

  return {
    ok: true,
    version: SPAQ.VERSION,
    generatedAt: spaNowIso_(),
    timezone: SPAQ.TIMEZONE,
    businessDate: dayKey,
    schedule: {
      start: '14:00',
      end: '23:00',
      slots: SPAQ.SLOT_TIMES
    },
    spaStatus: publicTherapists.some(function(item) {
      return item.currentStatus !== SPAQ.STATUS.CLOSED;
    }) ? SPAQ.STATUS.OPEN : SPAQ.STATUS.CLOSED,
    therapists: publicTherapists
  };
}

function readTherapists_(queueSS) {
  const sheet = requireSheet_(queueSS, SPAQ.THERAPIST_SHEET);
  const display = sheet.getDataRange().getDisplayValues();
  if (display.length < 2) return [];

  const headers = mapHeaderNames_(display[0]);
  const idIndex = requireHeaderIndex_(headers, ['ID']);
  const thIndex = requireHeaderIndex_(headers, ['Name_TH', 'Name TH']);
  const enIndex = requireHeaderIndex_(headers, ['Name_EN', 'Name EN']);
  const photoIndex = requireHeaderIndex_(headers, ['Photo_URL', 'Photo URL']);
  const workingIndex = requireHeaderIndex_(headers, ['Working_Today', 'Working Today']);

  const result = [];
  for (let r = 1; r < display.length; r++) {
    const id = cleanText_(display[r][idIndex]);
    if (!id) continue;
    result.push({
      id: id,
      nameTh: cleanText_(display[r][thIndex]) || ('หมอนวด ' + id),
      nameEn: cleanText_(display[r][enIndex]) || ('Therapist ' + id),
      photoUrl: cleanText_(display[r][photoIndex]),
      workingToday: isWorkingValue_(display[r][workingIndex])
    });
  }
  return result;
}

function ensureQueueRows_(queueSS) {
  const sheet = requireSheet_(queueSS, SPAQ.QUEUE_SHEET);
  const range = sheet.getDataRange();
  const display = range.getDisplayValues();
  const headers = mapRequiredHeaders_(display[0], SPAQ_QUEUE_HEADERS);
  const therapists = readTherapists_(queueSS);
  const existing = {};

  for (let r = 1; r < display.length; r++) {
    const id = cleanText_(display[r][headers.therapistId]);
    const slot = normalizeClockText_(display[r][headers.timeSlot]);
    if (id && slot) existing[id + '|' + slot] = true;
  }

  const newRows = [];
  therapists.forEach(function(therapist) {
    SPAQ.SLOT_TIMES.forEach(function(slot) {
      const key = therapist.id + '|' + slot;
      if (!existing[key]) {
        const row = new Array(display[0].length).fill('');
        row[headers.therapistId] = therapist.id;
        row[headers.timeSlot] = slot;
        row[headers.status] = therapist.workingToday && slot !== '23:00' ?
          SPAQ.STATUS.AVAILABLE : SPAQ.STATUS.CLOSED;
        newRows.push(row);
      }
    });
  });

  if (newRows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length)
      .setValues(newRows);
  }
}

function resetQueueData_(queueSS) {
  ensureQueueRows_(queueSS);
  const sheet = requireSheet_(queueSS, SPAQ.QUEUE_SHEET);
  const range = sheet.getDataRange();
  const display = range.getDisplayValues();
  const headers = mapRequiredHeaders_(display[0], SPAQ_QUEUE_HEADERS);
  const therapists = readTherapists_(queueSS);
  const working = {};
  therapists.forEach(function(item) { working[item.id] = item.workingToday; });

  const rowCount = Math.max(sheet.getLastRow() - 1, 0);
  if (!rowCount) return;

  const statusValues = [];
  for (let r = 1; r < display.length; r++) {
    const id = cleanText_(display[r][headers.therapistId]);
    const slot = normalizeClockText_(display[r][headers.timeSlot]);
    const canOpen = Boolean(working[id]) && slot !== '23:00';
    statusValues.push([canOpen ? SPAQ.STATUS.AVAILABLE : SPAQ.STATUS.CLOSED]);
  }
  sheet.getRange(2, headers.status + 1, statusValues.length, 1).setValues(statusValues);

  const clearStart = headers.customer + 1;
  const clearEnd = Math.max(
    headers.room, headers.phone, headers.course, headers.courseHours,
    headers.price, headers.customerType, headers.note,
    headers.actualStart, headers.actualEnd, headers.salesperson
  ) + 1;
  sheet.getRange(2, clearStart, rowCount, clearEnd - clearStart + 1).clearContent();
}

function createQueueBackup_(queueSS, businessDateKey) {
  const source = requireSheet_(queueSS, SPAQ.QUEUE_SHEET);
  const stamp = Utilities.formatDate(new Date(), SPAQ.TIMEZONE, 'yyyyMMdd_HHmmss');
  const baseName = 'SPAQ_BACKUP_' + businessDateKey.replace(/-/g, '') + '_' + stamp;
  let name = baseName.substring(0, 99);
  let suffix = 1;
  while (queueSS.getSheetByName(name)) {
    name = (baseName.substring(0, 94) + '_' + suffix++).substring(0, 99);
  }
  const backup = source.copyTo(queueSS).setName(name);
  backup.hideSheet();
}

function ensureHelperSheets_(queueSS) {
  let config = queueSS.getSheetByName(SPAQ.CONFIG_SHEET);
  if (!config) config = queueSS.insertSheet(SPAQ.CONFIG_SHEET);
  config.getRange('A1:C1').setValues([['Key', 'Value', 'Description']]);
  config.getRange('D1:G1').setValues([[
    'Therapist_ID', 'Registry_Name', 'Active', 'Note'
  ]]);

  const defaults = [
    ['SYSTEM_VERSION', SPAQ.VERSION, 'เวอร์ชันระบบ'],
    ['TIMEZONE', SPAQ.TIMEZONE, 'เขตเวลาระบบ'],
    ['ALERT_EMAIL', '', 'อีเมลแจ้งเตือนเมื่อ Export/Reset มีปัญหา'],
    ['PUBLIC_REFRESH_SECONDS', 10, 'รอบรีเฟรชหน้า Display']
  ];
  defaults.forEach(function(row, index) {
    const targetRow = index + 2;
    if (!config.getRange(targetRow, 1).getValue()) {
      config.getRange(targetRow, 1, 1, 3).setValues([row]);
    }
  });
  config.setFrozenRows(1);

  let log = queueSS.getSheetByName(SPAQ.LOG_SHEET);
  if (!log) log = queueSS.insertSheet(SPAQ.LOG_SHEET);
  if (!log.getRange(1, 1).getValue()) {
    log.getRange(1, 1, 1, 5).setValues([[
      'Timestamp', 'Event', 'Result', 'Details', 'Version'
    ]]);
  }

  let audit = queueSS.getSheetByName(SPAQ.AUDIT_SHEET);
  if (!audit) audit = queueSS.insertSheet(SPAQ.AUDIT_SHEET);
  if (!audit.getRange(1, 1).getValue()) {
    audit.getRange(1, 1, 1, 8).setValues([[
      'Timestamp', 'Business_Date', 'Booking_ID', 'Source_Row',
      'Action', 'Target_Sheet', 'Target_Row', 'Details'
    ]]);
  }
  if (!audit.isSheetHidden()) audit.hideSheet();
}

function syncTherapistConfig_(queueSS, registrySS) {
  ensureHelperSheets_(queueSS);
  const sheet = queueSS.getSheetByName(SPAQ.CONFIG_SHEET);
  const therapists = readTherapists_(queueSS);
  const existing = {};

  if (sheet.getLastRow() >= 2) {
    const rows = sheet.getRange(2, 4, sheet.getLastRow() - 1, 4).getDisplayValues();
    rows.forEach(function(row, index) {
      const id = cleanText_(row[0]);
      if (id) {
        existing[id] = {
          row: index + 2,
          registryName: cleanText_(row[1]),
          active: row[2],
          note: row[3]
        };
      }
    });
  }

  therapists.forEach(function(therapist) {
    const current = existing[therapist.id];
    if (current) {
      if (!current.registryName) {
        const discovered = discoverRegistryName_(registrySS, therapist);
        sheet.getRange(current.row, 5).setValue(discovered || therapist.nameTh);
      }
      sheet.getRange(current.row, 6).setValue(therapist.workingToday ? 'TRUE' : 'FALSE');
      return;
    }

    const discovered = discoverRegistryName_(registrySS, therapist);
    sheet.appendRow([
      '', '', '',
      therapist.id,
      discovered || therapist.nameTh,
      therapist.workingToday ? 'TRUE' : 'FALSE',
      discovered ? 'ตรวจพบชื่อเดิมจากทะเบียนอัตโนมัติ' : 'กรุณาตรวจชื่อเต็มก่อนใช้งาน'
    ]);
  });
}

function getTherapistConfig_(queueSS, therapists) {
  const sheet = requireSheet_(queueSS, SPAQ.CONFIG_SHEET);
  const result = {};
  const fallback = {};
  therapists.forEach(function(item) { fallback[item.id] = item; });

  if (sheet.getLastRow() < 2) return result;
  const rows = sheet.getRange(2, 4, sheet.getLastRow() - 1, 4).getDisplayValues();
  rows.forEach(function(row) {
    const id = cleanText_(row[0]);
    if (!id) return;
    result[id] = {
      registryName: cleanText_(row[1]) || (fallback[id] ? fallback[id].nameTh : ''),
      active: isWorkingValue_(row[2]),
      note: cleanText_(row[3])
    };
  });
  return result;
}

function discoverRegistryName_(registrySS, therapist) {
  const sheets = registrySS.getSheets().slice().reverse();
  const nicknames = [therapist.nameTh, therapist.nameEn]
    .map(function(value) { return normalizeForMatch_(value); })
    .filter(Boolean);

  for (let s = 0; s < Math.min(sheets.length, 12); s++) {
    const sheet = sheets[s];
    if (sheet.getLastRow() < SPAQ.REGISTRY_FIRST_DATA_ROW) continue;
    const values = sheet.getRange(
      SPAQ.REGISTRY_FIRST_DATA_ROW,
      9,
      sheet.getLastRow() - SPAQ.REGISTRY_FIRST_DATA_ROW + 1,
      1
    ).getDisplayValues();

    for (let r = 0; r < values.length; r++) {
      const fullName = cleanText_(values[r][0]);
      if (!fullName) continue;
      const parentheses = fullName.match(/\(([^)]+)\)/g) || [];
      const aliases = parentheses.map(function(part) {
        return normalizeForMatch_(part.replace(/[()]/g, ''));
      });
      if (aliases.some(function(alias) { return nicknames.indexOf(alias) >= 0; })) {
        return fullName;
      }
    }
  }
  return '';
}

function getOrCreateMonthlyRegistrySheet_(registrySS, businessDate) {
  const name = thaiMonthSheetName_(businessDate);
  let sheet = registrySS.getSheetByName(name);
  if (sheet) {
    ensureRegistryTechnicalColumn_(sheet);
    updateRegistryMonthTitle_(sheet, name);
    return sheet;
  }

  const previous = findPreviousMonthlySheet_(registrySS, businessDate);
  if (previous) {
    sheet = previous.copyTo(registrySS);
    sheet.setName(name);
    if (sheet.getMaxRows() >= SPAQ.REGISTRY_FIRST_DATA_ROW) {
      sheet.getRange(
        SPAQ.REGISTRY_FIRST_DATA_ROW,
        1,
        sheet.getMaxRows() - SPAQ.REGISTRY_FIRST_DATA_ROW + 1,
        sheet.getMaxColumns()
      ).clearContent();
    }
  } else {
    sheet = registrySS.insertSheet(name);
    sheet.getRange(1, 1).setValue('ทะเบียนประวัติผู้รับบริการ 42C Spa');
    sheet.getRange(2, 1, 1, 13).setValues([[
      'ที่', 'วันที่', 'ชื่อ', 'ห้อง', 'เบอร์โทร', 'คอส',
      'เวลา', 'ชั่วโมง', 'พนักงานนวด', 'ราคา',
      'พนักงานขาย', 'หมายเหตุ', 'SPAQ_BOOKING_ID'
    ]]);
    sheet.setFrozenRows(2);
  }

  updateRegistryMonthTitle_(sheet, name);
  ensureRegistryTechnicalColumn_(sheet);
  return sheet;
}

function findPreviousMonthlySheet_(registrySS, businessDate) {
  for (let offset = 1; offset <= 24; offset++) {
    const date = new Date(
      businessDate.getFullYear(),
      businessDate.getMonth() - offset,
      1,
      12, 0, 0
    );
    const exact = registrySS.getSheetByName(thaiMonthSheetName_(date));
    if (exact) return exact;

    const legacy = registrySS.getSheetByName(SPAQ_THAI_MONTHS[date.getMonth()]);
    if (legacy) return legacy;
  }
  return null;
}

function updateRegistryMonthTitle_(sheet, name) {
  const width = Math.max(sheet.getLastColumn(), 13);
  const values = sheet.getRange(1, 1, 1, width).getDisplayValues()[0];
  let found = false;

  for (let c = 0; c < values.length; c++) {
    if (/เดือน/.test(values[c])) {
      sheet.getRange(1, c + 1).setValue('เดือน ' + name);
      found = true;
      break;
    }
  }

  if (!found) {
    sheet.getRange(1, 4).setValue('เดือน ' + name);
  }
}

function ensureRegistryTechnicalColumn_(sheet) {
  if (sheet.getMaxColumns() < 13) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), 13 - sheet.getMaxColumns());
  }
  sheet.getRange(SPAQ.REGISTRY_HEADER_ROW, 13).setValue('SPAQ_BOOKING_ID');
  sheet.hideColumns(13);
}

function getRegistryIdRows_(sheet) {
  ensureRegistryTechnicalColumn_(sheet);
  const result = {};
  if (sheet.getLastRow() < SPAQ.REGISTRY_FIRST_DATA_ROW) return result;

  const ids = sheet.getRange(
    SPAQ.REGISTRY_FIRST_DATA_ROW,
    13,
    sheet.getLastRow() - SPAQ.REGISTRY_FIRST_DATA_ROW + 1,
    1
  ).getDisplayValues();

  ids.forEach(function(row, index) {
    const id = cleanText_(row[0]);
    if (id) result[id] = index + SPAQ.REGISTRY_FIRST_DATA_ROW;
  });
  return result;
}

function getNextRegistryNumber_(sheet) {
  if (sheet.getLastRow() < SPAQ.REGISTRY_FIRST_DATA_ROW) return 1;
  const values = sheet.getRange(
    SPAQ.REGISTRY_FIRST_DATA_ROW,
    1,
    sheet.getLastRow() - SPAQ.REGISTRY_FIRST_DATA_ROW + 1,
    1
  ).getValues();

  let max = 0;
  values.forEach(function(row) {
    const number = Number(row[0]);
    if (Number.isFinite(number)) max = Math.max(max, number);
  });
  return max + 1;
}

function writeAudit_(items) {
  if (!items || !items.length) return;
  const queueSS = SpreadsheetApp.openById(SPAQ.QUEUE_SPREADSHEET_ID);
  ensureHelperSheets_(queueSS);
  const sheet = queueSS.getSheetByName(SPAQ.AUDIT_SHEET);
  const now = new Date();
  const rows = items.map(function(item) {
    return [
      now,
      item.businessDate || '',
      item.bookingId || '',
      item.sourceRow || '',
      item.action || '',
      item.targetSheet || '',
      item.targetRow || '',
      item.details || ''
    ];
  });
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  sheet.getRange(sheet.getLastRow() - rows.length + 1, 1, rows.length, 1)
    .setNumberFormat('yyyy-mm-dd hh:mm:ss');
}

function logEvent_(eventName, result, details) {
  try {
    const queueSS = SpreadsheetApp.openById(SPAQ.QUEUE_SPREADSHEET_ID);
    ensureHelperSheets_(queueSS);
    const sheet = queueSS.getSheetByName(SPAQ.LOG_SHEET);
    sheet.appendRow([new Date(), eventName, result, details || '', SPAQ.VERSION]);
  } catch (ignored) {
    console.error(eventName + ': ' + details);
  }
}

function sendAlert_(subject, body) {
  try {
    const queueSS = SpreadsheetApp.openById(SPAQ.QUEUE_SPREADSHEET_ID);
    const config = queueSS.getSheetByName(SPAQ.CONFIG_SHEET);
    if (!config || config.getLastRow() < 2) return;
    const rows = config.getRange(2, 1, config.getLastRow() - 1, 2).getDisplayValues();
    let email = '';
    rows.forEach(function(row) {
      if (cleanText_(row[0]) === 'ALERT_EMAIL') email = cleanText_(row[1]);
    });
    if (!email) return;
    MailApp.sendEmail(
      email,
      '[42C Spa Queue] ' + subject,
      body + '\n\nเวลา: ' + spaNowIso_()
    );
  } catch (error) {
    console.error('Alert failed: ' + error);
  }
}

function validateRequiredSheets_(queueSS) {
  const queue = requireSheet_(queueSS, SPAQ.QUEUE_SHEET);
  requireSheet_(queueSS, SPAQ.THERAPIST_SHEET);
  const headers = queue.getRange(
    SPAQ.QUEUE_HEADER_ROW,
    1,
    1,
    queue.getLastColumn()
  ).getDisplayValues()[0];
  mapRequiredHeaders_(headers, SPAQ_QUEUE_HEADERS);
}

function mapRequiredHeaders_(headerRow, definitions) {
  const normalized = mapHeaderNames_(headerRow);
  const result = {};
  Object.keys(definitions).forEach(function(key) {
    result[key] = requireHeaderIndex_(normalized, definitions[key]);
  });
  return result;
}

function mapHeaderNames_(headerRow) {
  const map = {};
  headerRow.forEach(function(value, index) {
    map[normalizeHeader_(value)] = index;
  });
  return map;
}

function requireHeaderIndex_(headerMap, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    const key = normalizeHeader_(aliases[i]);
    if (Object.prototype.hasOwnProperty.call(headerMap, key)) return headerMap[key];
  }
  throw new Error('ไม่พบหัวตาราง: ' + aliases.join(' / '));
}

function normalizeHeader_(value) {
  return cleanText_(value).toLowerCase().replace(/[\s_\-]+/g, '');
}

function requireSheet_(spreadsheet, name) {
  const sheet = spreadsheet.getSheetByName(name);
  if (!sheet) throw new Error('ไม่พบแท็บ "' + name + '"');
  return sheet;
}

function getBusinessDateKey_() {
  const props = PropertiesService.getScriptProperties();
  let key = props.getProperty(SPAQ.PROPERTY_ACTIVE_DATE);
  if (!key) {
    key = initialBusinessDateKey_();
    props.setProperty(SPAQ.PROPERTY_ACTIVE_DATE, key);
  }
  return key;
}

function initialBusinessDateKey_() {
  const now = new Date();
  const hour = Number(Utilities.formatDate(now, SPAQ.TIMEZONE, 'H'));
  const effective = new Date(now.getTime());
  if (hour < 10) effective.setDate(effective.getDate() - 1);
  return Utilities.formatDate(effective, SPAQ.TIMEZONE, 'yyyy-MM-dd');
}

function dateFromKey_(key) {
  const parts = String(key).split('-').map(Number);
  return new Date(parts[0], parts[1] - 1, parts[2], 12, 0, 0);
}

function thaiMonthSheetName_(date) {
  return SPAQ_THAI_MONTHS[date.getMonth()] + ' ' + (date.getFullYear() + 543);
}

function parseTimeMinutes_(value, displayValue) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    const parts = Utilities.formatDate(value, SPAQ.TIMEZONE, 'H:mm').split(':');
    return Number(parts[0]) * 60 + Number(parts[1]);
  }

  if (typeof value === 'number' && Number.isFinite(value)) {
    const fraction = ((value % 1) + 1) % 1;
    return Math.round(fraction * 24 * 60) % (24 * 60);
  }

  const text = cleanText_(displayValue || value);
  const match = text.match(/(\d{1,2})\s*[:.]\s*(\d{1,2})/);
  if (!match) return null;
  const hour = Number(match[1]);
  const minute = Number(match[2]);
  if (hour < 0 || hour > 24 || minute < 0 || minute > 59) return null;
  if (hour === 24 && minute !== 0) return null;
  return hour === 24 ? 0 : hour * 60 + minute;
}

function minutesToClock_(minutes) {
  const normalized = ((minutes % (24 * 60)) + 24 * 60) % (24 * 60);
  return pad2_(Math.floor(normalized / 60)) + ':' + pad2_(normalized % 60);
}

function formatRegistryTimeRange_(start, end) {
  return minutesToClock_(start).replace(':', '.') + '-' +
    minutesToClock_(end).replace(':', '.');
}

function roundHours_(hours) {
  return Math.round(hours * 100) / 100;
}

function buildBookingId_(booking) {
  const stableParts = [
    booking.businessDate,
    booking.therapistId,
    minutesToClock_(booking.startMinutes),
    minutesToClock_(booking.endMinutes),
    normalizeForMatch_(booking.customer),
    normalizeForMatch_(booking.room),
    normalizeForMatch_(booking.phone),
    normalizeForMatch_(booking.course)
  ];
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    stableParts.join('|'),
    Utilities.Charset.UTF_8
  );
  return digest.map(function(byte) {
    const value = byte < 0 ? byte + 256 : byte;
    return ('0' + value.toString(16)).slice(-2);
  }).join('').substring(0, 24);
}

function combineRegistryNote_(customerType, note) {
  const parts = [];
  if (customerType) parts.push(customerType);
  if (note && parts.indexOf(note) < 0) parts.push(note);
  return parts.join(' | ');
}

function normalizeStatus_(value) {
  const text = cleanText_(value);
  const compact = text.replace(/\s+/g, '');
  const known = {};
  known[SPAQ.STATUS.AVAILABLE.replace(/\s+/g, '')] = SPAQ.STATUS.AVAILABLE;
  known[SPAQ.STATUS.BUSY.replace(/\s+/g, '')] = SPAQ.STATUS.BUSY;
  known[SPAQ.STATUS.LAST.replace(/\s+/g, '')] = SPAQ.STATUS.LAST;
  known[SPAQ.STATUS.OPEN.replace(/\s+/g, '')] = SPAQ.STATUS.OPEN;
  known[SPAQ.STATUS.CLOSED.replace(/\s+/g, '')] = SPAQ.STATUS.CLOSED;
  return known[compact] || text;
}

function publicStatusEnglish_(status) {
  const map = {};
  map[SPAQ.STATUS.AVAILABLE] = 'Available';
  map[SPAQ.STATUS.BUSY] = 'Occupied';
  map[SPAQ.STATUS.LAST] = 'Last Booking';
  map[SPAQ.STATUS.OPEN] = 'Spa Open';
  map[SPAQ.STATUS.CLOSED] = 'Spa Closed';
  return map[status] || 'Checking';
}

function isWorkingValue_(value) {
  const normalized = normalizeForMatch_(value);
  return ['true', '1', 'yes', 'y', 'active', 'ทำงาน', 'เปิด'].indexOf(normalized) >= 0;
}

function normalizeClockText_(value) {
  const minutes = parseTimeMinutes_(value, value);
  return minutes === null ? '' : minutesToClock_(minutes);
}

function normalizeForMatch_(value) {
  return cleanText_(value).toLowerCase().replace(/[\s.\-_/]+/g, '');
}

function cleanText_(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

function pad2_(number) {
  return ('0' + number).slice(-2);
}

function spaNowIso_() {
  return Utilities.formatDate(new Date(), SPAQ.TIMEZONE, "yyyy-MM-dd'T'HH:mm:ssXXX");
}

/**
 * Safe parser/idempotency checks that do not read or write Sheets.
 */
function runSpaQueueSelfTest() {
  const tests = [];
  function expect(name, actual, expected) {
    const pass = JSON.stringify(actual) === JSON.stringify(expected);
    tests.push({ name: name, pass: pass, actual: actual, expected: expected });
    if (!pass) throw new Error('Self-test failed: ' + name);
  }

  expect('time colon', parseTimeMinutes_('', '14:13'), 853);
  expect('time dot', parseTimeMinutes_('', '15.43'), 943);
  expect('midnight', parseTimeMinutes_('', '0:00'), 0);
  expect('clock', minutesToClock_(24 * 60), '00:00');
  expect('hours', roundHours_((17 * 60 + 40 - (16 * 60 + 10)) / 60), 1.5);
  expect('status', normalizeStatus_('คิวสุดท้ายของวัน'), SPAQ.STATUS.LAST);
  expect('buddhist month', thaiMonthSheetName_(new Date(2026, 7, 1)), 'สิงหาคม 2569');
  expect('closing hour', SPAQ.END_HOUR, 23);

  console.log(JSON.stringify(tests, null, 2));
  return tests;
}