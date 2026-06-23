// ═══════════════════════════════════════════════════════════════════
// Bravo Korea Loan — Google Apps Script
// 기능: 리드 저장 + 5분마다 Facebook Page 메시지 알림
//
// ─── 설정 방법 ───────────────────────────────────────────────────
// 1. Google Sheets 새 파일 생성
// 2. 확장프로그램 → Apps Script → 이 코드 전체 붙여넣기
// 3. 아래 CONFIG 값 채우기
// 4. 저장 후 '배포' → '새 배포' → 웹 앱
//    실행 계정: 나 / 액세스 권한: 모든 사용자
// 5. 배포 URL → HTML 파일의 APPS_SCRIPT_URL 에 붙여넣기
// 6. setupTrigger() 함수 한 번 실행 → 5분 자동 알림 활성화
//
// ─── Facebook Page Token 발급 ────────────────────────────────────
// 1. developers.facebook.com → 내 앱 → Graph API Explorer
// 2. Page Access Token 생성 (페이지 관리자 권한 필요)
// 3. 토큰은 장기 토큰으로 교환 권장 (60일 유효)
// 4. FACEBOOK_RECIPIENT_ID = 알림 받을 Facebook 계정의 PSID
//    (페이지에 메시지를 먼저 보낸 사람만 수신 가능 — Facebook 정책)
// ═══════════════════════════════════════════════════════════════════

const CONFIG = {
  SHEET_NAME: 'Leads',
  
  // Facebook Page API
  FB_PAGE_TOKEN: 'YOUR_FB_PAGE_ACCESS_TOKEN',  // ← 여기에 입력
  FB_RECIPIENT_ID: 'YOUR_FB_PSID',             // ← 알림 받을 PSID
  
  // 알림 설정
  NOTIFY_INTERVAL_MINUTES: 5,
  UNREAD_COL: 11,   // K열 = 확인여부
};

const FLAG_MAP = {
  'Philippines': '🇵🇭', 'Nepal': '🇳🇵', 'Sri Lanka': '🇱🇰',
  'Indonesia': '🇮🇩', 'Myanmar': '🇲🇲', 'Cambodia': '🇰🇭',
  'Bangladesh': '🇧🇩', 'Uzbekistan': '🇺🇿', 'Kyrgyzstan': '🇰🇬',
  'China': '🇨🇳', 'Vietnam': '🇻🇳', 'Thailand': '🇹🇭',
  'Mongolia': '🇲🇳',
};

// ── 리드 저장 (POST) ────────────────────────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEET_NAME);
      const headers = ['타임스탬프','전화번호','국적','비자','재직기간','연봉구간','결과','대출상품','최대한도','유입경로','확인여부'];
      sheet.getRange(1,1,1,headers.length).setValues([headers]);
      const hr = sheet.getRange(1,1,1,headers.length);
      hr.setBackground('#003087').setFontColor('#fff').setFontWeight('bold');
      sheet.setFrozenRows(1);
      [160,140,100,80,120,130,90,130,110,100,80].forEach((w,i)=>sheet.setColumnWidth(i+1,w));
    }

    const now = new Date();
    const kst = new Date(now.getTime() + 9*60*60*1000);
    const ts = Utilities.formatDate(kst, 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');

    sheet.appendRow([
      ts,
      data.phone || '(미입력)',
      data.nationality || '',
      data.visa || '',
      data.months || '',
      data.salary || '',
      data.result || '',
      data.product || '',
      data.amount || '',
      data.source || 'direct',
      '미확인',  // K열 — 확인여부 초기값
    ]);

    const lastRow = sheet.getLastRow();
    // 결과 색상
    if (data.result === 'eligible') {
      sheet.getRange(lastRow, 7).setBackground('#e8f5e9').setFontColor('#1b5e20');
    } else if (data.result === 'ineligible') {
      sheet.getRange(lastRow, 7).setBackground('#fff8e1').setFontColor('#e65100');
    }
    // 미확인 셀 강조
    sheet.getRange(lastRow, CONFIG.UNREAD_COL)
      .setBackground('#fff3e0').setFontColor('#e65100').setFontWeight('bold');

    return ContentService
      .createTextOutput(JSON.stringify({status:'ok', row:lastRow}))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({status:'error', message:err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── GET 테스트 ───────────────────────────────────────────────────
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({status:'ok', message:'Bravo Korea Lead Collector running!'}))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── 5분 알림 트리거 ──────────────────────────────────────────────
function checkAndNotify() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // 미확인 행 수집
  const data = sheet.getRange(2, 1, lastRow-1, CONFIG.UNREAD_COL).getValues();
  const unread = [];

  data.forEach((row, i) => {
    if (row[CONFIG.UNREAD_COL - 1] === '미확인') {
      unread.push({
        rowNum: i + 2,
        ts:         row[0],
        phone:      row[1],
        nationality:row[2],
        visa:       row[3],
        months:     row[4],
        salary:     row[5],
        result:     row[6],
        product:    row[7],
        amount:     row[8],
      });
    }
  });

  if (unread.length === 0) return;

  // 메시지 구성
  const msg = buildNotificationMessage(unread);

  // Facebook Messenger 발송
  sendFbMessage(msg);
}

function buildNotificationMessage(unread) {
  const total = unread.length;
  const eligible = unread.filter(r => r.result === 'eligible').length;
  const ineligible = total - eligible;

  let msg = `🏦 Bravo Korea 신규 상담 알림\n`;
  msg += `━━━━━━━━━━━━━━━━━━\n`;
  msg += `📊 미확인 ${total}건 (✅ ${eligible}건 / ⚠️ ${ineligible}건)\n`;
  msg += `━━━━━━━━━━━━━━━━━━\n\n`;

  unread.forEach((r, i) => {
    const flag = FLAG_MAP[r.nationality] || '🌐';
    const resultIcon = r.result === 'eligible' ? '✅' : '⚠️';
    const amtText = r.amount ? ` / ₩${Number(r.amount).toLocaleString()}` : '';
    msg += `${i+1}. ${flag} ${r.nationality} | ${r.visa} | ${r.result === 'eligible' ? r.product : '미충족'}${amtText}\n`;
    msg += `   📞 ${r.phone}  ${resultIcon}\n`;
    msg += `   🕐 ${r.ts}\n\n`;
  });

  msg += `━━━━━━━━━━━━━━━━━━\n`;
  msg += `👉 구글 시트에서 확인 후 "확인" 처리하세요.`;

  return msg;
}

function sendFbMessage(text) {
  if (!CONFIG.FB_PAGE_TOKEN || CONFIG.FB_PAGE_TOKEN === 'YOUR_FB_PAGE_ACCESS_TOKEN') {
    Logger.log('FB Token 미설정 — 메시지 발송 건너뜀:\n' + text);
    return;
  }

  const url = 'https://graph.facebook.com/v19.0/me/messages?access_token=' + CONFIG.FB_PAGE_TOKEN;
  const payload = {
    recipient: { id: CONFIG.FB_RECIPIENT_ID },
    message: { text: text.substring(0, 2000) }, // FB 메시지 2000자 제한
    messaging_type: 'MESSAGE_TAG',
    tag: 'CONFIRMED_EVENT_UPDATE',
  };

  try {
    UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });
    Logger.log('FB 알림 발송 완료: ' + unread.length + '건');
  } catch(e) {
    Logger.log('FB 발송 실패: ' + e.toString());
  }
}

// ── 시트에서 "확인" 버튼 처리 ────────────────────────────────────
// 시트에서 K열 셀을 선택 후 이 함수 실행하면 "확인" 처리됨
function markAsRead() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  const cell = sheet.getActiveCell();
  const col = cell.getColumn();
  const row = cell.getRow();
  if (col === CONFIG.UNREAD_COL && row > 1) {
    cell.setValue('확인').setBackground('#e8f5e9').setFontColor('#1b5e20').setFontWeight('normal');
    SpreadsheetApp.getUi().alert('✅ ' + row + '행 확인 처리 완료');
  } else {
    SpreadsheetApp.getUi().alert('K열(확인여부)의 셀을 선택 후 실행해주세요.');
  }
}

// 전체 일괄 확인처리
function markAllAsRead() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const range = sheet.getRange(2, CONFIG.UNREAD_COL, lastRow-1, 1);
  const vals = range.getValues();
  vals.forEach((v,i) => {
    if (v[0] === '미확인') {
      const cell = sheet.getRange(i+2, CONFIG.UNREAD_COL);
      cell.setValue('확인').setBackground('#e8f5e9').setFontColor('#1b5e20').setFontWeight('normal');
    }
  });
  SpreadsheetApp.getUi().alert('✅ 전체 미확인 건 처리 완료');
}

// ── 트리거 설정 (최초 1회만 실행) ──────────────────────────────
function setupTrigger() {
  // 기존 트리거 제거
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'checkAndNotify') {
      ScriptApp.deleteTrigger(t);
    }
  });
  // 5분마다 실행
  ScriptApp.newTrigger('checkAndNotify')
    .timeBased()
    .everyMinutes(CONFIG.NOTIFY_INTERVAL_MINUTES)
    .create();
  SpreadsheetApp.getUi().alert('✅ 5분 알림 트리거 설정 완료!\n\nFB Token과 PSID를 CONFIG에 입력했는지 확인하세요.');
}

// ── 커스텀 메뉴 (시트 열면 상단에 메뉴 추가) ─────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏦 Bravo Korea')
    .addItem('✅ 선택행 확인처리', 'markAsRead')
    .addItem('✅ 전체 확인처리', 'markAllAsRead')
    .addSeparator()
    .addItem('🔔 알림 트리거 설정 (최초 1회)', 'setupTrigger')
    .addItem('📨 지금 바로 알림 테스트', 'checkAndNotify')
    .addToUi();
}
