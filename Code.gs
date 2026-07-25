/**
 * 호캠프 리턴즈 — 사이트 실시간 재고 관리용 Apps Script
 * (구글 스프레드시트 하나를 "사이트 재고 DB"로 사용합니다)
 *
 * 설치 방법은 함께 받은 "설치_안내.md"를 참고하세요.
 */

var SHEET_NAME = '사이트재고';

// 사이트 구역 구성 — 여기 값만 바꾸면 initSites() 실행 시 자동으로 사이트 목록이 다시 생성됩니다.
var ZONES = [
  { zone: 'A', start: 1,  end: 20 },
  { zone: 'C', start: 23, end: 38 },
  { zone: 'D', start: 41, end: 42 },
  { zone: 'F', start: 50, end: 54 },
  { zone: 'G', start: 57, end: 57 },
  { zone: 'H', start: 58, end: 68 }
];

// 번호 규칙에서 벗어나는 개별 사이트를 추가로 넣고 싶을 때 여기에 추가하세요. (예: 30번 자리를 30-1/30-2로 분할)
var EXTRA_SITES = [
  { zone: 'C', id: 'C-30-1' },
  { zone: 'C', id: 'C-30-2' }
];

// ZONES 범위 안에서 타입/가격이 다른 특별 사이트 구간. initSites() 실행 시 자동 반영됩니다.
var SPECIAL_SITE_RANGES = [
  { zone: 'C', start: 31, end: 38, type: '캠핑카·카라반·트레일러 전용 대형 사이트', price: 200000 }
];

function findSpecialSite_(zone, n) {
  for (var i = 0; i < SPECIAL_SITE_RANGES.length; i++) {
    var r = SPECIAL_SITE_RANGES[i];
    if (r.zone === zone && n >= r.start && n <= r.end) return r;
  }
  return null;
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  return sheet;
}

/**
 * 최초 1회(또는 매년 시즌 초기화 시) 이 함수를 직접 실행하세요.
 * 위쪽 상단 드롭다운에서 initSites 선택 후 ▶ 실행 버튼을 누르면 됩니다.
 * 시트를 비우고 ZONES 설정에 맞춰 사이트 목록을 새로 채웁니다.
 */
function initSites() {
  var sheet = getSheet_();
  sheet.clear();
  sheet.appendRow(['SiteID', '구역', '타입', '추가금액(원)', '상태', '예약자명', '연락처', '예약시각']);
  ZONES.forEach(function(z) {
    for (var n = z.start; n <= z.end; n++) {
      var id = z.zone + '-' + n;
      var special = findSpecialSite_(z.zone, n);
      if (special) {
        sheet.appendRow([id, z.zone, special.type, special.price, '예약가능', '', '', '']);
      } else {
        sheet.appendRow([id, z.zone, '일반사이트', 0, '예약가능', '', '', '']);
      }
    }
  });
  EXTRA_SITES.forEach(function(s) {
    sheet.appendRow([s.id, s.zone, '일반사이트', 0, '예약가능', '', '', '']);
  });
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 8);
}

/**
 * GET 요청 — 사이트 목록 + 현재 상태 조회 (booking.html 페이지 로드 시 호출됨)
 */
function doGet(e) {
  var sheet = getSheet_();
  var data = sheet.getDataRange().getValues();
  var rows = data.slice(1);
  var list = rows
    .filter(function(r) { return r[0]; })
    .map(function(r) {
      return {
        id: r[0],
        zone: r[1],
        type: r[2],
        price: Number(r[3]) || 0,
        status: r[4]
        // 예약자명/연락처/예약시각은 개인정보이므로 외부에 노출하지 않습니다.
      };
    });
  return jsonOut_({ success: true, sites: list });
}

/**
 * POST 요청 — 사이트 예약 처리 (booking.html 제출 버튼 클릭 시 호출됨)
 * action=reserve  : 사이트 예약 확정 (재고 차감)
 * action=release   : 예약 취소 후 재고 복원 (관리자가 직접 취소 처리할 때 사용, 선택 사항)
 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (err) {
    return jsonOut_({ success: false, message: '서버가 혼잡합니다. 잠시 후 다시 시도해 주세요.' });
  }

  try {
    var params = e.parameter || {};
    var action = params.action;

    if (action === 'reserve') {
      return reserveSite_(params);
    } else if (action === 'release') {
      return releaseSite_(params);
    }
    return jsonOut_({ success: false, message: '알 수 없는 요청입니다.' });
  } finally {
    lock.releaseLock();
  }
}

function reserveSite_(params) {
  var siteId = params.siteId;
  var name = params.name || '';
  var phone = params.phone || '';

  if (!siteId) {
    return jsonOut_({ success: false, message: 'siteId가 필요합니다.' });
  }

  var sheet = getSheet_();
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === siteId) {
      if (data[i][4] === '예약완료') {
        return jsonOut_({ success: false, message: '이미 예약된 사이트입니다. 다른 사이트를 선택해 주세요.' });
      }
      var row = i + 1;
      sheet.getRange(row, 5).setValue('예약완료');
      sheet.getRange(row, 6).setValue(name);
      sheet.getRange(row, 7).setValue(phone);
      sheet.getRange(row, 8).setValue(new Date());
      return jsonOut_({ success: true, siteId: siteId });
    }
  }
  return jsonOut_({ success: false, message: '존재하지 않는 사이트입니다.' });
}

/**
 * 예약 취소 시 재고를 되돌립니다. (관리자가 스프레드시트에서 직접 상태를 "예약가능"으로
 * 바꿔도 되고, 필요하면 이 함수를 별도 관리자 도구에서 action=release로 호출해도 됩니다.)
 */
function releaseSite_(params) {
  var siteId = params.siteId;
  var sheet = getSheet_();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === siteId) {
      var row = i + 1;
      sheet.getRange(row, 5).setValue('예약가능');
      sheet.getRange(row, 6).setValue('');
      sheet.getRange(row, 7).setValue('');
      sheet.getRange(row, 8).setValue('');
      return jsonOut_({ success: true });
    }
  }
  return jsonOut_({ success: false, message: '존재하지 않는 사이트입니다.' });
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
