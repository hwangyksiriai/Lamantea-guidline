/**
 * 라만테아 12월 캠페인 신청서 → 구글시트 '시트1' 저장
 *
 * 설치 방법
 *  1. 시트에서 [확장 프로그램] → [Apps Script] 열기
 *  2. 기본 Code.gs 내용을 지우고 이 파일 내용 전체 붙여넣기 → 저장
 *  3. [배포] → [새 배포] → 유형: 웹 앱
 *       - 실행 계정: 나
 *       - 액세스 권한: 모든 사용자
 *  4. 권한 승인 후 발급된 웹 앱 URL(…/exec)을 가이드 페이지의 GAS_WEB_APP_URL에 넣기
 *
 * 코드를 수정한 뒤에는 [배포] → [배포 관리] → 수정(연필) → 버전: 새 버전 으로 다시 배포해야 반영됩니다.
 */

const SHEET_ID = '1b_4cLkLdqw0cUfb4pOOtvkLjhzypyHyenz4sd6bHYOI';
const SHEET_NAME = '시트1';
const TIMEZONE = 'Asia/Seoul';

// 열 순서 (key = 가이드 페이지에서 보내는 값)
const COLUMNS = [
  { key: 'submittedAt', label: '제출일시', width: 150 },
  { key: 'price',       label: '고료',     width: 80 },
  { key: 'name',        label: '이름',     width: 100 },
  { key: 'instagram',   label: '인스타그램', width: 230 },
  { key: 'phone',       label: '휴대폰',   width: 130, text: true },
  { key: 'email',       label: '이메일',   width: 190 },
  { key: 'zipcode',     label: '우편번호', width: 80,  text: true },
  { key: 'address',     label: '주소',     width: 340 },   // 배송지 주소 + 상세 주소
  { key: 'note',        label: '희망 제작 스타일', width: 320 },
  { key: 'productLine', label: '제공 제품', width: 260 }
];

// 고료별 색 — 행 전체는 옅게, 고료 칸은 진하게
const PRICE_COLORS = {
  '5만원':  { row: '#EEF6EA', cell: '#C9E3BD' },
  '10만원': { row: '#EAF2FB', cell: '#BFD8F2' },
  '20만원': { row: '#F4ECF8', cell: '#DCC5EA' }
};
const DEFAULT_COLOR = { row: '#FFFFFF', cell: '#EEEEEE' };

// 고료 요청이 의심되면 노란색으로 표시할 칸
const FEE_CHECK_KEYS = ['name', 'note'];
const FEE_HIGHLIGHT = '#FFE14D';
const FEE_PATTERN = /(원고료|고료|단가|비용|금액|협찬비|광고비|제작비|수수료|페이(?!지)|만\s*원|\d[\d,]*\s*원|\bfee\b|\bpay\b|\bcost\b)/i;

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    data.submittedAt = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
    appendApplication_(data);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// 브라우저로 웹 앱 URL을 열었을 때 연결 확인용
function doGet() {
  return json_({ ok: true, sheet: SHEET_NAME });
}

function appendApplication_(data) {
  const sheet = getSheet_();
  ensureHeader_(sheet);

  const row = sheet.getLastRow() + 1;
  const values = COLUMNS.map(c => String(data[c.key] == null ? '' : data[c.key]).trim());
  const range = sheet.getRange(row, 1, 1, COLUMNS.length);

  // 앞자리 0 보존: 값을 넣기 전에 텍스트 서식으로 지정
  COLUMNS.forEach((c, i) => {
    if (c.text) sheet.getRange(row, i + 1).setNumberFormat('@');
  });
  range.setValues([values]);
  range.setVerticalAlignment('middle').setWrap(true);

  // 고료별 색
  const color = PRICE_COLORS[values[colIndex_('price')]] || DEFAULT_COLOR;
  range.setBackground(color.row);
  sheet.getRange(row, colIndex_('price') + 1).setBackground(color.cell).setFontWeight('bold').setHorizontalAlignment('center');

  // 이름·희망 스타일에 고료 요청이 있으면 노란색
  FEE_CHECK_KEYS.forEach(key => {
    const i = colIndex_(key);
    if (FEE_PATTERN.test(values[i])) {
      sheet.getRange(row, i + 1).setBackground(FEE_HIGHLIGHT).setFontWeight('bold');
    }
  });
}

function ensureHeader_(sheet) {
  if (sheet.getLastRow() > 0) return;
  const header = sheet.getRange(1, 1, 1, COLUMNS.length);
  header.setValues([COLUMNS.map(c => c.label)])
    .setBackground('#363D21')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 34);
  COLUMNS.forEach((c, i) => {
    sheet.setColumnWidth(i + 1, c.width);
    // 휴대폰·우편번호 열 전체를 텍스트 서식으로
    if (c.text && sheet.getMaxRows() > 1) sheet.getRange(2, i + 1, sheet.getMaxRows() - 1, 1).setNumberFormat('@');
  });
}

function getSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  return ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
}

function colIndex_(key) {
  return COLUMNS.findIndex(c => c.key === key);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * 테스트용 — Apps Script 편집기에서 이 함수를 선택해 [실행]하면
 * 시트1에 샘플 3줄(5만원·10만원 고료요청·20만원)이 들어갑니다. 확인 후 행을 지워주세요.
 */
function testAppend() {
  const base = {
    productLine: '라만테아 EVO 선물 패키지 (500ml 병 1 + 250ml 캔 2)',
    instagram: 'https://instagram.com/test',
    phone: '010-0123-4567',
    email: 'test@example.com',
    zipcode: '06035',
    address: '서울 강남구 가로수길 5 101호'
  };
  [
    Object.assign({}, base, { price: '5만원',  name: '테스트A', note: '연말 홈파티 테이블' }),
    Object.assign({}, base, { price: '10만원', name: '테스트B', note: '언박싱 위주로, 고료 15만원 희망합니다' }),
    Object.assign({}, base, { price: '20만원', name: '테스트C', note: '' })
  ].forEach(d => {
    d.submittedAt = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
    appendApplication_(d);
  });
}
