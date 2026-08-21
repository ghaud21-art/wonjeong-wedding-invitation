/**
 * 모바일 청첩장 백엔드 (Google Apps Script)
 * 배포: 웹 앱 / 액세스 "전체(누구나)" / 실행 계정 "나"
 * 코드 수정 후에는 "배포 관리 → 수정 → 새 버전"으로 업데이트 (URL 유지)
 */
const SHEET_ID = '1sEuNQ8Abrk0E4nTlhxsqrBsY6ejHurV6Tw3jfOnGfLk';            // "원정이 청첩장 - 설정" 시트
const GALLERY_FOLDER_ID = '1FHTxbFIM5IprAIqBhZvYHFGp3FgpaETI';   // 갤러리 루트 폴더 (스튜디오 / 스냅사진 / 메인 사진 서브폴더)
const GUEST_PHOTO_FOLDER_ID = '1SDG1lrCekux_vqzOLzdQz7sw6iOQYGRO'; // 하객사진 폴더

function ss_() {
  return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ── 시트 자동 생성 + 예시값 ── */
function ensureSheets_() {
  const ss = ss_();
  let cfg = ss.getSheetByName('설정');
  if (!cfg) {
    cfg = ss.insertSheet('설정');
    cfg.getRange(1, 1, 1, 4).setValues([['항목', '내용', '입력 예시', '키(수정 금지)']]);
    const rows = [
      ['신랑 이름', '', '김철수', 'GROOM'], ['신부 이름', '', '이영희', 'BRIDE'],
      ['공유용 영문 이름(선택)', '', 'Chulsoo & Younghee', 'SCRIPT'],
      ['예식 일시 (형식을 반드시 지켜주세요)', '', '2026-06-20T14:00:00+09:00', 'WEDDING'],
      ['예식 일시 안내 문구', '', '2026년 6월 20일 토요일 14:00', 'DATE_TEXT'],
      ['예식 일시 짧은 문구', '', '2026. 06. 20  |  토요일 14:00', 'DATE_SHORT'],
      ['예식장 이름', '', '○○웨딩홀 3층 그랜드홀', 'VENUE'], ['예식장 주소', '', '서울시 ○○구 ○○로 123', 'ADDR'],
      ['인사말 첫째 줄', '', '저희 두 사람, 새로운 시작을 함께해 주세요.', 'INTRO1'],
      ['인사말 둘째 줄', '', '따뜻한 마음으로 축복해 주시면 감사하겠습니다.', 'INTRO2'],
      ['신랑측 아버지 성함', '', '아버지 성함', 'GROOM_FATHER'], ['신랑측 어머니 성함', '', '어머니 성함', 'GROOM_MOTHER'],
      ['신부측 아버지 성함', '', '아버지 성함', 'BRIDE_FATHER'], ['신부측 어머니 성함', '', '어머니 성함', 'BRIDE_MOTHER'],
      ['카카오내비 링크', '', 'https://kko.to/...', 'NAVI_KAKAO'], ['네이버지도 링크', '', 'https://naver.me/...', 'NAVI_NAVER'], ['티맵 링크', '', 'https://tmap.life/...', 'NAVI_TMAP'],
      ['지하철 안내', '', '2호선 ○○역 3번 출구 도보 5분', 'TRAFFIC_SUBWAY'], ['버스 안내', '', '146, 341 ○○정류장 하차', 'TRAFFIC_BUS'],
      ['자가용 안내', '', '내비게이션에 "○○웨딩홀" 검색', 'TRAFFIC_CAR'], ['주차 안내', '', '건물 지하 2시간 무료', 'TRAFFIC_PARKING'],
      ['식사 안내 시간', '', '10:30 - 12:30 (8층)', 'MEAL_TIME'],
      ['식사 안내 설명', '', '귀한 걸음 해주신 분들을 위해 뷔페 식사가 마련되어 있습니다.\n연회장은 홀과 다른 건물에 위치해 있습니다.', 'MEAL_INFO'],
      ['주차 안내 시간', '', '2시간 무료 (지하 2층)', 'PARKING_TIME'],
      ['주차 안내 설명', '', '주차장은 홀과 다른 건물에 있으니 안내 데스크를 이용해주세요.', 'PARKING_INFO']
    ];
    cfg.getRange(2, 1, rows.length, 4).setValues(rows);
    // 날짜/시간이 Date 타입으로 자동 변환되며 생기는 시간대 오차 방지: '내용' 열 텍스트 서식 고정
    cfg.getRange('B:B').setNumberFormat('@');
    cfg.setColumnWidth(1, 200); cfg.setColumnWidth(2, 220); cfg.setColumnWidth(4, 90);
  }
  let acc = ss.getSheetByName('계좌');
  if (!acc) {
    acc = ss.insertSheet('계좌');
    acc.getRange(1, 1, 1, 6).setValues([['측', '관계', '이름', '은행', '계좌번호', '카카오페이 링크']]);
    acc.getRange(2, 1, 2, 6).setValues([
      ['신랑측', '신랑', '', '', '', ''],
      ['신부측', '신부', '', '', '', '']
    ]);
    acc.getRange('E:E').setNumberFormat('@');
  }
  let gb = ss.getSheetByName('방명록');
  if (!gb) {
    gb = ss.insertSheet('방명록');
    gb.getRange(1, 1, 1, 5).setValues([['시각', '이름', '메시지', '비밀번호', 'ID']]);
  }
  let rs = ss.getSheetByName('참석');
  if (!rs) {
    rs = ss.insertSheet('참석');
    rs.getRange(1, 1, 1, 7).setValues([['시각', '이름', '연락처', '하객측', '참석여부', '식사여부', '동행인원(본인포함)']]);
    rs.getRange('C:C').setNumberFormat('@');
  }
}

/* ── 설정 읽기 ──
 * 행이 없는 항목은 응답에 아예 포함하지 않음(프론트가 코드 기본값 유지),
 * 행이 있고 비워둔 항목은 빈 문자열로 포함(프론트가 빈 값으로 덮어씀). */
function getConfig_() {
  ensureSheets_();
  const ss = ss_();
  const out = {};
  const cfg = ss.getSheetByName('설정').getDataRange().getValues();
  for (let i = 1; i < cfg.length; i++) {
    const key = String(cfg[i][3]).trim(); // D열의 기술 키를 기준으로 읽음 (A열은 사람이 보는 한글 라벨)
    if (key) out[key] = String(cfg[i][1]);
  }
  // 부모님 줄: 가운뎃점 조합은 프론트에서 처리 (빈 값 제외 후 · 연결)
  if ('GROOM_FATHER' in out || 'GROOM_MOTHER' in out) out.GROOM_PARENTS = [out.GROOM_FATHER || '', out.GROOM_MOTHER || ''];
  if ('BRIDE_FATHER' in out || 'BRIDE_MOTHER' in out) out.BRIDE_PARENTS = [out.BRIDE_FATHER || '', out.BRIDE_MOTHER || ''];
  if ('TRAFFIC_SUBWAY' in out) out.TRAFFIC = [['지하철', out.TRAFFIC_SUBWAY || ''], ['버스', out.TRAFFIC_BUS || ''], ['자가용', out.TRAFFIC_CAR || ''], ['주차', out.TRAFFIC_PARKING || '']];
  // 계좌
  const acc = ss.getSheetByName('계좌').getDataRange().getValues();
  const g = [], b = [];
  for (let i = 1; i < acc.length; i++) {
    const [side, rel, name, bank, num, kpay] = acc[i].map(String);
    if (!rel.trim()) continue;
    const row = { rel: rel, name: name, bank: bank, num: num, kpay: kpay.trim() };
    (side.indexOf('신랑') > -1 ? g : b).push(row);
  }
  if (g.length) out.ACC_G = g;
  if (b.length) out.ACC_B = b;
  return out;
}

/* ── 드라이브 갤러리 ── */
function listFolderImages_(folder) {
  const files = folder.getFiles();
  const arr = [];
  while (files.hasNext()) {
    const f = files.next();
    if (String(f.getMimeType()).indexOf('image/') === 0) {
      arr.push({
        name: f.getName(),
        thumb: 'https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w800',
        full: 'https://drive.google.com/uc?export=view&id=' + f.getId()
      });
    }
  }
  arr.sort(function (a, b) { return a.name.localeCompare(b.name); });
  return arr;
}
/* ── 방명록 행 찾기: ID열이 비어있는 옛 데이터는 행 번호로 대체 식별 ── */
function guestRowId_(values, i) {
  return String(values[i][4] || '').trim() || ('r' + i);
}
function findGuestRowIndex_(values, id) {
  for (let i = 1; i < values.length; i++) {
    if (guestRowId_(values, i) === id) return i;
  }
  return -1;
}

function getGallery_() {
  if (!GALLERY_FOLDER_ID) return { tabs: {}, main: null };
  const root = DriveApp.getFolderById(GALLERY_FOLDER_ID);
  const subs = root.getFolders();
  const tabs = {}; let main = null;
  while (subs.hasNext()) {
    const sub = subs.next();
    const imgs = listFolderImages_(sub);
    if (sub.getName() === '메인 사진') { if (imgs.length) main = imgs[0].full; }
    else tabs[sub.getName()] = imgs;
  }
  return { tabs: tabs, main: main };
}

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'config';
  try {
    if (action === 'config') return json_({ ok: true, config: getConfig_() });
    if (action === 'gallery') return json_({ ok: true, gallery: getGallery_() });
    if (action === 'guestbook') {
      ensureSheets_();
      const v = ss_().getSheetByName('방명록').getDataRange().getValues();
      const list = [];
      // 비밀번호(4번째 열)는 절대 응답에 포함하지 않음
      for (let i = v.length - 1; i >= 1; i--) list.push({ id: guestRowId_(v, i), time: String(v[i][0]), name: String(v[i][1]), msg: String(v[i][2]) });
      return json_({ ok: true, list: list });
    }
    if (action === 'photos') {
      if (!GUEST_PHOTO_FOLDER_ID) return json_({ ok: true, list: [] });
      return json_({ ok: true, list: listFolderImages_(DriveApp.getFolderById(GUEST_PHOTO_FOLDER_ID)) });
    }
    return json_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    // 프론트는 CORS 회피를 위해 Content-Type: text/plain;charset=utf-8 로 전송
    const body = JSON.parse(e.postData.contents);
    ensureSheets_();
    const ss = ss_();
    const now = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
    if (body.action === 'guestbook') {
      ss.getSheetByName('방명록').appendRow([now, body.name, body.msg, body.pw || '', body.id || '']);
      return json_({ ok: true });
    }
    if (body.action === 'guestbook_edit' || body.action === 'guestbook_delete') {
      const sheet = ss.getSheetByName('방명록');
      const values = sheet.getDataRange().getValues();
      const idx = findGuestRowIndex_(values, String(body.id || ''));
      if (idx === -1) return json_({ ok: false, error: '메시지를 찾을 수 없습니다' });
      const savedPw = String(values[idx][3] || '');
      if (savedPw !== String(body.pw || '')) return json_({ ok: false, error: '비밀번호가 일치하지 않습니다' });
      if (body.action === 'guestbook_edit') sheet.getRange(idx + 1, 3).setValue(body.msg || '');
      else sheet.deleteRow(idx + 1);
      return json_({ ok: true });
    }
    if (body.action === 'rsvp') {
      ss.getSheetByName('참석').appendRow([now, body.name, body.phone, body.side === 'groom' ? '신랑측' : '신부측', body.attend === 'y' ? '참석' : '불참', { y: '식사함', n: '식사 안 함', u: '미정' }[body.meal] || '미정', body.count]);
      return json_({ ok: true });
    }
    if (body.action === 'upload') {
      // body.data: base64, body.type: mime, body.name: 파일명
      if (!GUEST_PHOTO_FOLDER_ID) return json_({ ok: false, error: 'folder not set' });
      const blob = Utilities.newBlob(Utilities.base64Decode(body.data), body.type, body.name || ('guest_' + now + '.jpg'));
      DriveApp.getFolderById(GUEST_PHOTO_FOLDER_ID).createFile(blob);
      return json_({ ok: true });
    }
    return json_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}
