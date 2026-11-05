/**
 * 121 체크 → 구글시트 취합 (Apps Script 웹앱)
 * 시트 1개당 한 번만 설정: 확장 프로그램 > Apps Script > 이 코드 붙여넣기 > 배포 > 웹 앱
 *   실행: 나 / 액세스: 모든 사용자 → 나온 주소를 index.html 의 SHEET_URL 에 붙여넣기
 * 시트 '체크' = 한 줄에 (나, 상대, 단계). 멤버가 저장할 때마다 그 사람 줄만 통째로 교체한다.
 * 로그인: 페이지가 이름(me)과 전화번호 뒤 4자리(pin)를 함께 보낸다. 아래 PINS와 맞아야만 읽고 쓴다(남의 체크를 못 지운다).
 *   PINS의 '____' 자리는 저장소에 번호를 올리지 않으려는 자리표시다. 실제 번호는 Apps Script 편집기에만 넣는다.
 * 시트 '매트릭스' = 메뉴 '121 체크 > 매트릭스 만들기'로 생성. 둘 중 한쪽이라도 체크하면 그 단계로 본다.
 */
const NAMES = ["", "패스포트", "아는", "신뢰", "수익"];
const PINS = {
  "송승훈": "____",
  "이상호": "____",
  "조현우": "____",
  "김근우": "____",
  "김경태": "____",
  "김지현": "____",
  "심학봉": "____",
  "이화춘": "____",
  "이소연": "____",
  "임춘식": "____",
  "정명수": "____",
  "홍정택": "____",
  "박진성": "____",
  "박미성": "____",
  "이채홍": "____",
  "조은영": "____",
  "박선영": "____",
  "정상현": "____",
  "이도현": "____",
  "마루이": "____",
  "문성우": "____",
  "윤민수": "____",
  "이수민": "____",
  "박지형": "____",
  "박병준": "____",
  "조진성": "____",
  "이해경": "____",
  "이은성": "____",
  "오예준": "____",
  "이훈": "____",
  "김여련": "____",
  "최소은": "____",
  "이다경": "____",
  "송이": "____",
  "이서원": "____",
  "이헌영": "____",
  "김성일": "____"
};
function ok_(me, pin) { return !!me && Object.prototype.hasOwnProperty.call(PINS, me) && PINS[me] === String(pin || ""); }

function sheet_() {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName("체크");
  if (!s) { s = ss.insertSheet("체크"); s.appendRow(["갱신시각", "나", "상대", "단계(1~4)", "단계명"]); s.setFrozenRows(1); }
  return s;
}

function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  if (!ok_(d.me, d.pin)) return ContentService.createTextOutput("denied");
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const s = sheet_(), last = s.getLastRow();
    if (last > 1) {
      const me = s.getRange(2, 2, last - 1, 1).getValues();
      for (let i = me.length - 1; i >= 0; i--) if (me[i][0] === d.me) s.deleteRow(i + 2);
    }
    const now = new Date();
    // 페이지가 보낸 단계명(st)을 기준으로 번호를 정한다. 페이지·스크립트 버전이 달라도 번호와 이름이 어긋나지 않는다.
    const out = d.rows.map(r => { const lv = r.st ? NAMES.indexOf(r.st) : r.lv; return lv > 0 ? [now, d.me, r.p, lv, NAMES[lv]] : null; }).filter(Boolean);
    if (out.length) s.getRange(s.getLastRow() + 1, 1, out.length, 5).setValues(out);
  } finally { lock.releaseLock(); }
  return ContentService.createTextOutput("ok");
}

// 본인 기존 체크 불러오기(본인 줄만 돌려준다)
function doGet(e) {
  if (!ok_(e.parameter.me, e.parameter.pin)) return ContentService.createTextOutput(JSON.stringify({ error: "auth" })).setMimeType(ContentService.MimeType.JSON);
  const s = sheet_(), last = s.getLastRow(), ans = {}, st = {};
  if (last > 1) s.getRange(2, 1, last - 1, 5).getValues().forEach(r => { if (r[1] === e.parameter.me) { ans[r[2]] = r[3]; st[r[2]] = r[4]; } });
  return ContentService.createTextOutput(JSON.stringify({ ans, st })).setMimeType(ContentService.MimeType.JSON);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu("121 체크").addItem("매트릭스 만들기", "matrix").addToUi();
}

function matrix() {
  const ss = SpreadsheetApp.getActive(), s = sheet_(), last = s.getLastRow();
  const rows = last > 1 ? s.getRange(2, 2, last - 1, 3).getValues() : [];
  const names = [...new Set(rows.flatMap(r => [r[0], r[1]]))].sort();
  const lv = {}; // "a|b" -> 둘 중 높은 단계
  rows.forEach(r => { [r[0] + "|" + r[1], r[1] + "|" + r[0]].forEach(k => lv[k] = Math.max(lv[k] || 0, r[2])); });
  const grid = [[""].concat(names, ["121 한 사람 수"])];
  names.forEach(a => {
    const line = names.map(b => a === b ? "·" : (lv[a + "|" + b] ? NAMES[lv[a + "|" + b]] : ""));
    grid.push([a].concat(line, [line.filter(x => x && x !== "·").length]));
  });
  let m = ss.getSheetByName("매트릭스") || ss.insertSheet("매트릭스");
  m.clear();
  if (names.length) m.getRange(1, 1, grid.length, grid[0].length).setValues(grid);
  m.setFrozenRows(1); m.setFrozenColumns(1);
}
