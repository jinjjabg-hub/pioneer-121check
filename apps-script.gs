/**
 * 121 체크 → 구글시트 취합 (Apps Script 웹앱)
 * 시트 1개당 한 번만 설정: 확장 프로그램 > Apps Script > 이 코드 붙여넣기 > 배포 > 웹 앱
 *   실행: 나 / 액세스: 모든 사용자 → 나온 주소를 index.html 의 SHEET_URL 에 붙여넣기
 * 시트 '체크' = 한 줄에 (나, 상대, 단계). 멤버가 저장할 때마다 그 사람 줄만 통째로 교체한다.
 * 시트 '매트릭스' = 메뉴 '121 체크 > 매트릭스 만들기'로 생성. 둘 중 한쪽이라도 체크하면 그 단계로 본다.
 */
const NAMES = ["", "아는", "신뢰", "수익", "패스포트"];

function sheet_() {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName("체크");
  if (!s) { s = ss.insertSheet("체크"); s.appendRow(["갱신시각", "나", "상대", "단계(1~4)", "단계명"]); s.setFrozenRows(1); }
  return s;
}

function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const s = sheet_(), last = s.getLastRow();
    if (last > 1) {
      const me = s.getRange(2, 2, last - 1, 1).getValues();
      for (let i = me.length - 1; i >= 0; i--) if (me[i][0] === d.me) s.deleteRow(i + 2);
    }
    const now = new Date();
    const out = d.rows.filter(r => r.lv > 0).map(r => [now, d.me, r.p, r.lv, NAMES[r.lv]]);
    if (out.length) s.getRange(s.getLastRow() + 1, 1, out.length, 5).setValues(out);
  } finally { lock.releaseLock(); }
  return ContentService.createTextOutput("ok");
}

// 본인 기존 체크 불러오기(본인 줄만 돌려준다)
function doGet(e) {
  const s = sheet_(), last = s.getLastRow(), ans = {};
  if (last > 1) s.getRange(2, 1, last - 1, 5).getValues().forEach(r => { if (r[1] === e.parameter.me) ans[r[2]] = r[3]; });
  return ContentService.createTextOutput(JSON.stringify({ ans })).setMimeType(ContentService.MimeType.JSON);
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
