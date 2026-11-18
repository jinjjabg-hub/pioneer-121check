/**
 * 121 체크 → 구글시트 취합 (Apps Script 웹앱)
 * 시트 1개당 한 번만 설정: 확장 프로그램 > Apps Script > 이 코드 붙여넣기 > 배포 > 웹 앱
 *   실행: 나 / 액세스: 모든 사용자 → 나온 주소를 index.html 의 SHEET_URL 에 붙여넣기
 * 시트 '체크' = 한 줄에 (나, 상대, 단계). 멤버가 저장할 때마다 그 사람 줄만 통째로 교체한다.
 * 시트 '매트릭스' = 메뉴 '121 체크 > 매트릭스 만들기'로 생성. 둘 중 한쪽이라도 체크하면 그 단계로 본다.
 *
 * 시트 '멤버' = 멤버 명단의 원본(이름·회사·업종·역할·번호). 멤버 추가·삭제는 여기 한 줄만 고치면 된다.
 *   - 페이지는 ?members=1 로 명단을 받아 간다(번호는 빼고 보낸다). 줄 순서가 화면 순서.
 *   - 로그인: 페이지가 이름(me)과 전화번호 뒤 4자리(pin)를 보내면 이 탭의 번호와 맞아야만 읽고 쓴다.
 *   - 121 일정(Firebase)용 번호는 이 탭이 바뀔 때 Firestore 문서 private/pins 로 자동 복사된다.
 *     필요한 것: 스크립트 속성 FIREBASE_SA = Firebase 서비스 계정 키(JSON 파일 내용 전체),
 *     메뉴 '121 체크 > 자동 보내기 켜기' 한 번.
 */
const NAMES = ["", "패스포트", "아는", "신뢰", "수익"];
const MEMBER_SHEET = "멤버";
const MEMBER_HEAD = ["이름", "회사", "업종", "역할(앰배서더 등)", "전화번호 뒤 4자리"];
const PAGE_URL = "https://jinjjabg-hub.github.io/pioneer-121check/";

// 멤버 탭 읽기: [{n, co, f, role, pin}] (이름 빈 줄은 건너뛴다)
function members_() {
  const s = SpreadsheetApp.getActive().getSheetByName(MEMBER_SHEET);
  if (!s || s.getLastRow() < 2) return [];
  return s.getRange(2, 1, s.getLastRow() - 1, 5).getDisplayValues()
    .map(r => ({ n: r[0].trim(), co: r[1].trim(), f: r[2].trim(), role: r[3].trim(), pin: r[4].trim() }))
    .filter(m => m.n);
}
function ok_(me, pin) { return !!me && /^[0-9]{4}$/.test(String(pin || "")) && members_().some(m => m.n === me && m.pin === String(pin)); }

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
  if (e.parameter.members) {   // 공개 명단(번호 제외). 시트 수정 후 최대 5분 안에 반영
    const c = CacheService.getScriptCache(); let out = c.get("members");
    if (!out) { out = JSON.stringify({ members: members_().map(m => m.role ? { n: m.n, co: m.co, f: m.f, role: m.role } : { n: m.n, co: m.co, f: m.f }) }); c.put("members", out, 300); }
    return ContentService.createTextOutput(out).setMimeType(ContentService.MimeType.JSON);
  }
  if (!ok_(e.parameter.me, e.parameter.pin)) return ContentService.createTextOutput(JSON.stringify({ error: "auth" })).setMimeType(ContentService.MimeType.JSON);
  const s = sheet_(), last = s.getLastRow(), ans = {}, st = {};
  if (last > 1) s.getRange(2, 1, last - 1, 5).getValues().forEach(r => { if (r[1] === e.parameter.me) { ans[r[2]] = r[3]; st[r[2]] = r[4]; } });
  return ContentService.createTextOutput(JSON.stringify({ ans, st })).setMimeType(ContentService.MimeType.JSON);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu("121 체크")
    .addItem("매트릭스 만들기", "matrix")
    .addSeparator()
    .addItem("멤버 탭 만들기 (처음 한 번)", "setupMembers")
    .addItem("번호를 Firebase로 보내기", "syncPinsNow")
    .addItem("자동 보내기 켜기 (처음 한 번)", "setupAuto")
    .addToUi();
}

/* ===== 멤버 탭 ===== */
// 처음 한 번: 지금 페이지에 있는 명단으로 '멤버' 탭을 만든다.
// 예전 코드의 PINS 목록을 이 파일 맨 아래에 붙여 두고 실행하면 번호도 같이 옮겨진다.
function setupMembers() {
  const ss = SpreadsheetApp.getActive(), ui = SpreadsheetApp.getUi();
  if (ss.getSheetByName(MEMBER_SHEET)) return ui.alert("'멤버' 탭이 이미 있어요. 그 탭을 직접 고쳐 주세요.");
  const html = UrlFetchApp.fetch(PAGE_URL).getContentText(), m = html.match(/\/\*MEMBERS\*\/([\s\S]*?)\/\*END\*\//);
  if (!m) return ui.alert("페이지에서 명단을 찾지 못했어요.");
  const old = (typeof PINS !== "undefined") ? PINS : {};
  const rows = JSON.parse(m[1]).map(x => [x.n, x.co || "", x.f || "", x.role || "", /^[0-9]{4}$/.test(old[x.n] || "") ? old[x.n] : ""]);
  const s = ss.insertSheet(MEMBER_SHEET);
  s.getRange(1, 1, 1, 5).setValues([MEMBER_HEAD]).setFontWeight("bold");
  s.getRange(2, 5, Math.max(rows.length, 1) + 200, 1).setNumberFormat("@");   // 0으로 시작하는 번호가 숫자로 바뀌지 않게
  if (rows.length) s.getRange(2, 1, rows.length, 5).setValues(rows);
  s.setFrozenRows(1); s.autoResizeColumns(1, 5);
  const miss = rows.filter(r => !r[4]).map(r => r[0]);
  ui.alert(`'멤버' 탭을 만들었어요 (${rows.length}명).` + (miss.length ? `\n번호가 비어 있는 사람: ${miss.join(", ")}\n직접 채워 주세요.` : ""));
}

/* ===== 번호 → Firebase (Firestore 문서 private/pins) ===== */
function fbToken_() {
  const raw = PropertiesService.getScriptProperties().getProperty("FIREBASE_SA");
  if (!raw) throw new Error("스크립트 속성 FIREBASE_SA(서비스 계정 키)가 없어요.");
  const sa = JSON.parse(raw), now = Math.floor(Date.now() / 1000);
  const b64 = o => Utilities.base64EncodeWebSafe(JSON.stringify(o)).replace(/=+$/, "");
  const input = b64({ alg: "RS256", typ: "JWT" }) + "." + b64({ iss: sa.client_email, scope: "https://www.googleapis.com/auth/datastore", aud: "https://oauth2.googleapis.com/token", iat: now, exp: now + 3600 });
  const sig = Utilities.base64EncodeWebSafe(Utilities.computeRsaSha256Signature(input, sa.private_key)).replace(/=+$/, "");
  const r = UrlFetchApp.fetch("https://oauth2.googleapis.com/token", { method: "post", muteHttpExceptions: true, payload: { grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer", assertion: input + "." + sig } });
  if (r.getResponseCode() !== 200) throw new Error("Firebase 인증 실패: " + r.getContentText().slice(0, 200));
  return { token: JSON.parse(r.getContentText()).access_token, project: sa.project_id };
}
// 멤버 탭의 이름·번호를 통째로 보낸다(탭에서 지운 사람은 Firebase에서도 빠진다). 결과는 멤버 탭 G1에 적는다.
function syncPins_() {
  const list = members_(), fields = {}, bad = [];
  list.forEach(m => { if (/^[0-9]{4}$/.test(m.pin)) fields[m.n] = { stringValue: m.pin }; else bad.push(m.n); });
  if (!Object.keys(fields).length) throw new Error("보낼 번호가 없어요. '멤버' 탭을 확인해 주세요.");
  const t = fbToken_();
  const r = UrlFetchApp.fetch(`https://firestore.googleapis.com/v1/projects/${t.project}/databases/(default)/documents/private/pins`, {
    method: "patch", contentType: "application/json", muteHttpExceptions: true,
    headers: { Authorization: "Bearer " + t.token }, payload: JSON.stringify({ fields })
  });
  if (r.getResponseCode() !== 200) throw new Error("Firebase 저장 실패: " + r.getContentText().slice(0, 200));
  return { n: Object.keys(fields).length, bad };
}
function note_(msg) { const s = SpreadsheetApp.getActive().getSheetByName(MEMBER_SHEET); if (s) s.getRange("G1").setValue(msg); }
function stamp_() { return Utilities.formatDate(new Date(), "Asia/Seoul", "M/d HH:mm"); }
function syncAndNote_() {
  CacheService.getScriptCache().remove("members");
  try { const r = syncPins_(); note_(`✓ ${stamp_()} Firebase에 ${r.n}명 보냄` + (r.bad.length ? ` · 번호 없음/형식 오류: ${r.bad.join(", ")}` : "")); return r; }
  catch (e) { note_(`✗ ${stamp_()} ${e.message}`); throw e; }
}
function syncPinsNow() {
  const r = syncAndNote_();
  SpreadsheetApp.getUi().alert(`Firebase에 ${r.n}명의 번호를 보냈어요.` + (r.bad.length ? `\n번호가 없거나 4자리가 아닌 사람(로그인 안 됨): ${r.bad.join(", ")}` : ""));
}
// 멤버 탭을 고치면(onEdit) 또는 줄을 넣고 지우면(onChange) 자동으로 보낸다. 칸 수정은 onEdit만 처리해 두 번 보내지 않는다
function onMemberChange(e) {
  if (e && e.range && e.range.getSheet().getName() !== MEMBER_SHEET) return;
  if (e && e.changeType && !/INSERT_ROW|REMOVE_ROW|OTHER/.test(e.changeType)) return;
  try { syncAndNote_(); } catch (err) {}
}
function setupAuto() {
  const ss = SpreadsheetApp.getActive();
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === "onMemberChange").forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("onMemberChange").forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger("onMemberChange").forSpreadsheet(ss).onChange().create();
  SpreadsheetApp.getUi().alert("자동 보내기를 켰어요. 이제 '멤버' 탭을 고치면 Firebase에도 바로 반영돼요.");
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
