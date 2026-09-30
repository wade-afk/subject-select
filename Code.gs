/**
 * 과목 선택 수합 스크립트 (학교별 구글 시트에 붙여 쓰는 앱스크립트)
 * ----------------------------------------------------------------------
 * 과목 선택 페이지에서 학생이 [제출]을 누르면 이 시트에 한 줄씩 기록합니다.
 *   - '1학년' · '2학년' 탭 : 현재 학년별 제출 결과. 학생 1명 = 1줄, 고른 과목 칸에 1
 *                            (1행 학기 / 2행 선택 묶음 / 3행 과목명, 4행부터 학생)
 *   - '집계' 탭            : 과목별 신청 인원 (제출될 때마다 다시 계산)
 * 같은 학번이 다시 제출하면 그 줄을 덮어씁니다.
 *
 * 설치: 구글 시트 → 확장 프로그램 → Apps Script → 이 코드 붙여넣기 → 저장
 *       → 배포 → 새 배포 → 유형: 웹 앱 / 실행: 나 / 액세스 권한: 모든 사용자 → 웹 앱 URL 복사
 * 코드를 고친 뒤에는 [배포 관리 → ✏ → 버전: 새 버전 → 배포]를 해야 반영됩니다.
 */

const VERSION = 1;
const FIXED = ["제출시각", "학번", "반", "번호", "이름", "선택과목"];
const HEAD = 3;          // 머리글 행 수 (학기 / 선택 묶음 / 과목명)
const SEP = "\u0001";
const SUMMARY = "집계";

function onOpen(){
  SpreadsheetApp.getUi().createMenu("📋 과목 선택")
    .addItem("집계 새로고침", "refreshSummary")
    .addToUi();
}

// ===================== 웹 앱 진입점 =====================
function doGet(){
  return json_({ ok:true, service:"subject-select", version:VERSION, sheet:SpreadsheetApp.getActiveSpreadsheet().getName() });
}

function doPost(e){
  let body;
  try{ body = JSON.parse(e.postData.contents); }catch(err){ return json_({ ok:false, error:"bad json" }); }
  const lock = LockService.getScriptLock();
  try{ lock.waitLock(25000); }catch(err){ return json_({ ok:false, error:"접속이 많아요. 잠시 후 다시 제출해 주세요." }); }
  try{ return json_(record_(body)); }
  catch(err){ return json_({ ok:false, error:String(err) }); }
  finally{ lock.releaseLock(); }
}

// ===================== 제출 기록 =====================
function record_(b){
  const grade = clean_(b.grade), sid = clean_(b.studentId), name = clean_(b.name);
  if(!/^[1-3]$/.test(grade)) return { ok:false, error:"학년 정보가 없어요." };
  if(!sid || !name) return { ok:false, error:"학번과 이름이 필요해요." };
  const structure = Array.isArray(b.structure) ? b.structure : [];
  const selected = Array.isArray(b.selected) ? b.selected : [];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = gradeSheet_(ss, grade + "학년");
  const cols = ensureColumns_(sh, structure);

  const labels = {};   // 학기+묶음 → 2행 머리글
  structure.forEach(g => { labels[clean_(g.sem) + SEP + clean_(g.group)] = label_(g); });
  const row = new Array(sh.getLastColumn()).fill("");
  row[0] = new Date(); row[1] = sid; row[2] = clean_(b.classNo); row[3] = clean_(b.number); row[4] = name;
  const bySem = {};    // 학기 → 고른 과목 (선택과목 칸에 "2학년 1학기: 물리학, 화학 / 2학년 2학기: …"로 요약)
  selected.forEach(s => {
    const sem = clean_(s.sem), subject = clean_(s.subject);
    (bySem[sem] = bySem[sem] || []).push(subject);
    const c = cols[[sem, labels[sem + SEP + clean_(s.group)], subject].join(SEP)];
    if(c) row[c - 1] = 1;
  });
  row[5] = Object.keys(bySem).map(sem => sem + ": " + bySem[sem].join(", ")).join(" / ");

  // 같은 학번이 이미 있으면 그 줄을 덮어씀
  const last = Math.max(sh.getLastRow(), HEAD);
  let target = last + 1;
  if(last > HEAD){
    const ids = sh.getRange(HEAD + 1, 2, last - HEAD, 1).getValues();
    for(let i = 0; i < ids.length; i++) if(String(ids[i][0]).trim() === sid){ target = HEAD + 1 + i; break; }
  }
  sh.getRange(target, 1, 1, row.length).setValues([row]);
  updateSummary_(ss);
  return { ok:true, updated: target <= last };
}

/** 학년 탭을 찾거나 만듭니다. 새 스프레드시트의 빈 첫 시트는 그대로 이름만 바꿔 씁니다. */
function gradeSheet_(ss, name){
  let sh = ss.getSheetByName(name);
  if(sh) return sh;
  const sheets = ss.getSheets();
  sh = (sheets.length === 1 && sheets[0].getLastRow() === 0) ? sheets[0].setName(name) : ss.insertSheet(name);
  sh.getRange(HEAD, 1, 1, FIXED.length).setValues([FIXED]);
  sh.getRange(1, 1, HEAD, FIXED.length).setFontWeight("bold").setBackground("#EEF2FA");
  sh.setFrozenRows(HEAD);
  sh.setFrozenColumns(5);
  return sh;
}

/** 머리글(1~3행)에 없는 과목 열을 추가하고, "학기|묶음|과목" → 열 번호 표를 돌려줍니다. */
function ensureColumns_(sh, structure){
  const lastCol = Math.max(sh.getLastColumn(), FIXED.length);
  const head = sh.getRange(1, 1, HEAD, lastCol).getValues();
  const cols = {};
  for(let c = FIXED.length; c < lastCol; c++) cols[[head[0][c], head[1][c], head[2][c]].map(String).join(SEP)] = c + 1;
  const add = [[], [], []];
  structure.forEach(g => {
    const sem = clean_(g.sem), label = label_(g);
    (Array.isArray(g.subjects) ? g.subjects : []).forEach(s => {
      const key = [sem, label, clean_(s)].join(SEP);
      if(cols[key]) return;
      cols[key] = lastCol + add[0].length + 1;
      add[0].push(sem); add[1].push(label); add[2].push(clean_(s));
    });
  });
  if(add[0].length) sh.getRange(1, lastCol + 1, HEAD, add[0].length).setValues(add).setFontWeight("bold").setBackground("#EEF2FA");
  return cols;
}

function label_(g){
  return clean_(g.group) + " · 택" + (parseInt(g.pick, 10) || 1) + (g.optional ? " (희망)" : "");
}

// ===================== 집계 =====================
function refreshSummary(){ updateSummary_(SpreadsheetApp.getActiveSpreadsheet()); }

/** '집계' 탭을 다시 씁니다: 학년 탭마다 과목 열의 1을 세어 과목별 신청 인원을 기록. */
function updateSummary_(ss){
  const out = [["학년", "학기", "선택 묶음", "과목", "신청 인원"]];
  ss.getSheets().forEach(sh => {
    const name = sh.getName();
    if(!/^[1-3]학년$/.test(name)) return;
    const lastCol = sh.getLastColumn();
    if(lastCol <= FIXED.length) return;
    const v = sh.getRange(1, 1, Math.max(sh.getLastRow(), HEAD), lastCol).getValues();
    for(let c = FIXED.length; c < lastCol; c++){
      let n = 0;
      for(let r = HEAD; r < v.length; r++) if(Number(v[r][c]) === 1) n++;
      out.push([name, v[0][c], v[1][c], v[2][c], n]);
    }
  });
  const agg = ss.getSheetByName(SUMMARY) || ss.insertSheet(SUMMARY);
  agg.clearContents();
  agg.getRange(1, 1, out.length, 5).setValues(out);
  agg.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#EEF2FA");
  agg.setFrozenRows(1);
}

// ===================== 유틸 =====================
/** 한 줄 문자열로 정리. =,+,-,@ 로 시작하면 수식으로 실행되지 않게 ' 를 붙입니다. */
function clean_(v){
  return v == null ? "" : String(v).replace(/[\r\n\t]/g, " ").trim().slice(0, 200).replace(/^[=+\-@]/, "'$&");
}
function json_(obj){ return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
