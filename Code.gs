/**********************************************************************
 * 교회 부서 회계장부 — 자료 보관 서버
 *
 * 고치는 방법
 *  1) script.google.com 에서 만들어 둔 프로젝트를 엽니다.
 *  2) 기존 코드를 모두 지우고 이 파일 내용을 붙여넣습니다.
 *  3) Ctrl+S 로 저장합니다.
 *  4) 배포 → 배포 관리 → 연필(수정) → 버전을 '새 버전'으로 → 배포
 *     ※ 저장만 하면 반영되지 않습니다. 반드시 새 버전으로 배포하세요.
 *     ※ 웹 앱 주소는 그대로 유지되므로 프로그램은 고치지 않아도 됩니다.
 **********************************************************************/

var PASSWORD = 'church2026';        // 회계장부 프로그램의 SERVER_KEY 와 같아야 합니다
var FOLDER   = '교회회계자료';       // 드라이브에 만들어질 폴더 이름

function doPost(e) {
  var out;
  try {
    var req = JSON.parse(e.postData.contents);
    if (String(req.key) !== String(PASSWORD)) throw new Error('암호가 맞지 않습니다');

    var folder = getFolder_();
    if      (req.action === 'save') out = save_(folder, req);
    else if (req.action === 'list') out = list_(folder);
    else if (req.action === 'get')  out = get_(folder, req);
    else if (req.action === 'ping') out = { ok: true, folder: FOLDER };
    else throw new Error('알 수 없는 요청입니다');
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return ContentService.createTextOutput('교회 회계장부 보관 서버가 동작 중입니다.');
}

/* ---------------- 내부 함수 ---------------- */

function getFolder_() {
  var it = DriveApp.getFoldersByName(FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER);
}

function clean_(name) {
  return String(name).replace(/[\/\\:*?"<>|]/g, '_');
}

function fileName_(name, year) {
  return clean_(name) + '_' + year + '.json';
}

/* 파일 이름 해석 — 뒤에 (t) (1) 같은 꼬리표가 붙어도 알아본다 */
function parseName_(fn) {
  var m = String(fn).match(/^(.*?)_(\d{1,4})\s*(?:\([^)]*\))?\s*\.json$/i);
  return m ? { name: m[1], year: Number(m[2]) } : null;
}

/* 이름·연도가 같은 파일을 찾는다.
   이름이 정확한 파일이 있으면 폴더를 뒤지지 않고 바로 쓴다(빠른 길). */
function findFiles_(folder, name, year) {
  var want = clean_(name), out = [];
  var exact = folder.getFilesByName(fileName_(name, year));   // 색인으로 바로 찾기
  while (exact.hasNext()) out.push(exact.next());
  if (out.length === 1) return out;                           // 하나뿐이면 그대로

  if (!out.length) {                                          // 꼬리표가 붙은 것까지 훑는다
    var it = folder.getFiles();
    while (it.hasNext()) {
      var f = it.next();
      var p = parseName_(f.getName());
      if (p && p.name === want && p.year === Number(year)) out.push(f);
    }
  }
  out.sort(function (a, b) { return b.getLastUpdated().getTime() - a.getLastUpdated().getTime(); });
  return out;
}

function save_(folder, req) {
  if (!req.name) throw new Error('부서명이 비어 있습니다');
  if (!req.data) throw new Error('저장할 자료가 없습니다');

  var now  = new Date().getTime();
  var body = JSON.stringify({ name: req.name, year: req.year, at: now, data: req.data });
  var fn   = fileName_(req.name, req.year);
  var list = findFiles_(folder, req.name, req.year);

  if (list.length) {
    list[0].setContent(body);          // 가장 최근 것에 덮어쓴다
    if (list[0].getName() !== fn) list[0].setName(fn);   // 꼬리표가 있으면 이름도 바로잡는다
    for (var i = 1; i < list.length; i++) list[i].setTrashed(true);   // 중복은 휴지통으로
  } else {
    folder.createFile(fn, body, MimeType.PLAIN_TEXT);
  }
  return { ok: true, at: now, file: fn };
}

function list_(folder) {
  var seen = {}, out = [], it = folder.getFiles();
  while (it.hasNext()) {
    var f = it.next();
    var p = parseName_(f.getName());
    if (!p) continue;
    if (p.name.charAt(0) === '_') continue;          // 사용자 명단 같은 내부 기록은 뺀다
    var k = p.name + '_' + p.year, t = f.getLastUpdated().getTime();
    if (!seen[k] || seen[k].at < t) seen[k] = { name: p.name, year: p.year, at: t };
  }
  for (var key in seen) out.push(seen[key]);
  out.sort(function (a, b) {
    if (a.name === b.name) return b.year - a.year;
    return a.name < b.name ? -1 : 1;
  });
  return { ok: true, list: out };
}

function get_(folder, req) {
  var list = findFiles_(folder, req.name, req.year);
  if (!list.length) throw new Error(fileName_(req.name, req.year) + ' 파일을 찾지 못했습니다');
  return { ok: true, rec: JSON.parse(list[0].getBlob().getDataAsString('UTF-8')) };
}
