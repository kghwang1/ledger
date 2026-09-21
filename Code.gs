/**********************************************************************
 * 교회 부서 회계장부 · 가계부 — 자료 보관 서버  (2026-09 · 드라이브 목록 추가)
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
    else if (req.action === 'upfile') out = upfile_(folder, req);
    else if (req.action === 'delfile') out = delfile_(req);
    else if (req.action === 'drive') out = drive_(folder, req);
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

/* ── 첨부 문서(PDF 등) ── */

function attachFolder_(folder) {
  var it = folder.getFoldersByName('첨부');
  return it.hasNext() ? it.next() : folder.createFolder('첨부');
}

function upfile_(folder, req) {
  if (!req.data) throw new Error('올릴 파일이 없습니다');
  var m = String(req.data).match(/^data:([^;]+);base64,([\s\S]*)$/);
  if (!m) throw new Error('파일 형식을 알 수 없습니다');

  var name = clean_(req.name || '문서.pdf');
  var blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], name);
  var f = attachFolder_(folder).createFile(blob);

  try {                                  // 주소를 아는 사람은 볼 수 있게 (미리보기용)
    f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {}

  return { ok: true, id: f.getId(), name: f.getName(), size: f.getSize() };
}

function delfile_(req) {
  if (!req.id) throw new Error('지울 파일을 알 수 없습니다');
  try { DriveApp.getFileById(req.id).setTrashed(true); } catch (e) {}
  return { ok: true };
}

/* ── 자료 탭 › 구글 드라이브 목록 ──
   root 가 있으면 그 폴더, 없으면 교회회계자료 › (구분)자료 폴더를 보여 준다 */
function drive_(folder, req) {
  var root;
  if (req.root) {
    root = DriveApp.getFolderById(req.root);
  } else {
    var nm = clean_(req.rootName || '자료');
    var it = folder.getFoldersByName(nm);
    root = it.hasNext() ? it.next() : folder.createFolder(nm);
  }
  var cur = req.id ? DriveApp.getFolderById(req.id) : root;

  var folders = [], files = [];
  var fi = cur.getFolders();
  while (fi.hasNext()) {
    var f = fi.next();
    if (f.isTrashed()) continue;
    folders.push({ id: f.getId(), name: f.getName(), at: f.getLastUpdated().getTime() });
  }
  var xi = cur.getFiles();
  while (xi.hasNext()) {
    var x = xi.next();
    if (x.isTrashed()) continue;
    files.push({ id: x.getId(), name: x.getName(), mime: x.getMimeType(),
                 size: x.getSize(), at: x.getLastUpdated().getTime(), url: x.getUrl() });
  }
  var byName = function (a, b) { return a.name.localeCompare(b.name, 'ko'); };
  folders.sort(byName); files.sort(byName);

  return { ok: true,
           root: { id: root.getId(), name: root.getName(), url: root.getUrl() },
           cur:  { id: cur.getId(),  name: cur.getName(),  url: cur.getUrl() },
           folders: folders, files: files };
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
