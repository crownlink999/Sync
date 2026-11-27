// SVODKI: приёмник файлов: ПК -> этот скрипт (script.google.com) -> GitHub, ветка inbox -> Actions собирает данные -> сайт
// Настройки: Project Settings -> Script properties:
//   GH_TOKEN = токен GitHub (fine-grained; доступ к репозиторию crownlink999/Sync, Contents: Read and write)
//   KEY      = длинный пароль (тот же, что во 2-й строке send.txt)
//   REPO     = crownlink999/Sync   (по умолчанию)
//   BRANCH   = inbox               (по умолчанию)
// Ветка inbox каждый раз пересоздаётся ОДНИМ коммитом без истории (старые версии больших xlsx не копятся в репозитории).
// Файл лежит как input/<slot>/<имя>; новый файл того же slot заменяет прежний. Большие файлы приходят кусками name.part001of002.
function doGet(e) {
  // Проверка из браузера: ...exec?check=ВАШ_KEY  -> покажет, принимает ли GitHub токен и виден ли репозиторий
  var p = PropertiesService.getScriptProperties();
  if (e && e.parameter && e.parameter.check && e.parameter.check === p.getProperty('KEY')) {
    var res = {repo: p.getProperty('REPO') || 'crownlink999/Sync', branch: p.getProperty('BRANCH') || 'inbox'};
    var r = UrlFetchApp.fetch('https://api.github.com/repos/' + res.repo, {muteHttpExceptions: true,
      headers: {Authorization: 'Bearer ' + (p.getProperty('GH_TOKEN') || '').trim(), Accept: 'application/vnd.github+json'}});
    res.githubStatus = r.getResponseCode();
    try { var j = JSON.parse(r.getContentText()); res.message = j.message || 'ok'; res.canPush = j.permissions ? j.permissions.push : null; } catch (x) {}
    return ContentService.createTextOutput(JSON.stringify(res, null, 1)).setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput('SVODKI ok');
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(120000);
    var p = PropertiesService.getScriptProperties();
    var body = JSON.parse(e.postData.contents);
    if (!p.getProperty('KEY') || body.key !== p.getProperty('KEY')) return out({ok: false, error: 'bad key'});
    var repo = p.getProperty('REPO') || 'crownlink999/Sync';
    var branch = p.getProperty('BRANCH') || 'inbox';
    var files = body.files || [];
    if (!files.length) return out({ok: false, error: 'no files'});
    var base = 'https://api.github.com/repos/' + repo;

    var ref = gh(base + '/git/ref/heads/' + branch, 'get', null, true);   // 404 = ветки ещё нет
    var baseTree = null, existing = [];
    if (ref) {
      baseTree = gh(base + '/git/commits/' + ref.object.sha).tree.sha;
      existing = gh(base + '/git/trees/' + baseTree + '?recursive=1').tree.filter(function (t) { return t.type === 'blob'; });
    }
    var entries = [];
    files.forEach(function (f) {
      var slot = safe(f.slot || 'misc'), name = safe(f.name);
      var parts = f.parts || 1, part = f.part || 1;
      var suffix = parts > 1 ? '.part' + pad3(part) + 'of' + pad3(parts) : '';
      var path = 'input/' + slot + '/' + name + suffix;
      var blob = gh(base + '/git/blobs', 'post', {content: f.b64, encoding: 'base64'});
      entries.push({path: path, mode: '100644', type: 'blob', sha: blob.sha});
      // убрать прежние файлы этого slot: другое имя или другая нарезка на части
      existing.forEach(function (t) {
        if (t.path.indexOf('input/' + slot + '/') !== 0 || t.path === path) return;
        var b = t.path.replace(/\.part\d+of\d+$/, '');
        var m = t.path.match(/\.part\d+of(\d+)$/);
        var tot = m ? parseInt(m[1], 10) : 1;
        if (b !== 'input/' + slot + '/' + name || tot !== parts) entries.push({path: t.path, mode: '100644', type: 'blob', sha: null});
      });
    });
    var tp = {tree: entries};
    if (baseTree) tp.base_tree = baseTree;
    var nt = gh(base + '/git/trees', 'post', tp);
    var c = gh(base + '/git/commits', 'post', {message: 'Файлы с рабочего ПК: ' + files.length + ' шт.', tree: nt.sha, parents: []});
    if (ref) gh(base + '/git/refs/heads/' + branch, 'patch', {sha: c.sha, force: true});
    else gh(base + '/git/refs', 'post', {ref: 'refs/heads/' + branch, sha: c.sha});
    return out({ok: true, files: files.length, commit: c.sha});
  } catch (err) {
    return out({ok: false, error: String(err)});
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function safe(s) { return String(s).replace(/[\\\/]/g, '_'); }
function pad3(n) { return ('000' + n).slice(-3); }

function gh(url, method, payload, allow404) {
  var opt = {method: method || 'get', muteHttpExceptions: true, contentType: 'application/json',
    headers: {Authorization: 'Bearer ' + String(PropertiesService.getScriptProperties().getProperty('GH_TOKEN') || '').trim(),
              Accept: 'application/vnd.github+json'}};
  if (payload) opt.payload = JSON.stringify(payload);
  var r = UrlFetchApp.fetch(url, opt);
  var code = r.getResponseCode();
  if (allow404 && code === 404) return null;
  if (code >= 300) throw new Error('GitHub ' + code + ': ' + r.getContentText().slice(0, 300));
  return JSON.parse(r.getContentText());
}
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
