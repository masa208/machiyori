/**
 * まち頼 サーバー（GAS）v2  — 2026/11 本番化 第1段階
 * キーやトークンはコードに書かない。すべて「プロジェクトの設定 → スクリプト プロパティ」に入れる。
 *   OPENAI_KEY            … OpenAIのAPIキー（必須）
 *   OPENAI_MODEL          … 使うモデル名（任意。未設定なら gpt-4o-mini）
 *   LINE_LOGIN_CHANNEL_ID … LINEログインのチャネルID（任意。未設定なら LIFF ID の先頭の数字）
 *   APP_URL               … アプリのURL（パスワード再設定メールのリンクに使う）
 *   TEST_MODE             … on にするとテスト用の業者操作が使える（本番公開前に off にする）
 */
var SHEET_ID = '1Ba2Nq-mBTQSuXf5Ae0wXCTQqBluTv8BN4akILmHOJlw';
var LIFF_ID = '2011196111-mkgGx8sO';

/* ========== 旧アプリ（kensetsu_kanri.html）との互換 ========== */
var HEADERS = ['id','lineUserId','name','company','role','categories','venue','areas','contact','bio','status','createdAt'];
function sheet() { return SpreadsheetApp.openById(SHEET_ID).getSheetByName('members'); }

function doGet(e) {
  var p = (e && e.parameter) || {};
  var out;
  try {
    if (p.action === 'register') out = registerMember(p);
    else if (p.action === 'update') out = updateMember(p);
    else if (p.action === 'notify') out = { status: 'ok' };            // 公式LINEは使わないため何もしない
    else if (p.action === 'classify') out = classify(p.text);
    else if (p.action === 'api' && p.payload) out = api(JSON.parse(p.payload)); // 新アプリのJSONP予備
    else out = { status: 'ok', members: listMembers() };
  } catch (err) { out = { status: 'error', message: String(err) }; }
  var json = JSON.stringify(out);
  if (p.callback) return ContentService.createTextOutput(p.callback + '(' + json + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

/* 新アプリはここ（fetch POST・text/plain）で呼ぶ */
function doPost(e) {
  var out;
  try {
    var body = JSON.parse(e.postData.contents);
    if (body && body.events) return ContentService.createTextOutput('ok'); // 旧LINE Webhook（無視）
    out = api(body);
  } catch (err) { out = { ok: false, error: String(err && err.message || err) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function listMembers() {
  var rows = sheet().getDataRange().getValues(); var head = rows.shift();
  return rows.filter(function (r) { return r[0]; }).map(function (r) { var o = {}; head.forEach(function (h, i) { o[h] = r[i]; }); return o; });
}
function registerMember(p) {
  var sh = sheet(), rows = sh.getDataRange().getValues();
  var vals = [p.name || '', p.company || '', p.role || 'customer', p.categories || '', p.venue || '', p.areas || '', p.contact || '', p.bio || ''];
  for (var i = 1; i < rows.length; i++) if (rows[i][1] === p.lineUserId) { sh.getRange(i + 1, 3, 1, 8).setValues([vals]); return { status: 'ok', id: rows[i][0], updated: true }; }
  var id = 'M' + Date.now(), st = (p.role === 'contractor') ? 'pending' : 'active';
  sh.appendRow([id, p.lineUserId].concat(vals, [st, new Date().toISOString()]));
  return { status: 'ok', id: id, memberStatus: st };
}
function updateMember(p) {
  var sh = sheet(), rows = sh.getDataRange().getValues(), col = HEADERS.indexOf('status') + 1;
  for (var i = 1; i < rows.length; i++) if (rows[i][0] === p.id) { if (p.memberStatus) sh.getRange(i + 1, col).setValue(p.memberStatus); return { status: 'ok' }; }
  return { status: 'error', message: 'not found' };
}

/* ========== 設定・共通 ========== */
function prop(k, d) { var v = PropertiesService.getScriptProperties().getProperty(k); return v == null || v === '' ? d : v; }
function ss() { return SpreadsheetApp.openById(SHEET_ID); }
function now() { return new Date().toISOString(); }
function uid(p) { return (p || '') + Utilities.getUuid().replace(/-/g, '').slice(0, 16); }

var TABLES = {
  users: ['id','provider','lineUserId','email','passHash','salt','name','callName','area','role','chara','points','memo','status','createdAt'],
  sessions: ['token','userId','createdAt','lastSeen'],
  resets: ['token','userId','expires','used'],
  cases: ['id','userId','title','bunya','what','place','building','when','detail','ask','estOnly','photos','status','stopBids','postponeAt','closedReason','createdAt','updatedAt'],
  bids: ['id','caseId','vendorId','state','firstMsgAt','fee','helpers','createdAt','updatedAt','readAt','amount'],
  messages: ['id','caseId','bidId','from','type','body','createdAt'],
  notifications: ['id','userId','text','link','read','createdAt'],
  reviews: ['id','caseId','bidId','vendorId','userId','stars','text','createdAt'],
  inquiries: ['id','userId','type','body','reply','status','createdAt','repliedAt'],
  concierge: ['id','userId','role','text','createdAt'],
  vendors: ['id','name','category','area','star','reviews','googleUrl','isTest','createdAt','userId']
};

/** 初回だけ手動で実行：新しいシートを作る */
function setupV2() {
  var book = ss();
  Object.keys(TABLES).forEach(function (n) {
    var sh = book.getSheetByName(n);
    if (!sh) { sh = book.insertSheet(n); sh.appendRow(TABLES[n]); sh.setFrozenRows(1); }
  });
  installTriggers();
  Logger.log('setupV2 完了');
}
function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'hourlyJobs') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('hourlyJobs').timeBased().everyHours(1).create();
}

/* シート読み書き（小さなORM） */
function tbl(n) { return ss().getSheetByName(n); }
function all(n) {
  var sh = tbl(n), v = sh.getDataRange().getValues(), h = v.shift();
  return v.filter(function (r) { return r[0] !== ''; }).map(function (r, i) {
    var o = { _row: i + 2 }; h.forEach(function (k, j) { var x = r[j]; o[k] = (x instanceof Date) ? x.toISOString() : x; }); return o;
  });
}
function find(n, f) { var a = all(n); for (var i = 0; i < a.length; i++) if (f(a[i])) return a[i]; return null; }
function insert(n, o) { var h = TABLES[n]; tbl(n).appendRow(h.map(function (k) { return o[k] === undefined ? '' : o[k]; })); return o; }
function update(n, o, patch) {
  var h = TABLES[n], sh = tbl(n);
  Object.keys(patch).forEach(function (k) { var c = h.indexOf(k); if (c >= 0) { sh.getRange(o._row, c + 1).setValue(patch[k]); o[k] = patch[k]; } });
  return o;
}
function pub(o) { var r = {}; Object.keys(o).forEach(function (k) { if (k !== '_row' && k !== 'passHash' && k !== 'salt') r[k] = o[k]; }); return r; }
function J(s, d) { try { return s ? JSON.parse(s) : d; } catch (e) { return d; } }

/* ========== 入口 ========== */
function api(b) {
  var a = b.a, u = null;
  var open = { loginLine: 1, loginEmail: 1, registerEmail: 1, resetRequest: 1, resetDo: 1 };
  if (!open[a]) { u = auth(b.token); if (!u) return { ok: false, error: 'auth' }; }
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    migrateV3();
    switch (a) {
      case 'loginLine': return loginLine(b);
      case 'registerEmail': return registerEmail(b);
      case 'loginEmail': return loginEmail(b);
      case 'resetRequest': return resetRequest(b);
      case 'resetDo': return resetDo(b);
      case 'me': return { ok: true, user: pub(u) };
      case 'saveProfile': return saveProfile(u, b);
      case 'logout': return logout(b.token);
      case 'withdraw': return withdraw(u);
      case 'home': return home(u);
      case 'concierge': return concierge(u, b);
      case 'uploadPhoto': return uploadPhoto(u, b);
      case 'postCase': return postCase(u, b);
      case 'getCase': return getCase(u, b.caseId);
      case 'caseAction': return caseAction(u, b);
      case 'getChat': return getChat(u, b.bidId, b.since);
      case 'sendMsg': return sendMsg(u, b);
      case 'cardAction': return cardAction(u, b);
      case 'postReview': return postReview(u, b);
      case 'notifs': return notifs(u, b);
      case 'inquiry': return inquiry(u, b);
      case 'inquiries': return { ok: true, items: all('inquiries').filter(function (q) { return q.userId === u.id; }).reverse().map(pub) };
      case 'test': return testOp(u, b);
    }
    var pr = proApi(u, b); if (pr) return pr;
    return { ok: false, error: 'unknown action' };
  } finally { lock.releaseLock(); }
}

/* ========== ログイン ========== */
function newSession(userId) { var t = uid('s'); insert('sessions', { token: t, userId: userId, createdAt: now(), lastSeen: now() }); return t; }
function auth(token) {
  if (!token) return null;
  var s = find('sessions', function (x) { return x.token === token; }); if (!s) return null;
  var u = find('users', function (x) { return x.id === s.userId; });
  if (!u || u.status === 'withdrawn' || u.status === 'stopped') return null;
  return u;
}
function loginLine(b) {
  var cid = prop('LINE_LOGIN_CHANNEL_ID', LIFF_ID.split('-')[0]);
  var info;
  if (b.idToken) {
    var res = UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/verify', { method: 'post', payload: { id_token: b.idToken, client_id: cid }, muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return { ok: false, error: 'LINEの確認に失敗しました' };
    info = JSON.parse(res.getContentText()); // sub = LINEユーザーID
  } else if (b.accessToken) { // openid権限がない場合：アクセストークンを確かめてからプロフィールを取る
    var v = UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/verify?access_token=' + encodeURIComponent(b.accessToken), { muteHttpExceptions: true });
    if (v.getResponseCode() !== 200 || String(JSON.parse(v.getContentText()).client_id) !== String(cid)) return { ok: false, error: 'LINEの確認に失敗しました' };
    var pr = UrlFetchApp.fetch('https://api.line.me/v2/profile', { headers: { Authorization: 'Bearer ' + b.accessToken }, muteHttpExceptions: true });
    if (pr.getResponseCode() !== 200) return { ok: false, error: 'LINEの確認に失敗しました' };
    var p = JSON.parse(pr.getContentText()); info = { sub: p.userId, name: p.displayName };
  } else return { ok: false, error: 'LINEの確認に失敗しました' };
  var u = find('users', function (x) { return x.lineUserId === info.sub && x.status !== 'withdrawn'; });
  if (!u) u = insert('users', { id: uid('U'), provider: 'line', lineUserId: info.sub, name: info.name || '', callName: '', role: 'customer', chara: 'sis', points: 0, memo: '', status: 'active', createdAt: now() });
  return { ok: true, token: newSession(u.id), user: pub(u) };
}
function hashPw(pw, salt) {
  var h = pw + salt;
  for (var i = 0; i < 200; i++) h = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + salt, Utilities.Charset.UTF_8));
  return h;
}
function registerEmail(b) {
  var email = String(b.email || '').trim().toLowerCase();
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return { ok: false, error: 'メールアドレスの形が正しくありません' };
  if (String(b.password || '').length < 8) return { ok: false, error: 'パスワードは8文字以上にしてください' };
  if (find('users', function (x) { return x.email === email && x.status !== 'withdrawn'; })) return { ok: false, error: 'このメールアドレスはすでに登録されています' };
  var salt = uid('');
  var u = insert('users', { id: uid('U'), provider: 'email', email: email, passHash: hashPw(b.password, salt), salt: salt, name: '', callName: '', role: 'customer', chara: 'sis', points: 0, memo: '', status: 'active', createdAt: now() });
  return { ok: true, token: newSession(u.id), user: pub(u) };
}
function loginEmail(b) {
  var email = String(b.email || '').trim().toLowerCase();
  var u = find('users', function (x) { return x.email === email && x.status !== 'withdrawn'; });
  if (!u || hashPw(String(b.password || ''), u.salt) !== u.passHash) return { ok: false, error: 'メールアドレスかパスワードが違います' };
  return { ok: true, token: newSession(u.id), user: pub(u) };
}
function resetRequest(b) {
  var email = String(b.email || '').trim().toLowerCase();
  var u = find('users', function (x) { return x.email === email && x.status !== 'withdrawn'; });
  if (u) {
    var t = uid('r'); insert('resets', { token: t, userId: u.id, expires: new Date(Date.now() + 3600e3).toISOString(), used: '' });
    MailApp.sendEmail(email, '【まち頼】パスワードの再設定', 'まち頼のパスワードを再設定するには、1時間以内に次のリンクを開いてください。\n\n' + prop('APP_URL', '') + '?reset=' + t + '\n\nお心当たりがない場合は、このメールを無視してください。');
  }
  return { ok: true }; // 登録の有無は返さない
}
function resetDo(b) {
  var r = find('resets', function (x) { return x.token === b.resetToken; });
  if (!r || r.used || new Date(r.expires) < new Date()) return { ok: false, error: 'リンクの有効期限が切れています。もう一度お試しください' };
  if (String(b.password || '').length < 8) return { ok: false, error: 'パスワードは8文字以上にしてください' };
  var u = find('users', function (x) { return x.id === r.userId; }); var salt = uid('');
  update('users', u, { passHash: hashPw(b.password, salt), salt: salt }); update('resets', r, { used: now() });
  return { ok: true, token: newSession(u.id), user: pub(u) };
}
function logout(token) { var s = find('sessions', function (x) { return x.token === token; }); if (s) tbl('sessions').deleteRow(s._row); return { ok: true }; }
function saveProfile(u, b) {
  var p = {}; ['callName', 'area', 'chara', 'name'].forEach(function (k) { if (b[k] !== undefined) p[k] = String(b[k]).slice(0, 40); });
  update('users', u, p); return { ok: true, user: pub(u) };
}
function withdraw(u) {
  var active = all('cases').filter(function (c) { return c.userId === u.id && ['open', 'postponed', 'contracted'].indexOf(c.status) >= 0; });
  if (active.length) return { ok: false, error: '進行中の依頼があります。完了するか取り下げてから退会してください' };
  update('users', u, { status: 'withdrawn' }); return { ok: true };
}

/* ========== 通知BOX ========== */
function notify(userId, text, link) { userId = ownerOf(userId); if (userId) insert('notifications', { id: uid('N'), userId: userId, text: text, link: link || '', read: '', createdAt: now() }); }
function notifs(u, b) {
  var a = all('notifications').filter(function (n) { return n.userId === u.id; });
  if (b.markRead) a.forEach(function (n) { if (!n.read) update('notifications', n, { read: now() }); });
  return { ok: true, items: a.slice(-50).reverse().map(pub) };
}

/* ========== ホーム ========== */
function vendorMap() { var m = {}; all('vendors').forEach(function (v) { m[v.id] = pub(v); }); return m; }
function home(u) {
  var cs = all('cases').filter(function (c) { return c.userId === u.id && c.status !== 'withdrawn'; });
  var bids = all('bids'), vm = vendorMap();
  var list = cs.map(function (c) {
    var bs = bids.filter(function (x) { return x.caseId === c.id; });
    return { c: pub(c), bids: bs.map(function (x) { var o = pub(x); o.vendor = vm[x.vendorId] || null; return o; }) };
  }).reverse();
  var unread = all('notifications').filter(function (n) { return n.userId === u.id && !n.read; }).length;
  return { ok: true, user: pub(u), cases: list, unread: unread };
}

/* ========== AIコンシェルジュ ========== */
function concierge(u, b) {
  var key = prop('OPENAI_KEY', ''); if (!key) return { ok: false, error: 'OPENAI_KEY が設定されていません' };
  if (b.reset) { /* 新しい話題：履歴は残し、区切りだけ入れる */ insert('concierge', { id: uid('C'), userId: u.id, role: 'divider', text: '', createdAt: now() }); if (b.opening) insert('concierge', { id: uid('C'), userId: u.id, role: 'assistant', text: String(b.opening).slice(0, 300), createdAt: now() }); }
  var hist = all('concierge').filter(function (m) { return m.userId === u.id; });
  var start = 0; hist.forEach(function (m, i) { if (m.role === 'divider') start = i + 1; });
  hist = hist.slice(start).slice(-30);
  if (b.text) insert('concierge', { id: uid('C'), userId: u.id, role: 'user', text: String(b.text).slice(0, 2000), createdAt: now() });
  var ctx = 'キャラクター：' + (u.chara === 'bro' ? '大工のお兄さん' : '建築士のお姉さん') + '／おなじみ度：' + (u.points >= 30 ? '頼れる存在' : u.points >= 10 ? '話し相手' : '顔見知り') +
    '／名前：' + (u.callName || 'なし') + '／覚えておくことメモ：' + (u.memo || 'なし') + '／今日：' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy年M月d日') + '／住まい：' + (u.area || '不明');
  var msgs = [{ role: 'system', content: PERSONA }, { role: 'system', content: '【アプリから】' + ctx }];
  hist.forEach(function (m) { if (m.role === 'user' || m.role === 'assistant') msgs.push({ role: m.role, content: m.text }); });
  if (b.text) msgs.push({ role: 'user', content: String(b.text) });
  var res = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
    method: 'post', contentType: 'application/json', headers: { Authorization: 'Bearer ' + key },
    payload: JSON.stringify({ model: prop('OPENAI_MODEL', 'gpt-4o-mini'), messages: msgs, temperature: 0.7, max_tokens: 700 }), muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) return { ok: false, error: 'AIの応答に失敗しました（' + res.getResponseCode() + '）' };
  var txt = JSON.parse(res.getContentText()).choices[0].message.content || '';
  insert('concierge', { id: uid('C'), userId: u.id, role: 'assistant', text: txt, createdAt: now() });
  var sheetJ = null, m = txt.match(/\[\[SHEET\]\](\{.*\})/);
  if (m) { sheetJ = J(m[1], null); txt = txt.replace(m[0], '').trim(); }
  if (!u.points || u.points < 100) update('users', u, { points: (Number(u.points) || 0) + 1 });
  return { ok: true, reply: txt, sheet: sheetJ };
}
/* 旧アプリ用の分類（キーはプロパティから） */
function classify(text) {
  if (!text) return { status: 'error', message: 'no text' };
  var key = prop('OPENAI_KEY', '');
  var patterns = ['ame：雨漏り・水が落ちてくる','shimi：壁や天井のシミ・変色・カビ','mizu：水が出ない・水漏れ・給湯器','tsumari：排水のつまり・流れが悪い','yuka：床が沈む・きしむ・傾き','tobira：ドア・窓・ふすまの建て付け','nioi：変なにおい・異臭','oto：壁裏や天井裏の音・異音','denki：電気がつかない・コンセント・エアコン','niwa：庭・外構・駐車場・草・水たまり','kabe：屋根や外壁のひび・剥がれ・塗装','reform：リフォーム・増築・新築・間取り変更','souji：掃除・片付け・不用品処分・解体','dannetsu：寒い・暑い・結露・すきま風・断熱'];
  var prompt = '住まいの困りごとの相談文を読んで、下の一覧から最も近いものを1つ選び、IDだけを返してください。\nどれにも当てはまらない、または判断に迷う場合は none と返してください。\n説明や記号は不要です。IDのみを返してください。\n\n【一覧】\n' + patterns.join('\n') + '\n\n【相談文】\n' + text;
  var res = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', { method: 'post', contentType: 'application/json', headers: { Authorization: 'Bearer ' + key }, payload: JSON.stringify({ model: 'gpt-4o-mini', messages: [{ role: 'user', content: prompt }], temperature: 0, max_tokens: 10 }), muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) return { status: 'error', code: res.getResponseCode() };
  var id = (JSON.parse(res.getContentText()).choices[0].message.content || '').trim().toLowerCase().replace(/[^a-z]/g, '');
  return { status: 'ok', pattern: id || 'none' };
}

/* ========== 写真（Googleドライブ） ========== */
function photoFolder() {
  var id = prop('PHOTO_FOLDER_ID', '');
  if (id) return DriveApp.getFolderById(id);
  var f = DriveApp.createFolder('まち頼_写真'); PropertiesService.getScriptProperties().setProperty('PHOTO_FOLDER_ID', f.getId()); return f;
}
function uploadPhoto(u, b) {
  var m = String(b.dataUrl || '').match(/^data:(image\/\w+);base64,(.+)$/); if (!m) return { ok: false, error: '画像の形式が違います' };
  var blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], 'p_' + Date.now() + '.jpg');
  var file = photoFolder().createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { ok: true, url: 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1200' };
}

/* ========== 案件 ========== */
function postCase(u, b) {
  var s = b.sheet || {};
  if (!s.what && !s.detail) return { ok: false, error: '依頼したいことか、詳しい内容を入れてください' };
  var c = insert('cases', {
    id: uid('K'), userId: u.id, title: String(s.what || s.detail).slice(0, 30), bunya: s.bunya || 'その他・まず相談したい', what: s.what || '', place: s.place || u.area || '',
    building: s.building || '', when: s.when || '', detail: s.detail || '', ask: s.ask || '', estOnly: s.estOnly ? 'TRUE' : '', photos: JSON.stringify((b.photos || []).slice(0, 3)),
    status: 'open', stopBids: '', postponeAt: '', closedReason: '', createdAt: now(), updatedAt: now()
  });
  notify(u.id, '「' + c.title + '」を投稿しました。専門家からの返事をお待ちください', 'case:' + c.id);
  // 対応できる業者が0件の場合の表示（業者側の機能は第2段階で作る）
  if (!all('vendors').length) notify(u.id, '対応できる専門家を、運営が地元のつながりからお探しします。1営業日以内にご連絡します', 'case:' + c.id);
  return { ok: true, caseId: c.id };
}
function myCase(u, id) { var c = find('cases', function (x) { return x.id === id; }); return (c && c.userId === u.id) ? c : null; }
function getCase(u, id) {
  var c = myCase(u, id); if (!c) return { ok: false, error: '見つかりません' };
  var vm = vendorMap(), msgs = all('messages').filter(function (m) { return m.caseId === id; });
  var bs = all('bids').filter(function (x) { return x.caseId === id; }).map(function (x) {
    var o = pub(x); o.vendor = vm[x.vendorId] || null; var mm = msgs.filter(function (m) { return m.bidId === x.id; });
    o.last = mm.length ? pub(mm[mm.length - 1]) : null; if (o.last) o.last.body = J(o.last.body, {}); o.reviewed = !!find('reviews', function (q) { return q.bidId === x.id; }); o.unread = mm.filter(function (m) { return m.from !== u.id && m.createdAt > (x.readAt || ''); }).length; return o;
  });
  var o = pub(c); o.photos = J(c.photos, []);
  return { ok: true, c: o, bids: bs };
}
function caseAction(u, b) {
  var c = myCase(u, b.caseId); if (!c) return { ok: false, error: '見つかりません' };
  var bs = all('bids').filter(function (x) { return x.caseId === c.id && ['lost', 'declined'].indexOf(x.state) < 0; });
  if (b.op === 'stopBids') { update('cases', c, { stopBids: 'TRUE', updatedAt: now() }); }
  if (b.op === 'withdraw') {
    if (bs.some(function (x) { return x.state === 'won'; })) return { ok: false, error: '成約後は取り下げできません。業者にキャンセルを相談してください' };
    update('cases', c, { status: 'withdrawn', closedReason: b.reason || '', updatedAt: now() });
    bs.forEach(function (x) { sysMsg(x, '依頼者が依頼を取り下げました'); notify(x.vendorId, '「' + c.title + '」が取り下げられました'); });
  }
  if (b.op === 'postpone') {
    if (!b.at || new Date(b.at) <= new Date()) return { ok: false, error: '再開の日時は、これから先の日時を選んでください' };
    update('cases', c, { status: 'postponed', postponeAt: b.at, updatedAt: now() });
    bs.forEach(function (x) { sysMsg(x, '依頼者が先送りしました。再開予定：' + fmt(b.at)); notify(x.vendorId, '「' + c.title + '」が先送りされました（再開予定 ' + fmt(b.at) + '）'); });
  }
  if (b.op === 'resume') {
    update('cases', c, { status: 'open', postponeAt: '', updatedAt: now() });
    bs.forEach(function (x) { sysMsg(x, '依頼者が相談を再開しました'); notify(x.vendorId, '「' + c.title + '」の相談が再開されました'); });
  }
  return getCase(u, c.id);
}
function fmt(iso) { return Utilities.formatDate(new Date(iso), 'Asia/Tokyo', 'M/d H:mm'); }

/* ========== チャット ========== */
function myBid(u, bidId) { var x = find('bids', function (y) { return y.id === bidId; }); if (!x) return null; var c = myCase(u, x.caseId); return c ? { bid: x, c: c } : null; }
function sysMsg(bid, text) { insert('messages', { id: uid('m'), caseId: bid.caseId, bidId: bid.id, from: 'sys', type: 'sys', body: JSON.stringify({ text: text }), createdAt: now() }); }
function getChat(u, bidId, since) {
  var r = myBid(u, bidId); if (!r) return { ok: false, error: '見つかりません' };
  var ms = all('messages').filter(function (m) { return m.bidId === bidId && (!since || m.createdAt > since); }).map(function (m) { var o = pub(m); o.body = J(m.body, {}); return o; });
  update('bids', r.bid, { readAt: now() });
  var v = find('vendors', function (x) { return x.id === r.bid.vendorId; });
  return { ok: true, bid: pub(r.bid), vendor: v ? pub(v) : null, c: pub(r.c), msgs: ms };
}
function sendMsg(u, b) {
  var r = myBid(u, b.bidId); if (!r) return { ok: false, error: '見つかりません' };
  if (['lost', 'declined'].indexOf(r.bid.state) >= 0) return { ok: false, error: 'このやり取りは終了しています' };
  var text = String(b.text || '').slice(0, 2000);
  var body = b.img ? { img: b.img, text: text } : { text: text };
  insert('messages', { id: uid('m'), caseId: r.c.id, bidId: r.bid.id, from: u.id, type: b.img ? 'img' : 'text', body: JSON.stringify(body), createdAt: now() });
  if (!r.bid.firstMsgAt) { // 立候補料は依頼者の1通目で発生（見積もりだけの案件は無料）
    update('bids', r.bid, { firstMsgAt: now(), fee: r.c.estOnly ? 0 : 500 });
  }
  notify(r.bid.vendorId, '「' + r.c.title + '」に新しいメッセージがあります', 'bid:' + r.bid.id);
  return getChat(u, r.bid.id, b.since);
}
/* カード（提案・見積もり・金額変更・完了報告）への返事 */
function cardAction(u, b) {
  var r = myBid(u, b.bidId); if (!r) return { ok: false, error: '見つかりません' };
  var m = find('messages', function (x) { return x.id === b.msgId && x.bidId === r.bid.id; }); if (!m) return { ok: false, error: 'カードが見つかりません' };
  var body = J(m.body, {}); if (body.done) return { ok: false, error: 'すでに返事済みです' };
  var c = r.c, x = r.bid;
  if (m.type === 'prop') {
    if (b.op === 'pick') { body.done = '現地調査の日時が決まりました：' + b.slot; update('bids', x, { state: 'fixd', updatedAt: now() }); sysMsg(x, '現地調査の日時が決まりました：' + b.slot + '（住所が業者に伝わります）'); notify(x.vendorId, '「' + c.title + '」の現地調査が ' + b.slot + ' に決まりました', 'bid:' + x.id); }
    if (b.op === 'counter') { body.done = '候補を送りました。業者が24時間以内に時刻を決めます'; sysMsg(x, '依頼者から候補が届きました：' + b.slot); notify(x.vendorId, '「' + c.title + '」で依頼者から調査日の候補が届きました', 'bid:' + x.id); }
  } else if (m.type === 'est') {
    if (b.op === 'yes') {
      body.done = 'この内容で依頼しました（成約）'; update('bids', x, { state: 'won', amount: Number(body.amt), updatedAt: now() }); update('cases', c, { status: 'contracted', updatedAt: now() });
      sysMsg(x, '依頼が決まりました（成約）'); notify(x.vendorId, '「' + c.title + '」で依頼が決まりました（成約）', 'bid:' + x.id);
      all('bids').filter(function (o) { return o.caseId === c.id && o.id !== x.id && ['lost', 'declined'].indexOf(o.state) < 0; }).forEach(function (o) {
        update('bids', o, { state: 'other', updatedAt: now() }); sysMsg(o, '他社で決定済みになりました'); notify(o.vendorId, '「' + c.title + '」は他社で決定済みになりました', 'bid:' + o.id);
      });
    }
    if (b.op === 'no') { body.done = '見送りました'; update('bids', x, { state: 'lost', updatedAt: now() }); sysMsg(x, '依頼者が見送りました'); notify(x.vendorId, '「' + c.title + '」で見積もりが見送られました', 'bid:' + x.id); }
  } else if (m.type === 'chg') {
    if (b.op === 'ok') { body.done = '承諾しました。金額は ' + Number(body.to).toLocaleString() + '円 になります'; body.st = 'ok'; update('bids', x, { amount: Number(body.to), updatedAt: now() }); sysMsg(x, '追加の作業が承諾されました：合計 ' + Number(body.to).toLocaleString() + '円'); notify(x.vendorId, '「' + c.title + '」の追加の作業が承諾されました', 'bid:' + x.id); }
    if (b.op === 'ng') { body.done = '承諾しないと伝えました。追加の作業は行われません'; body.st = 'ng'; sysMsg(x, '追加の作業は承諾されませんでした。もとの見積もりの範囲で作業します'); notify(x.vendorId, '「' + c.title + '」の追加の作業は承諾されませんでした', 'bid:' + x.id); }
  } else if (m.type === 'rep') {
    if (b.op === 'fin') { body.done = '完了を確認しました'; completeCase(c, x, '依頼者が完了を確認しました'); }
    if (b.op === 'notyet') { body.done = '「まだ終わっていない」と伝えました'; body.notyet = (body.notyet || 0) + 1; sysMsg(x, '依頼者から「まだ終わっていない」と連絡がありました'); notify(x.vendorId, '「' + c.title + '」で依頼者から「まだ終わっていない」と連絡がありました', 'bid:' + x.id); }
  } else if (m.type === 'visit') {
    body.done = '回答ありがとうございました：' + b.op;
    if (b.op === '来なかった') { sysMsg(x, '依頼者から「現地調査に来なかった」と報告がありました'); notifyAdmin('現地調査に業者が来なかった', c, x); }
  }
  update('messages', m, { body: JSON.stringify(body) });
  return getChat(u, x.id);
}
function completeCase(c, x, why) {
  update('bids', x, { state: 'done', updatedAt: now() }); update('cases', c, { status: 'done', updatedAt: now() });
  sysMsg(x, '工事が完了しました（' + why + '）'); notify(x.vendorId, '「' + c.title + '」が完了しました', 'bid:' + x.id);
  notify(c.userId, '「' + c.title + '」が完了しました。よければ口コミを書いてください', 'case:' + c.id);
  all('bids').filter(function (o) { return o.caseId === c.id && o.state === 'other'; }).forEach(function (o) { update('bids', o, { state: 'lost', updatedAt: now() }); });
}
function postReview(u, b) {
  var r = myBid(u, b.bidId); if (!r || r.bid.state !== 'done') return { ok: false, error: '完了した依頼だけ口コミを書けます' };
  if (find('reviews', function (x) { return x.bidId === r.bid.id; })) return { ok: false, error: 'すでに口コミを書いています' };
  var st = Math.max(1, Math.min(5, Number(b.stars) || 0));
  insert('reviews', { id: uid('R'), caseId: r.c.id, bidId: r.bid.id, vendorId: r.bid.vendorId, userId: u.id, stars: st, text: String(b.text || '').slice(0, 1000), createdAt: now() });
  notify(r.bid.vendorId, '「' + r.c.title + '」に口コミが届きました');
  var v = find('vendors', function (x) { return x.id === r.bid.vendorId; });
  return { ok: true, googleUrl: v ? v.googleUrl : '' };
}

/* ========== 問い合わせ ========== */
function inquiry(u, b) {
  if (!b.type || !String(b.body || '').trim()) return { ok: false, error: '種類と内容を入れてください' };
  insert('inquiries', { id: uid('Q'), userId: u.id, type: b.type, body: String(b.body).slice(0, 3000), reply: '', status: 'open', createdAt: now(), repliedAt: '' });
  return { ok: true };
}
function notifyAdmin(what, c, x) { insert('inquiries', { id: uid('Q'), userId: 'system', type: '【要対応】' + what, body: '案件：' + (c ? c.title + '（' + c.id + '）' : '') + (x ? ' 業者：' + x.vendorId : ''), reply: '', status: 'open', createdAt: now(), repliedAt: '' }); }

/* ========== 毎時の自動処理 ========== */
function hourlyJobs() {
  var t = new Date();
  // 先送りの予告（1週間前）と当日
  all('cases').filter(function (c) { return c.status === 'postponed' && c.postponeAt; }).forEach(function (c) {
    var at = new Date(c.postponeAt), d = at - t;
    if (d <= 7 * 864e5 && d > 7 * 864e5 - 3600e3) notify(c.userId, '「' + c.title + '」の再開予定まで、あと1週間です', 'case:' + c.id);
    if (d <= 0 && d > -3600e3) notify(c.userId, '「' + c.title + '」の再開予定の日になりました。再開するか、もう少し先にするか、取りやめるかを選んでください', 'case:' + c.id);
  });
  // 完了報告から7日で自動完了
  all('messages').filter(function (m) { return m.type === 'rep'; }).forEach(function (m) {
    var body = J(m.body, {}); if (body.done || (t - new Date(m.createdAt)) < 7 * 864e5) return;
    var x = find('bids', function (b) { return b.id === m.bidId; }), c = x && find('cases', function (k) { return k.id === x.caseId; });
    if (!x || !c || c.status === 'done') return; body.done = '7日たったため、自動的に完了しました'; update('messages', m, { body: JSON.stringify(body) }); completeCase(c, x, '完了報告から7日経過');
  });
}

/* ========== テスト用：業者の動きを再現（TEST_MODE=on のときだけ） ========== */
function testOp(u, b) {
  if (prop('TEST_MODE', 'off') !== 'on') return { ok: false, error: 'テストモードではありません' };
  var c = myCase(u, b.caseId); if (!c) return { ok: false, error: '見つかりません' };
  var names = [['山田工務店（テスト）', 'リフォーム', 4.8, 23], ['松本水道設備（テスト）', '水のトラブル', 4.6, 11], ['中信設備（テスト）', '機器設置・修理', 4.2, 5]];
  if (b.op === 'bid') {
    var have = all('bids').filter(function (x) { return x.caseId === c.id; }).length; if (have >= 5) return { ok: false, error: '5社まで' };
    var n = names[have % 3], v = insert('vendors', { id: uid('V'), name: n[0], category: n[1], area: '松本市', star: n[2], reviews: n[3], googleUrl: 'https://www.google.com/maps', isTest: 'TRUE', createdAt: now() });
    var x = insert('bids', { id: uid('B'), caseId: c.id, vendorId: v.id, state: 'chat', firstMsgAt: '', fee: '', helpers: '', createdAt: now(), updatedAt: now() });
    insert('messages', { id: uid('m'), caseId: c.id, bidId: x.id, from: v.id, type: 'text', body: JSON.stringify({ text: 'はじめまして、' + n[0] + 'です。よろしければ、気になっていることを教えてください。' }), createdAt: now() });
    notify(u.id, '「' + c.title + '」に' + n[0] + 'が立候補しました', 'bid:' + x.id); return getCase(u, c.id);
  }
  var x2 = find('bids', function (y) { return y.id === b.bidId && y.caseId === c.id; }); if (!x2) return { ok: false, error: '業者を選んでください' };
  var card = null;
  if (b.op === 'msg') card = ['text', { text: b.text || 'ありがとうございます。一度見せていただければ正確な金額を出せます。' }];
  if (b.op === 'prop') { card = ['prop', { rows: [['部分補修の場合', '8,000円〜12,000円'], ['交換が必要な場合', '25,000円〜45,000円']], todo: '状態の確認', dur: 30, estDue: '調査後3日以内', slots: [[day(2), '10:00', '12:00'], [day(4), '13:00', '15:00']] }]; update('bids', x2, { state: 'prop' }); }
  if (b.op === 'est') { card = ['est', { amt: 38500, memo: '本体と取り付け工事一式。古いものの処分を含みます。', start: day(10), end: day(10), pay: '工事完了後に一括・銀行振込', due: day(3) }]; update('bids', x2, { state: 'est' }); }
  if (b.op === 'chg') { if (x2.state !== 'won') return { ok: false, error: '成約後に試せます' }; card = ['chg', { ad: 4400, from: Number(x2.amount), to: Number(x2.amount) + 4400, why: '追加で部品の交換が必要になったため' }]; }
  if (b.op === 'rep') { if (x2.state !== 'won') return { ok: false, error: '成約後に試せます' }; card = ['rep', { day: day(0), what: '交換工事一式', amt: Number(x2.amount) }]; }
  if (b.op === 'visit') card = ['visit', {}];
  if (!card) return { ok: false, error: '操作が違います' };
  insert('messages', { id: uid('m'), caseId: c.id, bidId: x2.id, from: x2.vendorId, type: card[0], body: JSON.stringify(card[1]), createdAt: now() });
  notify(u.id, '「' + c.title + '」に業者から新しい連絡があります', 'bid:' + x2.id);
  return getCase(u, c.id);
}
function day(n) { var d = new Date(Date.now() + n * 864e5); return Utilities.formatDate(d, 'Asia/Tokyo', 'M/d（E）').replace('Mon', '月').replace('Tue', '火').replace('Wed', '水').replace('Thu', '木').replace('Fri', '金').replace('Sat', '土').replace('Sun', '日'); }
