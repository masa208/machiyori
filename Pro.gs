/**
 * まち頼 Proモード（業者側）— テスト版
 * いまは「自分の依頼に、自分が業者として立候補して流れを試す」ための作り。
 * 本番の業者登録・カテゴリ別の新着配信は、業者登録フォームのあとで作る。
 */

/* 列が足りないシートに見出しを足す（1回だけ動く） */
function migrateV3() {
  if (prop('SCHEMA', '') === '3') return;
  ['vendors', 'bids'].forEach(function (n) {
    var sh = tbl(n), h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    TABLES[n].forEach(function (k) { if (h.indexOf(k) < 0) { sh.getRange(1, sh.getLastColumn() + 1).setValue(k); h.push(k); } });
  });
  PropertiesService.getScriptProperties().setProperty('SCHEMA', '3');
}

/* 業者あての通知は、その業者のアカウントに届ける */
function ownerOf(id) {
  if (String(id).charAt(0) !== 'V') return id;
  var v = find('vendors', function (x) { return x.id === id; });
  return v && v.userId ? v.userId : '';
}

function myVendor(u, create) {
  var v = find('vendors', function (x) { return x.userId === u.id; });
  if (!v && create) v = insert('vendors', { id: uid('V'), name: (u.callName || u.name || 'あなた') + '工務店（テスト）', category: 'テスト', area: u.area || '', star: 5, reviews: 0, googleUrl: 'https://www.google.com/maps', isTest: 'TRUE', createdAt: now(), userId: u.id });
  return v;
}

function proApi(u, b) {
  if (prop('TEST_MODE', 'off') !== 'on') return { ok: false, error: 'Proモードはテストモードのときだけ使えます' };
  migrateV3();
  switch (b.a) {
    case 'proHome': return proHome(u);
    case 'proBid': return proBid(u, b);
    case 'proChat': return proChat(u, b.bidId, b.since);
    case 'proSend': return proSend(u, b);
    case 'proCard': return proCard(u, b);
    case 'proUndo': return proUndo(u, b);
    case 'testReset': return testReset(u);
  }
  return null;
}

function proHome(u) {
  var v = myVendor(u, true), bids = all('bids'), cases = all('cases');
  var mine = bids.filter(function (x) { return x.vendorId === v.id; });
  var mineCase = {}; mine.forEach(function (x) { mineCase[x.caseId] = 1; });
  // テスト版：新着は「自分が出した依頼」のうち、まだ立候補していないもの
  var fresh = cases.filter(function (c) {
    if (c.userId !== u.id || c.status !== 'open' || c.stopBids || mineCase[c.id]) return false;
    return bids.filter(function (x) { return x.caseId === c.id; }).length < 5;
  }).map(pub).reverse();
  var run = mine.map(function (x) { var c = cases.filter(function (k) { return k.id === x.caseId; })[0]; var o = pub(x); o.c = c ? pub(c) : null; return o; })
    .filter(function (o) { return o.c && o.c.status !== 'withdrawn'; }).reverse();
  return { ok: true, vendor: pub(v), fresh: fresh, run: run };
}

function proBid(u, b) {
  var v = myVendor(u, true), c = find('cases', function (x) { return x.id === b.caseId; });
  if (!c || c.status !== 'open') return { ok: false, error: 'この依頼には立候補できません' };
  if (find('bids', function (x) { return x.caseId === c.id && x.vendorId === v.id; })) return { ok: false, error: 'すでに立候補しています' };
  if (all('bids').filter(function (x) { return x.caseId === c.id; }).length >= 5) return { ok: false, error: '立候補は5社までです' };
  var x = insert('bids', { id: uid('B'), caseId: c.id, vendorId: v.id, state: 'chat', firstMsgAt: '', fee: '', helpers: '', createdAt: now(), updatedAt: now(), readAt: '', amount: '' });
  insert('messages', { id: uid('m'), caseId: c.id, bidId: x.id, from: v.id, type: 'text', body: JSON.stringify({ text: 'はじめまして、' + v.name + 'です。くわしい状況を教えていただけますか？' }), createdAt: now() });
  notify(c.userId, '「' + c.title + '」に' + v.name + 'が立候補しました', 'bid:' + x.id);
  return { ok: true, bidId: x.id };
}

function proBidOf(u, bidId) {
  var v = myVendor(u, false); if (!v) return null;
  var x = find('bids', function (y) { return y.id === bidId && y.vendorId === v.id; }); if (!x) return null;
  return { v: v, bid: x, c: find('cases', function (k) { return k.id === x.caseId; }) };
}

function proChat(u, bidId, since) {
  var r = proBidOf(u, bidId); if (!r) return { ok: false, error: '見つかりません' };
  var ms = all('messages').filter(function (m) { return m.bidId === bidId && (!since || m.createdAt > since); }).map(function (m) { var o = pub(m); o.body = J(m.body, {}); return o; });
  var o = pub(r.c); o.photos = J(r.c.photos, []);
  return { ok: true, bid: pub(r.bid), vendor: pub(r.v), c: o, msgs: ms };
}

function proSend(u, b) {
  var r = proBidOf(u, b.bidId); if (!r) return { ok: false, error: '見つかりません' };
  if (['lost', 'declined', 'other'].indexOf(r.bid.state) >= 0) return { ok: false, error: 'このやり取りは終了しています' };
  var body = b.img ? { img: b.img, text: String(b.text || '') } : { text: String(b.text || '').slice(0, 2000) };
  insert('messages', { id: uid('m'), caseId: r.c.id, bidId: r.bid.id, from: r.v.id, type: b.img ? 'img' : 'text', body: JSON.stringify(body), createdAt: now() });
  notify(r.c.userId, '「' + r.c.title + '」に' + r.v.name + 'からメッセージがあります', 'bid:' + r.bid.id);
  return proChat(u, r.bid.id, b.since);
}

/* 業者が送るカード：prop（現地調査の提案）/ est（見積もり）/ chg（追加の作業の申請）/ rep（完了報告） */
function proCard(u, b) {
  var r = proBidOf(u, b.bidId); if (!r) return { ok: false, error: '見つかりません' };
  var x = r.bid, c = r.c, d = b.body || {}, label = '';
  if (b.type === 'prop') {
    if (!d.rows || !d.rows.length || !d.slots || !d.slots.length) return { ok: false, error: '工事費の目安と候補日時を入れてください' };
    update('bids', x, { state: 'prop', updatedAt: now() }); label = '現地調査の提案';
  } else if (b.type === 'est') {
    if (!(Number(d.amt) > 0) || !d.start || !d.end) return { ok: false, error: '金額と日付を入れてください' };
    if (d.end < d.start) return { ok: false, error: '完了予定日は着手予定日と同じか、それより後にしてください' };
    d.due = fmtD(new Date(Date.now() + 3 * 864e5));
    update('bids', x, { state: 'est', updatedAt: now() }); label = '見積もり';
  } else if (b.type === 'chg') {
    if (x.state !== 'won') return { ok: false, error: '成約後に申請できます' };
    if (!(Number(d.ad) > 0) || !d.why) return { ok: false, error: '金額と理由を入れてください' };
    d.from = Number(x.amount); d.to = d.from + Number(d.ad); label = '追加の作業の申請';
  } else if (b.type === 'rep') {
    if (x.state !== 'won') return { ok: false, error: '成約後に報告できます' };
    if (all('messages').some(function (m) { return m.bidId === x.id && m.type === 'rep' && !J(m.body, {}).done; })) return { ok: false, error: 'すでに完了報告を送っています' };
    var cap = Number(x.amount);
    d.amt = d.amt ? Math.min(Number(d.amt), cap) : cap; label = '完了報告';
  } else return { ok: false, error: '種類が違います' };
  insert('messages', { id: uid('m'), caseId: c.id, bidId: x.id, from: r.v.id, type: b.type, body: JSON.stringify(d), createdAt: now() });
  notify(c.userId, '「' + c.title + '」に' + r.v.name + 'から' + label + 'が届きました', 'bid:' + x.id);
  return proChat(u, x.id);
}
function fmtD(d) { return Utilities.formatDate(d, 'Asia/Tokyo', 'M/d'); }

/* 完了報告の取り消し（依頼者が確認する前だけ） */
function proUndo(u, b) {
  var r = proBidOf(u, b.bidId); if (!r) return { ok: false, error: '見つかりません' };
  var m = find('messages', function (y) { return y.id === b.msgId && y.bidId === r.bid.id && y.type === 'rep'; });
  if (!m || J(m.body, {}).done) return { ok: false, error: '取り消せません（依頼者が確認済みです）' };
  tbl('messages').deleteRow(m._row);
  sysMsg(r.bid, '業者が完了報告を取り消しました');
  return proChat(u, r.bid.id);
}

/* テストを最初からやり直す：自分の依頼・チャット・お知らせ・相談の履歴を消す */
function testReset(u) {
  var v = myVendor(u, false);
  var caseIds = {}; all('cases').forEach(function (c) { if (c.userId === u.id) caseIds[c.id] = 1; });
  var bidIds = {}; all('bids').forEach(function (x) { if (caseIds[x.caseId] || (v && x.vendorId === v.id)) bidIds[x.id] = 1; });
  var testVendors = {}; all('bids').forEach(function (x) { if (caseIds[x.caseId]) testVendors[x.vendorId] = 1; });
  delWhere('messages', function (m) { return caseIds[m.caseId] || bidIds[m.bidId]; });
  delWhere('reviews', function (q) { return caseIds[q.caseId]; });
  delWhere('bids', function (x) { return bidIds[x.id]; });
  delWhere('cases', function (c) { return caseIds[c.id]; });
  delWhere('notifications', function (n) { return n.userId === u.id; });
  delWhere('concierge', function (m) { return m.userId === u.id; });
  delWhere('vendors', function (x) { return testVendors[x.id] && x.isTest && !x.userId; });
  return { ok: true };
}
function delWhere(n, f) {
  var a = all(n).filter(f).map(function (o) { return o._row; }).sort(function (p, q) { return q - p; });
  var sh = tbl(n); a.forEach(function (row) { sh.deleteRow(row); });
}
