/**
 * 巡回チャット バックエンド
 * ・新規のスプレッドシートに紐づけた Apps Script として使う（現行の巡回管理システムとは完全に別）
 * ・個人のGoogleアカウントでデプロイすること（会社Workspaceだと外部から403になる）
 * ・スクリプトプロパティに次の2つを登録しておくこと
 *     SERVICE_ACCOUNT_JSON : Firebaseからダウンロードした秘密鍵(.json)の中身を丸ごと
 *     CHAT_ACCESS_CODE     : スタッフ全員に共有するPASS（自由に決める）
 *     CHAT_INTEGRATION_KEY : 巡回管理システムとの連携専用キー（PASSとは別の長い文字列。巡回管理側にも同じ値を登録）
 *
 * 機能：エリア別の部屋（東海／関東）、既読管理（既読人数・既読者一覧・ここから未読）、エリア別プッシュ通知
 */

var SHEET_MSG = 'メッセージ';
var SHEET_DEV = 'デバイス登録';
var SHEET_READ = '既読状態';
var AREAS = ['東海', '関東'];
var SYSTEM_SENDER = '巡回管理システム'; // 緊急指令カードの投稿者名
var PROPS = PropertiesService.getScriptProperties();

function doGet() {
  return ContentService.createTextOutput('junkai-chat api ok');
}

function doPost(e) {
  var res;
  try {
    var req = JSON.parse(e.postData.contents);
    if (req.action === 'command') {
      // 巡回管理システムからの投稿。スタッフ用のPASSとは別の連携キー（CHAT_INTEGRATION_KEY）で確認する
      return ContentService.createTextOutput(JSON.stringify(postCommand_(req))).setMimeType(ContentService.MimeType.JSON);
    }
    var expected = PROPS.getProperty('CHAT_ACCESS_CODE');
    if (!expected || String(req.code) !== expected) throw new Error('PASSが違います。');
    if (req.name !== undefined && req.name !== '' && !isValidName_(req.name)) {
      throw new Error('NAME_INVALID: お名前は、電話番号・メールアドレス以外（例：松井 武史）で入力してください。');
    }
    switch (req.action) {
      case 'register': res = registerDevice_(req); break;
      case 'send': res = sendMessage_(req); break;
      case 'list': res = listMessages_(req); break;
      case 'read': res = markRead_(req); break;
      case 'unsend': res = unsendMessage_(req); break;
      case 'search': res = searchMessages_(req); break;
      default: throw new Error('不明な操作です。');
    }
  } catch (err) {
    res = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- シート ---------- */

function getSheet_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
  }
  // 以前のバージョンで作った既存シートに、後から増えた列の見出しを足す
  var now = sh.getRange(1, 1, 1, headers.length).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if (!now[i]) sh.getRange(1, i + 1).setValue(headers[i]);
  }
  return sh;
}

// 返信先ID・メンション（名前をカンマ区切り）・取消日時（空でなければ取り消し済み。本文は管理用にシートへ残る）
function msgSheet_() { return getSheet_(SHEET_MSG, ['メッセージID', '投稿日時', '投稿者', '本文', 'エリア', '返信先ID', 'メンション', '取消日時']); }
function devSheet_() { return getSheet_(SHEET_DEV, ['ユーザー名', 'FCMトークン', '登録日時', 'エリア']); }
// 既読は「ユーザー×エリアごとに、どのメッセージまで読んだか」の1行だけを持つ（メッセージ数×人数で行が増えないようにするため）
function readSheet_() { return getSheet_(SHEET_READ, ['ユーザー名', '最終既読メッセージID', '更新日時', 'エリア']); }

/** 表示名として使えるか。電話番号・メールアドレスのような文字列は不可（同じ人が別名で何人にも見えるのを防ぐ） */
function isValidName_(n) {
  n = String(n || '').trim();
  if (!n || n.length > 30) return false;
  if (n.indexOf('@') !== -1) return false;
  if (/\d{5,}/.test(n)) return false;          // 数字が5桁以上続く（電話番号など）
  if (/^[\d\s\-+()]+$/.test(n)) return false;   // 数字と記号だけ
  return true;
}

function normArea_(a) {
  a = String(a || '').trim();
  if (AREAS.indexOf(a) === -1) throw new Error('エリアが不正です。');
  return a;
}

/** メッセージ全件を [ID, 日時, 投稿者, 本文, エリア, 返信先ID, メンション, 取消日時] の配列で返す（シート上の並び順のまま） */
function allMsgs_() {
  var sh = msgSheet_();
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 8).getValues();
}

/** クライアントから来たメンション一覧を安全な形に整える（名前の文字列だけ・最大20件） */
function cleanMentions_(v) {
  var out = [];
  (Array.isArray(v) ? v : []).forEach(function (m) {
    m = String(m || '').replace(/,/g, '').trim();
    if (m && m.length <= 30 && out.indexOf(m) === -1 && out.length < 20) out.push(m);
  });
  return out;
}

function roomOf_(all, area) {
  return all.filter(function (r) { return String(r[4] || '') === area; });
}

function indexMap_(room) {
  var map = {};
  for (var i = 0; i < room.length; i++) map[String(room[i][0])] = i;
  return map;
}

/** 既読位置の一覧 [{name, id, area}] */
function readPointers_() {
  var sh = readSheet_();
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 4).getValues().map(function (r) {
    return { name: String(r[0] || ''), id: String(r[1] || ''), area: String(r[3] || '') };
  });
}

/* ---------- API ---------- */

function registerDevice_(req) {
  var name = String(req.name || '').trim();
  var token = String(req.token || '').trim();
  var areas = (Array.isArray(req.areas) ? req.areas : []).filter(function (a) { return AREAS.indexOf(a) !== -1; });
  if (!name || !token) throw new Error('名前と通知トークンが必要です。');
  if (areas.length === 0) throw new Error('エリアを選んでください。');
  var areaText = areas.join(',');

  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sh = devSheet_();
    var last = sh.getLastRow();
    if (last >= 2) {
      var tokens = sh.getRange(2, 2, last - 1, 1).getValues();
      for (var i = 0; i < tokens.length; i++) {
        if (String(tokens[i][0]) === token) {
          sh.getRange(i + 2, 1).setValue(name);
          sh.getRange(i + 2, 3).setValue(new Date());
          sh.getRange(i + 2, 4).setValue(areaText);
          return { ok: true };
        }
      }
    }
    sh.appendRow([name, token, new Date(), areaText]);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 巡回管理システムからの「指令カード」投稿。専用の連携キー（スクリプトプロパティ CHAT_INTEGRATION_KEY）で確認する。
 * req: { key, area('東海'|'関東'), title, text }
 */
function postCommand_(req) {
  var expected = PROPS.getProperty('CHAT_INTEGRATION_KEY');
  if (!expected || String(req.key) !== expected) return { ok: false, error: '連携キーが違います。' };
  try {
    var area = normArea_(req.area);
    var title = String(req.title || 'お知らせ').substring(0, 60);
    var text = String(req.text || '').trim();
    if (!text) throw new Error('本文が空です。');
    if (text.length > 1800) text = text.substring(0, 1800) + '…';
    text = title + '\n' + text;

    var lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      msgSheet_().appendRow([Utilities.getUuid(), new Date(), SYSTEM_SENDER, text, area, '', '', '']);
    } finally {
      lock.releaseLock();
    }
    var pushInfo = '';
    try {
      pushInfo = pushToAll_('【' + area + '】' + title, text.length > 100 ? text.substring(0, 100) + '…' : text, '', area, [], '');
    } catch (err) {
      pushInfo = '通知エラー: ' + err;
      console.error(pushInfo);
    }
    return { ok: true, push: pushInfo };
  } catch (err) {
    return { ok: false, error: String(err && err.message ? err.message : err) };
  }
}

function sendMessage_(req) {
  var name = String(req.name || '').trim();
  var text = String(req.text || '').trim();
  var area = normArea_(req.area);
  if (!name) throw new Error('名前が必要です。');
  if (name === SYSTEM_SENDER) throw new Error('この名前は使えません。');
  if (!text) throw new Error('メッセージが空です。');
  if (text.length > 1000) throw new Error('メッセージは1000文字以内にしてください。');

  var mentions = cleanMentions_(req.mentions).filter(function (m) { return m !== name; });
  // 返信先は、同じ部屋に実在するメッセージIDだけ受け付ける
  var replyTo = String(req.replyTo || '').trim();
  if (replyTo && indexMap_(roomOf_(allMsgs_(), area))[replyTo] === undefined) replyTo = '';

  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    msgSheet_().appendRow([Utilities.getUuid(), new Date(), name, text, area, replyTo, mentions.join(','), '']);
  } finally {
    lock.releaseLock();
  }

  // 通知の失敗でメッセージ送信自体は失敗扱いにしない
  var pushInfo = '';
  try {
    pushInfo = pushToAll_('【' + area + '】' + name, text.length > 100 ? text.substring(0, 100) + '…' : text,
      String(req.token || ''), area, mentions, '【' + area + '】' + name + 'さんからあなたへ');
  } catch (err) {
    pushInfo = '通知エラー: ' + err;
    console.error(pushInfo);
  }
  return { ok: true, push: pushInfo };
}

/**
 * 指定エリアの部屋のメッセージ（新しい100件）を返す。
 *  readBy      : そのメッセージを読んだ人の名前（投稿者本人は除く）
 *  myLast      : 呼び出した本人が、この部屋で最後に読んだメッセージID（「ここから未読」用）
 *  unreadAreas : 呼び出した本人の担当エリアのうち、この部屋以外で未読があるエリア（タブの印用）
 */
function listMessages_(req) {
  var area = normArea_(req.area);
  var me = String(req.name || '').trim();
  if (me === SYSTEM_SENDER) throw new Error('この名前は使えません。');
  var myAreas =(Array.isArray(req.areas) ? req.areas : []).filter(function (a) { return AREAS.indexOf(a) !== -1; });

  var all = allMsgs_();
  var room = roomOf_(all, area);
  var idx = indexMap_(room);
  var pointers = readPointers_();

  var readers = [];
  var myLast = '';
  pointers.forEach(function (p) {
    if (p.area !== area || !p.name) return;
    var pos = idx[p.id];
    if (pos !== undefined) readers.push({ name: p.name, pos: pos });
    if (me && p.name === me) myLast = p.id;
  });

  var out = [];
  for (var i = Math.max(0, room.length - 100); i < room.length; i++) {
    var r = room[i];
    var readBy = [];
    readers.forEach(function (rd) {
      if (rd.pos >= i && rd.name !== String(r[2])) readBy.push(rd.name); // 投稿者本人は数えない
    });
    var deleted = !!r[7];
    var quote = null; // 返信先の引用（古くて100件より前のメッセージでも、同じ部屋にあれば引ける）
    if (!deleted && r[5]) {
      var qp = idx[String(r[5])];
      if (qp !== undefined) {
        var q = room[qp];
        quote = { id: q[0], name: q[2], deleted: !!q[7], text: q[7] ? '' : String(q[3]).substring(0, 60) };
      }
    }
    out.push({
      id: r[0],
      at: Utilities.formatDate(new Date(r[1]), 'Asia/Tokyo', 'M/d HH:mm'),
      name: r[2],
      text: deleted ? '' : r[3], // 取り消し済みは本文を返さない
      deleted: deleted,
      replyTo: quote,
      mentions: deleted ? [] : String(r[6] || '').split(',').filter(function (m) { return m; }),
      readBy: readBy
    });
  }

  // メンション候補：この部屋で発言した人・既読を付けた人・このエリアで通知登録している人
  var seen = {};
  room.forEach(function (r) { if (r[2]) seen[String(r[2])] = 1; });
  pointers.forEach(function (p) { if (p.area === area && p.name) seen[p.name] = 1; });
  var dsh = devSheet_();
  var dlast = dsh.getLastRow();
  if (dlast >= 2) {
    dsh.getRange(2, 1, dlast - 1, 4).getValues().forEach(function (r) {
      if (r[0] && (',' + String(r[3] || '') + ',').indexOf(',' + area + ',') !== -1) seen[String(r[0])] = 1;
    });
  }
  delete seen[SYSTEM_SENDER];
  var members = Object.keys(seen).sort();

  // 他のエリアの部屋に、自分がまだ読んでいない他人のメッセージがあるか
  var unreadAreas = [];
  if (me) {
    myAreas.forEach(function (a) {
      if (a === area) return;
      var other = roomOf_(all, a);
      var oidx = indexMap_(other);
      var pos = -1;
      pointers.forEach(function (p) {
        if (p.area === a && p.name === me && oidx[p.id] !== undefined) pos = oidx[p.id];
      });
      for (var j = pos + 1; j < other.length; j++) {
        if (String(other[j][2]) !== me) { unreadAreas.push(a); break; }
      }
    });
  }
  return { ok: true, messages: out, myLast: myLast, unreadAreas: unreadAreas, members: members };
}

/** 自分の投稿を取り消す（他人の投稿は不可）。行は残し、取消日時を入れる。既読位置はそのまま有効 */
function unsendMessage_(req) {
  var name = String(req.name || '').trim();
  var id = String(req.id || '').trim();
  var area = normArea_(req.area);
  if (!name || !id) throw new Error('名前とメッセージIDが必要です。');
  if (name === SYSTEM_SENDER) throw new Error('システムの投稿は取り消せません。');

  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sh = msgSheet_();
    var last = sh.getLastRow();
    if (last < 2) throw new Error('メッセージが見つかりません。');
    var rows = sh.getRange(2, 1, last - 1, 8).getValues();
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]) !== id) continue;
      if (String(rows[i][4] || '') !== area) throw new Error('メッセージが見つかりません。');
      if (String(rows[i][2]) !== name) throw new Error('取り消せるのは自分のメッセージだけです。');
      if (!rows[i][7]) sh.getRange(i + 2, 8).setValue(new Date());
      return { ok: true };
    }
    throw new Error('メッセージが見つかりません。');
  } finally {
    lock.releaseLock();
  }
}

/** この部屋の全履歴から、本文・投稿者名にキーワードを含むメッセージを新しい順に最大50件返す（取り消し済みは対象外） */
function searchMessages_(req) {
  var area = normArea_(req.area);
  var q = String(req.q || '').trim().substring(0, 50);
  if (!q) return { ok: true, results: [], total: 0 };
  var lower = q.toLowerCase();
  var room = roomOf_(allMsgs_(), area);
  var results = [];
  var total = 0;
  for (var i = room.length - 1; i >= 0; i--) {
    var r = room[i];
    if (r[7]) continue;
    var name = String(r[2]);
    var flat = String(r[3]).replace(/\s+/g, ' ');
    if ((name + '\n' + flat).toLowerCase().indexOf(lower) === -1) continue;
    total++;
    if (results.length >= 50) continue;
    var pos = flat.toLowerCase().indexOf(lower);
    var start = Math.max(0, pos - 30);
    results.push({
      id: r[0],
      at: Utilities.formatDate(new Date(r[1]), 'Asia/Tokyo', 'yyyy/M/d HH:mm'),
      name: name,
      snippet: (start > 0 ? '…' : '') + flat.substr(start, 120) + (start + 120 < flat.length ? '…' : '')
    });
  }
  return { ok: true, results: results, total: total };
}

/** 「このメッセージまで読んだ」を記録する。位置は進む方向にだけ更新（古い通信で巻き戻らない） */
function markRead_(req) {
  var name = String(req.name || '').trim();
  var lastId = String(req.lastId || '').trim();
  var area = normArea_(req.area);
  if (!name || !lastId) throw new Error('名前とメッセージIDが必要です。');
  var idx = indexMap_(roomOf_(allMsgs_(), area));
  var newPos = idx[lastId];
  if (newPos === undefined) throw new Error('メッセージが見つかりません。');

  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sh = readSheet_();
    var last = sh.getLastRow();
    if (last >= 2) {
      var rows = sh.getRange(2, 1, last - 1, 4).getValues();
      for (var i = 0; i < rows.length; i++) {
        if (String(rows[i][0]) === name && String(rows[i][3] || '') === area) {
          var oldPos = idx[String(rows[i][1])];
          if (oldPos === undefined || newPos > oldPos) {
            sh.getRange(i + 2, 2, 1, 2).setValues([[lastId, new Date()]]);
          }
          return { ok: true };
        }
      }
    }
    sh.appendRow([name, lastId, new Date(), area]);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- Push通知（FCM HTTP v1） ---------- */

/**
 * area を指定すると、そのエリアを担当にしている端末だけに送る（空なら全端末）。
 * mentions に名前（または「全員」）があれば、その人の端末には mentionTitle（あなた宛）の通知にする。
 */
function pushToAll_(title, body, excludeToken, area, mentions, mentionTitle) {
  mentions = mentions || [];
  var sa = JSON.parse(PROPS.getProperty('SERVICE_ACCOUNT_JSON'));
  var accessToken = getAccessToken_(sa);
  var sh = devSheet_();
  var last = sh.getLastRow();
  if (last < 2) return '送信先なし';
  var rows = sh.getRange(2, 1, last - 1, 4).getValues();

  var targets = [];
  var requests = [];
  rows.forEach(function (r, idx) {
    var token = String(r[1]);
    if (!token || token === excludeToken) return;
    if (area && (',' + String(r[3] || '') + ',').indexOf(',' + area + ',') === -1) return; // 別エリアの端末には送らない
    targets.push(idx + 2); // シート上の行番号
    var forMe = mentions.length > 0 && mentionTitle && (mentions.indexOf('全員') !== -1 || mentions.indexOf(String(r[0])) !== -1);
    requests.push({
      url: 'https://fcm.googleapis.com/v1/projects/' + sa.project_id + '/messages:send',
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + accessToken },
      muteHttpExceptions: true,
      payload: JSON.stringify({
        message: {
          token: token,
          data: { title: String(forMe ? mentionTitle : title), body: String(body) },
          webpush: { headers: { Urgency: 'high', TTL: '3600' } }
        }
      })
    });
  });
  if (requests.length === 0) return '送信先なし';

  var responses = UrlFetchApp.fetchAll(requests);
  var ok = 0;
  var deadRows = [];
  responses.forEach(function (res, i) {
    var code = res.getResponseCode();
    if (code === 200) { ok++; return; }
    var text = res.getContentText();
    if (code === 404 || text.indexOf('UNREGISTERED') !== -1 || text.indexOf('INVALID_ARGUMENT') !== -1) {
      deadRows.push(targets[i]); // 無効になったトークンは後で削除
    } else {
      console.error('FCM送信エラー ' + code + ': ' + text);
    }
  });

  if (deadRows.length > 0) {
    var lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      deadRows.sort(function (a, b) { return b - a; }).forEach(function (r) { sh.deleteRow(r); });
    } finally {
      lock.releaseLock();
    }
  }
  return ok + '件送信成功 / 無効トークン削除' + deadRows.length + '件';
}

function b64url_(str) {
  return Utilities.base64EncodeWebSafe(str, Utilities.Charset.UTF_8).replace(/=+$/, '');
}

/** サービスアカウントでGoogleのアクセストークンを取得する（50分キャッシュ） */
function getAccessToken_(sa) {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('fcm_access_token');
  if (cached) return cached;

  var now = Math.floor(Date.now() / 1000);
  var header = { alg: 'RS256', typ: 'JWT' };
  var claim = {
    iss: sa.client_email,
    scope: 'https://www.googleapis.com/auth/firebase.messaging',
    aud: 'https://oauth2.googleapis.com/token',
    iat: now,
    exp: now + 3600
  };
  var unsigned = b64url_(JSON.stringify(header)) + '.' + b64url_(JSON.stringify(claim));
  var signature = Utilities.computeRsaSha256Signature(unsigned, sa.private_key);
  var jwt = unsigned + '.' + Utilities.base64EncodeWebSafe(signature).replace(/=+$/, '');

  var res = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method: 'post',
    payload: { grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer', assertion: jwt },
    muteHttpExceptions: true
  });
  var data = JSON.parse(res.getContentText());
  if (!data.access_token) throw new Error('アクセストークン取得に失敗: ' + res.getContentText());
  cache.put('fcm_access_token', data.access_token, 3000);
  return data.access_token;
}

/* ---------- 動作確認用（エディタから実行） ---------- */

/** 初回の権限承認と、登録済み端末へのテスト通知。結果は「実行ログ」に出る。 */
function testPush() {
  var result = pushToAll_('テスト通知', 'Apps Scriptからのテストです', '', '');
  console.log(result);
}

/**
 * 【1回だけ実行】エリア機能を入れる前に投稿された古いメッセージ（エリア欄が空）を、指定エリアの部屋に移す。
 * 実行しなくても動作に問題はないが、その場合は古いメッセージはどの部屋にも表示されない。
 * 使い方：下の '東海' を移したいエリアに変えて、エディタでこの関数を選んで実行する。
 */
function migrateOldMessages() {
  var targetArea = '東海';
  var sh = msgSheet_();
  var last = sh.getLastRow();
  if (last < 2) return;
  var rng = sh.getRange(2, 5, last - 1, 1);
  var vals = rng.getValues().map(function (r) { return [r[0] ? r[0] : targetArea]; });
  rng.setValues(vals);
  console.log('エリア欄が空のメッセージを「' + targetArea + '」にしました。');
}


/* ============================================================================
 * 名前の整理（Apps Scriptのエディタから手で実行する）
 *  1) showChatUsers を実行 → 実行ログに、使われている名前の一覧（メッセージ数・端末数）が出る
 *  2) 下の RENAMES に「変更前の名前: 変更後の名前」を書く（例：電話番号やメールアドレスで入っていた人）
 *  3) applyRenames を実行 → メッセージ・端末登録・既読状態の名前が一括で置き換わる
 *  ※ 同じ人が複数の名前で入っている場合は、全部を同じ変更後の名前に向ければ1人にまとまる
 * ========================================================================== */
var RENAMES = {
  // '080-0000-0000': '山田 太郎',
  // 'taro_yamada@example.co.jp': '山田 太郎'
};

function showChatUsers() {
  var count = {};
  function add(n, k) { n = String(n || ''); if (!n) return; if (!count[n]) count[n] = { msg: 0, dev: 0, read: 0 }; count[n][k]++; }
  allMsgs_().forEach(function (r) { add(r[2], 'msg'); });
  var d = devSheet_(); if (d.getLastRow() >= 2) d.getRange(2, 1, d.getLastRow() - 1, 1).getValues().forEach(function (r) { add(r[0], 'dev'); });
  var rd = readSheet_(); if (rd.getLastRow() >= 2) rd.getRange(2, 1, rd.getLastRow() - 1, 1).getValues().forEach(function (r) { add(r[0], 'read'); });
  Object.keys(count).sort().forEach(function (n) {
    var c = count[n];
    Logger.log((isValidName_(n) ? '　' : '要確認 ') + n + '　メッセージ' + c.msg + '件／端末' + c.dev + '台／既読' + c.read + '行');
  });
}

function applyRenames() {
  var keys = Object.keys(RENAMES);
  if (keys.length === 0) { Logger.log('RENAMES が空です。変更前→変更後を書いてから実行してください。'); return; }
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    keys.forEach(function (oldName) { renameChatUser_(oldName, RENAMES[oldName]); });
  } finally {
    lock.releaseLock();
  }
}

function renameChatUser_(oldName, newName) {
  oldName = String(oldName).trim(); newName = String(newName).trim();
  if (!oldName || !newName || oldName === newName) return;
  if (!isValidName_(newName)) throw new Error('変更後の名前が使えません: ' + newName);
  if (newName === SYSTEM_SENDER) throw new Error('この名前は使えません。');
  var n = { msg: 0, mention: 0, dev: 0, read: 0, readMerged: 0 };

  // メッセージ：投稿者(C列)とメンション(G列)
  var ms = msgSheet_(), ml = ms.getLastRow();
  if (ml >= 2) {
    var rows = ms.getRange(2, 1, ml - 1, 8).getValues(), cc = [], gg = [];
    rows.forEach(function (r) {
      var who = String(r[2]);
      if (who === oldName) { who = newName; n.msg++; }
      var men = String(r[6] || '').split(',').filter(function (m) { return m; });
      var changed = false;
      men = men.map(function (m) { if (m === oldName) { changed = true; return newName; } return m; });
      if (changed) n.mention++;
      cc.push([who]); gg.push([men.join(',')]);
    });
    ms.getRange(2, 3, cc.length, 1).setValues(cc);
    ms.getRange(2, 7, gg.length, 1).setValues(gg);
  }
  // 端末登録（A列）
  var ds = devSheet_(), dl = ds.getLastRow();
  if (dl >= 2) {
    var dv = ds.getRange(2, 1, dl - 1, 1).getValues().map(function (r) { if (String(r[0]) === oldName) { n.dev++; return [newName]; } return [r[0]]; });
    ds.getRange(2, 1, dv.length, 1).setValues(dv);
  }
  // 既読状態：変更後の名前の行が同じエリアにすでにあれば、より新しい方（読み進んだ位置）を残して1行にまとめる
  var rs = readSheet_(), rl = rs.getLastRow();
  if (rl >= 2) {
    var all = allMsgs_();
    var rr = rs.getRange(2, 1, rl - 1, 4).getValues();
    var del = [];
    for (var i = 0; i < rr.length; i++) {
      if (String(rr[i][0]) !== oldName) continue;
      var area = String(rr[i][3] || '');
      var idx = indexMap_(roomOf_(all, area));
      var twin = -1;
      for (var j = 0; j < rr.length; j++) { if (j !== i && String(rr[j][0]) === newName && String(rr[j][3] || '') === area) { twin = j; break; } }
      if (twin === -1) {
        rs.getRange(i + 2, 1).setValue(newName); rr[i][0] = newName; n.read++;
      } else {
        var po = idx[String(rr[i][1])], pt = idx[String(rr[twin][1])];
        if (po !== undefined && (pt === undefined || po > pt)) rs.getRange(twin + 2, 2, 1, 2).setValues([[rr[i][1], rr[i][2]]]);
        del.push(i + 2); n.readMerged++;
      }
    }
    del.sort(function (a, b) { return b - a; }).forEach(function (row) { rs.deleteRow(row); });
  }
  Logger.log('「' + oldName + '」→「' + newName + '」 メッセージ' + n.msg + '件／メンション' + n.mention + '件／端末' + n.dev + '台／既読' + n.read + '行（統合' + n.readMerged + '行）');
}
