// ============================================================
// 穴センサー GAS スクリプト
// デプロイ方法: Apps Script エディタに貼り付け → ウェブアプリとして公開
//   実行: 自分、アクセス: 全員（匿名を含む）
// トリガー: checkNewRows → スプレッドシートの onChange (変更時)
// ============================================================

// ─── 設定 ───────────────────────────────────────────
const CONFIG = {
  SHEET_ID:       '1S9U_AR4dM8tKTUTKx3_wAJBLW5x4zB3h7annjv7z8iw',
  LOG_SHEET:      '係数サマリー',
  SUB_SHEET:      'subscriptions',             // 購読シート名
  LAST_ROW_KEY:   'lastProcessedRow',          // Script Properties キー
  LATEST_KEY:     'latestSensorData',          // 最新センサーデータ

  VAPID_PRIVATE:  'lOxrs0RwvXg4pbwmchDSlX_HJvNj8yKtUqUdLwg7Ves',
  VAPID_PUBLIC:   'BO13tsTjl2y_vuX84DIzUbbWUgndqDKnvi7CF-9kkeK5ZBjeTRck4m5X8zKFLgN_-8erCil_UC4Ei1tE5fgmM-M',
  VAPID_MAILTO:   'mailto:noreply@example.com',
};

// ─── doPost: 購読エンドポイントを保存 ───────────────
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.action !== 'subscribe') return _ok({ status: 'ignored' });

    const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    let sheet   = ss.getSheetByName(CONFIG.SUB_SHEET);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SUB_SHEET);
      sheet.appendRow(['endpoint', 'p256dh', 'auth', 'registered_at']);
    }

    // 重複チェック
    const data     = sheet.getDataRange().getValues();
    const existing = data.slice(1).some(r => r[0] === body.endpoint);
    if (!existing) {
      sheet.appendRow([body.endpoint, body.p256dh, body.auth, new Date().toISOString()]);
    }
    return _ok({ status: 'ok' });
  } catch (err) {
    return _ok({ status: 'error', message: String(err) });
  }
}

// ─── doGet: 最新センサーデータ or healthcheck ────────
function doGet(e) {
  const action = e?.parameter?.action || '';
  if (action === 'latest') {
    const raw  = PropertiesService.getScriptProperties().getProperty(CONFIG.LATEST_KEY);
    const data = raw ? JSON.parse(raw) : { active: false };
    return _json(data);
  }
  return _json({ status: 'ok' });
}

// ─── checkNewRows: トリガーから呼び出し ──────────────
function checkNewRows() {
  const props = PropertiesService.getScriptProperties();
  const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.LOG_SHEET);
  if (!sheet) return;

  const currentLastRow = sheet.getLastRow();

  // 初回実行時: 既存行をスキップして現在の最終行を記録して終了
  if (!props.getProperty(CONFIG.LAST_ROW_KEY)) {
    props.setProperty(CONFIG.LAST_ROW_KEY, String(currentLastRow));
    return;
  }

  const lastDone = parseInt(props.getProperty(CONFIG.LAST_ROW_KEY), 10);
  if (currentLastRow <= lastDone) return;

  for (let row = lastDone + 1; row <= currentLastRow; row++) {
    try {
      _processRow(sheet, row);
    } catch (err) {
      console.error('Row ' + row + ' error: ' + err);
    }
  }
  props.setProperty(CONFIG.LAST_ROW_KEY, String(currentLastRow));
}

function _processRow(sheet, row) {
  // F列（index 5）の完全ログJSONを取得
  const raw = sheet.getRange(row, 6).getValue();
  const log = _tryParse(raw);
  if (!log) return;

  const raceInfo = log.race_info;
  const snapshot = log.snapshot;
  const prediction = log.prediction;
  if (!raceInfo || !snapshot) return;

  // 条件判定
  const grade     = raceInfo.grade;
  const windSpeed = raceInfo.wind?.speed;
  const tenun     = raceInfo.tenun;

  const isHit =
    grade     === 'a-kyu' &&
    windSpeed >= 1.5      &&
    windSpeed <  3.1      &&
    tenun     === 33;

  if (!isHit) return;

  // R1/R2/R3: snapshot.scores.final.seiten を値の降順ソート
  const seiten  = snapshot?.scores?.final?.seiten || {};
  const ranking = Object.entries(seiten)
    .sort((a, b) => b[1] - a[1])
    .map(e => e[0]);
  const R1 = ranking[0], R2 = ranking[1], R3 = ranking[2];

  // L: prediction.kouten から正規表現抽出
  const koutenHtml = prediction?.kouten || '';
  const lMatch     = koutenHtml.match(/特異点：(\d+)/);
  const L          = lMatch ? lMatch[1] : ranking[3] || '?';

  // 買い目5点
  const betsResult = [
    `${R1}-${R2}-${L}`,
    `${R1}-${L}-${R2}`,
    `${R1}-${L}-${R3}`,
    `${R3}-${R2}-${R1}`,
    `${R2}-${R1}-${L}`,
  ];

  const payload = {
    active: true,
    bank:   raceInfo.bank || '',
    wind:   windSpeed,
    tenun,
    R1, R2, R3, L,
    bets:   betsResult,
    ts:     new Date().toISOString(),
  };

  // Script Properties に保存（doGet で返す用）
  PropertiesService.getScriptProperties().setProperty(
    CONFIG.LATEST_KEY, JSON.stringify(payload)
  );

  // 全購読者に Push 送信
  _sendPushToAll(payload);
}

// ─── Push送信 ────────────────────────────────────────
function _sendPushToAll() {
  const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SUB_SHEET);
  if (!sheet) return;

  const rows   = sheet.getDataRange().getValues().slice(1); // ヘッダー除く
  const failed = [];

  rows.forEach((row, i) => {
    const endpoint = row[0];
    if (!endpoint) return;
    try {
      const code = _sendVapidPush(endpoint);
      if (code === 410 || code === 404) failed.push(i + 2); // 1-indexed + header
    } catch (err) {
      console.error(`Push failed for row ${i + 2}: ${err}`);
    }
  });

  // 無効な購読を削除（後ろから削除してインデックスズレを防ぐ）
  failed.reverse().forEach(r => sheet.deleteRow(r));
}

function _sendVapidPush(endpoint) {
  const { jwt, pubKey } = _createVapidJwt(endpoint);
  const resp = UrlFetchApp.fetch(endpoint, {
    method: 'post',
    headers: {
      'Authorization': `vapid t=${jwt},k=${pubKey}`,
      'TTL': '86400',
    },
    muteHttpExceptions: true,
  });
  return resp.getResponseCode();
}

// ─── VAPID JWT 生成（P-256 ECDSA、GAS V8 BigInt使用）───
function _createVapidJwt(endpoint) {
  const origin = endpoint.match(/^https?:\/\/[^\/]+/)[0];
  const now    = Math.floor(Date.now() / 1000);
  const _b64u  = s => Utilities.base64EncodeWebSafe(s).replace(/=+$/, '');

  const header  = _b64u(JSON.stringify({ typ: 'JWT', alg: 'ES256' }));
  const payload = _b64u(JSON.stringify({
    aud: origin,
    exp: now + 43200,
    sub: CONFIG.VAPID_MAILTO,
  }));
  const signingInput = `${header}.${payload}`;
  const signature    = _p256Sign(CONFIG.VAPID_PRIVATE, signingInput);

  return { jwt: `${signingInput}.${signature}`, pubKey: CONFIG.VAPID_PUBLIC };
}

// ─── P-256 ECDSA（RFC 6979、GAS V8 BigInt）───────────
function _p256Sign(privateKeyB64u, message) {
  // P-256 パラメータ
  const p  = BigInt('0xFFFFFFFF00000001000000000000000000000000FFFFFFFFFFFFFFFFFFFFFFFF');
  const a  = p - BigInt(3);
  const n  = BigInt('0xFFFFFFFF00000000FFFFFFFFFFFFFFFFBCE6FAADA7179E84F3B9CAC2FC632551');
  const Gx = BigInt('0x6B17D1F2E12C4247F8BCE6E563A440F277037D812DEB33A0F4A13945D898C296');
  const Gy = BigInt('0x4FE342E2FE1A7F9B8EE7EB4A7C0F9E162BCE33576B315ECECBB6406837BF51F5');
  const G  = [Gx, Gy];

  // フィールド演算
  const _fm = (x, m) => { let r = x % m; return r < BigInt(0) ? r + m : r; };
  const _fi = (x, m) => {
    let [r0, r1, s0, s1] = [_fm(x, m), m, BigInt(1), BigInt(0)];
    while (r1 !== BigInt(0)) {
      const q = r0 / r1;
      [r0, r1] = [r1, r0 - q * r1];
      [s0, s1] = [s1, s0 - q * s1];
    }
    return _fm(s0, m);
  };

  // 楕円曲線演算
  const _add = (P, Q) => {
    if (!P) return Q; if (!Q) return P;
    const [x1, y1] = P, [x2, y2] = Q;
    if (x1 === x2) {
      if (y1 !== y2) return null;
      const m = _fm(BigInt(3) * x1 * x1 + a, p) * _fi(BigInt(2) * y1, p) % p;
      const x3 = _fm(m * m - BigInt(2) * x1, p);
      return [x3, _fm(m * (x1 - x3) - y1, p)];
    }
    const m  = _fm(y2 - y1, p) * _fi(_fm(x2 - x1, p), p) % p;
    const x3 = _fm(m * m - x1 - x2, p);
    return [x3, _fm(m * (x1 - x3) - y1, p)];
  };
  const _mul = (k, P) => {
    let R = null, Q = P;
    while (k > BigInt(0)) { if (k & BigInt(1)) R = _add(R, Q); Q = _add(Q, Q); k >>= BigInt(1); }
    return R;
  };

  // バイト変換ユーティリティ
  const _toHex = b => (b & 0xFF).toString(16).padStart(2, '0');
  const _bytes32 = v => {
    const h = v.toString(16).padStart(64, '0');
    return new Uint8Array(Array.from({length: 32}, (_, i) => parseInt(h.slice(i*2, i*2+2), 16)));
  };
  const _fromBytes = arr => BigInt('0x' + Array.from(arr).map(_toHex).join(''));
  const _cat = (...arrs) => {
    const r = new Uint8Array(arrs.reduce((s, a) => s + a.length, 0));
    let o = 0; arrs.forEach(a => { r.set(a, o); o += a.length; });
    return r;
  };

  // GAS HMAC-SHA256（signed byte → unsigned byte 変換必須）
  const _hmac = (key, data) => {
    const k = Array.from(key).map(b => b > 127 ? b - 256 : b);
    const d = Array.from(data).map(b => b > 127 ? b - 256 : b);
    return Uint8Array.from(Utilities.computeHmacSha256Signature(d, k).map(b => b & 0xFF));
  };

  // GAS SHA-256
  const _sha256 = msg => {
    const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, msg);
    return Uint8Array.from(raw.map(b => b & 0xFF));
  };

  // base64url デコード
  const _b64uDec = s => {
    const pad = s.replace(/-/g, '+').replace(/_/g, '/') + '==='.slice((s.length + 3) % 4);
    return Uint8Array.from(Utilities.base64Decode(pad));
  };

  // RFC 6979 決定論的 k 生成
  const _rfcK = (privB, hashB) => {
    let V = new Uint8Array(32).fill(1);
    let K = new Uint8Array(32).fill(0);
    K = _hmac(K, _cat(V, new Uint8Array([0]), privB, hashB));
    V = _hmac(K, V);
    K = _hmac(K, _cat(V, new Uint8Array([1]), privB, hashB));
    V = _hmac(K, V);
    for (let i = 0; i < 100; i++) {
      V = _hmac(K, V);
      const k = _fromBytes(V);
      if (k >= BigInt(1) && k < n) return k;
      K = _hmac(K, _cat(V, new Uint8Array([0])));
      V = _hmac(K, V);
    }
    throw new Error('RFC 6979: k generation failed');
  };

  // 署名
  const privKey = _fromBytes(_b64uDec(privateKeyB64u));
  const hashB   = _sha256(message);
  const z       = _fromBytes(hashB);
  const k       = _rfcK(_bytes32(privKey), hashB);
  const [rx]    = _mul(k, G);
  const r       = _fm(rx, n);
  const s       = _fm(_fi(k, n) * _fm(z + r * privKey, n), n);

  // 署名バイト (r || s) 64bytes → base64url
  const sig = _cat(_bytes32(r), _bytes32(s));
  return Utilities.base64EncodeWebSafe(sig).replace(/=+$/, '');
}

// ─── ユーティリティ ───────────────────────────────────
function _tryParse(str) {
  try { return typeof str === 'string' ? JSON.parse(str) : str; } catch (_) { return null; }
}
function _ok(obj)   { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
function _json(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
