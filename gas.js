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
  PUSH_STATUS_KEY: 'lastPushStatus',
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

// ─── onChange トリガー自動設定 ───────────────────────
function _ensureTrigger() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('triggerConfigured') === 'true') return;
  try {
    const existing = ScriptApp.getProjectTriggers()
      .filter(t => t.getHandlerFunction() === 'checkNewRows');
    if (existing.length === 0) {
      ScriptApp.newTrigger('checkNewRows')
        .forSpreadsheet(SpreadsheetApp.openById(CONFIG.SHEET_ID))
        .onChange()
        .create();
    }
    props.setProperty('triggerConfigured', 'true');
  } catch (_) {}
}

// ─── doGet: 最新センサーデータ / 診断 / healthcheck ──
function doGet(e) {
  _ensureTrigger();
  const action = e?.parameter?.action || '';

  if (action === 'latest') {
    const raw  = PropertiesService.getScriptProperties().getProperty(CONFIG.LATEST_KEY);
    const data = raw ? JSON.parse(raw) : { active: false };
    return _json(data);
  }

  if (action === 'diag') {
    const props    = PropertiesService.getScriptProperties();
    const lastRow  = props.getProperty(CONFIG.LAST_ROW_KEY);
    const ss       = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const logSheet = ss.getSheetByName(CONFIG.LOG_SHEET);
    const subSheet = ss.getSheetByName(CONFIG.SUB_SHEET);
    const subCount = subSheet ? Math.max(0, subSheet.getLastRow() - 1) : 0;
    const raw      = props.getProperty(CONFIG.LATEST_KEY);
    const triggerCount = ScriptApp.getProjectTriggers()
      .filter(t => t.getHandlerFunction() === 'checkNewRows').length;
    const pushStatusRaw = props.getProperty(CONFIG.PUSH_STATUS_KEY);
    return _json({
      lastProcessedRow:  lastRow ? parseInt(lastRow) : null,
      currentLastRow:    logSheet ? logSheet.getLastRow() : null,
      subscriptionCount: subCount,
      triggerCount,
      latestSensorData:  raw ? JSON.parse(raw) : null,
      lastPushStatus:    pushStatusRaw ? JSON.parse(pushStatusRaw) : null,
    });
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

  _resetLastRow(props, currentLastRow);
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
    grade     === 'a-kyu'  &&
    windSpeed >= 1.5       &&
    windSpeed <  3.1       &&
    Number(tenun) === 33;  // 文字列 "33" も許容

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
function _sendPushToAll(payload) {
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
      PropertiesService.getScriptProperties().setProperty(
        CONFIG.PUSH_STATUS_KEY,
        JSON.stringify({ code: 0, body: String(err), endpoint: endpoint.slice(-30), ts: new Date().toISOString() })
      );
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
      'Authorization':  `vapid t=${jwt},k=${pubKey}`,
      'TTL':            '86400',
      'Content-Type':   'application/octet-stream',
      'Content-Length': '0',
    },
    muteHttpExceptions: true,
  });
  const code = resp.getResponseCode();
  const body = resp.getContentText().slice(0, 200);
  PropertiesService.getScriptProperties().setProperty(
    CONFIG.PUSH_STATUS_KEY,
    JSON.stringify({ code, body, endpoint: endpoint.slice(-30), ts: new Date().toISOString() })
  );
  return code;
}

// ─── VAPID JWT 生成 ──────────────────────────────────
function _createVapidJwt(endpoint) {
  const origin = endpoint.match(/^https?:\/\/[^\/]+/)[0];
  const now    = Math.floor(Date.now() / 1000);
  const b64u   = s => Utilities.base64EncodeWebSafe(s).replace(/=+$/, '');

  const header  = b64u(JSON.stringify({ typ: 'JWT', alg: 'ES256' }));
  const payload = b64u(JSON.stringify({
    aud: origin,
    exp: now + 43200,
    sub: CONFIG.VAPID_MAILTO,
  }));
  const sigInput = `${header}.${payload}`;
  return { jwt: `${sigInput}.${_p256Sign(CONFIG.VAPID_PRIVATE, sigInput)}`, pubKey: CONFIG.VAPID_PUBLIC };
}

// ─── P-256 ECDSA（RFC 6979、GAS V8 BigInt）───────────
// 修正: Uint8Array → 通常 Array 統一 / BigInt() コンストラクタ / padding修正 / sha256明示バイト変換
function _p256Sign(privKeyB64u, message) {
  // P-256 曲線パラメータ
  const p  = BigInt('0xFFFFFFFF00000001000000000000000000000000FFFFFFFFFFFFFFFFFFFFFFFF');
  const a  = BigInt('0xFFFFFFFF00000001000000000000000000000000FFFFFFFFFFFFFFFFFFFFFFFC');
  const n  = BigInt('0xFFFFFFFF00000000FFFFFFFFFFFFFFFFBCE6FAADA7179E84F3B9CAC2FC632551');
  const Gx = BigInt('0x6B17D1F2E12C4247F8BCE6E563A440F277037D812DEB33A0F4A13945D898C296');
  const Gy = BigInt('0x4FE342E2FE1A7F9B8EE7EB4A7C0F9E162BCE33576B315ECECBB6406837BF51F5');

  // フィールド演算
  const mod = (x, m) => { const r = x % m; return r < BigInt(0) ? r + m : r; };
  const inv = (x, m) => {
    let [a, b, u, v] = [mod(x, m), m, BigInt(1), BigInt(0)];
    while (a !== BigInt(0)) {
      const q = b / a;
      [a, b] = [b - q * a, a];
      [u, v] = [v - q * u, u];
    }
    return mod(v, m);
  };

  // 楕円曲線演算
  const add = (P, Q) => {
    if (!P) return Q;
    if (!Q) return P;
    const [x1, y1] = P, [x2, y2] = Q;
    let lam;
    if (x1 === x2) {
      if (y1 !== y2) return null;
      lam = mod(BigInt(3) * x1 * x1 + a, p) * inv(mod(BigInt(2) * y1, p), p) % p;
    } else {
      lam = mod(y2 - y1, p) * inv(mod(x2 - x1, p), p) % p;
    }
    const x3 = mod(lam * lam - x1 - x2, p);
    return [x3, mod(lam * (x1 - x3) - y1, p)];
  };
  const mul = (k, P) => {
    let R = null, Q = [...P];
    for (; k > BigInt(0); k >>= BigInt(1)) { if (k & BigInt(1)) R = add(R, Q); Q = add(Q, Q); }
    return R;
  };

  // バイト変換（通常 Array のみ使用）
  const b64uDec = s => {
    const b = s.replace(/-/g, '+').replace(/_/g, '/');
    const r = b.length % 4;
    return Array.from(Utilities.base64Decode(r ? b + '='.repeat(4 - r) : b)).map(x => x & 0xFF);
  };
  const toBI   = arr => BigInt('0x' + arr.map(b => b.toString(16).padStart(2, '0')).join(''));
  const to32B  = v   => {
    const h = v.toString(16).padStart(64, '0');
    return Array.from({length: 32}, (_, i) => parseInt(h.slice(i * 2, i * 2 + 2), 16));
  };
  const cat    = (...a) => [].concat(...a);

  // HMAC-SHA256（GAS signed byte 変換）
  const hmac = (key, data) => Array.from(
    Utilities.computeHmacSha256Signature(
      data.map(b => b > 127 ? b - 256 : b),
      key.map(b => b > 127 ? b - 256 : b)
    )
  ).map(b => b & 0xFF);

  // SHA-256（string → ASCII bytes → hash）
  const sha256 = str => Array.from(
    Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      str.split('').map(c => c.charCodeAt(0) > 127 ? c.charCodeAt(0) - 256 : c.charCodeAt(0))
    )
  ).map(b => b & 0xFF);

  // RFC 6979 決定論的 k 生成
  const rfcK = (dB, hB) => {
    let V = new Array(32).fill(1);
    let K = new Array(32).fill(0);
    K = hmac(K, cat(V, [0], dB, hB));
    V = hmac(K, V);
    K = hmac(K, cat(V, [1], dB, hB));
    V = hmac(K, V);
    for (let i = 0; i < 100; i++) {
      V = hmac(K, V);
      const k = toBI(V);
      if (k >= BigInt(1) && k < n) return k;
      K = hmac(K, cat(V, [0]));
      V = hmac(K, V);
    }
    throw new Error('RFC 6979 failed');
  };

  // 署名本体
  const dB  = b64uDec(privKeyB64u);
  const hB  = sha256(message);
  const d   = toBI(dB);
  const z   = toBI(hB);
  const k   = rfcK(dB, hB);
  const [rx] = mul(k, [Gx, Gy]);
  const r   = mod(rx, n);
  const s   = mod(inv(k, n) * mod(z + r * d, n), n);

  // r || s (64 bytes) → base64url
  return Utilities.base64EncodeWebSafe(to32B(r).concat(to32B(s))).replace(/=+$/, '');
}

// ─── ユーティリティ ───────────────────────────────────
function _tryParse(str) {
  try { return typeof str === 'string' ? JSON.parse(str) : str; } catch (_) { return null; }
}
function _ok(obj)   { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
function _json(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }

// ======== テスト用ここから（完了後この関数と上記呼び出し1行を削除）========
function _resetLastRow(props, currentLastRow) {
  const lastDone = parseInt(props.getProperty(CONFIG.LAST_ROW_KEY), 10);
  if (currentLastRow < lastDone) {
    props.setProperty(CONFIG.LAST_ROW_KEY, String(currentLastRow));
  }
}
// ======== テスト用ここまで ========
