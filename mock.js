/*
 * 本機預覽用：在網址加上 ?mock=1 時，於瀏覽器內模擬 Google Apps Script，
 * 直接執行 ../gas/Code.gs 的程式碼，資料存在這台電腦的瀏覽器裡。
 * 正式使用時不會用到這個檔案（沒有 ?mock=1 就不會啟動）。
 *   學生預覽：index.html?mock=1&as=s1110101@student.ksvs.kh.edu.tw
 *   教師預覽：admin.html?mock=1
 */
(function () {
  if (!/[?&]mock=1/.test(location.search)) return;
  var CFG = window.APP_CONFIG || {};
  var STORE = 'checkin_mock_db';
  var GAS_PATH = CFG.MOCK_GAS_PATH || (/\/web\//.test(location.pathname) ? '../gas/Code.gs' : 'gas/Code.gs');

  function load() {
    try {
      return JSON.parse(localStorage.getItem(STORE) || 'null', function (k, v) {
        return v && typeof v === 'object' && v.__d ? new Date(v.__d) : v;
      });
    } catch (e) { return null; }
  }
  var db = load() || { sheets: {}, order: [], props: {}, fmt: {}, files: 0 };
  function save() {
    var copy = JSON.stringify(db, function (k, v) {
      var raw = this[k];
      return raw instanceof Date ? { __d: raw.toISOString() } : v;
    });
    try { localStorage.setItem(STORE, copy); } catch (e) { /* ignore */ }
  }

  function colIndex(letters) { var n = 0; for (var i = 0; i < letters.length; i++) n = n * 26 + letters.charCodeAt(i) - 64; return n; }

  function coerce(sheetName, col, v) {
    var textCols = db.fmt[sheetName] || [];
    if (typeof v !== 'string' || textCols.indexOf(col) >= 0) return v;
    if (/^\d{4}-\d{2}-\d{2}$/.test(v)) { var p = v.split('-'); return new Date(+p[0], +p[1] - 1, +p[2]); }
    if (/^-?\d+(\.\d+)?$/.test(v) && !/^0\d/.test(v)) return Number(v);
    return v;
  }

  function Sheet(name) {
    var rows = function () { return db.sheets[name]; };
    var self = {
      getName: function () { return name; },
      getLastRow: function () { return rows().length; },
      setFrozenRows: function () {},
      getDataRange: function () {
        return { getValues: function () { return rows().map(function (r) { return r.slice(); }); } };
      },
      appendRow: function (vals) { rows().push(vals.map(function (v, i) { return coerce(name, i + 1, v); })); save(); },
      deleteRow: function (r) { rows().splice(r - 1, 1); save(); },
      getRange: function (r, c, nr, nc) {
        if (typeof r === 'string') {
          var m = /^([A-Z]+):([A-Z]+)$/.exec(r);
          var from = colIndex(m[1]), to = colIndex(m[2]);
          return {
            setNumberFormat: function (f) {
              if (f !== '@') return;
              var arr = db.fmt[name] = db.fmt[name] || [];
              for (var i = from; i <= to; i++) if (arr.indexOf(i) < 0) arr.push(i);
              save();
            }
          };
        }
        nr = nr || 1; nc = nc || 1;
        var rng = {
          setValues: function (vals) {
            for (var i = 0; i < nr; i++) {
              while (rows().length < r + i) rows().push([]);
              var row = rows()[r + i - 1];
              for (var j = 0; j < nc; j++) row[c + j - 1] = coerce(name, c + j, vals[i][j]);
            }
            save(); return rng;
          },
          setValue: function (v) { return rng.setValues([[v]]); },
          clearContent: function () {
            for (var i = 0; i < nr; i++) { var row = rows()[r + i - 1]; if (row) for (var j = 0; j < nc; j++) row[c + j - 1] = ''; }
            // 移除結尾空白列
            while (rows().length && rows()[rows().length - 1].every(function (x) { return x === '' || x === undefined; })) rows().pop();
            save(); return rng;
          },
          setFontWeight: function () { return rng; },
          setBackground: function () { return rng; }
        };
        return rng;
      }
    };
    return self;
  }

  var ss = {
    getId: function () { return 'mock-sheet'; },
    getUrl: function () { return 'about:blank#mock-spreadsheet'; },
    getSheetByName: function (n) { return db.sheets[n] ? Sheet(n) : null; },
    insertSheet: function (n) { db.sheets[n] = []; db.order.push(n); save(); return Sheet(n); },
    getSheets: function () { return db.order.map(Sheet); },
    deleteSheet: function (sh) { delete db.sheets[sh.getName()]; db.order.splice(db.order.indexOf(sh.getName()), 1); save(); }
  };

  function pad(n) { return ('0' + n).slice(-2); }
  function b64(str) { return btoa(unescape(encodeURIComponent(str))).replace(/\+/g, '-').replace(/\//g, '_'); }
  var cache = {};
  function Folder(id, name) {
    return {
      getId: function () { return id; },
      getFoldersByName: function (n) { var done = false; return { hasNext: function () { return !done && false; }, next: function () { done = true; } }; },
      createFolder: function (n) { return Folder(id + '/' + n, n); },
      createFile: function (blob) { db.files++; save(); return { getUrl: function () { return 'https://drive.google.com/mock/' + db.files + '/' + encodeURIComponent(blob.name); } }; }
    };
  }

  var stubs = {
    SpreadsheetApp: { getActiveSpreadsheet: function () { return ss; }, openById: function () { return ss; } },
    PropertiesService: { getScriptProperties: function () {
      return { getProperty: function (k) { return db.props[k] || null; }, setProperty: function (k, v) { db.props[k] = v; save(); } };
    } },
    CacheService: { getScriptCache: function () {
      return { get: function (k) { return cache[k] || null; }, put: function (k, v) { cache[k] = v; } };
    } },
    LockService: { getScriptLock: function () { return { tryLock: function () { return true; }, releaseLock: function () {} }; } },
    Utilities: {
      DigestAlgorithm: { SHA_256: 'sha256' },
      Charset: { UTF_8: 'utf8' },
      computeDigest: function (alg, s) { return s; },
      base64EncodeWebSafe: function (s) { return b64(String(s)); },
      base64Decode: function (s) { return { length: Math.floor(String(s).length * 0.75) }; },
      newBlob: function (bytes, type, name) { return { name: name, type: type }; },
      formatDate: function (d, tz, p) {
        return p.replace('yyyy', d.getFullYear()).replace('MM', pad(d.getMonth() + 1)).replace('dd', pad(d.getDate()))
          .replace('HH', pad(d.getHours())).replace('mm', pad(d.getMinutes())).replace('ss', pad(d.getSeconds()));
      }
    },
    Session: { getScriptTimeZone: function () { return 'Asia/Taipei'; }, getEffectiveUser: function () { return { getEmail: function () { return 'teacher@gmail.com'; } }; } },
    UrlFetchApp: { fetch: function (url) {
      var tok = decodeURIComponent(url.split('id_token=')[1] || '');
      var info = JSON.parse(atob(tok.split('.')[1]));
      info.aud = db.props.CLIENT_ID; info.email_verified = 'true';
      return { getResponseCode: function () { return 200; }, getContentText: function () { return JSON.stringify(info); } };
    } },
    DriveApp: { getFolderById: function (id) { return Folder(id); }, createFolder: function (n) { return Folder('root/' + n, n); }, getRootFolder: function () { return Folder('root'); } },
    ContentService: { MimeType: { JSON: 'json' }, createTextOutput: function (s) { return { content: s, setMimeType: function () { return this; } }; } },
    Logger: { log: function (m) { console.log('[GAS]', m); } }
  };

  var gasPromise = null;
  function gas() {
    if (gasPromise) return gasPromise;
    gasPromise = fetch(GAS_PATH).then(function (r) { return r.text(); }).then(function (code) {
      var names = Object.keys(stubs);
      var fn = new Function(names.join(','), code + '\n;return { doPost: doPost, setup: setup };');
      var g = fn.apply(null, names.map(function (n) { return stubs[n]; }));
      if (!db.order.length) {
        g.setup();
        db.props.TEACHER_EMAILS = 'teacher@gmail.com';
        db.props.CLIENT_ID = 'mock-client';
        save();
      }
      return g;
    });
    return gasPromise;
  }

  window.MockAPI = {
    call: function (payload) {
      return gas().then(function (g) {
        return new Promise(function (res) { setTimeout(res, 150); }).then(function () {
          var out = g.doPost({ postData: { contents: JSON.stringify(payload) } });
          return JSON.parse(out.content);
        });
      });
    },
    reset: function () { localStorage.removeItem(STORE); location.reload(); }
  };
})();
