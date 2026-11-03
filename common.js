/* 共用：Google 登入、呼叫後端、小工具 */
(function () {
  var CFG = window.APP_CONFIG || {};
  var MOCK = /[?&]mock=1/.test(location.search);
  var TOKEN_KEY = 'checkin_id_token';

  function decodeJwt(t) {
    try {
      var p = t.split('.')[1].replace(/-/g, '+').replace(/_/g, '/');
      var json = decodeURIComponent(atob(p).split('').map(function (c) {
        return '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2);
      }).join(''));
      return JSON.parse(json);
    } catch (e) { return null; }
  }

  var Auth = {
    token: null,
    profile: null,
    onLogin: null,
    /** 初始化登入。opts: { buttonEl, hd, onLogin } */
    init: function (opts) {
      Auth.onLogin = opts.onLogin;
      if (MOCK) {
        var who = new URLSearchParams(location.search).get('as') || (opts.hd ? 's1@' + CFG.STUDENT_DOMAIN : 'teacher@gmail.com');
        Auth.set('mock.' + btoa(JSON.stringify({ email: who, name: who.split('@')[0], exp: Date.now() / 1000 + 3600 })) + '.x');
        return;
      }
      try {
        var saved = sessionStorage.getItem(TOKEN_KEY);
        if (saved && Auth.valid(saved)) { Auth.set(saved); return; }
      } catch (e) { /* 無痕模式等情況 */ }
      waitForGsi(function () {
        google.accounts.id.initialize({
          client_id: CFG.CLIENT_ID,
          callback: function (resp) { Auth.set(resp.credential); },
          auto_select: true,
          hd: opts.hd || undefined,
          ux_mode: 'popup'
        });
        google.accounts.id.renderButton(opts.buttonEl, { theme: 'filled_blue', size: 'large', text: 'signin_with', shape: 'pill', locale: 'zh-TW' });
        google.accounts.id.prompt();
      });
    },
    valid: function (t) {
      var p = decodeJwt(t);
      return p && p.exp * 1000 > Date.now() + 60000;
    },
    set: function (t) {
      Auth.token = t;
      Auth.profile = decodeJwt(t);
      try { sessionStorage.setItem(TOKEN_KEY, t); } catch (e) { /* ignore */ }
      if (Auth.onLogin) Auth.onLogin(Auth.profile);
    },
    logout: function () {
      try { sessionStorage.removeItem(TOKEN_KEY); } catch (e) { /* ignore */ }
      if (window.google && google.accounts) {
        google.accounts.id.disableAutoSelect();
      }
      location.href = location.pathname + (MOCK ? location.search : '');
    }
  };

  function waitForGsi(cb) {
    if (window.google && google.accounts && google.accounts.id) return cb();
    setTimeout(function () { waitForGsi(cb); }, 100);
  }

  /** 呼叫後端 API */
  function api(action, data) {
    if (Auth.token && !MOCK && !Auth.valid(Auth.token)) {
      toast('登入已過期，請重新登入', true);
      setTimeout(Auth.logout, 1500);
      return Promise.reject(new Error('登入已過期'));
    }
    var payload = { action: action, idToken: Auth.token, data: data || {} };
    var p = MOCK
      ? window.MockAPI.call(payload)
      : fetch(CFG.API_URL, { method: 'POST', body: JSON.stringify(payload), redirect: 'follow' })
          .then(function (r) { return r.json(); });
    return p.then(function (res) {
      if (!res.ok) {
        if (/過期|請先登入/.test(res.error)) setTimeout(Auth.logout, 1500);
        throw new Error(res.error);
      }
      return res.data;
    }, function (err) {
      if (err && err.message && !/過期/.test(err.message)) throw new Error('連線失敗，請檢查網路後再試（' + err.message + '）');
      throw err;
    });
  }

  var toastTimer;
  function toast(msg, isErr) {
    var el = document.getElementById('toast');
    if (!el) { el = document.createElement('div'); el.id = 'toast'; document.body.appendChild(el); }
    el.className = 'toast' + (isErr ? ' err' : '');
    el.textContent = msg;
    el.style.display = 'block';
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { el.style.display = 'none'; }, isErr ? 5000 : 2500);
  }

  function esc(s) {
    return String(s === undefined || s === null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }

  function $(sel, root) { return (root || document).querySelector(sel); }
  function $all(sel, root) { return Array.prototype.slice.call((root || document).querySelectorAll(sel)); }

  /** 讀檔為 base64；大圖片先壓縮到長邊 1600px */
  function readFile(file) {
    return new Promise(function (resolve, reject) {
      var isImg = /^image\/(jpeg|png|webp)$/.test(file.type);
      var reader = new FileReader();
      reader.onerror = function () { reject(new Error('讀取檔案失敗：' + file.name)); };
      reader.onload = function () {
        var dataUrl = reader.result;
        if (!isImg || file.size < 800 * 1024) return resolve(pack(file.name, file.type, dataUrl));
        var img = new Image();
        img.onload = function () {
          var max = 1600, w = img.width, h = img.height;
          var s = Math.min(1, max / Math.max(w, h));
          var c = document.createElement('canvas');
          c.width = Math.round(w * s); c.height = Math.round(h * s);
          c.getContext('2d').drawImage(img, 0, 0, c.width, c.height);
          var name = file.name.replace(/\.(png|webp|jpe?g)$/i, '') + '.jpg';
          resolve(pack(name, 'image/jpeg', c.toDataURL('image/jpeg', 0.85)));
        };
        img.onerror = function () { resolve(pack(file.name, file.type, dataUrl)); };
        img.src = dataUrl;
      };
      reader.readAsDataURL(file);
    });
    function pack(name, type, dataUrl) {
      return { name: name, type: type || 'application/octet-stream', base64: String(dataUrl).split(',')[1] || '' };
    }
  }

  window.App = { CFG: CFG, MOCK: MOCK, Auth: Auth, api: api, toast: toast, esc: esc, $: $, $all: $all, readFile: readFile };
  if (CFG.TITLE) document.title = CFG.TITLE + (document.title ? '｜' + document.title : '');
})();
