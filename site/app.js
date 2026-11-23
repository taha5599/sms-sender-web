'use strict';

/* SmsSender web 1.0 — talks to the local SmsSenderBridge.
   The Excel file and the message go straight from this page to the
   bridge on the user's own PC (http://127.0.0.1:18711). Render never
   sees them. If this page is served BY the bridge itself, calls are
   same-origin (BRIDGE = ''); otherwise the page calls 127.0.0.1.
   (R77) The bridge address is configurable on the page and saved in
   localStorage, so the page can reach a bridge on another machine
   (a LAN address or a tunnel URL) instead of this PC.
   Every /api/* call carries the X-SmsBridge header, which a cross-site
   form post cannot set — that is what keeps other websites from
   driving the modem. */

/* (R77) bridge address: the page can point at a bridge on another
   machine (a LAN address or a tunnel URL) instead of this PC. The
   address is saved in localStorage (BRIDGE_KEY) and restored on the
   next visit; empty means the old default - same-origin when this page
   is served by the bridge itself, otherwise http://127.0.0.1:18711.
   Every fetch reads the address fresh, so changing it takes effect
   on the next call. */
var BRIDGE_KEY = 'smsSenderBridge';
var BRIDGE_DEFAULT = 'http://127.0.0.1:18711';
/* (R78) Normalise a typed/saved bridge address to a bare origin:
   trim, prepend http:// when there is no scheme (R78 M2 - a bare
   host:port would otherwise fetch as a relative URL against the page
   origin), then keep only the origin (R78 M1 - a saved path would
   misroute every call to /path/api/...). (R79) Only http(s) schemes
   are usable by fetch - any other scheme returns '' so the caller
   falls back to the default. (R79) A bare host typed without a
   scheme and without a port gets the bridge port 18711; a value that
   already carried a scheme never gets a port appended. (R80) The
   18711 append applies only to a LOCAL bridge address (IP literal,
   localhost, a single-label LAN hostname or a .local name) - a bare
   public hostname keeps its default port, so a tunnel host typed
   without a scheme still works. (R80) A dangling colon with no port
   digits (e.g. 192.168.1.50:) counts as NO port. Returns ''
   when the value is empty or unparseable, so the caller falls back
   to the default. */
function normaliseBridgeAddress(raw) {
  var v = String(raw == null ? '' : raw).trim();
  if (!v) return '';
  var hadScheme = v.indexOf('://') !== -1;
  if (!hadScheme) v = 'http://' + v;
  var u;
  try { u = new URL(v); } catch (e) { return ''; }
  if (u.protocol !== 'http:' && u.protocol !== 'https:') return '';
  if (!hadScheme && !rawHasPort(raw) && isLocalBridgeHost(u.hostname)) u.port = '18711';
  return u.origin;
}
/* (R79) True when the raw (scheme-less) value already carries a port:
   a ':' in the authority after any userinfo, or after ']' for IPv6.
   (R80) A colon counts only when digits follow it: a dangling colon
   (e.g. 192.168.1.50:) is NOT a port, so the bridge port is appended. */
function rawHasPort(raw) {
  var auth = String(raw).trim();
  var cut = auth.search(/[\/?#]/);
  if (cut !== -1) auth = auth.slice(0, cut);
  var at = auth.lastIndexOf('@');
  if (at !== -1) auth = auth.slice(at + 1);
  if (auth.charAt(0) === '[') {
    var close = auth.indexOf(']');
    if (close === -1) return false;
    return /^:\d/.test(auth.slice(close + 1));
  }
  return /:\d/.test(auth);
}
/* (R80) True when the parsed hostname is a LOCAL bridge address: an IP
   literal (v4/v6), localhost, a single-label LAN hostname, or a .local
   (mDNS) name. A bare public hostname (e.g. a tunnel host typed without
   a scheme) is NOT local, so it keeps its default port. (R81) Trailing
   dots are stripped first: the fully-qualified root form (mypc.,
   mybridge.local.) is local too, not just localhost. */
function isLocalBridgeHost(h) {
  h = String(h || '').toLowerCase().replace(/\.+$/, '');
  if (h === 'localhost') return true;
  if (h.charAt(0) === '[' && h.charAt(h.length - 1) === ']') return true;
  if (/^(?:\d{1,3}\.){3}\d{1,3}$/.test(h)) return true;
  if (h.indexOf('.') === -1) return true;
  return h.slice(-6) === '.local';
}
function getBridgeBase() {
  var saved = '';
  try { saved = localStorage.getItem(BRIDGE_KEY) || ''; } catch (e) { /* storage unavailable */ }
  var base = normaliseBridgeAddress(saved);
  if (base) return base;
  return (location.hostname === '127.0.0.1' || location.hostname === 'localhost')
    ? '' : BRIDGE_DEFAULT;
}
/* (R78) The stored value is always the normalised origin (or ''), so
   what is saved, what is fetched and what is displayed agree. */
function saveBridgeAddress() {
  try { localStorage.setItem(BRIDGE_KEY, normaliseBridgeAddress($('bridgeAddressInput').value)); } catch (e) { /* storage unavailable */ }
}
/* (R78) The input shows the EFFECTIVE base: the normalised origin, or
   the page's own origin on the bridge-served page, where the effective
   base is '' (same-origin) - never a stale raw value or the 127.0.0.1
   default while the page is served from localhost. */
function displayBridgeBase() {
  var base = getBridgeBase();
  return base || location.origin;
}
function restoreBridgeAddress() {
  $('bridgeAddressInput').value = displayBridgeBase();
}
/* (R78) Mixed content: from an https page the browser blocks fetches to
   an http:// public host (only loopback is exempt), so warn visibly
   instead of failing silently as "الوسيط لا يعمل". */
function isBridgeMixedContent() {
  if (location.protocol !== 'https:') return false;
  var base = getBridgeBase();
  if (!base || base.indexOf('http://') !== 0) return false;
  var host = '';
  try { host = new URL(base).hostname; } catch (e) { return false; }
  return host !== '127.0.0.1' && host !== 'localhost' && host !== '[::1]' && host !== '::1';
}
function updateBridgeWarning() {
  var el = $('bridgeMixedWarning');
  if (el) el.classList.toggle('hidden', !isBridgeMixedContent());
}
var BRIDGE_HEADER = 'X-SmsBridge';
var BRIDGE_TOKEN = 'sms-web-1.0';

var modems = [];          // {port, label, detected, checked, reason}
var previewNumbers = [];  // valid numbers from the last preview
var currentJobId = null;
var pollTimer = null;

function $(id) { return document.getElementById(id); }

function escapeHtml(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

/* ---------------- bridge calls ---------------- */

function api(path, opts) {
  opts = opts || {};
  var headers = {};
  headers[BRIDGE_HEADER] = BRIDGE_TOKEN;
  if (opts.headers) {
    for (var k in opts.headers) headers[k] = opts.headers[k];
  }
  return fetch(getBridgeBase() + path, {
    method: opts.method || 'GET',
    headers: headers,
    body: opts.body || undefined
  }).then(function (res) {
    return res.json().then(function (data) {
      // (R77W) C-76W-L1: a null / non-object / array body is a malformed
      // response, not an empty result - reject it with a friendly Arabic
      // error instead of leaking a JS engine TypeError (or silently
      // rendering an empty inbox) to every api() caller.
      if (data === null || typeof data !== 'object' || Array.isArray(data)) {
        var bad = new Error('رد الوسيط غير صالح');
        bad.status = res.status;
        bad.data = data;
        throw bad;
      }
      if (!res.ok || data.ok === false) {
        var msg = (data && data.error) ? data.error : ('خطأ ' + res.status);
        var err = new Error(msg);
        err.status = res.status;
        err.data = data;
        throw err;
      }
      return data;
    }, function () {
      // (R78W) C-77W-L1: res.json() itself rejects on a non-JSON or
      // empty body (HTML error page from a broken proxy, truncated
      // response) - convert it to the same friendly Arabic error
      // instead of leaking the SyntaxError verbatim. Rejection
      // handler (not .catch) so the guard's own error and the
      // bridge's error message above are never swallowed.
      var bad = new Error('رد الوسيط غير صالح');
      bad.status = res.status;
      throw bad;
    });
  });
}

function setBridgeStatus(state, text) {
  var el = $('bridgeStatus');
  el.className = 'status ' + state;
  el.textContent = text;
  $('bridgeHelp').classList.toggle('hidden', state === 'ok');
}

function checkBridge(autoDetect) {
  setBridgeStatus('unknown', 'جارٍ التحقق من الوسيط…');
  return api('/api/status').then(function (data) {
    setBridgeStatus('ok', 'الوسيط شغال — الإصدار ' + (data.version || ''));
    if (autoDetect) return detectModems();
  }).catch(function () {
    setBridgeStatus('down', 'الوسيط لا يعمل');
  });
}

/* ---------------- modem list ---------------- */

function upsertModem(port, label, detected, checked, reason) {
  var m = null;
  for (var i = 0; i < modems.length; i++) {
    if (modems[i].port === port) { m = modems[i]; break; }
  }
  if (!m) {
    m = { port: port, label: label || port, detected: !!detected, checked: !!checked, reason: reason || '' };
    modems.push(m);
  } else {
    if (label) m.label = label;
    if (detected !== undefined) m.detected = detected;
    if (checked !== undefined) m.checked = checked;
    if (reason !== undefined) m.reason = reason || '';
  }
  return m;
}

function renderModemList() {
  var list = $('modemList');
  if (modems.length === 0) {
    list.innerHTML = '<p class="empty">لا توجد مودمات بعد — اضغط «كشف تلقائي».</p>';
  } else {
    var html = '';
    for (var i = 0; i < modems.length; i++) {
      var m = modems[i];
      // (Round 70W) B-65W-L1: an answered-but-excluded port is NOT "لم يرد" —
      // it gets the distinct tag "مستبعد" plus the bridge's reason. A port with
      // no reason at all is the one that truly did not reply.
      var tag = 'لم يرد';
      if (m.detected) tag = 'مودم';
      else if (m.reason) tag = 'مستبعد';
      html += '<label class="modem-row' + (m.detected ? ' detected' : '') + (m.reason ? ' excluded' : '') + '">'
        + '<input type="checkbox" data-port="' + escapeHtml(m.port) + '"'
        + (m.checked ? ' checked' : '') + '>'
        + '<span class="modem-port">' + escapeHtml(m.port) + '</span>'
        + '<span class="modem-label">' + escapeHtml(m.label) + '</span>'
        + '<span class="modem-tag">' + tag + '</span>'
        + (m.reason ? '<span class="modem-reason">' + escapeHtml(m.reason) + '</span>' : '')
        + '</label>';
    }
    list.innerHTML = html;
  }
  var anyChecked = modems.some(function (m) { return m.checked; });
  $('btnTest').disabled = !anyChecked;
  $('chkAll').checked = modems.length > 0 && modems.every(function (m) { return m.checked; });
  updateSendButton();
}

$('modemList').addEventListener('change', function (ev) {
  var port = ev.target.getAttribute && ev.target.getAttribute('data-port');
  if (!port) return;
  for (var i = 0; i < modems.length; i++) {
    if (modems[i].port === port) modems[i].checked = ev.target.checked;
  }
  renderModemList();
});

$('chkAll').addEventListener('change', function (ev) {
  modems.forEach(function (m) { m.checked = ev.target.checked; });
  renderModemList();
});

function loadPorts() {
  return api('/api/ports').then(function (data) {
    (data.ports || []).forEach(function (p) {
      upsertModem(p.port, p.display || p.description || p.port, undefined, undefined);
    });
    renderModemList();
  }).catch(function (err) {
    showError('تعذر تحديث المنافذ: ' + err.message);
  });
}

function detectModems() {
  return api('/api/detect', { method: 'POST' }).then(function (data) {
    (data.results || []).forEach(function (r) {
      var usable = r.kind === 'Answered' && !r.isHuaweiData && !r.isNmea && !r.isHuaweiModem;
      // (Round 70W) B-65W-L2: prefer the bridge's display name (WMI description)
      // over the ATI line; fall back to the ATI label only when display is empty.
      var label = r.display || (r.label ? (r.label + ' — ' + r.port) : r.port);
      upsertModem(r.port, label, usable, usable ? true : undefined, r.reason || '');
    });
    (data.modems || []).forEach(function (port) {
      upsertModem(port, undefined, true, true);
    });
    renderModemList();
    if (!data.modems || data.modems.length === 0) {
      showError('لم يعثر الكشف التلقائي على أي مودم يرد على AT. تأكد من توصيل الفلاشة وإغلاق برنامج الإنترنت الخاص بها (مثل Mobile Partner).');
    } else {
      clearError();
    }
  }).catch(function (err) {
    showError('تعذر الكشف التلقائي: ' + err.message);
  });
}

function testSelected() {
  var ports = checkedPorts();
  if (ports.length === 0) { showError('علّم مودماً واحداً على الأقل للاختبار.'); return; }
  $('testResults').innerHTML = '<p class="hint">جارٍ اختبار ' + ports.join('، ') + '…</p>';
  api('/api/test', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ ports: ports.join(',') })
  }).then(function (data) {
    var html = '';
    (data.results || []).forEach(function (r) {
      html += '<div class="test-row ' + (r.ok ? 'ok' : 'bad') + '">'
        + '<strong>' + escapeHtml(r.port) + '</strong> — '
        + (r.ok
          ? ('الشريحة: ' + escapeHtml(r.sim || '؟') + ' — الإشارة: ' + escapeHtml(r.signal || '؟'))
          : ('فشل: ' + escapeHtml(r.error || 'لا يوجد رد')))
        + '</div>';
    });
    $('testResults').innerHTML = html;
  }).catch(function (err) {
    $('testResults').innerHTML = '<div class="test-row bad">فشل الاختبار: ' + escapeHtml(err.message) + '</div>';
  });
}

function checkedPorts() {
  return modems.filter(function (m) { return m.checked; }).map(function (m) { return m.port; });
}

/* ---------------- preview ---------------- */

function selectedFile() {
  var input = $('fileInput');
  return input.files && input.files[0] ? input.files[0] : null;
}

function previewFile() {
  var file = selectedFile();
  if (!file) { showError('اختر ملف Excel أو CSV أولاً.'); return; }
  var fd = new FormData();
  fd.append('file', file);
  $('btnPreview').disabled = true;
  api('/api/preview', { method: 'POST', body: fd }).then(function (data) {
    previewNumbers = data.numbers || [];
    $('previewCounts').innerHTML =
      '<span>إجمالي الخلايا: <strong>' + data.total + '</strong></span>'
      + '<span>صالحة: <strong>' + data.valid + '</strong></span>'
      + '<span>مكررة: <strong>' + data.duplicates + '</strong></span>'
      + '<span>غير صالحة: <strong>' + data.invalid + '</strong></span>'
      + (data.sheet ? '<span>الورقة: <strong>' + escapeHtml(data.sheet) + '</strong></span>' : '');
    var tbody = $('previewTable').querySelector('tbody');
    var html = '';
    previewNumbers.forEach(function (n, i) {
      html += '<tr><td>' + (i + 1) + '</td><td dir="ltr">' + escapeHtml(n) + '</td><td>صالح</td></tr>';
    });
    tbody.innerHTML = html;
    $('previewTable').classList.toggle('hidden', previewNumbers.length === 0);
    updateSendButton();
    clearError();
  }).catch(function (err) {
    showError('تعذرت المعاينة: ' + err.message);
    previewNumbers = [];
    updateSendButton();
  }).finally(function () {
    $('btnPreview').disabled = false;
  });
}

/* ---------------- send / progress ---------------- */

function updateSendButton() {
  $('btnSend').disabled = !(previewNumbers.length > 0 && checkedPorts().length > 0 && !currentJobId);
}

function sendJob() {
  var file = selectedFile();
  var message = $('messageInput').value.trim();
  var ports = checkedPorts();
  if (!file) { showError('اختر ملف Excel أو CSV أولاً.'); return; }
  if (previewNumbers.length === 0) { showError('اعرض معاينة الأرقام أولاً.'); return; }
  if (!message) { showError('اكتب نص الرسالة أولاً.'); return; }
  if (ports.length === 0) { showError('علّم مودماً واحداً على الأقل.'); return; }
  var fd = new FormData();
  fd.append('file', file);
  fd.append('message', message);
  fd.append('delaySeconds', getDelaySeconds());
  fd.append('ports', ports.join(','));
  clearError();
  // (Round 61) C-60-M2/B-60-M1: disable the send button BEFORE the request, not in .then - a double-click fired
  // two requests inside the server's probe window and both created jobs. Re-enabled in .catch via updateSendButton.
  $('btnSend').disabled = true;
  api('/api/send', { method: 'POST', body: fd }).then(function (data) {
    currentJobId = data.jobId;
    $('btnSend').disabled = true;
    $('btnStop').classList.remove('hidden');
    $('progressArea').classList.remove('hidden');
    $('logBox').classList.remove('hidden');
    $('resultsTable').classList.add('hidden');
    $('logBox').innerHTML = '';
    if (data.excluded && data.excluded.length > 0) {
      var ex = data.excluded.map(function (e) { return e.port + ' (' + e.reason + ')'; }).join('، ');
      showError('مودمات مستبعدة قبل البدء: ' + ex);
    }
    pollTimer = setInterval(pollProgress, 1000);
    pollProgress();
  }).catch(function (err) {
    showError('تعذر بدء الإرسال: ' + err.message);
    updateSendButton(); // (Round 61) C-60-M2: re-enable after a failed start (the button was disabled before the request).
    if (err.data && err.data.excluded) {
      var ex = err.data.excluded.map(function (e) { return e.port + ' (' + e.reason + ')'; }).join('، ');
      showError('تعذر بدء الإرسال: ' + err.message + ' — ' + ex);
    }
  });
}

function pollProgress() {
  if (!currentJobId) return;
  api('/api/progress?job=' + encodeURIComponent(currentJobId)).then(function (data) {
    var total = data.total || 0;
    // (Round 61) C-60-M1/A-60-L1: stopped rows count toward done, so the bar reaches 100% on a stop.
    var done = (data.sent || 0) + (data.failed || 0) + (data.cancelled || 0) + (data.unconfirmed || 0); // (Round 80P) P2: unconfirmed rows count toward done, so the bar reaches 100%.
    var pct = total > 0 ? Math.round(done / total * 100) : 0;
    $('progressFill').style.width = pct + '%';
    $('progressText').textContent = done + ' / ' + total + ' — نجح ' + (data.sent || 0) + '، فشل ' + (data.failed || 0) + ((data.cancelled || 0) > 0 ? '، أُوقف ' + data.cancelled : '') + ((data.unconfirmed || 0) > 0 ? '، لم يتم تأكيد ' + data.unconfirmed + ' (قد تكون أُرسلت — تحقق قبل إعادة الإرسال)' : '') + ((data.unconfirmedPartial || 0) > 0 ? ' (منها ' + data.unconfirmedPartial + ' أُرسلت مقاطع قبل عدم التأكيد — إعادة الإرسال ستعيد إرسال المقاطع التي تم إرسالها)' : '') + (data.stopped ? ' — تم الإيقاف' : ''); // (Round 80P) P2: unconfirmed totals clause, mirrors the desktop summary.
    var mh = '';
    (data.modems || []).forEach(function (m) {
      var state = m.status === 'dead' ? ' — توقف: ' + m.reason
        : m.status === 'excluded' ? ' — مستبعد: ' + m.reason
        : m.status === 'done' ? ' — انتهى'
        : (m.current ? ' — يرسل ' + m.current : ' — جارٍ');
      mh += '<div class="modem-progress ' + m.status + '"><strong>' + escapeHtml(m.port) + '</strong>: نجح '
        + m.sent + '، فشل ' + m.failed + ((m.cancelled || 0) > 0 ? '، أُوقف ' + m.cancelled : '') + ((m.unconfirmed || 0) > 0 ? '، لم يتم تأكيد ' + m.unconfirmed : '') + escapeHtml(state) + '</div>'; // (Round 80P) P2: per-modem unconfirmed, mirrors the desktop per-modem line. // (Round 62) B-61-L1: per-modem stopped rows read "أُوقف", not "فشل".
    });
    $('modemProgress').innerHTML = mh;
    $('logBox').innerHTML = (data.log || []).map(escapeHtml).join('<br>');
    $('logBox').scrollTop = $('logBox').scrollHeight;
    var tbody = $('resultsTable').querySelector('tbody');
    var rh = '';
    (data.results || []).forEach(function (r, i) {
      rh += '<tr><td>' + (i + 1) + '</td><td dir="ltr">' + escapeHtml(r.number) + '</td>'
        + '<td>' + escapeHtml(r.port || '—') + '</td>'
        + '<td>' + (r.ok ? 'تم' : (r.cancelled ? 'تم الإيقاف' : (r.unconfirmed ? 'لم يتم التأكيد' : 'فشل'))) + '</td>' // (Round 61) C-60-M1: stopped rows read "تم الإيقاف", not "فشل".
        + '<td>' + escapeHtml(r.message || '') + '</td></tr>';
    });
    tbody.innerHTML = rh;
    $('resultsTable').classList.toggle('hidden', (data.results || []).length === 0);
    if (data.done) finishJob(data);
  }).catch(function (err) {
    showError('تعذر متابعة التقدم: ' + err.message);
  });
}

function finishJob(data) {
  clearInterval(pollTimer);
  pollTimer = null;
  currentJobId = null;
  $('btnStop').classList.add('hidden');
  updateSendButton();
  if (data.fatalError) showError(data.fatalError);
}

function stopJob() {
  if (!currentJobId) return;
  api('/api/stop', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ jobId: currentJobId })
  }).catch(function (err) { showError('تعذر الإيقاف: ' + err.message); });
}

/* ---------------- inbox (R75W) ---------------- */

/* Read-only inbox: a fresh GET /api/inbox read on every refresh (the
   bridge reads AT+CMGL=4 live each time — nothing cached, nothing
   deleted or marked read). Reuses api() so the call carries the
   X-SmsBridge header and the configured bridge address. */
/* (Round 76W) C-L2: Arabic plural for the inbox count. */
function inboxCountText(n) {
  if (n === 1) return '1 رسالة';
  if (n === 2) return '2 رسالتان';
  return n + ' رسائل';
}

function loadInbox() {
  var btn = $('btnRefreshInbox');
  btn.disabled = true;
  $('inboxStatus').textContent = 'جارٍ قراءة الوارد…';
  $('inboxError').classList.add('hidden');
  api('/api/inbox').then(function (data) {
    /* (Round 76W) C-L3: a missing messages field means an empty inbox
       (as before); a present-but-not-array payload is malformed and
       fails with a friendly Arabic error, never a raw JS engine
       message; null entries inside the array are skipped. */
    var raw = (data && data.messages !== undefined) ? data.messages : [];
    if (!Array.isArray(raw)) throw new Error('رد الوسيط غير صالح');
    var msgs = raw.filter(function (m) { return m !== null && typeof m === 'object'; });
    /* (Round 76W) C-L2: correct Arabic plural, and the port only when present. */
    $('inboxStatus').textContent = msgs.length === 0 ? '' : inboxCountText(msgs.length) + ((data && data.port) ? ' — المنفذ ' + data.port : '');
    if (msgs.length === 0) {
      $('inboxList').innerHTML = '<p class="empty">لا توجد رسائل واردة على الشريحة.</p>';
    } else {
      var html = '';
      msgs.forEach(function (m) {
        html += '<div class="inbox-msg"><div class="inbox-head">'
          + '<span class="inbox-from" dir="ltr">' + escapeHtml(m.from) + '</span>'
          + '<span class="inbox-time" dir="ltr">' + escapeHtml(m.time) + '</span>'
          + '<span class="inbox-status">' + escapeHtml(m.status) + '</span></div>'
          + '<div class="inbox-text">' + escapeHtml(m.text) + '</div></div>';
      });
      $('inboxList').innerHTML = html;
    }
  }).catch(function (err) {
    $('inboxError').textContent = 'تعذرت قراءة الوارد: ' + err.message;
    $('inboxError').classList.remove('hidden');
    $('inboxStatus').textContent = '';
    /* (Round 76W) C-L1: never present a stale list as current after a
       failed refresh - clear it and say the last read failed. */
    $('inboxList').innerHTML = '<p class="empty">لا توجد بيانات واردة معروضة — آخر محاولة قراءة فشلت.</p>';
  }).finally(function () {
    btn.disabled = false;
  });
}

/* ---------------- misc ---------------- */

function showError(text) {
  var el = $('sendError');
  el.textContent = text;
  el.classList.remove('hidden');
}

function clearError() {
  var el = $('sendError');
  el.textContent = '';
  el.classList.add('hidden');
}

function updateCharCount() {
  var text = $('messageInput').value;
  var isArabic = /[\u0600-\u06FF]/.test(text);
  var perSeg = isArabic ? 70 : 160;
  var segs = text.length === 0 ? 0 : Math.ceil(text.length / perSeg);
  $('charCount').textContent = text.length === 0 ? ''
    : text.length + ' حرف — حوالي ' + segs + ' رسالة لكل رقم (' + (isArabic ? 'عربي' : 'لاتيني') + ')';
}

/* ---------------- wiring ---------------- */

/* (R76) delay: the user types the seconds or taps a preset.
   (R76b) No upper cap: 2 seconds minimum, up to 86400 (24h) — the
   86400 is only the overflow guard, mirroring the bridge.
   Clamp lives in getDelaySeconds() — called at send time, on change
   (blur/Enter) and on restore — NOT on every keystroke, so typing "12"
   is never rewritten to "2" mid-typing. The value is saved to
   localStorage raw on input (same pattern as the message, W1) and
   restored on the next visit. */
var DELAY_KEY = 'smsSenderDelay';
function getDelaySeconds() {
  var v = parseInt($('delayInput').value, 10);
  if (isNaN(v)) v = 5;
  if (v < 2) v = 2;
  if (v > 86400) v = 86400;
  $('delayInput').value = v;
  return v;
}
function saveDelay() {
  try { localStorage.setItem(DELAY_KEY, $('delayInput').value); } catch (e) { /* storage unavailable */ }
}
function restoreDelay() {
  try {
    var saved = localStorage.getItem(DELAY_KEY);
    if (saved !== null && saved !== '') { $('delayInput').value = saved; getDelaySeconds(); }
  } catch (e) { /* storage unavailable */ }
}

$('btnRecheck').addEventListener('click', function () { checkBridge(true); });
$('btnRefreshPorts').addEventListener('click', loadPorts);
$('btnDetect').addEventListener('click', detectModems);
$('btnTest').addEventListener('click', testSelected);
$('btnRefreshInbox').addEventListener('click', loadInbox);
$('btnPreview').addEventListener('click', previewFile);
$('btnSend').addEventListener('click', sendJob);
$('btnStop').addEventListener('click', stopJob);
/* (R65) W1: message persistence — the message the user wrote is saved
   to this browser's localStorage on every keystroke and restored on the
   next visit, so reopening the page finds it. First visit stays EMPTY:
   nothing is prefilled; only the user's own saved text comes back.
   Storage failures are ignored — sending still works. */
var MESSAGE_KEY = 'smsSenderMessage';
function saveMessage() {
  try { localStorage.setItem(MESSAGE_KEY, $('messageInput').value); } catch (e) { /* storage unavailable */ }
}
function restoreMessage() {
  try {
    var saved = localStorage.getItem(MESSAGE_KEY);
    if (saved) { $('messageInput').value = saved; updateCharCount(); }
  } catch (e) { /* storage unavailable */ }
}

$('messageInput').addEventListener('input', updateCharCount);
$('messageInput').addEventListener('input', saveMessage);
restoreMessage();
$('delayInput').addEventListener('input', saveDelay);
$('delayInput').addEventListener('change', function () { getDelaySeconds(); saveDelay(); });
document.querySelectorAll('.delay-presets button').forEach(function (b) {
  b.addEventListener('click', function () {
    $('delayInput').value = b.getAttribute('data-delay');
    getDelaySeconds();
    saveDelay();
  });
});
restoreDelay();
$('fileInput').addEventListener('change', function () {
  previewNumbers = [];
  $('previewCounts').innerHTML = '';
  $('previewTable').classList.add('hidden');
  $('btnPreview').disabled = !selectedFile();
  updateSendButton();
});

$('bridgeAddressInput').addEventListener('input', saveBridgeAddress);
$('bridgeAddressInput').addEventListener('input', updateBridgeWarning);
$('bridgeAddressInput').addEventListener('change', function () { saveBridgeAddress(); restoreBridgeAddress(); updateBridgeWarning(); checkBridge(true); });
restoreBridgeAddress();
updateBridgeWarning();

checkBridge(true);
