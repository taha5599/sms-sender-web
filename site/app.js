'use strict';

/* SmsSender web 1.0 — talks to the local SmsSenderBridge.
   The Excel file and the message go straight from this page to the
   bridge on the user's own PC (http://127.0.0.1:18711). Render never
   sees them. If this page is served BY the bridge itself, calls are
   same-origin (BRIDGE = ''); otherwise the page calls 127.0.0.1.
   Every /api/* call carries the X-SmsBridge header, which a cross-site
   form post cannot set — that is what keeps other websites from
   driving the modem. */

var BRIDGE = (location.hostname === '127.0.0.1' || location.hostname === 'localhost')
  ? '' : 'http://127.0.0.1:18711';
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
  return fetch(BRIDGE + path, {
    method: opts.method || 'GET',
    headers: headers,
    body: opts.body || undefined
  }).then(function (res) {
    return res.json().then(function (data) {
      if (!res.ok || data.ok === false) {
        var msg = (data && data.error) ? data.error : ('خطأ ' + res.status);
        var err = new Error(msg);
        err.status = res.status;
        err.data = data;
        throw err;
      }
      return data;
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
  fd.append('delaySeconds', $('delaySelect').value);
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
    var done = (data.sent || 0) + (data.failed || 0) + (data.cancelled || 0);
    var pct = total > 0 ? Math.round(done / total * 100) : 0;
    $('progressFill').style.width = pct + '%';
    $('progressText').textContent = done + ' / ' + total + ' — نجح ' + (data.sent || 0) + '، فشل ' + (data.failed || 0) + ((data.cancelled || 0) > 0 ? '، أُوقف ' + data.cancelled : '') + (data.stopped ? ' — تم الإيقاف' : '');
    var mh = '';
    (data.modems || []).forEach(function (m) {
      var state = m.status === 'dead' ? ' — توقف: ' + m.reason
        : m.status === 'excluded' ? ' — مستبعد: ' + m.reason
        : m.status === 'done' ? ' — انتهى'
        : (m.current ? ' — يرسل ' + m.current : ' — جارٍ');
      mh += '<div class="modem-progress ' + m.status + '"><strong>' + escapeHtml(m.port) + '</strong>: نجح '
        + m.sent + '، فشل ' + m.failed + ((m.cancelled || 0) > 0 ? '، أُوقف ' + m.cancelled : '') + escapeHtml(state) + '</div>'; // (Round 62) B-61-L1: per-modem stopped rows read "أُوقف", not "فشل".
    });
    $('modemProgress').innerHTML = mh;
    $('logBox').innerHTML = (data.log || []).map(escapeHtml).join('<br>');
    $('logBox').scrollTop = $('logBox').scrollHeight;
    var tbody = $('resultsTable').querySelector('tbody');
    var rh = '';
    (data.results || []).forEach(function (r, i) {
      rh += '<tr><td>' + (i + 1) + '</td><td dir="ltr">' + escapeHtml(r.number) + '</td>'
        + '<td>' + escapeHtml(r.port || '—') + '</td>'
        + '<td>' + (r.ok ? 'تم' : (r.cancelled ? 'تم الإيقاف' : 'فشل')) + '</td>' // (Round 61) C-60-M1: stopped rows read "تم الإيقاف", not "فشل".
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

$('btnRecheck').addEventListener('click', function () { checkBridge(true); });
$('btnRefreshPorts').addEventListener('click', loadPorts);
$('btnDetect').addEventListener('click', detectModems);
$('btnTest').addEventListener('click', testSelected);
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
$('fileInput').addEventListener('change', function () {
  previewNumbers = [];
  $('previewCounts').innerHTML = '';
  $('previewTable').classList.add('hidden');
  $('btnPreview').disabled = !selectedFile();
  updateSendButton();
});

checkBridge(true);
