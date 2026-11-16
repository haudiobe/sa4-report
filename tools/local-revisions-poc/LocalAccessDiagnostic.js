/**
 * DISPOSABLE PROOF OF CONCEPT -- not part of any release bundle.
 *
 * Can the browser context of an Apps Script dialog read the meeting-room
 * server's folder listing (http://10.10.10.10/...)? Push this file next to a
 * report's files into a DISPOSABLE copy only, open the copy on the meeting-room
 * network, and run localAccessDiagnostic() from the Apps Script editor. The
 * dialog runs every probe in the browser and shows the result as text to copy.
 * Nothing is stored, nothing is sent anywhere but to the meeting-room server.
 *
 * Probes (all in the dialog's own iframe, an https page of googleusercontent.com):
 *   context    origin, isSecureContext, iframe depth, user agent
 *   permission navigator.permissions 'local-network-access' (Chrome LNA), if known
 *   cors       fetch(url)                         -- would give the listing text
 *   cors+lna   fetch(url, { targetAddressSpace }) -- Chrome's Local Network Access opt-in
 *   no-cors    fetch(url, { mode: 'no-cors' })    -- opaque: reachable, but unreadable
 *   image      <img src="http://10.10.10.10/...">  -- passive mixed content (often upgraded/blocked)
 * Only "cors" or "cors+lna" returning the listing text would make direct reading feasible.
 */

var LOCAL_ACCESS_DIAGNOSTIC_DEFAULT_URL_ = 'http://10.10.10.10/ftp/SA/SA4/Inbox/Drafts/MBS/';

function localAccessDiagnostic() {
  const name = DocumentApp.getActiveDocument().getName();
  if (name.indexOf('DISPOSABLE TEST COPY') !== 0) throw new Error('Run this only in a disposable copy (its name starts with "DISPOSABLE TEST COPY").');
  const html = HtmlService.createHtmlOutput(localAccessDiagnosticHtml_(LOCAL_ACCESS_DIAGNOSTIC_DEFAULT_URL_)).setWidth(640).setHeight(520);
  DocumentApp.getUi().showModalDialog(html, 'Local access diagnostic (PoC)');
}

/** Pure: the dialog page. */
function localAccessDiagnosticHtml_(url) {
  const esc = function (v) { return String(v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;'); };
  return '<!DOCTYPE html><html><head><base target="_top"><style>body{font:13px Arial,sans-serif;margin:10px}pre{background:#f4f4f4;padding:6px;white-space:pre-wrap;max-height:330px;overflow:auto}input{width:100%}</style></head><body>' +
    '<div>Folder to probe:</div><input id="u" value="' + esc(url) + '"> <button id="go">Run probes</button>' +
    '<pre id="out">Not run yet.</pre><div>Copy everything above and send it back. Nothing is stored.</div>' +
    '<script>' + localAccessDiagnosticClient_.toString() + '\nlocalAccessDiagnosticClient_(document, window);</script></body></html>';
}

/** The probes, run in the browser (shipped as its own source; no server call). */
function localAccessDiagnosticClient_(doc, win) {
  var out = doc.getElementById('out');
  var lines = [];
  function log(k, v) { lines.push(k + ': ' + v); out.textContent = lines.join('\n'); }
  function depth() { var d = 0, w = win; try { while (w !== w.parent && d < 10) { w = w.parent; d++; } } catch (e) { d = d + '+ (cross-origin)'; } return d; }
  function withTimeout(p, ms) { return Promise.race([p, new Promise(function (_, rej) { setTimeout(function () { rej(new Error('timeout ' + ms + ' ms')); }, ms); })]); }
  function summarize(text) { var links = (String(text).match(/<a\s/gi) || []).length; return 'read ' + String(text).length + ' chars, ' + links + ' links'; }
  doc.getElementById('go').onclick = function () {
    lines = [];
    var url = doc.getElementById('u').value.trim();
    log('time', new Date().toISOString());
    log('url', url);
    log('origin', win.location.origin);
    log('isSecureContext', win.isSecureContext);
    log('iframe depth', depth());
    log('userAgent', win.navigator.userAgent);
    var steps = [];
    steps.push(function () {
      if (!win.navigator.permissions || !win.navigator.permissions.query) { log('permission local-network-access', 'Permissions API not available'); return Promise.resolve(); }
      return win.navigator.permissions.query({ name: 'local-network-access' })
        .then(function (s) { log('permission local-network-access', s.state); })
        .catch(function (e) { log('permission local-network-access', 'not known to this browser (' + e.message + ')'); });
    });
    steps.push(function () {
      return withTimeout(win.fetch(url, { cache: 'no-store' }), 8000)
        .then(function (r) { return r.text().then(function (t) { log('cors', 'HTTP ' + r.status + ', type ' + r.type + ', ' + summarize(t) + '  <-- READABLE'); }); })
        .catch(function (e) { log('cors', 'blocked: ' + e.name + ': ' + e.message); });
    });
    steps.push(function () {
      return withTimeout(win.fetch(url, { cache: 'no-store', targetAddressSpace: 'local' }), 8000)
        .then(function (r) { return r.text().then(function (t) { log('cors+lna', 'HTTP ' + r.status + ', type ' + r.type + ', ' + summarize(t) + '  <-- READABLE'); }); })
        .catch(function (e) { log('cors+lna', 'blocked: ' + e.name + ': ' + e.message); });
    });
    steps.push(function () {
      return withTimeout(win.fetch(url, { mode: 'no-cors', cache: 'no-store' }), 8000)
        .then(function (r) { log('no-cors', 'type ' + r.type + ' (opaque = the request went out; its content cannot be read)'); })
        .catch(function (e) { log('no-cors', 'blocked: ' + e.name + ': ' + e.message); });
    });
    steps.push(function () {
      return new Promise(function (resolve) {
        var img = new Image();
        var done = function (r) { log('image', r); resolve(); };
        img.onload = function () { done('loaded (passive mixed content allowed; still says nothing about reading the listing)'); };
        img.onerror = function () { done('error (blocked, upgraded to https, or no such image)'); };
        setTimeout(function () { done('timeout'); }, 8000);
        img.src = url.replace(/^(https?:\/\/[^\/]+).*$/, '$1') + '/favicon.ico?probe=' + Date.now();
      });
    });
    steps.reduce(function (p, step) { return p.then(step); }, Promise.resolve()).then(function () {
      var readable = lines.some(function (l) { return /READABLE/.test(l); });
      log('VERDICT', readable ? 'the dialog CAN read the listing (direct import feasible)' : 'the dialog CANNOT read the listing (direct import not feasible here)');
    });
  };
}
