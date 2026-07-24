/* PaperTrail for Word — taskpane logic (v1.0)
 *
 * Everything but Inspect, ported from the web app (app.html) and the extension:
 *   • Read the open Word document (Office.js)
 *   • StyleMatch — full 8-metric client-side engine (stylematch.js), no transmit
 *   • Verify — Consistency + Data & Methods + Compare + Citations, calling the
 *     SAME edge functions as the web app (generate-report / verify-data-integrity /
 *     verify-citations), polled, then rendered IN-PANE with the SAME renderer
 *     (reports.js → window.__wpaReports.*) inside an isolated <iframe>.
 *   • Token wallet + buy-tokens (Lemon Squeezy, checkout[custom][user_id])
 *   • Oral Defense — reads the essay and hands off to the web wizard (least friction)
 *
 * Contracts verified against papertrail-oral/app.html (2026-07). Token costs:
 * Consistency/Data/Compare = 1; Citations = 2+ (server-priced, over-cap chooser).
 * report_json is nulled immediately after a successful fetch (privacy posture).
 */
'use strict';

var SUPABASE_URL = 'https://ktzrdhiqhidexunucuqp.supabase.co';
var ANON_KEY =
  'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Imt0enJkaGlxaGlkZXh1bnVjdXFwIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzI2MTQ2ODMsImV4cCI6MjA4ODE5MDY4M30.24Gbyh0C-wbDqQ31MD-ttXoCfLwjcWvjs9-UXAaqmHU';
var ADDIN_BASE = 'https://papertrailacademic.com/word-addin/';
var WEB_APP_URL = 'https://app.papertrailacademic.com/app';
var CHECKOUT_URL = 'https://papertrailacademic.lemonsqueezy.com/checkout/buy/bb899eea-e9af-47ac-a757-052c5f189e8b'; // Verify tokens
var SM_SUB_CHECKOUT = 'https://papertrailacademic.lemonsqueezy.com/checkout/buy/33e2150f-8332-42d5-ac7d-d22a8bc6b646'; // StyleMatch subscription ($1.99/mo · $19.99/yr)

var MIN_WORDS = 500;   // Consistency + Compare + StyleMatch
var MIN_WORDS_SHORT = 300; // Data & Methods + Citations
var SM_FREE_RUN_LIMIT = 3; // lifetime free StyleMatch runs for registered free users (mirrors extension)

var docText = '';
var session = null;    // { access_token, email, user_id }
var vt = 'consistency';
var verifyRunning = false;

// StyleMatch's engine (stylematch.js) exposes its API only under the
// window.__wpaStyleMatch namespace. This file was ported from the extension,
// where smWords/smComputeScores were globals. Shim them so the ported call
// sites work unchanged. Lookup is deferred to call time (load-order safe).
function smWords(t) { return window.__wpaStyleMatch.smWords(t); }
function smComputeScores(a, b) { return window.__wpaStyleMatch.smComputeScores(a, b); }

var $ = function (id) { return document.getElementById(id); };
function wc(t) { return t && t.trim() ? smWords(t.trim()).length : 0; }
function esc(s) { return String(s == null ? '' : s).replace(/[&<>]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]; }); }

Office.onReady(function (info) {
  if (info.host !== Office.HostType.Word) {
    $('docinfo').textContent = 'This add-in runs inside Microsoft Word.';
    return;
  }
  $('btn-read').onclick = readDocument;
  $('acct').onclick = onAccountClick;
  $('btn-sm').onclick = runStyleMatch;
  $('btn-verify').onclick = runVerifyDispatch;
  $('btn-oral').onclick = createOral;

  // tabs
  var tabs = document.querySelectorAll('.tabs button');
  tabs.forEach(function (b) {
    b.onclick = function () {
      tabs.forEach(function (x) { x.classList.remove('active'); });
      b.classList.add('active');
      document.querySelectorAll('.tab').forEach(function (p) { p.classList.remove('active'); });
      $('tab-' + b.getAttribute('data-tab')).classList.add('active');
    };
  });

  // verify type segments
  document.querySelectorAll('#verify-type button').forEach(function (b) {
    b.onclick = function () {
      document.querySelectorAll('#verify-type button').forEach(function (x) { x.classList.remove('active'); });
      b.classList.add('active');
      vt = b.getAttribute('data-vt');
      applyVerifyType();
    };
  });

  // word counts
  $('sampleA').addEventListener('input', function () {
    var n = wc($('sampleA').value);
    var el = $('sm-wc'); el.textContent = n + ' words';
    el.className = 'wc ' + (n === 0 ? '' : n >= MIN_WORDS ? 'ok' : 'bad');
  });
  $('ctrlText').addEventListener('input', function () {
    var n = wc($('ctrlText').value);
    var el = $('ctrl-wc'); el.textContent = n + ' words';
    el.className = 'wc ' + (n === 0 ? '' : n >= MIN_WORDS ? 'ok' : 'bad');
  });

  // report overlay
  $('report-back').onclick = function () { $('report-overlay').classList.add('hidden'); $('report-frame').srcdoc = ''; };
  $('report-print').onclick = function () { try { $('report-frame').contentWindow.print(); } catch (e) {} };

  applyVerifyType();
});

/* ── Read the document ─────────────────────────────────────────── */
function readDocument() {
  $('docinfo').textContent = 'Reading…';
  Word.run(function (ctx) {
    var body = ctx.document.body;
    body.load('text');
    return ctx.sync().then(function () {
      docText = body.text || '';
      var n = wc(docText);
      $('docinfo').innerHTML = '<b>' + n.toLocaleString() + ' words</b> read ' +
        (n < MIN_WORDS_SHORT ? '<span class="pill warn">very short</span>' : '<span class="pill ok">ready</span>');
      $('btn-sm').disabled = false;
      $('btn-verify').disabled = false; // clickable even when signed out → preflight routes to sign-in / token purchase
      $('btn-oral').disabled = false;
    });
  }).catch(function (e) { $('docinfo').textContent = 'Could not read the document: ' + (e.message || e); });
}

/* ── StyleMatch (client-side) ──────────────────────────────────── */
var SM_METRICS = [
  ['grammaticalSim', 'Function words'],
  ['discourseSim', 'Discourse markers'],
  ['discoursePositionSim', 'Marker placement'],
  ['punctuationSim', 'Punctuation'],
  ['sentenceLengthSim', 'Sentence length'],
  ['sentLengthHistSim', 'Sentence rhythm'],
  ['ttrSim', 'Type-token ratio'],
  ['fkSim', 'Reading grade'],
  ['informalitySim', 'Informality']
];
function smBand(v) {
  if (v >= 70) return ['Similar', 'var(--teal)'];
  if (v >= 45) return ['Notable', 'var(--warn)'];
  return ['Divergent', 'var(--red)'];
}
function smShow(html) { $('sm-results').classList.remove('hidden'); $('sm-results').innerHTML = html; }

function runStyleMatch() {
  var a = $('sampleA').value.trim();
  var wa = wc(a), wb = wc(docText);
  if (!docText) { smShow('<div class="muted">Read the document first.</div>'); return; }
  if (wa < MIN_WORDS || wb < MIN_WORDS) {
    smShow('<div class="muted">Both samples need at least 500 words (controlled: ' + wa + ', document: ' + wb + ').</div>');
    return;
  }
  if (!session) {
    smShow('<div class="muted">Sign in to run StyleMatch — new accounts get <b>3 free analyses</b>. <a class="link" id="sm-signin">Sign in</a></div>');
    var sb = $('sm-signin'); if (sb) sb.onclick = signIn;
    return;
  }
  // Gate against the shared Supabase profile — single source of truth across all surfaces.
  smShow('<div class="muted">Checking your plan…</div>');
  refreshBalance().then(function () {
    var isSub = session.subscription_status === 'active' || session.subscription_status === 'cancelled';
    var runsUsed = session.stylematch_runs_used || 0;
    if (!isSub && runsUsed >= SM_FREE_RUN_LIMIT) { smShow(''); showSubModal(); return; }
    smRenderResults(a);
    var hasCredits = (session.credit_balance || 0) > 0;
    if (!isSub && !hasCredits) incrementStyleMatchRun(); // credits are for Verify; only pure-free runs count
  });
}

function smRenderResults(a) {
  var s = smComputeScores(a, docText);
  var rows = SM_METRICS.map(function (m) {
    var v = Math.round(s[m[0]]);
    if (isNaN(v)) return '';
    var band = smBand(v);
    return '<div class="mrow"><span class="n">' + m[1] + '</span>' +
      '<span class="bar"><span style="width:' + v + '%;background:' + band[1] + '"></span></span>' +
      '<span class="v" style="color:' + band[1] + '">' + v + '% ' + band[0] + '</span></div>';
  }).join('');
  var footer = '<div class="muted">Divergence ratings, not a verdict — you read the data. Genre-constrained writing (lab reports, TOK) can inflate similarity.</div>';
  var isSub = session && (session.subscription_status === 'active' || session.subscription_status === 'cancelled');
  if (session && !isSub) {
    if ((session.credit_balance || 0) > 0) {
      footer += '<div class="muted"><a class="link" id="sm-sub-link">Subscribe to StyleMatch</a> for unlimited analyses.</div>';
    } else {
      var remaining = Math.max(0, SM_FREE_RUN_LIMIT - (session.stylematch_runs_used || 0) - 1);
      footer += '<div class="muted">' + remaining + ' of ' + SM_FREE_RUN_LIMIT + ' free analyses remaining. <a class="link" id="sm-sub-link">Subscribe</a> for unlimited.</div>';
    }
  }
  smShow(rows + footer);
  var sl = $('sm-sub-link'); if (sl) sl.onclick = showSubModal;
}

function incrementStyleMatchRun() {
  if (!session || !session.user_id) return;
  fetch(SUPABASE_URL + '/rest/v1/rpc/increment_stylematch_run',
    { method: 'POST', headers: { 'Content-Type': 'application/json', apikey: ANON_KEY, Authorization: 'Bearer ' + session.access_token }, body: JSON.stringify({ user_id: session.user_id }) })
    .then(function () { session.stylematch_runs_used = (session.stylematch_runs_used || 0) + 1; })
    .catch(function () {}); // silent — may under-count; never surfaced (mirrors extension)
}

function showSubModal() {
  var existing = $('sm-sub-modal');
  if (existing) { existing.classList.remove('hidden'); return; }
  var m = document.createElement('div');
  m.id = 'sm-sub-modal';
  m.style.cssText = 'position:fixed;inset:0;background:rgba(26,34,53,.5);display:flex;align-items:center;justify-content:center;z-index:60;padding:16px';
  m.innerHTML =
    '<div style="background:#fff;border-radius:12px;max-width:340px;width:100%;padding:20px;box-shadow:0 16px 56px rgba(74,111,165,.25)">' +
      '<div style="font-family:Lora,serif;font-weight:700;font-size:16px;color:#1a2235;margin-bottom:6px">Subscribe to StyleMatch</div>' +
      '<p style="font-size:12.5px;color:#3d4f6e;line-height:1.6;margin:0 0 14px">You’ve used your 3 free StyleMatch analyses. Subscribe for unlimited authorship-consistency checks across Word, Google Docs, and the web app.</p>' +
      '<div style="display:flex;gap:8px;margin-bottom:12px">' +
        '<div style="flex:1;border:1px solid #dce4ef;border-radius:8px;padding:10px;text-align:center"><div style="font-weight:700;color:#1a2235">$1.99<span style="font-size:11px;color:#6b7a96">/mo</span></div><div style="font-size:11px;color:#6b7a96">Monthly</div></div>' +
        '<div style="flex:1;border:1px solid #C9A84C;border-radius:8px;padding:10px;text-align:center;background:#fdf6e3"><div style="font-weight:700;color:#1a2235">$19.99<span style="font-size:11px;color:#6b7a96">/yr</span></div><div style="font-size:11px;color:#a8892f">Save ~16%</div></div>' +
      '</div>' +
      '<button class="btn" id="sm-sub-go" style="margin-bottom:8px">Subscribe</button>' +
      '<button class="btn ghost" id="sm-sub-cancel">Not now</button>' +
    '</div>';
  document.body.appendChild(m);
  $('sm-sub-go').onclick = function () { openCheckout(SM_SUB_CHECKOUT); };
  $('sm-sub-cancel').onclick = function () { m.classList.add('hidden'); };
}

/* ── Account / sign-in ─────────────────────────────────────────── */
function onAccountClick() {
  if (session) { openCheckout(); return; } // signed in → chip buys tokens
  signIn();
}
function signIn() {
  Office.context.ui.displayDialogAsync(ADDIN_BASE + 'auth-dialog.html',
    { height: 55, width: 30, displayInIframe: false },
    function (res) {
      if (res.status !== Office.AsyncResultStatus.Succeeded) return;
      var dlg = res.value;
      dlg.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
        try {
          var msg = JSON.parse(arg.message);
          if (msg.type === 'auth') {
            session = msg;
            $('btn-verify').disabled = !docText;
            refreshBalance();
          }
        } catch (e) {}
        dlg.close();
      });
    });
}
function refreshBalance() {
  if (!session) return Promise.resolve();
  return fetch(SUPABASE_URL + '/rest/v1/users?id=eq.' + encodeURIComponent(session.user_id) + '&select=subscription_status,credit_balance,stylematch_runs_used',
    { headers: { apikey: ANON_KEY, Authorization: 'Bearer ' + session.access_token } })
    .then(function (r) { return r.json(); })
    .then(function (rows) {
      var row = (rows && rows[0]) || {};
      session.subscription_status  = row.subscription_status || 'free';
      session.credit_balance       = typeof row.credit_balance === 'number' ? row.credit_balance : 0;
      session.stylematch_runs_used = row.stylematch_runs_used || 0;
      var chip = $('acct');
      chip.className = 'chip tokens';
      chip.textContent = session.email + ' · ' + session.credit_balance + ' 🪙';
    }).catch(function () {});
}
function openCheckout(base) {
  base = base || CHECKOUT_URL; // default: Verify tokens
  var url = session ? base + '?checkout[custom][user_id]=' + encodeURIComponent(session.user_id) + '&checkout[email]=' + encodeURIComponent(session.email) : base;
  openExternal(url);
}
function openExternal(url) {
  try {
    if (Office.context.ui.openBrowserWindow) { Office.context.ui.openBrowserWindow(url); return; }
  } catch (e) {}
  window.open(url, '_blank');
}

/* ── Verify — shared ───────────────────────────────────────────── */
var VERIFY_INTRO = {
  consistency: '<b>Consistency Report.</b> The document is read in ~150-word chunks and measured on six dimensions per chunk — register, vocabulary, sentence complexity, argument depth, error density, cohesion — with a heatmap of where the writing shifts. Text is sent for analysis and not retained. <b>1 token.</b>',
  data: '<b>Data &amp; Methods Report.</b> For lab reports, IAs, data-based essays. Every reported statistic — t-tests, p-values, means, SDs, correlations, percentage changes — is <b>recomputed in code</b> from the numbers in the document. Also checks protocol arithmetic and "see Table N" references. Text is sent for analysis and not retained. <b>1 token.</b>',
  compare: '<b>Compare Report.</b> The submitted document is analysed against a controlled writing sample, with your on-device StyleMatch scores as anchors. Text is sent for analysis and not retained. <b>1 token.</b>',
  citations: '<b>Verify Citations.</b> Every citation is identified, then each source is checked against what the essay says it says — quotes verified where reachable via live web search. Keep the works-cited list in the document. Text is sent for analysis and not retained. <b>2+ tokens (you confirm the count first).</b>'
};
function applyVerifyType() {
  $('verify-intro').innerHTML = VERIFY_INTRO[vt];
  $('compare-wrap').classList.toggle('hidden', vt !== 'compare');
  var labels = { consistency: '🔬 Run Consistency — 1 token', data: '🧪 Run Data & Methods — 1 token', compare: '⚖️ Run Compare — 1 token', citations: '📚 Run Citations' };
  $('btn-verify').textContent = labels[vt];
}
function getMeta() {
  return { studentName: $('m-student').value.trim(), assignment: $('m-assignment').value.trim() || 'Word add-in', course: $('m-course').value.trim() };
}
function vStatus(msg, isErr) { var el = $('verify-status'); el.className = 'status' + (isErr ? ' err' : ''); el.innerHTML = msg; }

function preflight(minWords) {
  if (verifyRunning) return false;
  if (!session) { vStatus('Sign in to run Verify.', true); signIn(); return false; }
  if (!docText) { vStatus('Read the document first.', true); return false; }
  if ((session.credit_balance || 0) <= 0) { vStatus('You have no PaperTrail Tokens — opening checkout. Each report costs 1 token (Citations 2+).', true); openCheckout(CHECKOUT_URL); return false; }
  if (wc(docText) < minWords) { vStatus('This report needs at least ' + minWords + ' words in the document.', true); return false; }
  return true;
}
function runVerifyDispatch() {
  // Refresh the shared profile first so the token check (and just-completed
  // purchases / subscriptions) reflect the current Supabase state.
  if (session) { refreshBalance().then(_verifyDispatch); } else { _verifyDispatch(); }
}
function _verifyDispatch() {
  if (vt === 'consistency') return runReport('consistency');
  if (vt === 'data') return runDataIntegrity();
  if (vt === 'compare') return runReport('compare');
  if (vt === 'citations') return runCitations();
}

/* Render a full report in a large Office dialog (task panes can't be widened by
   the add-in). Falls back to the in-pane overlay if the dialog can't open. */
var _reportDlg = null;
function showReport(html) {
  if (_reportDlg) { try { _reportDlg.close(); } catch (e) {} _reportDlg = null; }
  var canDialog = Office.context && Office.context.ui && Office.context.ui.displayDialogAsync;
  if (!canDialog) { $('report-frame').srcdoc = html; $('report-overlay').classList.remove('hidden'); return; }
  Office.context.ui.displayDialogAsync(ADDIN_BASE + 'report-viewer.html',
    { height: 86, width: 62, displayInIframe: true },
    function (res) {
      if (res.status !== Office.AsyncResultStatus.Succeeded) {
        $('report-frame').srcdoc = html; $('report-overlay').classList.remove('hidden'); return;
      }
      var dlg = res.value; _reportDlg = dlg;
      dlg.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
        if (arg.message === 'ready') { sendReportToDialog(dlg, html); }
        else if (arg.message === 'close') { try { dlg.close(); } catch (e) {} _reportDlg = null; }
      });
      dlg.addEventHandler(Office.EventType.DialogEventReceived, function () { _reportDlg = null; });
    });
}

/* Office dialog messages are size-capped, so stream the report HTML in chunks
   the viewer reassembles in order. */
function sendReportToDialog(dlg, html) {
  if (!dlg.messageChild) return;
  var CHUNK = 7000;
  var total = Math.ceil(html.length / CHUNK) || 1;
  dlg.messageChild(JSON.stringify({ t: 'begin', total: total }));
  for (var i = 0; i < total; i++) {
    dlg.messageChild(JSON.stringify({ t: 'chunk', i: i, s: html.substr(i * CHUNK, CHUNK) }));
  }
  dlg.messageChild(JSON.stringify({ t: 'end' }));
}

/* Poll the reports row until complete, then hand report_json to `build` */
function pollReport(reportId, opts, build) {
  var interval = opts.interval || 2500, max = opts.max || 200, attempts = 0, finished = false, timer = null;
  function stop() { if (timer) clearInterval(timer); }
  function tick() {
    if (finished) return;
    attempts++;
    fetch(SUPABASE_URL + '/rest/v1/reports?id=eq.' + encodeURIComponent(reportId) + '&select=status,report_json,error_message',
      { headers: { apikey: ANON_KEY, Authorization: 'Bearer ' + session.access_token } })
      .then(function (r) { return r.json(); })
      .then(function (rows) {
        var row = rows && rows[0];
        if (!row) { if (attempts >= max) { finished = true; stop(); done('Report not found. Your token was not charged.'); } return; }
        if (row.status === 'complete' && row.report_json) {
          finished = true; stop();
          var rd = row.report_json; if (!rd.reportId) rd.reportId = reportId;
          verifyRunning = false;
          vStatus('Report ready.');
          showReport(build(rd));
          refreshBalance();
          // privacy: null report_json after fetch
          fetch(SUPABASE_URL + '/rest/v1/reports?id=eq.' + encodeURIComponent(reportId),
            { method: 'PATCH', headers: { apikey: ANON_KEY, Authorization: 'Bearer ' + session.access_token, 'Content-Type': 'application/json', Prefer: 'return=minimal' }, body: JSON.stringify({ report_json: null }) }).catch(function () {});
        } else if (row.status === 'failed') {
          finished = true; stop();
          done(row.error_message === 'no_quantitative_content'
            ? 'No statistics or tables were found to recompute — for prose essays use Consistency instead. No tokens were used.'
            : (row.error_message || 'Generation failed. Your token was not charged.'));
        } else if (attempts >= max) {
          finished = true; stop();
          done('This is taking unusually long. If a report appears on your web dashboard shortly, view it there; otherwise try again.');
        } else {
          vStatus('Generating… (typically 1–3 minutes) <span class="pill info">' + attempts + '</span>');
        }
      })
      .catch(function () { if (attempts >= max) { finished = true; stop(); done('Lost connection while waiting for the report.'); } });
  }
  function done(msg) { verifyRunning = false; $('btn-verify').disabled = false; vStatus(msg, true); }
  tick(); timer = setInterval(tick, interval);
}

/* Consistency (single) + Compare (comparative) → generate-report */
function runReport(kind) {
  var comparative = kind === 'compare';
  if (!preflight(MIN_WORDS)) return;
  var meta = getMeta();
  var body = { sampleB: docText, analysisMode: comparative ? 'comparative' : 'single', metadata: meta };
  if (comparative) {
    var ctrl = $('ctrlText').value.trim();
    if (wc(ctrl) < MIN_WORDS) { vStatus('Compare needs a controlled sample of at least 500 words.', true); return; }
    body.sampleA = ctrl;
    body.algorithmicScores = smComputeScores(ctrl, docText);
  }
  verifyRunning = true; $('btn-verify').disabled = true; vStatus('Submitting…');
  fetch(SUPABASE_URL + '/functions/v1/generate-report',
    { method: 'POST', headers: { 'Content-Type': 'application/json', apikey: ANON_KEY, Authorization: 'Bearer ' + session.access_token }, body: JSON.stringify(body) })
    .then(function (res) { return res.json().then(function (data) { return { status: res.status, data: data }; }); })
    .then(function (r) {
      if (handleCommonErrors(r)) return;
      if (r.status === 202 && r.data.reportId) {
        pollReport(r.data.reportId, { interval: 2000, max: 200 }, function (rd) { return window.__wpaReports.smBuildVerifyReportHTML(rd, comparative, meta, ''); });
      } else if (r.status === 200 && (r.data.report || r.data.reportId)) {
        var rd = r.data.report || r.data; if (!rd.reportId) rd.reportId = r.data.reportId;
        verifyRunning = false; $('btn-verify').disabled = false; vStatus('Report ready.');
        showReport(window.__wpaReports.smBuildVerifyReportHTML(rd, comparative, meta, '')); refreshBalance();
      } else { failVerify('The Verify service returned an unexpected response. Your token was not charged.'); }
    })
    .catch(function () { failVerify('Could not reach the Verify service — your token was not charged.'); });
}

/* Data & Methods → verify-data-integrity */
function runDataIntegrity() {
  if (!preflight(MIN_WORDS_SHORT)) return;
  var meta = getMeta();
  verifyRunning = true; $('btn-verify').disabled = true; vStatus('Submitting…');
  fetch(SUPABASE_URL + '/functions/v1/verify-data-integrity',
    { method: 'POST', headers: { 'Content-Type': 'application/json', apikey: ANON_KEY, Authorization: 'Bearer ' + session.access_token }, body: JSON.stringify({ essay: docText, access_token: session.access_token, metadata: meta }) })
    .then(function (res) { return res.json().then(function (data) { return { status: res.status, data: data }; }); })
    .then(function (r) {
      if (handleCommonErrors(r)) return;
      if (r.data && r.data.ok === false) { return failVerify(r.data.message || 'Nothing to check. No tokens were used.'); }
      if (r.status === 202 && r.data.reportId) {
        pollReport(r.data.reportId, { interval: 2500, max: 120 }, function (rd) { return window.__wpaReports.smBuildDataIntegrityReportHTML(rd, meta, ''); });
      } else { failVerify('The Verify service returned an unexpected response. No tokens were used.'); }
    })
    .catch(function () { failVerify('Could not reach the Verify service — no tokens were used.'); });
}

/* Citations → verify-citations (parse → over-cap chooser → run) */
function runCitations() {
  if (!preflight(MIN_WORDS_SHORT)) return;
  var meta = getMeta();
  verifyRunning = true; $('btn-verify').disabled = true; vStatus('Finding citations…');
  citPost({ action: 'parse', essay: docText, metadata: meta }).then(function (r) {
    if (handleCommonErrors(r)) return;
    if (r.data && r.data.ok === false) { return failVerify(r.data.message || 'No citations were found to check. No tokens were used.'); }
    if (r.data && r.data.ok && r.data.reportId) {
      if (r.data.overCap) {
        vStatus('<b>' + r.data.citationCount + ' citations found.</b> You are not charged until you choose.<br>' +
          '<button class="link" id="cit-capped">Check first 25 — ' + r.data.costIfCapped + ' tokens</button> &nbsp;·&nbsp; ' +
          '<button class="link" id="cit-all">Check all ' + r.data.citationCount + ' — ' + r.data.costIfAll + ' tokens</button>');
        $('cit-capped').onclick = function () { citRun(r.data.reportId, 'capped', meta); };
        $('cit-all').onclick = function () { citRun(r.data.reportId, 'all', meta); };
      } else { citRun(r.data.reportId, 'all', meta); }
    } else { failVerify('The Verify service returned an unexpected response. No tokens were used.'); }
  }).catch(function () { failVerify('Could not reach the Verify service — no tokens were used.'); });
}
function citPost(payload) {
  return fetch(SUPABASE_URL + '/functions/v1/verify-citations',
    { method: 'POST', headers: { 'Content-Type': 'application/json', apikey: ANON_KEY, Authorization: 'Bearer ' + session.access_token }, body: JSON.stringify(Object.assign({ access_token: session.access_token }, payload)) })
    .then(function (res) { return res.json().then(function (data) { return { status: res.status, data: data }; }); });
}
function citRun(reportId, cap, meta) {
  vStatus('Checking citations… (this can take two to four minutes)');
  citPost({ action: 'run', reportId: reportId, cap: cap, essay: docText, metadata: meta }).then(function (r) {
    if (handleCommonErrors(r)) return;
    if (r.status === 202 && r.data.reportId) {
      pollReport(r.data.reportId, { interval: 2500, max: 200 }, function (rd) { return window.__wpaReports.smBuildCitationsReportHTML(rd, meta, ''); });
    } else { failVerify('The Verify service returned an unexpected response. Your tokens were not charged.'); }
  }).catch(function () { failVerify('Could not reach the Verify service — your tokens were not charged.'); });
}

function handleCommonErrors(r) {
  if (r.status === 401) { failVerify('Your session expired. Sign in again and retry.'); return true; }
  if (r.status === 402 || (r.data && r.data.reason === 'insufficient_credits')) {
    failVerify((r.data && r.data.message) || 'Not enough PaperTrail Tokens.'); openCheckout(); return true;
  }
  if (r.data && r.data.error) { failVerify(r.data.error); return true; }
  return false;
}
function failVerify(msg) { verifyRunning = false; $('btn-verify').disabled = false; vStatus(msg, true); }

/* ── Oral Defense — hand off to the web wizard ─────────────────── */
function createOral() {
  if (!docText) { $('oral-status').textContent = 'Read the document first.'; return; }
  var status = $('oral-status');
  var go = function (copied) {
    openExternal(WEB_APP_URL + '#oral');
    status.className = 'status';
    status.innerHTML = (copied ? 'Essay copied to your clipboard. ' : '') +
      'The PaperTrail web app is opening — choose <b>🎙️ Oral Defense</b>, ' + (copied ? 'paste the essay (Ctrl/Cmd+V)' : 'paste the essay') +
      ', and follow the steps to generate questions and send the student their link.';
  };
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(docText).then(function () { go(true); }).catch(function () { go(false); });
  } else { go(false); }
}
