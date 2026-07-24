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
var CHECKOUT_URL = 'https://papertrailacademic.lemonsqueezy.com/checkout/buy/bb899eea-e9af-47ac-a757-052c5f189e8b';

var MIN_WORDS = 500;   // Consistency + Compare + StyleMatch
var MIN_WORDS_SHORT = 300; // Data & Methods + Citations

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
      $('btn-verify').disabled = !session;
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
function runStyleMatch() {
  var a = $('sampleA').value.trim();
  var wa = wc(a), wb = wc(docText);
  if (!docText) { $('sm-results').classList.remove('hidden'); $('sm-results').innerHTML = '<div class="muted">Read the document first.</div>'; return; }
  if (wa < MIN_WORDS || wb < MIN_WORDS) {
    $('sm-results').classList.remove('hidden');
    $('sm-results').innerHTML = '<div class="muted">Both samples need at least 500 words (controlled: ' + wa + ', document: ' + wb + ').</div>';
    return;
  }
  var s = smComputeScores(a, docText);
  var rows = SM_METRICS.map(function (m) {
    var v = Math.round(s[m[0]]);
    if (isNaN(v)) return '';
    var band = smBand(v);
    return '<div class="mrow"><span class="n">' + m[1] + '</span>' +
      '<span class="bar"><span style="width:' + v + '%;background:' + band[1] + '"></span></span>' +
      '<span class="v" style="color:' + band[1] + '">' + v + '% ' + band[0] + '</span></div>';
  }).join('');
  $('sm-results').innerHTML = rows +
    '<div class="muted">Divergence ratings, not a verdict — you read the data. Genre-constrained writing (lab reports, TOK) can inflate similarity.</div>';
  $('sm-results').classList.remove('hidden');
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
  if (!session) return;
  fetch(SUPABASE_URL + '/rest/v1/users?id=eq.' + encodeURIComponent(session.user_id) + '&select=credit_balance',
    { headers: { apikey: ANON_KEY, Authorization: 'Bearer ' + session.access_token } })
    .then(function (r) { return r.json(); })
    .then(function (rows) {
      var bal = rows && rows[0] ? rows[0].credit_balance : '—';
      var chip = $('acct');
      chip.className = 'chip tokens';
      chip.textContent = session.email + ' · ' + bal + ' 🪙';
    }).catch(function () {});
}
function openCheckout() {
  var url = session ? CHECKOUT_URL + '?checkout[custom][user_id]=' + encodeURIComponent(session.user_id) + '&checkout[email]=' + encodeURIComponent(session.email) : CHECKOUT_URL;
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
  if (!session) { vStatus('Sign in first (top-right).', true); return false; }
  if (!docText) { vStatus('Read the document first.', true); return false; }
  if (wc(docText) < minWords) { vStatus('This report needs at least ' + minWords + ' words in the document.', true); return false; }
  return true;
}
function runVerifyDispatch() {
  if (vt === 'consistency') return runReport('consistency');
  if (vt === 'data') return runDataIntegrity();
  if (vt === 'compare') return runReport('compare');
  if (vt === 'citations') return runCitations();
}

/* Render a full report HTML string into the isolated iframe overlay */
function showReport(html) {
  $('report-frame').srcdoc = html;
  $('report-overlay').classList.remove('hidden');
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
