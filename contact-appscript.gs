/**
 * ZIENSYS — Contact form handler (Google Apps Script Web App)
 * -----------------------------------------------------------------
 * What it does when the website contact form is submitted:
 *   1. Appends the submission as a row in the bound Google Sheet.
 *   2. Emails the details to Info@ziensys.com
 *      with CC to preethamreddy2808@gmail.com and pathipakasunil@gmail.com
 *      and Reply-To set to the person who filled the form.
 *
 * SETUP
 *   1. Create a Google Sheet. Extensions ▸ Apps Script.
 *   2. Paste this file in as Code.gs and Save.
 *   3. Deploy ▸ New deployment ▸ type "Web app"
 *        - Execute as: Me
 *        - Who has access: Anyone
 *      Copy the /exec URL it gives you.
 *   4. Point the website form at that URL (see FORM_WIRING.md / the snippet
 *      at the bottom of this file).
 *   5. First run will ask you to authorise Gmail + Sheets — accept.
 */

// ─── Config ──────────────────────────────────────────────────────
var TO_EMAIL   = 'Info@ziensys.com';
var CC_EMAILS  = 'preethamreddy2808@gmail.com,pathipakasunil@gmail.com';
var SHEET_NAME = 'Contact Submissions';
var EMAIL_SUBJECT_PREFIX = 'New Website Contact — ';
// ─────────────────────────────────────────────────────────────────

/**
 * Handles POST requests from the contact form.
 * Accepts either form-encoded fields or a JSON body.
 */
function doPost(e) {
  try {
    var data = parseRequest_(e);

    var name    = String(data.name    || '').trim();
    var email   = String(data.email   || '').trim();
    var message = String(data.message || '').trim();
    var phone   = String(data.phone   || '').trim();      // optional
    var company = String(data.company || '').trim();      // optional

    if (!name || !email || !message) {
      return json_({ ok: false, error: 'Missing required fields (name, email, message).' });
    }

    var timestamp = new Date();

    saveToSheet_({
      timestamp: timestamp,
      name: name,
      email: email,
      phone: phone,
      company: company,
      message: message
    });

    sendNotificationEmail_({
      timestamp: timestamp,
      name: name,
      email: email,
      phone: phone,
      company: company,
      message: message
    });

    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

/** Simple GET so you can open the URL and confirm the app is live. */
function doGet() {
  return json_({ ok: true, service: 'ZIENSYS contact handler' });
}

// ─── Helpers ─────────────────────────────────────────────────────

function parseRequest_(e) {
  if (!e) return {};
  // JSON body
  if (e.postData && e.postData.type === 'application/json') {
    try { return JSON.parse(e.postData.contents) || {}; } catch (ignore) {}
  }
  // form-encoded / multipart
  if (e.parameter && Object.keys(e.parameter).length) {
    return e.parameter;
  }
  // raw form-encoded fallback
  if (e.postData && e.postData.contents) {
    var out = {};
    e.postData.contents.split('&').forEach(function (pair) {
      var kv = pair.split('=');
      if (kv.length === 2) {
        out[decodeURIComponent(kv[0])] = decodeURIComponent(kv[1].replace(/\+/g, ' '));
      }
    });
    return out;
  }
  return {};
}

function saveToSheet_(row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Timestamp', 'Name', 'Email', 'Phone', 'Company', 'Message']);
    sheet.setFrozenRows(1);
  }
  sheet.appendRow([
    row.timestamp,
    row.name,
    row.email,
    row.phone,
    row.company,
    row.message
  ]);
}

function sendNotificationEmail_(row) {
  var tz = Session.getScriptTimeZone();
  var when = Utilities.formatDate(row.timestamp, tz, 'yyyy-MM-dd HH:mm');

  var lines = [
    'A new contact form submission was received on the ZIENSYS website.',
    '',
    'Name:     ' + row.name,
    'Email:    ' + row.email,
    'Phone:    ' + (row.phone || '—'),
    'Company:  ' + (row.company || '—'),
    'Received: ' + when,
    '',
    'Message:',
    row.message
  ];

  var htmlBody =
    '<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#0f172a;line-height:1.6">' +
      '<p>A new contact form submission was received on the <strong>ZIENSYS</strong> website.</p>' +
      '<table cellpadding="6" style="border-collapse:collapse">' +
        tr_('Name', row.name) +
        tr_('Email', '<a href="mailto:' + escapeHtml_(row.email) + '">' + escapeHtml_(row.email) + '</a>') +
        tr_('Phone', row.phone || '—') +
        tr_('Company', row.company || '—') +
        tr_('Received', when) +
      '</table>' +
      '<p style="margin-top:16px"><strong>Message</strong><br>' +
        escapeHtml_(row.message).replace(/\n/g, '<br>') +
      '</p>' +
    '</div>';

  MailApp.sendEmail({
    to: TO_EMAIL,
    cc: CC_EMAILS,
    replyTo: row.email,
    subject: EMAIL_SUBJECT_PREFIX + row.name,
    body: lines.join('\n'),
    htmlBody: htmlBody,
    name: 'ZIENSYS Website'
  });
}

function tr_(label, value) {
  return '<tr>' +
    '<td style="color:#64748b;vertical-align:top"><strong>' + label + '</strong></td>' +
    '<td>' + value + '</td>' +
  '</tr>';
}

function escapeHtml_(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ───────────────────────────────────────────────────────────────────
   FRONT-END WIRING (put this in index.html, replacing the current
   FormSubmit action on #ziensys-contact-form)

   <script>
     const CONTACT_ENDPOINT = 'https://script.google.com/macros/s/AKfycbx2YrbNBusUkc8rtpz-xJtq7Foch0F5OcaGkGlfclXpMvLoPFSzIwDHSZb8Cbbb_j5Upw/exec';
     const form = document.getElementById('ziensys-contact-form');
     form.addEventListener('submit', async (ev) => {
       ev.preventDefault();
       const btn = form.querySelector('.contact-glass-submit');
       const original = btn.textContent;
       btn.disabled = true; btn.textContent = 'Sending…';
       try {
         await fetch(CONTACT_ENDPOINT, {
           method: 'POST',
           body: new FormData(form)          // name / email / message fields
         });
         btn.textContent = 'Message sent ✓';
         form.reset();
       } catch (e) {
         btn.textContent = 'Failed — try again';
         btn.disabled = false;
       }
       setTimeout(() => { btn.textContent = original; btn.disabled = false; }, 4000);
     });
   </script>
   ─────────────────────────────────────────────────────────────────── */
