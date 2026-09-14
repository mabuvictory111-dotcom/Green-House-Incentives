/**
 * Green House Incentives — Claims backend (Google Sheets + Apps Script)
 * --------------------------------------------------------------------
 * 1. Creates a "Claims" log in your Google Sheet.
 * 2. Receives applications from the website (doPost).
 * 3. Emails the applicant their tracking ID instantly.
 * 4. Adds a menu so you can send Approval / More-info / Paid emails from the sheet.
 *
 * SETUP: see SETUP-GUIDE.md in the project folder.
 */

/* ------------------------- CONFIG ------------------------- */
var CFG = {
  SHEET_NAME: 'Claims',
  FROM_NAME: 'Green House Incentives',
  REPLY_TO: 'greenhouseincentives@gmail.com', // where replies land
  FROM_ADDRESS: 'greenhouseincentives@gmail.com', // MUST be a verified "Send mail as" address, or leave blank
  OWNER_EMAIL: '',      // leave blank to notify the account running the script
  DEPOSIT_DEFAULT: 0,   // pre-filled deposit amount in the approval prompt (you can override per claim)
  /* Max rebate per appliance category — keep in step with the website. */
  REBATES: {
    'Refrigerator / freezer': 5500,
    'Washing machine': 4500,
    'Heat-pump water heater': 8500,
    'Solar water heater': 7500,
    'Induction cooktop': 3500,
    'Air conditioner (inverter)': 6500
  },
  DEPOSIT_INSTRUCTIONS: 'Reply to this email and we will send you the deposit payment details.',
  BRAND_GREEN: '#0f9d58',
  SITE_URL: 'https://greenhouseincentives.com'
};

var COL = {
  TS: 1, ID: 2, NAME: 3, EMAIL: 4, PHONE: 5, STATE: 6, APPLIANCE: 7,
  DATE: 8, BRAND: 9, PRICE: 10, STORE: 11, RECEIPT: 12, NOTES: 13,
  STATUS: 14, AMOUNT: 15, UPDATED: 16, LOG: 17,
  DEPOSIT: 18, DEPOSIT_PAID: 19, PHOTO: 20
};

var HEADERS = ['Timestamp', 'Tracking ID', 'Full name', 'Email', 'Phone', 'State', 'Appliance',
  'Purchase date', 'Brand & model', 'Price (USD)', 'Retailer', 'Receipt #', 'Notes',
  'Status', 'Rebate amount', 'Last updated', 'Email log',
  'Deposit required', 'Deposit received', 'Receipt photo'];

var STATUS = {
  RECEIVED: 'Received',
  REVIEW: 'Under review',
  NEED_INFO: 'More info needed',
  APPROVED: 'Approved',
  DECLINED: 'Declined',
  PAID: 'Paid',
  AWAITING_DEPOSIT: 'Awaiting deposit'
};

/* ------------------------- SHEET SETUP ------------------------- */

function setupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CFG.SHEET_NAME) || ss.insertSheet(CFG.SHEET_NAME);
  sh.clear();
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sh.getRange(1, 1, 1, HEADERS.length)
    .setFontWeight('bold').setBackground('#0b3d2c').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  for (var i = 1; i <= HEADERS.length; i++) sh.autoResizeColumn(i);
  sh.getRange('A1:T1').setWrap(false);
  ss.toast('Claims sheet ready.', 'Green House Incentives', 5);
  return sh;
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CFG.SHEET_NAME);
  if (!sh) sh = setupSheet();
  return sh;
}

function ownerEmail_() {
  return CFG.OWNER_EMAIL || Session.getActiveUser().getEmail();
}

/* ------------------------- WEB ENDPOINTS ------------------------- */

/** Website form posts here. Client sends JSON as text/plain (avoids CORS preflight). */
function doPost(e) {
  try {
    var raw = e && e.postData && e.postData.contents ? e.postData.contents : '';
    var d = JSON.parse(raw);

    var sh = getSheet_();
    var id = (d.id || genId_()).toString().toUpperCase();
    var now = new Date();

    sh.appendRow([
      now, id, d.name || '', d.email || '', d.phone || '', d.state || '', d.appliance || '',
      d.date || '', d.brand || '', d.price || '', d.store || '', d.receipt || '', d.notes || '',
      STATUS.RECEIVED, '', now, ''
    ]);
    var rowNum = sh.getLastRow();
    sh.getRange(rowNum, COL.ID).setFontWeight('bold');
    SpreadsheetApp.flush();

    var errs = [];

    /* Receipt photo: decoded and filed in Drive, link stored in the sheet. */
    if (d.photo) {
      try {
        sh.getRange(rowNum, COL.PHOTO).setValue(saveReceiptPhoto_(d, id));
      } catch (pe) {
        errs.push('RECEIPT PHOTO FAILED: ' + pe);
      }
    }

    /* Emails are wrapped separately so a mail failure is written into the
       "Email log" column instead of silently killing the whole claim. */
    try { sendReceivedEmail_(d, id); }
    catch (e1) { errs.push('APPLICANT EMAIL FAILED: ' + e1); }
    try { notifyOwner_(d, id); }
    catch (e2) { errs.push('OWNER EMAIL FAILED: ' + e2); }
    if (errs.length) {
      sh.getRange(rowNum, COL.LOG).setValue(
        Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') +
        ' — ' + errs.join(' | '));
    }

    return jsonOut_({ ok: true, id: id, emailErrors: errs.length ? errs : null });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

/**
 * Status lookup from the site. Called as JSONP (a <script> tag), because Apps Script
 * cannot send CORS headers. Only returns non-sensitive fields, and only when the
 * tracking ID AND the email on file both match.
 */
function doGet(e) {
  var cb = (e && e.parameter && e.parameter.callback) || 'callback';
  var out;
  try {
    var id = ((e.parameter.id) || '').toString().trim().toUpperCase();
    var email = ((e.parameter.email) || '').toString().trim().toLowerCase();
    var res = lookup_(id, email);
    out = { ok: true, found: res.found, id: id, status: res.status,
            appliance: res.appliance, amount: res.amount, updated: res.updated };
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  return ContentService
    .createTextOutput(cb + '(' + JSON.stringify(out) + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function lookup_(id, email) {
  var sh = getSheet_();
  var last = sh.getLastRow();
  if (last < 2 || !id) return { found: false };
  var vals = sh.getRange(2, 1, last - 1, COL.PHOTO).getValues();
  for (var i = 0; i < vals.length; i++) {
    var rowId = String(vals[i][COL.ID - 1] || '').trim().toUpperCase();
    var rowEmail = String(vals[i][COL.EMAIL - 1] || '').trim().toLowerCase();
    if (rowId === id && rowEmail === email) {
      return {
        found: true,
        status: vals[i][COL.STATUS - 1] || STATUS.RECEIVED,
        appliance: vals[i][COL.APPLIANCE - 1] || '',
        amount: vals[i][COL.AMOUNT - 1] || '',
        updated: vals[i][COL.UPDATED - 1] ? new Date(vals[i][COL.UPDATED - 1]).toISOString() : ''
      };
    }
  }
  return { found: false };
}

function genId_() {
  var c = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789', s = '';
  for (var i = 0; i < 6; i++) s += c.charAt(Math.floor(Math.random() * c.length));
  return 'GHI-' + s;
}

/* ------------------------- EMAILS ------------------------- */

/**
 * Sends the message. If CFG.FROM_ADDRESS is set AND that address is verified as a
 * "Send mail as" alias in Gmail, the mail goes out with that exact From header.
 * Otherwise it falls back to GmailApp (which sends as the account running the script).
 */
function send_(to, subject, html, logRow) {
  if (CFG.FROM_ADDRESS) {
    try { sendFromAlias_(to, subject, html, CFG.FROM_ADDRESS); }
    catch (e) {
      GmailApp.sendEmail(to, subject, stripTags_(html), {
        htmlBody: html, name: CFG.FROM_NAME, replyTo: CFG.REPLY_TO
      });
    }
  } else {
    GmailApp.sendEmail(to, subject, stripTags_(html), {
      htmlBody: html, name: CFG.FROM_NAME, replyTo: CFG.REPLY_TO
    });
  }
  if (logRow) {
    var sh = getSheet_();
    var cell = sh.getRange(logRow, COL.LOG);
    var prev = cell.getValue();
    var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    cell.setValue((prev ? prev + '\n' : '') + stamp + ' — ' + subject);
    var up = sh.getRange(logRow, COL.UPDATED);
    up.setValue(new Date());
  }
}

/**
 * Builds and sends a raw RFC-822 message so the From header is exactly
 * CFG.FROM_ADDRESS. Requires the Gmail advanced service (Services → Gmail API)
 * and the address verified under Gmail ▸ Settings ▸ Accounts ▸ Send mail as.
 */
function sendFromAlias_(to, subject, html, fromAddress) {
  var raw =
    'From: ' + CFG.FROM_NAME + ' <' + fromAddress + '>\r\n' +
    'To: ' + to + '\r\n' +
    'Subject: ' + subject + '\r\n' +
    'Reply-To: ' + (CFG.REPLY_TO || fromAddress) + '\r\n' +
    'MIME-Version: 1.0\r\n' +
    'Content-Type: text/html; charset=UTF-8\r\n' +
    '\r\n' + html;
  var encoded = Utilities.base64EncodeWebSafe(
    Utilities.newBlob(raw, 'text/plain', 'UTF-8').getBytes());
  Gmail.Users.Messages.send({ raw: encoded }, 'me');
}

function stripTags_(html) {
  return html.replace(/<br\s*\/?>/gi, '\n').replace(/<\/p>/gi, '\n\n').replace(/<[^>]+>/g, '');
}

function layout_(title, bodyHtml) {
  return '<div style="font-family:Arial,Helvetica,sans-serif;max-width:600px;margin:0 auto;color:#12331f">' +
    '<div style="background:#0b3d2c;padding:22px;border-radius:14px 14px 0 0;text-align:center">' +
    '<div style="color:#fff;font-size:20px;font-weight:bold">&#127807; Green House Incentives</div>' +
    '<div style="color:#a7d8bd;font-size:13px;margin-top:4px">Cash back for energy-efficient homes</div></div>' +
    '<div style="background:#ffffff;padding:26px;border:1px solid #e2efe7;border-top:none">' +
    '<h2 style="margin:0 0 14px;font-size:19px;color:#0b3d2c">' + title + '</h2>' +
    bodyHtml +
    '</div>' +
    '<div style="background:#f2f8f4;padding:16px;border:1px solid #e2efe7;border-top:none;' +
    'border-radius:0 0 14px 14px;font-size:12px;color:#5b6f63;text-align:center">' +
    'Green House Incentives &middot; Austin, TX<br>' +
    'You are receiving this because you submitted a rebate claim with us. ' +
    'If that wasn\u2019t you, reply to this email and we\u2019ll close it.<br>' +
    '<b>Add greenhouseincentives@gmail.com to your contacts</b> so our replies reach your inbox.<br>' +
    'We will <b>never</b> ask for your bank password, card PIN, or SSN.</div></div>';
}

function idBox_(id) {
  return '<div style="background:#eef7f1;border:1px dashed #0f9d58;border-radius:10px;' +
    'padding:14px;text-align:center;margin:18px 0">' +
    '<div style="font-size:12px;letter-spacing:.08em;color:#4a6b58">TRACKING ID</div>' +
    '<div style="font-size:24px;font-weight:bold;color:#0b3d2c">' + id + '</div></div>';
}

/** 1 — automatic, sent the moment the application arrives. */
function sendReceivedEmail_(d, id) {
  var body =
    '<p>Hi ' + (d.name || 'there').toString().split(' ')[0] + ',</p>' +
    '<p>Thanks for applying. Your claim is in the queue and nothing else is needed from you right now.</p>' +
    idBox_(id) +
    '<table style="width:100%;font-size:14px;border-collapse:collapse">' +
    row_('Appliance', d.appliance) + row_('Brand &amp; model', d.brand) +
    row_('Purchase date', d.date) + row_('Retailer', d.store) +
    row_('Purchase price', d.price ? '$' + d.price : '') +
    row_('State', d.state) + row_('Status', 'Received') +
    '</table>' +
    '<p style="margin-top:20px">Most claims are reviewed within <b>10 business days</b>. ' +
    'We\u2019ll email you the decision \u2014 no need to check back.</p>' +
    '<p style="font-size:13px;color:#5b6f63">Keep your receipt until the claim closes; we may ask for a clearer photo.</p>';
  send_(d.email, 'Your Green House Incentives claim ' + id, layout_('We got your claim &#127807;', body));
}

function notifyOwner_(d, id) {
  var body = '<p>New claim submitted on the site.</p>' +
    '<table style="width:100%;font-size:14px;border-collapse:collapse">' +
    row_('Tracking ID', id) + row_('Name', d.name) + row_('Email', d.email) +
    row_('Phone', d.phone) + row_('State', d.state) + row_('Appliance', d.appliance) +
    row_('Brand &amp; model', d.brand) + row_('Price', d.price ? '$' + d.price : '') +
    row_('Retailer', d.store) + row_('Receipt #', d.receipt) + row_('Notes', d.notes) +
    '</table>' +
    '<p style="margin-top:18px">Open the sheet to review, then use the menu to send an ' +
    '<b>Approval</b> or <b>More info</b> email.</p>';
  var to = ownerEmail_();
  if (to) send_(to, 'New claim ' + id + ' — ' + (d.appliance || ''), layout_('New claim submitted', body));
}

function row_(k, v) {
  if (v === undefined || v === null || v === '') return '';
  return '<tr><td style="padding:6px 0;color:#5b6f63;width:38%">' + k + '</td>' +
         '<td style="padding:6px 0;font-weight:bold">' + v + '</td></tr>';
}

/* ------------------------- SHEET MENU ------------------------- */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🌿 Green House')
    .addItem('1. Set up / reset Claims sheet', 'setupSheet')
    .addSeparator()
    .addItem('\U0001F4F7 View receipt photo (selected row)', 'showReceiptPhoto')
    .addItem('Request a CLEARER receipt photo', 'menuRequestClearerPhoto')
    .addSeparator()
    .addItem('Send APPROVAL email (selected row)', 'menuSendApproval')
    .addItem('Send MORE-INFO email (selected row)', 'menuSendInfo')
    .addItem('Send PAID email (selected row)', 'menuSendPaid')
    .addItem('Confirm DEPOSIT received \u2192 activate transfer', 'menuConfirmDeposit')
    .addSeparator()
    .addItem('Send a test confirmation email', 'sendTestEmail')
    .addItem('Diagnose email problems', 'diagnoseEmail')
    .addToUi();
}

function selectedRows_() {
  var sh = getSheet_();
  var r = sh.getActiveRange();
  if (!r) return [];
  var out = [], start = Math.max(2, r.getRow()), end = Math.min(sh.getLastRow(), r.getLastRow());
  for (var i = start; i <= end; i++) out.push(i);
  return out;
}

function rowData_(rowNum) {
  var sh = getSheet_();
  var v = sh.getRange(rowNum, 1, 1, COL.PHOTO).getValues()[0];
  return {
    row: rowNum, id: v[COL.ID - 1], name: v[COL.NAME - 1], email: v[COL.EMAIL - 1],
    phone: v[COL.PHONE - 1], state: v[COL.STATE - 1], appliance: v[COL.APPLIANCE - 1],
    date: v[COL.DATE - 1], brand: v[COL.BRAND - 1], price: v[COL.PRICE - 1],
    store: v[COL.STORE - 1], receipt: v[COL.RECEIPT - 1], notes: v[COL.NOTES - 1],
    status: v[COL.STATUS - 1], amount: v[COL.AMOUNT - 1],
    deposit: v[COL.DEPOSIT - 1], depositPaid: v[COL.DEPOSIT_PAID - 1],
    photo: v[COL.PHOTO - 1]
  };
}

function setStatus_(rowNum, status, amount) {
  var sh = getSheet_();
  sh.getRange(rowNum, COL.STATUS).setValue(status);
  if (amount !== undefined && amount !== null && amount !== '') {
    sh.getRange(rowNum, COL.AMOUNT).setValue(amount);
  }
  sh.getRange(rowNum, COL.UPDATED).setValue(new Date());
}

/** 3 — APPROVAL: you pick the row, confirm the amount, and the applicant gets their approval email. */
function menuSendApproval() {
  var sh = getSheet_();
  var rows = selectedRows_();
  if (!rows.length) { alert_('Select a claim row first (click any cell in the row).'); return; }
  var ui = SpreadsheetApp.getUi();

  var rowNum = rows[0];
  var d = rowData_(rowNum);
  if (!d.email) { alert_('That row has no email address.'); return; }

  var catMax = CFG.REBATES[d.appliance] || 0;
  var suggested = d.amount || (catMax || '');
  var res = ui.prompt('Approve claim ' + d.id,
    'Rebate amount in USD (numbers only).' +
    (catMax ? ' Max for ' + d.appliance + ': $' + catMax + '.' : ''),
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var amount = String(res.getResponseText()).replace(/[^0-9.]/g, '');
  if (!amount) amount = String(suggested).replace(/[^0-9.]/g, '');
  if (!amount) { alert_('No amount entered — nothing sent.'); return; }

  /* Security deposit required to activate the transfer. */
  var depRes = ui.prompt('Security deposit for ' + d.id,
    'Deposit amount in USD required to activate this transfer (0 for none):',
    ui.ButtonSet.OK_CANCEL);
  if (depRes.getSelectedButton() !== ui.Button.OK) return;
  var deposit = String(depRes.getResponseText()).replace(/[^0-9.]/g, '');
  if (!deposit) deposit = String(CFG.DEPOSIT_DEFAULT || 0);
  deposit = parseFloat(deposit) || 0;
  sh.getRange(rowNum, COL.DEPOSIT).setValue(deposit);

  var depositBlock = deposit > 0
    ? '<div style="background:#fff8e6;border-left:4px solid #e6a700;padding:14px;margin:18px 0">' +
      '<h3 style="font-size:15px;margin:0 0 8px;color:#8a6100">Security deposit required to activate your transfer</h3>' +
      '<p style="margin:0 0 8px">To activate your transfer, a security deposit of <b>$' + deposit + '</b> is required. ' +
      'This deposit ensures that funds go directly to verified participants and protects the program against fraudulent claims.</p>' +
      '<ul style="margin:0 0 10px;padding-left:20px;line-height:1.6">' +
      '<li><b>Security deposit due:</b> $' + deposit + '</li>' +
      '<li><b>Transfer amount:</b> $' + amount + ' — released once your deposit is confirmed</li>' +
      '<li><b>How to pay:</b> ' + CFG.DEPOSIT_INSTRUCTIONS + '</li></ul>' +
      '<p style="margin:0">Your claim is held as <b>Awaiting deposit</b> until we confirm it. ' +
      'The deposit is then applied and your transfer is activated and paid on the next payment run.</p></div>'
    : '';

  var body =
    '<p>Good news, ' + (d.name || 'there').toString().split(' ')[0] + ' — your claim is <b>approved</b>.</p>' +
    idBox_(d.id) +
    '<div style="background:#eef7f1;border-radius:10px;padding:16px;text-align:center;margin:16px 0">' +
    '<div style="font-size:13px;color:#4a6b58">APPROVED REBATE</div>' +
    '<div style="font-size:30px;font-weight:bold;color:#0f9d58">$' + amount + '</div></div>' +
    '<table style="width:100%;font-size:14px;border-collapse:collapse">' +
    row_('Appliance', d.appliance) + row_('Brand &amp; model', d.brand) +
    row_('Retailer', d.store) + row_('Purchase date', d.date) + row_('State', d.state) +
    '</table>' +
    depositBlock +
    '<h3 style="font-size:15px;margin:22px 0 6px">Next step: choose how you\u2019d like to be paid</h3>' +
    '<p style="margin:0 0 6px">Just reply to this email with one of the following:</p>' +
    '<ul style="margin:0 0 16px;padding-left:20px;line-height:1.6">' +
    '<li><b>Direct deposit</b> — routing number and account number. We\u2019ll confirm the deposit by email.</li>' +
    '<li><b>Mailed check</b> — the mailing address to send it to.</li></ul>' +
    '<p style="margin:0">Once we have that, payment goes out within <b>4\u20136 weeks</b>.</p>' +
    '<div style="background:#fff8e6;border-left:4px solid #e6a700;padding:12px;margin:18px 0;font-size:13px">' +
    '<b>Stay scam-safe:</b> we will never ask for your online banking password, your card PIN, ' +
    'a one-time code, or your Social Security number. If anyone asks for those claiming to be us, it isn\u2019t us.</div>';

  send_(d.email, 'Your rebate is approved — claim ' + d.id,
    layout_('Your rebate is approved &#127881;', body), rowNum);
  setStatus_(rowNum, deposit > 0 ? STATUS.AWAITING_DEPOSIT : STATUS.APPROVED, amount);
  alert_('Approval email sent to ' + d.email +
    (deposit > 0 ? '\n\nHeld as Awaiting deposit until the $' + deposit + ' deposit is confirmed.' : ''));
}

/** 2 — MORE INFO: ask the applicant for whatever is blocking the review. */
function menuSendInfo() {
  var rows = selectedRows_();
  if (!rows.length) { alert_('Select a claim row first.'); return; }
  var ui = SpreadsheetApp.getUi();
  var rowNum = rows[0], d = rowData_(rowNum);
  if (!d.email) { alert_('That row has no email address.'); return; }

  var res = ui.prompt('What do you need from the applicant?',
    'e.g. "a clearer photo of the receipt showing the model number"',
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var need = res.getResponseText() || 'a clearer photo of your receipt';

  var body =
    '<p>Hi ' + (d.name || 'there').toString().split(' ')[0] + ',</p>' +
    '<p>We\u2019re reviewing claim <b>' + d.id + '</b> and just need one more thing before we can decide:</p>' +
    '<div style="background:#eef7f1;border-radius:10px;padding:14px;margin:14px 0;font-weight:bold">' + need + '</div>' +
    '<p>Reply to this email with the document attached and we\u2019ll pick the review right back up. ' +
    'Claims are held for <b>30 days</b> while we wait.</p>' +
    row_('Appliance', d.appliance) + row_('Brand &amp; model', d.brand);
  var bodyHtml = '<p>Hi ' + (d.name || 'there').toString().split(' ')[0] + ',</p>' +
    '<p>We\u2019re reviewing claim <b>' + d.id + '</b> and just need one more thing before we can decide:</p>' +
    '<div style="background:#eef7f1;border-radius:10px;padding:14px;margin:14px 0;font-weight:bold">' + need + '</div>' +
    '<p>Reply to this email with the document attached and we\u2019ll pick the review right back up. ' +
    'Claims are held for <b>30 days</b> while we wait.</p>';

  send_(d.email, 'One more item needed for claim ' + d.id,
    layout_('We need a quick extra item', bodyHtml), rowNum);
  setStatus_(rowNum, STATUS.NEED_INFO);
  alert_('More-info email sent to ' + d.email);
}

/** 3b — DEPOSIT RECEIVED: activate the transfer. */
function menuConfirmDeposit() {
  var rows = selectedRows_();
  if (!rows.length) { alert_('Select a claim row first (click any cell in the row).'); return; }
  var sh = getSheet_();
  var rowNum = rows[0], d = rowData_(rowNum);
  if (!d.email) { alert_('That row has no email address.'); return; }

  var deposit = parseFloat(d.deposit) || 0;
  sh.getRange(rowNum, COL.DEPOSIT_PAID).setValue(new Date());

  var body =
    '<p>Hi ' + (d.name || 'there').toString().split(' ')[0] + ',</p>' +
    '<p>We have your security deposit of <b>$' + deposit + '</b> for claim <b>' + d.id + '</b>. Thank you.</p>' +
    '<div style="background:#eef7f1;border-radius:10px;padding:16px;margin:16px 0;text-align:center">' +
    '<div style="font-size:13px;color:#4a6b58">TRANSFER ACTIVATED</div>' +
    '<div style="font-size:28px;font-weight:bold;color:#0f9d58">$' + (d.amount || '0') + '</div>' +
    '<div style="margin-top:6px;font-size:13px">Deposit received: $' + deposit + '</div></div>' +
    '<p>Your transfer is now <b>active</b>. Payment is issued on the next payment run, within ' +
    '<b>4\u20136 weeks</b>, and we\u2019ll email you the moment it goes out.</p>';

  send_(d.email, 'Deposit confirmed \u2014 transfer active \u2014 claim ' + d.id,
    layout_('Your transfer is active &#128640;', body), rowNum);
  setStatus_(rowNum, STATUS.APPROVED);
  alert_('Deposit recorded. Transfer activated for ' + d.id + '.');
}

/* ------------------------- RECEIPT PHOTOS ------------------------- */

function getReceiptFolder_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('RECEIPT_FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var folders = DriveApp.getFoldersByName('GHI Receipts');
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('GHI Receipts');
  props.setProperty('RECEIPT_FOLDER_ID', folder.getId());
  return folder;
}

function saveReceiptPhoto_(d, id) {
  var clean = String(d.photo).replace(/^data:image\/[^;]+;base64,/, '');
  var bytes = Utilities.base64Decode(clean);
  var name = String(d.photoName || 'receipt').replace(/[^\w.\-() ]/g, '_');
  var blob = Utilities.newBlob(bytes, 'image/jpeg', id + ' - ' + name);
  var file = getReceiptFolder_().createFile(blob);
  file.setDescription('Receipt for claim ' + id);
  return file.getUrl();
}

function photoFileId_(url) {
  url = String(url || '');
  var m = url.match(/\/d\/([-\w]+)/) || url.match(/[?&]id=([-\w]+)/);
  if (m) return m[1];
  return /^[-\w]{25,}$/.test(url) ? url : '';
}

/** View the receipt photo for the selected row. */
function showReceiptPhoto() {
  var rows = selectedRows_();
  if (!rows.length) { alert_('Select a claim row first (click any cell in the row).'); return; }
  var d = rowData_(rows[0]);
  var fileId = photoFileId_(d.photo);
  if (!fileId) { alert_('No receipt photo stored for ' + d.id + '.'); return; }

  var file;
  try { file = DriveApp.getFileById(fileId); }
  catch (e) { alert_('Could not open that photo: ' + e); return; }

  var b64 = Utilities.base64Encode(file.getBlob().getBytes());
  var html =
    '<div style="font-family:Arial,Helvetica,sans-serif;padding:12px;color:#12331f">' +
    '<h3 style="margin:0 0 4px">' + d.id + ' \u2014 ' + (d.appliance || '') + '</h3>' +
    '<p style="margin:0 0 10px;color:#5b6f63;font-size:13px">' +
      'Name: <b>' + d.name + '</b> &middot; Retailer: <b>' + d.store + '</b> &middot; ' +
      'Receipt #: <b>' + d.receipt + '</b> &middot; Date: <b>' + d.date + '</b> &middot; ' +
      'Price: <b>' + (d.price ? '$' + d.price : '') + '</b></p>' +
    '<img src="data:image/jpeg;base64,' + b64 + '" ' +
      'style="max-width:100%;border:1px solid #d6ebe0;border-radius:10px">' +
    '<p style="margin:12px 0 0;font-size:13px">Close this window, then use the ' +
      '<b>🌿 Green House</b> menu to <b>approve</b> the claim or ' +
      '<b>request a clearer photo</b>.</p></div>';

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(760).setHeight(760),
    'Receipt \u2014 ' + d.id);
}

/** Ask the applicant for a clearer photo. */
function menuRequestClearerPhoto() {
  var rows = selectedRows_();
  if (!rows.length) { alert_('Select a claim row first.'); return; }
  var ui = SpreadsheetApp.getUi();
  var rowNum = rows[0], d = rowData_(rowNum);
  if (!d.email) { alert_('That row has no email address.'); return; }

  var res = ui.prompt('What is unreadable?',
    'e.g. "the total and the purchase date are cut off"',
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var reason = res.getResponseText() || 'part of the receipt is unreadable';

  var body =
    '<p>Hi ' + (d.name || 'there').toString().split(' ')[0] + ',</p>' +
    '<p>Thanks for sending your receipt for claim <b>' + d.id + '</b>. ' +
    'We can\u2019t read it clearly enough yet \u2014 ' + reason + '.</p>' +
    '<p>Please reply to this email with a new photo where we can clearly see:</p>' +
    '<ul style="line-height:1.7">' +
    '<li>The <b>retailer name</b></li>' +
    '<li>The <b>date of purchase</b></li>' +
    '<li>The <b>total paid</b></li>' +
    '<li>The <b>model number</b>, if it appears</li></ul>' +
    '<p>Tips: lay the receipt flat, use good light, and make sure all four corners are in frame.</p>' +
    '<p>We\u2019ll hold your claim for <b>30 days</b> while we wait.</p>';

  send_(d.email, 'Clearer receipt photo needed \u2014 claim ' + d.id,
    layout_('We need a clearer photo of your receipt', body), rowNum);
  setStatus_(rowNum, STATUS.NEED_INFO);
  alert_('Clearer-photo request sent to ' + d.email);
}

/** 4 — PAID: confirm the money is on its way. */
function menuSendPaid() {
  var rows = selectedRows_();
  if (!rows.length) { alert_('Select a claim row first.'); return; }
  var ui = SpreadsheetApp.getUi();
  var rowNum = rows[0], d = rowData_(rowNum);
  if (!d.email) { alert_('That row has no email address.'); return; }

  var res = ui.prompt('Payment details',
    'e.g. "Direct deposit of $4,500 sent 12 March — allow 1\u20133 business days"',
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var detail = res.getResponseText() || 'Your payment has been issued.';

  var body =
    '<p>Hi ' + (d.name || 'there').toString().split(' ')[0] + ',</p>' +
    '<p>Your rebate for claim <b>' + d.id + '</b> has been sent.</p>' +
    '<div style="background:#eef7f1;border-radius:10px;padding:16px;margin:16px 0">' +
    '<div style="font-size:13px;color:#4a6b58">AMOUNT PAID</div>' +
    '<div style="font-size:26px;font-weight:bold;color:#0f9d58">' +
    (d.amount ? '$' + d.amount : '') + '</div>' +
    '<div style="margin-top:8px">' + detail + '</div></div>' +
    '<p>Thanks for making your home more efficient &#127807;</p>';

  send_(d.email, 'Your rebate payment has been sent — claim ' + d.id,
    layout_('Payment sent &#128176;', body), rowNum);
  setStatus_(rowNum, STATUS.PAID);
  alert_('Paid email sent to ' + d.email);
}

/** Run this from the editor: it tells you exactly why mail is or isn't working. */
function diagnoseEmail() {
  var out = [];
  try { out.push('Script runs as: ' + Session.getActiveUser().getEmail()); }
  catch (e) { out.push('Could not read the active user: ' + e); }
  try { out.push('Daily mail quota left: ' + MailApp.getRemainingDailyQuota()); }
  catch (e) { out.push('Could not read mail quota: ' + e); }
  out.push('Mail is sent by this Google account: ' + Session.getActiveUser().getEmail());
  out.push('Configured From address: ' + (CFG.FROM_ADDRESS || '(none — using the account above)'));
  try {
    if (CFG.FROM_ADDRESS) {
      sendFromAlias_(Session.getActiveUser().getEmail(),
        'GHI from-address test ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm:ss'),
        '<p>This message should show From: ' + CFG.FROM_ADDRESS + '</p>', CFG.FROM_ADDRESS);
      out.push('Alias send: SUCCESS — check that the From line reads ' + CFG.FROM_ADDRESS + '.');
    } else {
      MailApp.sendEmail(Session.getActiveUser().getEmail(),
        'GHI mail test ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm:ss'),
        'If you can read this, sending works.');
      out.push('Test send: SUCCESS — check your inbox.');
    }
  } catch (e) {
    out.push('Send FAILED: ' + e);
    out.push('');
    out.push('If the error mentions Gmail is not defined: Services → + → Gmail API → Add.');
    out.push('If it says the address is not verified: Gmail → Settings → Accounts → Send mail as → add and confirm ' + (CFG.FROM_ADDRESS || 'your address') + '.');
  }
  SpreadsheetApp.getUi().alert(out.join('\n'));
}

function sendTestEmail() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('Send a test confirmation email to:', 'you@gmail.com', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var to = res.getResponseText().trim();
  var fake = { name: 'Test Applicant', email: to, appliance: 'Heat-pump water heater',
    brand: 'Brand X, Model Y', date: '2026-05-14', store: 'Home Depot',
    price: '2400', state: 'Texas', phone: '', receipt: 'INV-12345', notes: '' };
  var id = genId_();
  sendReceivedEmail_(fake, id);
  alert_('Test email sent to ' + to + ' with ID ' + id);
}

function alert_(msg) {
  SpreadsheetApp.getUi().alert(msg);
}
