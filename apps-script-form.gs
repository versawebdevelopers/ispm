/**
 * ISPM appointment form → email
 *
 * FIRST-TIME AUTH (required once after updates):
 * 1. Select function: authorizeMail → Run → Allow (Gmail / send email)
 * 2. Deploy → Manage deployments → pencil → New version → Deploy
 *
 * Test: YOUR_EXEC_URL?test=1
 *
 * Quota (Google Workspace): 1,500 emails/day via MailApp/GmailApp.
 * Resets daily — there is no separate monthly cap beyond that daily limit.
 */

/**
 * Who receives appointment leads (ISPM / client inbox).
 * Change this to whatever address they want — testing or production.
 */
var TO_EMAIL = 'info@sportsandpainmanagement.com';

/** Display name shown as the sender (From name) */
var FROM_NAME = 'Versa Creative';

/**
 * Exact Gmail "Send mail as" alias address.
 * Leave blank to auto-pick a non-flegall alias, or your primary account.
 * Reply-To is always the patient's form email, so Reply goes to the patient.
 */
var FROM_EMAIL = '';

function authorizeMail() {
  var me = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
  var from = resolveFrom_();
  var aliases = GmailApp.getAliases();
  var remaining = MailApp.getRemainingDailyQuota();

  GmailApp.sendEmail(me, 'ISPM form — mail permission OK', '', {
    name: FROM_NAME,
    from: from,
    htmlBody:
      '<p style="font-family:Arial,sans-serif;font-size:14px;color:#1a2b3c">' +
      '<strong>Mail is authorized.</strong><br><br>' +
      'Script account: ' +
      escape_(me) +
      '<br>Sending as: ' +
      escape_(FROM_NAME) +
      ' &lt;' +
      escape_(from) +
      '&gt;<br>Aliases: ' +
      escape_(aliases.length ? aliases.join(', ') : '(none)') +
      '<br>Daily quota remaining: ' +
      remaining +
      ' / 1500<br>Recipient for form leads: ' +
      escape_(TO_EMAIL) +
      '</p>'
  });

  Logger.log(
    'Auth mail from=' +
      from +
      ' aliases=' +
      aliases.join(',') +
      ' quota=' +
      remaining
  );
}

function doPost(e) {
  try {
    var data = (e && e.parameter) || {};

    if (data._honey) {
      return json_({ ok: true, skipped: 'honeypot' });
    }

    var required = ['first_name', 'last_name', 'email', 'phone', 'privacy_consent'];
    for (var i = 0; i < required.length; i++) {
      var key = required[i];
      if (!data[key] || String(data[key]).trim() === '') {
        return json_({ ok: false, error: 'Missing: ' + key });
      }
    }

    sendAppointment_(data);
    return json_({
      ok: true,
      sent: true,
      to: TO_EMAIL,
      from: resolveFrom_(),
      quota: MailApp.getRemainingDailyQuota()
    });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  var params = (e && e.parameter) || {};
  if (params.test === '1') {
    try {
      sendAppointment_({
        patient: 'new',
        location: 'humble',
        first_name: 'Test',
        last_name: 'Email',
        date_of_birth: '1967-12-12',
        email: TO_EMAIL,
        phone: '7133333456',
        zip_code: '77040',
        insurance: 'Blue Cross Blue Shield',
        reason: 'Manual test via ?test=1',
        privacy_consent: 'yes',
        sms_consent: 'yes'
      });
      return ContentService.createTextOutput(
        'Test email sent to ' +
          TO_EMAIL +
          ' from ' +
          resolveFrom_() +
          '. Quota left: ' +
          MailApp.getRemainingDailyQuota()
      );
    } catch (err) {
      return ContentService.createTextOutput('Test email FAILED: ' + String(err));
    }
  }

  return ContentService.createTextOutput(
    'ISPM form endpoint is live. Use POST from the website. Append ?test=1 to send a test email.'
  );
}

function sendAppointment_(data) {
  var first = properName_(data.first_name);
  var last = properName_(data.last_name);
  var fullName = (first + ' ' + last).trim() || 'Unknown';
  var patient = formatPatient_(data.patient);
  var location = formatLocation_(data.location);
  var email = String(data.email || '').trim();
  var phone = formatPhone_(data.phone);
  var phoneHref = String(data.phone || '').replace(/[^\d+]/g, '');
  var insurance = String(data.insurance || '—').trim() || '—';
  var reason = String(data.reason || '—').trim() || '—';
  var dob = formatDob_(data.date_of_birth);
  var zip = String(data.zip_code || '—').trim() || '—';
  var privacy = data.privacy_consent === 'yes' ? 'Agreed' : '—';
  var sms = data.sms_consent === 'yes' ? 'Opted in' : 'Not opted in';
  var submitted = new Date().toLocaleString('en-US', {
    timeZone: 'America/Chicago',
    dateStyle: 'medium',
    timeStyle: 'short'
  });
  var subject = 'ISPM appointment — ' + fullName + ' (' + location + ')';
  var from = resolveFrom_();

  var html =
    '<!DOCTYPE html><html><body style="margin:0;padding:0;background:#eef2f6;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#eef2f6;padding:24px 12px;">' +
    '<tr><td align="center">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:560px;background:#ffffff;border-radius:12px;overflow:hidden;border:1px solid #d7e0ea;">' +
    '<tr><td style="background:#193558;padding:28px 32px;">' +
    '<p style="margin:0 0 6px;font-family:Arial,Helvetica,sans-serif;font-size:12px;letter-spacing:0.06em;color:#8eabc1;">Interventional Sports &amp; Pain Management</p>' +
    '<h1 style="margin:0;font-family:Arial,Helvetica,sans-serif;font-size:22px;line-height:1.3;font-weight:700;color:#ffffff;">New Appointment Request</h1>' +
    '</td></tr>' +
    '<tr><td style="padding:20px 32px;background:#f7fafc;border-bottom:1px solid #e4ebf2;">' +
    '<p style="margin:0 0 4px;font-family:Arial,Helvetica,sans-serif;font-size:18px;font-weight:700;color:#193558;">' +
    escape_(fullName) +
    '</p>' +
    '<p style="margin:0;font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#2c5578;">' +
    escape_(patient) +
    ' · ' +
    escape_(location) +
    ' Office</p>' +
    '</td></tr>' +
    '<tr><td style="padding:8px 32px 28px;font-family:Arial,Helvetica,sans-serif;color:#1a2b3c;">' +
    sectionTitle_('Contact') +
    fieldRow_('Email', email ? '<a href="mailto:' + escape_(email) + '" style="color:#193558;text-decoration:underline;">' + escape_(email) + '</a>' : '—', false) +
    fieldRow_('Phone', phone !== '—' ? '<a href="tel:' + escape_(phoneHref) + '" style="color:#193558;text-decoration:underline;">' + escape_(phone) + '</a>' : '—', true) +
    fieldRow_('Zip Code', escape_(zip), false) +
    sectionTitle_('Visit Details') +
    fieldRow_('Patient Type', escape_(patient), false) +
    fieldRow_('Preferred Location', escape_(location), true) +
    fieldRow_('Date Of Birth', escape_(dob), false) +
    fieldRow_('Insurance', escape_(insurance), true) +
    fieldRow_('Reason For Visit', escape_(reason).replace(/\n/g, '<br>'), false) +
    sectionTitle_('Consents') +
    fieldRow_('Privacy / Terms', escape_(privacy), false) +
    fieldRow_('SMS Updates', escape_(sms), true) +
    '<p style="margin:24px 0 0;font-size:12px;line-height:1.5;color:#6b7c8f;">Submitted ' +
    escape_(submitted) +
    ' (CT)<br>Reply to this email to respond directly to the patient.</p>' +
    '</td></tr>' +
    '<tr><td style="padding:16px 32px;background:#142a46;font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#8eabc1;">' +
    'Humble · 281-570-2961 &nbsp;|&nbsp; Baytown · 832-514-6300' +
    '</td></tr>' +
    '</table></td></tr></table></body></html>';

  var plain =
    'New Appointment Request\n\n' +
    fullName +
    ' · ' +
    patient +
    ' · ' +
    location +
    '\nEmail: ' +
    email +
    '\nPhone: ' +
    phone +
    '\nDate Of Birth: ' +
    dob +
    '\nZip Code: ' +
    zip +
    '\nInsurance: ' +
    insurance +
    '\nReason For Visit: ' +
    reason +
    '\nPrivacy / Terms: ' +
    privacy +
    '\nSMS Updates: ' +
    sms +
    '\nSubmitted: ' +
    submitted;

  GmailApp.sendEmail(TO_EMAIL, subject, plain, {
    htmlBody: html,
    name: FROM_NAME,
    from: from,
    replyTo: email || TO_EMAIL
  });
}

function resolveFrom_() {
  if (FROM_EMAIL) return FROM_EMAIL;
  var aliases = GmailApp.getAliases() || [];
  for (var i = 0; i < aliases.length; i++) {
    var a = String(aliases[i] || '').toLowerCase();
    if (a && a.indexOf('flegall') === -1) return aliases[i];
  }
  if (aliases.length) return aliases[0];
  return Session.getEffectiveUser().getEmail() || Session.getActiveUser().getEmail();
}

function sectionTitle_(label) {
  return (
    '<p style="margin:22px 0 10px;font-size:13px;font-weight:700;color:#f28b80;">' +
    escape_(label) +
    '</p>'
  );
}

function fieldRow_(label, valueHtml, alt) {
  var bg = alt ? '#f7fafc' : '#ffffff';
  return (
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;">' +
    '<tr>' +
    '<td width="38%" style="padding:10px 12px;background:' +
    bg +
    ';font-size:13px;font-weight:700;color:#2c5578;vertical-align:top;border-bottom:1px solid #e4ebf2;">' +
    escape_(label) +
    '</td>' +
    '<td style="padding:10px 12px;background:' +
    bg +
    ';font-size:14px;color:#1a2b3c;vertical-align:top;border-bottom:1px solid #e4ebf2;">' +
    valueHtml +
    '</td>' +
    '</tr></table>'
  );
}

function formatPatient_(v) {
  v = String(v || '').toLowerCase().trim();
  if (v === 'new') return 'New Patient';
  if (v === 'returning') return 'Returning Patient';
  return titleCase_(v) || '—';
}

function formatLocation_(v) {
  v = String(v || '').toLowerCase().trim();
  if (v === 'humble') return 'Humble';
  if (v === 'baytown') return 'Baytown';
  return titleCase_(v) || '—';
}

function formatDob_(v) {
  v = String(v || '').trim();
  if (!v) return '—';
  var m = v.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return v;
  var months = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
  ];
  return months[parseInt(m[2], 10) - 1] + ' ' + parseInt(m[3], 10) + ', ' + m[1];
}

function formatPhone_(v) {
  var d = String(v || '').replace(/\D/g, '');
  if (d.length === 11 && d.charAt(0) === '1') d = d.slice(1);
  if (d.length === 10) {
    return '(' + d.slice(0, 3) + ') ' + d.slice(3, 6) + '-' + d.slice(6);
  }
  return String(v || '').trim() || '—';
}

function properName_(s) {
  s = String(s || '').trim();
  if (!s) return '';
  return s
    .toLowerCase()
    .split(/\s+/)
    .map(function (w) {
      return w ? w.charAt(0).toUpperCase() + w.slice(1) : '';
    })
    .join(' ');
}

function titleCase_(s) {
  s = String(s || '').replace(/[_-]+/g, ' ').trim();
  if (!s) return '';
  return s.charAt(0).toUpperCase() + s.slice(1).toLowerCase();
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}

function escape_(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
