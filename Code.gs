/**
 * ITP Test Form – Backend (Google Apps Script)
 * - Serves the web form (index.html)
 * - Saves PDF to Google Drive, emails it as an attachment, pushes a LINE group message with the link
 *
 * Script Properties (Project Settings > Script properties):
 *   DEFAULT_TO  : default email recipients (optional, comma separated)
 *   LINE_TOKEN  : LINE Messaging API – Channel access token (long-lived)
 *   LINE_TO     : LINE groupId / userId to push to (type "id" in the group to get it)
 *   FOLDER_ID   : Drive folder id for PDFs (optional – auto-created "ITP Test Reports")
 *   API_KEY     : optional – required only when the form is hosted outside Apps Script
 */
const PROPS = PropertiesService.getScriptProperties();

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ITP Test Form – Sky Ground Residence')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) {}
  // LINE webhook
  if (body.events) { handleLineWebhook_(body.events); return ContentService.createTextOutput('OK'); }
  // Form submit from an externally hosted page
  try {
    const key = PROPS.getProperty('API_KEY');
    if (key && body.key !== key) throw new Error('API key ไม่ถูกต้อง');
    return json_(submitReport(body));
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

/** Called by the form: google.script.run.submitReport(payload) */
function submitReport(p) {
  if (!p || !p.pdf) throw new Error('ไม่มีไฟล์ PDF');
  const blob = Utilities.newBlob(Utilities.base64Decode(p.pdf), 'application/pdf', p.fileName || 'ITP_Report.pdf');
  const file = getFolder_().createFile(blob);
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  const url = file.getUrl();
  const res = { ok: true, url: url, email: false, line: false };
  try { res.folderUrl = getFolder_().getUrl(); } catch (e) {}
  res.appUrl = appUrl_();
  let xlsxBlob = null;
  if (p.xlsx) {
    xlsxBlob = Utilities.newBlob(Utilities.base64Decode(p.xlsx), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', p.xlsxName || 'ITP_Report.xlsx');
    const xf = getFolder_().createFile(xlsxBlob);
    try { xf.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
    res.xlsxUrl = xf.getUrl();
  }
  const L = buildLinks_(res);   // set after all URLs are known

  if (p.sendEmail) {
    const to = String(p.to || PROPS.getProperty('DEFAULT_TO') || '').trim();
    if (!to) res.emailError = 'ไม่มีอีเมลผู้รับ';
    else {
      try {
        const atts = xlsxBlob ? [blob, xlsxBlob] : [blob];
        const size = atts.reduce(function (s, b) { return s + b.getBytes().length; }, 0);
        const big = size > 20 * 1024 * 1024;
        const opt = { cc: p.cc || '', name: 'ITP Test Form' };
        if (!big) opt.attachments = atts;
        const note = big ? '\n(ไฟล์ใหญ่เกิน แนบเป็นลิงก์แทน)' : '';
        opt.htmlBody = '<div style="font-family:Arial,Tahoma,sans-serif;font-size:14px;line-height:1.6">' +
          escHtml_(p.summary || '').replace(/\n/g, '<br>') + '<br><br>' + L.html + escHtml_(note).replace(/\n/g, '<br>') + '</div>';
        MailApp.sendEmail(to, p.subject || 'ITP Test Report', (p.summary || '') + '\n\n' + L.text + note, opt);
        res.email = true;
      } catch (e) { res.emailError = String(e.message || e); }
    }
  }

  if (p.sendLine) {
    const token = PROPS.getProperty('LINE_TOKEN'), lineTo = PROPS.getProperty('LINE_TO');
    if (!token || !lineTo) res.lineError = 'ยังไม่ได้ตั้งค่า LINE_TOKEN / LINE_TO';
    else {
      try { pushLine_(token, lineTo, (p.summary || '') + '\n\n' + L.text); res.line = true; }
      catch (e) { res.lineError = String(e.message || e); }
    }
  }
  return res;
}

/** Web app URL (/exec) – the same link that is used for the form and the LINE webhook */
function appUrl_() {
  try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; }
}

/** One shared set of links used by the web page, e-mail and LINE */
function buildLinks_(res) {
  const items = [];
  if (res.url) items.push({ icon: '📄', label: 'PDF', url: res.url });
  if (res.xlsxUrl) items.push({ icon: '📊', label: 'Excel', url: res.xlsxUrl });
  if (res.folderUrl) items.push({ icon: '📁', label: 'โฟลเดอร์รายงานทั้งหมด', url: res.folderUrl });
  if (res.appUrl) items.push({ icon: '📝', label: 'เปิดฟอร์ม ITP Test Form', url: res.appUrl });
  return {
    text: items.map(function (i) { return i.icon + ' ' + i.label + ': ' + i.url; }).join('\n'),
    html: items.map(function (i) { return i.icon + ' <a href="' + i.url + '">' + i.label + '</a>'; }).join('<br>')
  };
}

function escHtml_(t) {
  return String(t == null ? '' : t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function pushLine_(token, to, text) {
  const r = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({ to: to, messages: [{ type: 'text', text: text.slice(0, 4900) }] })
  });
  if (r.getResponseCode() !== 200) throw new Error('LINE ' + r.getResponseCode() + ': ' + r.getContentText());
}

/** Add the bot to a group, type "id" → bot replies with the groupId to put in LINE_TO */
function handleLineWebhook_(events) {
  const token = PROPS.getProperty('LINE_TOKEN');
  events.forEach(function (ev) {
    const src = ev.source || {};
    const id = src.groupId || src.roomId || src.userId || '';
    if (id) PROPS.setProperty('LAST_SOURCE_ID', id);
    const isId = ev.type === 'message' && ev.message && ev.message.type === 'text' && ev.message.text.trim().toLowerCase() === 'id';
    if (token && ev.replyToken && (isId || ev.type === 'join')) {
      UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        headers: { Authorization: 'Bearer ' + token },
        payload: JSON.stringify({ replyToken: ev.replyToken, messages: [{ type: 'text', text: 'ID ของห้องนี้:\n' + id + '\n\nนำไปใส่ใน Script property: LINE_TO' }] })
      });
    }
  });
}

function getFolder_() {
  const id = PROPS.getProperty('FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const it = DriveApp.getFoldersByName('ITP Test Reports');
  const f = it.hasNext() ? it.next() : DriveApp.createFolder('ITP Test Reports');
  PROPS.setProperty('FOLDER_ID', f.getId());
  return f;
}

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/** Run once from the editor to grant Drive / Gmail / external-request permissions */
function setup() {
  getFolder_();
  Logger.log('Folder: ' + getFolder_().getUrl());
  Logger.log('Mail quota remaining today: ' + MailApp.getRemainingDailyQuota());
  UrlFetchApp.fetch('https://api.line.me', { muteHttpExceptions: true });
}

/** Optional: test LINE push after setting LINE_TOKEN / LINE_TO */
function testLine() { pushLine_(PROPS.getProperty('LINE_TOKEN'), PROPS.getProperty('LINE_TO'), '✅ ทดสอบส่งจาก ITP Test Form'); }
