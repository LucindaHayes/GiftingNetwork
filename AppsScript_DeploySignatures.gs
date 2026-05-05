/**
 * =====================================================================
 *  GiftingNetwork - Gmail Signature Deployment
 *  ---------------------------------------------------------------------
 *  Reads employees from the "Employees" sheet in the bound spreadsheet
 *  and pushes a standardized HTML signature into each user's Gmail
 *  account via the Gmail API, using a domain-wide-delegated service
 *  account to impersonate each user.
 *
 *  Required Script Property:
 *    SERVICE_ACCOUNT_KEY_JSON = full JSON key file of your service account
 *
 *  Required Library:
 *    OAuth2 by Google (script id below, added as "OAuth2")
 *    1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF
 * =====================================================================
 */

// ============ CONFIG ============
const CONFIG = {
  SHEET_NAME: 'Employees',
  LOG_SHEET_NAME: 'DeployLog',
  LOGO_URL: 'https://www.giftingnetwork.com/GNLogo.png',
  COMPANY_WEB: 'https://www.giftingnetwork.com',
  COMPANY_SOCIAL: {
    linkedin: 'https://www.linkedin.com/company/giftingnetwork',
    facebook: 'https://www.facebook.com/giftingnetwork',
    twitter:  'https://twitter.com/giftingnetwork'
  },
  ICON_BASE: 'https://www.giftingnetwork.com/icons',
  LOCKUP_URL: 'https://www.giftingnetwork.com/icons/gn-lockup.png',
  // Set STYLE to:
  //   'v1' = labeled style (Web / Email word labels)
  //   'v2' = modern FinTech minimalist (icon + HTML wordmark)
  //   'v3' = lockup style (single GN logo+wordmark image on top)
  STYLE: 'v3',
  SCOPES: ['https://www.googleapis.com/auth/gmail.settings.basic']
};

// ============ MENU ============
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('GiftingNetwork')
    .addItem('Preview signature HTML (row 2)', 'previewSignature_')
    .addItem('Deploy to ONE user (row 2)...',  'deployToSingle_')
    .addSeparator()
    .addItem('Deploy to ALL users',            'deployToAll_')
    .addSeparator()
    .addItem('Check service-account config',   'checkConfig_')
    .addToUi();
}

// ============ PUBLIC ENTRY POINTS ============

function deployToAll_() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    'Deploy signatures to ALL users?',
    'This will overwrite each listed employee\'s Gmail signature. Continue?',
    ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;

  const rows = readEmployees_();
  const log  = getOrCreateLogSheet_();

  let ok = 0, fail = 0;
  rows.forEach(r => {
    try {
      const html = buildSignature_(r);
      setSignatureForUser_(r.email, html);
      logResult_(log, r.email, 'OK', '');
      ok++;
    } catch (err) {
      logResult_(log, r.email, 'FAIL', String(err && err.message || err));
      fail++;
    }
  });

  ui.alert('Deploy complete',
    'Success: ' + ok + '\nFailed:  ' + fail + '\n\nSee the "' + CONFIG.LOG_SHEET_NAME + '" tab for details.',
    ui.ButtonSet.OK);
}

function deployToSingle_() {
  const rows = readEmployees_();
  if (!rows.length) throw new Error('No employees found on "' + CONFIG.SHEET_NAME + '" sheet.');
  const first = rows[0];
  const html = buildSignature_(first);
  setSignatureForUser_(first.email, html);
  SpreadsheetApp.getUi().alert('Signature pushed to ' + first.email);
}

function previewSignature_() {
  const rows = readEmployees_();
  if (!rows.length) throw new Error('No employees found on "' + CONFIG.SHEET_NAME + '" sheet.');
  const html = buildSignature_(rows[0]);
  const out = HtmlService.createHtmlOutput(html).setWidth(700).setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(out, 'Preview: ' + rows[0].email);
}

function checkConfig_() {
  const ui = SpreadsheetApp.getUi();
  const raw = PropertiesService.getScriptProperties().getProperty('SERVICE_ACCOUNT_KEY_JSON');
  if (!raw) { ui.alert('Missing Script Property SERVICE_ACCOUNT_KEY_JSON. See SETUP_GUIDE.md step 4.'); return; }
  try {
    const k = JSON.parse(raw);
    if (!k.client_email || !k.private_key) throw new Error('Key JSON missing client_email / private_key');
    ui.alert('Service account looks valid.\n\nclient_email:\n' + k.client_email +
             '\n\nMake sure this client_email (or its Unique ID) is authorized in Admin Console > Security > API Controls > Domain-wide Delegation, with scope:\n' +
             CONFIG.SCOPES.join(' '));
  } catch (e) {
    ui.alert('SERVICE_ACCOUNT_KEY_JSON is not valid JSON:\n' + e.message);
  }
}

// ============ CORE ============

function readEmployees_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!sh) throw new Error('Sheet "' + CONFIG.SHEET_NAME + '" not found.');
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values.shift().map(h => String(h).trim().toLowerCase());
  const idx = name => headers.indexOf(name);
  const iEmail = idx('email'), iName = idx('name'),
        iTitle = idx('title'), iLink = idx('linkedin'), iPhone = idx('phone');
  if (iEmail < 0 || iName < 0) {
    throw new Error('Employees sheet must have at least "Email" and "Name" columns.');
  }
  return values
    .filter(r => r[iEmail])
    .map(r => ({
      email:    String(r[iEmail]).trim(),
      name:     String(r[iName]).trim(),
      title:    iTitle >= 0 ? String(r[iTitle] || '').trim() : '',
      linkedin: iLink  >= 0 ? String(r[iLink]  || '').trim() : '',
      phone:    iPhone >= 0 ? String(r[iPhone] || '').trim() : ''
    }));
}

function buildSignature_(u) {
  if (CONFIG.STYLE === 'v1') return buildSignatureV1_(u);
  if (CONFIG.STYLE === 'v3') return buildSignatureV3_(u);
  return buildSignatureV2_(u);
}

/** v3 - Lockup style (full logo+wordmark image on top, divider, details below). */
function buildSignatureV3_(u) {
  const teal = '#00777B';
  const soc = CONFIG.COMPANY_SOCIAL;
  const linkedinHref = u.linkedin || soc.linkedin;
  const icon = (href, name, file) =>
    '<a href="' + href + '" style="text-decoration:none; border:0; margin-right:6px;">' +
      '<img src="' + CONFIG.ICON_BASE + '/' + file + '.png" ' +
        'alt="' + name + '" width="16" height="16" ' +
        'style="display:inline-block; border:0; vertical-align:middle;"></a>';
  const phoneLine = u.phone
    ? '<div style="font-size:12px; color:#1f1f1f;"><a href="tel:' + u.phone.replace(/[^0-9+]/g,'') +
      '" style="color:#1f1f1f; text-decoration:none;">' + escapeHtml_(u.phone) + '</a></div>'
    : '';
  return '' +
  '<table cellpadding="0" cellspacing="0" border="0" style="font-family:Arial,Helvetica,sans-serif; color:#1f1f1f; line-height:1.4;">' +
    '<tr><td style="padding-bottom:8px;">' +
      '<a href="' + CONFIG.COMPANY_WEB + '" style="text-decoration:none; border:0;">' +
        '<img src="' + CONFIG.LOCKUP_URL + '" alt="GiftingNetwork" width="240" ' +
          'style="display:block; width:240px; height:auto; border:0;"></a>' +
    '</td></tr>' +
    '<tr><td style="border-top:2px solid ' + teal + '; padding-top:10px;">' +
      '<div style="font-size:15px; font-weight:bold; color:#1f1f1f; letter-spacing:0.2px;">' + escapeHtml_(u.name) + '</div>' +
      (u.title ? '<div style="font-size:11px; color:#5a5a5a; margin-top:1px; text-transform:uppercase; letter-spacing:0.6px;">' + escapeHtml_(u.title) + '</div>' : '') +
      '<div style="font-size:12px; margin-top:8px; color:#1f1f1f;">' +
        '<a href="mailto:' + u.email + '" style="color:#1f1f1f; text-decoration:none;">' + escapeHtml_(u.email) + '</a></div>' +
      '<div style="font-size:12px; color:#1f1f1f;">' +
        '<a href="' + CONFIG.COMPANY_WEB + '" style="color:#1f1f1f; text-decoration:none;">giftingnetwork.com</a></div>' +
      phoneLine +
      '<div style="margin-top:10px;">' +
        icon(linkedinHref, 'LinkedIn', 'linkedin') +
        icon(soc.facebook, 'Facebook', 'facebook') +
        icon(soc.twitter, 'X', 'x') +
      '</div>' +
    '</td></tr>' +
  '</table>';
}

/** v1 - Labeled style (Web / Email word labels). */
function buildSignatureV1_(u) {
  const teal = '#00777B';
  const soc = CONFIG.COMPANY_SOCIAL;
  const linkedinHref = u.linkedin || soc.linkedin;
  const icon = (href, name, file) =>
    '<a href="' + href + '" style="text-decoration:none; border:0; margin-right:6px;">' +
      '<img src="' + CONFIG.ICON_BASE + '/' + file + '.png" ' +
        'alt="' + name + '" width="18" height="18" ' +
        'style="display:inline-block; border:0; vertical-align:middle;"></a>';
  const phoneLine = u.phone
    ? '<div style="font-size:12px; margin-top:2px; color:#333333;"><span style="color:' + teal +
      '; font-weight:bold;">Phone</span>&nbsp;<a href="tel:' + u.phone.replace(/[^0-9+]/g,'') +
      '" style="color:#333333; text-decoration:none;">' + escapeHtml_(u.phone) + '</a></div>'
    : '';
  return '' +
  '<table cellpadding="0" cellspacing="0" border="0" style="font-family:Arial,Helvetica,sans-serif; color:#333333; line-height:1.35;">' +
    '<tr>' +
      '<td style="vertical-align:middle; padding-right:16px;">' +
        '<a href="' + CONFIG.COMPANY_WEB + '" style="text-decoration:none; border:0;">' +
          '<img src="' + CONFIG.LOGO_URL + '" alt="GiftingNetwork" width="90" ' +
            'style="display:block; width:90px; height:auto; border:0;"></a></td>' +
      '<td style="border-left:2px solid ' + teal + '; padding:6px 0;">&nbsp;</td>' +
      '<td style="vertical-align:middle; padding-left:16px;">' +
        '<div style="font-size:15px; font-weight:bold; color:#222222; letter-spacing:0.2px;">' + escapeHtml_(u.name) + '</div>' +
        (u.title ? '<div style="font-size:11px; color:#666666; margin-top:2px; text-transform:uppercase; letter-spacing:0.6px;">' + escapeHtml_(u.title) + '</div>' : '') +
        '<div style="font-size:16px; font-weight:bold; margin-top:8px;">' +
          '<span style="color:' + teal + ';">Gifting</span><span style="color:#000000;">Network</span></div>' +
        '<div style="font-size:12px; margin-top:6px; color:#333333;">' +
          '<span style="color:' + teal + '; font-weight:bold;">Web</span>&nbsp;' +
          '<a href="' + CONFIG.COMPANY_WEB + '" style="color:#333333; text-decoration:none;">www.giftingnetwork.com</a></div>' +
        '<div style="font-size:12px; margin-top:2px; color:#333333;">' +
          '<span style="color:' + teal + '; font-weight:bold;">Email</span>&nbsp;' +
          '<a href="mailto:' + u.email + '" style="color:#333333; text-decoration:none;">' + escapeHtml_(u.email) + '</a></div>' +
        phoneLine +
        '<div style="margin-top:10px;">' +
          icon(linkedinHref, 'LinkedIn', 'linkedin') +
          icon(soc.facebook, 'Facebook', 'facebook') +
          icon(soc.twitter, 'X', 'x') +
        '</div>' +
      '</td>' +
    '</tr>' +
  '</table>';
}

/** v2 - Modern FinTech style (no word labels). */
function buildSignatureV2_(u) {
  const teal = '#00777B';
  const soc = CONFIG.COMPANY_SOCIAL;
  const linkedinHref = u.linkedin || soc.linkedin;
  const icon = (href, name, file) =>
    '<a href="' + href + '" style="text-decoration:none; border:0; margin-right:6px;">' +
      '<img src="' + CONFIG.ICON_BASE + '/' + file + '.png" ' +
        'alt="' + name + '" width="16" height="16" ' +
        'style="display:inline-block; border:0; vertical-align:middle;"></a>';
  const phoneLine = u.phone
    ? '<div style="font-size:12px; color:#1f1f1f;"><a href="tel:' + u.phone.replace(/[^0-9+]/g,'') +
      '" style="color:#1f1f1f; text-decoration:none;">' + escapeHtml_(u.phone) + '</a></div>'
    : '';
  return '' +
  '<table cellpadding="0" cellspacing="0" border="0" style="font-family:Arial,Helvetica,sans-serif; color:#1f1f1f; line-height:1.4;">' +
    '<tr>' +
      '<td style="vertical-align:middle; padding-right:18px;">' +
        '<a href="' + CONFIG.COMPANY_WEB + '" style="text-decoration:none; border:0;">' +
          '<img src="' + CONFIG.LOGO_URL + '" alt="GiftingNetwork" width="80" ' +
            'style="display:block; width:80px; height:auto; border:0;"></a></td>' +
      '<td style="border-left:2px solid ' + teal + '; padding:8px 0;">&nbsp;</td>' +
      '<td style="vertical-align:middle; padding-left:18px;">' +
        '<div style="font-size:15px; font-weight:bold; color:#1f1f1f; letter-spacing:0.2px;">' + escapeHtml_(u.name) + '</div>' +
        (u.title ? '<div style="font-size:11px; color:#5a5a5a; margin-top:1px;">' + escapeHtml_(u.title) + '</div>' : '') +
        '<div style="font-size:15px; font-weight:bold; margin-top:10px;">' +
          '<span style="color:' + teal + ';">Gifting</span><span style="color:#000000;">Network</span></div>' +
        '<div style="font-size:12px; margin-top:6px; color:#1f1f1f;">' +
          '<a href="mailto:' + u.email + '" style="color:#1f1f1f; text-decoration:none;">' + escapeHtml_(u.email) + '</a></div>' +
        '<div style="font-size:12px; color:#1f1f1f;">' +
          '<a href="' + CONFIG.COMPANY_WEB + '" style="color:#1f1f1f; text-decoration:none;">giftingnetwork.com</a></div>' +
        phoneLine +
        '<div style="margin-top:10px;">' +
          icon(linkedinHref, 'LinkedIn', 'linkedin') +
          icon(soc.facebook, 'Facebook', 'facebook') +
          icon(soc.twitter, 'X', 'x') +
        '</div>' +
      '</td>' +
    '</tr>' +
  '</table>';
}

function setSignatureForUser_(userEmail, html) {
  const token = getDelegatedToken_(userEmail);
  const listResp = UrlFetchApp.fetch(
    'https://gmail.googleapis.com/gmail/v1/users/me/settings/sendAs',
    { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
  );
  if (listResp.getResponseCode() >= 300) {
    throw new Error('List sendAs failed (' + listResp.getResponseCode() + '): ' + listResp.getContentText());
  }
  const sendAsList = JSON.parse(listResp.getContentText()).sendAs || [];
  const primary = sendAsList.find(s => s.isPrimary) || sendAsList.find(s => s.sendAsEmail === userEmail);
  if (!primary) throw new Error('No primary sendAs found for ' + userEmail);
  const patchUrl = 'https://gmail.googleapis.com/gmail/v1/users/me/settings/sendAs/' +
                   encodeURIComponent(primary.sendAsEmail);
  const patchResp = UrlFetchApp.fetch(patchUrl, {
    method: 'patch',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({ signature: html }),
    muteHttpExceptions: true
  });
  if (patchResp.getResponseCode() >= 300) {
    throw new Error('Patch signature failed (' + patchResp.getResponseCode() + '): ' + patchResp.getContentText());
  }
}

function getDelegatedToken_(userEmail) {
  const raw = PropertiesService.getScriptProperties().getProperty('SERVICE_ACCOUNT_KEY_JSON');
  if (!raw) throw new Error('Missing Script Property SERVICE_ACCOUNT_KEY_JSON.');
  const key = JSON.parse(raw);
  const service = OAuth2.createService('gn-sig-' + userEmail)
    .setTokenUrl('https://oauth2.googleapis.com/token')
    .setPrivateKey(key.private_key)
    .setIssuer(key.client_email)
    .setSubject(userEmail)
    .setPropertyStore(PropertiesService.getScriptProperties())
    .setCache(CacheService.getScriptCache())
    .setScope(CONFIG.SCOPES.join(' '));
  service.reset();
  if (!service.hasAccess()) {
    throw new Error('OAuth failure for ' + userEmail + ': ' + service.getLastError());
  }
  return service.getAccessToken();
}

function getOrCreateLogSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.LOG_SHEET_NAME);
    sh.appendRow(['Timestamp', 'Email', 'Status', 'Detail']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function logResult_(sh, email, status, detail) {
  sh.appendRow([new Date(), email, status, detail]);
}

function escapeHtml_(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
