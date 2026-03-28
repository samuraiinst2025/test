const FREEE = {
CLIENT_ID: '650276714605798',
CLIENT_SECRET: '9259P92-mAP5puwqAsluskIgeNMCB0Af8W1C9KI6FjbirOaTaiHVz9lFAfX8hY6SL5PoGtPdBI_kuI1HO63Png',
AUTH_URL: 'https://accounts.secure.freee.co.jp/public_api/authorize',
TOKEN_URL: 'https://accounts.secure.freee.co.jp/public_api/token',
API_BASE: 'https://api.freee.co.jp',
SCOPE: 'read',
};

function getService_() {
  return OAuth2.createService('freee')
    .setAuthorizationBaseUrl(FREEE.AUTH_URL)
    .setTokenUrl(FREEE.TOKEN_URL)
    .setClientId(FREEE.CLIENT_ID)
    .setClientSecret(FREEE.CLIENT_SECRET)
    .setCallbackFunction('authCallback')
    // 生徒さん環境でも安全に「ユーザーごと」に持つのが無難
    .setPropertyStore(PropertiesService.getUserProperties())
    .setScope(FREEE.SCOPE)
    // refresh token を取るための定番（必要に応じて）
    .setParam('access_type', 'offline')
    .setParam('prompt', 'consent');
}

function logRedirectUri() {
  const service = getService_();
  Logger.log('Redirect URI: ' + service.getRedirectUri());
  SpreadsheetApp.getUi().alert('Redirect URI をログに出しました（表示 > ログ）');
}

function authorize() {
  const service = getService_();
  if (service.hasAccess()) {
    SpreadsheetApp.getUi().alert('すでに認可済みです');
    return;
  }
  const url = service.getAuthorizationUrl();
  // 修正点1: HTMLテンプレートリテラル内の不要なバックスラッシュを削除
  const html = HtmlService.createHtmlOutput(`<p>このリンクを開いてfreeeで許可してください：</p><p><a href="${url}" target="_blank">freeeを認可する</a></p><p>許可後にこのスプレッドシートへ戻って、③疎通テストを実行してください。</p>`).setWidth(420).setHeight(220);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'freee 認可');
}

function authCallback(request) {
  const service = getService_();
  const authorized = service.handleCallback(request);
  
  if (authorized) {
    // 修正点2: 文字列を囲むシングルクォートを追加し、不要なバックスラッシュを削除
    return HtmlService.createHtmlOutput('認可OK、タブを閉じてスプレッドシートに戻り、③疎通テストを実行してください。<script>window.top.close();</script>');
  } else {
    return HtmlService.createHtmlOutput('認可NG。もう一度②からやり直してください。');
  }
}

function getAccessToken_() {
  const service = getService_();
  if (!service.hasAccess()) throw new Error('未認可です。② 認可する を実行してください。');
  return service.getAccessToken();
}

function testPing() {
  const me = apiGet_('/api/1/users/me', {});
  console.log(me); // これで十分（Cloudログに出る）
  log_('PING', JSON.stringify(me).slice(0, 2000)); // logシートに書くならこちら
  return me;
}

function testCompanies() {
  const res = apiGet_('/api/1/companies', {});
  log_('COMPANIES', JSON.stringify(res).slice(0, 2000));
  return res;
}

function apiGet_(path, params) {
  const token = getAccessToken_();
  const qs = Object.keys(params || {}).filter(k => params[k] !== '' && params[k] !== null && params[k] !== undefined).map(k => encodeURIComponent(k) + '=' + encodeURIComponent(String(params[k]))).join('&');
  
  const url = FREEE.API_BASE + path + (qs ? `?${qs}` : '');
  const resp = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true,
  });
  const code = resp.getResponseCode();
  const text = resp.getContentText();
  if (code < 200 || code >= 300) throw new Error(`${code} ${text}`);
  return text ? JSON.parse(text) : {};
}

function log_(tag, message) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('log') || ss.insertSheet('log');
  if (sh.getLastRow() === 0) sh.appendRow(['timestamp', 'tag', 'message']);
  sh.appendRow([new Date(), tag, message]);
}