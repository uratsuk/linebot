/**
 * 貴金属相場表 PDF 自動送信
 *
 * 「催事用 貴金属相場表(自動計算)」の Apps Script に追加して使う。
 * 平日 10:02 に「相場表」シートを PDF にして Gmail で送る。
 *
 * 初回だけ setupSoubaPdfTrigger() を手動実行して権限を許可する。
 * 止めるときは removeSoubaPdfTriggers() を実行する。
 *
 * PDF 書き出しに Drive の権限が要るため、次の行は消さないこと。
 * DriveApp.getFiles()
 */

const SOUBA_PDF = {
  to: 'uratsu.k@gmail.com',
  sheetName: '相場表',
  range: 'A1:K29',
  settingsSheetName: '設定',
  dateCell: 'B9',          // 設定シートの「日付」(相場の取得日)
  sendHour: 10,
  sendMinute: 2,
  retryMinutes: 3,         // 本日分の相場がまだ入っていないときの再確認間隔
  giveUpHour: 11,          // この時刻を過ぎたら、本日分が無くても注意書きつきで送る
};

const SOUBA_PDF_DAILY = 'soubaPdfDaily';
const SOUBA_PDF_SEND = 'sendSoubaPdf';
const SOUBA_PDF_SENT_KEY = 'soubaPdfSentDate';

/** 初回に1回だけ実行する。毎朝9時台に、その日の 10:02 の送信を予約する。 */
function setupSoubaPdfTrigger() {
  removeSoubaPdfTriggers();
  ScriptApp.newTrigger(SOUBA_PDF_DAILY).timeBased().everyDays(1).atHour(9).create();
  // 今日の 10:02 がまだ先なら、今日の分もすぐ予約する
  soubaPdfDaily();
}

/** 自動送信を止める。 */
function removeSoubaPdfTriggers() {
  ScriptApp.getProjectTriggers()
    .filter(t => [SOUBA_PDF_DAILY, SOUBA_PDF_SEND].includes(t.getHandlerFunction()))
    .forEach(t => ScriptApp.deleteTrigger(t));
}

/** 毎朝実行。平日なら 10:02 ちょうどの1回きりのトリガーを作る。 */
function soubaPdfDaily() {
  const now = new Date();
  if (!isWeekday_(now)) return;
  const at = todayAt_(SOUBA_PDF.sendHour, SOUBA_PDF.sendMinute);
  if (at <= now) return;
  deleteSendTriggers_();
  ScriptApp.newTrigger(SOUBA_PDF_SEND).timeBased().at(at).create();
}

/** 相場表を PDF にしてメールする。手動で実行してもよい(今日まだ送っていなければ送る)。 */
function sendSoubaPdf() {
  deleteSendTriggers_();
  const now = new Date();
  const today = fmt_(now, 'yyyy/MM/dd');
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty(SOUBA_PDF_SENT_KEY) === today) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const isTodayRate = soubaDate_(ss) === today;

  // 本日分の相場がまだなら少し待って再確認する
  if (!isTodayRate && now.getHours() < SOUBA_PDF.giveUpHour) {
    ScriptApp.newTrigger(SOUBA_PDF_SEND).timeBased()
      .after(SOUBA_PDF.retryMinutes * 60 * 1000).create();
    return;
  }

  SpreadsheetApp.flush();
  const pdf = exportSheetPdf_(ss, SOUBA_PDF.sheetName, SOUBA_PDF.range)
    .setName(`貴金属相場表_${fmt_(now, 'yyyyMMdd')}.pdf`);

  const subject = `【相場表】${fmt_(now, 'M/d')} 貴金属相場表` + (isTodayRate ? '' : '(本日分未取得)');
  const body = isTodayRate
    ? `本日 ${fmt_(now, 'M/d')} の貴金属相場表です。\n\n${ss.getUrl()}`
    : `本日分の相場がまだ取得できていないため、前回(${soubaDate_(ss) || '不明'})の相場表を送ります。\n\n${ss.getUrl()}`;
  GmailApp.sendEmail(SOUBA_PDF.to, subject, body, { attachments: [pdf] });
  props.setProperty(SOUBA_PDF_SENT_KEY, today);
}

function exportSheetPdf_(ss, sheetName, range) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`シート「${sheetName}」がありません`);
  const params = {
    format: 'pdf',
    gid: sheet.getSheetId(),
    range: range,
    size: 'A4',
    portrait: 'true',
    fitw: 'true',
    gridlines: 'false',
    printtitle: 'false',
    sheetnames: 'false',
    pagenum: 'UNDEFINED',
    horizontal_alignment: 'CENTER',
    top_margin: '0.4',
    bottom_margin: '0.4',
    left_margin: '0.4',
    right_margin: '0.4',
  };
  const query = Object.keys(params).map(k => `${k}=${encodeURIComponent(params[k])}`).join('&');
  const url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?${query}`;
  const res = UrlFetchApp.fetch(url, {
    headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` },
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200) {
    throw new Error(`PDF の書き出しに失敗しました (${res.getResponseCode()})`);
  }
  return res.getBlob();
}

/** 設定シートの「日付」を yyyy/MM/dd で返す。 */
function soubaDate_(ss) {
  const sheet = ss.getSheetByName(SOUBA_PDF.settingsSheetName);
  if (!sheet) return '';
  const v = sheet.getRange(SOUBA_PDF.dateCell).getValue();
  if (v instanceof Date) return fmt_(v, 'yyyy/MM/dd');
  return String(v).trim().replace(/-/g, '/');
}

function deleteSendTriggers_() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === SOUBA_PDF_SEND)
    .forEach(t => ScriptApp.deleteTrigger(t));
}

function isWeekday_(d) {
  const day = Number(fmt_(d, 'u')); // 1=月 ... 7=日
  return day <= 5;
}

function todayAt_(h, m) {
  const d = new Date();
  d.setHours(h, m, 0, 0);
  return d;
}

function fmt_(d, pattern) {
  return Utilities.formatDate(d, 'Asia/Tokyo', pattern);
}
