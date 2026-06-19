/**
 * AI Safety Agent Survey — response collection webhook
 * Supports:
 * - randomized situation block order
 * - randomized 12-agent order within each situation
 * - training + warning situation responses
 *
 * 이 파일은 Google Apps Script(스프레드시트에 연결된 웹앱)의 사본입니다.
 * 실제 동작 코드는 Apps Script 편집기에 있으며, 여기서는 버전 관리용으로 보관합니다.
 * 코드를 바꾸면 편집기에 붙여넣고 Deploy → Manage deployments → New version 으로 재배포하세요.
 */
const SHEET_NAME = 'Responses';

const BASE_HEADERS = [
  'timestamp',
  'prolificPid',
  'studyId',
  'sessionId',

  // IRB 동의 — 값은 'agree' / 'disagree' (이름 자체가 아니라 '동의 여부'를 저장).
  // 셋 다 agree여야 설문이 진행되므로 저장되는 값은 항상 agree. 핵심 증빙은 consentTimestamp.
  'consentName',          // 7번: 이름 수집·사용에 동의했는가
  'consentFutureUse',     // 8번: 향후 연구·논문·학회 활용에 동의했는가
  'consentParticipate',   // 최종: 연구 참여에 동의했는가
  'consentTimestamp',     // 동의(Continue) 클릭 시각 (ISO 문자열)

  'firstName',
  'lastName',
  'age',
  'participantGender',
  'country',
  'education',
  'yearsOnSite',

  'riskInjury',
  'riskOwnWork',
  'riskOwnAccident',
  'aiUsage',
  'aiFamiliarity',
  'aiFamiliarityConstruction',
  'aiTrust',

  'situationOrder',
  'trainingOrder',
  'warningOrder'
];

const AGENT_KEYS = [
  'HU-MA-HW', 'HU-MA-LW',
  'HU-FE-HW', 'HU-FE-LW',
  'HM-MA-HW', 'HM-MA-LW',
  'HM-FE-HW', 'HM-FE-LW',
  'RB-MA-HW', 'RB-MA-LW',
  'RB-FE-HW', 'RB-FE-LW'
];

const SITUATIONS = ['training', 'warning'];

const EVAL_ITEMS = [
  'compliance',
  'preference',
  'persuasiveness',
  'trust',
  'comfort',
  'attention'
];

const IGNORED_KEYS = [
  'scenarioOrder'
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonOut({ result: 'error', message: 'no postData' });
    }

    const data = JSON.parse(e.postData.contents);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

    let headers = getExistingHeaders(sheet);
    const preferredHeaders = buildPreferredHeaders();
    const payloadHeaders = Object.keys(data).filter(function (k) {
      return IGNORED_KEYS.indexOf(k) === -1;
    });

    headers = mergeHeaders(headers, preferredHeaders);
    headers = mergeHeaders(headers, payloadHeaders);

    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    const row = headers.map(function (h) {
      const v = data[h];
      return (v === undefined || v === null) ? '' : v;
    });

    sheet.appendRow(row);

    return jsonOut({
      result: 'success',
      row: sheet.getLastRow()
    });

  } catch (err) {
    return jsonOut({
      result: 'error',
      message: String(err)
    });

  } finally {
    lock.releaseLock();
  }
}

function buildPreferredHeaders() {
  const headers = BASE_HEADERS.slice();

  SITUATIONS.forEach(function (situation) {
    AGENT_KEYS.forEach(function (agentKey) {
      EVAL_ITEMS.forEach(function (item) {
        headers.push(situation + '_' + agentKey + '_' + item);
      });
    });
  });

  return headers;
}

function getExistingHeaders(sheet) {
  if (sheet.getLastRow() === 0 || sheet.getLastColumn() === 0) {
    return [];
  }

  return sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0]
    .filter(function (h) {
      return h !== '';
    });
}

function mergeHeaders(existingHeaders, newHeaders) {
  const merged = existingHeaders.slice();

  newHeaders.forEach(function (h) {
    if (merged.indexOf(h) === -1) {
      merged.push(h);
    }
  });

  return merged;
}

function doGet() {
  return jsonOut({
    result: 'ok',
    message: 'AI Safety Agent survey webhook is running.'
  });
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
