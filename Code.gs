// ═══════════════════════════════════════════════════════════
//  IELTS CITY — OTR TOOL BCD  v3
//  Google Apps Script — Code.gs
//  Deploy as: Web App | Execute as: User accessing the web app
//                      | Who has access: Anyone (IELTS City)
// ═══════════════════════════════════════════════════════════

const VERCEL_URL = 'https://icacaotr.vercel.app/api/generate-otr';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('OTR Tool BCD — IELTS City')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function generateOTR(sheetUrl) {
  try {
    const sheetId = extractSheetId(sheetUrl);
    if (!sheetId) throw new Error('URL không hợp lệ. Vui lòng dán đúng link Google Sheets.');

    const ss = SpreadsheetApp.openById(sheetId);

    const ovSheet = getSheet(ss, 'Tổng quan');
    const sSheet  = getSheet(ss, 'S');
    const w1Sheet = getSheet(ss, 'W1');
    const w2Sheet = getSheet(ss, 'W2');
    const lSheet  = getSheet(ss, 'L');
    const rSheet  = getSheet(ss, 'R');

    const candidateName    = str(ovSheet, 'C4');
    const rawDate          = ovSheet.getRange('C5').getValue();
    const testDate         = formatDate(rawDate);
    const photoRawUrl      = str(ovSheet, 'C8');
    const overallScore     = ovSheet.getRange('F5').getValue();
    const listeningCorrect = ovSheet.getRange('C10').getValue();
    const listeningScore   = ovSheet.getRange('F10').getValue();
    const readingCorrect   = ovSheet.getRange('C13').getValue();
    const readingScore     = ovSheet.getRange('F13').getValue();
    const writingScore     = ovSheet.getRange('F17').getValue();
    const speakingScore    = sSheet.getRange('E4').getValue();
    const w1Score          = w1Sheet.getRange('E4').getValue();
    const w2Score          = w2Sheet.getRange('E4').getValue();

    const fileName   = ss.getName();
    const classMatch = fileName.match(/([A-Z]{1,3}\d{3,4})/);
    const classId    = classMatch ? classMatch[1] : 'N/A';

    const photo = fetchPhoto(photoRawUrl);

    const speaking = {
      criteria: [
        { name: str(sSheet,'B6'),  score: str(sSheet,'E6'),
          rows: [row(sSheet,8), row(sSheet,9), row(sSheet,10)] },
        { name: str(sSheet,'B12'), score: str(sSheet,'E12'),
          rows: [row(sSheet,14), row(sSheet,15), row(sSheet,16)] },
        { name: str(sSheet,'B18'), score: str(sSheet,'E18'),
          rows: [row(sSheet,20), row(sSheet,21)] },
        { name: str(sSheet,'B23'), score: str(sSheet,'E23'),
          rows: [row(sSheet,25), row(sSheet,26), row(sSheet,27)] }
      ]
    };

    const w1 = {
      criteria: [
        { name: str(w1Sheet,'B6'),  score: str(w1Sheet,'E6'),
          rows: [row(w1Sheet,8), row(w1Sheet,9), row(w1Sheet,10)] },
        { name: str(w1Sheet,'B12'), score: str(w1Sheet,'E12'),
          rows: [row(w1Sheet,14), row(w1Sheet,15), row(w1Sheet,16)] },
        { name: str(w1Sheet,'B18'), score: str(w1Sheet,'E18'),
          rows: [row(w1Sheet,20), row(w1Sheet,21)] },
        { name: str(w1Sheet,'B23'), score: str(w1Sheet,'E23'),
          rows: [row(w1Sheet,25), row(w1Sheet,26)] }
      ]
    };

    const w2 = {
      criteria: [
        { name: str(w2Sheet,'B6'),  score: str(w2Sheet,'E6'),
          rows: [row(w2Sheet,8), row(w2Sheet,9), row(w2Sheet,10)] },
        { name: str(w2Sheet,'B12'), score: str(w2Sheet,'E12'),
          rows: [row(w2Sheet,14), row(w2Sheet,15), row(w2Sheet,16)] },
        { name: str(w2Sheet,'B18'), score: str(w2Sheet,'E18'),
          rows: [row(w2Sheet,20), row(w2Sheet,21)] },
        { name: str(w2Sheet,'B23'), score: str(w2Sheet,'E23'),
          rows: [row(w2Sheet,25), row(w2Sheet,26)] }
      ]
    };

    const listeningWrong = extractWrong(lSheet);
    const readingWrong   = extractWrong(rSheet);

    const payload = {
      candidate: { name: candidateName, testDate, classId, overall: overallScore },
      scores: {
        listening: listeningScore, reading: readingScore,
        writing: writingScore,     speaking: speakingScore,
        w1Score, w2Score,
        listeningCorrect, readingCorrect
      },
      photo,
      speaking, w1, w2, listeningWrong, readingWrong
    };

    const resp = UrlFetchApp.fetch(VERCEL_URL, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    if (resp.getResponseCode() !== 200) {
      throw new Error('Lỗi từ server: ' + resp.getContentText().substring(0, 300));
    }

    const b64    = Utilities.base64Encode(resp.getContent());
    const dlName = 'OTR_' + candidateName.replace(/\s+/g,'_') + '_' + classId + '.docx';
    return { ok: true, file: b64, name: dlName };

  } catch(e) {
    return { ok: false, error: e.message };
  }
}

function getSheet(ss, name) {
  const ws = ss.getSheetByName(name);
  if (!ws) throw new Error('Không tìm thấy sheet: "' + name + '"');
  return ws;
}

function str(sheet, cell) {
  const v = sheet.getRange(cell).getValue();
  return (v !== null && v !== undefined) ? String(v) : '';
}

function row(sheet, rowNum) {
  const vals = sheet.getRange(rowNum, 1, 1, 5).getValues()[0];
  return [
    String(vals[1] || ''),
    String(vals[2] || ''),
    String(vals[3] || ''),
    String(vals[4] || '')
  ];
}

function extractWrong(sheet) {
  const data = sheet.getRange('B8:E47').getValues();
  return data
    .map((r,i) => ({ q: i+1, status: r[3] }))
    .filter(x => x.status === false || String(x.status).toLowerCase() === 'false')
    .map(x => x.q);
}

function fetchPhoto(rawUrl) {
  if (!rawUrl || !rawUrl.trim()) return null;
  const url = rawUrl.trim();

  const driveMatch = url.match(/(?:\/file\/d\/|[?&]id=)([a-zA-Z0-9_-]{10,})/);
  if (driveMatch) {
    const fileId = driveMatch[1];

    try {
      const blob = DriveApp.getFileById(fileId).getBlob();
      const ct   = blob.getContentType().toLowerCase();
      const type = ct.includes('png') ? 'png' : 'jpg';
      return { b64: Utilities.base64Encode(blob.getBytes()), type };
    } catch(e) {}

    try {
      const dlUrl = 'https://drive.google.com/uc?export=download&id=' + fileId;
      const resp  = UrlFetchApp.fetch(dlUrl, {
        muteHttpExceptions: true,
        followRedirects: true,
        headers: { 'User-Agent': 'Mozilla/5.0' }
      });
      if (resp.getResponseCode() === 200) {
        const ct = (resp.getHeaders()['Content-Type'] || '').toLowerCase();
        if (ct.startsWith('image/')) {
          return { b64: Utilities.base64Encode(resp.getContent()), type: ct.includes('png') ? 'png' : 'jpg' };
        }
        const resp2 = UrlFetchApp.fetch(dlUrl + '&confirm=t', { muteHttpExceptions: true, followRedirects: true });
        if (resp2.getResponseCode() === 200) {
          const ct2 = (resp2.getHeaders()['Content-Type'] || '').toLowerCase();
          if (ct2.startsWith('image/')) {
            return { b64: Utilities.base64Encode(resp2.getContent()), type: ct2.includes('png') ? 'png' : 'jpg' };
          }
        }
      }
    } catch(e) {}

    return null;
  }

  try {
    const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    if (resp.getResponseCode() !== 200) return null;
    const ct = (resp.getHeaders()['Content-Type'] || '').toLowerCase();
    if (!ct.startsWith('image/')) return null;
    return { b64: Utilities.base64Encode(resp.getContent()), type: ct.includes('png') ? 'png' : 'jpg' };
  } catch(e) {
    return null;
  }
}

function formatDate(raw) {
  if (!raw) return '';
  if (raw instanceof Date) {
    const dd = String(raw.getDate()).padStart(2,'0');
    const mm = String(raw.getMonth()+1).padStart(2,'0');
    return dd+'/'+mm+'/'+raw.getFullYear();
  }
  return String(raw);
}

function extractSheetId(url) {
  const m = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  return m ? m[1] : null;
}
