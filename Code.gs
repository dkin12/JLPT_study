// Google Apps Script — JLPT 단어장 백엔드
// 배포: 확장 프로그램 > Apps Script > 배포 > 웹앱으로 배포
// 실행 계정: 나(본인), 액세스 권한: 모든 사용자

var SHEET_NAME = '단어장';

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['id', 'kanji', 'furigana', 'meaning', 'tags', 'created']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// 모든 요청을 doGet으로 처리 (JSONP) — POST는 CORS 문제로 사용 안 함
function doGet(e) {
  var action = e.parameter.action;
  var callback = e.parameter.callback;
  var sheet = getSheet();
  var result;

  if (action === 'getAll') {
    var rows = sheet.getDataRange().getValues();
    var words = rows.slice(1).map(function(row) {
      return {
        id: row[0], kanji: row[1], furigana: row[2],
        meaning: row[3], tags: row[4] ? row[4].split(',') : [], created: row[5]
      };
    }).filter(function(w) { return w.id; });
    result = { ok: true, words: words };

  } else if (action === 'add') {
    var words = JSON.parse(e.parameter.data);
    words.forEach(function(w) {
      sheet.appendRow([w.id, w.kanji, w.furigana, w.meaning, (w.tags||[]).join(','), new Date().toISOString()]);
    });
    result = { ok: true, added: words.length };

  } else if (action === 'delete') {
    var id = String(e.parameter.id);
    var rows = sheet.getDataRange().getValues();
    for (var i = rows.length - 1; i >= 1; i--) {
      if (String(rows[i][0]) === id) { sheet.deleteRow(i + 1); break; }
    }
    result = { ok: true };

  } else {
    result = { ok: false, error: 'unknown action' };
  }

  var text = JSON.stringify(result);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + text + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(text).setMimeType(ContentService.MimeType.JSON);
}

