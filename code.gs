// Classroom Gesture Quiz - code.gs
// Deploy as Web App: Execute as Me, Anyone can access

var SS_NAME = "Student Walking Quiz";
var VOCAB = "Vocabulary";
var RESULTS = "Results";

function doGet(e) {
  var action = e && e.parameter && e.parameter.action;

  if (action === 'questions') {
    return respond(getQuestionsList());
  }

  // Default: return info page
  var html = '<h2>Gesture Quiz API running</h2><p>Use ?action=questions to get questions.</p>';
  return HtmlService.createHtmlOutput(html);
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    if (data.action === 'save') {
      var ss = getSheet();
      var sheet = ss.getSheetByName(RESULTS);
      var score = (data.answer === data.correct_answer) ? 1 : 0;
      sheet.appendRow([new Date().toISOString(), data.question_id, data.answer, score]);
      return respond({ ok: true, score: score });
    }
  } catch (ex) {
    return respond({ ok: false, error: ex.toString() });
  }
  return respond({ ok: false, error: 'Unknown action' });
}

function respond(obj) {
  var output = ContentService.createTextOutput(JSON.stringify(obj));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

function getQuestionsList() {
  try {
    var ss = getSheet();
    var sheet = ss.getSheetByName(VOCAB);
    var data = sheet.getDataRange().getValues();
    var list = [];
    for (var i = 1; i < data.length; i++) {
      var r = data[i];
      if (r[0] !== '' && r[1] !== '') {
        list.push({
          id: String(r[0]),
          question: String(r[1]),
          type: String(r[2] || 'text'),
          correct_answer: String(r[3]).toUpperCase().trim()
        });
      }
    }
    return { ok: true, list: list };
  } catch (ex) {
    return { ok: false, list: [] };
  }
}

function getSheet() {
  var files = DriveApp.getFilesByName(SS_NAME);
  if (files.hasNext()) {
    return SpreadsheetApp.open(files.next());
  }
  var ss = SpreadsheetApp.create(SS_NAME);
  var v = ss.getActiveSheet();
  v.setName(VOCAB);
  v.getRange('A1:D1').setValues([['id','question','type','correct_answer']]);
  v.getRange('A1:D1').setFontWeight('bold').setBackground('#4a90d9').setFontColor('#ffffff');
  var d = [
    [1,'https://upload.wikimedia.org/wikipedia/commons/thumb/1/15/Red_Apple.jpg/800px-Red_Apple.jpg','image','LEFT'],
    [2,'Elephant','text','RIGHT'],
    [3,'https://upload.wikimedia.org/wikipedia/commons/thumb/3/3a/Cat03.jpg/481px-Cat03.jpg','image','RIGHT'],
    [4,'Sun','text','LEFT'],
    [5,'Dog','text','RIGHT'],
    [6,'Moon','text','LEFT'],
    [7,'Ocean','text','RIGHT'],
    [8,'Tree','text','LEFT'],
    [9,'Car','text','RIGHT'],
    [10,'Bird','text','LEFT']
  ];
  v.getRange(2, 1, d.length, 4).setValues(d);
  var r = ss.insertSheet(RESULTS);
  r.getRange('A1:D1').setValues([['timestamp','question_id','answer','score']]);
  r.getRange('A1:D1').setFontWeight('bold').setBackground('#4a90d9').setFontColor('#ffffff');
  return ss;
}

function initSetup() {
  getSheet();
  Logger.log('Done');
}

function insertSampleQuestions() {
  var ss = getSheet();
  var sheet = ss.getSheetByName(VOCAB);
  var last = sheet.getLastRow();
  if (last > 1) sheet.getRange(2, 1, last - 1, 4).clearContent();
  var d = [
    [1,'https://upload.wikimedia.org/wikipedia/commons/thumb/1/15/Red_Apple.jpg/800px-Red_Apple.jpg','image','LEFT'],
    [2,'Elephant','text','RIGHT'],
    [3,'https://upload.wikimedia.org/wikipedia/commons/thumb/3/3a/Cat03.jpg/481px-Cat03.jpg','image','RIGHT'],
    [4,'Sun','text','LEFT'],
    [5,'Dog','text','RIGHT'],
    [6,'Moon','text','LEFT'],
    [7,'Ocean','text','RIGHT'],
    [8,'Tree','text','LEFT'],
    [9,'Car','text','RIGHT'],
    [10,'Bird','text','LEFT']
  ];
  sheet.getRange(2, 1, d.length, 4).setValues(d);
  Logger.log('Inserted ' + d.length + ' questions');
}
