/**
 * Google Apps Script for Recruitment Test System
 */

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Hệ thống Test Tuyển dụng')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Hệ thống Test')
    .addItem('Khởi tạo hệ thống', 'initSystem')
    .addToUi();
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function initSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let qbSheet = ss.getSheetByName('QUESTIONBANK');
  if (!qbSheet) {
    qbSheet = ss.insertSheet('QUESTIONBANK');
    const headers = [
      'ID', 'Category', 'PositionLevel', 'Difficulty', 'Version', 'Status', 
      'Question', 'Desc', 'Image', 'Type', 'Required', 
      'Option 1', 'Option 2', 'Option 3', 'Option 4', 'Other', 
      'Points', 'Correct Answer', 'Correct Feedback', 'Correct URL', 
      'Incorrect Feedback', 'Incorrect URL'
    ];
    qbSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    setupValidation(qbSheet);
    addSampleQuestions(qbSheet);
  }

  let configSheet = ss.getSheetByName('CONFIG');
  if (!configSheet) {
    configSheet = ss.insertSheet('CONFIG');
    const headers = ['Position', 'IQ_Count', 'EQ_Count', 'ProblemSolving_Count', 'Leadership_Count', 'Personality_Count', 'Duration_Minutes'];
    configSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    configSheet.appendRow(['Staff', 5, 5, 5, 5, 5, 30]);
  }

  if (!ss.getSheetByName('ANSWERS')) {
    let ansSheet = ss.insertSheet('ANSWERS');
    ansSheet.getRange(1, 1, 1, 7).setValues([['Timestamp', 'Candidate Email', 'Question ID', 'Question Text', 'Candidate Answer', 'Is Correct', 'Points']]).setFontWeight('bold');
  }

  if (!ss.getSheetByName('RESULTS')) {
    let resSheet = ss.insertSheet('RESULTS');
    resSheet.getRange(1, 1, 1, 7).setValues([['Timestamp', 'FullName', 'Email', 'Phone', 'Position', 'Total Score', 'DISC/MBTI Raw Data']]).setFontWeight('bold');
  }
  
  return "Hệ thống đã khởi tạo thành công!";
}

function setupValidation(sheet) {
  const rules = {
    'B': ['Personality', 'IQ', 'EQ', 'ProblemSolving', 'Leadership'],
    'C': ['All', 'Staff', 'Manager', 'Senior'],
    'F': ['Active', 'Inactive', 'Review'],
    'J': ['MULTIPLECHOICE', 'CHECKBOX', 'PARAGRAPH', 'SCALE', 'DROPDOWN'],
    'K': ['TRUE', 'FALSE']
  };
  for (let col in rules) {
    let range = sheet.getRange(col + "2:" + col + "1000");
    let rule = SpreadsheetApp.newDataValidation().requireValueInList(rules[col], true).build();
    range.setDataValidation(rule);
  }
}

function addSampleQuestions(sheet) {
  const samples = [
    ['PS001', 'ProblemSolving', 'Manager', 2, 1, 'Active', 'Câu 1: Dự án trễ tiến độ, bạn xử lý thế nào?', '', '', 'MULTIPLECHOICE', 'TRUE', 'Báo cáo', 'Đánh giá nội bộ', 'Thuê thêm', 'Báo KH', '', 1, 'Đánh giá nội bộ', 'Đúng!', '', 'Sai!', ''],
    ['PER001', 'Personality', 'All', 1, 1, 'Active', 'Câu 1: Khi có xung đột nhóm, bạn sẽ?', '', '', 'MULTIPLECHOICE', 'TRUE', 'Quyết định ngay', 'Thảo luận', 'Giữ hòa khí', 'Phân tích', '', 0, '', 'D, I, S, C mapping...', '', '', ''],
    ['IQ001', 'IQ', 'All', 1, 1, 'Active', 'Câu 1: 2, 4, 8, 16... số tiếp theo?', '', '', 'MULTIPLECHOICE', 'TRUE', '24', '30', '32', '64', '', 1, '32', 'Đúng!', '', 'Sai!', '']
  ];
  sheet.getRange(2, 1, samples.length, samples[0].length).setValues(samples);
}

function getPositions() {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CONFIG').getDataRange().getValues();
  return data.slice(1).map(r => r[0]);
}

function generateTest(candidateInfo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = ss.getSheetByName('CONFIG').getDataRange().getValues().find(r => r[0] === candidateInfo.position);
  if (!config) throw new Error("Không tìm thấy cấu hình.");

  const counts = { IQ: config[1], EQ: config[2], ProblemSolving: config[3], Leadership: config[4], Personality: config[5] };
  const allQ = ss.getSheetByName('QUESTIONBANK').getDataRange().getValues();
  const headers = allQ[0];
  const activeQ = allQ.slice(1).filter(q => q[headers.indexOf('Status')] === 'Active');

  const testQuestions = [];
  for (let cat in counts) {
    let pool = activeQ.filter(q => q[headers.indexOf('Category')] === cat).sort(() => Math.random() - 0.5);
    testQuestions.push(...pool.slice(0, counts[cat]));
  }

  return {
    questions: testQuestions.map(q => ({
      id: q[0], question: q[6], desc: q[7], image: q[8], type: q[9], required: q[10],
      options: [q[11], q[12], q[13], q[14], q[15]].filter(o => o !== ""), points: q[16]
    })),
    duration: config[6],
    candidateInfo: candidateInfo
  };
}

function processResults(submission) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const qb = ss.getSheetByName('QUESTIONBANK').getDataRange().getValues();
  const ts = new Date();
  let score = 0, rawAI = [];

  submission.answers.forEach(ans => {
    const q = qb.find(r => r[0] === ans.id);
    if (!q) return;
    let correct = String(ans.answer).trim() === String(q[17]).trim();
    let pts = correct ? (Number(q[16]) || 0) : 0;
    score += pts;
    if (q[1] === 'Personality' || q[9] === 'PARAGRAPH' || q[1] === 'Leadership') {
      rawAI.push(`Q: ${ans.question} | A: ${ans.answer}`);
    }
    ss.getSheetByName('ANSWERS').appendRow([ts, submission.candidateInfo.email, ans.id, ans.question, ans.answer, correct, pts]);
  });

  ss.getSheetByName('RESULTS').appendRow([ts, submission.candidateInfo.fullName, submission.candidateInfo.email, submission.candidateInfo.phone, submission.candidateInfo.position, score, rawAI.join('\n')]);
  return "Nộp bài thành công!";
}
