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
    .addItem('Xuất kết quả ứng viên', 'showExportPrompt')
    .addToUi();
}

function showExportPrompt() {
  const html = HtmlService.createHtmlOutputFromFile('ExportModal')
      .setWidth(600)
      .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Xuất kết quả bài Test');
}

function generateExportData(email) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const qbData = ss.getSheetByName('QUESTIONBANK').getDataRange().getValues();
  const ansData = ss.getSheetByName('ANSWERS').getDataRange().getValues();

  // Filter answers by email
  const candidateAnswers = ansData.filter(r => r[1] === email);
  if (candidateAnswers.length === 0) return "Không tìm thấy dữ liệu cho email này.";

  let output = `KẾT QUẢ BÀI TEST - Ứng viên: ${email}\n==============================================\n\n`;

  candidateAnswers.forEach(ans => {
    const qID = ans[2];
    const candidateAnswer = ans[4];
    const isCorrect = ans[5];
    const qData = qbData.find(q => q[0] === qID);

    if (qData) {
      output += `[${qData[1]}] Câu hỏi: ${qData[6]}\n`;
      output += `- Ứng viên chọn: ${candidateAnswer}\n`;
      output += `- Đáp án đúng của hệ thống: ${qData[17]}\n`;
      output += `- Kết quả chấm tự động: ${isCorrect ? 'ĐÚNG' : 'SAI'}\n\n`;
    }
  });

  return output;
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
    const headers = ['Position', 'IQ_1_Count', 'IQ_2_Count', 'IQ_3_Count', 'EQ_1_Count', 'EQ_2_Count', 'EQ_3_Count', 'Problem_Solving_1_Count', 'Problem_Solving_2_Count', 'Problem_Solving_3_Count', 'Leadership_1_Count', 'Leadership_2_Count', 'Leadership_3_Count', 'Personality_1_Count', 'Duration_Minutes'];
    configSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    configSheet.appendRow(['Staff', 2, 2, 1, 2, 2, 1, 2, 2, 1, 2, 2, 1, 5, 30]);
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
    'B': ['Personality', 'IQ', 'EQ', 'Problem_Solving', 'Leadership'],
    'C': ['All', 'Staff', 'Manager', 'Senior'],
    'F': ['Active', 'Inactive', 'Review'],
    'J': ['MULTIPLE_CHOICE', 'CHECKBOX', 'PARAGRAPH', 'SCALE', 'DROPDOWN'],
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
    ['PS001', 'Problem_Solving', 'Manager', 2, 1, 'Active', 'Câu 1: Dự án trễ tiến độ, bạn xử lý thế nào?', '', '', 'MULTIPLE_CHOICE', 'TRUE', 'Báo cáo', 'Đánh giá nội bộ', 'Thuê thêm', 'Báo KH', '', 1, 'Đánh giá nội bộ', 'Đúng!', '', 'Sai!', ''],
    ['PER001', 'Personality', 'All', 1, 1, 'Active', 'Câu 1: Khi có xung đột nhóm, bạn sẽ?', '', '', 'MULTIPLE_CHOICE', 'TRUE', 'Quyết định ngay', 'Thảo luận', 'Giữ hòa khí', 'Phân tích', '', 0, '', 'D, I, S, C mapping...', '', '', ''],
    ['IQ001', 'IQ', 'All', 1, 1, 'Active', 'Câu 1: 2, 4, 8, 16... số tiếp theo?', '', '', 'MULTIPLE_CHOICE', 'TRUE', '24', '30', '32', '64', '', 1, '32', 'Đúng!', '', 'Sai!', '']
  ];
  sheet.getRange(2, 1, samples.length, samples[0].length).setValues(samples);
}

function getPositions() {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CONFIG').getDataRange().getValues();
  return data.slice(1).map(r => r[0]);
}

function generateTest(candidateInfo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configData = ss.getSheetByName('CONFIG').getDataRange().getValues();
  const configHeaders = configData[0];
  const config = configData.find(r => r[0] === candidateInfo.position);
  if (!config) throw new Error("Không tìm thấy cấu hình.");

  const diffCounts = {}; // e.g. { "IQ": { "1": 2, "2": 2, "3": 1 }, "Personality": { "ANY": 5 } }
  for (let i = 1; i < configHeaders.length; i++) {
    const h = configHeaders[i];
    if (h.endsWith('_Count')) {
      const parts = h.replace('_Count', '').split('_');
      let cat = parts[0];
      let diff = "ANY";

      // Handle categories with multiple underscores (like Problem_Solving)
      if (parts.length >= 2 && !isNaN(parts[parts.length - 1])) {
        diff = parts.pop();
        cat = parts.join('_');
      } else if (parts.length > 1) {
        cat = parts.join('_');
      }

      if (!diffCounts[cat]) diffCounts[cat] = {};
      diffCounts[cat][diff] = Number(config[i]) || 0;
    }
  }

  const allQ = ss.getSheetByName('QUESTIONBANK').getDataRange().getValues();
  const headers = allQ[0];
  const activeQ = allQ.slice(1).filter(q => {
    return q[headers.indexOf('Status')] === 'Active' &&
           (q[headers.indexOf('PositionLevel')] === 'All' || q[headers.indexOf('PositionLevel')] === candidateInfo.position);
  });

  const testQuestions = [];
  for (let cat in diffCounts) {
    for (let diff in diffCounts[cat]) {
      let countNeeded = diffCounts[cat][diff];
      if (countNeeded <= 0) continue;

      let pool = activeQ.filter(q => {
        let matchCat = q[headers.indexOf('Category')] === cat;
        if (!matchCat) return false;
        if (diff === "ANY") return true;
        return String(q[headers.indexOf('Difficulty')]) === diff;
      }).sort(() => Math.random() - 0.5);

      testQuestions.push(...pool.slice(0, countNeeded));
    }
  }

  return {
    questions: testQuestions.map(q => ({
      id: q[0], category: q[1], question: q[6], desc: q[7], image: q[8], type: q[9], required: q[10],
      options: [q[11], q[12], q[13], q[14], q[15]].filter(o => o !== ""), points: q[16]
    })),
    duration: config[configHeaders.indexOf('Duration_Minutes')],
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

    let correct = false;
    let candidateAns = String(ans.answer).trim();
    let trueAns = String(q[17]).trim();

    if (q[9] === 'CHECKBOX') {
      let candParts = candidateAns.split(',').map(s => s.trim().toLowerCase()).filter(s => s);
      let trueParts = trueAns.split(',').map(s => s.trim().toLowerCase()).filter(s => s);
      candParts.sort();
      trueParts.sort();
      correct = (candParts.join(',') === trueParts.join(','));
    } else {
      correct = candidateAns.toLowerCase() === trueAns.toLowerCase();
    }

    let pts = correct ? (Number(q[16]) || 0) : 0;
    score += pts;

    if (q[1] === 'Personality' || q[9] === 'PARAGRAPH' || (Number(q[16]) || 0) === 0) {
      rawAI.push(`Q: ${ans.question} | A: ${ans.answer}`);
    }

    ss.getSheetByName('ANSWERS').appendRow([ts, submission.candidateInfo.email, ans.id, ans.question, ans.answer, correct, pts]);
  });

  ss.getSheetByName('RESULTS').appendRow([ts, submission.candidateInfo.fullName, submission.candidateInfo.email, submission.candidateInfo.phone, submission.candidateInfo.position, score, rawAI.join('\n')]);
  return "Nộp bài thành công!";
}
