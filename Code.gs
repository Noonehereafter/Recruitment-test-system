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
    .addItem('Mở Admin Panel', 'showAdminSidebar')
    .addItem('Xuất kết quả ứng viên', 'showExportPrompt')
    .addToUi();
}

function showAdminSidebar() {
  const html = HtmlService.createTemplateFromFile('AdminSidebar').evaluate().setTitle('Admin Panel');
  SpreadsheetApp.getUi().showSidebar(html);
}

function showExportPrompt() {
  const html = HtmlService.createHtmlOutputFromFile('ExportModal')
      .setWidth(600)
      .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Xuất kết quả bài Test');
}

function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

function assignQuestionsToCandidate(email, position, idsString) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('ASSIGNMENTS');
  if (!sheet) return "Lỗi: Không tìm thấy sheet ASSIGNMENTS.";

  // Parse space-separated IDs to comma-separated
  let ids = idsString.split(/\s+/).map(id => id.trim()).filter(id => id.length > 0).join(',');

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email) {
      sheet.getRange(i + 1, 2).setValue(position);
      sheet.getRange(i + 1, 3).setValue(ids);
      return "Đã cập nhật chỉ định câu hỏi thành công!";
    }
  }

  sheet.appendRow([email, position, ids]);
  return "Đã thêm chỉ định câu hỏi thành công!";
}

function getCandidateReport(email) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const qbData = ss.getSheetByName('QUESTIONBANK').getDataRange().getValues();
  const ansData = ss.getSheetByName('ANSWERS').getDataRange().getValues();
  const resData = ss.getSheetByName('RESULTS').getDataRange().getValues();

  const cRes = resData.find(r => r[2] === email);
  if (!cRes) return null;

  const candidateAnswers = ansData.filter(r => r[1] === email);

  let catScores = {};
  let paragraphAnswers = [];

  candidateAnswers.forEach(ans => {
    const qID = ans[2];
    const pts = Number(ans[6]) || 0;
    const qData = qbData.find(q => q[0] === qID);

    if (qData) {
      let cat = qData[1];
      if (!catScores[cat]) catScores[cat] = 0;
      catScores[cat] += pts;

      if (qData[9] === 'PARAGRAPH') {
        paragraphAnswers.push({ id: qID, question: qData[6], answer: ans[4], currentScore: pts });
      }
    }
  });

  return {
    email: email,
    totalScore: cRes[5],
    timeTaken: cRes[7],
    discProfile: cRes[12],
    catScores: catScores,
    paragraphAnswers: paragraphAnswers
  };
}

function updateParagraphScore(email, qId, newScore) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ansSheet = ss.getSheetByName('ANSWERS');
  const resSheet = ss.getSheetByName('RESULTS');

  const ansData = ansSheet.getDataRange().getValues();
  let scoreDiff = 0;

  for (let i = 1; i < ansData.length; i++) {
    if (ansData[i][1] === email && ansData[i][2] === qId) {
      const oldScore = Number(ansData[i][6]) || 0;
      scoreDiff = Number(newScore) - oldScore;
      ansSheet.getRange(i + 1, 7).setValue(Number(newScore));
      break;
    }
  }

  if (scoreDiff !== 0) {
    const resData = resSheet.getDataRange().getValues();
    for (let i = 1; i < resData.length; i++) {
      if (resData[i][2] === email) {
        const currentTotal = Number(resData[i][5]) || 0;
        resSheet.getRange(i + 1, 6).setValue(currentTotal + scoreDiff);
        break;
      }
    }
  }
  return true;
}

function getConfigData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName('CONFIG');
  if (!configSheet) return { headers: [], data: [] };

  const values = configSheet.getDataRange().getValues();
  if (values.length < 1) return { headers: [], data: [] };

  return {
    headers: values[0],
    data: values.slice(1)
  };
}

function updateConfigData(configArray) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName('CONFIG');
  if (!configSheet) return "Lỗi: Không tìm thấy sheet CONFIG.";

  // configArray is expected to be a 2D array including headers
  configSheet.clearContents();
  configSheet.getRange(1, 1, configArray.length, configArray[0].length).setValues(configArray);
  configSheet.getRange(1, 1, 1, configArray[0].length).setFontWeight('bold');
  return "Cập nhật Configs thành công!";
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
      if (qData[1] !== 'Personality' && qData[9] !== 'PARAGRAPH') {
        output += `- Đáp án đúng của hệ thống: ${qData[17]}\n`;
        output += `- Kết quả chấm tự động: ${isCorrect ? 'ĐÚNG' : 'SAI'}\n`;
      }
      output += `\n`;
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
    resSheet.getRange(1, 1, 1, 13).setValues([['Timestamp', 'FullName', 'Email', 'Phone', 'Position', 'Total Score', 'DISC/MBTI Raw Data', 'Time Taken (s)', 'DISC_D', 'DISC_I', 'DISC_S', 'DISC_C', 'DISC_Profile']]).setFontWeight('bold');
  }

  if (!ss.getSheetByName('AUTOSAVES')) {
    let autoSheet = ss.insertSheet('AUTOSAVES');
    autoSheet.getRange(1, 1, 1, 3).setValues([['Timestamp', 'Email', 'JSON Data']]).setFontWeight('bold');
  }

  if (!ss.getSheetByName('ASSIGNMENTS')) {
    let assSheet = ss.insertSheet('ASSIGNMENTS');
    assSheet.getRange(1, 1, 1, 3).setValues([['Email', 'Position', 'Assigned Question IDs (Comma separated)']]).setFontWeight('bold');
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
  const configData = ss.getSheetByName('CONFIG').getDataRange().getValues();
  const configHeaders = configData[0];
  const config = configData.find(r => r[0] === candidateInfo.position);
  if (!config) throw new Error("Không tìm thấy cấu hình.");

  const allQ = ss.getSheetByName('QUESTIONBANK').getDataRange().getValues();
  const headers = allQ[0];
  let testQuestions = [];

  // Check if assigned specifically
  let assignmentsSheet = ss.getSheetByName('ASSIGNMENTS');
  let specificAssignment = null;
  if (assignmentsSheet) {
    const assData = assignmentsSheet.getDataRange().getValues();
    specificAssignment = assData.find(r => r[0] === candidateInfo.email);
  }

  if (specificAssignment && specificAssignment[2]) {
    const assignedIds = specificAssignment[2].toString().split(',').map(s => s.trim());
    testQuestions = allQ.slice(1).filter(q => assignedIds.includes(q[0].toString()));
  } else {
    // Regular random logic
    const diffCounts = {};
    for (let i = 1; i < configHeaders.length; i++) {
      const h = configHeaders[i];
      if (h.endsWith('_Count')) {
        const parts = h.replace('_Count', '').split('_');
        let cat = parts[0];
        let diff = "ANY";
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

    const activeQ = allQ.slice(1).filter(q => {
      return q[headers.indexOf('Status')] === 'Active' &&
             (q[headers.indexOf('PositionLevel')] === 'All' || q[headers.indexOf('PositionLevel')] === candidateInfo.position);
    });

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
  let discCounts = {D: 0, I: 0, S: 0, C: 0};

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

    if (q[1] === 'Personality') {
       if (candidateAns === String(q[11]).trim()) discCounts.D++;
       else if (candidateAns === String(q[12]).trim()) discCounts.I++;
       else if (candidateAns === String(q[13]).trim()) discCounts.S++;
       else if (candidateAns === String(q[14]).trim()) discCounts.C++;
    }

    if (q[1] === 'Personality' || q[9] === 'PARAGRAPH' || (Number(q[16]) || 0) === 0) {
      rawAI.push(`Q: ${ans.question} | A: ${ans.answer}`);
    }

    ss.getSheetByName('ANSWERS').appendRow([ts, submission.candidateInfo.email, ans.id, ans.question, ans.answer, correct, pts]);
  });

  // Calculate DISC Profile
  let discArray = [
    {type: 'D', count: discCounts.D},
    {type: 'I', count: discCounts.I},
    {type: 'S', count: discCounts.S},
    {type: 'C', count: discCounts.C}
  ];
  discArray.sort((a, b) => b.count - a.count);
  let discProfile = `${discArray[0].type}-${discArray[1].type}`;

  ss.getSheetByName('RESULTS').appendRow([ts, submission.candidateInfo.fullName, submission.candidateInfo.email, submission.candidateInfo.phone, submission.candidateInfo.position, score, rawAI.join('\n'), submission.timeTaken || 0, discCounts.D, discCounts.I, discCounts.S, discCounts.C, discProfile]);
  return "Nộp bài thành công!";
}

function saveProgress(email, data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('AUTOSAVES');
  if (!sheet) return;
  const ts = new Date();

  // Try to find existing row
  const rData = sheet.getDataRange().getValues();
  for (let i = 1; i < rData.length; i++) {
    if (rData[i][1] === email) {
      sheet.getRange(i + 1, 1).setValue(ts);
      sheet.getRange(i + 1, 3).setValue(JSON.stringify(data));
      return;
    }
  }

  // If not found, append
  sheet.appendRow([ts, email, JSON.stringify(data)]);
}
