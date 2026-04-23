/**
 * Google Apps Script for Recruitment Test System
 */

function doGet(e) {
  if (e && e.parameter && e.parameter.page === 'admin') {
    return HtmlService.createTemplateFromFile('AdminPanel')
      .evaluate()
      .setTitle('Admin Panel - Hệ thống Test')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
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
    .addItem('Mở Admin Panel (Tab mới)', 'showAdminModal')
    .addItem('Xuất kết quả ứng viên', 'showExportPrompt')
    .addToUi();
}

function showAdminModal() {
  const url = ScriptApp.getService().getUrl() + '?page=admin';
  const html = HtmlService.createHtmlOutput(`<script>window.open('${url}', '_blank'); google.script.host.close();</script>`)
    .setWidth(300).setHeight(100);
  SpreadsheetApp.getUi().showModalDialog(html, 'Đang mở Admin Panel...');
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

function assignQuestionsToCandidate(type, value, idsString) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('ASSIGNMENTS');
  if (!sheet) return "Lỗi: Không tìm thấy sheet ASSIGNMENTS.";

  // Parse space-separated IDs to comma-separated
  let ids = idsString.split(/\s+/).map(id => id.trim()).filter(id => id.length > 0).join(',');

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === type && data[i][1] === value) {
      sheet.getRange(i + 1, 3).setValue(ids);
      return "Đã cập nhật chỉ định câu hỏi thành công!";
    }
  }

  sheet.appendRow([type, value, ids]);
  return "Đã thêm chỉ định câu hỏi thành công!";
}

function getSystemSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('SYSTEM_SETTINGS');
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  let settings = {};
  for(let i=1; i<data.length; i++){
    settings[data[i][0]] = data[i][1];
  }
  return settings;
}

function saveSystemSettings(settingsObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('SYSTEM_SETTINGS');
  if (!sheet) return "Lỗi: Không tìm thấy SYSTEM_SETTINGS";
  const data = sheet.getDataRange().getValues();

  for(let key in settingsObj) {
    let found = false;
    for(let i=1; i<data.length; i++){
      if(data[i][0] === key) {
        sheet.getRange(i+1, 2).setValue(settingsObj[key]);
        found = true;
        break;
      }
    }
    if(!found) {
      sheet.appendRow([key, settingsObj[key]]);
    }
  }
  return "Lưu cài đặt thành công!";
}

function getDashboardStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resSheet = ss.getSheetByName('RESULTS');
  if(!resSheet) return { total: 0, avgScore: 0, positionCounts: {} };

  const data = resSheet.getDataRange().getValues();
  if(data.length < 2) return { total: 0, avgScore: 0, positionCounts: {} };

  let totalCandidates = data.length - 1;
  let totalScoreSum = 0;
  let posCounts = {};

  for(let i=1; i<data.length; i++){
    if(!data[i][2]) continue; // skip empty rows

    let score = Number(data[i][5]) || 0;
    totalScoreSum += score;

    let pos = data[i][4] || "Chưa xác định";
    if(!posCounts[pos]) posCounts[pos] = 0;
    posCounts[pos]++;
  }

  return {
    total: totalCandidates,
    avgScore: (totalScoreSum / totalCandidates).toFixed(2),
    positionCounts: posCounts
  };
}

function getCandidatesList() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const resSheet = ss.getSheetByName('RESULTS');
    if(!resSheet) return [];

    const data = resSheet.getDataRange().getValues();
    if(data.length < 2) return []; // Only headers or empty

    let list = [];
    for(let i=1; i<data.length; i++){
      // Skip empty rows
      if(!data[i][2]) continue;

      list.push({
        timestamp: data[i][0] ? new Date(data[i][0]).toISOString() : '',
        fullName: data[i][1] || '',
        email: data[i][2] || '',
        phone: data[i][3] || '',
        position: data[i][4] || '',
        totalScore: data[i][5] || 0
      });
    }
    // Return sorted by newest first
    return list.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
  } catch (error) {
    throw new Error("Lỗi Backend: " + error.message);
  }
}

function getBulkReports(emails) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const qbData = ss.getSheetByName('QUESTIONBANK').getDataRange().getValues();
  const ansData = ss.getSheetByName('ANSWERS').getDataRange().getValues();
  const resData = ss.getSheetByName('RESULTS').getDataRange().getValues();

  let reports = [];

  emails.forEach(email => {
    const cRes = resData.find(r => r[2] === email);
    if (!cRes) return;

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

    reports.push({
      email: email,
      fullName: cRes[1],
      totalScore: cRes[5],
      timeTaken: cRes[7],
      discProfile: cRes[12],
      catScores: catScores,
      paragraphAnswers: paragraphAnswers
    });
  });

  return reports;
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

  // Determine Pass/Fail status
  let passStatus = "CHƯA XÁC ĐỊNH";
  const configData = ss.getSheetByName('CONFIG').getDataRange().getValues();
  const configHeaders = configData[0];
  const config = configData.find(r => r[0] === cRes[4]); // Match by position
  if (config) {
    const passIdx = configHeaders.indexOf('Pass_Score');
    if (passIdx !== -1 && config[passIdx] !== "") {
      const requiredScore = Number(config[passIdx]);
      passStatus = Number(cRes[5]) >= requiredScore ? "ĐẠT" : "KHÔNG ĐẠT";
    }
  }

  return {
    email: email,
    totalScore: cRes[5],
    timeTaken: cRes[7],
    discProfile: cRes[12],
    catScores: catScores,
    paragraphAnswers: paragraphAnswers,
    passStatus: passStatus
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

  let values = configSheet.getDataRange().getValues();
  if (values.length < 1) return { headers: [], data: [] };

  let headers = values[0];

  // Dynamic parsing of new categories from QUESTIONBANK
  let qbSheet = ss.getSheetByName('QUESTIONBANK');
  if (qbSheet) {
    let qbData = qbSheet.getDataRange().getValues();
    if (qbData.length > 1) {
      let qbHeaders = qbData[0];
      let catIdx = qbHeaders.indexOf('Category');
      let diffIdx = qbHeaders.indexOf('Difficulty');

      if (catIdx !== -1 && diffIdx !== -1) {
        let uniqueCombinations = new Set();
        for (let i = 1; i < qbData.length; i++) {
          let cat = String(qbData[i][catIdx]).trim();
          let diff = String(qbData[i][diffIdx]).trim();
          if (cat && diff) {
            uniqueCombinations.add(`${cat}_${diff}_Count`);
          }
        }

        let headerChanged = false;
        uniqueCombinations.forEach(comb => {
          if (!headers.includes(comb)) {
            // Check if column is placed before Duration_Minutes
            let insertIdx = headers.indexOf('Duration_Minutes');
            if (insertIdx === -1) insertIdx = headers.length;

            headers.splice(insertIdx, 0, comb);
            headerChanged = true;

            // Add 0 to all data rows for the new column
            for (let r = 1; r < values.length; r++) {
              values[r].splice(insertIdx, 0, 0);
            }
          }
        });

        if (headerChanged) {
          configSheet.clearContents();
          configSheet.getRange(1, 1, values.length, headers.length).setValues(values);
          configSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
        }
      }
    }
  }

  return {
    headers: headers,
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
  const resData = ss.getSheetByName('RESULTS').getDataRange().getValues();

  const cRes = resData.find(r => r[2] === email);
  const candidateAnswers = ansData.filter(r => r[1] === email);

  if (candidateAnswers.length === 0 || !cRes) return "Không tìm thấy dữ liệu cho email này.";

  let output = `[BÁO CÁO KẾT QUẢ BÀI TEST - SYSTEM EXPORT]\n`;
  output += `==============================================\n`;
  output += `THÔNG TIN ỨNG VIÊN:\n`;
  output += `Họ tên: ${cRes[1]}\nEmail: ${cRes[2]}\nSố điện thoại: ${cRes[3]}\nVị trí: ${cRes[4]}\n`;
  output += `Tổng điểm: ${cRes[5]}\nThời gian làm bài: ${cRes[7]} giây\n`;
  output += `DISC Profile: ${cRes[12]}\n`;
  output += `==============================================\n\n`;

  // Group answers by category
  let grouped = {};
  candidateAnswers.forEach(ans => {
    const qID = ans[2];
    const qData = qbData.find(q => q[0] === qID);
    if (qData) {
      let cat = qData[1];
      if (!grouped[cat]) grouped[cat] = [];
      grouped[cat].push({ ansData: ans, qData: qData });
    }
  });

  const diffMap = { "1": "Dễ", "2": "Trung bình", "3": "Khó" };

  for (let cat in grouped) {
    output += `--- PHẦN: ${cat.toUpperCase()} ---\n`;
    grouped[cat].forEach(item => {
      const qData = item.qData;
      const ansData = item.ansData;
      const difficulty = diffMap[String(qData[3])] || qData[3];

      output += `Q: ${qData[6]} (Độ khó: ${difficulty})\n`;
      output += `Trả lời: ${ansData[4]}\n`;

      if (cat === 'Personality') {
        output += `Đánh giá DISC nội bộ: Đáp án này thuộc nhóm tính cách nào đó (D, I, S, C).\n`;
      } else if (qData[9] === 'PARAGRAPH') {
        output += `Loại câu hỏi: TỰ LUẬN (PARAGRAPH)\n`;
        output += `Điểm được HR/AI đánh giá: ${ansData[6]}\n`;
      } else {
        output += `Đáp án đúng: ${qData[17]}\n`;
        output += `Kết quả: ${ansData[5] ? 'ĐÚNG' : 'SAI'} (Điểm: ${ansData[6]})\n`;
      }
      output += `\n`;
    });
  }

  output += `[END OF REPORT]`;
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
    const headers = ['Position', 'IQ_1_Count', 'IQ_2_Count', 'IQ_3_Count', 'EQ_1_Count', 'EQ_2_Count', 'EQ_3_Count', 'Problem_Solving_1_Count', 'Problem_Solving_2_Count', 'Problem_Solving_3_Count', 'Leadership_1_Count', 'Leadership_2_Count', 'Leadership_3_Count', 'Personality_1_Count', 'Duration_Minutes', 'Pass_Score'];
    configSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    configSheet.appendRow(['Nhân viên văn phòng', 2, 2, 1, 2, 2, 1, 2, 2, 1, 2, 2, 1, 5, 30, 10]);
    configSheet.appendRow(['Quản lý cấp trung', 1, 2, 2, 1, 2, 2, 1, 2, 2, 1, 2, 2, 5, 45, 15]);
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
    assSheet.getRange(1, 1, 1, 3).setValues([['Assign_Type', 'Assign_Value', 'Assigned_Question_IDs']]).setFontWeight('bold');
  }

  if (!ss.getSheetByName('SYSTEM_SETTINGS')) {
    let setSheet = ss.insertSheet('SYSTEM_SETTINGS');
    setSheet.getRange(1, 1, 1, 2).setValues([['Key', 'Value']]).setFontWeight('bold');
    setSheet.appendRow(['Logo_URL', '']);
    setSheet.appendRow(['Anti_Cheat_Enabled', 'TRUE']);
    setSheet.appendRow(['Anti_Cheat_Max_Violations', '3']);
    setSheet.appendRow(['HR_Notification_Email', '']);
  }
  
  return "Hệ thống đã khởi tạo thành công!";
}

function setupValidation(sheet) {
  const rules = {
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

  // Prevent duplicate submissions
  const resSheet = ss.getSheetByName('RESULTS');
  if (resSheet) {
    const resData = resSheet.getDataRange().getValues();
    const existing = resData.find(r => r[2] === candidateInfo.email);
    if (existing) {
      throw new Error("Email này đã hoàn thành bài test và được ghi nhận hệ thống. Không thể thi lại.");
    }
  }

  const configData = ss.getSheetByName('CONFIG').getDataRange().getValues();
  const configHeaders = configData[0];
  const config = configData.find(r => r[0] === candidateInfo.position);
  if (!config) throw new Error("Không tìm thấy cấu hình cho vị trí này.");

  const allQ = ss.getSheetByName('QUESTIONBANK').getDataRange().getValues();
  const headers = allQ[0];
  let testQuestions = [];

  // Check if assigned specifically
  let assignmentsSheet = ss.getSheetByName('ASSIGNMENTS');
  let specificAssignment = null;
  if (assignmentsSheet) {
    const assData = assignmentsSheet.getDataRange().getValues();
    // Prioritize Email, then Phone, then Position
    specificAssignment = assData.find(r => r[0] === 'Email' && r[1] === candidateInfo.email) ||
                         assData.find(r => r[0] === 'Phone' && r[1] === candidateInfo.phone) ||
                         assData.find(r => r[0] === 'Position' && r[1] === candidateInfo.position);
  }

  if (specificAssignment && specificAssignment[2]) {
    const assignedIds = specificAssignment[2].toString().split(',').map(s => s.trim());
    testQuestions = allQ.slice(1).filter(q => assignedIds.includes(q[0].toString()));
  } else {
    // Regular random logic based on Difficulty mapping
    const diffCounts = {};
    for (let i = 1; i < configHeaders.length; i++) {
      const h = configHeaders[i];
      if (h.endsWith('_Count')) {
        let namePart = h.replace('_Count', '');
        let lastUnderscore = namePart.lastIndexOf('_');

        let cat = namePart;
        let diff = "ANY";

        if (lastUnderscore !== -1) {
            let possibleDiff = namePart.substring(lastUnderscore + 1);
            if (!isNaN(possibleDiff)) {
                diff = possibleDiff;
                cat = namePart.substring(0, lastUnderscore);
            }
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
    candidateInfo: candidateInfo,
    settings: getSystemSettings()
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

  // Notification Logic
  try {
    const settings = getSystemSettings();
    const hrEmail = settings['HR_Notification_Email'];
    if (hrEmail && hrEmail.trim() !== '') {
      const subject = `[Test System] Có ứng viên mới hoàn thành bài test: ${submission.candidateInfo.fullName}`;
      const body = `Hệ thống vừa ghi nhận bài làm mới.\n\nThông tin:\n- Ứng viên: ${submission.candidateInfo.fullName}\n- Email: ${submission.candidateInfo.email}\n- Vị trí: ${submission.candidateInfo.position}\n- Tổng điểm ban đầu: ${score}\n- DISC Profile: ${discProfile}\n\nVui lòng truy cập Admin Panel để xem chi tiết và chấm điểm tự luận.`;
      MailApp.sendEmail(hrEmail.trim(), subject, body);
    }
  } catch(e) {
    // Ignore email errors to not fail the submission
  }

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
