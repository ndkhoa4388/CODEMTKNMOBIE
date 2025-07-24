// ==========================================================================================
// HÀM TIỆN ÍCH VÀ THIẾT LẬP BAN ĐẦU (GIỮ NGUYÊN)
// ==========================================================================================
function transformGoogleDriveUrl(url) {
  if (url.includes('drive.google.com')) {
    var fileId = url.match(/[-\w]{25,}/);
    if (fileId) {
      return 'https://drive.google.com/uc?export=view&id=' + fileId[0];
    }
  }
  return url;
}

// ==========================================================================================
// HÀM XỬ LÝ ĐĂNG NHẬP VÀ BÀI KIỂM TRA (GIỮ NGUYÊN)
// ==========================================================================================
function checkLogin(employeeCode) {
  var sheet = SpreadsheetApp.openById('1YRQySkiMm-_y_18bKLlX8isbvxzb2rPyuC0bry-DdlM').getSheetByName('NS');
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { isValid: false, fullName: "" };
  var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim() === employeeCode.toString().trim()) {
      return { isValid: true, fullName: data[i][1] || '' };
    }
  }
  return { isValid: false, fullName: "" };
}

function getTestGroups() {
  var props = PropertiesService.getScriptProperties();
  var cachedGroups = props.getProperty('testGroups');
  if (cachedGroups) {
    Logger.log('Returning test groups from cache.');
    return JSON.parse(cachedGroups);
  }
  var sheet = SpreadsheetApp.openById('1YRQySkiMm-_y_18bKLlX8isbvxzb2rPyuC0bry-DdlM').getSheetByName('CH');
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var data = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  var uniqueGroups = [...new Set(data.flat())].filter(group => group !== "");
  props.setProperty('testGroups', JSON.stringify(uniqueGroups));
  return uniqueGroups;
}

function getRandomQuestions(testGroup) {
  var sheet = SpreadsheetApp.openById('1YRQySkiMm-_y_18bKLlX8isbvxzb2rPyuC0bry-DdlM').getSheetByName('CH');
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error('Không có câu hỏi nào trong nhóm ' + testGroup);
  var data = sheet.getRange(2, 2, lastRow - 1, 5).getValues(); // Nhóm, Câu hỏi, Đúng, Sai1, Sai2
  var questions = [];
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === testGroup) { // data[i][0] là cột Nhóm (cột B trong sheet CH)
      questions.push({
        question: data[i][1], // Câu hỏi (C)
        correct: data[i][2],  // Đáp án đúng (D)
        wrong1: data[i][3],   // Đáp án sai 1 (E)
        wrong2: data[i][4]    // Đáp án sai 2 (F)
      });
    }
  }
  if (questions.length < 5) {
    throw new Error('Không đủ 5 câu hỏi trong nhóm ' + testGroup + '. Số câu hỏi tìm thấy: ' + questions.length);
  }
  // Shuffling questions
  for (let i = questions.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [questions[i], questions[j]] = [questions[j], questions[i]];
  }
  return questions.slice(0, 5);
}

function checkAnswersAndSave(userAnswers, questions, employeeCode, fullName, testGroup) {
  var score = 0;
  var questionList = [];
  var answerList = [];
  var correctAnswerList = [];

  for (var i = 0; i < 5; i++) {
    if (i < questions.length) {
      questionList.push(questions[i].question.trim());
    } else {
      questionList.push("Không có câu hỏi");
    }
    if (i < userAnswers.length) {
      answerList.push(userAnswers[i] ? userAnswers[i].trim() : "Không trả lời");
    } else {
      answerList.push("Không trả lời");
    }
    if (i < questions.length) {
      correctAnswerList.push(questions[i].correct.trim());
    } else {
      correctAnswerList.push("Không có đáp án");
    }
    if (i < userAnswers.length && i < questions.length && userAnswers[i] === questions[i].correct) {
      score++;
    }
  }

  var wrongCount = 5 - score;
  var questionsString = questionList.join("|");
  var answersString = answerList.join("|");
  var correctAnswersString = correctAnswerList.join("|");
  var correctWrong = score + "/" + wrongCount;
  var sheet = SpreadsheetApp.openById('1YRQySkiMm-_y_18bKLlX8isbvxzb2rPyuC0bry-DdlM').getSheetByName('CTL');
  var timestamp = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss");

  sheet.appendRow([employeeCode, fullName, timestamp, testGroup, questionsString, answersString, correctWrong, score, correctAnswersString]);
  return score;
}

// ==========================================================================================
// HÀM LẤY TÀI LIỆU ĐÀO TẠO (ĐÃ CẬP NHẬT ĐỂ SỬ DỤNG CACHING)
// ==========================================================================================
function getTrainingDocuments() {
  var props = PropertiesService.getScriptProperties();
  var cachedDocuments = props.getProperty('trainingDocumentsCache');
  if (cachedDocuments) {
    Logger.log('Returning training documents from cache.');
    return JSON.parse(cachedDocuments);
  }

  Logger.log('Fetching training documents from sheet and caching.');
  var sheet = SpreadsheetApp.openById('1YRQySkiMm-_y_18bKLlX8isbvxzb2rPyuC0bry-DdlM').getSheetByName('Tai_lieu');
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var data = sheet.getRange(2, 2, lastRow - 1, 5).getValues(); // B: skill, C: link, D: unused, E: note, F: group
  var documents = [];
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] && data[i][1]) {
      var skill = data[i][0].toString().trim();
      var link = data[i][1];
      var note = data[i][3] ? data[i][3].toString().trim() : "Không có ghi chú.";
      var group = data[i][4] ? data[i][4].toString().trim() : "Không xác định"; // Cột F (index 4)
      documents.push({
        skill: skill,
        link: link,
        note: note,
        group: group
      });
    }
  }
  props.setProperty('trainingDocumentsCache', JSON.stringify(documents)); // Cache the data
  return documents;
}

// ==========================================================================================
// HÀM LẤY LỊCH SỬ BÀI LÀM (GIỮ NGUYÊN)
// ==========================================================================================
function getTestHistory(employeeCode) {
  try {
    Logger.log('getTestHistory called with employeeCode: ' + employeeCode + ' (type: ' + typeof employeeCode + ')');
    var spreadsheet = SpreadsheetApp.openById('1YRQySkiMm-_y_18bKLlX8isbvxzb2rPyuC0bry-DdlM');
    var sheet = spreadsheet.getSheetByName('CTL');
    if (!sheet) throw new Error('Không tìm thấy sheet "CTL".');
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    var data = sheet.getRange(2, 1, lastRow - 1, 9).getDisplayValues(); // Lấy giá trị hiển thị để đảm bảo mã NV là text
    Logger.log('Data fetched from CTL sheet (first 5 rows): ' + JSON.stringify(data.slice(0, 5)));
    var tests = [];
    var inputCodeStr = String(employeeCode).trim();

    for (var i = 0; i < data.length; i++) {
      var sheetCodeStr = String(data[i][0]).trim();
      if (sheetCodeStr === inputCodeStr) {
        tests.push({
          employeeCode: sheetCodeStr,
          fullName: data[i][1] || '',
          timestamp: data[i][2] || '',
          testGroup: data[i][3] || '',
          questions: data[i][4] || '',
          answers: data[i][5] || '',
          correctWrong: data[i][6] || '',
          score: data[i][7] || 0,
          correctAnswers: data[i][8] || ''
        });
      }
    }
    Logger.log('Total tests found for ' + inputCodeStr + ': ' + tests.length);
    return tests;
  } catch (error) {
    Logger.log('Error in getTestHistory: ' + error.message + ' Stack: ' + error.stack);
    throw new Error('Lỗi khi tra cứu lịch sử bài làm: ' + error.message);
  }
}

// ==========================================================================================
// HÀM LẤY DỮ LIỆU NHÂN SỰ (ĐÃ CẬP NHẬT ĐỂ SỬ DỤNG CACHING)
// ==========================================================================================
function getPersonnelData() {
  try {
    var props = PropertiesService.getScriptProperties();
    var cachedPersonnel = props.getProperty('personnelDataCache');
    if (cachedPersonnel) {
      Logger.log('Returning personnel data from cache.');
      return JSON.parse(cachedPersonnel);
    }

    Logger.log('Fetching personnel data from sheet and caching.');
    var spreadsheet = SpreadsheetApp.openById('1YRQySkiMm-_y_18bKLlX8isbvxzb2rPyuC0bry-DdlM');
    var sheet = spreadsheet.getSheetByName('Skill_Matrix');
    if (!sheet) {
      throw new Error('Không tìm thấy sheet "Skill_Matrix". Vui lòng kiểm tra tên sheet.');
    }

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    var lastColumn = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];

    var imageColumnIndex = headers.indexOf('Hình ảnh') + 1;
    var mtcvColumnIndex = headers.indexOf('MTCV') + 1;
    var replacement1ColumnIndex = headers.indexOf('THAY THẾ 1') + 1;
    var replacement2ColumnIndex = headers.indexOf('THAY THẾ 2') + 1;
    var genderColumnIndex = headers.indexOf('GIỚI TÍNH') + 1;

    if (imageColumnIndex === 0) throw new Error('Không tìm thấy cột "Hình ảnh" trong sheet Skill_Matrix');
    if (mtcvColumnIndex === 0) throw new Error('Không tìm thấy cột "MTCV" trong sheet Skill_Matrix');
    if (replacement1ColumnIndex === 0) throw new Error('Không tìm thấy cột "THAY THẾ 1" trong sheet Skill_Matrix');
    if (replacement2ColumnIndex === 0) throw new Error('Không tìm thấy cột "THAY THẾ 2" trong sheet Skill_Matrix');
    if (genderColumnIndex === 0) throw new Error('Không tìm thấy cột "GIỚI TÍNH" trong sheet Skill_Matrix');
    var dataRange = sheet.getRange(2, 2, lastRow - 1, lastColumn - 1);
    var data = dataRange.getValues();

    var personnel = [];
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var skills = [];
      for (var colSheet = 15; colSheet <= 28; colSheet++) {
        var skillName = sheet.getRange(1, colSheet).getValue().toString().trim();
        var skillValue = row[colSheet - 2] || 0;
        skills.push({ name: skillName, current: parseFloat(skillValue) || 0 });
      }

      var targetSkillsData = [];
      for (var colSheet = 29; colSheet <= 42; colSheet++) {
        var targetName = sheet.getRange(1, colSheet).getValue().toString().replace("Mục tiêu - ", "").trim();
        var targetValue = row[colSheet - 2] || 0;
        targetSkillsData.push({ name: targetName, target: parseFloat(targetValue) || 0 });
      }

      var matchedSkills = [];
      skills.forEach(function(skill) {
        var targetObj = targetSkillsData.find(function(ts) { return ts.name === skill.name; });
        matchedSkills.push({
            name: skill.name,
            current: skill.current,
            target: targetObj ? targetObj.target : 0
        });
      });
      var overallSkills = [
        { name: sheet.getRange(1, 43).getValue(), current: parseFloat(row[43 - 2]) || 0, target: parseFloat(row[44 - 2]) || 0 },
        { name: sheet.getRange(1, 46).getValue(), current: parseFloat(row[46 - 2]) || 0, target: parseFloat(row[47 - 2]) || 0 },
        { name: sheet.getRange(1, 45).getValue(), current: ((parseFloat(row[45 - 2]) || 0) * 100).toFixed(2), target: 100 }
      ];
      var replacements = [];
      var r1 = sheet.getRange(i + 2, replacement1ColumnIndex).getValue();
      var r2 = sheet.getRange(i + 2, replacement2ColumnIndex).getValue();
      if (r1 && r1.toString().trim()) replacements.push(r1.toString().trim());
      if (r2 && r2.toString().trim()) replacements.push(r2.toString().trim());

      var rawImageUrl = sheet.getRange(i + 2, imageColumnIndex).getValue();
      var imageUrl = rawImageUrl ? transformGoogleDriveUrl(rawImageUrl.toString().trim()) : '';
      var mtcvUrl = sheet.getRange(i + 2, mtcvColumnIndex).getValue();
      var jobDescriptionLink = mtcvUrl ? mtcvUrl.toString().trim() : '';

      var gender = sheet.getRange(i + 2, genderColumnIndex).getValue();
      var radarChartSkillsData = [];
      var radarSkillStartSheetCol = 20;
      var radarSkillEndSheetCol = 28;

      for (var rCol = radarSkillStartSheetCol; rCol <= radarSkillEndSheetCol; rCol++) {
        var skillNameRadar = sheet.getRange(1, rCol).getValue().toString().trim();
        var skillValueRadar = row[rCol - 2];
        var numericSkillValueRadar = parseFloat(skillValueRadar);
        if (isNaN(numericSkillValueRadar)) {
            numericSkillValueRadar = 0;
        }

        radarChartSkillsData.push({
          name: skillNameRadar,
          value: numericSkillValueRadar
        });
      }

      const skillToSwap1_Name = "Kỹ năng thuyết trình";
      const skillToSwap2_Name = "Sử dụng Google Sheets, Google drive";

      let skill1_Index = -1;
      let skill2_Index = -1;
      for (let k = 0; k < radarChartSkillsData.length; k++) {
            if (radarChartSkillsData[k].name === skillToSwap1_Name) {
                skill1_Index = k;
            }
            if (radarChartSkillsData[k].name === skillToSwap2_Name) {
                skill2_Index = k;
            }
      }

      if (skill1_Index !== -1 && skill2_Index !== -1 && skill1_Index !== skill2_Index) {
            const tempSkill = radarChartSkillsData[skill1_Index];
            radarChartSkillsData[skill1_Index] = radarChartSkillsData[skill2_Index];
            radarChartSkillsData[skill2_Index] = tempSkill;
            Logger.log('Swapped radar skills: "' + skillToSwap1_Name + '" and "' + skillToSwap2_Name + '"');
      } else {
            if (skill1_Index === -1) Logger.log('Radar skill to swap not found: "' + skillToSwap1_Name + '"');
            if (skill2_Index === -1) Logger.log('Radar skill to swap not found: "' + skillToSwap2_Name + '"');
      }

      personnel.push({
        employeeCode: row[0] ? row[0].toString().trim() : 'N/A',
        fullName: row[1] || 'N/A',
        age: row[5] || 'N/A',
        seniority: row[6] || 'N/A',
        position: row[8] || 'N/A',
        materialGroup: row[9] || 'N/A',
        overallSkills: overallSkills,
        imageUrl: imageUrl,
        jobDescriptionLink: jobDescriptionLink,
        skills: matchedSkills,
        replacements: replacements,
        radarSkills: radarChartSkillsData,
        gender: gender ? gender.toString().trim() : 'Không xác định'
      });
    }
    props.setProperty('personnelDataCache', JSON.stringify(personnel)); // Cache the data
    return personnel;
  } catch (error) {
    Logger.log('Error in getPersonnelData: ' + error.message + ' Stack: ' + error.stack);
    throw new Error('Lỗi khi lấy dữ liệu nhân sự: ' + error.message);
  }
}


// ==========================================================================================
// NEW: HÀM LẤY DỮ LIỆU LỘ TRÌNH VÀ CẬP NHẬT (ROADMAP)
// ==========================================================================================
function getRoadmapData() {
  try {
    var props = PropertiesService.getScriptProperties();
    var cachedRoadmap = props.getProperty('roadmapDataCache');
    if (cachedRoadmap) {
      Logger.log('Returning roadmap data from cache.');
      return JSON.parse(cachedRoadmap);
    }

    Logger.log('Fetching roadmap data from sheet and caching.');
    var spreadsheet = SpreadsheetApp.openById('1YRQySkiMm-_y_18bKLlX8isbvxzb2rPyuC0bry-DdlM');
    var sheet = spreadsheet.getSheetByName('Lo_trinh');
    if (!sheet) {
      throw new Error('Không tìm thấy sheet "Lo_trinh". Vui lòng kiểm tra tên sheet.');
    }

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    var data = sheet.getRange(2, 1, lastRow - 1, 9).getDisplayValues(); // Get display values for dates/numbers

    var roadmap = [];
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      // Check if the row has meaningful data (e.g., roadmapName or task is not empty)
      if (row[0].toString().trim() === '' && row[2].toString().trim() === '') {
        continue; // Skip empty or incomplete rows
      }

      roadmap.push({
        rowIndex: i + 2, // Store actual sheet row index (1-based)
        roadmapName: row[0] || '', // LỘ TRÌNH
        step: row[1] || '', // BƯỚC
        task: row[2] || '', // CÔNG VIỆC
        startDate: row[3] || '', // BẮT ĐẦU (dd-mm-yyyy)
        endDate: row[4] || '', // KẾT THÚC (dd-mm-yyyy)
        startWeek: row[5] || '', // TUẦN BẮT ĐẦU
        endWeek: row[6] || '', // TUẦN KẾT THÚC
        duration: row[7] || '', // THỜI GIAN
        progress: parseFloat(String(row[8]).replace('%', '')) || 0 // CẬP NHẬT TIẾN ĐỘ (convert to number)
      });
    }
    props.setProperty('roadmapDataCache', JSON.stringify(roadmap)); // Cache the data
    return roadmap;
  } catch (error) {
    Logger.log('Error in getRoadmapData: ' + error.message + ' Stack: ' + error.stack);
    throw new Error('Lỗi khi lấy dữ liệu lộ trình: ' + error.message);
  }
}

function updateRoadmapProgress(rowIndex, progressValue) {
  try {
    var spreadsheet = SpreadsheetApp.openById('1YRQySkiMm-_y_18bKLlX8isbvxzb2rPyuC0bry-DdlM');
    var sheet = spreadsheet.getSheetByName('Lo_trinh');
    if (!sheet) {
      throw new Error('Không tìm thấy sheet "Lo_trinh". Vui lòng kiểm tra tên sheet.');
    }

    // Column I (index 8 in 0-based array, so col 9 in 1-based sheet)
    sheet.getRange(rowIndex, 9).setValue(progressValue + '%');
    Logger.log('Updated roadmap progress for row ' + rowIndex + ' to ' + progressValue + '%');
    // Clear roadmap cache after update
    PropertiesService.getScriptProperties().deleteProperty('roadmapDataCache');
    return { success: true };
  } catch (error) {
    Logger.log('Error in updateRoadmapProgress: ' + error.message + ' Stack: ' + error.stack);
    return { success: false, message: error.message };
  }
}

// Hàm mới để xóa tất cả các bộ nhớ đệm
function clearAllCaches() {
  PropertiesService.getScriptProperties().deleteProperty('trainingDocumentsCache');
  PropertiesService.getScriptProperties().deleteProperty('personnelDataCache');
  PropertiesService.getScriptProperties().deleteProperty('testGroups');
  PropertiesService.getScriptProperties().deleteProperty('roadmapDataCache'); // NEW: Clear roadmap cache
  Logger.log('All caches cleared.');
}

// Hàm này sẽ được kích hoạt mỗi khi dữ liệu trong bảng tính thay đổi
function onSpreadsheetChange() {
  clearAllCaches();
  // Gọi hàm xóa tất cả bộ nhớ đệm
  Logger.log('Caches cleared automatically due to spreadsheet change.');
}

// ==========================================================================================
// HÀM XỬ LÝ YÊU CẦU API TỪ CLIENT (NETLIFY)
// ==========================================================================================
function doGet(e) {
  var action = e.parameter.action;
  var result;

  try {
    switch (action) {
      case 'checkLogin':
        result = checkLogin(e.parameter.employeeCode);
        break;
      case 'getTestGroups':
        result = getTestGroups();
        break;
      case 'getRandomQuestions':
        result = getRandomQuestions(e.parameter.testGroup);
        break;
      case 'getTrainingDocuments':
        result = getTrainingDocuments();
        break;
      case 'getTestHistory':
        result = getTestHistory(e.parameter.employeeCode);
        break;
      case 'getPersonnelData':
        result = getPersonnelData();
        break;
      case 'getRoadmapData':
        result = getRoadmapData();
        break;
      case 'clearAllCaches':
        clearAllCaches();
        result = { success: true, message: 'All caches cleared.' };
        break;
      default:
        return ContentService.createTextOutput(JSON.stringify({ error: 'Invalid action', action: action }))
          .setMimeType(ContentService.MimeType.JSON);
    }

    // Cần thêm tiêu đề CORS cho phản hồi GET
    var output = ContentService.createTextOutput(JSON.stringify(result));
    output.setMimeType(ContentService.MimeType.JSON);
    output.setHeaders({ 'Access-Control-Allow-Origin': '*' }); // Cho phép mọi tên miền
    return output;

  } catch (error) {
    Logger.log('API Error for action ' + action + ': ' + error.message + ' Stack: ' + error.stack);
    var errorOutput = ContentService.createTextOutput(JSON.stringify({ error: error.message, stack: error.stack }));
    errorOutput.setMimeType(ContentService.MimeType.JSON);
    errorOutput.setHeaders({ 'Access-Control-Allow-Origin': '*' });
    return errorOutput;
  }
}

function doPost(e) {
  var action = e.parameter.action;
  var result;

  try {
    var requestBody = JSON.parse(e.postData.contents);

    switch (action) {
      case 'checkAnswersAndSave':
        result = checkAnswersAndSave(
          requestBody.userAnswers,
          requestBody.questions,
          requestBody.employeeCode,
          requestBody.fullName,
          requestBody.testGroup
        );
        break;
      case 'updateRoadmapProgress':
        result = updateRoadmapProgress(
          requestBody.rowIndex,
          requestBody.progressValue
        );
        break;
      default:
        return ContentService.createTextOutput(JSON.stringify({ error: 'Invalid POST action', action: action }))
          .setMimeType(ContentService.MimeType.JSON);
    }

    // Cần thêm tiêu đề CORS cho phản hồi POST
    var output = ContentService.createTextOutput(JSON.stringify(result));
    output.setMimeType(ContentService.MimeType.JSON);
    output.setHeaders({ 'Access-Control-Allow-Origin': '*' }); // Cho phép mọi tên miền
    return output;

  } catch (error) {
    Logger.log('API POST Error for action ' + action + ': ' + error.message + ' Stack: ' + error.stack);
    var errorOutput = ContentService.createTextOutput(JSON.stringify({ error: error.message, stack: error.stack }));
    errorOutput.setMimeType(ContentService.MimeType.JSON);
    errorOutput.setHeaders({ 'Access-Control-Allow-Origin': '*' });
    return errorOutput;
  }
}

// Hàm này là bắt buộc cho CORS Preflight request (OPTIONS)
function doOptions(e) {
  var output = ContentService.createTextOutput('');
  output.setHeaders({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Access-Control-Max-Age': 86400 // Cache preflight for 1 day
  });
  return output;
}
// Hàm này là bắt buộc cho CORS Preflight request (OPTIONS)
function doOptions(e) {
  var output = ContentService.createTextOutput('');
  output.setHeaders({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type', // Quan trọng cho các yêu cầu POST với Content-Type: application/json
    'Access-Control-Max-Age': 86400 // Cache preflight for 1 day
  });
  return output;
}